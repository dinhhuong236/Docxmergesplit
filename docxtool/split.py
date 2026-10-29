"""Tách một file .docx thành nhiều file, mỗi file giữ nguyên style, cài đặt trang,
header/footer, đánh số danh sách... như file gốc."""
from __future__ import annotations

import copy
import os
import re
import unicodedata
from dataclasses import dataclass, field

from lxml import etree

from .docmodel import Numbering, StyleSheet, resolve_numbering
from .ooxml import (O, R, W14, PPR_ORDER, get_ppr, has_visible_content, insert_ordered, local,
                    text_of, w, wa)
from .package import DocxError, Package

MODES = {
    "heading": "Theo đề mục (Heading)",
    "section": "Theo ngắt section",
    "pagebreak": "Theo ngắt trang thủ công",
    "page": "Theo số trang",
}
PROP_TAGS = {w("pPr"), w("rPr"), w("sdtPr"), w("customXmlPr"), w("smartTagPr")}
NOTE_RELS = ("footnotes", "endnotes", "comments")
PRUNABLE_REL_SUFFIX = ("/image", "/hyperlink", "/header", "/footer", "/chart", "/oleObject",
                       "/package", "/video", "/audio", "/media", "/diagramData", "/diagramLayout",
                       "/diagramQuickStyle", "/diagramColors", "/aFChunk", "/control")


@dataclass
class SplitOptions:
    mode: str = "heading"
    level: int = 1              # heading 1..level là điểm tách
    group: int = 1              # số phần (hoặc số trang ở chế độ trang) mỗi file
    ranges: str = ""            # "1-3, 4, 5-8" — để trống = tách đều theo `group`
    name_pattern: str = "{so} - {tieude}"
    keep_numbering: bool = True


@dataclass
class Unit:
    no: int
    start: int
    end: int
    title: str
    first_page: int = 0
    last_page: int = 0


@dataclass
class SplitResult:
    files: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# =============================================================================
# Chuẩn bị: phân tích tài liệu
# =============================================================================
class Prepared:
    def __init__(self, path, opts: SplitOptions):
        self.path = path
        self.opts = opts
        self.pkg = Package.open(path)
        self.docname = self.pkg.main_document_name()
        self.doc = self.pkg.xml(self.docname)
        body = self.doc.find(w("body"))
        if body is None:
            raise DocxError("Tài liệu không có nội dung.")
        self.sheet = StyleSheet(self.pkg, self.docname)
        self.numbering = Numbering(self.pkg, self.docname)
        self.warnings = []
        children = list(body)
        self.body_sect = None
        if children and children[-1].tag == w("sectPr"):
            self.body_sect = children.pop()
        for c in list(body):
            body.remove(c)
        self.break_after = set()
        if opts.mode in ("pagebreak", "page"):
            children = self._split_explicit_breaks(children)
        self.blocks = children
        self._sections()

    # ---- tách đoạn tại ngắt trang thủ công -----------------------------------
    def _split_explicit_breaks(self, blocks):
        out = []
        for b in blocks:
            if b.tag != w("p"):
                out.append(b)
                continue
            cur = b
            while True:
                br = None
                for x in cur.iter(w("br")):
                    if wa(x, "type") == "page" and not _in_textbox(x, cur):
                        br = x
                        break
                if br is None:
                    out.append(cur)
                    break
                p1, p2 = split_paragraph(cur, br, self.sheet, self.numbering)
                out.append(p1)
                self.break_after.add(id(p1))
                cur = p2
        return out

    # ---- section ---------------------------------------------------------------
    def _sections(self):
        self.sect_of = []
        self.sections = []
        k = 0
        for b in self.blocks:
            self.sect_of.append(k)
            s = b.find(f".//{w('pPr')}/{w('sectPr')}") if b.tag != w("tbl") else None
            if s is not None:
                self.sections.append(s)
                k += 1
        self.sections.append(self.body_sect if self.body_sect is not None else etree.Element(w("sectPr")))
        # header/footer hiệu lực (section sau thừa kế section trước)
        self.eff_refs = []
        cur = {}
        for s in self.sections:
            cur = dict(cur)
            for kind in ("headerReference", "footerReference"):
                for r in s.findall(w(kind)):
                    cur[(kind, wa(r, "type", "default"))] = r.get(f"{{{R}}}id")
            self.eff_refs.append(cur)

    # ---- đề mục ------------------------------------------------------------------
    def outline_level(self, p):
        ppr = p.find(w("pPr"))
        if ppr is not None:
            ol = ppr.find(w("outlineLvl"))
            if ol is not None:
                return int(wa(ol, "val", "9"))
            ps = ppr.find(w("pStyle"))
            sid = wa(ps, "val") if ps is not None else None
        else:
            sid = None
        sid = sid or self.sheet.defaults.get("paragraph")
        if sid:
            el = self.sheet.find_ppr(sid, "outlineLvl")
            if el is not None:
                return int(wa(el, "val", "9"))
            # style tên "heading N" không ghi outlineLvl
            m = re.match(r"heading (\d)$", self.sheet.name(sid).lower())
            if m:
                return int(m.group(1)) - 1
        return 9

    def page_break_before(self, p):
        ppr = p.find(w("pPr"))
        el = ppr.find(w("pageBreakBefore")) if ppr is not None else None
        if el is None:
            ps = ppr.find(w("pStyle")) if ppr is not None else None
            sid = wa(ps, "val") if ps is not None else self.sheet.defaults.get("paragraph")
            if sid:
                el = self.sheet.find_ppr(sid, "pageBreakBefore")
        return el is not None and wa(el, "val", "1") not in ("0", "false", "off")

    def sect_break_newpage(self, b):
        if b.tag == w("tbl"):
            return False
        s = b.find(f".//{w('pPr')}/{w('sectPr')}")
        if s is None:
            return False
        # kiểu bắt đầu của section KẾ TIẾP quyết định có sang trang không
        idx = self.sections.index(s)
        nxt = self.sections[idx + 1] if idx + 1 < len(self.sections) else None
        t = nxt.find(w("type")) if nxt is not None else None
        return t is None or wa(t, "val", "nextPage") != "continuous"

    # ---- tạo các "phần" ---------------------------------------------------------------
    def units(self):
        mode = self.opts.mode
        n = len(self.blocks)
        starts = [0]
        titles = {}
        pages = None
        if mode == "heading":
            lvl = max(1, min(9, int(self.opts.level)))
            for i, b in enumerate(self.blocks):
                if b.tag == w("p") and self.outline_level(b) < lvl and text_of(b):
                    if i not in starts:
                        starts.append(i)
                    titles[i] = text_of(b, 80)
            if not titles:
                self.warnings.append(f"Không tìm thấy đề mục cấp 1..{lvl} nào – không tách được.")
        elif mode == "section":
            for i, b in enumerate(self.blocks):
                if b.tag != w("tbl") and b.find(f".//{w('pPr')}/{w('sectPr')}") is not None and i + 1 < n:
                    starts.append(i + 1)
            if len(starts) == 1:
                self.warnings.append("Tài liệu chỉ có 1 section.")
        elif mode == "pagebreak":
            for i, b in enumerate(self.blocks):
                if i > 0 and b.tag == w("p") and self.page_break_before(b):
                    starts.append(i)
                if (id(b) in self.break_after or self.sect_break_newpage(b)) and i + 1 < n:
                    starts.append(i + 1)
            if len(set(starts)) == 1:
                self.warnings.append("Không có ngắt trang thủ công nào.")
        elif mode == "page":
            starts, pages = self._page_flow()
        else:
            raise DocxError(f"Chế độ tách không hợp lệ: {mode}")
        starts = sorted(set(starts))
        bounds = starts + [n]
        out = []
        for k in range(len(starts)):
            s, e = bounds[k], bounds[k + 1]
            if s >= e:
                continue
            seg = self.blocks[s:e]
            if not any(has_visible_content(x) for x in seg) and mode != "page":
                # đoạn trống (vd. chỉ là dòng trắng trước đề mục đầu tiên) -> gộp vào phần sau
                if k + 1 < len(starts):
                    bounds[k + 1] = s
                    continue
            if mode == "heading" and s in titles:
                title = titles[s]
            elif mode == "heading":
                title = "(Phần mở đầu)"
            else:
                title = next((text_of(x, 60) for x in seg if text_of(x)), "")
            u = Unit(no=len(out) + 1, start=s, end=e, title=title)
            if pages:
                u.first_page = pages[s]
                u.last_page = max(u.first_page, pages[e] - 1 if e in pages else self.total_pages)
            out.append(u)
        if not out and n:
            out.append(Unit(no=1, start=0, end=n, title=""))
        return out

    def _page_flow(self):
        lrpb_total = sum(1 for b in self.blocks for _ in b.iter(w("lastRenderedPageBreak")))
        if lrpb_total == 0:
            self.warnings.append(
                "File chưa có thông tin phân trang do Word lưu (có thể được tạo bởi phần mềm khác). "
                "Chỉ tách được tại ngắt trang thủ công/section. Hãy mở file bằng Word, lưu lại rồi "
                "thử lại để tách đúng theo trang.")
        starts, page_of = [0], {}
        page, pending, has_content = 1, 0, False
        for i, b in enumerate(self.blocks):
            sn, internal = False, 0
            if b.tag == w("p"):
                sn = i > 0 and self.page_break_before(b)
                seen_content = False
                for x in b.iter():
                    if x.tag == w("lastRenderedPageBreak"):
                        if seen_content:
                            internal += 1
                        else:
                            sn = sn or i > 0
                    elif x.tag in (w("t"), w("drawing"), w("pict"), w("object"), w("tab"), w("sym")):
                        if x.tag != w("t") or (x.text or "").strip():
                            seen_content = True
            else:
                internal = sum(1 for _ in b.iter(w("lastRenderedPageBreak")))
            if i > 0 and (pending or sn):
                page += max(pending, 1)
                if has_content:
                    starts.append(i)
                    has_content = False
            page_of[i] = page
            pending = internal
            if id(b) in self.break_after or self.sect_break_newpage(b):
                pending += 1
            has_content = has_content or has_visible_content(b)
        self.total_pages = page
        self.page_of = page_of
        return starts, page_of

    # ---- đánh số danh sách liên tục -------------------------------------------------
    def numbering_snapshots(self, starts):
        snaps = {}
        if self.numbering.root is None:
            return snaps
        state, seen = {}, set()
        num = self.numbering
        for i, b in enumerate(self.blocks):
            if i in starts:
                snaps[i] = ({k: dict(v) for k, v in state.items()}, set(seen))
            for p in b.iter(w("p")):
                r = resolve_numbering(p, self.sheet, num)
                if not r:
                    continue
                nid, ilvl = r
                aid = num.abstract_id(nid)
                st = state.setdefault(aid, {})
                if nid not in seen:
                    seen.add(nid)
                    for lv in range(9):
                        if num.has_start_override(nid, lv):
                            st[lv] = num.start(nid, lv) - 1
                            for k in [k for k in st if k > lv]:
                                del st[k]
                val = st.get(ilvl, num.start(nid, ilvl) - 1) + 1
                st[ilvl] = val
                for k in [k for k in st if k > ilvl]:
                    del st[k]
        return snaps


# =============================================================================
def _in_textbox(x, stop):
    for a in x.iterancestors():
        if a is stop:
            return False
        if a.tag == w("txbxContent"):
            return True
    return False


def split_paragraph(p, marker, sheet, numbering):
    """Cắt đoạn văn `p` tại phần tử `marker` (bỏ marker). Trả về (p_trước, p_sau)."""
    path = []
    el = marker
    while el is not p:
        path.append(el)
        el = el.getparent()
    path.reverse()

    def shallow(x):
        c = etree.Element(x.tag, attrib=dict(x.attrib), nsmap=x.nsmap)
        return c

    def rec(node, i, before, after):
        target = path[i]
        seen = False
        for ch in node:
            if ch is target:
                seen = True
                if i == len(path) - 1:
                    continue
                bb, aa = shallow(ch), shallow(ch)
                rec(ch, i + 1, bb, aa)
                if any(c.tag not in PROP_TAGS for c in bb):
                    before.append(bb)
                if any(c.tag not in PROP_TAGS for c in aa):
                    after.append(aa)
            elif ch.tag in PROP_TAGS and not seen:
                before.append(copy.deepcopy(ch))
                after.append(copy.deepcopy(ch))
            elif not seen:
                before.append(copy.deepcopy(ch))
            else:
                after.append(copy.deepcopy(ch))

    p1, p2 = shallow(p), shallow(p)
    rec(p, 0, p1, p2)
    # p1 không giữ sectPr (section kết thúc ở cuối đoạn gốc)
    ppr1 = p1.find(w("pPr"))
    if ppr1 is not None:
        s = ppr1.find(w("sectPr"))
        if s is not None:
            ppr1.remove(s)
    ppr2 = p2.find(w("pPr"))
    if ppr2 is not None:
        pb = ppr2.find(w("pageBreakBefore"))
        if pb is not None:
            ppr2.remove(pb)
    # phần sau không được đánh số lại
    r = resolve_numbering(p, sheet, numbering) if numbering is not None else None
    if r:
        nid, ilvl = r
        ppr2 = get_ppr(p2, create=True)
        lv = numbering.level(nid, ilvl)
        ind = lv.find(f"{w('pPr')}/{w('ind')}") if lv is not None else None
        np_ = ppr2.find(w("numPr"))
        if np_ is not None:
            ppr2.remove(np_)
        np_ = etree.Element(w("numPr"))
        iv = etree.SubElement(np_, w("ilvl"))
        iv.set(w("val"), "0")
        ni = etree.SubElement(np_, w("numId"))
        ni.set(w("val"), "0")
        insert_ordered(ppr2, np_, PPR_ORDER)
        if ind is not None and ppr2.find(w("ind")) is None:
            c = copy.deepcopy(ind)
            for a in ("hanging", "firstLine", "hangingChars", "firstLineChars"):
                c.attrib.pop(w(a), None)
            insert_ordered(ppr2, c, PPR_ORDER)
    # paraId mới cho phần sau
    p2.attrib.pop(f"{{{W14}}}paraId", None)
    p2.attrib.pop(f"{{{W14}}}textId", None)
    return p1, p2


# =============================================================================
# Chọn phần -> file
# =============================================================================
def parse_ranges(spec: str):
    out = []
    for part in re.split(r"[,;]", spec or ""):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d+)\s*(?:-|–|đến|den)\s*(\d+)", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
        elif part.isdigit():
            a = b = int(part)
        else:
            raise DocxError(f"Khoảng không hợp lệ: '{part}'. Ví dụ đúng: 1-3, 5, 7-9")
        if a < 1 or b < a:
            raise DocxError(f"Khoảng không hợp lệ: '{part}'")
        out.append((a, b))
    return out


def plan_files(units, opts: SplitOptions):
    """Trả về danh sách nhóm unit, mỗi nhóm thành 1 file."""
    by_page = opts.mode == "page"
    key = (lambda u: u.first_page) if by_page else (lambda u: u.no)
    if opts.ranges.strip():
        groups = []
        for a, b in parse_ranges(opts.ranges):
            g = [u for u in units if a <= key(u) <= b]
            if g:
                groups.append(g)
        return groups
    n = max(1, int(opts.group or 1))
    groups, cur, bucket = [], [], None
    for u in units:
        bk = (key(u) - 1) // n
        if cur and bk != bucket:
            groups.append(cur)
            cur = []
        cur.append(u)
        bucket = bk
    if cur:
        groups.append(cur)
    return groups


def safe_filename(s: str, maxlen=80) -> str:
    s = unicodedata.normalize("NFC", s)
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', " ", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:maxlen].rstrip(" .")


# =============================================================================
# Dựng file con
# =============================================================================
def set_refs(sect, refs: dict):
    for r in sect.findall(w("headerReference")) + sect.findall(w("footerReference")):
        sect.remove(r)
    items = sorted(refs.items(), key=lambda kv: (kv[0][0] != "headerReference", kv[0][1]))
    for i, ((kind, typ), rid) in enumerate(items):
        el = etree.Element(w(kind))
        el.set(w("type"), typ)
        el.set(f"{{{R}}}id", rid)
        sect.insert(i, el)


def build_chunk(prep: Prepared, s: int, e: int, snaps) -> Package:
    pkg = prep.pkg.clone()
    doc = pkg.xml(prep.docname)
    body = doc.find(w("body"))
    blocks = [copy.deepcopy(b) for b in prep.blocks[s:e]]
    # section của các khối trong phần này
    for i, b in enumerate(blocks):
        if b.tag == w("tbl"):
            continue
        sp = b.find(f".//{w('pPr')}/{w('sectPr')}")
        if sp is not None:
            set_refs(sp, prep.eff_refs[prep.sect_of[s + i]])
    last = blocks[-1]
    sp = last.find(f".//{w('pPr')}/{w('sectPr')}") if last.tag != w("tbl") else None
    if sp is not None:
        sp.getparent().remove(sp)
        body_sect = sp
        k = prep.sect_of[e - 1]
    else:
        k = prep.sect_of[e - 1]
        body_sect = copy.deepcopy(prep.sections[k])
    set_refs(body_sect, prep.eff_refs[k])
    for b in blocks:
        body.append(b)
    body.append(body_sect)

    if prep.opts.keep_numbering and s in snaps:
        fix_numbering(prep, pkg, blocks, snaps[s])
    prune(pkg, prep.docname)
    return pkg


def fix_numbering(prep: Prepared, pkg: Package, blocks, snap):
    state, seen = snap
    num = prep.numbering
    part = num.part
    if part is None:
        return
    root = pkg.xml(part)
    nums = {wa(n, "numId"): n for n in root.findall(w("num"))}
    paras = []
    for b in blocks:
        for p in b.iter(w("p")):
            r = resolve_numbering(p, prep.sheet, num)
            if r:
                paras.append(r)
    done_abs = set()
    for idx, (nid, ilvl) in enumerate(paras):
        aid = num.abstract_id(nid)
        if aid in done_abs:
            continue
        done_abs.add(aid)
        if nid not in seen and num.has_start_override(nid, ilvl):
            continue
        st = state.get(aid, {})
        overrides = {}
        val = st.get(ilvl, num.start(nid, ilvl) - 1) + 1
        if val != num.start(nid, ilvl):
            overrides[ilvl] = val
        later_levels = {lv for (n2, lv) in paras[idx + 1:] if num.abstract_id(n2) == aid}
        for lv in range(ilvl):
            if lv in st:
                v = st[lv] + 1 if lv in later_levels else st[lv]
                if v != num.start(nid, lv):
                    overrides[lv] = v
        if not overrides or nid not in nums:
            continue
        n_el = nums[nid]
        for lv, v in overrides.items():
            ov = None
            for o in n_el.findall(w("lvlOverride")):
                if wa(o, "ilvl") == str(lv):
                    ov = o
            if ov is None:
                ov = etree.SubElement(n_el, w("lvlOverride"))
                ov.set(w("ilvl"), str(lv))
            so = ov.find(w("startOverride"))
            if so is None:
                so = etree.Element(w("startOverride"))
                ov.insert(0, so)
            so.set(w("val"), str(v))


# =============================================================================
# Dọn dẹp file con
# =============================================================================
def prune(pkg: Package, docname: str):
    doc = pkg.xml(docname)
    body = doc.find(w("body"))
    # --- bookmark / vùng comment / quyền chỉnh sửa bị cắt dở ---------------------
    for start_tag, end_tag in (("bookmarkStart", "bookmarkEnd"),
                               ("commentRangeStart", "commentRangeEnd"),
                               ("permStart", "permEnd")):
        starts = {wa(x, "id") for x in body.iter(w(start_tag))}
        ends = {wa(x, "id") for x in body.iter(w(end_tag))}
        for x in list(body.iter(w(start_tag), w(end_tag))):
            other = ends if x.tag == w(start_tag) else starts
            if wa(x, "id") not in other:
                x.getparent().remove(x)
    # --- trường (field) bị cắt dở --------------------------------------------------
    stack, orphans = [], []
    for fc in body.iter(w("fldChar")):
        t = wa(fc, "fldCharType")
        if t == "begin":
            stack.append(fc)
        elif t == "separate":
            if not stack:
                orphans.append(fc)
        elif t == "end":
            if stack:
                stack.pop()
            else:
                orphans.append(fc)
    for fc in orphans:
        run = fc.getparent()
        run.remove(fc)
        if run.tag == w("r") and all(c.tag == w("rPr") for c in run):
            run.getparent().remove(run)
    if stack:
        paras = list(body.iter(w("p")))
        if paras:
            lastp = paras[-1]
            for _ in stack:
                r = etree.SubElement(lastp, w("r"))
                fc = etree.SubElement(r, w("fldChar"))
                fc.set(w("fldCharType"), "end")
    # --- ghi chú, comment không còn dùng ----------------------------------------------
    refs = {
        "footnotes": ("footnote", {wa(x, "id") for x in body.iter(w("footnoteReference"))}),
        "endnotes": ("endnote", {wa(x, "id") for x in body.iter(w("endnoteReference"))}),
        "comments": ("comment", {wa(x, "id") for x in body.iter(w("commentReference"), w("commentRangeStart"))}),
    }
    for kind, (tag, used) in refs.items():
        part = pkg.related_part(docname, kind)
        if not part or not pkg.has(part):
            continue
        root = pkg.xml(part)
        for c in list(root.findall(w(tag))):
            if not wa(c, "type") and wa(c, "id") not in used:
                root.remove(c)
        if kind == "comments":
            _prune_comment_extras(pkg, docname, root)
    # --- quan hệ không còn dùng -------------------------------------------------------
    used = set()
    for el in doc.iter():
        if not isinstance(el.tag, str):
            continue
        for k, v in el.attrib.items():
            if k.startswith("{" + R + "}") or k == f"{{{O}}}relid":
                used.add(v)
    rels = pkg.rels(docname)
    if rels is not None:
        has_dgm = False
        for r in list(rels):
            t = r.get("Type", "")
            if t.endswith(PRUNABLE_REL_SUFFIX) and r.get("Id") not in used:
                rels.remove(r)
            elif t.endswith("/diagramData"):
                has_dgm = True
        if not has_dgm:
            for r in list(rels):
                if r.get("Type", "").endswith("/diagramDrawing"):
                    rels.remove(r)
    pkg.remove_orphans()


def _prune_comment_extras(pkg, docname, comments_root):
    para_ids = {p.get(f"{{{W14}}}paraId") for p in comments_root.iter(w("p"))}
    para_ids.discard(None)
    durable = set()
    for kind in ("commentsExtended", "commentsIds", "commentsExtensible"):
        part = pkg.related_part(docname, kind)
        if not part or not pkg.has(part):
            continue
        root = pkg.xml(part)
        for el in list(root):
            if not isinstance(el.tag, str):
                continue
            attrs = {local(k): v for k, v in el.attrib.items()}
            if "paraId" in attrs:
                if attrs["paraId"] not in para_ids:
                    root.remove(el)
                elif "durableId" in attrs:
                    durable.add(attrs["durableId"])
            elif "durableId" in attrs and kind == "commentsExtensible" and durable and attrs["durableId"] not in durable:
                root.remove(el)


# =============================================================================
def analyze(path, opts: SplitOptions):
    prep = Prepared(path, opts)
    return prep, prep.units()


def split_file(path, out_dir, opts: SplitOptions | None = None, log=print) -> SplitResult:
    opts = opts or SplitOptions()
    prep, units = analyze(path, opts)
    res = SplitResult(warnings=list(prep.warnings))
    groups = plan_files(units, opts)
    if not groups:
        raise DocxError("Không có phần nào để xuất (kiểm tra lại khoảng đã nhập).")
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(path))[0]
    width = max(2, len(str(len(groups))))
    starts = {g[0].start for g in groups}
    snaps = prep.numbering_snapshots(starts) if opts.keep_numbering else {}
    for i, g in enumerate(groups, 1):
        s, e = g[0].start, g[-1].end
        if any(g[j].end != g[j + 1].start for j in range(len(g) - 1)):
            # các phần không liền nhau (hiếm) -> lấy từ đầu đến cuối
            pass
        title = g[0].title.strip("() ") if g[0].title else ""
        pages = f"{g[0].first_page}-{g[-1].last_page}" if opts.mode == "page" else ""
        so = str(i).zfill(width)
        fields = {
            "so": so,
            "tieude": safe_filename(title, 60) or f"Phan {so}",
            "ten": base,
            "trang": pages,
        }
        try:
            name = opts.name_pattern.format(**fields)
        except (KeyError, IndexError, ValueError):
            name = "{so} - {tieude}".format(**fields)
        name = safe_filename(name, 120) or f"{base}_{so}"
        out = os.path.join(out_dir, name + ".docx")
        k = 2
        while os.path.exists(out):
            out = os.path.join(out_dir, f"{name} ({k}).docx")
            k += 1
        pkg = build_chunk(prep, s, e, snaps)
        pkg.save(out)
        res.files.append(out)
        log(f"Đã tạo: {os.path.basename(out)}")
    return res

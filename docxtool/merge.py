"""Gộp nhiều file .docx thành một, giữ nguyên định dạng.

Quy tắc style:
  * Style của file sau giống HỆT (định dạng hiệu lực giống nhau) một style cùng tên ở file
    kết quả  -> dùng chung, không tạo bản sao.
  * Cùng tên nhưng định dạng khác -> tạo style mới tên "<Tên>_1", "<Tên>_2"...
  * Tên chưa có -> chép nguyên style sang.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field

from lxml import etree

from .docmodel import Numbering, StyleSheet, Theme, bake_defaults, resolve_theme
from .ooxml import (CT_FOOTER, CT_HEADER, O, PPR_ORDER, R, RT, TBLPR_ORDER, V, W, W14, WP,
                    get_ppr, insert_ordered, new_nsid, new_para_id, w, wa)
from .package import DocxError, Package

SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
              "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
              "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
              "printerSettings", "sectPrChange"]
VML_ID_RE = re.compile(r"^_x0000_([a-z])(\d+)$")


@dataclass
class MergeOptions:
    new_page: bool = True          # mỗi file bắt đầu ở trang mới (False = section liên tục)
    keep_headers: bool = True      # giữ header/footer riêng của từng file


@dataclass
class MergeReport:
    files: int = 0
    styles_shared: list = field(default_factory=list)
    styles_renamed: list = field(default_factory=list)
    styles_added: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# =============================================================================
class RemapCtx:
    def __init__(self):
        self.style_map = {}
        self.num_map = {}
        self.fn_map, self.en_map, self.cm_map = {}, {}, {}
        self.bm_offset = 0
        self.bm_names = {}
        self.rel_fn = None
        self.default_pstyle = None
        self.default_tblstyle = None
        self.theme = None
        self.strip_colors = False
        self.vml_map = {}
        self.vml_next = None   # callable trả số id VML mới
        self.drop_gobacks = True


def remap_content(elements, ctx: RemapCtx):
    """Đổi các tham chiếu (style, numbering, quan hệ, ghi chú, bookmark...) trong nội dung
    được chép từ tài liệu nguồn sang để khớp với tài liệu đích."""
    bm_drop = []
    for root in elements:
        for el in root.iter():
            tag = el.tag
            if not isinstance(tag, str):
                continue
            # --- quan hệ (ảnh, link, header, chart...) ---------------------
            if ctx.rel_fn is not None:
                for k, v in list(el.attrib.items()):
                    if k.startswith("{" + R + "}") or k == f"{{{O}}}relid":
                        el.set(k, ctx.rel_fn(v))
            # --- VML shape id -------------------------------------------------
            if tag.startswith("{" + V + "}") and el.get("id") and ctx.vml_next:
                m = VML_ID_RE.match(el.get("id"))
                if m:
                    old = el.get("id")
                    if old not in ctx.vml_map:
                        ctx.vml_map[old] = f"_x0000_{m.group(1)}{ctx.vml_next()}"
                    el.set("id", ctx.vml_map[old])
            if not tag.startswith("{" + W + "}"):
                continue
            name = tag[len(W) + 2:]
            if name in ("pStyle", "rStyle", "tblStyle"):
                v = wa(el, "val")
                if v in ctx.style_map:
                    el.set(w("val"), ctx.style_map[v])
            elif name == "numId" and el.getparent() is not None and el.getparent().tag == w("numPr"):
                v = wa(el, "val")
                if v in ctx.num_map:
                    el.set(w("val"), ctx.num_map[v])
            elif name == "footnoteReference":
                v = wa(el, "id")
                if v in ctx.fn_map:
                    el.set(w("id"), ctx.fn_map[v])
            elif name == "endnoteReference":
                v = wa(el, "id")
                if v in ctx.en_map:
                    el.set(w("id"), ctx.en_map[v])
            elif name in ("commentReference", "commentRangeStart", "commentRangeEnd"):
                v = wa(el, "id")
                if v in ctx.cm_map:
                    el.set(w("id"), ctx.cm_map[v])
            elif name in ("bookmarkStart", "bookmarkEnd"):
                v = wa(el, "id")
                if name == "bookmarkStart" and ctx.drop_gobacks and wa(el, "name") == "_GoBack":
                    bm_drop.append(v)
                if v is not None and v.lstrip("-").isdigit():
                    el.set(w("id"), str(int(v) + ctx.bm_offset))
                nm = wa(el, "name")
                if nm and nm.lower() in ctx.bm_names:
                    el.set(w("name"), ctx.bm_names[nm.lower()])
            elif name == "hyperlink":
                a = wa(el, "anchor")
                if a and a.lower() in ctx.bm_names:
                    el.set(w("anchor"), ctx.bm_names[a.lower()])
            elif name == "instrText" and ctx.bm_names and el.text:
                el.text = _rename_in_field(el.text, ctx.bm_names)
            elif name == "fldSimple" and ctx.bm_names and wa(el, "instr"):
                el.set(w("instr"), _rename_in_field(wa(el, "instr"), ctx.bm_names))
        # --- style mặc định ------------------------------------------------
        if ctx.default_pstyle:
            for p in root.iter(w("p")):
                ppr = get_ppr(p, create=True)
                if ppr.find(w("pStyle")) is None:
                    ps = etree.Element(w("pStyle"))
                    ps.set(w("val"), ctx.default_pstyle)
                    ppr.insert(0, ps)
        if ctx.default_tblstyle:
            for t in root.iter(w("tbl")):
                tp = t.find(w("tblPr"))
                if tp is None:
                    tp = etree.Element(w("tblPr"))
                    t.insert(0, tp)
                if tp.find(w("tblStyle")) is None:
                    ts = etree.Element(w("tblStyle"))
                    ts.set(w("val"), ctx.default_tblstyle)
                    insert_ordered(tp, ts, TBLPR_ORDER)
    if ctx.theme is not None or ctx.strip_colors:
        resolve_theme(elements, ctx.theme, ctx.strip_colors)
    if bm_drop:
        dropped = {str(int(v) + ctx.bm_offset) for v in bm_drop if v and v.lstrip("-").isdigit()}
        for root in elements:
            for el in list(root.iter(w("bookmarkStart"), w("bookmarkEnd"))):
                if wa(el, "id") in dropped:
                    el.getparent().remove(el)


def _rename_in_field(text, names):
    def sub(m):
        word = m.group(0)
        return names.get(word.lower(), word)
    return re.sub(r"[A-Za-z_][\w]*", sub, text)


def make_rel_importer(src: Package, sname: str, dst: Package, dname: str, part_cache: dict):
    cache = {}

    def f(rid):
        if rid in cache:
            return cache[rid]
        r = src.rel_by_id(sname, rid)
        if r is None:
            return rid
        if r.get("TargetMode") == "External":
            new = dst.add_rel(dname, r.get("Type"), r.get("Target"), external=True)
        else:
            t = src.resolve(sname, r.get("Target"))
            if not src.has(t):
                return rid
            np_ = dst.import_part(src, t, part_cache)
            new = dst.add_rel(dname, r.get("Type"), dst.relative(dname, np_))
        cache[rid] = new
        return new
    return f


# =============================================================================
class MergeTarget:
    def __init__(self, pkg: Package):
        self.pkg = pkg
        self.docname = pkg.main_document_name()
        self.doc = pkg.xml(self.docname)
        self.body = self.doc.find(w("body"))
        if self.body is None:
            raise DocxError("File đầu tiên không có phần thân (body).")
        self.report = MergeReport(files=1)
        self.vml_counter = self._max_vml() + 1
        settings = pkg.related_part(self.docname, "settings")
        self.even_odd = False
        if settings and pkg.has(settings):
            eo = pkg.xml(settings).find(w("evenAndOddHeaders"))
            self.even_odd = eo is not None and wa(eo, "val", "1") not in ("0", "false", "off")

    def stories(self):
        names = [self.docname]
        for r in self.pkg.rel_list(self.docname):
            t = r.get("Type", "")
            if r.get("TargetMode") == "External":
                continue
            if t.endswith(("/header", "/footer", "/footnotes", "/endnotes", "/comments")):
                n = self.pkg.resolve(self.docname, r.get("Target"))
                if self.pkg.has(n) and n not in names:
                    names.append(n)
        return names

    def _max_vml(self):
        mx = 1024
        for s in self.stories():
            for el in self.pkg.xml(s).iter():
                if isinstance(el.tag, str) and el.tag.startswith("{" + V + "}") and el.get("id"):
                    m = VML_ID_RE.match(el.get("id"))
                    if m:
                        mx = max(mx, int(m.group(2)))
        return mx

    def next_vml(self):
        self.vml_counter += 1
        return self.vml_counter

    # ------------------------------------------------------------------
    def finalize(self):
        """Đảm bảo id ảnh (docPr) và paraId là duy nhất trong toàn tài liệu."""
        n = 1
        used_para = set()
        for s in self.stories():
            root = self.pkg.xml(s)
            for el in root.iter(f"{{{WP}}}docPr"):
                el.set("id", str(n))
                n += 1
            for el in root.iter():
                if not isinstance(el.tag, str):
                    continue
                pid = el.get(f"{{{W14}}}paraId")
                if pid is not None:
                    if pid in used_para:
                        el.set(f"{{{W14}}}paraId", new_para_id(used_para))
                    else:
                        used_para.add(pid)


# =============================================================================
def compute_style_map(ss: StyleSheet, sn: Numbering, ds: StyleSheet, dn: Numbering,
                      st: Theme, dt: Theme):
    """Quyết định style nguồn nào dùng chung, style nào đổi tên."""
    dst_by_name = {}
    for cid in ds.styles:
        dst_by_name.setdefault((ds.stype(cid), ds.name(cid).lower()), []).append(cid)
    dst_canon = {}

    def dcanon(cid):
        if cid not in dst_canon:
            dst_canon[cid] = ds.style_canonical(cid, None, dn, dt)
        return dst_canon[cid]

    forced = set()
    for _round in range(20):
        style_map, merged, new = {}, set(), {}
        used_names = {ds.name(c).lower() for c in ds.styles}
        used_ids = {c.lower() for c in ds.styles}
        for sid in ss.topo_order():
            stype = ss.stype(sid)
            nm = ss.name(sid)
            parent = ss.based_on(sid)
            mparent = style_map.get(parent, parent) if parent else None
            match = None
            if sid not in forced:
                canon = ss.style_canonical(sid, mparent, sn, st)
                cands = list(dst_by_name.get((stype, nm.lower()), []))
                pat = re.compile(re.escape(nm.lower()) + r"_\d+$")
                for (t, n2), ids in dst_by_name.items():
                    if t == stype and pat.match(n2):
                        cands.extend(ids)
                for cid in cands:
                    if dcanon(cid) == canon:
                        match = cid
                        break
            if match:
                style_map[sid] = match
                merged.add(sid)
                continue
            # tạo style mới
            if nm.lower() not in used_names:
                new_name = nm
            else:
                k = 1
                while f"{nm}_{k}".lower() in used_names:
                    k += 1
                new_name = f"{nm}_{k}"
            base_id = sid if new_name == nm else f"{sid}_{new_name.rsplit('_', 1)[1]}"
            new_id, k = base_id, 2
            while new_id.lower() in used_ids:
                new_id = f"{base_id}_{k}"
                k += 1
            used_names.add(new_name.lower())
            used_ids.add(new_id.lower())
            style_map[sid] = new_id
            new[sid] = (new_id, new_name)
        # kiểm tra cặp style liên kết (Heading 1 <-> Heading 1 Char)
        bad = set()
        for sid in ss.styles:
            lk = ss.link(sid)
            if not lk or lk not in ss.styles:
                continue
            if sid in merged and lk in merged:
                if ds.link(style_map[sid]) != style_map[lk]:
                    bad |= {sid, lk}
            elif sid in merged and lk not in merged:
                bad.add(sid)
        bad -= forced
        if not bad:
            return style_map, merged, new
        forced |= bad
    return style_map, merged, new


# =============================================================================
def import_numbering(tgt: MergeTarget, src: Package, sdoc: str, sn: Numbering, dn: Numbering,
                     ctx: RemapCtx, part_cache):
    if sn.root is None or not sn.nums:
        return
    if dn.root is None:
        dn = Numbering(tgt.pkg, tgt.docname, create=True)
    droot = dn.root
    # bullet hình ảnh
    pic_map = {}
    spics = sn.root.findall(w("numPicBullet"))
    if spics:
        rel_fn = make_rel_importer(src, sn.part, tgt.pkg, dn.part, part_cache)
        existing = [int(wa(x, "numPicBulletId", "0")) for x in droot.findall(w("numPicBullet"))]
        nxt = max(existing, default=-1) + 1
        first_abs = droot.find(w("abstractNum"))
        for pb in spics:
            c = copy.deepcopy(pb)
            pic_map[wa(pb, "numPicBulletId")] = str(nxt)
            c.set(w("numPicBulletId"), str(nxt))
            nxt += 1
            sub = RemapCtx()
            sub.rel_fn = rel_fn
            remap_content([c], sub)
            last_pb = droot.findall(w("numPicBullet"))
            if last_pb:
                last_pb[-1].addnext(c)
            elif first_abs is not None:
                first_abs.addprevious(c)
            else:
                droot.insert(0, c)
    abs_ids = [int(x) for x in dn.abstracts if x and x.isdigit()]
    num_ids = [int(x) for x in dn.nums if x and x.isdigit()]
    next_abs = max(abs_ids, default=-1) + 1
    next_num = max(num_ids, default=0) + 1
    abs_map = {}
    for num_id, num in sn.nums.items():
        if num_id in ctx.num_map:
            continue
        aid = sn.abstract_id(num_id)
        if aid not in abs_map and aid in sn.abstracts:
            a = copy.deepcopy(sn.abstracts[aid])
            a.set(w("abstractNumId"), str(next_abs))
            abs_map[aid] = str(next_abs)
            next_abs += 1
            ns = a.find(w("nsid"))
            if ns is not None:
                ns.set(w("val"), new_nsid())
            for el in a.iter(w("pStyle"), w("styleLink"), w("numStyleLink"), w("lvlPicBulletId")):
                v = wa(el, "val")
                if el.tag == w("lvlPicBulletId"):
                    if v in pic_map:
                        el.set(w("val"), pic_map[v])
                elif v in ctx.style_map:
                    el.set(w("val"), ctx.style_map[v])
            resolve_theme([a], ctx.theme, ctx.strip_colors)
            existing = droot.findall(w("abstractNum"))
            if existing:
                existing[-1].addnext(a)
            else:
                pbs = droot.findall(w("numPicBullet"))
                first_num = droot.find(w("num"))
                if pbs:
                    pbs[-1].addnext(a)
                elif first_num is not None:
                    first_num.addprevious(a)
                else:
                    droot.insert(0, a)
        n = copy.deepcopy(num)
        n.set(w("numId"), str(next_num))
        ctx.num_map[num_id] = str(next_num)
        next_num += 1
        an = n.find(w("abstractNumId"))
        if an is not None and wa(an, "val") in abs_map:
            an.set(w("val"), abs_map[wa(an, "val")])
        nums = droot.findall(w("num"))
        if nums:
            nums[-1].addnext(n)
        else:
            cleanup = droot.find(w("numIdMacAtCleanup"))
            if cleanup is not None:
                cleanup.addprevious(n)
            else:
                droot.append(n)


def import_styles(tgt, ss: StyleSheet, ds: StyleSheet, new: dict, ctx: RemapCtx,
                  dd_differ: bool, report: MergeReport, fname: str):
    from .ooxml import STYLE_ORDER
    for sid in ss.topo_order():
        if sid not in new:
            continue
        new_id, new_name = new[sid]
        el = copy.deepcopy(ss.styles[sid])
        el.set(w("styleId"), new_id)
        el.attrib.pop(w("default"), None)
        nm = el.find(w("name"))
        if nm is None:
            nm = insert_ordered(el, etree.Element(w("name")), STYLE_ORDER)
        old_name = wa(nm, "val", sid)
        nm.set(w("val"), new_name)
        for tag in ("basedOn", "next", "link"):
            r = el.find(w(tag))
            if r is not None and wa(r, "val") in ctx.style_map:
                r.set(w("val"), ctx.style_map[wa(r, "val")])
        for nid in el.iter(w("numId")):
            if wa(nid, "val") in ctx.num_map:
                nid.set(w("val"), ctx.num_map[wa(nid, "val")])
        if el.find(w("basedOn")) is None and dd_differ:
            bake_defaults(el, ss.dd_ppr, ss.dd_rpr)
        resolve_theme([el], ctx.theme, ctx.strip_colors)
        # chèn sau style cuối cùng
        last = ds.root.findall(w("style"))
        if last:
            last[-1].addnext(el)
        else:
            ds.root.append(el)
        if new_name != old_name:
            report.styles_renamed.append(f"{fname}: '{old_name}' → '{new_name}'")
        else:
            report.styles_added.append(f"{fname}: '{new_name}'")


def import_notes(tgt: MergeTarget, src: Package, sdoc: str, kind: str, ids_needed: set,
                 ctx: RemapCtx, part_cache, out_map: dict):
    """kind = footnotes | endnotes | comments"""
    if not ids_needed:
        return
    tag = {"footnotes": "footnote", "endnotes": "endnote", "comments": "comment"}[kind]
    spart = src.related_part(sdoc, kind)
    if not spart or not src.has(spart):
        return
    sroot = src.xml(spart)
    dpart = tgt.pkg.related_part(tgt.docname, kind)
    if not dpart or not tgt.pkg.has(dpart):
        # tài liệu đích chưa có -> chép nguyên part
        dpart = tgt.pkg.import_part(src, spart, part_cache)
        tgt.pkg.add_rel(tgt.docname, RT + kind, tgt.pkg.relative(tgt.docname, dpart))
        droot = tgt.pkg.xml(dpart)
        for c in droot.findall(w(tag)):
            out_map[wa(c, "id")] = wa(c, "id")
        sub = copy.copy(ctx)
        sub.rel_fn = None
        sub.default_pstyle = ctx.default_pstyle
        remap_content(list(droot.findall(w(tag))), sub)
        return
    droot = tgt.pkg.xml(dpart)
    ids = [int(wa(c, "id", "0")) for c in droot.findall(w(tag)) if wa(c, "id", "").lstrip("-").isdigit()]
    nxt = max(ids + [0]) + 1
    sub = copy.copy(ctx)
    sub.rel_fn = make_rel_importer(src, spart, tgt.pkg, dpart, part_cache)
    for c in sroot.findall(w(tag)):
        cid = wa(c, "id")
        if wa(c, "type") or cid not in ids_needed:
            continue
        nc = copy.deepcopy(c)
        nc.set(w("id"), str(nxt))
        out_map[cid] = str(nxt)
        nxt += 1
        remap_content([nc], sub)
        droot.append(nc)


def blank_hf_part(tgt: MergeTarget, kind: str) -> str:
    """Tạo header/footer trống (để file sau không 'thừa kế' header của file trước)."""
    tag = "hdr" if kind == "header" else "ftr"
    name = tgt.pkg.unique_name(f"word/{kind}_blank.xml")
    root = etree.Element(w(tag), nsmap={"w": W, "r": R})
    etree.SubElement(root, w("p"))
    tgt.pkg.set_xml(name, root)
    tgt.pkg.set_content_type(name, CT_HEADER if kind == "header" else CT_FOOTER)
    return tgt.pkg.add_rel(tgt.docname, RT + kind, tgt.pkg.relative(tgt.docname, name))


def set_sect_type(sect, value):
    t = sect.find(w("type"))
    if t is None:
        t = insert_ordered(sect, etree.Element(w("type")), SECT_ORDER)
    t.set(w("val"), value)


def merge_fonts(tgt: MergeTarget, src: Package, sdoc: str):
    sp = src.related_part(sdoc, "fontTable")
    dp = tgt.pkg.related_part(tgt.docname, "fontTable")
    if not sp or not dp or not src.has(sp) or not tgt.pkg.has(dp):
        return
    droot = tgt.pkg.xml(dp)
    names = {wa(f, "name", "").lower() for f in droot.findall(w("font"))}
    for f in src.xml(sp).findall(w("font")):
        if wa(f, "name", "").lower() in names:
            continue
        c = copy.deepcopy(f)
        for ch in list(c):
            if ch.tag.startswith(w("embed")):
                c.remove(ch)
        droot.append(c)
        names.add(wa(f, "name", "").lower())


# =============================================================================
def import_document(tgt: MergeTarget, src: Package, fname: str, opts: MergeOptions, log=print):
    pkg = tgt.pkg
    sdoc = src.main_document_name()
    sbody = src.xml(sdoc).find(w("body"))
    if sbody is None:
        tgt.report.warnings.append(f"{fname}: không có nội dung, bỏ qua.")
        return
    report = tgt.report
    ss = StyleSheet(src, sdoc)
    sn = Numbering(src, sdoc)
    ds = StyleSheet(pkg, tgt.docname, create=True)
    dn = Numbering(pkg, tgt.docname)
    st, dt = Theme(src, sdoc), Theme(pkg, tgt.docname)
    themes_differ = st.signature != dt.signature and bool(st.signature)

    ctx = RemapCtx()
    ctx.vml_next = tgt.next_vml
    if themes_differ:
        ctx.theme = st
        ctx.strip_colors = True

    # ---- 1. style ----------------------------------------------------------
    style_map, merged, new = compute_style_map(ss, sn, ds, dn, st, dt)
    ctx.style_map = style_map
    for sid in sorted(merged):
        report.styles_shared.append(f"{fname}: '{ss.name(sid)}'")
    dd_differ = ss.dd_canonical(st) != ds.dd_canonical(dt) if ss.root is not None else False

    # numbering của style dùng chung -> trỏ về numbering của file đích
    for sid in merged:
        sp = ss.styles[sid].find(f"{w('pPr')}/{w('numPr')}/{w('numId')}")
        dp_ = ds.styles[style_map[sid]].find(f"{w('pPr')}/{w('numPr')}/{w('numId')}")
        if sp is not None and dp_ is not None:
            ctx.num_map.setdefault(wa(sp, "val"), wa(dp_, "val"))

    part_cache = {}
    # ---- 2. numbering ------------------------------------------------------
    import_numbering(tgt, src, sdoc, sn, dn, ctx, part_cache)
    # ---- 3. thêm style mới --------------------------------------------------
    import_styles(tgt, ss, ds, new, ctx, dd_differ, report, fname)
    ds.refresh()

    # style mặc định (đoạn văn không ghi style)
    for typ, attr in (("paragraph", "default_pstyle"), ("table", "default_tblstyle")):
        sdef = ss.defaults.get(typ)
        if sdef:
            mapped = style_map.get(sdef, sdef)
            if mapped != ds.defaults.get(typ):
                setattr(ctx, attr, mapped)

    # ---- 4. bookmark --------------------------------------------------------
    dst_bm_ids = [int(wa(b, "id")) for b in tgt.body.iter(w("bookmarkStart"))
                  if wa(b, "id", "").lstrip("-").isdigit()]
    ctx.bm_offset = max(dst_bm_ids + [0]) + 1
    dst_names = {wa(b, "name", "").lower() for b in tgt.body.iter(w("bookmarkStart"))}
    for b in sbody.iter(w("bookmarkStart")):
        nm = wa(b, "name", "")
        if nm.lower() in dst_names and nm != "_GoBack":
            k = 2
            base = nm[:36]
            while f"{base}_{k}".lower() in dst_names:
                k += 1
            ctx.bm_names[nm.lower()] = f"{base}_{k}"
            dst_names.add(f"{base}_{k}".lower())

    # ---- 5. nội dung -----------------------------------------------------
    children = list(sbody)
    src_sect = None
    if children and children[-1].tag == w("sectPr"):
        src_sect = children.pop()
    blocks = [copy.deepcopy(c) for c in children]
    src_sect = copy.deepcopy(src_sect) if src_sect is not None else None

    # ghi chú / chú thích / comment
    def ids(tagname):
        return {wa(e, "id") for b in blocks for e in b.iter(w(tagname))}
    import_notes(tgt, src, sdoc, "footnotes", ids("footnoteReference"), ctx, part_cache, ctx.fn_map)
    import_notes(tgt, src, sdoc, "endnotes", ids("endnoteReference"), ctx, part_cache, ctx.en_map)
    import_notes(tgt, src, sdoc, "comments", ids("commentReference") | ids("commentRangeStart"),
                 ctx, part_cache, ctx.cm_map)

    ctx.rel_fn = make_rel_importer(src, sdoc, pkg, tgt.docname, part_cache)
    if not opts.keep_headers:
        for sp_ in [src_sect] + [b.find(f".//{w('sectPr')}") for b in blocks]:
            if sp_ is not None:
                for r in sp_.findall(w("headerReference")) + sp_.findall(w("footerReference")):
                    sp_.remove(r)
    remap_content(blocks + ([src_sect] if src_sect is not None else []), ctx)

    # header/footer vừa chép sang cần đổi style/numbering
    hf_ctx = copy.copy(ctx)
    hf_ctx.rel_fn = None
    for sname, dname in part_cache.items():
        ctype = pkg.content_type(dname)
        if ctype in (CT_HEADER, CT_FOOTER) and src.has(sname):
            remap_content([pkg.xml(dname)], hf_ctx)

    # ---- 6. section ---------------------------------------------------------
    body = tgt.body
    tsect = body[-1] if len(body) and body[-1].tag == w("sectPr") else None
    if tsect is not None:
        body.remove(tsect)
        last = body[-1] if len(body) else None
        if last is not None and last.tag == w("p") and last.find(f"{w('pPr')}/{w('sectPr')}") is None:
            insert_ordered(get_ppr(last, create=True), tsect, PPR_ORDER)
        else:
            p = etree.SubElement(body, w("p"))
            insert_ordered(get_ppr(p, create=True), tsect, PPR_ORDER)
    if src_sect is None:
        src_sect = copy.deepcopy(tsect) if tsect is not None else etree.Element(w("sectPr"))
        for r in src_sect.findall(w("headerReference")) + src_sect.findall(w("footerReference")):
            src_sect.remove(r)

    first_sect = None
    for b in blocks:
        first_sect = b.find(f".//{w('pPr')}/{w('sectPr')}")
        if first_sect is not None:
            break
    if first_sect is None:
        first_sect = src_sect

    if opts.keep_headers:
        types = ["default"]
        tp = first_sect.find(w("titlePg"))
        if tp is not None and wa(tp, "val", "1") not in ("0", "false", "off"):
            types.append("first")
        if tgt.even_odd:
            types.append("even")
        for kind in ("header", "footer"):
            have = {wa(r, "type", "default") for r in first_sect.findall(w(f"{kind}Reference"))}
            for t in types:
                if t not in have:
                    rid = blank_hf_part(tgt, kind)
                    ref = etree.Element(w(f"{kind}Reference"))
                    ref.set(w("type"), t)
                    ref.set(f"{{{R}}}id", rid)
                    refs = first_sect.findall(w("headerReference")) + first_sect.findall(w("footerReference"))
                    if kind == "header" or not refs:
                        first_sect.insert(0, ref)
                    else:
                        refs[-1].addnext(ref)
    else:
        # bỏ header riêng -> dùng tiếp header của file trước
        pass
    set_sect_type(first_sect, "nextPage" if opts.new_page else "continuous")

    for b in blocks:
        body.append(b)
    body.append(src_sect)

    merge_fonts(tgt, src, sdoc)
    report.files += 1


def merge_files(paths, out_path, opts: MergeOptions | None = None, log=print) -> MergeReport:
    opts = opts or MergeOptions()
    paths = [str(p) for p in paths]
    if len(paths) < 2:
        raise DocxError("Cần ít nhất 2 file để gộp.")
    import os
    log(f"Mở file gốc: {os.path.basename(paths[0])}")
    tgt = MergeTarget(Package.open(paths[0]))
    for p in paths[1:]:
        name = os.path.basename(p)
        log(f"Gộp: {name}")
        import_document(tgt, Package.open(p), name, opts, log)
    tgt.finalize()
    tgt.pkg.save(out_path)
    log(f"Đã lưu: {out_path}")
    return tgt.report

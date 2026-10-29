"""Mô hình hóa styles.xml, numbering.xml, theme của một tài liệu Word."""
from __future__ import annotations

import copy
import hashlib

from lxml import etree

from .ooxml import (A, CT_NUMBERING, CT_STYLES, PPR_ORDER, RPR_ORDER, RT, STYLE_ORDER, W,
                    canonical, local, w, wa)

UI_TAGS = {w(t) for t in ("name", "aliases", "autoRedefine", "hidden", "uiPriority", "semiHidden",
                          "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose",
                          "personalReply", "rsid", "next", "link")}
THEME_FONT_ATTRS = (("asciiTheme", "ascii"), ("hAnsiTheme", "hAnsi"),
                    ("eastAsiaTheme", "eastAsia"), ("cstheme", "cs"))
MERGE_ATTR_TAGS = {w(t) for t in ("rFonts", "spacing", "ind", "lang")}


# =============================================================================
# Theme
# =============================================================================
class Theme:
    def __init__(self, pkg, docname):
        self.fonts: dict[str, str] = {}
        self.signature = b""
        part = pkg.related_part(docname, "theme")
        if not part or not pkg.has(part):
            return
        root = pkg.xml(part)
        fs = root.find(f".//{{{A}}}fontScheme")
        cs = root.find(f".//{{{A}}}clrScheme")
        sig = b""
        for el in (fs, cs):
            if el is not None:
                c = copy.deepcopy(el)
                c.attrib.pop("name", None)
                sig += canonical(c)
        self.signature = sig
        if fs is None:
            return
        for kind in ("major", "minor"):
            f = fs.find(f"{{{A}}}{kind}Font")
            if f is None:
                continue
            latin = f.find(f"{{{A}}}latin")
            ea = f.find(f"{{{A}}}ea")
            cs_ = f.find(f"{{{A}}}cs")
            if latin is not None and latin.get("typeface"):
                self.fonts[f"{kind}Ascii"] = latin.get("typeface")
                self.fonts[f"{kind}HAnsi"] = latin.get("typeface")
            if ea is not None and ea.get("typeface"):
                self.fonts[f"{kind}EastAsia"] = ea.get("typeface")
            if cs_ is not None and cs_.get("typeface"):
                self.fonts[f"{kind}Bidi"] = cs_.get("typeface")


def resolve_theme(elements, theme: Theme | None, strip_colors: bool):
    """Đổi font/màu tham chiếu theme thành giá trị cụ thể (dùng khi 2 file khác theme)."""
    for root in elements:
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            if theme is not None and el.tag == w("rFonts"):
                for tattr, attr in THEME_FONT_ATTRS:
                    v = el.get(w(tattr))
                    if v and theme.fonts.get(v):
                        el.set(w(attr), theme.fonts[v])
                        del el.attrib[w(tattr)]
            if strip_colors and el.tag.startswith("{" + W):
                col = el.get(w("color")) if el.tag != w("color") else el.get(w("val"))
                if col and col.lower() != "auto":
                    for a in ("themeColor", "themeTint", "themeShade"):
                        el.attrib.pop(w(a), None)
                fill = el.get(w("fill"))
                if fill and fill.lower() != "auto":
                    for a in ("themeFill", "themeFillTint", "themeFillShade"):
                        el.attrib.pop(w(a), None)


# =============================================================================
# Styles
# =============================================================================
def merge_props(target, source, order):
    """Bổ sung vào `target` (pPr/rPr) các thuộc tính từ `source` mà target chưa có."""
    for child in source:
        if not isinstance(child.tag, str) or child.tag in (w("rPr"), w("sectPr"), w("pPrChange"), w("rPrChange")):
            continue
        ex = target.find(child.tag)
        if ex is None:
            from .ooxml import insert_ordered
            insert_ordered(target, copy.deepcopy(child), order)
        elif child.tag in MERGE_ATTR_TAGS:
            for k, v in child.attrib.items():
                if k not in ex.attrib:
                    ex.set(k, v)


def bake_defaults(style_el, dd_ppr, dd_rpr):
    """Ghi các thuộc tính mặc định của tài liệu (docDefaults) vào style gốc."""
    from .ooxml import insert_ordered
    stype = wa(style_el, "type", "paragraph")
    if stype == "numbering":
        return
    if dd_ppr is not None and len(dd_ppr) and stype in ("paragraph", "table"):
        ppr = style_el.find(w("pPr"))
        if ppr is None:
            ppr = insert_ordered(style_el, etree.Element(w("pPr")), STYLE_ORDER)
        merge_props(ppr, dd_ppr, PPR_ORDER)
    if dd_rpr is not None and len(dd_rpr):
        rpr = style_el.find(w("rPr"))
        if rpr is None:
            rpr = insert_ordered(style_el, etree.Element(w("rPr")), STYLE_ORDER)
        merge_props(rpr, dd_rpr, RPR_ORDER)


class StyleSheet:
    def __init__(self, pkg, docname, create=False):
        self.pkg = pkg
        self.part = pkg.related_part(docname, "styles")
        if self.part is None and create:
            self.part = pkg.unique_name("word/styles.xml")
            root = etree.Element(w("styles"), nsmap={"w": W})
            pkg.set_xml(self.part, root)
            pkg.set_content_type(self.part, CT_STYLES)
            pkg.add_rel(docname, RT + "styles", pkg.relative(docname, self.part))
        self.root = pkg.xml(self.part) if self.part else None
        self.refresh()

    def refresh(self):
        self.styles: dict[str, etree._Element] = {}
        self.defaults: dict[str, str] = {}
        self.dd_ppr = self.dd_rpr = None
        if self.root is None:
            return
        for st in self.root.findall(w("style")):
            sid = wa(st, "styleId")
            if sid is None:
                continue
            self.styles[sid] = st
            if wa(st, "default") in ("1", "true", "on"):
                self.defaults.setdefault(wa(st, "type", "paragraph"), sid)
        dd = self.root.find(w("docDefaults"))
        if dd is not None:
            self.dd_ppr = dd.find(f"{w('pPrDefault')}/{w('pPr')}")
            self.dd_rpr = dd.find(f"{w('rPrDefault')}/{w('rPr')}")

    # --- tra cứu ---------------------------------------------------------
    def stype(self, sid):
        st = self.styles.get(sid)
        return wa(st, "type", "paragraph") if st is not None else None

    def name(self, sid):
        st = self.styles.get(sid)
        if st is None:
            return sid
        n = st.find(w("name"))
        return wa(n, "val", sid) if n is not None else sid

    def based_on(self, sid):
        st = self.styles.get(sid)
        if st is None:
            return None
        b = st.find(w("basedOn"))
        return wa(b, "val") if b is not None else None

    def link(self, sid):
        st = self.styles.get(sid)
        if st is None:
            return None
        b = st.find(w("link"))
        return wa(b, "val") if b is not None else None

    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(sid)
            sid = self.based_on(sid)
        return out

    def find_ppr(self, sid, tag):
        """Tìm phần tử con `tag` trong pPr theo chuỗi kế thừa style."""
        for s in self.chain(sid):
            ppr = self.styles[s].find(w("pPr"))
            if ppr is not None:
                el = ppr.find(w(tag))
                if el is not None:
                    return el
        return None

    def topo_order(self):
        order, seen = [], set()

        def visit(s, stack=()):
            if s in seen or s in stack:
                return
            b = self.based_on(s)
            if b and b in self.styles:
                visit(b, stack + (s,))
            seen.add(s)
            order.append(s)

        for s in self.styles:
            visit(s)
        return order

    def dd_canonical(self, theme):
        dd = self.root.find(w("docDefaults")) if self.root is not None else None
        if dd is None:
            return b""
        c = copy.deepcopy(dd)
        resolve_theme([c], theme, True)
        return canonical(c)

    def style_canonical(self, sid, mapped_basedon, numbering, theme):
        el = copy.deepcopy(self.styles[sid])
        for a in ("styleId", "default", "customStyle"):
            el.attrib.pop(w(a), None)
        for ch in list(el):
            if ch.tag in UI_TAGS or not isinstance(ch.tag, str):
                el.remove(ch)
        b = el.find(w("basedOn"))
        if b is not None:
            b.set(w("val"), mapped_basedon or wa(b, "val"))
        else:
            bake_defaults(el, self.dd_ppr, self.dd_rpr)
        for nid in el.iter(w("numId")):
            v = wa(nid, "val")
            if v and v != "0" and numbering is not None:
                nid.set(w("val"), "SIG" + hashlib.sha1(numbering.signature(v)).hexdigest())
        resolve_theme([el], theme, True)
        return canonical(el)


# =============================================================================
# Numbering
# =============================================================================
class Numbering:
    def __init__(self, pkg, docname, create=False):
        self.pkg = pkg
        self.part = pkg.related_part(docname, "numbering")
        if self.part is None and create:
            self.part = pkg.unique_name("word/numbering.xml")
            root = etree.Element(w("numbering"), nsmap={"w": W})
            pkg.set_xml(self.part, root)
            pkg.set_content_type(self.part, CT_NUMBERING)
            pkg.add_rel(docname, RT + "numbering", pkg.relative(docname, self.part))
        self.root = pkg.xml(self.part) if self.part else None
        self._sig = {}
        self.refresh()

    def refresh(self):
        self.nums, self.abstracts = {}, {}
        if self.root is None:
            return
        for a in self.root.findall(w("abstractNum")):
            self.abstracts[wa(a, "abstractNumId")] = a
        for n in self.root.findall(w("num")):
            self.nums[wa(n, "numId")] = n

    def abstract_id(self, num_id):
        n = self.nums.get(num_id)
        if n is None:
            return None
        a = n.find(w("abstractNumId"))
        return wa(a, "val") if a is not None else None

    def abstract(self, num_id):
        return self.abstracts.get(self.abstract_id(num_id))

    def signature(self, num_id):
        if num_id in self._sig:
            return self._sig[num_id]
        sig = b"none"
        a = self.abstract(num_id)
        if a is not None:
            c = copy.deepcopy(a)
            c.attrib.pop(w("abstractNumId"), None)
            for ch in list(c):
                if ch.tag in (w("nsid"), w("tmpl"), w("name")):
                    c.remove(ch)
            for el in c.iter():
                for k in list(el.attrib):
                    if local(k) in ("tplc", "tentative"):
                        del el.attrib[k]
            sig = canonical(c)
            for o in self.nums[num_id].findall(w("lvlOverride")):
                sig += canonical(o)
        self._sig[num_id] = sig
        return sig

    def level(self, num_id, ilvl):
        n = self.nums.get(num_id)
        if n is None:
            return None
        for o in n.findall(w("lvlOverride")):
            if wa(o, "ilvl") == str(ilvl):
                lv = o.find(w("lvl"))
                if lv is not None:
                    return lv
        a = self.abstract(num_id)
        if a is None:
            return None
        for lv in a.findall(w("lvl")):
            if wa(lv, "ilvl") == str(ilvl):
                return lv
        return None

    def start(self, num_id, ilvl):
        n = self.nums.get(num_id)
        if n is not None:
            for o in n.findall(w("lvlOverride")):
                if wa(o, "ilvl") == str(ilvl):
                    so = o.find(w("startOverride"))
                    if so is not None:
                        return int(wa(so, "val", "0"))
        lv = self.level(num_id, ilvl)
        if lv is not None:
            s = lv.find(w("start"))
            if s is not None:
                try:
                    return int(wa(s, "val", "0"))
                except ValueError:
                    return 0
        return 0

    def has_start_override(self, num_id, ilvl):
        n = self.nums.get(num_id)
        if n is None:
            return False
        for o in n.findall(w("lvlOverride")):
            if wa(o, "ilvl") == str(ilvl) and o.find(w("startOverride")) is not None:
                return True
        return False

    def ilvl_for_style(self, num_id, style_id):
        a = self.abstract(num_id)
        if a is None:
            return None
        for lv in a.findall(w("lvl")):
            ps = lv.find(w("pStyle"))
            if ps is not None and wa(ps, "val") == style_id:
                return int(wa(lv, "ilvl", "0"))
        return None


def resolve_numbering(p, sheet: StyleSheet, numbering: Numbering):
    """Trả về (numId, ilvl) thực sự của đoạn văn (xét cả numbering kế thừa từ style)."""
    if numbering is None or numbering.root is None:
        return None
    ppr = p.find(w("pPr"))
    num_id = ilvl = None
    sid = None
    if ppr is not None:
        np_ = ppr.find(w("numPr"))
        ps = ppr.find(w("pStyle"))
        sid = wa(ps, "val") if ps is not None else None
        if np_ is not None:
            n = np_.find(w("numId"))
            lv = np_.find(w("ilvl"))
            if n is not None:
                num_id = wa(n, "val")
            if lv is not None:
                ilvl = int(wa(lv, "val", "0"))
    if sid is None:
        sid = sheet.defaults.get("paragraph")
    if num_id is None and sid:
        for s in sheet.chain(sid):
            ppr2 = sheet.styles[s].find(w("pPr"))
            np_ = ppr2.find(w("numPr")) if ppr2 is not None else None
            if np_ is None:
                continue
            n = np_.find(w("numId"))
            lv = np_.find(w("ilvl"))
            if num_id is None and n is not None:
                num_id = wa(n, "val")
            if ilvl is None and lv is not None:
                ilvl = int(wa(lv, "val", "0"))
            if num_id is not None:
                break
        if num_id and ilvl is None:
            ilvl = numbering.ilvl_for_style(num_id, sid)
    if not num_id or num_id == "0" or num_id not in numbering.nums:
        return None
    return num_id, (ilvl or 0)

"""Hằng số namespace và các hàm tiện ích làm việc với XML của Word (OOXML)."""
from __future__ import annotations

import copy
import random
import re

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"

NSMAP = {"w": W, "r": R, "wp": WP, "w14": W14, "v": V, "o": O, "a": A}

CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"


def w(tag: str) -> str:
    """Tên đầy đủ (Clark notation) của thẻ trong namespace w."""
    return f"{{{W}}}{tag}"


def wa(el, name: str, default=None):
    """Đọc thuộc tính w:name."""
    return el.get(w(name), default)


def xp(el, path: str):
    return el.xpath(path, namespaces=NSMAP)


def local(tag) -> str:
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def is_w(el, name: str) -> bool:
    return isinstance(el.tag, str) and el.tag == w(name)


def parser() -> etree.XMLParser:
    return etree.XMLParser(
        remove_blank_text=False,
        resolve_entities=False,
        no_network=True,
        huge_tree=True,
    )


def parse(data: bytes):
    return etree.fromstring(data, parser())


def serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Thứ tự phần tử con theo schema (để chèn phần tử mới đúng chỗ)
# ---------------------------------------------------------------------------
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
    "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
    "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]
STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
    "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose",
    "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
]
TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
    "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
]
NUM_ORDER = ["abstractNumId", "lvlOverride"]
SECT_REF_TAGS = ("headerReference", "footerReference")


def insert_ordered(parent, child, order):
    """Chèn child vào parent theo đúng thứ tự schema."""
    name = local(child.tag)
    try:
        idx = order.index(name)
    except ValueError:
        parent.append(child)
        return child
    for i, existing in enumerate(parent):
        en = local(existing.tag)
        if en in order and order.index(en) > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def get_or_add(parent, name, order):
    el = parent.find(w(name))
    if el is None:
        el = etree.Element(w(name))
        insert_ordered(parent, el, order)
    return el


def get_ppr(p, create=False):
    ppr = p.find(w("pPr"))
    if ppr is None and create:
        ppr = etree.Element(w("pPr"))
        p.insert(0, ppr)
    return ppr


def para_style(p):
    ppr = p.find(w("pPr"))
    if ppr is None:
        return None
    ps = ppr.find(w("pStyle"))
    return None if ps is None else wa(ps, "val")


def text_of(el, limit=None) -> str:
    """Lấy chữ hiển thị của một phần tử (bỏ qua chữ đã xóa trong track changes)."""
    out = []
    for t in el.iter(w("t"), w("tab")):
        if t.tag == w("tab"):
            out.append(" ")
        else:
            out.append(t.text or "")
        if limit and sum(len(x) for x in out) > limit * 2:
            break
    s = re.sub(r"\s+", " ", "".join(out)).strip()
    return s[:limit] if limit else s


VISIBLE_TAGS = {w("t"), w("drawing"), w("pict"), w("object"), w("sym"), w("tab")}


def has_visible_content(el) -> bool:
    for x in el.iter():
        if x.tag in VISIBLE_TAGS:
            if x.tag == w("t") and not (x.text or "").strip():
                continue
            return True
        if x.tag == w("tbl"):
            return True
    return False


def new_para_id(used: set) -> str:
    while True:
        v = "%08X" % random.randint(1, 0x7FFFFFFF)
        if v not in used:
            used.add(v)
            return v


def new_nsid() -> str:
    return "%08X" % random.randint(0x10000000, 0xFFFFFFFE)


def canonical(el) -> bytes:
    """Chuỗi chuẩn hóa để so sánh hai phần tử XML có giống nhau không."""
    c = copy.deepcopy(el)
    for x in c.iter():
        if isinstance(x.tag, str):
            # bỏ khoảng trắng do định dạng file
            if x.text is not None and not x.text.strip() and len(x):
                x.text = None
            if x.tail is not None and not x.tail.strip():
                x.tail = None
    return etree.tostring(c, method="c14n", exclusive=True)

"""Đọc / ghi gói .docx (ZIP + Open Packaging Conventions) ở mức part và relationship."""
from __future__ import annotations

import copy
import hashlib
import posixpath
import zipfile

from lxml import etree

from .ooxml import CT_NS, REL_NS, RT, STRICT_W, W, parse, serialize


class DocxError(Exception):
    pass


class Package:
    def __init__(self):
        self.parts: dict[str, bytes] = {}
        self._xml: dict[str, etree._Element] = {}
        self.source = None
        self._media_hash: dict[str, str] | None = None

    # ------------------------------------------------------------------ I/O
    @classmethod
    def open(cls, path) -> "Package":
        pkg = cls()
        pkg.source = str(path)
        try:
            zf = zipfile.ZipFile(path)
        except zipfile.BadZipFile as e:
            raise DocxError(f"'{path}' không phải file .docx hợp lệ (có thể là .doc cũ hoặc file bị mã hóa).") from e
        with zf:
            for info in zf.infolist():
                name = info.filename
                if info.is_dir():
                    continue
                norm = posixpath.normpath(name.replace("\\", "/")).lstrip("/")
                if norm.startswith("..") or norm != name.lstrip("/"):
                    continue  # bỏ entry bất thường
                pkg.parts[norm] = zf.read(info)
        if "[Content_Types].xml" not in pkg.parts:
            raise DocxError(f"'{path}' thiếu [Content_Types].xml – không phải file Word.")
        doc = pkg.main_document_name()
        if doc is None:
            raise DocxError(f"Không tìm thấy nội dung chính trong '{path}'.")
        root = pkg.xml(doc)
        if root.tag.startswith("{" + STRICT_W):
            raise DocxError(
                f"'{path}' đang lưu ở định dạng 'Strict Open XML'. Hãy mở trong Word và "
                "Lưu thành 'Word Document (.docx)' thông thường rồi thử lại.")
        if root.tag != f"{{{W}}}document":
            raise DocxError(f"'{path}' không phải tài liệu Word.")
        return pkg

    def save(self, path):
        order = ["[Content_Types].xml", "_rels/.rels"]
        names = order + sorted(n for n in self.all_names() if n not in order)
        tmp = str(path) + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for n in names:
                if self.has(n):
                    zf.writestr(n, self.get_bytes(n))
        import os
        os.replace(tmp, path)

    def clone(self) -> "Package":
        p = Package()
        p.source = self.source
        p.parts = dict(self.parts)
        p._xml = {k: copy.deepcopy(v) for k, v in self._xml.items()}
        return p

    # ---------------------------------------------------------------- parts
    def all_names(self):
        return set(self.parts) | set(self._xml)

    def has(self, name) -> bool:
        return name in self._xml or name in self.parts

    def xml(self, name):
        if name not in self._xml:
            if name not in self.parts:
                raise KeyError(name)
            self._xml[name] = parse(self.parts[name])
        return self._xml[name]

    def set_xml(self, name, root):
        self._xml[name] = root
        self.parts.pop(name, None)

    def get_bytes(self, name) -> bytes:
        if name in self._xml:
            return serialize(self._xml[name])
        return self.parts[name]

    def set_bytes(self, name, data: bytes):
        self._xml.pop(name, None)
        self.parts[name] = data
        self._media_hash = None

    def remove(self, name):
        self._xml.pop(name, None)
        self.parts.pop(name, None)
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName") == "/" + name:
                ct.remove(o)
        rn = self.rels_name(name)
        if self.has(rn):
            self._xml.pop(rn, None)
            self.parts.pop(rn, None)

    def unique_name(self, name: str) -> str:
        existing = {n.lower() for n in self.all_names()}
        if name.lower() not in existing:
            return name
        d, b = posixpath.split(name)
        stem, ext = posixpath.splitext(b)
        i = 2
        while True:
            cand = posixpath.join(d, f"{stem}_{i}{ext}")
            if cand.lower() not in existing:
                return cand
            i += 1

    # -------------------------------------------------------- content types
    def content_type(self, name: str):
        ct = self.xml("[Content_Types].xml")
        for o in ct.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName", "").lower() == "/" + name.lower():
                return o.get("ContentType")
        ext = posixpath.splitext(name)[1][1:].lower()
        return self.default_ct(ext)

    def default_ct(self, ext):
        ct = self.xml("[Content_Types].xml")
        for d in ct.findall(f"{{{CT_NS}}}Default"):
            if d.get("Extension", "").lower() == ext:
                return d.get("ContentType")
        return None

    def set_content_type(self, name: str, ctype: str):
        if not ctype:
            return
        ct = self.xml("[Content_Types].xml")
        ext = posixpath.splitext(name)[1][1:].lower()
        dflt = self.default_ct(ext)
        if dflt == ctype:
            return
        if dflt is None and ext not in ("xml", "rels", "") and not ctype.endswith("+xml"):
            el = etree.SubElement(ct, f"{{{CT_NS}}}Default")
            el.set("Extension", ext)
            el.set("ContentType", ctype)
            # Default phải đứng trước Override
            ct.remove(el)
            first_ov = ct.find(f"{{{CT_NS}}}Override")
            if first_ov is None:
                ct.append(el)
            else:
                first_ov.addprevious(el)
            return
        for o in ct.findall(f"{{{CT_NS}}}Override"):
            if o.get("PartName", "").lower() == "/" + name.lower():
                o.set("ContentType", ctype)
                return
        el = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        el.set("PartName", "/" + name)
        el.set("ContentType", ctype)

    # -------------------------------------------------------- relationships
    @staticmethod
    def rels_name(part: str) -> str:
        d, b = posixpath.split(part)
        return posixpath.join(d, "_rels", b + ".rels")

    def rels(self, part: str, create=False):
        rn = self.rels_name(part)
        if not self.has(rn):
            if not create:
                return None
            root = etree.Element(f"{{{REL_NS}}}Relationships", nsmap={None: REL_NS})
            self.set_xml(rn, root)
        return self.xml(rn)

    def rel_list(self, part: str):
        root = self.rels(part)
        if root is None:
            return []
        return list(root.findall(f"{{{REL_NS}}}Relationship"))

    def rel_by_id(self, part, rid):
        for r in self.rel_list(part):
            if r.get("Id") == rid:
                return r
        return None

    @staticmethod
    def resolve(part: str, target: str) -> str:
        target = target.split("#", 1)[0]
        if target.startswith("/"):
            return posixpath.normpath(target.lstrip("/"))
        base = posixpath.dirname(part)
        return posixpath.normpath(posixpath.join(base, target))

    @staticmethod
    def relative(from_part: str, to_part: str) -> str:
        base = posixpath.dirname(from_part) or "."
        return posixpath.relpath(to_part, base)

    def add_rel(self, part: str, rtype: str, target: str, external=False, rid=None) -> str:
        root = self.rels(part, create=True)
        ids = {r.get("Id") for r in root}
        if rid is None or rid in ids:
            n = len(ids) + 1
            while f"rId{n}" in ids:
                n += 1
            rid = f"rId{n}"
        el = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
        el.set("Id", rid)
        el.set("Type", rtype)
        el.set("Target", target)
        if external:
            el.set("TargetMode", "External")
        return rid

    def related_part(self, part: str, type_suffix: str):
        for r in self.rel_list(part):
            if r.get("Type", "").endswith("/" + type_suffix) and r.get("TargetMode") != "External":
                return self.resolve(part, r.get("Target"))
        return None

    def main_document_name(self):
        for r in self.rel_list(""):
            if r.get("Type", "").endswith("/officeDocument"):
                name = self.resolve("", r.get("Target"))
                if self.has(name):
                    return name
        return "word/document.xml" if self.has("word/document.xml") else None

    # -------------------------------------------------------- copy between packages
    def media_hash(self):
        if self._media_hash is None:
            self._media_hash = {}
            for n, data in self.parts.items():
                if n.startswith("word/media/"):
                    self._media_hash.setdefault(hashlib.sha1(data).hexdigest(), n)
        return self._media_hash

    def import_part(self, src: "Package", sname: str, cache: dict, dir_hint: str | None = None) -> str:
        """Sao chép part `sname` (và các part con của nó) từ src sang gói này.
        Trả về tên part mới. `cache` tránh sao chép lặp."""
        if sname in cache:
            return cache[sname]
        data = src.get_bytes(sname)
        ctype = src.content_type(sname)
        srels = src.rels(sname)
        # ảnh giống hệt nhau thì dùng chung
        if srels is None and sname.startswith("word/media/"):
            h = hashlib.sha1(data).hexdigest()
            existing = self.media_hash().get(h)
            if existing and self.content_type(existing) == ctype:
                cache[sname] = existing
                return existing
        dname = self.unique_name(sname if dir_hint is None else posixpath.join(dir_hint, posixpath.basename(sname)))
        cache[sname] = dname
        if sname in src._xml:
            self.set_xml(dname, copy.deepcopy(src._xml[sname]))
        else:
            self.parts[dname] = data
            if dname.startswith("word/media/") and self._media_hash is not None:
                self._media_hash.setdefault(hashlib.sha1(data).hexdigest(), dname)
        self.set_content_type(dname, ctype)
        if srels is not None:
            droot = self.rels(dname, create=True)
            for r in srels.findall(f"{{{REL_NS}}}Relationship"):
                nr = copy.deepcopy(r)
                if r.get("TargetMode") != "External":
                    target = src.resolve(sname, r.get("Target"))
                    if src.has(target):
                        new = self.import_part(src, target, cache)
                        frag = r.get("Target").split("#", 1)
                        nr.set("Target", self.relative(dname, new) + ("#" + frag[1] if len(frag) > 1 else ""))
                droot.append(nr)
        return dname

    def reachable_parts(self):
        seen = set()
        stack = [""]
        while stack:
            part = stack.pop()
            for r in self.rel_list(part):
                if r.get("TargetMode") == "External":
                    continue
                t = self.resolve(part, r.get("Target"))
                if t not in seen and self.has(t):
                    seen.add(t)
                    stack.append(t)
        return seen

    def remove_orphans(self):
        keep = self.reachable_parts()
        for n in list(self.all_names()):
            if n == "[Content_Types].xml" or "/_rels/" in "/" + n or n.endswith(".rels"):
                continue
            if n not in keep:
                self.remove(n)
        # rels của part không còn tồn tại
        for n in list(self.all_names()):
            if n.endswith(".rels") and n != "_rels/.rels":
                d, b = posixpath.split(n)
                owner = posixpath.join(posixpath.dirname(d), b[:-5])
                if not self.has(owner):
                    self._xml.pop(n, None)
                    self.parts.pop(n, None)

    RT = RT

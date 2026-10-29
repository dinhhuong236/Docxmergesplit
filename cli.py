#!/usr/bin/env python3
"""Dòng lệnh DocxTool.

  python3 cli.py merge  out.docx a.docx b.docx c.docx [--continuous] [--no-headers]
  python3 cli.py split  in.docx  thu_muc_ra  [--mode heading|section|pagebreak|page]
                        [--level 1] [--group 1] [--ranges "1-3,5"] [--name "{so} - {tieude}"]
                        [--restart-numbering]
  python3 cli.py preview in.docx [--mode ...] [--level 1]
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxtool import DocxError, MergeOptions, SplitOptions, analyze, merge_files, split_file  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(description="Gộp / tách file Word giữ nguyên định dạng")
    sub = ap.add_subparsers(dest="cmd", required=True)

    m = sub.add_parser("merge", help="gộp nhiều file")
    m.add_argument("output")
    m.add_argument("inputs", nargs="+")
    m.add_argument("--continuous", action="store_true", help="không sang trang mới giữa các file")
    m.add_argument("--no-headers", action="store_true", help="không giữ header/footer riêng từng file")

    for name in ("split", "preview"):
        s = sub.add_parser(name)
        s.add_argument("input")
        if name == "split":
            s.add_argument("outdir")
        s.add_argument("--mode", default="heading", choices=["heading", "section", "pagebreak", "page"])
        s.add_argument("--level", type=int, default=1)
        s.add_argument("--group", type=int, default=1)
        s.add_argument("--ranges", default="")
        s.add_argument("--name", default="{so} - {tieude}")
        s.add_argument("--restart-numbering", action="store_true")

    a = ap.parse_args(argv)
    try:
        if a.cmd == "merge":
            rep = merge_files(a.inputs, a.output, MergeOptions(new_page=not a.continuous,
                                                                keep_headers=not a.no_headers))
            for s in rep.styles_renamed:
                print("  đổi tên style:", s)
        else:
            opts = SplitOptions(mode=a.mode, level=a.level, group=a.group, ranges=a.ranges,
                                name_pattern=a.name, keep_numbering=not a.restart_numbering)
            if a.cmd == "preview":
                prep, units = analyze(a.input, opts)
                for u in units:
                    pg = f"  [trang {u.first_page}-{u.last_page}]" if u.first_page else ""
                    print(f"{u.no:>4}. {u.title}{pg}")
                for wmsg in prep.warnings:
                    print("⚠", wmsg)
            else:
                res = split_file(a.input, a.outdir, opts)
                for wmsg in res.warnings:
                    print("⚠", wmsg)
    except DocxError as e:
        print("Lỗi:", e, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

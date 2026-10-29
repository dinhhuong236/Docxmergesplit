#!/usr/bin/env python3
"""Giao diện DocxTool cho macOS: gộp và tách file Word (.docx) giữ nguyên định dạng."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docxtool import (MODES, DocxError, MergeOptions, SplitOptions, analyze,  # noqa: E402
                      merge_files, split_file)

APP_NAME = "DocxTool – Gộp & Tách Word"


def reveal_in_finder(path):
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path] if os.path.isfile(path) else ["open", path])


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("820x640")
        self.minsize(720, 560)
        self.q: queue.Queue = queue.Queue()
        self.busy = False

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        self.merge_tab = ttk.Frame(nb, padding=10)
        self.split_tab = ttk.Frame(nb, padding=10)
        nb.add(self.merge_tab, text="  Gộp file  ")
        nb.add(self.split_tab, text="  Tách file  ")
        self._build_merge()
        self._build_split()

        logf = ttk.LabelFrame(self, text="Nhật ký", padding=6)
        logf.pack(fill="both", padx=10, pady=(0, 10))
        self.log_box = tk.Text(logf, height=8, wrap="word", state="disabled")
        self.log_box.pack(fill="both", expand=True)
        self.after(100, self._poll)

    # ------------------------------------------------------------------ chung
    def log(self, msg):
        self.q.put(("log", msg))

    def _poll(self):
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "log":
                    self.log_box.configure(state="normal")
                    self.log_box.insert("end", data + "\n")
                    self.log_box.see("end")
                    self.log_box.configure(state="disabled")
                elif kind == "done":
                    self.busy = False
                    title, msg, reveal = data
                    if messagebox.askyesno(title, msg + "\n\nMở thư mục kết quả trong Finder?"):
                        reveal_in_finder(reveal)
                elif kind == "error":
                    self.busy = False
                    messagebox.showerror("Lỗi", data)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def run_bg(self, fn):
        if self.busy:
            messagebox.showinfo(APP_NAME, "Đang xử lý, vui lòng đợi…")
            return
        self.busy = True

        def worker():
            try:
                fn()
            except DocxError as e:
                self.q.put(("error", str(e)))
            except Exception as e:  # noqa: BLE001
                import traceback
                self.log(traceback.format_exc())
                self.q.put(("error", f"Lỗi không mong muốn: {e}"))
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------ GỘP
    def _build_merge(self):
        f = self.merge_tab
        ttk.Label(f, text="Thứ tự file trong danh sách = thứ tự nội dung trong file kết quả. "
                          "File đầu tiên là file gốc (giữ cài đặt chung của tài liệu).",
                  wraplength=760).pack(anchor="w")
        row = ttk.Frame(f)
        row.pack(fill="both", expand=True, pady=6)
        self.merge_list = tk.Listbox(row, selectmode="extended", height=10)
        self.merge_list.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(row, command=self.merge_list.yview)
        sb.pack(side="left", fill="y")
        self.merge_list.configure(yscrollcommand=sb.set)
        self.merge_paths: list[str] = []
        btns = ttk.Frame(row)
        btns.pack(side="left", fill="y", padx=(8, 0))
        for text, cmd in (("Thêm file…", self.m_add), ("Xóa", self.m_remove), ("Lên ▲", lambda: self.m_move(-1)),
                          ("Xuống ▼", lambda: self.m_move(1)), ("Sắp xếp A→Z", self.m_sort), ("Xóa hết", self.m_clear)):
            ttk.Button(btns, text=text, command=cmd, width=14).pack(pady=2)

        opt = ttk.LabelFrame(f, text="Tùy chọn", padding=6)
        opt.pack(fill="x")
        self.m_newpage = tk.BooleanVar(value=True)
        self.m_headers = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="Mỗi file bắt đầu ở trang mới (bỏ chọn = nối liền mạch)",
                        variable=self.m_newpage).pack(anchor="w")
        ttk.Checkbutton(opt, text="Giữ header/footer, khổ giấy, lề riêng của từng file",
                        variable=self.m_headers).pack(anchor="w")
        ttk.Label(opt, foreground="#555", wraplength=760,
                  text="Style trùng tên & giống hệt định dạng sẽ được dùng chung; trùng tên nhưng khác "
                       "định dạng sẽ đổi thành 'Tên_1', 'Tên_2'…").pack(anchor="w", pady=(4, 0))
        ttk.Button(f, text="Gộp và lưu…", command=self.m_run).pack(anchor="e", pady=(8, 0))

    def _refresh_mlist(self):
        self.merge_list.delete(0, "end")
        for i, p in enumerate(self.merge_paths, 1):
            self.merge_list.insert("end", f"{i:>2}. {os.path.basename(p)}   —   {os.path.dirname(p)}")

    def m_add(self):
        files = filedialog.askopenfilenames(title="Chọn file Word", filetypes=[("Word", "*.docx *.docm"), ("Tất cả", "*")])
        self.merge_paths.extend(files)
        self._refresh_mlist()

    def m_remove(self):
        for i in sorted(self.merge_list.curselection(), reverse=True):
            del self.merge_paths[i]
        self._refresh_mlist()

    def m_clear(self):
        self.merge_paths.clear()
        self._refresh_mlist()

    def m_sort(self):
        self.merge_paths.sort(key=lambda p: os.path.basename(p).lower())
        self._refresh_mlist()

    def m_move(self, d):
        sel = list(self.merge_list.curselection())
        if not sel:
            return
        order = sel if d < 0 else sel[::-1]
        for i in order:
            j = i + d
            if 0 <= j < len(self.merge_paths):
                self.merge_paths[i], self.merge_paths[j] = self.merge_paths[j], self.merge_paths[i]
        self._refresh_mlist()
        for i in sel:
            if 0 <= i + d < len(self.merge_paths):
                self.merge_list.selection_set(i + d)

    def m_run(self):
        if len(self.merge_paths) < 2:
            messagebox.showinfo(APP_NAME, "Hãy thêm ít nhất 2 file.")
            return
        out = filedialog.asksaveasfilename(title="Lưu file gộp", defaultextension=".docx",
                                           initialdir=os.path.dirname(self.merge_paths[0]),
                                           initialfile="File_gop.docx", filetypes=[("Word", "*.docx")])
        if not out:
            return
        if os.path.abspath(out) in map(os.path.abspath, self.merge_paths):
            messagebox.showerror("Lỗi", "Không được ghi đè lên một file đang gộp.")
            return
        paths = list(self.merge_paths)
        opts = MergeOptions(new_page=self.m_newpage.get(), keep_headers=self.m_headers.get())

        def job():
            rep = merge_files(paths, out, opts, log=self.log)
            if rep.styles_renamed:
                self.log("Style đổi tên do khác định dạng:")
                for s in rep.styles_renamed:
                    self.log("   " + s)
            for wmsg in rep.warnings:
                self.log("⚠ " + wmsg)
            self.q.put(("done", ("Xong", f"Đã gộp {rep.files} file.\n"
                                         f"Style dùng chung: {len(rep.styles_shared)} • đổi tên: {len(rep.styles_renamed)}", out)))
        self.run_bg(job)

    # ------------------------------------------------------------------ TÁCH
    def _build_split(self):
        f = self.split_tab
        r1 = ttk.Frame(f)
        r1.pack(fill="x")
        ttk.Label(r1, text="File cần tách:").pack(side="left")
        self.s_path = tk.StringVar()
        ttk.Entry(r1, textvariable=self.s_path).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(r1, text="Chọn…", command=self.s_pick).pack(side="left")

        opt = ttk.LabelFrame(f, text="Cách tách", padding=6)
        opt.pack(fill="x", pady=8)
        self.s_mode = tk.StringVar(value="heading")
        g = ttk.Frame(opt)
        g.pack(fill="x")
        for i, (k, v) in enumerate(MODES.items()):
            ttk.Radiobutton(g, text=v, value=k, variable=self.s_mode,
                            command=self.s_mode_changed).grid(row=i // 2, column=i % 2, sticky="w", padx=4, pady=1)
        r = ttk.Frame(opt)
        r.pack(fill="x", pady=(6, 0))
        ttk.Label(r, text="Cấp đề mục:").pack(side="left")
        self.s_level = tk.IntVar(value=1)
        self.s_level_box = ttk.Spinbox(r, from_=1, to=9, width=4, textvariable=self.s_level)
        self.s_level_box.pack(side="left", padx=(4, 16))
        self.s_group_label = ttk.Label(r, text="Số phần mỗi file:")
        self.s_group_label.pack(side="left")
        self.s_group = tk.IntVar(value=1)
        ttk.Spinbox(r, from_=1, to=9999, width=6, textvariable=self.s_group).pack(side="left", padx=(4, 16))
        r2 = ttk.Frame(opt)
        r2.pack(fill="x", pady=(6, 0))
        self.s_range_label = ttk.Label(r2, text="Hoặc chọn phần (vd 1-3, 5, 7-9):")
        self.s_range_label.pack(side="left")
        self.s_ranges = tk.StringVar()
        ttk.Entry(r2, textvariable=self.s_ranges, width=28).pack(side="left", padx=4)
        ttk.Label(r2, text="(mỗi khoảng = 1 file)", foreground="#555").pack(side="left")
        r3 = ttk.Frame(opt)
        r3.pack(fill="x", pady=(6, 0))
        ttk.Label(r3, text="Tên file:").pack(side="left")
        self.s_pattern = tk.StringVar(value="{so} - {tieude}")
        ttk.Entry(r3, textvariable=self.s_pattern, width=28).pack(side="left", padx=4)
        ttk.Label(r3, text="{so} số thứ tự · {tieude} tiêu đề · {ten} tên file gốc · {trang} số trang",
                  foreground="#555").pack(side="left")
        self.s_keepnum = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="Giữ số thứ tự danh sách/đề mục như file gốc (không đánh lại từ 1)",
                        variable=self.s_keepnum).pack(anchor="w", pady=(6, 0))

        pv = ttk.LabelFrame(f, text="Xem trước các phần", padding=6)
        pv.pack(fill="both", expand=True)
        cols = ("no", "title", "pages")
        self.s_tree = ttk.Treeview(pv, columns=cols, show="headings", height=7)
        self.s_tree.heading("no", text="Phần")
        self.s_tree.heading("title", text="Bắt đầu bằng")
        self.s_tree.heading("pages", text="Trang")
        self.s_tree.column("no", width=60, anchor="center", stretch=False)
        self.s_tree.column("pages", width=90, anchor="center", stretch=False)
        self.s_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(pv, command=self.s_tree.yview)
        sb.pack(side="left", fill="y")
        self.s_tree.configure(yscrollcommand=sb.set)

        b = ttk.Frame(f)
        b.pack(fill="x", pady=(8, 0))
        ttk.Button(b, text="Xem trước", command=self.s_preview).pack(side="left")
        ttk.Button(b, text="Tách và lưu…", command=self.s_run).pack(side="right")
        self.s_mode_changed()

    def s_mode_changed(self):
        m = self.s_mode.get()
        self.s_level_box.configure(state="normal" if m == "heading" else "disabled")
        if m == "page":
            self.s_group_label.configure(text="Số trang mỗi file:")
            self.s_range_label.configure(text="Hoặc khoảng trang (vd 1-3, 5, 7-9):")
        else:
            self.s_group_label.configure(text="Số phần mỗi file:")
            self.s_range_label.configure(text="Hoặc chọn phần (vd 1-3, 5, 7-9):")

    def s_pick(self):
        p = filedialog.askopenfilename(title="Chọn file Word", filetypes=[("Word", "*.docx *.docm"), ("Tất cả", "*")])
        if p:
            self.s_path.set(p)
            self.s_preview()

    def _sopts(self):
        try:
            level, group = int(self.s_level.get()), int(self.s_group.get())
        except (tk.TclError, ValueError):
            raise DocxError("Cấp đề mục và số phần phải là số.")
        return SplitOptions(mode=self.s_mode.get(), level=level, group=group,
                            ranges=self.s_ranges.get(), name_pattern=self.s_pattern.get() or "{so} - {tieude}",
                            keep_numbering=self.s_keepnum.get())

    def s_preview(self):
        path = self.s_path.get()
        if not os.path.isfile(path):
            messagebox.showinfo(APP_NAME, "Hãy chọn file cần tách.")
            return
        try:
            prep, units = analyze(path, self._sopts())
        except DocxError as e:
            messagebox.showerror("Lỗi", str(e))
            return
        self.s_tree.delete(*self.s_tree.get_children())
        for u in units:
            pages = f"{u.first_page}–{u.last_page}" if u.first_page else ""
            self.s_tree.insert("", "end", values=(u.no, u.title, pages))
        self.log(f"{os.path.basename(path)}: {len(units)} phần.")
        for wmsg in prep.warnings:
            self.log("⚠ " + wmsg)

    def s_run(self):
        path = self.s_path.get()
        if not os.path.isfile(path):
            messagebox.showinfo(APP_NAME, "Hãy chọn file cần tách.")
            return
        try:
            opts = self._sopts()
        except DocxError as e:
            messagebox.showerror("Lỗi", str(e))
            return
        base = os.path.splitext(os.path.basename(path))[0]
        parent = filedialog.askdirectory(title="Chọn thư mục lưu kết quả", initialdir=os.path.dirname(path))
        if not parent:
            return
        out_dir = os.path.join(parent, f"{base} (đã tách)")

        def job():
            res = split_file(path, out_dir, opts, log=self.log)
            for wmsg in res.warnings:
                self.log("⚠ " + wmsg)
            self.q.put(("done", ("Xong", f"Đã tạo {len(res.files)} file trong:\n{out_dir}", out_dir)))
        self.run_bg(job)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

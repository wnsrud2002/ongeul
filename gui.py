#!/usr/bin/env python3
"""온글 — 문서를 Markdown으로 (창 화면).

명령줄(mdmaker.py)과 같은 변환·검사 경로(run_batch)를 그대로 쓴다.

표준 라이브러리 tkinter만 사용한다. 변환은 작업 스레드에서 돌리고 결과만
큐로 받아 화면을 갱신한다(변환 중에도 창이 멈추지 않게).
"""
from __future__ import annotations

import argparse
import queue
import subprocess
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import mdmaker as M

STATUS_TEXT = {
    M.SUCCESS: "완전 보존 검증 통과",
    M.PARTIAL: "누락·변형 확인됨",
    M.UNVERIFIED: "보존 확인 불가",
    M.FAILED: "변환 실패",
    M.UNSUPPORTED: "미지원 형식",
    M.SKIPPED: "이전 결과 재사용",
    "conflict": "기존 결과와 충돌",
}

COLORS = {
    "bg": "#F1F3F9", "surface": "#FFFFFF", "line": "#DDE1EA",
    "text": "#172033", "muted": "#667085", "brand": "#5946D2",
    "brand_dark": "#342887", "success": "#18794E", "warning": "#9A6700",
    "danger": "#C0362C",
}
OCR_VALUES = {"사용 안 함": "off", "자동": "auto", "항상": "force"}
PDF_VALUES = {"만들지 않음": "off", "변환 결과": "result", "원본 문서": "source", "둘 다": "both"}
XLSX_VALUES = {"자동": "auto", "Markdown 표": "md", "TSV 코드 블록": "tsv"}


def open_path(path: Path) -> None:
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], shell=False, check=False)
        elif sys.platform.startswith("win"):
            import os
            os.startfile(str(path))            # noqa: S606
        else:
            subprocess.run(["xdg-open", str(path)], shell=False, check=False)
    except OSError as exc:
        messagebox.showerror("열지 못했다", str(exc))


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.inputs: list[Path] = []
        self.q: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop = threading.Event()
        self.counts: dict = {}
        root.title("온글 — 문서를 안전하게 Markdown으로")
        root.geometry("1040x760")
        root.minsize(860, 640)
        root.configure(background=COLORS["bg"])
        self._build()
        modifier = "Command" if sys.platform == "darwin" else "Control"
        root.bind("<%s-o>" % modifier, lambda _e: self.pick_files())
        root.bind("<%s-Return>" % modifier, lambda _e: self.start())
        self.root.after(100, self._drain)

    # ---------------------------------------------------------------- 화면
    def _build(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["surface"], relief="solid", borderwidth=1)
        style.configure("Hero.TFrame", background=COLORS["brand_dark"])
        style.configure("Title.TLabel", background=COLORS["brand_dark"], foreground="white",
                        font=("TkDefaultFont", 24, "bold"))
        style.configure("HeroSubtitle.TLabel", background=COLORS["brand_dark"], foreground="#DAD6FF",
                        font=("TkDefaultFont", 11))
        style.configure("Subtitle.TLabel", background=COLORS["bg"], foreground=COLORS["muted"],
                        font=("TkDefaultFont", 11))
        style.configure("Section.TLabel", background=COLORS["surface"], foreground=COLORS["text"],
                        font=("TkDefaultFont", 13, "bold"))
        style.configure("Body.TLabel", background=COLORS["surface"], foreground=COLORS["text"])
        style.configure("Muted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])
        style.configure("Primary.TButton", background=COLORS["brand"], foreground="white",
                        borderwidth=0, padding=(18, 10), font=("TkDefaultFont", 11, "bold"))
        style.map("Primary.TButton", background=[("active", COLORS["brand_dark"]),
                                                  ("disabled", "#B8B4DE")])
        style.configure("Secondary.TButton", background="#EEF0F5", foreground=COLORS["text"],
                        borderwidth=0, padding=(14, 9))
        style.map("Secondary.TButton", background=[("active", "#E1E4EB")])
        style.configure("Stop.TButton", background="#FEECEB", foreground=COLORS["danger"],
                        borderwidth=0, padding=(14, 9))
        style.configure("TCheckbutton", background=COLORS["surface"], foreground=COLORS["text"])
        style.configure("TCombobox", padding=6)
        style.configure("TEntry", padding=7)
        style.configure("Results.Treeview", background="white", fieldbackground="white",
                        foreground=COLORS["text"], rowheight=34, borderwidth=0)
        style.configure("Results.Treeview.Heading", background="#F0F2F6",
                        foreground="#475467", relief="flat", padding=(8, 8),
                        font=("TkDefaultFont", 10, "bold"))
        style.map("Results.Treeview", background=[("selected", "#E9E7FF")],
                  foreground=[("selected", COLORS["text"])])
        style.configure("Brand.Horizontal.TProgressbar", troughcolor="#E9E7F8",
                        background=COLORS["brand"], borderwidth=0)

        shell = ttk.Frame(self.root, style="App.TFrame", padding=(30, 24, 30, 22))
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(4, weight=1)

        header = ttk.Frame(shell, style="Hero.TFrame", padding=(22, 18))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        brand = ttk.Frame(header, style="Hero.TFrame")
        brand.grid(row=0, column=0, sticky="w")
        ttk.Label(brand, text="온글", style="Title.TLabel").pack(anchor="w")
        ttk.Label(brand, text="문서를 빠짐없이 확인하고 Markdown으로 바꿉니다",
                  style="HeroSubtitle.TLabel").pack(anchor="w", pady=(3, 0))
        tk.Label(header, text="●  네트워크 전송 없음", bg="#FFFFFF", fg=COLORS["success"],
                 padx=12, pady=7, font=("TkDefaultFont", 10, "bold")).grid(
                     row=0, column=1, sticky="e")

        source = ttk.Frame(shell, style="Card.TFrame", padding=(20, 16))
        source.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        source.columnconfigure(2, weight=1)
        ttk.Label(source, text="1  문서 선택", style="Section.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))
        ttk.Button(source, text="파일 선택", style="Primary.TButton",
                   command=self.pick_files).grid(row=1, column=0, sticky="w")
        ttk.Button(source, text="폴더 선택", style="Secondary.TButton",
                   command=self.pick_dir).grid(row=1, column=1, sticky="w", padx=(8, 14))
        self.in_var = tk.StringVar(value="변환할 파일이나 폴더를 선택하세요")
        ttk.Label(source, textvariable=self.in_var, style="Body.TLabel").grid(
            row=1, column=2, columnspan=2, sticky="w")

        ttk.Label(source, text="저장 위치", style="Muted.TLabel").grid(
            row=2, column=0, sticky="w", pady=(14, 0))
        self.out_var = tk.StringVar(value="")
        self.out_var.trace_add("write", self._sync_actions)
        ttk.Entry(source, textvariable=self.out_var).grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=(8, 8), pady=(14, 0))
        ttk.Button(source, text="변경", style="Secondary.TButton", command=self.pick_out).grid(
            row=2, column=3, sticky="e", pady=(14, 0))

        opt = ttk.Frame(shell, style="Card.TFrame", padding=(20, 16))
        opt.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        for i in range(6):
            opt.columnconfigure(i, weight=1 if i in (1, 3, 5) else 0)
        ttk.Label(opt, text="2  변환 설정", style="Section.TLabel").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=False)
        self.reuse = tk.BooleanVar(value=False)
        self.make_md = tk.BooleanVar(value=True)
        checks = ttk.Frame(opt, style="Card.TFrame")
        checks.grid(row=1, column=0, columnspan=6, sticky="w", pady=(0, 12))
        ttk.Checkbutton(checks, text="하위 폴더 포함", variable=self.recursive).pack(side="left")
        ttk.Checkbutton(checks, text="기존 결과 덮어쓰기", variable=self.overwrite).pack(
            side="left", padx=(18, 0))
        ttk.Checkbutton(checks, text="검증된 결과 재사용", variable=self.reuse).pack(
            side="left", padx=(18, 0))
        ttk.Checkbutton(checks, text="Markdown 파일 남기기", variable=self.make_md).pack(
            side="left", padx=(18, 0))

        ttk.Label(opt, text="이미지 글자 읽기", style="Muted.TLabel").grid(row=2, column=0, sticky="w")
        self.ocr = tk.StringVar(value="사용 안 함")
        ttk.Combobox(opt, textvariable=self.ocr, values=tuple(OCR_VALUES), width=13,
                     state="readonly").grid(row=2, column=1, sticky="ew", padx=(8, 22))
        self.ocr_lang = tk.StringVar(value="kor+eng")
        ttk.Label(opt, text="OCR 언어", style="Muted.TLabel").grid(row=2, column=2, sticky="w")
        self.ocr_lang_entry = ttk.Entry(opt, textvariable=self.ocr_lang, width=10)
        self.ocr_lang_entry.grid(
            row=2, column=3, sticky="ew", padx=(8, 22))
        self.ocr.trace_add("write", self._sync_actions)

        ttk.Label(opt, text="분할 문자 수", style="Muted.TLabel").grid(row=2, column=4, sticky="w")
        self.split = tk.StringVar(value="0")
        ttk.Entry(opt, textvariable=self.split, width=10).grid(row=2, column=5, sticky="ew", padx=(8, 0))

        ttk.Label(opt, text="엑셀 표", style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        self.xlsx_table = tk.StringVar(value="자동")
        ttk.Combobox(opt, textvariable=self.xlsx_table, values=tuple(XLSX_VALUES), width=13,
                     state="readonly").grid(row=3, column=1, sticky="ew", padx=(8, 22), pady=(10, 0))
        ttk.Label(opt, text="PDF 출력", style="Muted.TLabel").grid(row=3, column=2, sticky="w", pady=(10, 0))
        self.pdf = tk.StringVar(value="만들지 않음")
        ttk.Combobox(opt, textvariable=self.pdf, values=tuple(PDF_VALUES), width=13,
                     state="readonly").grid(row=3, column=3, sticky="ew", padx=(8, 22), pady=(10, 0))
        ttk.Label(opt, text="PDF는 읽기용이며 Markdown 검증 상태에는 영향을 주지 않습니다",
                  style="Muted.TLabel").grid(row=3, column=4, columnspan=2, sticky="w", pady=(10, 0))

        run = ttk.Frame(shell, style="App.TFrame")
        run.grid(row=3, column=0, sticky="ew", pady=(2, 12))
        run.columnconfigure(3, weight=1)
        self.run_btn = ttk.Button(run, text="변환 시작", style="Primary.TButton",
                                  command=self.start, state="disabled")
        self.run_btn.grid(row=0, column=0, sticky="w")
        self.stop_btn = ttk.Button(run, text="중단", style="Stop.TButton",
                                   command=self.stop.set, state="disabled")
        self.stop_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.open_btn = ttk.Button(run, text="결과 폴더 열기", style="Secondary.TButton",
                                   command=lambda: open_path(Path(self.out_var.get() or ".")))
        self.open_btn.grid(row=0, column=2, sticky="w", padx=(8, 18))
        self.bar = ttk.Progressbar(run, mode="determinate", style="Brand.Horizontal.TProgressbar")
        self.bar.grid(row=0, column=3, sticky="ew", padx=(0, 12))
        self.progress_text = tk.StringVar(value="준비됨")
        ttk.Label(run, textvariable=self.progress_text, style="Subtitle.TLabel").grid(
            row=0, column=4, sticky="e")

        results = ttk.Frame(shell, style="Card.TFrame", padding=(20, 16))
        results.grid(row=4, column=0, sticky="nsew")
        results.columnconfigure(0, weight=1)
        results.rowconfigure(1, weight=1)
        head = ttk.Frame(results, style="Card.TFrame")
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        head.columnconfigure(0, weight=1)
        ttk.Label(head, text="3  변환 결과", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(head, text="행을 두 번 누르면 결과를 엽니다", style="Muted.TLabel").grid(
            row=0, column=1, sticky="e")
        cols = ("status", "meaning", "file", "note")
        self.tree = ttk.Treeview(results, columns=cols, show="headings", height=9,
                                 style="Results.Treeview")
        for c, t, w, stretch in (("status", "상태", 90, False), ("meaning", "검증 결과", 160, False),
                                 ("file", "파일", 250, True), ("note", "안내", 360, True)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, minwidth=70, anchor="w", stretch=stretch)
        self.tree.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", self._open_row)
        self.tree.tag_configure(M.SUCCESS, background="#ECFDF3", foreground=COLORS["success"])
        self.tree.tag_configure(M.SKIPPED, background="#EEF4FF", foreground="#2458A6")
        self.tree.tag_configure(M.UNVERIFIED, background="#FFF9E8", foreground=COLORS["warning"])
        self.tree.tag_configure(M.PARTIAL, background="#FFF4E8", foreground="#A04B00")
        self.tree.tag_configure(M.FAILED, background="#FEF0EF", foreground=COLORS["danger"])
        self.tree.tag_configure(M.UNSUPPORTED, background="#F3F4F6", foreground="#5F6673")
        self.tree.tag_configure("conflict", background="#FEF0EF", foreground=COLORS["danger"])
        self.rows: dict = {}
        self.empty = tk.Label(results, text="아직 변환한 문서가 없습니다\n위에서 파일이나 폴더를 선택해 시작하세요",
                              bg="white", fg=COLORS["muted"], justify="center",
                              font=("TkDefaultFont", 11), pady=22)
        self.empty.place(relx=0.5, rely=0.58, anchor="center")

        self.summary = tk.StringVar(
            value="대기 중 · 원본은 변경하지 않으며, 검증을 통과한 결과만 완료로 표시합니다.")
        self.summary_label = tk.Label(shell, textvariable=self.summary, anchor="w",
                                      bg="#ECEAFB", fg=COLORS["brand_dark"], padx=14, pady=10,
                                      font=("TkDefaultFont", 10, "bold"))
        self.summary_label.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        self._sync_actions()

    def _sync_actions(self, *_args) -> None:
        running = bool(self.worker and self.worker.is_alive())
        self.run_btn["state"] = "normal" if self.inputs and self.out_var.get() and not running else "disabled"
        self.ocr_lang_entry["state"] = "normal" if self.ocr.get() != "사용 안 함" else "disabled"

    # ---------------------------------------------------------------- 입력
    def pick_files(self) -> None:
        got = filedialog.askopenfilenames(title="변환할 파일")
        if got:
            self.inputs = [Path(g) for g in got]
            names = ", ".join(p.name for p in self.inputs[:3])
            tail = " 외 %d개" % (len(self.inputs) - 3) if len(self.inputs) > 3 else ""
            self.in_var.set("파일 %d개 · %s%s" % (len(self.inputs), names, tail))
            self._suggest_out(self.inputs[0].parent)
            self._sync_actions()

    def pick_dir(self) -> None:
        got = filedialog.askdirectory(title="변환할 폴더")
        if got:
            self.inputs = [Path(got)]
            self.in_var.set("폴더 · %s" % got)
            self._suggest_out(Path(got).parent)
            self._sync_actions()

    def pick_out(self) -> None:
        got = filedialog.askdirectory(title="결과를 저장할 폴더")
        if got:
            self.out_var.set(got)
            self._sync_actions()

    def _suggest_out(self, base: Path) -> None:
        if not self.out_var.get():
            self.out_var.set(str(base / "mdmaker_출력"))

    def _open_row(self, _event) -> None:
        sel = self.tree.selection()
        if sel and self.rows.get(sel[0]):
            open_path(self.rows[sel[0]])

    # ---------------------------------------------------------------- 실행
    def start(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        if not self.inputs or not self.out_var.get():
            messagebox.showwarning("입력 확인", "입력과 출력 폴더를 모두 고르세요.")
            return
        try:
            split = int(self.split.get() or 0)
        except ValueError:
            messagebox.showwarning("분할", "분할 문자 수는 숫자여야 한다.")
            return
        opts = argparse.Namespace(
            out=self.out_var.get(), recursive=self.recursive.get(),
            overwrite=self.overwrite.get(), reuse=self.reuse.get(), encoding=None,
            xlsx_table=XLSX_VALUES[self.xlsx_table.get()], max_cells=2_000_000,
            ocr=OCR_VALUES[self.ocr.get()], ocr_lang=self.ocr_lang.get(), split_chars=split,
            pdf=PDF_VALUES[self.pdf.get()], pdf_engine="builtin", pdf_font=None,
            no_md=not self.make_md.get())
        out_root = Path(opts.out)
        src = self.inputs[0]
        if len(self.inputs) == 1 and src.is_dir():
            root, files = src, M.collect(src, out_root, opts.recursive)
            single = False
        else:
            root, files = src.parent, [p for p in self.inputs if p.is_file()]
            single = True
        err = M.check_options(files, opts)
        if err:
            messagebox.showwarning("옵션", err)
            return
        if not files:
            messagebox.showwarning("입력", "처리할 파일이 없다.")
            return
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.rows.clear()
        self.counts = {}
        self.bar["maximum"] = len(files)
        self.bar["value"] = 0
        self.progress_text.set("0 / %d" % len(files))
        self.empty.place(relx=0.5, rely=0.58, anchor="center")
        self.stop.clear()
        self.run_btn["state"] = "disabled"
        self.stop_btn["state"] = "normal"
        self.summary.set("변환 중 · 원본은 그대로 두고 결과를 검사하고 있습니다.")
        self.summary_label.configure(bg="#ECEAFB", fg=COLORS["brand_dark"])
        self.worker = threading.Thread(
            target=self._work, args=(files, root, out_root, opts, single), daemon=True)
        self.worker.start()

    def _work(self, files, root, out_root, opts, single) -> None:
        limits = M.Limits(max_cells=opts.max_cells)
        try:
            for item in M.run_batch(files, root, out_root, opts, limits, single=single):
                self.q.put(item)
                if self.stop.is_set():
                    break
        except Exception as exc:                  # 창이 조용히 죽지 않게 한다
            self.q.put({"error": "%s: %s" % (type(exc).__name__, exc)})
        self.q.put({"done": True, "stopped": self.stop.is_set()})

    def _drain(self) -> None:
        try:
            while True:
                item = self.q.get_nowait()
                if item.get("done"):
                    self.stop_btn["state"] = "disabled"
                    if item.get("stopped"):
                        self.summary.set("중단됨 · 이미 변환한 결과는 저장되었습니다.")
                        self.progress_text.set("중단됨")
                        self.summary_label.configure(bg="#FFF4E8", fg="#A04B00")
                    else:
                        self.progress_text.set("완료")
                        self._summarize(final=True)
                    self._sync_actions()
                elif item.get("error"):
                    self.progress_text.set("오류")
                    messagebox.showerror("변환 중 오류", item["error"])
                else:
                    self._add(item)
        except queue.Empty:
            pass
        self.root.after(120, self._drain)

    def _add(self, item) -> None:
        st = item["status"]
        self.counts[st] = self.counts.get(st, 0) + 1
        note = "; ".join(item["warnings"][:2])
        iid = self.tree.insert("", "end", values=(st, STATUS_TEXT.get(st, st),
                                                    str(item["rel"]), note), tags=(st,))
        self.empty.place_forget()
        self.rows[iid] = item["target"]
        self.bar["value"] += 1
        self.progress_text.set("%d / %d" % (int(self.bar["value"]), int(self.bar["maximum"])))
        self._summarize()

    def _summarize(self, final: bool = False) -> None:
        parts = " · ".join("%s %d" % (STATUS_TEXT.get(k, k), v)
                           for k, v in sorted(self.counts.items()))
        done = M.exit_code(self.counts) == 0
        tail = ""
        if final:
            tail = ("  ·  모든 입력이 완전 보존 검증을 통과했다."
                    if done else
                    "  ·  success가 아닌 결과는 완료가 아니다. `_incomplete` 폴더와 경고를 확인할 것.")
            self.summary_label.configure(
                bg="#EAF8F1" if done else "#FFF4E8",
                fg=COLORS["success"] if done else "#A04B00")
        self.summary.set((parts or "결과 없음") + tail)


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if "--selftest" in argv:
        # 앱 묶음이 제대로 서는지 창을 띄우지 않고 확인한다.
        root = tk.Tk()
        root.withdraw()
        app = App(root)
        root.update_idletasks()
        ok = root.winfo_reqwidth() > 400 and len(app.tree["columns"]) == 4
        root.destroy()
        print("온글 자체 점검: %s (파이썬 %s)"
              % ("정상" if ok else "실패", sys.version.split()[0]))
        return 0 if ok else 1
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())

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
from tkinter import font as tkfont

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
    "bg": "#F6F7F9", "surface": "#FFFFFF", "line": "#EAECF0", "line_strong": "#D0D5DD",
    "hover": "#F2F4F7", "track": "#EEF0F4", "text": "#101828", "muted": "#475467",
    "faint": "#98A2B3", "brand": "#4F46E5", "brand_dark": "#4338CA", "brand_soft": "#EEF0FF",
    "success": "#067647", "success_soft": "#ECFDF3", "warning": "#B54708", "danger": "#D92D20",
}
OCR_VALUES = {"사용 안 함": "off", "자동": "auto", "항상": "force"}
PDF_VALUES = {"만들지 않음": "off", "변환 결과": "result", "원본 문서": "source", "둘 다": "both"}
XLSX_VALUES = {"자동": "auto", "Markdown 표": "md", "TSV 코드 블록": "tsv"}


def ui_font_family(root) -> str:
    """OS마다 기본 한글 UI 글꼴이 달라서, 있는 것 중 첫째를 쓴다."""
    have = set(tkfont.families(root))
    for fam in ("Apple SD Gothic Neo", "Malgun Gothic", "맑은 고딕", "Pretendard",
                "Noto Sans CJK KR", "Noto Sans KR", "NanumGothic"):
        if fam in have:
            return fam
    return tkfont.nametofont("TkDefaultFont").actual("family")


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
        root.geometry("1060x880")
        root.minsize(900, 760)
        root.configure(background=COLORS["bg"])
        self._build()
        modifier = "Command" if sys.platform == "darwin" else "Control"
        root.bind("<%s-o>" % modifier, lambda _e: self.pick_files())
        root.bind("<%s-Return>" % modifier, lambda _e: self.start())
        self.root.after(100, self._drain)

    # ---------------------------------------------------------------- 화면
    def _build(self) -> None:
        C = COLORS
        fam = ui_font_family(self.root)
        for name, size, weight in (("TkDefaultFont", 13, "normal"), ("TkTextFont", 13, "normal"),
                                   ("TkHeadingFont", 12, "bold"), ("TkMenuFont", 13, "normal")):
            try:
                tkfont.nametofont(name).configure(family=fam, size=size, weight=weight)
            except tk.TclError:
                pass
        F = lambda size, weight="normal": (fam, size, weight)   # noqa: E731
        self.root.option_add("*TCombobox*Listbox.font", F(13))
        self.root.option_add("*TCombobox*Listbox.selectBackground", C["brand_soft"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", C["text"])
        self.root.option_add("*TCombobox*Listbox.borderWidth", 0)

        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        # clam은 테두리를 bordercolor·lightcolor·darkcolor 세 색으로 입체 있게 그린다.
        # 셋을 같은 색으로 맞춰야 평평한 요즘 모양이 된다.
        flat = dict(bordercolor=C["line"], lightcolor=C["surface"], darkcolor=C["surface"])
        style.configure(".", background=C["surface"], foreground=C["text"], font=F(13),
                        focuscolor=C["surface"], **flat)
        style.configure("App.TFrame", background=C["bg"])
        style.configure("Card.TFrame", background=C["surface"])
        style.configure("Title.TLabel", background=C["bg"], foreground=C["text"], font=F(22, "bold"))
        style.configure("Subtitle.TLabel", background=C["bg"], foreground=C["muted"], font=F(12))
        style.configure("Section.TLabel", background=C["surface"], foreground=C["text"],
                        font=F(14, "bold"))
        style.configure("Step.TLabel", background=C["surface"], foreground=C["brand"],
                        font=F(14, "bold"))
        style.configure("Body.TLabel", background=C["surface"], foreground=C["text"])
        style.configure("Muted.TLabel", background=C["surface"], foreground=C["muted"], font=F(12))
        style.configure("Hint.TLabel", background=C["surface"], foreground=C["faint"], font=F(11))

        def button(name, bg, fg, hover, border=None, disabled_bg=None, bold=False):
            style.configure(name, background=bg, foreground=fg, bordercolor=border or bg,
                            lightcolor=bg, darkcolor=bg, focusthickness=0, relief="flat",
                            padding=(18, 9), font=F(13, "bold" if bold else "normal"))
            style.map(name,
                      background=[("disabled", disabled_bg or bg), ("pressed", hover),
                                  ("active", hover)],
                      lightcolor=[("disabled", disabled_bg or bg), ("active", hover)],
                      darkcolor=[("disabled", disabled_bg or bg), ("active", hover)],
                      bordercolor=[("disabled", disabled_bg or border or bg),
                                   ("active", border or hover)],
                      foreground=[("disabled", C["faint"] if not bold else "#FFFFFF")])
        button("Primary.TButton", C["brand"], "#FFFFFF", C["brand_dark"],
               disabled_bg="#C7C4F4", bold=True)
        button("Secondary.TButton", C["hover"], C["text"], "#E4E7EC", disabled_bg=C["hover"])
        button("Stop.TButton", "#FEF3F2", C["danger"], "#FEE4E2", disabled_bg=C["hover"])

        field = dict(fieldbackground=C["surface"], padding=(10, 7), insertcolor=C["text"],
                     bordercolor=C["line_strong"], lightcolor=C["surface"], darkcolor=C["surface"])
        focus = dict(bordercolor=[("focus", C["brand"])], lightcolor=[("focus", C["brand"])])
        style.configure("TEntry", **field)
        style.map("TEntry", fieldbackground=[("disabled", C["bg"])],
                  foreground=[("disabled", C["faint"])], **focus)
        # 테마 화살표는 입체 상자가 붙어 나와서, 작은 ⌄ 그림을 입력칸 테두리 안에 넣는다.
        self._chevron = tk.PhotoImage(width=10, height=6)
        for x in range(10):
            y = x if x < 5 else 9 - x
            self._chevron.put(C["muted"], (x, y))
            if 0 < x < 9:
                self._chevron.put(C["muted"], (x, y - 1 if x < 5 else y - 1))
        if "Flat.Combobox.downarrow" not in style.element_names():
            style.element_create("Flat.Combobox.downarrow", "image", self._chevron,
                                 sticky="", border=0)
        # clam의 콤보 필드는 오른쪽 선을 화살표에 맡겨서, 네 변을 다 그리는 입력칸 테두리를 빌린다.
        style.layout("TCombobox", [("Entry.field", {"sticky": "nswe", "children": [
            ("Flat.Combobox.downarrow", {"side": "right", "sticky": "ns"}),
            ("Combobox.padding", {"sticky": "nswe", "children": [
                ("Combobox.textarea", {"sticky": "nswe"})]})]})])
        style.configure("TCombobox", **dict(field, padding=(10, 7, 12, 7)))
        style.map("TCombobox", fieldbackground=[("readonly", C["surface"])],
                  selectbackground=[("readonly", C["surface"])],
                  selectforeground=[("readonly", C["text"])],
                  **focus)

        style.configure("Results.Treeview", background=C["surface"], fieldbackground=C["surface"],
                        foreground=C["text"], rowheight=38, borderwidth=0, font=F(12), **flat)
        style.layout("Results.Treeview", [("Results.Treeview.treearea", {"sticky": "nswe"})])
        style.configure("Results.Treeview.Heading", background=C["surface"], foreground=C["muted"],
                        relief="flat", padding=(10, 8), font=F(11, "bold"),
                        bordercolor=C["line"], lightcolor=C["surface"], darkcolor=C["line"])
        style.map("Results.Treeview.Heading", background=[("active", C["surface"])])
        style.map("Results.Treeview", background=[("selected", C["brand_soft"])],
                  foreground=[("selected", C["text"])])
        style.configure("Brand.Horizontal.TProgressbar", troughcolor=C["track"],
                        background=C["brand"], thickness=6, arrowsize=6, borderwidth=0,
                        bordercolor=C["track"], lightcolor=C["brand"], darkcolor=C["brand"])
        # 화살표 없는 얇은 스크롤바
        style.layout("Slim.Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {
            "sticky": "ns", "children": [("Vertical.Scrollbar.thumb",
                                          {"expand": "1", "sticky": "nswe"})]})])
        style.configure("Slim.Vertical.TScrollbar", troughcolor=C["surface"], background="#D0D5DD",
                        bordercolor=C["surface"], lightcolor="#D0D5DD", darkcolor="#D0D5DD",
                        gripcount=0, arrowsize=8, width=8)
        style.map("Slim.Vertical.TScrollbar", background=[("active", "#98A2B3")])

        def card(row, weight=False):
            # ttk 프레임의 solid 테두리는 안쪽 자식 프레임까지 선을 그어서, 바깥 한 겹만
            # 1px 선으로 두르는 tk 프레임을 쓴다.
            outer = tk.Frame(shell, bg=C["surface"], highlightthickness=1,
                             highlightbackground=C["line"], highlightcolor=C["line"])
            outer.grid(row=row, column=0, sticky="nsew" if weight else "ew", pady=(0, 14))
            inner = ttk.Frame(outer, style="Card.TFrame", padding=(22, 18))
            inner.pack(fill="both", expand=True)
            return inner

        def section(parent, step, title, cols):
            box = ttk.Frame(parent, style="Card.TFrame")
            box.grid(row=0, column=0, columnspan=cols, sticky="ew", pady=(0, 14))
            box.columnconfigure(2, weight=1)
            ttk.Label(box, text=step, style="Step.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(box, text=title, style="Section.TLabel").grid(
                row=0, column=1, sticky="w", padx=(8, 0))
            return box

        shell = ttk.Frame(self.root, style="App.TFrame", padding=(32, 26, 32, 22))
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(4, weight=1)

        header = ttk.Frame(shell, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        header.columnconfigure(1, weight=1)
        tk.Label(header, text="온", bg=C["brand"], fg="#FFFFFF", width=2,
                 font=F(18, "bold"), padx=6, pady=2).grid(row=0, column=0, rowspan=2,
                                                          sticky="w", padx=(0, 14))
        ttk.Label(header, text="온글", style="Title.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(header, text="문서를 빠짐없이 확인하고 Markdown으로 바꿉니다",
                  style="Subtitle.TLabel").grid(row=1, column=1, sticky="nw")
        tk.Label(header, text="●  오프라인 · 외부 전송 없음", bg=C["success_soft"],
                 fg=C["success"], padx=12, pady=6, font=F(11, "bold")).grid(
                     row=0, column=2, rowspan=2, sticky="e")

        source = card(1)
        source.columnconfigure(2, weight=1)
        section(source, "01", "문서 선택", 4)
        ttk.Button(source, text="파일 선택", style="Primary.TButton",
                   command=self.pick_files).grid(row=1, column=0, sticky="w")
        ttk.Button(source, text="폴더 선택", style="Secondary.TButton",
                   command=self.pick_dir).grid(row=1, column=1, sticky="w", padx=(8, 16))
        self.in_var = tk.StringVar(value="변환할 파일이나 폴더를 선택하세요")
        ttk.Label(source, textvariable=self.in_var, style="Muted.TLabel").grid(
            row=1, column=2, columnspan=2, sticky="w")

        ttk.Label(source, text="저장 위치", style="Muted.TLabel").grid(
            row=2, column=0, sticky="w", pady=(16, 0))
        self.out_var = tk.StringVar(value="")
        self.out_var.trace_add("write", self._sync_actions)
        ttk.Entry(source, textvariable=self.out_var).grid(
            row=2, column=1, columnspan=2, sticky="ew", padx=(8, 8), pady=(16, 0))
        ttk.Button(source, text="변경", style="Secondary.TButton", command=self.pick_out).grid(
            row=2, column=3, sticky="e", pady=(16, 0))

        opt = card(2)
        for i in range(6):
            opt.columnconfigure(i, weight=1 if i in (1, 3, 5) else 0)
        section(opt, "02", "변환 설정", 6)
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=False)
        self.reuse = tk.BooleanVar(value=False)
        self.make_md = tk.BooleanVar(value=True)
        checks = ttk.Frame(opt, style="Card.TFrame")
        checks.grid(row=1, column=0, columnspan=6, sticky="w", pady=(0, 16))
        for i, (text, var) in enumerate((("하위 폴더 포함", self.recursive),
                                         ("기존 결과 덮어쓰기", self.overwrite),
                                         ("검증된 결과 재사용", self.reuse),
                                         ("Markdown 파일 남기기", self.make_md))):
            self._check(checks, text, var, F).pack(side="left", padx=(0 if i == 0 else 26, 0))

        ttk.Label(opt, text="이미지 글자 읽기", style="Muted.TLabel").grid(row=2, column=0, sticky="w")
        self.ocr = tk.StringVar(value="사용 안 함")
        ttk.Combobox(opt, textvariable=self.ocr, values=tuple(OCR_VALUES), width=13,
                     state="readonly").grid(row=2, column=1, sticky="ew", padx=(10, 24))
        self.ocr_lang = tk.StringVar(value="kor+eng")
        ttk.Label(opt, text="OCR 언어", style="Muted.TLabel").grid(row=2, column=2, sticky="w")
        self.ocr_lang_entry = ttk.Entry(opt, textvariable=self.ocr_lang, width=10)
        self.ocr_lang_entry.grid(row=2, column=3, sticky="ew", padx=(10, 24))
        self.ocr.trace_add("write", self._sync_actions)

        ttk.Label(opt, text="분할 문자 수", style="Muted.TLabel").grid(row=2, column=4, sticky="w")
        self.split = tk.StringVar(value="0")
        ttk.Entry(opt, textvariable=self.split, width=10).grid(row=2, column=5, sticky="ew", padx=(10, 0))

        ttk.Label(opt, text="엑셀 표", style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.xlsx_table = tk.StringVar(value="자동")
        ttk.Combobox(opt, textvariable=self.xlsx_table, values=tuple(XLSX_VALUES), width=13,
                     state="readonly").grid(row=3, column=1, sticky="ew", padx=(10, 24), pady=(12, 0))
        ttk.Label(opt, text="PDF 출력", style="Muted.TLabel").grid(row=3, column=2, sticky="w", pady=(12, 0))
        self.pdf = tk.StringVar(value="만들지 않음")
        ttk.Combobox(opt, textvariable=self.pdf, values=tuple(PDF_VALUES), width=13,
                     state="readonly").grid(row=3, column=3, sticky="ew", padx=(10, 24), pady=(12, 0))
        ttk.Label(opt, text="PDF는 읽기용이며 검증 상태에는 영향을 주지 않습니다",
                  style="Hint.TLabel").grid(row=3, column=4, columnspan=2, sticky="w", pady=(12, 0))

        run = ttk.Frame(shell, style="App.TFrame")
        run.grid(row=3, column=0, sticky="ew", pady=(0, 14))
        run.columnconfigure(3, weight=1)
        self.run_btn = ttk.Button(run, text="변환 시작", style="Primary.TButton",
                                  command=self.start, state="disabled")
        self.run_btn.grid(row=0, column=0, sticky="w")
        self.stop_btn = ttk.Button(run, text="중단", style="Stop.TButton",
                                   command=self.stop.set, state="disabled")
        self.stop_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.open_btn = ttk.Button(run, text="결과 폴더 열기", style="Secondary.TButton",
                                   command=lambda: open_path(Path(self.out_var.get() or ".")))
        self.open_btn.grid(row=0, column=2, sticky="w", padx=(8, 20))
        self.bar = ttk.Progressbar(run, mode="determinate", style="Brand.Horizontal.TProgressbar")
        self.bar.grid(row=0, column=3, sticky="ew", padx=(0, 14))
        self.progress_text = tk.StringVar(value="준비됨")
        ttk.Label(run, textvariable=self.progress_text, style="Subtitle.TLabel").grid(
            row=0, column=4, sticky="e")

        results = card(4, weight=True)
        results.columnconfigure(0, weight=1)
        results.rowconfigure(1, weight=1)
        head = section(results, "03", "변환 결과", 2)
        ttk.Label(head, text="행을 두 번 누르면 결과를 엽니다", style="Hint.TLabel").grid(
            row=0, column=2, sticky="e")
        cols = ("status", "meaning", "file", "note")
        self.tree = ttk.Treeview(results, columns=cols, show="headings", height=9,
                                 style="Results.Treeview")
        for c, t, w, stretch in (("status", "상태", 100, False), ("meaning", "검증 결과", 170, False),
                                 ("file", "파일", 250, True), ("note", "안내", 360, True)):
            self.tree.heading(c, text=t, anchor="w")
            self.tree.column(c, width=w, minwidth=70, anchor="w", stretch=stretch)
        self.tree.grid(row=1, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview,
                               style="Slim.Vertical.TScrollbar")
        scroll.grid(row=1, column=1, sticky="ns", padx=(6, 0))
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", self._open_row)
        # 행 전체를 칠하면 무거워 보여서 글자색으로만 상태를 구분한다.
        for tag, color in ((M.SUCCESS, C["success"]), (M.SKIPPED, "#175CD3"),
                           (M.UNVERIFIED, C["warning"]), (M.PARTIAL, "#B54708"),
                           (M.FAILED, C["danger"]), (M.UNSUPPORTED, C["muted"]),
                           ("conflict", C["danger"])):
            self.tree.tag_configure(tag, foreground=color)
        self.rows: dict = {}
        self.empty = tk.Label(results, text="아직 변환한 문서가 없습니다\n위에서 파일이나 폴더를 선택해 시작하세요",
                              bg=C["surface"], fg=C["faint"], justify="center",
                              font=F(12), pady=22)
        self.empty.place(relx=0.5, rely=0.58, anchor="center")

        self.summary = tk.StringVar(
            value="대기 중 · 원본은 변경하지 않으며, 검증을 통과한 결과만 완료로 표시합니다.")
        self.summary_label = tk.Label(shell, textvariable=self.summary, anchor="w",
                                      bg=C["bg"], fg=C["muted"], padx=2, pady=4, font=F(12))
        self.summary_label.grid(row=5, column=0, sticky="ew")
        self._sync_actions()

    def _check(self, parent, text: str, var: tk.BooleanVar, F) -> tk.Frame:
        """clam 체크박스는 ✕ 표시밖에 못 그린다. 글자로 그리면 레티나에서도 선명하다."""
        C = COLORS
        row = tk.Frame(parent, bg=C["surface"], cursor="hand2", takefocus=1,
                       highlightthickness=0)
        # macOS에서는 Label의 highlight 테두리가 그려지지 않아, 1px 틈을 둔 바깥 프레임을 선으로 쓴다.
        edge = tk.Frame(row, bd=0, padx=1, pady=1)
        box = tk.Label(edge, width=2, font=F(10, "bold"), bd=0)
        box.pack()
        label = tk.Label(row, text=text, bg=C["surface"], fg=C["text"], font=F(13))
        edge.pack(side="left")
        label.pack(side="left", padx=(8, 0))

        def paint(*_):
            on = var.get()
            box.configure(text="✓" if on else "", fg="#FFFFFF",
                          bg=C["brand"] if on else C["surface"])
            edge.configure(bg=C["brand"] if on else C["line_strong"])

        def toggle(_e=None):
            var.set(not var.get())
        for w in (row, edge, box, label):
            w.bind("<Button-1>", toggle)
        row.bind("<space>", toggle)             # 키보드로도 켜고 끌 수 있게
        row.bind("<FocusIn>", lambda _e: label.configure(fg=C["brand"]))
        row.bind("<FocusOut>", lambda _e: label.configure(fg=C["text"]))
        var.trace_add("write", paint)
        paint()
        return row

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
        self.summary_label.configure(fg=COLORS["brand"])
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
                        self.summary_label.configure(fg=COLORS["warning"])
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
            self.summary_label.configure(fg=COLORS["success"] if done else COLORS["warning"])
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

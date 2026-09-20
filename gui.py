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
        root.title("온글 — 문서를 Markdown으로 · 로컬에서만 변환합니다")
        root.geometry("900x620")
        self._build()
        self.root.after(100, self._drain)

    # ---------------------------------------------------------------- 화면
    def _build(self) -> None:
        pad = {"padx": 6, "pady": 4}
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)

        ttk.Button(top, text="파일 고르기", command=self.pick_files).grid(row=0, column=0)
        ttk.Button(top, text="폴더 고르기", command=self.pick_dir).grid(row=0, column=1, padx=4)
        self.in_var = tk.StringVar(value="입력을 고르세요")
        ttk.Label(top, textvariable=self.in_var, width=70).grid(row=0, column=2, sticky="w")

        ttk.Button(top, text="출력 폴더", command=self.pick_out).grid(row=1, column=0, pady=4)
        self.out_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self.out_var, width=78).grid(row=1, column=1, columnspan=2,
                                                                sticky="we")

        opt = ttk.LabelFrame(self.root, text="옵션")
        opt.pack(fill="x", **pad)
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=False)
        self.reuse = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="하위 폴더까지", variable=self.recursive).grid(row=0, column=0)
        ttk.Checkbutton(opt, text="기존 결과 덮어쓰기", variable=self.overwrite).grid(row=0, column=1)
        ttk.Checkbutton(opt, text="이전 성공 결과 재사용", variable=self.reuse).grid(row=0, column=2)

        ttk.Label(opt, text="OCR").grid(row=0, column=3, padx=(16, 2))
        self.ocr = tk.StringVar(value="off")
        ttk.Combobox(opt, textvariable=self.ocr, values=("off", "auto", "force"), width=6,
                     state="readonly").grid(row=0, column=4)
        self.ocr_lang = tk.StringVar(value="kor+eng")
        ttk.Entry(opt, textvariable=self.ocr_lang, width=10).grid(row=0, column=5, padx=2)

        ttk.Label(opt, text="분할(문자)").grid(row=0, column=6, padx=(16, 2))
        self.split = tk.StringVar(value="0")
        ttk.Entry(opt, textvariable=self.split, width=8).grid(row=0, column=7)

        ttk.Label(opt, text="엑셀 표").grid(row=0, column=8, padx=(16, 2))
        self.xlsx_table = tk.StringVar(value="auto")
        ttk.Combobox(opt, textvariable=self.xlsx_table, values=("auto", "md", "tsv"), width=6,
                     state="readonly").grid(row=0, column=9)

        run = ttk.Frame(self.root)
        run.pack(fill="x", **pad)
        self.run_btn = ttk.Button(run, text="변환 시작", command=self.start)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(run, text="중단", command=self.stop.set, state="disabled")
        self.stop_btn.pack(side="left", padx=4)
        ttk.Button(run, text="결과 폴더 열기",
                   command=lambda: open_path(Path(self.out_var.get() or "."))).pack(side="left")
        self.bar = ttk.Progressbar(run, mode="determinate", length=320)
        self.bar.pack(side="right")

        cols = ("status", "meaning", "file", "note")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings", height=16)
        for c, t, w in (("status", "상태", 90), ("meaning", "뜻", 150),
                        ("file", "파일", 300), ("note", "경고", 330)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, **pad)
        self.tree.bind("<Double-1>", self._open_row)
        self.rows: dict = {}

        self.summary = tk.StringVar(
            value="success 만 완전 보존 검증을 통과한 결과다. 나머지는 원본을 대체할 수 없다.")
        ttk.Label(self.root, textvariable=self.summary, anchor="w").pack(fill="x", **pad)

    # ---------------------------------------------------------------- 입력
    def pick_files(self) -> None:
        got = filedialog.askopenfilenames(title="변환할 파일")
        if got:
            self.inputs = [Path(g) for g in got]
            self.in_var.set("파일 %d개" % len(self.inputs))
            self._suggest_out(self.inputs[0].parent)

    def pick_dir(self) -> None:
        got = filedialog.askdirectory(title="변환할 폴더")
        if got:
            self.inputs = [Path(got)]
            self.in_var.set(got)
            self._suggest_out(Path(got).parent)

    def pick_out(self) -> None:
        got = filedialog.askdirectory(title="결과를 저장할 폴더")
        if got:
            self.out_var.set(got)

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
            xlsx_table=self.xlsx_table.get(), max_cells=2_000_000, ocr=self.ocr.get(),
            ocr_lang=self.ocr_lang.get(), split_chars=split)
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
        self.stop.clear()
        self.run_btn["state"] = "disabled"
        self.stop_btn["state"] = "normal"
        self.worker = threading.Thread(
            target=self._work, args=(files, root, out_root, opts, single), daemon=True)
        self.worker.start()

    def _work(self, files, root, out_root, opts, single) -> None:
        limits = M.Limits(max_cells=opts.max_cells)
        try:
            for item in M.run_batch(files, root, out_root, opts, limits, single=single):
                self.q.put(item)
                if self.stop.is_set():
                    self.q.put({"stopped": True})
                    break
        except Exception as exc:                  # 창이 조용히 죽지 않게 한다
            self.q.put({"error": "%s: %s" % (type(exc).__name__, exc)})
        self.q.put({"done": True})

    def _drain(self) -> None:
        try:
            while True:
                item = self.q.get_nowait()
                if item.get("done"):
                    self.run_btn["state"] = "normal"
                    self.stop_btn["state"] = "disabled"
                    self._summarize(final=True)
                elif item.get("stopped"):
                    self.summary.set("중단했다. 여기까지의 결과만 저장돼 있다.")
                elif item.get("error"):
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
                                                  str(item["rel"]), note))
        self.tree.item(iid, tags=("ok",) if st == M.SUCCESS else ("check",))
        self.tree.tag_configure("check", background="#fff6e0")
        self.rows[iid] = item["target"]
        self.bar["value"] += 1
        self._summarize()

    def _summarize(self, final: bool = False) -> None:
        parts = ", ".join("%s %d" % (k, v) for k, v in sorted(self.counts.items()))
        done = M.exit_code(self.counts) == 0
        tail = ""
        if final:
            tail = ("  ·  모든 입력이 완전 보존 검증을 통과했다."
                    if done else
                    "  ·  success가 아닌 결과는 완료가 아니다. `_incomplete` 폴더와 경고를 확인할 것.")
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

"""창(GUI) 실행 화면.

1) 시작 창: 어휘 등급 파일·지문 파일을 자동으로 찾아 채워 두고, 바꾸고 싶으면 [찾아보기]로 고른다.
2) 분석이 끝나면 질적평가 입력 창: 지문마다 평가자 3명의 -3~+3 점수를 입력한다.
   빈칸은 '미입력'으로 처리하며, 나중에 결과 엑셀에 적고 --recalc로 다시 계산할 수도 있다.
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .qualitative import CRITERIA, pairwise_agreement, parse_score

FONT = ("맑은 고딕", 10)
FONT_B = ("맑은 고딕", 10, "bold")


def _open_file(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


class Launcher:
    def __init__(self, root, cfg, args):
        from .cli import default_output_path, search_dirs
        from .passages import find_passage_files
        from .vocab import find_vocab_file

        self.root, self.cfg = root, cfg
        self.queue = queue.Queue()
        root.title("ERI 텍스트 복잡도 계산기")
        root.geometry("760x560")
        root.minsize(640, 480)

        dirs = search_dirs()
        vocab = args.vocab or (find_vocab_file(dirs) or "")
        inputs = [str(p) for p in (args.inputs or find_passage_files(dirs))]

        frm = ttk.Frame(root, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="어휘 등급 파일", font=FONT_B).grid(row=0, column=0, sticky="w", pady=4)
        self.vocab_var = tk.StringVar(value=str(vocab))
        ttk.Entry(frm, textvariable=self.vocab_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="찾아보기", command=self.pick_vocab).grid(row=0, column=2)

        ttk.Label(frm, text="지문 파일", font=FONT_B).grid(row=1, column=0, sticky="nw", pady=4)
        self.listbox = tk.Listbox(frm, height=6, selectmode="extended", font=FONT)
        self.listbox.grid(row=1, column=1, sticky="nsew", padx=6)
        for p in inputs:
            self.listbox.insert("end", p)
        btns = ttk.Frame(frm)
        btns.grid(row=1, column=2, sticky="n")
        ttk.Button(btns, text="파일 추가", command=self.add_files).pack(fill="x")
        ttk.Button(btns, text="폴더 추가", command=self.add_folder).pack(fill="x", pady=2)
        ttk.Button(btns, text="선택 제거", command=self.remove_selected).pack(fill="x")
        ttk.Label(frm, text="지원 형식: 텍스트(.txt, 인코딩 자동), 엑셀/CSV(제목·본문 열), 워드(.docx), 폴더",
                  foreground="#555").grid(row=2, column=1, sticky="w", padx=6)

        ttk.Label(frm, text="기본 학교급", font=FONT_B).grid(row=3, column=0, sticky="w", pady=(10, 4))
        lv = ttk.Frame(frm)
        lv.grid(row=3, column=1, sticky="w", padx=6, pady=(10, 4))
        self.level_var = tk.StringVar(value=cfg.default_level)
        ttk.Radiobutton(lv, text="초등 (식1)", value="초등", variable=self.level_var).pack(side="left")
        ttk.Radiobutton(lv, text="중등 (식2)", value="중등", variable=self.level_var).pack(side="left", padx=10)
        ttk.Label(lv, text="※ 제목에 [초등] / [중등] / [5학년] 표시가 있으면 그것을 따릅니다.",
                  foreground="#555").pack(side="left")

        ttk.Label(frm, text="결과 파일", font=FONT_B).grid(row=4, column=0, sticky="w", pady=4)
        self.out_var = tk.StringVar(value=str(args.out or default_output_path(inputs[0] if inputs else None)))
        ttk.Entry(frm, textvariable=self.out_var).grid(row=4, column=1, sticky="ew", padx=6)
        ttk.Button(frm, text="바꾸기", command=self.pick_out).grid(row=4, column=2)

        self.ask_qual = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="분석 후 질적평가 점수 입력 창 열기 (평가자 3명, -3 ~ +3)",
                        variable=self.ask_qual).grid(row=5, column=1, sticky="w", padx=6, pady=6)

        self.run_btn = ttk.Button(frm, text="분석 시작", command=self.start)
        self.run_btn.grid(row=6, column=1, sticky="ew", padx=6, pady=4)
        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=4)
        self.log = tk.Text(frm, height=10, font=("Consolas", 9), state="disabled", bg="#f7f7f7")
        self.log.grid(row=8, column=0, columnspan=3, sticky="nsew")
        frm.rowconfigure(8, weight=1)
        root.after(100, self.poll)

    # --- 파일 고르기 ---
    def pick_vocab(self):
        p = filedialog.askopenfilename(title="어휘 등급 파일", filetypes=[("엑셀/CSV", "*.xlsx *.xlsm *.csv"), ("모든 파일", "*.*")])
        if p:
            self.vocab_var.set(p)

    def add_files(self):
        for p in filedialog.askopenfilenames(title="지문 파일", filetypes=[
                ("지문 파일", "*.txt *.md *.xlsx *.csv *.docx"), ("모든 파일", "*.*")]):
            self.listbox.insert("end", p)

    def add_folder(self):
        p = filedialog.askdirectory(title="지문 폴더")
        if p:
            self.listbox.insert("end", p)

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title="결과 저장", defaultextension=".xlsx",
                                         initialfile=Path(self.out_var.get()).name, filetypes=[("엑셀", "*.xlsx")])
        if p:
            self.out_var.set(p)

    # --- 분석 ---
    def write(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        inputs = list(self.listbox.get(0, "end"))
        vocab = self.vocab_var.get().strip()
        if not vocab or not Path(vocab).is_file():
            messagebox.showerror("어휘 등급 파일", "어휘 등급 엑셀 파일을 선택하세요.")
            return
        if not inputs:
            messagebox.showerror("지문 파일", "분석할 지문 파일을 추가하세요.")
            return
        self.cfg.default_level = self.level_var.get()
        self.run_btn.configure(state="disabled")
        self.progress.configure(value=0)
        threading.Thread(target=self.worker, args=(inputs, vocab), daemon=True).start()

    def worker(self, inputs, vocab):
        from .cli import run_analysis
        try:
            res = run_analysis(inputs, vocab, self.cfg,
                               log=lambda m: self.queue.put(("log", m)),
                               progress=lambda i, n: self.queue.put(("progress", i / n * 100)))
            self.queue.put(("done", res))
        except Exception as e:
            self.queue.put(("error", e))

    def poll(self):
        try:
            while True:
                kind, val = self.queue.get_nowait()
                if kind == "log":
                    self.write(val)
                elif kind == "progress":
                    self.progress.configure(value=val)
                elif kind == "error":
                    self.run_btn.configure(state="normal")
                    messagebox.showerror("오류", str(val))
                elif kind == "done":
                    self.run_btn.configure(state="normal")
                    self.finish(*val)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def finish(self, results, vocab, analyzer, passages, failed):
        if not results:
            messagebox.showerror("오류", "분석된 지문이 없습니다.")
            return
        if failed:
            messagebox.showwarning("일부 지문 오류", "\n".join(f"{t}: {e}" for t, e in failed))
        if self.ask_qual.get():
            RatingWindow(self.root, results, analyzer, lambda: self.save(results, vocab))
        else:
            self.save(results, vocab)

    def save(self, results, vocab):
        from .report import write_report
        out = Path(self.out_var.get())
        try:
            write_report(results, self.cfg, out, vocab.source)
        except PermissionError:
            messagebox.showerror("저장 실패", f"{out.name} 파일이 열려 있으면 닫고 다시 시도하세요.")
            return False
        self.write(f"결과 저장: {out}")
        if messagebox.askyesno("완료", f"결과를 저장했습니다.\n{out}\n\n파일을 열까요?"):
            _open_file(out)
        return True


class RatingWindow:
    """질적평가 점수 입력 (지문 × 평가자)."""

    def __init__(self, master, results, analyzer, on_saved):
        self.results, self.analyzer, self.on_saved = results, analyzer, on_saved
        cfg = analyzer.cfg
        self.n = cfg.rater_count
        win = self.win = tk.Toplevel(master)
        win.title("질적평가 점수 입력")
        win.geometry("860x560")
        win.transient(master)
        win.grab_set()

        top = ttk.Frame(win, padding=(12, 10))
        top.pack(fill="x")
        ttk.Label(top, text="정량 지수에 대한 보정값을 평가자별로 입력하세요 (-3 ~ +3, 빈칸 = 미입력).",
                  font=FONT_B).pack(anchor="w")
        ttk.Label(top, text="평가 기준: " + ", ".join(CRITERIA), foreground="#555").pack(anchor="w")

        outer = ttk.Frame(win)
        outer.pack(fill="both", expand=True, padx=12)
        canvas = tk.Canvas(outer, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas)
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.canvas = canvas

        heads = ["지문명", "학교급", "정량 지수"] + [f"평가자{i + 1}" for i in range(self.n)] + ["평균", "ERI"]
        for c, h in enumerate(heads):
            ttk.Label(body, text=h, font=FONT_B).grid(row=0, column=c, padx=4, pady=4)
        self.entries, self.mean_labels, self.eri_labels = [], [], []
        for r, res in enumerate(results, 1):
            ttk.Label(body, text=res.title, width=34, anchor="w").grid(row=r, column=0, sticky="w", padx=4)
            ttk.Label(body, text=res.level).grid(row=r, column=1)
            ttk.Label(body, text=f"{res.quant_rounded:.1f}").grid(row=r, column=2)
            row = []
            for k in range(self.n):
                v = tk.StringVar()
                e = ttk.Entry(body, textvariable=v, width=7, justify="center")
                e.grid(row=r, column=3 + k, padx=3, pady=2)
                v.trace_add("write", lambda *_a, i=r - 1: self.update_row(i))
                e.bind("<Return>", lambda ev, i=r - 1, k=k: self.move(i + 1, k))
                e.bind("<Down>", lambda ev, i=r - 1, k=k: self.move(i + 1, k))
                e.bind("<Up>", lambda ev, i=r - 1, k=k: self.move(i - 1, k))
                row.append((v, e))
            self.entries.append(row)
            ml = ttk.Label(body, text="-", width=6)
            ml.grid(row=r, column=3 + self.n)
            el = ttk.Label(body, text=f"{res.eri:.1f}", width=6, foreground="#c00000", font=FONT_B)
            el.grid(row=r, column=4 + self.n)
            self.mean_labels.append(ml)
            self.eri_labels.append(el)

        bottom = ttk.Frame(win, padding=10)
        bottom.pack(fill="x")
        self.agree_label = ttk.Label(bottom, text="평가자 일치도: -")
        self.agree_label.pack(side="left")
        ttk.Button(bottom, text="저장", command=self.save).pack(side="right")
        ttk.Button(bottom, text="질적평가 없이 저장", command=self.skip).pack(side="right", padx=6)
        if self.entries:
            win.after(150, self.entries[0][0][1].focus_set)

    def move(self, i, k):
        if 0 <= i < len(self.entries):
            e = self.entries[i][k][1]
            e.focus_set()
            self.canvas.yview_moveto(max(0, (i - 3) / max(1, len(self.entries))))
        return "break"

    def _values(self, i, strict=False):
        cfg = self.analyzer.cfg
        vals = []
        for v, _ in self.entries[i]:
            try:
                vals.append(parse_score(v.get(), cfg.qual_min, cfg.qual_max))
            except ValueError:
                if strict:
                    raise
                vals.append(None)
        return vals

    def update_row(self, i):
        vals = self._values(i)
        res = self.analyzer.apply_qualitative(self.results[i], vals)
        self.mean_labels[i].configure(text="-" if res.qual_mean is None else f"{res.qual_mean:+.2f}")
        self.eri_labels[i].configure(text=f"{res.eri:.1f}")
        mean, _ = pairwise_agreement([[self._values(r)[k] for r in range(len(self.entries))] for k in range(self.n)])
        if mean is not None:
            ok = mean >= self.analyzer.cfg.agreement_threshold
            self.agree_label.configure(text=f"평가자 일치도: {mean:.2f} ({'기준 충족' if ok else '기준 0.9 미달'})",
                                       foreground="#060" if ok else "#c00000")

    def save(self):
        for i, res in enumerate(self.results):
            try:
                vals = self._values(i, strict=True)
            except ValueError as e:
                messagebox.showerror("입력 오류", f"'{res.title}': {e}\n숫자를 -3 ~ +3 범위로 입력하세요.", parent=self.win)
                return
            self.analyzer.apply_qualitative(res, vals)
        if self.on_saved() is not False:
            self.win.destroy()

    def skip(self):
        for res in self.results:
            self.analyzer.apply_qualitative(res, [])
        if self.on_saved() is not False:
            self.win.destroy()


def launch(cfg, args):
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform.startswith("win") else "clam")
    except tk.TclError:
        pass
    Launcher(root, cfg, args)
    root.mainloop()
    return 0

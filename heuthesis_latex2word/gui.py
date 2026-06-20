from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .converter import ConversionOptions, convert_project


class Latex2WordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("HEU LaTeX 转 Word")
        self.geometry("720x360")
        self.resizable(False, False)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.auth_var = tk.StringVar()
        self.degree_var = tk.StringVar(value="master")
        self.status_var = tk.StringVar(value="请选择 HEU main.tex")
        self.report_path: Path | None = None
        self.output_path: Path | None = None

        self._build()

    def _build(self) -> None:
        pad = {"padx": 14, "pady": 8}
        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=18, pady=18)

        ttk.Label(frame, text="主文件").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.input_var, width=72).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="选择", command=self._choose_input).grid(row=0, column=2, **pad)

        ttk.Label(frame, text="输出 Word").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.output_var, width=72).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="保存为", command=self._choose_output).grid(row=1, column=2, **pad)

        ttk.Label(frame, text="学位类型").grid(row=2, column=0, sticky="w", **pad)
        degree_box = ttk.Combobox(
            frame,
            textvariable=self.degree_var,
            values=("master", "doctor"),
            state="readonly",
            width=18,
        )
        degree_box.grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(frame, text="签字声明页").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.auth_var, width=72).grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="可选", command=self._choose_auth).grid(row=3, column=2, **pad)

        ttk.Label(frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=3, sticky="w", **pad)

        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=3, sticky="e", pady=(22, 0))
        ttk.Button(buttons, text="开始转换", command=self._start).pack(side="left", padx=8)
        ttk.Button(buttons, text="打开输出", command=self._open_output).pack(side="left", padx=8)
        ttk.Button(buttons, text="查看报告", command=self._open_report).pack(side="left", padx=8)

    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("LaTeX", "*.tex"), ("All files", "*.*")])
        if not path:
            return
        self.input_var.set(path)
        output = Path(path).with_suffix(".docx")
        self.output_var.set(str(output))

    def _choose_output(self) -> None:
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if path:
            self.output_var.set(path)

    def _choose_auth(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("PDF/Image", "*.pdf;*.png;*.jpg;*.jpeg"), ("All files", "*.*")])
        if path:
            self.auth_var.set(path)

    def _start(self) -> None:
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        if not input_path or not output_path:
            messagebox.showwarning("缺少文件", "请选择输入 main.tex 和输出 docx。")
            return

        def worker() -> None:
            try:
                self.status_var.set("正在转换，请稍候...")
                result = convert_project(
                    ConversionOptions(
                        input_file=Path(input_path),
                        output_file=Path(output_path),
                        degree=self.degree_var.get(),
                        include_auth_scan=Path(self.auth_var.get()) if self.auth_var.get().strip() else None,
                    )
                )
                self.output_path = result.output_file
                self.report_path = result.report_file
                self.status_var.set(f"完成：{result.output_file}")
                messagebox.showinfo("转换完成", f"已生成：\n{result.output_file}")
            except Exception as exc:
                self.status_var.set("转换失败")
                messagebox.showerror("转换失败", str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _open_output(self) -> None:
        if self.output_path and self.output_path.exists():
            os.startfile(self.output_path)

    def _open_report(self) -> None:
        if self.report_path and self.report_path.exists():
            os.startfile(self.report_path)


def main() -> None:
    Latex2WordApp().mainloop()


if __name__ == "__main__":
    main()

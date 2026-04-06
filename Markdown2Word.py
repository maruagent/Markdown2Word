import sys
import os
import tempfile
import pypandoc
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime


def resource_path(relative_path):
    """実行ファイル(.exe)化された時に、内包されたファイルのパスを取得する魔法の関数"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_templates():
    """templatesフォルダ内の.docxファイル一覧を取得"""
    templates_dir = resource_path("templates")
    templates = {}

    # style.docx（デフォルト）があれば追加
    default_style = resource_path("style.docx")
    if os.path.exists(default_style):
        templates["デフォルト (style.docx)"] = default_style

    # templatesフォルダがあればその中の.docxを追加
    if os.path.exists(templates_dir):
        for f in sorted(os.listdir(templates_dir)):
            if f.endswith(".docx"):
                name = os.path.splitext(f)[0]
                templates[name] = os.path.join(templates_dir, f)

    return templates


def convert(markdown_text, ref_file, out_file):
    """MarkdownテキストをWordに変換する"""
    # 一時的なMarkdownファイルを作成して変換
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".md", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(markdown_text)
        tmp_path = tmp.name

    try:
        if ref_file and os.path.exists(ref_file):
            pypandoc.convert_file(
                tmp_path,
                "docx",
                outputfile=out_file,
                extra_args=[
                    "-f",
                    "markdown-yaml_metadata_block",
                    f"--reference-doc={ref_file}",
                ],
            )
        else:
            pypandoc.convert_file(
                tmp_path,
                "docx",
                outputfile=out_file,
                extra_args=["-f", "markdown-yaml_metadata_block"],
            )
    finally:
        os.unlink(tmp_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Markdown → Word 変換")
        self.resizable(True, True)
        self.minsize(600, 500)
        self._build_ui()
        self._load_templates()

        # ドラッグ&ドロップまたはコマンドライン引数でファイルが渡された場合
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self._load_md_file(sys.argv[1])

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # ── Markdownテキストエリア ──────────────────────────
        frame_md = tk.LabelFrame(self, text="Markdownテキスト", **pad)
        frame_md.pack(fill="both", expand=True, **pad)

        btn_frame = tk.Frame(frame_md)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="クリップボードから貼り付け", command=self._paste_clipboard).pack(side="left", padx=4, pady=4)
        tk.Button(btn_frame, text="ファイルから読み込み", command=self._open_file).pack(side="left", padx=4, pady=4)
        tk.Button(btn_frame, text="クリア", command=self._clear_text).pack(side="left", padx=4, pady=4)

        text_frame = tk.Frame(frame_md)
        text_frame.pack(fill="both", expand=True)
        self.text_area = tk.Text(text_frame, wrap="word", undo=True, font=("Consolas", 10))
        scrollbar = tk.Scrollbar(text_frame, command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.text_area.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # ── テンプレート選択 ────────────────────────────────
        frame_tmpl = tk.LabelFrame(self, text="テンプレート (style.docx)", **pad)
        frame_tmpl.pack(fill="x", **pad)

        tmpl_inner = tk.Frame(frame_tmpl)
        tmpl_inner.pack(fill="x", padx=4, pady=4)
        self.template_var = tk.StringVar()
        self.template_combo = ttk.Combobox(tmpl_inner, textvariable=self.template_var, state="readonly", width=40)
        self.template_combo.pack(side="left")
        tk.Button(tmpl_inner, text="参照...", command=self._browse_template).pack(side="left", padx=6)
        tk.Button(tmpl_inner, text="更新", command=self._load_templates).pack(side="left")

        # ── 出力ファイル名 ──────────────────────────────────
        frame_out = tk.LabelFrame(self, text="出力ファイル名", **pad)
        frame_out.pack(fill="x", **pad)

        out_inner = tk.Frame(frame_out)
        out_inner.pack(fill="x", padx=4, pady=4)
        self.out_var = tk.StringVar(value=self._default_filename())
        tk.Entry(out_inner, textvariable=self.out_var, width=40).pack(side="left")
        tk.Label(out_inner, text=".docx").pack(side="left")
        tk.Button(out_inner, text="保存先を選択...", command=self._browse_output).pack(side="left", padx=6)

        # ── 変換ボタン ──────────────────────────────────────
        tk.Button(
            self,
            text="Wordに変換して開く",
            font=("", 12, "bold"),
            bg="#0078D4",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            pady=8,
            command=self._run_convert,
        ).pack(fill="x", padx=10, pady=10)

        # ── ステータスバー ──────────────────────────────────
        self.status_var = tk.StringVar(value="準備完了")
        tk.Label(self, textvariable=self.status_var, anchor="w", fg="gray").pack(fill="x", padx=10, pady=(0, 6))

    def _default_filename(self):
        return datetime.now().strftime("%Y%m%d_document")

    def _load_templates(self):
        self.templates = get_templates()
        names = list(self.templates.keys())
        self.template_combo["values"] = names if names else ["(テンプレートなし)"]
        if names:
            self.template_combo.current(0)
        else:
            self.template_var.set("(テンプレートなし)")

    def _paste_clipboard(self):
        try:
            text = self.clipboard_get()
            self.text_area.delete("1.0", "end")
            self.text_area.insert("1.0", text)
            self.status_var.set("クリップボードから貼り付けました")
        except tk.TclError:
            messagebox.showwarning("貼り付け失敗", "クリップボードにテキストがありません。")

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Markdownファイルを選択",
            filetypes=[("Markdownファイル", "*.md"), ("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if path:
            self._load_md_file(path)

    def _load_md_file(self, path):
        with open(path, encoding="utf-8") as f:
            self.text_area.delete("1.0", "end")
            self.text_area.insert("1.0", f.read())
        # 出力ファイル名をMDファイル名ベースに設定
        base = os.path.splitext(os.path.basename(path))[0]
        self.out_var.set(os.path.join(os.path.dirname(path), base))
        self.status_var.set(f"読み込み完了: {os.path.basename(path)}")

    def _clear_text(self):
        self.text_area.delete("1.0", "end")
        self.out_var.set(self._default_filename())
        self.status_var.set("クリア済み")

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="テンプレートWordファイルを選択",
            filetypes=[("Wordファイル", "*.docx")],
        )
        if path:
            name = os.path.basename(path)
            self.templates[name] = path
            values = list(self.template_combo["values"])
            if name not in values:
                values.append(name)
                self.template_combo["values"] = values
            self.template_var.set(name)

    def _browse_output(self):
        current = self.out_var.get()
        init_dir = os.path.dirname(current) if os.path.dirname(current) else os.path.expanduser("~")
        init_file = os.path.basename(current) if not os.path.isdir(current) else self._default_filename()
        path = filedialog.asksaveasfilename(
            title="保存先を選択",
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".docx",
            filetypes=[("Wordファイル", "*.docx")],
        )
        if path:
            # 拡張子を除いて保存
            self.out_var.set(os.path.splitext(path)[0])

    def _run_convert(self):
        markdown_text = self.text_area.get("1.0", "end").strip()
        if not markdown_text:
            messagebox.showwarning("入力なし", "Markdownテキストを入力または貼り付けてください。")
            return

        # テンプレート解決
        selected = self.template_var.get()
        ref_file = self.templates.get(selected)

        # 出力パス
        out_base = self.out_var.get().strip()
        if not out_base:
            out_base = self._default_filename()
        if not out_base.endswith(".docx"):
            out_file = out_base + ".docx"
        else:
            out_file = out_base

        # 出力ディレクトリが指定されていない場合はダウンロードフォルダ
        if not os.path.dirname(out_file):
            downloads_dir = os.path.expanduser("~/Downloads")
            if not os.path.exists(downloads_dir):
                downloads_dir = os.path.expanduser("~")
            out_file = os.path.join(downloads_dir, out_file)

        self.status_var.set("変換中...")
        self.update()

        try:
            convert(markdown_text, ref_file, out_file)
            os.startfile(out_file)
            self.status_var.set(f"完了: {out_file}")
        except Exception as e:
            messagebox.showerror("エラー", f"変換に失敗しました。\n{e}")
            self.status_var.set("変換失敗")


if __name__ == "__main__":
    app = App()
    app.mainloop()

import sys
import os
import re
import threading
import tempfile
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime

import pypandoc
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

# --- アプリの見た目を今風（モダン）に設定 ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def resource_path(relative_path):
    """PyInstallerでexe化した後も、内蔵リソースへのパスを正しく解決する"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_templates():
    """templatesフォルダとルートのstyle.docxを走査してテンプレート辞書を返す"""
    templates = {}
    default_style = resource_path("style.docx")
    if os.path.exists(default_style):
        templates["デフォルト (style.docx)"] = default_style
    templates_dir = resource_path("templates")
    if os.path.exists(templates_dir):
        for f in sorted(os.listdir(templates_dir)):
            if f.lower().endswith(".docx"):
                name = os.path.splitext(f)[0]
                templates[name] = os.path.join(templates_dir, f)
    return templates


class ModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Markdown to Word")
        self.geometry("680x560")
        self.minsize(500, 400)

        # テキストエリアが画面いっぱいに広がるよう設定
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.templates = {}

        self._build_ui()
        self._load_templates()

        # exeアイコンにD&Dされた時の全自動処理
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self._handle_dnd(sys.argv[1])

    # ──────────────────────────────────────────
    # UI構築
    # ──────────────────────────────────────────
    def _build_ui(self):
        # 0. ファイル操作ボタン（上部）
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=15, pady=(12, 0), sticky="ew")
        top_frame.grid_columnconfigure(3, weight=1)

        ctk.CTkButton(top_frame, text="クリップボード貼り付け", width=160,
                      command=self._paste_clipboard).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(top_frame, text="ファイルから読み込み", width=150,
                      command=self._open_file).grid(row=0, column=1, padx=(0, 6))
        ctk.CTkButton(top_frame, text="クリア", width=70,
                      fg_color="gray40", hover_color="gray30",
                      command=self._clear_text).grid(row=0, column=2)

        # 1. テキスト入力エリア（画面いっぱいに広がる）
        self.textbox = ctk.CTkTextbox(self, font=("Consolas", 13), wrap="word")
        self.textbox.grid(row=1, column=0, padx=15, pady=8, sticky="nsew")

        # 2. テンプレート選択行
        tmpl_frame = ctk.CTkFrame(self, fg_color="transparent")
        tmpl_frame.grid(row=2, column=0, padx=15, pady=(0, 4), sticky="ew")
        tmpl_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(tmpl_frame, text="テンプレート:").grid(row=0, column=0, padx=(0, 8))
        self.template_var = ctk.StringVar(value="(自動装飾のみ)")
        self.template_combo = ctk.CTkComboBox(
            tmpl_frame, variable=self.template_var, state="readonly", width=280
        )
        self.template_combo.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ctk.CTkButton(tmpl_frame, text="参照...", width=70,
                      command=self._browse_template).grid(row=0, column=2, padx=(0, 6))
        ctk.CTkButton(tmpl_frame, text="更新", width=55,
                      fg_color="gray40", hover_color="gray30",
                      command=self._load_templates).grid(row=0, column=3)

        # 3. 出力ファイル名 ＋ 変換ボタン
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=3, column=0, padx=15, pady=(0, 8), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        self.out_var = ctk.StringVar(value=datetime.now().strftime("%Y%m%d_document"))
        self.entry_out = ctk.CTkEntry(
            bottom_frame, textvariable=self.out_var,
            font=("", 13), placeholder_text="出力ファイル名（拡張子不要）"
        )
        self.entry_out.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(bottom_frame, text="保存先...", width=80,
                      command=self._browse_output).grid(row=0, column=1, padx=(0, 6))

        self.btn_convert = ctk.CTkButton(
            bottom_frame, text="Wordに変換", width=120,
            font=("", 14, "bold"), command=self._start_conversion
        )
        self.btn_convert.grid(row=0, column=2)

        # 4. ステータスバー
        self.status_var = ctk.StringVar(value="準備完了")
        ctk.CTkLabel(self, textvariable=self.status_var,
                     anchor="w", text_color="gray60").grid(
            row=4, column=0, padx=16, pady=(0, 10), sticky="ew"
        )

    # ──────────────────────────────────────────
    # テンプレート操作
    # ──────────────────────────────────────────
    def _load_templates(self):
        self.templates = get_templates()
        names = list(self.templates.keys())
        if names:
            self.template_combo.configure(values=names)
            self.template_var.set(names[0])
        else:
            self.template_combo.configure(values=["(自動装飾のみ)"])
            self.template_var.set("(自動装飾のみ)")

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="テンプレートWordファイルを選択",
            filetypes=[("Wordファイル", "*.docx")],
        )
        if path:
            name = os.path.basename(path)
            self.templates[name] = path
            values = list(self.template_combo.cget("values"))
            if name not in values:
                values.append(name)
                self.template_combo.configure(values=values)
            self.template_var.set(name)

    # ──────────────────────────────────────────
    # ファイル操作
    # ──────────────────────────────────────────
    def _paste_clipboard(self):
        try:
            text = self.clipboard_get()
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", text)
            self.status_var.set("クリップボードから貼り付けました")
        except Exception:
            messagebox.showwarning("貼り付け失敗", "クリップボードにテキストがありません。")

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Markdownファイルを選択",
            filetypes=[("Markdownファイル", "*.md"), ("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if path:
            self._load_md_file(path)

    def _load_md_file(self, path):
        try:
            with open(path, encoding="utf-8") as f:
                content = f.read()
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", content)
            base = os.path.splitext(os.path.basename(path))[0]
            out_dir = os.path.dirname(path)
            self.out_var.set(os.path.join(out_dir, base))
            self.status_var.set(f"読み込み完了: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"ファイルの読み込みに失敗しました:\n{e}")

    def _clear_text(self):
        self.textbox.delete("1.0", "end")
        self.out_var.set(datetime.now().strftime("%Y%m%d_document"))
        self.status_var.set("クリア済み")

    def _browse_output(self):
        current = self.out_var.get().strip()
        init_dir = os.path.dirname(current) if os.path.dirname(current) else os.path.expanduser("~")
        init_file = os.path.basename(current) if current else datetime.now().strftime("%Y%m%d_document")
        path = filedialog.asksaveasfilename(
            title="保存先を選択",
            initialdir=init_dir,
            initialfile=init_file,
            defaultextension=".docx",
            filetypes=[("Wordファイル", "*.docx")],
        )
        if path:
            self.out_var.set(os.path.splitext(path)[0])

    # ──────────────────────────────────────────
    # D&D処理
    # ──────────────────────────────────────────
    def _handle_dnd(self, file_path):
        """D&Dされたファイルを読み込み、0.5秒後に自動で変換をスタートする"""
        self._load_md_file(file_path)
        self.after(500, self._start_conversion)

    # ──────────────────────────────────────────
    # 変換処理
    # ──────────────────────────────────────────
    def _start_conversion(self):
        text = self.textbox.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("入力なし", "Markdownテキストがありません。")
            return

        # ボタンの連打防止
        self.btn_convert.configure(text="変換中...", state="disabled")
        self.status_var.set("変換中...")

        out_base = self.out_var.get().strip()
        if not out_base:
            out_base = datetime.now().strftime("%Y%m%d_document")
        if out_base.endswith(".docx"):
            out_file = out_base
        else:
            out_file = out_base + ".docx"

        # 保存先フォルダが指定されていない場合はダウンロードフォルダへ
        if not os.path.dirname(out_file):
            downloads = os.path.expanduser("~/Downloads")
            out_file = os.path.join(downloads if os.path.exists(downloads) else os.path.expanduser("~"), out_file)

        # テンプレート解決
        selected = self.template_var.get()
        ref_file = self.templates.get(selected)

        # 画面をフリーズさせずに裏で変換処理を行う
        threading.Thread(
            target=self._process_conversion,
            args=(text, out_file, ref_file),
            daemon=True
        ).start()

    def _process_conversion(self, text, out_file, ref_file):
        try:
            # #の後のスペース忘れを自動修正
            text = re.sub(r'^(#{1,6})([^#\s])', r'\1 \2', text, flags=re.MULTILINE)

            with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as tmp:
                tmp.write(text)
                tmp_path = tmp.name

            # Pandocで変換
            extra_args = ["-f", "markdown", "-V", "lang=ja-JP"]
            if ref_file and os.path.exists(ref_file):
                extra_args.append(f"--reference-doc={ref_file}")

            pypandoc.convert_file(tmp_path, "docx", outputfile=out_file, extra_args=extra_args)
            os.unlink(tmp_path)

            # python-docxでスタイルを上書き適用
            self._apply_styles(out_file)

            self.after(0, lambda: self._finish_success(out_file))

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: messagebox.showerror("変換エラー", err_msg))
            self.after(0, lambda: self._reset_button())

    def _finish_success(self, out_file):
        self._reset_button()
        self.status_var.set(f"完了: {out_file}")
        os.startfile(out_file)

    def _reset_button(self):
        self.btn_convert.configure(text="Wordに変換", state="normal")
        if self.status_var.get() == "変換中...":
            self.status_var.set("準備完了")

    # ──────────────────────────────────────────
    # スタイル適用
    # ──────────────────────────────────────────
    def _apply_styles(self, docx_path):
        """Wordのデザインをpython-docxで強制上書き"""
        doc = Document(docx_path)

        def get_style(name):
            return doc.styles[name] if name in doc.styles else doc.styles.add_style(name, 1)

        # 標準 (游明朝・10.5pt)
        s_norm = get_style('Normal')
        s_norm.font.size = Pt(10.5)
        s_norm.font.name = '游明朝'
        if s_norm.font.element.rPr is None:
            s_norm.font.element.get_or_add_rPr()
        s_norm.font.element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')

        # 見出し1 (青・太字・18pt)
        s_h1 = get_style('Heading 1')
        s_h1.font.size = Pt(18)
        s_h1.font.bold = True
        s_h1.font.color.rgb = RGBColor(0, 112, 192)

        # 見出し2 (紺・太字・14pt)
        s_h2 = get_style('Heading 2')
        s_h2.font.size = Pt(14)
        s_h2.font.bold = True
        s_h2.font.color.rgb = RGBColor(0, 32, 96)

        # 見出し3 (黒・太字・12pt)
        s_h3 = get_style('Heading 3')
        s_h3.font.size = Pt(12)
        s_h3.font.bold = True
        s_h3.font.color.rgb = RGBColor(0, 0, 0)

        # 表スタイル
        for table in doc.tables:
            try:
                table.style = 'Light Shading Accent 1'
            except Exception:
                try:
                    table.style = 'Table Grid'
                except Exception:
                    pass

        doc.save(docx_path)


if __name__ == "__main__":
    app = ModernApp()
    app.mainloop()

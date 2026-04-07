import sys
import os
import re
import threading
import tempfile
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime

import pypandoc
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

# --- アプリの見た目を今風（モダン）に設定 ---
ctk.set_appearance_mode("System")  # PCのダーク/ライトモードに自動追従
ctk.set_default_color_theme("blue")

class ModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Markdown to Word")
        # どんなに小さくしてもレイアウトが崩れない設定
        self.geometry("600x450")
        self.minsize(400, 300)

        # 画面の伸び縮みに合わせてテキストエリアが自動調整される魔法
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()

        # 🌟 exeアイコンにD&Dされた時の全自動処理
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self._handle_dnd(sys.argv[1])

    def _build_ui(self):
        # 1. テキスト入力エリア（常に画面いっぱいに広がる）
        self.textbox = ctk.CTkTextbox(self, font=("Consolas", 14), wrap="word")
        self.textbox.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="nsew")

        # 2. 下部の操作エリア（絶対に画面外にはみ出さない）
        bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1)

        # 出力ファイル名入力欄
        self.out_var = ctk.StringVar(value=datetime.now().strftime("%Y%m%d_document"))
        self.entry_out = ctk.CTkEntry(bottom_frame, textvariable=self.out_var, font=("", 13), placeholder_text="出力ファイル名")
        self.entry_out.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        # 変換ボタン
        self.btn_convert = ctk.CTkButton(bottom_frame, text="Wordに変換", font=("", 14, "bold"), command=self._start_conversion)
        self.btn_convert.grid(row=0, column=1)

    def _handle_dnd(self, file_path):
        """D&Dされたファイルを読み込み、0.5秒後に自動で変換をスタートする"""
        try:
            with open(file_path, encoding="utf-8") as f:
                self.textbox.insert("1.0", f.read())
            self.out_var.set(os.path.splitext(file_path)[0])
            self.after(500, self._start_conversion)
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました:\n{e}")

    def _start_conversion(self):
        text = self.textbox.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("入力なし", "Markdownテキストがありません。")
            return

        # ボタンの連打防止
        self.btn_convert.configure(text="変換中...", state="disabled")
        
        out_base = self.out_var.get().strip()
        if not out_base.endswith(".docx"):
            out_base += ".docx"

        # 画面をフリーズさせずに裏で変換処理を行う
        threading.Thread(target=self._process_conversion, args=(text, out_base), daemon=True).start()

    def _process_conversion(self, text, out_file):
        try:
            # Markdownのスペース忘れ修正
            text = re.sub(r'^(#{1,6})([^#\s])', r'\1 \2', text, flags=re.MULTILINE)
            
            with tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8") as tmp:
                tmp.write(text)
                tmp_path = tmp.name

            # Wordへ変換
            pypandoc.convert_file(tmp_path, "docx", outputfile=out_file, extra_args=["-f", "markdown", "-V", "lang=ja-JP"])
            os.unlink(tmp_path)

            # スタイル適用
            self._apply_styles(out_file)

            # 成功したらWordを開いてボタンを元に戻す
            self.after(0, lambda: self._finish_success(out_file))

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("変換エラー", str(e)))
            self.after(0, lambda: self.btn_convert.configure(text="Wordに変換", state="normal"))

    def _finish_success(self, out_file):
        self.btn_convert.configure(text="Wordに変換", state="normal")
        os.startfile(out_file)

    def _apply_styles(self, docx_path):
        """Wordのデザインを美しく整える"""
        doc = Document(docx_path)

        def get_style(name):
            return doc.styles[name] if name in doc.styles else doc.styles.add_style(name, 1)

        # 標準
        s_norm = get_style('Normal')
        s_norm.font.size = Pt(10.5)
        s_norm.font.name = '游明朝'
        if s_norm.font.element.rPr is None: s_norm.font.element.get_or_add_rPr()
        s_norm.font.element.rPr.rFonts.set(qn('w:eastAsia'), '游明朝')

        # 見出し1
        s_h1 = get_style('Heading 1')
        s_h1.font.size = Pt(18)
        s_h1.font.bold = True
        s_h1.font.color.rgb = RGBColor(0, 112, 192)

        # 見出し2
        s_h2 = get_style('Heading 2')
        s_h2.font.size = Pt(14)
        s_h2.font.bold = True
        s_h2.font.color.rgb = RGBColor(0, 32, 96)

        # 見出し3
        s_h3 = get_style('Heading 3')
        s_h3.font.size = Pt(12)
        s_h3.font.bold = True
        s_h3.font.color.rgb = RGBColor(0, 0, 0)

        # 表
        for table in doc.tables:
            try: table.style = 'Light Shading Accent 1'
            except: pass

        doc.save(docx_path)

if __name__ == "__main__":
    app = ModernApp()
    app.mainloop()
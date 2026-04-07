import os
import shutil
import PyInstaller.__main__

# プログラム自身が自分のいるフォルダに移動してからビルドする
# → カレントディレクトリのズレによる「ファイルが見つからない」エラーを完全防止
current_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(current_dir)

print("==================================================")
print("  Master Builder: 究極の実行ファイル作成スクリプト")
print("==================================================")

target_file = "Markdown2Word_Master.py"
app_name = "Markdown2Word_Pro"

# 1. 徹底的なクリーンアップ
print("\n[1/3] 古いキャッシュを完全に消去しています...")
for folder in ["build", "dist"]:
    if os.path.exists(folder):
        shutil.rmtree(folder)
if os.path.exists(f"{app_name}.spec"):
    os.remove(f"{app_name}.spec")

# 2. PyInstallerオプションを組み立て
print("\n[2/3] コンパイルを開始します。これには数分かかる場合があります...")

pyinstaller_args = [
    target_file,
    f"--name={app_name}",
    "--noconsole",      # コマンドプロンプトを非表示
    "--onefile",        # 単一exeにまとめる
    "--clean",          # クリーンビルド

    # ── 必須：データファイルを含む依存ライブラリ ──
    "--collect-data=customtkinter",  # UI画像・テーマJSON等が必要
    "--collect-data=docx",          # python-docx のテンプレートXML等
    "--hidden-import=pypandoc",
    "--hidden-import=docx",
    "--hidden-import=customtkinter",
]

# templatesフォルダが存在する場合のみ同梱する
# (exe実行時に resource_path("templates") で参照できるようになる)
templates_src = os.path.join(current_dir, "templates")
if os.path.exists(templates_src):
    # Windows: --add-data=src;dest  (セミコロン区切り)
    pyinstaller_args.append(f"--add-data={templates_src};templates")
    print(f"    → templatesフォルダを同梱します: {templates_src}")
else:
    print("    → templatesフォルダが見つかりません。テンプレートなしでビルドします。")
    print("      (後から templates/*.docx を追加したい場合は再ビルドしてください)")

# style.docxが存在する場合のみ同梱
style_src = os.path.join(current_dir, "style.docx")
if os.path.exists(style_src):
    pyinstaller_args.append(f"--add-data={style_src};.")
    print(f"    → style.docxを同梱します: {style_src}")

try:
    PyInstaller.__main__.run(pyinstaller_args)
    print("\n[3/3] ビルド大成功！")
    print(f"  → dist/{app_name}.exe が作成されました。")
    print("  この .exe は単体で配布できます。")
    print("  .mdファイルをexeアイコンにD&Dするだけで変換できます！")

except Exception as e:
    print(f"\n[ERROR] ビルド中にエラーが発生しました: {e}")
    raise SystemExit(1)

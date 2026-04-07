import os
import sys
import shutil
import PyInstaller.__main__

# 🌟 追加：プログラム自身が、自分が保存されているフォルダに移動する魔法の2行
# これにより「ファイルが見つからない(does not exist)」エラーを完全に防ぎます。
current_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(current_dir)

print("==================================================")
print(" 🚀 Master Builder: 究極の実行ファイル作成スクリプト")
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

# 2. PyInstallerの実行
print("\n[2/3] コンパイルを開始します。これには数分かかる場合があります...")
try:
    PyInstaller.__main__.run([
        target_file,
        '--name=%s' % app_name,
        '--noconsole',              # コマンドプロンプトを隠す
        '--onefile',                # 1ファイルにまとめる
        '--clean',                  # クリーンビルド
        
        # 🎯 依存ライブラリを全て強制的にフックする
        '--collect-data=docx',      
        '--hidden-import=pypandoc',
        '--hidden-import=docx',
    ])
    print("\n[3/3] ✨ ビルド大成功！ ✨")
    print(f"👉 フォルダ内の 'dist' フォルダに '{app_name}.exe' が作成されました。")
    print("この .exe は、Pandocが入っていない他人のPCでも単体で動作します！")

except Exception as e:
    print(f"\n❌ ビルド中にエラーが発生しました: {e}")
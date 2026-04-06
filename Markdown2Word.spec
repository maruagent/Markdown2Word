# -*- mode: python ; coding: utf-8 -*-
import os

# exeに同梱するファイルを指定
# (実際のファイルパス, exe内での格納先)
datas = []

# style.docx が存在すれば同梱
if os.path.exists('style.docx'):
    datas.append(('style.docx', '.'))

# templates フォルダが存在すれば同梱
if os.path.exists('templates'):
    datas.append(('templates', 'templates'))

a = Analysis(
    ['Markdown2Word.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['pypandoc'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Markdown2Word',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # 黒いコンソール画面を非表示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

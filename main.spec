# -*- mode: python ; coding: utf-8 -*-

import os

# 工程根目录（即本 spec 文件所在目录）
project_dir = os.path.abspath(SPECPATH)

block_cipher = None

a = Analysis(
    [os.path.join(project_dir, 'main.py')],
    pathex=[project_dir],
    binaries=[],
    datas=[
        (os.path.join(project_dir, 'price.xlsx'), '.'),
        (os.path.join(project_dir, 'price_accessor.xlsx'), '.'),
        (os.path.join(project_dir, 'factory.xlsx'), '.'),
        (os.path.join(project_dir, 'config.yaml'), '.'),
        (os.path.join(project_dir, 'images'), 'images'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[os.path.join(project_dir, 'runtime_hook.py')],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='备货单工具_3.10.0_外发版',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='备货单工具_3.10.0_外发版',
)
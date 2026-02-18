a = Analysis(
    ['main.py'],
    pathex=['E:\\PyCharmProjects\\WP快通'],
    binaries=[],
    datas=[
        ('E:\\PyCharmProjects\\WP快通\\assets\\images', 'assets\\images'),
        ('E:\\PyCharmProjects\\WP快通\\src', 'src')
    ],
    # 核心保留：兜底防止漏装src模块，无模块报错的关键
    hiddenimports=['src', 'src.main_window'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'jupyter', 'ipython', 'matplotlib', 'numpy', 'pandas', 'scipy', 'notebook', 'seaborn'],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='WP快通',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_dir='E:\\PyCharmProjects\\WP快通\\upx',
    upx_args=['-9'],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['E:\\PyCharmProjects\\WP快通\\assets\\icons\\app.ico']
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_dir='E:\\PyCharmProjects\\WP快通\\upx',
    upx_args=['-9'],
    upx_exclude=[],
    name='WP快通',
)
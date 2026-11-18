# ExamSync.spec — PyInstaller build configuration
block_cipher = None

a = Analysis(
    ['desktop_app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # Bundle the HTML templates folder into the exe
        ('templates', 'templates'),
    ],
    hiddenimports=[
        # Flask internals
        'flask',
        'flask.templating',
        'jinja2',
        'jinja2.ext',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.routing',
        'werkzeug.exceptions',
        'werkzeug.middleware',
        'werkzeug.middleware.proxy_fix',

        # python-docx
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.parts',

        # Data / Excel / PDF
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'reportlab',
        'reportlab.platypus',
        'reportlab.lib.styles',
        'reportlab.lib.colors',
        'reportlab.lib.pagesizes',
        'reportlab.pdfgen',

        # pywebview Windows backends (include all — installer picks the right one)
        'webview',
        'webview.platforms.winforms',
        'webview.platforms.edgechromium',
        'webview.platforms.mshtml',
        'clr',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExamSync',                  # final exe name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                    # FALSE = no black terminal window, clean desktop app
    icon=None,                        # put 'icon.ico' here if you have one
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

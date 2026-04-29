# price_file_app.spec — PyInstaller build spec for Price File App
#
# Build with:
#   pyinstaller price_file_app.spec
#
# Output:  dist/PriceFileApp/PriceFileApp.exe
#
# Version format: YYYY.MM.N  (year, month, release number for that month)

APP_VERSION = '2026.04.03'

block_cipher = None

a = Analysis(
    ['qt_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Note: settings.ini, help.html, data/, brand_import_mappings.json, and
        # templates.json are NOT bundled here because the app resolves them relative
        # to the exe directory (os.path.dirname(sys.executable)).  Copy those
        # manually after the build — see the README Distribution section.
    ],
    hiddenimports=[
        'pyodbc',
        'openpyxl',
        'openpyxl.cell._writer',
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'streamlit',
        'tornado',
        'watchdog',
        'gitpython',
        'git',
        'pydeck',
        'altair',
        'boto3',
        'botocore',
        'IPython',
        'jinja2',
        'jsonschema',
        'matplotlib',
        'scipy',
        'sklearn',
        'PIL',
        'pytest',
        'setuptools',
        'pkg_resources',
    ],
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
    name='PriceFileApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
    version_file='version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PriceFileApp',
)

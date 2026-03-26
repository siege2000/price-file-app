# price_file_app.spec — PyInstaller build spec for Price File App
#
# Build with:
#   pyinstaller price_file_app.spec
#
# Output:  dist/PriceFileApp/PriceFileApp.exe
#
# After building, copy these alongside the exe before shipping:
#   settings.ini
#   data/  (entire folder)

block_cipher = None

a = Analysis(
    ['qt_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Brand import mapping config (read by specials module via __file__)
        ('specials/brand_import_mappings.json', 'specials'),
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
        # Streamlit and its dependency tree — not used in the Qt app
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
        'jinja2',          # pulled in by streamlit, not needed standalone
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
    upx=False,          # set True if UPX is installed, reduces size ~30%
    console=False,      # no console window (GUI app)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icon.ico',  # uncomment and supply a .ico file to set the exe icon
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

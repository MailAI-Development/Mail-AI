# -*- mode: python ; coding: utf-8 -*-
# Build from the repo root:   pyinstaller packaging/mailai.spec
# Output: dist/Mail AI v<APP_VERSION>.exe   (version is read from app/logic.py)
import os, re, sys

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
APP = os.path.join(ROOT, "app")

with open(os.path.join(APP, "logic.py"), encoding="utf-8") as f:
    VERSION = re.search(r'^APP_VERSION = "([^"]+)"', f.read(), re.M).group(1)

# Refuse to build an app that can't reach the extraction proxy.
sys.path.insert(0, APP)
try:
    import _build_secrets as _s
except ImportError:
    _s = None
for _name in ("PROXY_URL", "APP_TOKEN", "LICENSE_SECRET"):
    if not (os.environ.get(f"MAILAI_{_name}") or getattr(_s, _name, "")):
        raise SystemExit(f"Build aborted: {_name} not set (app/_build_secrets.py or MAILAI_{_name}).")

a = Analysis(
    [os.path.join(APP, "main.py")],
    pathex=[APP],
    binaries=[],
    datas=[(os.path.join(APP, "assets"), "assets")],
    hiddenimports=["win32timezone", "_build_secrets"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
splash = Splash(
    os.path.join(APP, "assets", "splash.png"),
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    splash,
    splash.binaries,
    [],
    name=f"Mail AI v{VERSION}",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[os.path.join(APP, "assets", "icon.ico")],
)

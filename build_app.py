#!/usr/bin/env python3
"""'온글' macOS 앱 묶음을 만든다. 표준 도구(iconutil)와 Pillow만 쓴다.

번들은 파이썬 소스를 그대로 담고, 실행기가 쓸 만한 파이썬을 찾아 창 화면을 띄운다.
빌드 산출물(.app)은 저장소에 넣지 않는다. 이 스크립트가 만드는 방법이다.
"""
from __future__ import annotations

import argparse
import plistlib
import shutil
import subprocess
import sys
from pathlib import Path

APP_NAME = "온글"
BUNDLE_ID = "local.mdmaker.ongeul"
# 모듈을 손으로 나열하면 새로 추가한 것을 빠뜨린다(실제로 xls.py 를 빠뜨렸다).
SKIP_SOURCES = ("build_app.py",)


def sources(project: Path) -> list[Path]:
    return sorted(p for p in project.glob("*.py")
                  if p.name not in SKIP_SOURCES and not p.name.startswith("test_"))
ICON_SIZES = (16, 32, 64, 128, 256, 512, 1024)
KOREAN_FONTS = ("/System/Library/Fonts/AppleSDGothicNeo.ttc",
                "/System/Library/Fonts/Supplemental/AppleGothic.ttf")


def draw_icon(size: int):
    """둥근 사각형 바탕에 '온' 한 글자. 16픽셀에서도 읽히도록 단순하게 둔다."""
    from PIL import Image, ImageDraw, ImageFont
    scale = 4 if size <= 256 else 2
    w = size * scale
    img = Image.new("RGBA", (w, w), (0, 0, 0, 0))
    grad = Image.new("RGB", (1, w))
    for y in range(w):                      # 위에서 아래로 남보라 → 보라
        t = y / max(1, w - 1)
        grad.putpixel((0, y), (int(46 + 60 * t), int(42 + 20 * t), int(120 + 60 * t)))
    grad = grad.resize((w, w))
    mask = Image.new("L", (w, w), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, w - 1, w - 1), radius=int(w * 0.22), fill=255)
    img.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(img)
    # 종이 한 장 위에 이름 첫 글자. 16픽셀로 줄여도 "문서"로 읽히는 쪽을 택했다.
    card = (w * 0.26, w * 0.17, w * 0.74, w * 0.83)
    d.rounded_rectangle(card, radius=int(w * 0.065), fill=(255, 255, 255, 242))
    font = None
    for path in KOREAN_FONTS:
        try:
            font = ImageFont.truetype(path, int(w * 0.33), index=0)
            break
        except OSError:
            continue
    if font is not None:
        d.text((w / 2, w * 0.42), "온", font=font, fill=(62, 46, 140, 255), anchor="mm")
    bar_h = max(1, int(w * 0.036))
    for i, ratio in enumerate((1.0, 0.62)):
        y = w * 0.615 + i * bar_h * 2.5
        d.rounded_rectangle((w * 0.345, y, w * 0.345 + w * 0.31 * ratio, y + bar_h),
                            radius=bar_h / 2, fill=(126, 110, 196, 255))
    return img.resize((size, size), Image.LANCZOS)


def make_icns(dest: Path) -> bool:
    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        print("Pillow가 없어 아이콘을 건너뛴다")
        return False
    iconset = dest.parent / "ongeul.iconset"
    shutil.rmtree(iconset, ignore_errors=True)
    iconset.mkdir(parents=True)
    for s in ICON_SIZES:
        draw_icon(s).save(iconset / ("icon_%dx%d.png" % (s, s)))
        if s <= 512:
            draw_icon(s * 2).save(iconset / ("icon_%dx%d@2x.png" % (s, s)))
    if not shutil.which("iconutil"):
        print("iconutil이 없어 icns를 만들지 못했다")
        return False
    r = subprocess.run(["iconutil", "-c", "icns", str(iconset), "-o", str(dest)],
                       capture_output=True, shell=False)
    shutil.rmtree(iconset, ignore_errors=True)
    if r.returncode != 0:
        print("iconutil 실패:", r.stderr.decode("utf-8", "replace")[:200])
        return False
    return True


LAUNCHER = """#!/bin/sh
# 온글 실행기: 창 화면에 필요한 파이썬을 찾아서 띄운다.
DIR=$(cd "$(dirname "$0")/../Resources/app" && pwd)
for PY in %s /usr/local/bin/python3 /opt/homebrew/bin/python3 /usr/bin/python3; do
    [ -x "$PY" ] || PY=$(command -v "$PY" 2>/dev/null) || continue
    [ -x "$PY" ] || continue
    if "$PY" -c 'import tkinter' 2>/dev/null; then
        exec "$PY" "$DIR/gui.py" "$@"
    fi
done
osascript -e 'display alert "온글을 실행할 수 없습니다" message "tkinter가 포함된 python3가 필요합니다. 터미널에서 python3 gui.py 로 실행해 보세요."'
exit 1
"""


def build(project: Path, out_dir: Path) -> Path:
    app = out_dir / (APP_NAME + ".app")
    shutil.rmtree(app, ignore_errors=True)
    macos = app / "Contents" / "MacOS"
    res = app / "Contents" / "Resources"
    (res / "app").mkdir(parents=True)
    macos.mkdir(parents=True)

    for f in sources(project):
        shutil.copy2(f, res / "app" / f.name)

    icon_ok = make_icns(res / "ongeul.icns")
    info = {
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleExecutable": "ongeul",
        "CFBundleIdentifier": BUNDLE_ID,
        "CFBundlePackageType": "APPL",
        "CFBundleShortVersionString": __import__("mdmaker").VERSION,
        "CFBundleVersion": __import__("mdmaker").VERSION,
        "LSMinimumSystemVersion": "11.0",
        "NSHighResolutionCapable": True,
        "LSApplicationCategoryType": "public.app-category.productivity",
        "NSHumanReadableCopyright": "로컬에서만 동작하며 문서를 외부로 보내지 않습니다.",
    }
    if icon_ok:
        info["CFBundleIconFile"] = "ongeul.icns"
    (app / "Contents" / "Info.plist").write_bytes(plistlib.dumps(info))

    launcher = macos / "ongeul"
    launcher.write_text(LAUNCHER % _quote(sys.executable), encoding="utf-8")
    launcher.chmod(0o755)
    return app


def _quote(p: str) -> str:
    return '"%s"' % p if " " in p else p


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="온글 앱 묶음 만들기")
    ap.add_argument("--out", default="dist", help="앱을 놓을 폴더")
    opts = ap.parse_args(argv)
    project = Path(__file__).resolve().parent
    sys.path.insert(0, str(project))
    out_dir = Path(opts.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    app = build(project, out_dir)
    print("만들었다:", app)
    print("실행:      open '%s'" % app)
    return 0


if __name__ == "__main__":
    sys.exit(main())

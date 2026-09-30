# -*- coding: utf-8 -*-
"""LINEスタンプ用の画像整形＆申請ZIP作成スクリプト

使い方:
    python build_line_stamp.py company/stamp/<セット名> [--main N] [--whitebg]

入力:  <セット名>/src/ にスタンプの元画像（PNG/JPG）を入れておく
出力:  <セット名>/out/ に main.png / tab.png / 01.png〜40.png と申請用ZIP

LINE Creators Market の仕様（2026年時点）:
  - スタンプ画像: 最大 370x320px、8/16/24/32/40個、PNG、背景透過、余白10px推奨
  - メイン画像: 240x240px / トークルームタブ画像: 96x74px
  - 各ファイル1MB以下、ZIPにまとめてアップロード
"""
import argparse
import sys
import zipfile
from pathlib import Path

from PIL import Image

STAMP_W, STAMP_H = 370, 320
MARGIN = 10
MAIN_SIZE = (240, 240)
TAB_SIZE = (96, 74)
VALID_COUNTS = [8, 16, 24, 32, 40]
MAX_BYTES = 1024 * 1024


def remove_white_bg(img: Image.Image, threshold: int = 240) -> Image.Image:
    """縁から繋がった白背景だけを透過にする（キャラ内部の白は残す）"""
    img = img.convert("RGBA")
    w, h = img.size
    px = img.load()

    def is_white(x, y):
        r, g, b, a = px[x, y]
        return a > 0 and r >= threshold and g >= threshold and b >= threshold

    seen = set()
    stack = [(x, y) for x in range(w) for y in (0, h - 1) if is_white(x, y)]
    stack += [(x, y) for y in range(h) for x in (0, w - 1) if is_white(x, y)]
    while stack:
        x, y = stack.pop()
        if (x, y) in seen or not (0 <= x < w and 0 <= y < h) or not is_white(x, y):
            continue
        seen.add((x, y))
        px[x, y] = (0, 0, 0, 0)
        stack += [(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)]
    return img


def fit_on_canvas(img: Image.Image, canvas_w: int, canvas_h: int, margin: int) -> Image.Image:
    """透明キャンバスの中央に、余白を確保して収める"""
    img = img.convert("RGBA")
    bbox = img.getbbox()  # 透明部分を除いた実体の範囲
    if bbox:
        img = img.crop(bbox)
    inner_w, inner_h = canvas_w - margin * 2, canvas_h - margin * 2
    scale = min(inner_w / img.width, inner_h / img.height)
    img = img.resize((max(1, round(img.width * scale)), max(1, round(img.height * scale))), Image.LANCZOS)
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    canvas.paste(img, ((canvas_w - img.width) // 2, (canvas_h - img.height) // 2), img)
    return canvas


def main():
    ap = argparse.ArgumentParser(description="LINEスタンプ申請セットを作る")
    ap.add_argument("set_dir", help="セットのフォルダ（中に src/ を置く）")
    ap.add_argument("--main", type=int, default=1, help="メイン画像に使う画像の番号（1始まり、既定=1）")
    ap.add_argument("--whitebg", action="store_true", help="白背景の画像を自動で透過にする")
    args = ap.parse_args()

    set_dir = Path(args.set_dir)
    src_dir = set_dir / "src"
    out_dir = set_dir / "out"
    if not src_dir.is_dir():
        sys.exit(f"エラー: {src_dir} がありません。src/ フォルダに元画像を入れてください。")

    files = sorted(p for p in src_dir.iterdir() if p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"))
    if not files:
        sys.exit(f"エラー: {src_dir} に画像がありません。")

    n = len(files)
    if n not in VALID_COUNTS:
        nearest = max([c for c in VALID_COUNTS if c <= n], default=None)
        msg = f"注意: 画像が{n}個です。申請できるのは {VALID_COUNTS} 個のいずれか。"
        if nearest:
            msg += f" 先頭{nearest}個で作成します。"
            files = files[:nearest]
            n = nearest
        else:
            sys.exit(msg + f" あと{VALID_COUNTS[0] - n}個必要です。")
        print(msg)

    out_dir.mkdir(parents=True, exist_ok=True)
    outputs = []
    for i, f in enumerate(files, 1):
        img = Image.open(f)
        if args.whitebg:
            img = remove_white_bg(img)
        stamp = fit_on_canvas(img, STAMP_W, STAMP_H, MARGIN)
        dest = out_dir / f"{i:02d}.png"
        stamp.save(dest, "PNG", optimize=True)
        outputs.append(dest)
        print(f"  {f.name} -> {dest.name} ({STAMP_W}x{STAMP_H})")

    main_idx = min(max(args.main, 1), n) - 1
    main_src = Image.open(files[main_idx])
    if args.whitebg:
        main_src = remove_white_bg(main_src)
    fit_on_canvas(main_src, *MAIN_SIZE, 6).save(out_dir / "main.png", "PNG", optimize=True)
    fit_on_canvas(main_src, *TAB_SIZE, 2).save(out_dir / "tab.png", "PNG", optimize=True)
    outputs += [out_dir / "main.png", out_dir / "tab.png"]
    print(f"  main.png (240x240) / tab.png (96x74) ← {files[main_idx].name}")

    too_big = [p.name for p in outputs if p.stat().st_size > MAX_BYTES]
    if too_big:
        sys.exit(f"エラー: 1MBを超えるファイルがあります: {too_big}")

    zip_path = set_dir / f"{set_dir.name}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in outputs:
            z.write(p, p.name)
    print(f"\n完成: {zip_path}（スタンプ{n}個 + main + tab）")
    print("→ LINE Creators Market の新規登録画面でこのZIPをアップロードしてください。")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""議事録（.docx）の配布前チェック。

配布していいかどうかを機械的に判定する。人が目視する前の関門で、
とーるさんが打つものではなく /my-gijiroku が裏で走らせる。

  python3 templates/check-minutes.py 議事録_2026-09-30_担当者会議.docx

追加の依存は無し（標準ライブラリだけ）。LibreOffice があれば
--preview で目視用の PNG も出す。

終了コード:
  0 … 配布可
  1 … 配布前に直すところがある
  2 … ファイルが読めない
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 差し替え忘れのプレースホルダ（〔…〕）
PLACEHOLDER = re.compile(r"〔[^〕]*〕")
# 「未定」「TBD」など、埋まっていない印
UNSET = re.compile(r"(未定|TBD|tbd|未記入|要記入)")
# 仮名化の形（A様・B様…）。これ以外の「◯◯様/さん」は実名の疑い
PSEUDONYM = re.compile(r"^[A-Z]様$")
REAL_NAME = re.compile(r"([一-龥ぁ-んァ-ヶーA-Za-z]{2,6})(様|さん|氏)")


class Report:
    def __init__(self):
        self.blockers = []   # 配布前に必ず直す
        self.warnings = []   # 見ておいた方がいい
        self.notes = []      # 参考情報

    def blocker(self, title, detail=""):
        self.blockers.append((title, detail))

    def warn(self, title, detail=""):
        self.warnings.append((title, detail))

    def note(self, title, detail=""):
        self.notes.append((title, detail))

    def render(self):
        def block(label, items, mark):
            if not items:
                return []
            out = [f"\n{label}"]
            for title, detail in items:
                out.append(f"  {mark} {title}")
                if detail:
                    for line in detail.splitlines():
                        out.append(f"      {line}")
            return out

        lines = []
        lines += block("■ 配布前に直すところ", self.blockers, "×")
        lines += block("■ 見ておいた方がいいところ", self.warnings, "△")
        lines += block("■ 参考", self.notes, "・")
        lines.append("")
        if self.blockers:
            lines.append(f"判定: 配布前に {len(self.blockers)} 件直す必要がある")
        elif self.warnings:
            lines.append(f"判定: 配布できる（気になる点 {len(self.warnings)} 件は目で確認）")
        else:
            lines.append("判定: 配布できる")
        return "\n".join(lines)


def read_docx(path):
    """本文の段落テキストと、表をセルの二次元配列で返す。"""
    z = zipfile.ZipFile(path)
    body = ET.fromstring(z.read("word/document.xml")).find(f"{W}body")

    def para_text(p):
        return "".join(t.text or "" for t in p.iter(f"{W}t"))

    paragraphs = [para_text(p) for p in body.findall(f"{W}p")]

    tables = []
    for tbl in body.findall(f"{W}tbl"):
        rows = []
        for tr in tbl.findall(f"{W}tr"):
            rows.append([
                "".join(para_text(p) for p in tc.findall(f"{W}p"))
                for tc in tr.findall(f"{W}tc")
            ])
        tables.append(rows)

    return z, paragraphs, tables


def body_only(paragraphs):
    """配布前チェック欄より前の段落だけを返す。

    あの欄は「［要確認］が残っていないか」のような点検文を含むので、
    そのまま検査に混ぜると自分自身を誤検知する。
    """
    out = []
    for t in paragraphs:
        if "配布前チェック" in t:
            break
        out.append(t)
    return out


def check_needs_confirmation(paragraphs, tables, report):
    """［要確認］が残っていたら配布できない。"""
    hits = [t for t in paragraphs if "要確認" in t]
    for rows in tables:
        for row in rows:
            for cellv in row:
                if "要確認" in cellv:
                    hits.append(cellv)
    # 節見出し自体（「［要確認］一覧」）は中身が空なら問題なし
    real = [h for h in hits if h.strip() not in ("［要確認］一覧",)
            and not h.strip().startswith("〔") and "要確認" in h]
    real = [h for h in real if not re.fullmatch(r"\d*\.?\s*［要確認］一覧", h.strip())]
    if real:
        sample = "\n".join(f"- {h.strip()[:70]}" for h in real[:5])
        report.blocker(f"［要確認］が {len(real)} 件残っている", sample)


def check_placeholders(paragraphs, tables, report):
    """テンプレートの〔…〕が残っていたら埋め忘れ。"""
    found = []
    for t in paragraphs:
        found += PLACEHOLDER.findall(t)
    for rows in tables:
        for row in rows:
            for cellv in row:
                found += PLACEHOLDER.findall(cellv)
    if found:
        sample = "\n".join(f"- {f[:60]}" for f in found[:5])
        report.blocker(f"埋めていない項目が {len(found)} 件ある", sample)


def check_todo_table(tables, report):
    """ToDo 表の担当・期限が空なら、あとで必ず揉める。"""
    for rows in tables:
        if not rows:
            continue
        header = [h.strip() for h in rows[0]]
        if not ("担当" in header and "期限" in header):
            continue
        owner_i, due_i = header.index("担当"), header.index("期限")
        empty, unset = [], []
        for row in rows[1:]:
            if len(row) <= max(owner_i, due_i):
                continue
            # 行自体が空（記入枠）ならスキップ
            if not any(c.strip() for c in row if c.strip() not in [str(i) for i in range(1, 100)]):
                continue
            owner, due = row[owner_i].strip(), row[due_i].strip()
            label = (row[1].strip() if len(row) > 1 else "")[:40] or "(内容なし)"
            if not owner or not due:
                empty.append(f"- {label} → 担当「{owner or '空欄'}」／期限「{due or '空欄'}」")
            elif UNSET.search(owner) or UNSET.search(due):
                unset.append(f"- {label} → 担当「{owner}」／期限「{due}」")
        if empty:
            report.blocker(f"ToDo の担当・期限が空欄（{len(empty)} 件）", "\n".join(empty[:5]))
        if unset:
            report.warn(f"ToDo の担当・期限が未定のまま（{len(unset)} 件）", "\n".join(unset[:5]))


def check_predistribution_block(paragraphs, report):
    """配布前チェック欄は配布時に削除する前提。"""
    if any("配布前チェック" in t for t in paragraphs):
        report.blocker("テンプレートの「配布前チェック」欄が残っている",
                       "配布用のファイルからは削除する")


def check_personal_names(paragraphs, tables, report):
    """仮名化されていない実名の疑いを拾う。判定は人がする。"""
    texts = list(paragraphs)
    for rows in tables:
        for row in rows:
            texts += row
    suspects = set()
    for t in texts:
        for name, honorific in REAL_NAME.findall(t):
            token = f"{name}{honorific}"
            if PSEUDONYM.match(token):
                continue
            # 職種・立場を表す語は実名ではない
            if name in ("利用者", "ご家族", "家族", "担当", "職員", "出席", "参加",
                        "ケアマネ", "管理者", "本人", "皆", "各"):
                continue
            suspects.add(token)
    if suspects:
        report.warn(
            f"仮名化されていない氏名の可能性（{len(suspects)} 件）",
            "\n".join(f"- {s}" for s in sorted(suspects)[:8])
            + "\n利用者・ご家族なら A様・B様 に置き換える。職員名は事業所の慣習に合わせる",
        )


def check_layout(z, tables, report):
    """体裁が壊れていないか（列幅の合計・フォント）。"""
    body = ET.fromstring(z.read("word/document.xml")).find(f"{W}body")
    sect = body.find(f"{W}sectPr")
    pg_sz, pg_mar = sect.find(f"{W}pgSz"), sect.find(f"{W}pgMar")
    page_w = int(pg_sz.get(f"{W}w"))
    left = int(pg_mar.get(f"{W}left"))
    right = int(pg_mar.get(f"{W}right"))
    content_w = page_w - left - right

    for i, tbl in enumerate(body.findall(f"{W}tbl"), 1):
        grid = tbl.find(f"{W}tblGrid")
        if grid is None:
            continue
        widths = [int(c.get(f"{W}w")) for c in grid.findall(f"{W}gridCol")]
        if sum(widths) != content_w:
            report.warn(
                f"{i} 番目の表の列幅の合計が本文幅とずれている",
                f"合計 {sum(widths)} / 本文幅 {content_w}。Word 以外で開くと崩れることがある",
            )

    try:
        styles = z.read("word/styles.xml").decode("utf-8")
        m = re.search(r'w:eastAsia="([^"]+)"', styles)
        if m:
            report.note(f"本文フォント: {m.group(1)}")
    except KeyError:
        pass
    report.note(f"表の数: {len(tables)}")


def make_preview(path, outdir):
    """LibreOffice があれば目視用の PNG を出す。無ければ None。"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None, "LibreOffice が無いのでプレビュー画像は作れない"
    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as profile:
        try:
            subprocess.run(
                [soffice, "--headless", f"-env:UserInstallation=file://{profile}",
                 "--convert-to", "pdf", "--outdir", outdir, path],
                check=True, capture_output=True, timeout=180,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            return None, f"LibreOffice の変換に失敗した: {e}"
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if not os.path.exists(pdf):
        return None, "PDF が出来なかった"
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-jpeg", "-r", "100", pdf,
                        os.path.join(outdir, "page")], check=False)
        pages = sorted(
            os.path.join(outdir, f) for f in os.listdir(outdir)
            if f.startswith("page-") and f.endswith(".jpg")
        )
        if pages:
            return pages, None
    return [pdf], "pdftoppm が無いので PDF のまま（画像にはできていない）"


def main():
    ap = argparse.ArgumentParser(description="議事録(.docx)の配布前チェック")
    ap.add_argument("docx", help="チェックする .docx")
    ap.add_argument("--preview", metavar="DIR",
                    help="目視用のプレビュー画像を出すディレクトリ")
    args = ap.parse_args()

    if not os.path.exists(args.docx):
        print(f"エラー: ファイルが無い → {args.docx}", file=sys.stderr)
        return 2
    try:
        z, paragraphs, tables = read_docx(args.docx)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
        print(f"エラー: .docx として読めない → {args.docx}\n  {e}", file=sys.stderr)
        return 2

    report = Report()
    # 配布前チェック欄の点検文を検査対象に含めない（自分自身を誤検知するため）
    body = body_only(paragraphs)
    check_needs_confirmation(body, tables, report)
    check_placeholders(body, tables, report)
    check_todo_table(tables, report)
    check_predistribution_block(paragraphs, report)
    check_personal_names(body, tables, report)
    check_layout(z, tables, report)

    print(f"配布前チェック: {os.path.basename(args.docx)}")
    print(report.render())

    if args.preview:
        pages, problem = make_preview(args.docx, args.preview)
        print()
        if pages:
            print("プレビュー（目で見て確認する）:")
            for p in pages:
                print(f"  {p}")
            if problem:
                print(f"  ※ {problem}")
        else:
            print(f"プレビューは作れなかった: {problem}")
            print("  → フォントの当たり方と改ページ位置は Word で開いて確認する")

    return 1 if report.blockers else 0


if __name__ == "__main__":
    sys.exit(main())

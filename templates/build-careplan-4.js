// ケアプラン第4表（サービス担当者会議の要点）を .docx で作る。
//
// 第4表は法定様式で形が決まっているため、事業所ごとに変わる議事録用の
// build-meeting-minutes-template.js とは分けて、この様式専用に組んでいる。
//
//   npm install docx
//   node templates/build-careplan-4.js                 → 空の様式
//   node templates/build-careplan-4.js <データ.json>   → 記入済み
//   node templates/build-careplan-4.js <データ.json> -o <出力.docx>
//
// データJSONのキーは様式の欄名に合わせてある（templates/profiles/careplan-4.json 参照）。
// AI が推測した箇所には値の中に ［AI提案・要確認］ を書く。ケアマネジャーが
// 確認してタグを削除するまで提出しない。

const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Packer, PageOrientation, Paragraph,
  Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const MARGIN = 850;                       // 15mm
// A4 横。docx-js は LANDSCAPE 指定のとき幅と高さを内部で入れ替えるので、
// ここには「縦向きの寸法」を渡す。実際の用紙は 16838 x 11906 になる。
const PORTRAIT_W = 11906, PORTRAIT_H = 16838;
const CW = PORTRAIT_H - MARGIN * 2;       // 横にしたときの本文幅 15138
const LABEL_W = 2400;                     // 左のラベル列
const PAIR = Math.floor((CW - LABEL_W) / 3);
const SHOZOKU_W = Math.round(PAIR * 0.6);
const SHIMEI_W = PAIR - SHOZOKU_W;
// 端数は最後の氏名列で吸収して、合計を本文幅にそろえる
const ATTEND_COLS = [SHOZOKU_W, SHIMEI_W, SHOZOKU_W, SHIMEI_W, SHOZOKU_W,
                     CW - LABEL_W - (SHOZOKU_W + SHIMEI_W) * 2 - SHOZOKU_W];

const FONT = "游ゴシック";
const RED = "C00000";     // 様式上の赤字（利用者・家族の出席、※備考）
const GRAY = "808080";
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const LINE = { style: BorderStyle.SINGLE, size: 6, color: "000000" };

const run = (text, opts = {}) => new TextRun({ text, ...opts });
const para = (children, opts = {}) => new Paragraph({ children, ...opts });
const txt = (text, opts = {}) => para([run(text, opts)]);
// Word は TextRun 内の \n を改行しないので、行ごとに段落へ分ける
const multiline = (text, opts = {}) =>
  String(text).split("\n").map(line => txt(line, opts));

/** 罫線なしのレイアウト用セル。underline:true で記入欄の下線を引く */
const layoutCell = (children, width, { underline = false, align } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: { top: NONE, left: NONE, right: NONE, bottom: underline ? LINE : NONE },
    margins: { top: 20, bottom: 20, left: 40, right: 40 },
    verticalAlign: VerticalAlign.BOTTOM,
    children: children.length ? children : [new Paragraph({ alignment: align })],
  });

const layoutRow = cells => new TableRow({ children: cells });
const layoutTable = rows => new Table({
  width: { size: CW, type: WidthType.DXA },
  borders: { top: NONE, bottom: NONE, left: NONE, right: NONE,
             insideHorizontal: NONE, insideVertical: NONE },
  rows,
});

/** 本表のセル */
const gridCell = (children, width, opts = {}) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  columnSpan: opts.colSpan,
  rowSpan: opts.rowSpan,
  verticalAlign: opts.valign || VerticalAlign.CENTER,
  margins: { top: 60, bottom: 60, left: 110, right: 110 },
  children: children.length ? children : [new Paragraph("")],
});

function buildDoc(d) {
  const v = k => d[k] || "";
  const attendees = d["出席者"] || [];
  const at = i => attendees[i] || ["", ""];
  const lines = k => (d[k] || []).flatMap(t => multiline(t));

  // ---- 表題行 ----
  const titleRow = layoutTable([layoutRow([
    layoutCell([new Paragraph({
      border: { top: LINE, bottom: LINE, left: LINE, right: LINE },
      alignment: AlignmentType.CENTER,
      children: [run("第４表", { bold: true })],
    })], 1600),
    layoutCell([para([run("サービス担当者会議の要点", { bold: true, size: 26 })],
      { alignment: AlignmentType.CENTER })], CW - 1600 - 4200),
    layoutCell([para([run("作成年月日　"), run(v("作成年月日") || "　　　年　　月　　日")],
      { alignment: AlignmentType.RIGHT })], 4200),
  ])]);

  // ---- 利用者名／計画作成者 ----
  const nameRow = layoutTable([layoutRow([
    layoutCell([txt("利用者名")], 1300),
    layoutCell([txt(v("利用者名"))], 3600, { underline: true }),
    layoutCell([txt("殿")], 700),
    layoutCell([txt("居宅サービス計画作成者（担当者）氏名")], 4200),
    layoutCell([txt(v("計画作成者"))], CW - 1300 - 3600 - 700 - 4200, { underline: true }),
  ])]);

  // ---- 開催日／場所／時間／回数 ----
  const openRow = layoutTable([layoutRow([
    layoutCell([txt("開催日")], 1100),
    layoutCell([txt(v("開催日") || "　　　年　　月　　日")], 2900, { underline: true }),
    layoutCell([txt("開催場所")], 1300),
    layoutCell([txt(v("開催場所"))], 3600, { underline: true }),
    layoutCell([txt("開催時間")], 1300),
    layoutCell([txt(v("開催時間"))], 2600, { underline: true }),
    layoutCell([txt("開催回数")], 1300),
    layoutCell([txt(v("開催回数"))], CW - 1100 - 2900 - 1300 - 3600 - 1300 - 2600 - 1300,
      { underline: true }),
  ])]);

  // ---- 会議出席者のラベルセル（赤字の補助欄を含む。4行ぶん結合）----
  const attendLabel = gridCell([
    para([run("会議出席者")], { alignment: AlignmentType.CENTER }),
    new Paragraph(""),
    txt("利用者・家族の出席", { color: RED, size: 16 }),
    txt(`　本人：【 ${v("本人出席")} 】`, { color: RED, size: 16 }),
    txt(`　家族：【 ${v("家族出席")} 】`, { color: RED, size: 16 }),
    txt(`　（続柄：${v("続柄") || "　　　"}）`, { color: RED, size: 16 }),
    new Paragraph(""),
    txt(`※備考 ${v("備考")}`, { color: RED, size: 16 }),
  ], LABEL_W, { rowSpan: 4, valign: VerticalAlign.TOP });

  // ---- 本文欄のラベル（2行ぶんの記入枠を持つ）----
  const bodyRow = (label, sub, contentParas) => new TableRow({
    children: [
      gridCell([
        para([run(label)], { alignment: AlignmentType.CENTER }),
        ...(sub ? [para([run(sub, { size: 17 })], { alignment: AlignmentType.CENTER })] : []),
      ], LABEL_W),
      gridCell(contentParas, CW - LABEL_W, { colSpan: 6, valign: VerticalAlign.TOP }),
    ],
  });

  const mainTable = new Table({
    columnWidths: [LABEL_W, ...ATTEND_COLS],
    width: { size: CW, type: WidthType.DXA },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          attendLabel,
          ...["所 属(職種)", "氏　名", "所 属(職種)", "氏　名", "所 属(職種)", "氏　名"]
            .map((h, i) => gridCell([para([run(h)], { alignment: AlignmentType.CENTER })], ATTEND_COLS[i])),
        ],
      }),
      ...[0, 1, 2].map(r => new TableRow({
        children: [0, 1, 2].flatMap(c => {
          const [shozoku, shimei] = at(r * 3 + c);
          return [
            gridCell(multiline(shozoku), ATTEND_COLS[c * 2]),
            gridCell(multiline(shimei), ATTEND_COLS[c * 2 + 1]),
          ];
        }),
      })),
      bodyRow("検討した項目", null, lines("検討した項目")),
      bodyRow("検討内容", null, lines("検討内容")),
      bodyRow("結論", null, lines("結論")),
      bodyRow("残された課題", "(次回の開催時期)", lines("残された課題")),
    ],
  });

  const children = [];
  if (d["下書き注記"]) {
    children.push(txt(
      "※これは下書きです。［AI提案・要確認］ の箇所を確認し、タグを削除してから提出してください。",
      { size: 16, color: GRAY }));
  }
  children.push(titleRow, new Paragraph(""), nameRow, openRow, new Paragraph(""), mainTable);
  if (d["下書き注記"]) {
    children.push(new Paragraph(""), txt(
      "本下書きは介護支援専門員の確認・修正を前提としています。［AI提案・要確認］ の箇所を確認し、タグを削除してから提出してください。",
      { size: 16, color: GRAY }));
  }

  return new Document({
    creator: "dev-notes / 第4表",
    title: "第４表 サービス担当者会議の要点",
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: PORTRAIT_W, height: PORTRAIT_H, orientation: PageOrientation.LANDSCAPE },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children,
    }],
  });
}

async function main() {
  const args = process.argv.slice(2);
  const oAt = args.indexOf("-o");
  let out = null;
  if (oAt !== -1) { out = args[oAt + 1]; args.splice(oAt, 2); }
  const dataPath = args[0] || null;

  let data = {};
  if (dataPath) {
    try {
      data = JSON.parse(fs.readFileSync(dataPath, "utf8"));
    } catch (e) {
      throw new Error(`データを読めなかった: ${dataPath}\n  ${e.message}`);
    }
  }

  const outPath = out || path.join(__dirname, "built",
    dataPath ? `${path.basename(dataPath, ".json")}.docx` : "careplan-4-blank.docx");
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  const buf = await Packer.toBuffer(buildDoc(data));
  fs.writeFileSync(outPath, buf);
  console.log(`wrote ${outPath} (${buf.length} bytes)${dataPath ? ` ← ${dataPath}` : " ← 空の様式"}`);
}

main().catch(e => { console.error(`エラー: ${e.message}`); process.exit(1); });

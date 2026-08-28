// 議事録テンプレート（.docx）のビルドスクリプト
//
//   npm install docx
//   node templates/build-meeting-minutes-template.js
//
// 出力: templates/meeting-minutes-template.docx
// 手順書: docs/meeting-minutes-automation.md

const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, HeadingLevel,
  LevelFormat, PageNumber, Packer, Paragraph, ShadingType, Table, TableCell,
  TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

// A4 縦・余白 20mm。本文の実効幅は 11906 - 1134*2 = 9638 DXA。
const CONTENT_WIDTH = 9638;
const MARGIN = 1134;

const FONT = "游ゴシック";
const TEAL = "0F766E";   // 見出し（となりにAI ブランドカラー）
const GRAY = "808080";   // 差し替え前提のプレースホルダ
const RULE = "D0D0D0";
const HEAD_BG = "E8F2F1"; // 表のヘッダ行

/** 差し替え前提の 〔…〕 プレースホルダ段落 */
const placeholder = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [new TextRun({ text, color: GRAY })],
  });

const cell = (children, { width, bg, header = false } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: bg ? { type: ShadingType.CLEAR, fill: bg, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: children.map((text) =>
      new Paragraph({
        children: [new TextRun({ text, bold: header, color: header ? undefined : GRAY })],
      })
    ),
  });

/** ヘッダ行 + 空の記入行を持つ表 */
const gridTable = (columnWidths, headers, rows) =>
  new Table({
    columnWidths,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) =>
          cell([h], { width: columnWidths[i], bg: HEAD_BG, header: true })
        ),
      }),
      ...rows.map((row) =>
        new TableRow({
          children: row.map((text, i) => cell([text], { width: columnWidths[i] })),
        })
      ),
    ],
  });

/** 会議情報の 2 列表（左が項目名） */
const infoTable = (pairs) => {
  const widths = [2200, CONTENT_WIDTH - 2200];
  return new Table({
    columnWidths: widths,
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    rows: pairs.map(([label, value]) =>
      new TableRow({
        children: [
          cell([label], { width: widths[0], bg: HEAD_BG, header: true }),
          cell([value], { width: widths[1] }),
        ],
      })
    ),
  });
};

const heading = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text })],
  });

const doc = new Document({
  creator: "dev-notes / meeting-minutes template",
  title: "議事録テンプレート",
  description: "録音→文字起こし→整形→Word の流れで使う議事録の雛形",
  styles: {
    default: {
      document: { run: { font: FONT, size: 21 } }, // 21 half-points = 10.5pt
      heading1: {
        run: { font: FONT, size: 26, bold: true, color: TEAL },
        paragraph: { spacing: { before: 320, after: 160 } },
      },
    },
    paragraphStyles: [
      {
        id: "DocTitle",
        name: "Doc Title",
        basedOn: "Normal",
        run: { font: FONT, size: 36, bold: true, color: TEAL },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 320 } },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "note-bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360, hanging: 240 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: { margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 4 } },
              children: [new TextRun({ text: "〔会議名〕", color: GRAY, size: 18 })],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ children: ["- ", PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES, " -"], size: 18 }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({ style: "DocTitle", children: [new TextRun("議 事 録")] }),

        infoTable([
          ["会議名", "〔例〕○○事業所 サービス担当者会議"],
          ["日時", "〔例〕2026年8月26日（水）14:00〜15:00"],
          ["場所", "〔例〕○○事業所 会議室 ／ オンライン（Zoom）"],
          ["出席者", "〔所属・氏名を列挙〕"],
          ["欠席者", "〔いなければ「なし」〕"],
          ["記録者", "〔氏名〕"],
        ]),

        heading("1. 議題"),
        placeholder("〔議題を番号付きで列挙する〕"),

        heading("2. 議論の要旨"),
        placeholder("〔議題ごとに 3〜5 行。誰が何を言ったかではなく、何が論点でどう整理されたかを書く〕"),

        heading("3. 決定事項"),
        gridTable(
          [800, 6000, 2838],
          ["No", "決定事項", "補足・根拠"],
          [["1", "", ""], ["2", "", ""], ["3", "", ""]]
        ),
        placeholder("※「決まったこと」だけを書く。検討中のものは 2. の要旨側に残す。", { spacing: { before: 120 } }),

        heading("4. ToDo"),
        gridTable(
          [800, 4838, 2000, 2000],
          ["No", "内容", "担当", "期限"],
          [["1", "", "", ""], ["2", "", "", ""], ["3", "", "", ""]]
        ),
        placeholder("※担当と期限が空欄の ToDo は残さない。決められないなら「次回までに決める」を ToDo にする。", { spacing: { before: 120 } }),

        heading("5. 次回予定"),
        placeholder("〔日時・場所・主な議題〕"),

        heading("6. ［要確認］一覧"),
        placeholder("〔文字起こしから読み取れなかった箇所・裏取りが必要な数字をここに集める。空になったら配布可〕"),

        new Paragraph({
          spacing: { before: 480, after: 120 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 8 } },
          children: [new TextRun({ text: "配布前チェック（この欄は配布時に削除する）", bold: true, size: 18, color: TEAL })],
        }),
        ...[
          "日付・会議名・出席者が合っているか",
          "［要確認］が残っていないか（残っていれば音声に戻って確認する）",
          "決定事項が「決まったこと」だけになっているか",
          "ToDo に担当と期限が全部入っているか",
          "利用者名などの個人情報が残っていないか（A様・B様に置き換える）",
        ].map((text) =>
          new Paragraph({
            numbering: { reference: "note-bullets", level: 0 },
            children: [new TextRun({ text, size: 18, color: GRAY })],
          })
        ),
      ],
    },
  ],
});

const out = path.join(__dirname, "meeting-minutes-template.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("wrote", out, buf.length, "bytes");
});

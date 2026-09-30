// 議事録テンプレート（.docx）のビルドスクリプト
//
// 書式プロファイル（JSON）を読んで、その体裁の空テンプレートを出す。
// 事業所ごとに違う様式へ合わせるための仕組み。プロファイルの作り方は
// docs/meeting-minutes-automation.md の「事業所のフォーマットに合わせる」を参照。
//
//   npm install docx
//   node templates/build-meeting-minutes-template.js
//     → templates/profiles/default.json から templates/meeting-minutes-template.docx
//
//   node templates/build-meeting-minutes-template.js templates/profiles/sample-office.json
//     → templates/built/sample-office.docx
//
//   node templates/build-meeting-minutes-template.js <profile.json> -o <out.docx>
//     → 出力先を明示する
//
// 実在の事業所のプロファイルは company/secretary/minutes/formats/ に置く
// （このリポジトリは公開のため、事業所名を含むものは入れない）。

const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, HeadingLevel,
  LevelFormat, PageNumber, PageOrientation, Packer, Paragraph, ShadingType,
  Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const MM_TO_DXA = 56.6929; // 1mm ≒ 56.69 DXA（1440 DXA = 1インチ）
const A4_SHORT = 11906;    // A4 210mm
const A4_LONG = 16838;     // A4 297mm

const GRAY = "808080";     // 差し替え前提のプレースホルダ
const RULE = "D0D0D0";
const HEAD_BG = "E8F2F1";  // 表のヘッダ行

// ---------------------------------------------------------------- プロファイル

function loadProfile(file) {
  let profile;
  try {
    profile = JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    throw new Error(`プロファイルを読めなかった: ${file}\n  ${e.message}`);
  }
  if (!Array.isArray(profile.sections) || profile.sections.length === 0) {
    throw new Error(`${file}: sections が空。最低1つは節が必要`);
  }
  if (!Array.isArray(profile.infoFields) || profile.infoFields.length === 0) {
    throw new Error(`${file}: infoFields が空。会議情報の項目を1つ以上入れる`);
  }
  for (const s of profile.sections) {
    if (!s.heading) throw new Error(`${file}: heading の無い節がある`);
    if (s.type === "table" && !Array.isArray(s.columns)) {
      throw new Error(`${file}: 表「${s.heading}」に columns が無い`);
    }
  }
  return profile;
}

/**
 * improvements で、お手本に無かった「こちらの改善点」を足す。
 * 項目名と並びはお手本のまま、足りない列と節だけを補う方針。
 */
function applyImprovements(profile) {
  const imp = profile.improvements || {};
  const sections = profile.sections.map((s) => ({ ...s }));

  // ToDo 表に担当／期限が無ければ列を足す（空欄のまま残せない作りにする）
  if (imp.todoOwnerDeadline) {
    const todo = sections.find((s) => s.role === "todo" && s.type === "table");
    if (todo) {
      todo.columns = [...todo.columns];
      for (const label of ["担当", "期限"]) {
        if (!todo.columns.some((c) => c.label === label)) {
          todo.columns.push({ label, weight: 3 });
        }
      }
    }
  }

  // ［要確認］一覧が無ければ末尾に足す（聞き取れなかった箇所の置き場）
  if (imp.needsCheckSection && !sections.some((s) => s.role === "needs-check")) {
    sections.push({
      heading: `${sections.length + 1}. ［要確認］一覧`,
      type: "text",
      role: "needs-check",
      placeholder: "〔文字起こしから読み取れなかった箇所・裏取りが必要な数字をここに集める。空になったら配布可〕",
    });
  }

  return { ...profile, sections };
}

// -------------------------------------------------------------------- 部品

const placeholder = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [new TextRun({ text, color: GRAY })],
  });

const cell = (text, { width, bg, header = false } = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: bg ? { type: ShadingType.CLEAR, fill: bg, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [
      new Paragraph({
        children: [new TextRun({ text, bold: header, color: header ? undefined : GRAY })],
      }),
    ],
  });

/**
 * 列幅を DXA で確定させる。width 指定があればそれを、無ければ weight で按分。
 * 合計は必ず本文幅に一致させる（ずれると Google ドキュメントで崩れる）。
 */
function resolveWidths(columns, contentWidth) {
  const hasExplicit = columns.some((c) => typeof c.width === "number");
  let widths;
  if (hasExplicit) {
    widths = columns.map((c) => c.width || 0);
  } else {
    const weights = columns.map((c) => c.weight || 1);
    const total = weights.reduce((a, b) => a + b, 0);
    widths = weights.map((w) => Math.round((contentWidth * w) / total));
  }
  // 端数は一番広い列に寄せて、合計を本文幅にそろえる
  const diff = contentWidth - widths.reduce((a, b) => a + b, 0);
  if (diff !== 0) {
    const widest = widths.indexOf(Math.max(...widths));
    widths[widest] += diff;
  }
  return widths;
}

const gridTable = (columns, rowCount, contentWidth) => {
  const widths = resolveWidths(columns, contentWidth);
  return new Table({
    columnWidths: widths,
    width: { size: contentWidth, type: WidthType.DXA },
    rows: [
      new TableRow({
        tableHeader: true,
        children: columns.map((c, i) =>
          cell(c.label, { width: widths[i], bg: HEAD_BG, header: true })
        ),
      }),
      ...Array.from({ length: rowCount }, (_, r) =>
        new TableRow({
          children: columns.map((c, i) =>
            // 1列目が No なら連番を薄く入れておく
            cell(c.label === "No" ? String(r + 1) : "", { width: widths[i] })
          ),
        })
      ),
    ],
  });
};

const infoTable = (fields, contentWidth, fixedLabelWidth) => {
  // infoLabelWidth を指定すればラベル列の幅を固定できる（未指定なら本文幅の 26%）
  const labelWidth = fixedLabelWidth || Math.min(2600, Math.round(contentWidth * 0.26));
  const widths = [labelWidth, contentWidth - labelWidth];
  return new Table({
    columnWidths: widths,
    width: { size: contentWidth, type: WidthType.DXA },
    rows: fields.map((f) =>
      new TableRow({
        children: [
          cell(f.label, { width: widths[0], bg: HEAD_BG, header: true }),
          cell(f.placeholder || "", { width: widths[1] }),
        ],
      })
    ),
  });
};

// ---------------------------------------------------------------- 組み立て

function buildDocument(profile) {
  const page = profile.page || {};
  const landscape = page.orientation === "landscape";
  const margin = Math.round((page.marginMm ?? 20) * MM_TO_DXA);
  // 横向きでも縦向きの寸法を渡す（docx-js が内部で入れ替える）
  const contentWidth = (landscape ? A4_LONG : A4_SHORT) - margin * 2;

  const font = (profile.font && profile.font.name) || "游ゴシック";
  const size = Math.round(((profile.font && profile.font.sizePt) || 10.5) * 2);
  const headingColor = profile.headingColor || "0F766E";

  const children = [
    new Paragraph({ style: "DocTitle", children: [new TextRun(profile.title || "議 事 録")] }),
    infoTable(profile.infoFields, contentWidth, profile.infoLabelWidth),
  ];

  for (const s of profile.sections) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 320, after: 160 },
        children: [new TextRun({ text: s.heading })],
      })
    );
    if (s.type === "table") {
      children.push(gridTable(s.columns, s.rows || 3, contentWidth));
      if (s.note) children.push(placeholder(s.note, { spacing: { before: 120 } }));
    } else {
      children.push(placeholder(s.placeholder || "〔　　　　　〕"));
    }
  }

  if (profile.improvements && profile.improvements.preDistributionChecklist) {
    children.push(
      new Paragraph({
        spacing: { before: 480, after: 120 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 8 } },
        children: [
          new TextRun({ text: "配布前チェック（この欄は配布時に削除する）", bold: true, size: 18, color: headingColor }),
        ],
      }),
      ...[
        "会議情報（日付・場所・出席者）が合っているか",
        "［要確認］が残っていないか（残っていれば音声に戻って確認する）",
        "決定事項が「決まったこと」だけになっているか",
        "ToDo に担当と期限が全部入っているか",
        "利用者名などの個人情報が残っていないか（A様・B様に置き換える）",
      ].map((text) =>
        new Paragraph({
          numbering: { reference: "note-bullets", level: 0 },
          children: [new TextRun({ text, size: 18, color: GRAY })],
        })
      )
    );
  }

  return new Document({
    creator: "dev-notes / meeting-minutes template",
    title: `議事録テンプレート（${profile.profileName || "profile"}）`,
    description: profile.note || "録音→文字起こし→整形→Word の流れで使う議事録の雛形",
    styles: {
      default: {
        document: { run: { font, size } },
        heading1: {
          run: { font, size: Math.round(size * 1.24), bold: true, color: headingColor },
          paragraph: { spacing: { before: 320, after: 160 } },
        },
      },
      paragraphStyles: [
        {
          id: "DocTitle",
          name: "Doc Title",
          basedOn: "Normal",
          run: { font, size: Math.round(size * 1.7), bold: true, color: headingColor },
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
          page: {
            size: landscape
              ? { width: A4_SHORT, height: A4_LONG, orientation: PageOrientation.LANDSCAPE }
              : undefined,
            margin: { top: margin, right: margin, bottom: margin, left: margin },
          },
        },
        headers: profile.headerText
          ? {
              default: new Header({
                children: [
                  new Paragraph({
                    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 4 } },
                    children: [new TextRun({ text: profile.headerText, color: GRAY, size: 18 })],
                  }),
                ],
              }),
            }
          : undefined,
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: ["- ", PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES, " -"],
                    size: 18,
                  }),
                ],
              }),
            ],
          }),
        },
        children,
      },
    ],
  });
}

// -------------------------------------------------------------------- CLI

function parseArgs(argv) {
  const args = argv.slice(2);
  const oAt = args.indexOf("-o");
  let out = null;
  if (oAt !== -1) {
    out = args[oAt + 1];
    if (!out) throw new Error("-o の後に出力パスが無い");
    args.splice(oAt, 2);
  }
  return { profilePath: args[0] || null, out };
}

// async にしておくと、プロファイル検証の同期 throw も下の catch で拾える
async function main() {
  const { profilePath, out } = parseArgs(process.argv);
  const here = __dirname;
  const file = profilePath || path.join(here, "profiles", "default.json");
  const profile = applyImprovements(loadProfile(file));

  const name = path.basename(file, ".json");
  const outPath =
    out ||
    (name === "default"
      ? path.join(here, "meeting-minutes-template.docx")
      : path.join(here, "built", `${name}.docx`));

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  return Packer.toBuffer(buildDocument(profile)).then((buf) => {
    fs.writeFileSync(outPath, buf);
    console.log(`wrote ${outPath} (${buf.length} bytes) ← ${file}`);
  });
}

main().catch((e) => {
  console.error(`エラー: ${e.message}`);
  process.exit(1);
});

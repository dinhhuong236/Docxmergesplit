// Tạo các file .docx mẫu để kiểm thử merge/split.
const fs = require("fs");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Header, Footer, ImageRun,
  Table, TableRow, TableCell, WidthType, AlignmentType, LevelFormat, PageOrientation,
  FootnoteReferenceRun, PageNumber, PageBreak, BorderStyle, ShadingType } = d;

function styles(h1color, h1size, normalFont) {
  return {
    default: { document: { run: { font: normalFont, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: h1size, bold: true, color: h1color, font: normalFont },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, italics: true, font: normalFont },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      { id: "GhiChu", name: "Ghi chu", basedOn: "Normal", run: { color: "C00000", italics: true } },
    ],
  };
}
const numbering = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "nums", levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }] },
  ],
};
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
function table(label) {
  const cell = (t, fill) => new TableCell({ borders: { top: border, bottom: border, left: border, right: border },
    width: { size: 4680, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph(t)] });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
    rows: [new TableRow({ children: [cell(label + " cột 1", "D5E8F0"), cell(label + " cột 2", "D5E8F0")] }),
           new TableRow({ children: [cell("a"), cell("b")] })] });
}
const img = (f) => new ImageRun({ type: "png", data: fs.readFileSync(f),
  transformation: { width: 160, height: 80 }, altText: { title: f, description: f, name: f } });

function doc({ name, h1color, h1size, font, landscape, image, chapters }) {
  const children = [new Paragraph({ children: [new TextRun({ text: `Trang bìa của ${name}`, bold: true, size: 40 })] })];
  chapters.forEach((ch, ci) => {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(ch)] }));
    children.push(new Paragraph({ children: [new TextRun(`Nội dung ${ch} của ${name}.`),
      new FootnoteReferenceRun(ci + 1)] }));
    children.push(new Paragraph({ style: "GhiChu", children: [new TextRun("Đây là ghi chú dùng style Ghi chu.")] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`Mục con của ${ch}`)] }));
    for (let i = 0; i < 3; i++)
      children.push(new Paragraph({ numbering: { reference: "nums", level: 0 }, children: [new TextRun(`Bước ${i + 1} trong ${ch}`)] }));
    children.push(new Paragraph({ numbering: { reference: "nums", level: 1 }, children: [new TextRun("ý phụ")] }));
    children.push(new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun("gạch đầu dòng")] }));
    children.push(table(ch));
    children.push(new Paragraph({ children: [img(image)] }));
    if (ci < chapters.length - 1) children.push(new Paragraph({ children: [new PageBreak()] }));
  });
  const footnotes = {};
  chapters.forEach((ch, ci) => { footnotes[ci + 1] = { children: [new Paragraph(`Chú thích ${ci + 1} của ${name}`)] }; });
  return new Document({
    styles: styles(h1color, h1size, font), numbering, footnotes,
    sections: [{
      properties: { page: { size: landscape ? { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE }
                                            : { width: 11906, height: 16838 },
                            margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun(`HEADER ${name}`)] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun(`${name} – trang `), new TextRun({ children: [PageNumber.CURRENT] })] })] }) },
      children,
    }],
  });
}

const specs = [
  { file: "A.docx", name: "File A", h1color: "1F4E79", h1size: 32, font: "Times New Roman", landscape: false, image: "red.png", chapters: ["Chương 1", "Chương 2"] },
  { file: "B.docx", name: "File B", h1color: "C55A11", h1size: 36, font: "Times New Roman", landscape: true, image: "red.png", chapters: ["Chương 3"] },
  { file: "C.docx", name: "File C", h1color: "C55A11", h1size: 36, font: "Times New Roman", landscape: false, image: "blue.png", chapters: ["Chương 4", "Chương 5"] },
  { file: "D.docx", name: "File D", h1color: "1F4E79", h1size: 32, font: "Arial", landscape: false, image: "blue.png", chapters: ["Chương 6"] },
];
(async () => {
  for (const s of specs) fs.writeFileSync(s.file, await Packer.toBuffer(doc(s)));
  console.log("ok");
})();

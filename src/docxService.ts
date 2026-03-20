import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  HeadingLevel,
  PageBreak,
  BorderStyle,
  TableLayoutType,
  HorizontalPositionAlign,
  VerticalPositionAlign,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  TextWrappingType,
  TextWrappingSide,
  Footer,
  Header,
  PageNumber,
  ExternalHyperlink,
  VerticalAlignSection,
} from "docx";
import * as fs from "fs";
import * as path from "path";
import { AiOutput, GenerateRequest } from "./schema";

// --- Constants ---
const PURPLE = "7030A0";
const FONT = "Aptos";
const SIZE_H1 = 36;    // 18pt
const SIZE_H2 = 32;    // 16pt
const SIZE_H3 = 28;    // 14pt
const SIZE_BODY = 24;  // 12pt

// --- Helpers ---

function heading1(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, bold: true, size: SIZE_H1, font: FONT, color: PURPLE })],
  });
}

function heading2(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    children: [new TextRun({ text, bold: true, size: SIZE_H2, font: FONT, color: PURPLE })],
  });
}

function heading3(text: string): Paragraph {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, size: SIZE_H3, font: FONT, color: PURPLE })],
  });
}

function bodyParagraph(text: string): Paragraph {
  return new Paragraph({
    spacing: { after: 150 },
    children: [new TextRun({ text, size: SIZE_BODY, font: FONT })],
  });
}

function bulletPoint(text: string): Paragraph {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: SIZE_BODY, font: FONT })],
  });
}

function hyperlinkBullet(label: string, url: string): Paragraph {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 80 },
    children: [
      new ExternalHyperlink({
        link: url,
        children: [
          new TextRun({ text: label, size: SIZE_BODY, font: FONT, color: "0563C1", underline: { type: "single" } }),
        ],
      }),
    ],
  });
}

/** If text looks like a URL, render as clickable hyperlink bullet; otherwise plain bullet */
function smartBullet(text: string): Paragraph {
  const trimmed = text.trim();
  if (trimmed.startsWith("http://") || trimmed.startsWith("https://")) {
    return hyperlinkBullet(trimmed, trimmed);
  }
  return bulletPoint(trimmed);
}

const cellBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
  left: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
  right: { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" },
};

function headerCell(text: string): TableCell {
  return new TableCell({
    borders: cellBorders,
    shading: { fill: "D9D9D9" },
    children: [
      new Paragraph({
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text, bold: true, size: SIZE_BODY, font: FONT })],
      }),
    ],
  });
}

function dataCell(text: string): TableCell {
  return new TableCell({
    borders: cellBorders,
    children: [
      new Paragraph({
        spacing: { before: 40, after: 40 },
        children: [new TextRun({ text, size: SIZE_BODY, font: FONT })],
      }),
    ],
  });
}

function makeTable(headers: string[], rows: string[][]): Table {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ children: headers.map((h) => headerCell(h)) }),
      ...rows.map(
        (row) => new TableRow({ children: row.map((cell) => dataCell(cell)) })
      ),
    ],
  });
}

// --- Title Page Helpers ---

function titleRunWhite(text: string, size: number, bold = true): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 80 },
    indent: { left: 720 },
    children: [new TextRun({ text, bold, size, font: FONT, color: "FFFFFF" })],
  });
}

function loadAsset(filename: string): Buffer {
  const assetPath = path.resolve(__dirname, "..", "assets", filename);
  return fs.readFileSync(assetPath);
}

function buildTitlePageChildren(data: AiOutput, currentDate: string): (Paragraph | Table)[] {
  const bgImageData = loadAsset("background.jpg");
  const logoImageData = loadAsset("commvault-logo.jpg");

  const titleChildren: (Paragraph | Table)[] = [];

  // Background image — A4 size (794x1123 px at 96 DPI)
  titleChildren.push(
    new Paragraph({
      children: [
        new ImageRun({
          type: "jpg",
          data: bgImageData,
          transformation: { width: 794, height: 1123 },
          floating: {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.PAGE,
              offset: 0,
            },
            verticalPosition: {
              relative: VerticalPositionRelativeFrom.PAGE,
              offset: 0,
            },
            behindDocument: true,
            wrap: {
              type: TextWrappingType.NONE,
              side: TextWrappingSide.BOTH_SIDES,
            },
            allowOverlap: true,
          },
        }),
      ],
    })
  );

  // Vertical space from top
  for (let i = 0; i < 4; i++) {
    titleChildren.push(emptyLine());
  }

  // Logo at top-left, same indent as text
  titleChildren.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: { left: 720 },
      spacing: { after: 200 },
      children: [
        new ImageRun({
          type: "jpg",
          data: logoImageData,
          transformation: { width: 200, height: 45 },
        }),
      ],
    })
  );

  // Client Name — main heading
  titleChildren.push(titleRunWhite(data.customerName, 72, true));
  // Commvault POC Document — sub heading
  titleChildren.push(titleRunWhite("Commvault POC Document", 40, false));
  // Date — sub heading
  titleChildren.push(titleRunWhite(currentDate, 32, false));

  return titleChildren;
}

function pageBreak(): Paragraph {
  return new Paragraph({ children: [new PageBreak()] });
}

function emptyLine(): Paragraph {
  return new Paragraph({ spacing: { after: 100 }, children: [] });
}

// --- Document Generation ---

export async function generateDocx(data: AiOutput, input: GenerateRequest): Promise<Buffer> {
  const currentDate = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  const children: (Paragraph | Table)[] = [];

  // ========== TITLE PAGE (separate section) ==========
  const titlePageChildren = buildTitlePageChildren(data, currentDate);

  // ========== PAGE 2+: CONTENT ==========
  children.push(heading1("Index"));
  const indexItems = [
    "1. Version History",
    "2. POC Team",
    "3. Executive Summary",
    "4. Objective",
    "5. Scope of the POC",
    "6. Workloads in Scope",
    "7. Hardware Requirements",
    "8. Networking and Firewall Requirements",
    "9. POC Prerequisites",
    "10. Test Cases",
    "11. Timeline",
    "12. POC Closure and Handover",
  ];
  indexItems.forEach((item) =>
    children.push(
      new Paragraph({
        spacing: { after: 150 },
        children: [new TextRun({ text: item, bold: true, size: SIZE_H3, font: FONT, color: "000000" })],
      })
    )
  );
  children.push(pageBreak());

  // ========== SECTION: Version History ==========
  children.push(heading1("1. Version History"));
  children.push(
    makeTable(
      ["Version #", "Revision Date", "Contributor's Name", "Revision Description"],
      [["1", currentDate, input.salesRepName, "Initial POC Draft"]]
    )
  );
  children.push(emptyLine());

  // ========== SECTION: POC Team ==========
  children.push(heading1("2. POC Team"));
  children.push(
    makeTable(
      ["Name", "Role", "Contact Info / Email"],
      [[input.salesRepName, input.salesRepRole, input.salesRepEmail]]
    )
  );
  children.push(emptyLine());

  // ========== SECTION: Executive Summary ==========
  children.push(heading1("3. Executive Summary"));
  children.push(bodyParagraph(data.executiveSummary));
  children.push(emptyLine());

  // ========== SECTION: Objective ==========
  children.push(heading1("4. Objective"));
  children.push(bodyParagraph(data.objective));
  children.push(emptyLine());

  // ========== SECTION: Scope of the POC ==========
  children.push(heading1("5. Scope of the POC"));

  children.push(heading2("In Scope"));
  data.scope.inScope.forEach((item) => children.push(bulletPoint(item)));

  children.push(heading2("Out of Scope"));
  data.scope.outOfScope.forEach((item) => children.push(bulletPoint(item)));

  children.push(heading2("Assumptions"));
  data.scope.assumptions.forEach((item) => children.push(bulletPoint(item)));
  children.push(emptyLine());

  // ========== SECTION: Workloads in Scope ==========
  children.push(heading1("6. Workloads in Scope"));
  children.push(
    makeTable(
      ["Workload Category", "Workload", "Deployment Type", "Location"],
      data.workloadsInScopeTable.map((r) => [r.workloadCategory, r.workload, r.deploymentType, r.location])
    )
  );
  children.push(emptyLine());

  // ========== SECTION: Hardware Requirements ==========
  children.push(heading1("7. Hardware Requirements"));

  children.push(heading2("Components"));
  children.push(
    makeTable(
      ["Component", "Quantity", "Role"],
      data.hardwareRequirements.componentsTable.map((r) => [r.component, String(r.quantity), r.role])
    )
  );
  children.push(emptyLine());

  children.push(heading2("Hardware Sizing"));
  children.push(
    makeTable(
      ["Component", "CPU", "Memory", "Storage", "OS"],
      data.hardwareRequirements.hardwareSizingTable.map((r) => [r.component, r.cpu, r.memory, r.storage, r.os])
    )
  );
  children.push(emptyLine());

  if (data.hardwareRequirements.documentationLinks.length > 0) {
    children.push(heading2("Documentation Links"));
    data.hardwareRequirements.documentationLinks.forEach((link) => children.push(smartBullet(link)));
    children.push(emptyLine());
  }

  // ========== SECTION: Networking and Firewall Requirements ==========
  children.push(heading1("8. Networking and Firewall Requirements"));
  children.push(
    makeTable(
      ["Port", "Protocol", "Purpose"],
      data.networkingFirewallTable.map((r) => [r.port, r.protocol, r.purpose])
    )
  );
  children.push(emptyLine());

  children.push(heading2("Documentation Links"));
  children.push(heading3("Recommended Antivirus Exclusions"));
  children.push(
    hyperlinkBullet(
      "Windows Exclusions",
      "https://documentation.commvault.com/11.40/software/antivirus_exclusions_for_windows.html"
    )
  );
  children.push(
    hyperlinkBullet(
      "Mac/Linux Exclusions",
      "https://documentation.commvault.com/11.40/software/recommended_antivirus_exclusions_for_unix_and_mac.html"
    )
  );
  children.push(emptyLine());

  // ========== SECTION: POC Prerequisites (GLOBAL TABLE) ==========
  children.push(heading1("9. POC Prerequisites"));
  children.push(
    makeTable(
      ["Category", "Prerequisite", "Customer Responsibility"],
      data.prerequisitesTable.map((r) => [r.category, r.prerequisite, r.customerResponsibility])
    )
  );
  children.push(emptyLine());

  // ========== SECTION: Test Cases (ALL workloads under one heading) ==========
  children.push(heading1("10. Test Cases"));
  for (let idx = 0; idx < data.testCasesByWorkload.length; idx++) {
    const tc = data.testCasesByWorkload[idx];
    children.push(heading2(`Workload ${idx + 1} : ${tc.workloadName}`));
    if (tc.rows.length > 0) {
      children.push(
        makeTable(
          ["Test Case", "Description", "Comments", "Result"],
          tc.rows.map((r) => [r.testCase, r.description, r.comments, r.result])
        )
      );
      children.push(emptyLine());
    }
  }

  // ========== SECTION: Timeline (ALL workloads under one heading) ==========
  children.push(heading1("11. Timeline"));
  for (let idx = 0; idx < data.timelinesByWorkload.length; idx++) {
    const tl = data.timelinesByWorkload[idx];
    children.push(heading2(`Workload ${idx + 1} : ${tl.workloadName}`));
    if (tl.rows.length > 0) {
      children.push(
        makeTable(
          ["Phase", "Date", "Workload", "Task"],
          tl.rows.map((r) => [r.phase, r.date, r.workload, r.task])
        )
      );
      children.push(emptyLine());
    }
  }

  // ========== SECTION: POC Closure and Handover ==========
  children.push(heading1("12. POC Closure and Handover"));
  children.push(bodyParagraph(data.pocClosureAndHandover));

  // ========== SIGNATURE BLOCK ==========
  children.push(emptyLine());
  children.push(emptyLine());

  const noBorders = {
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  };

  function sigCell(lines: { text: string; bold?: boolean }[]): TableCell {
    return new TableCell({
      borders: noBorders,
      width: { size: 50, type: WidthType.PERCENTAGE },
      children: lines.map(
        (line) =>
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: line.text, bold: line.bold ?? false, size: SIZE_BODY, font: FONT })],
          })
      ),
    });
  }

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          children: [
            sigCell([{ text: "________________________________" }, { text: "Name (Print)", bold: true }]),
            sigCell([{ text: "________________________________" }, { text: "Title (Print)", bold: true }]),
          ],
        }),
        new TableRow({
          children: [
            sigCell([{ text: "" }]),
            sigCell([{ text: "" }]),
          ],
        }),
        new TableRow({
          children: [
            sigCell([{ text: "________________________________" }, { text: "Signature", bold: true }]),
            sigCell([{ text: "________________________________" }, { text: "Date", bold: true }]),
          ],
        }),
      ],
    })
  );

  // ========== Build Document (two sections: title page + content) ==========
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 0, bottom: 0, left: 720, right: 720 },
            size: { width: 11906, height: 16838 },
          },
        },
        children: titlePageChildren as any[],
      },
      {
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
            size: { width: 11906, height: 16838 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new ImageRun({
                    type: "jpg",
                    data: loadAsset("commvault-logo.jpg"),
                    transformation: { width: 150, height: 34 },
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({ text: "Page | ", size: SIZE_BODY, font: FONT }),
                  new TextRun({ children: [PageNumber.CURRENT], size: SIZE_BODY, font: FONT }),
                ],
              }),
            ],
          }),
        },
        children: children as any[],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer as Buffer;
}

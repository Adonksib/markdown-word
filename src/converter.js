import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import * as docx from "docx";

const LIST_LEVELS = [0, 1, 2, 3, 4, 5];

const YAHEI = {
  ascii: "微软雅黑",
  hAnsi: "微软雅黑",
  eastAsia: "微软雅黑",
  cs: "微软雅黑",
};

const CONSOLAS = {
  ascii: "Consolas",
  hAnsi: "Consolas",
  eastAsia: "微软雅黑",
  cs: "Consolas",
};

const STYLES = {
  default: {
    document: {
      run: { font: YAHEI, size: 22 },
    },
  },
  paragraphStyles: [
    {
      id: "BodyText",
      name: "Body Text",
      basedOn: "Normal",
      next: "BodyText",
      quickFormat: true,
      paragraph: {
        spacing: { after: 0, line: 240 },
        indent: { left: 100 },
      },
      run: { font: YAHEI, size: 19 },
    },
    {
      id: "Title",
      name: "Title",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      paragraph: {
        spacing: { after: 120, line: 240 },
        indent: { left: 100 },
      },
      run: { font: YAHEI, size: 39, bold: true },
    },
    {
      id: "Heading1",
      name: "heading 1",
      basedOn: "Normal",
      next: "BodyText",
      quickFormat: true,
      paragraph: {
        spacing: { before: 200, after: 0, line: 240 },
        indent: { left: 100 },
        outlineLevel: 1,
      },
      run: { font: YAHEI, size: 29, bold: true },
    },
    {
      id: "Heading2",
      name: "heading 2",
      basedOn: "Normal",
      next: "BodyText",
      quickFormat: true,
      paragraph: {
        spacing: { before: 120, after: 0, line: 240 },
        indent: { left: 355, hanging: 256 },
        outlineLevel: 2,
      },
      run: { font: YAHEI, size: 23, bold: true },
    },
    {
      id: "ListParagraph",
      name: "List Paragraph",
      basedOn: "Normal",
      next: "ListParagraph",
      quickFormat: true,
      paragraph: {
        spacing: { after: 0, line: 240 },
        indent: { left: 490, hanging: 277 },
      },
      run: { font: YAHEI, size: 19 },
    },
    {
      id: "CodeBlock",
      name: "Code Block",
      basedOn: "Normal",
      next: "Normal",
      quickFormat: true,
      paragraph: {
        spacing: { before: 60, after: 60, line: 240 },
        indent: { left: 100 },
      },
      run: { font: CONSOLAS, size: 19 },
    },
  ],
};

const PAGE = {
  size: { width: 12240, height: 15840 },
  margin: {
    top: 860,
    right: 720,
    bottom: 280,
    left: 740,
    header: 720,
    footer: 720,
    gutter: 0,
  },
};

function bulletNumbering() {
  return {
    reference: "unordered-list",
    levels: LIST_LEVELS.map((level) => ({
      level,
      format: docx.LevelFormat.BULLET,
      text: "\u2022",
      alignment: docx.AlignmentType.START,
    })),
  };
}

function orderedNumbering(reference) {
  return {
    reference,
    levels: LIST_LEVELS.map((level) => ({
      level,
      format: docx.LevelFormat.DECIMAL,
      text: "%1.",
      alignment: docx.AlignmentType.START,
    })),
  };
}

function collectOrderedListIds(tree) {
  const ids = new Map();
  let counter = 0;

  const walk = (node) => {
    if (!node || typeof node !== "object") {
      return;
    }
    if (node.type === "list" && node.ordered) {
      ids.set(node, `ordered-list-${counter}`);
      counter += 1;
    }
    if (Array.isArray(node.children)) {
      for (const child of node.children) {
        walk(child);
      }
    }
  };

  walk(tree);
  return ids;
}

function inlineRuns(nodes, options = {}) {
  const runs = [];

  for (const node of nodes || []) {
    if (node.type === "text") {
      runs.push(new docx.TextRun({ text: node.value, ...options }));
    } else if (node.type === "strong") {
      runs.push(...inlineRuns(node.children, { ...options, bold: true }));
    } else if (node.type === "emphasis") {
      runs.push(...inlineRuns(node.children, { ...options, italics: true }));
    } else if (node.type === "delete") {
      runs.push(...inlineRuns(node.children, { ...options, strike: true }));
    } else if (node.type === "inlineCode") {
      runs.push(
        new docx.TextRun({
          text: node.value,
          font: CONSOLAS,
          ...options,
        }),
      );
    } else if (node.type === "break") {
      runs.push(new docx.TextRun({ text: "", break: 1, ...options }));
    } else if (node.type === "link") {
      const children = inlineRuns(node.children, options);
      if (node.url) {
        runs.push(
          new docx.ExternalHyperlink({
            link: node.url,
            children,
          }),
        );
      } else {
        runs.push(...children);
      }
    } else if (node.type === "image") {
      runs.push(
        new docx.TextRun({
          text: node.alt || "",
          italics: true,
          ...options,
        }),
      );
    }
    // Raw inline HTML is intentionally ignored for the minimal tool.
  }

  return runs;
}

function renderHeading(node) {
  const depth = node.depth;
  if (depth <= 1) {
    return new docx.Paragraph({
      style: "Title",
      children: inlineRuns(node.children, { color: "0077BA" }),
    });
  }
  if (depth === 2) {
    return new docx.Paragraph({
      style: "Heading1",
      children: inlineRuns(node.children),
    });
  }
  if (depth === 3) {
    return new docx.Paragraph({
      style: "Heading2",
      children: inlineRuns(node.children),
    });
  }
  return new docx.Paragraph({
    style: "BodyText",
    children: inlineRuns(node.children, { bold: true }),
  });
}

function renderParagraph(node, extra = {}) {
  return new docx.Paragraph({
    style: "BodyText",
    children: inlineRuns(node.children),
    ...extra,
  });
}

function renderCodeBlock(node) {
  const value = node.value.replace(/\n$/, "");
  const lines = value.split("\n");
  const runs = lines.map(
    (line, index) =>
      new docx.TextRun({
        text: line,
        ...(index > 0 ? { break: 1 } : {}),
      }),
  );

  return new docx.Paragraph({
    style: "CodeBlock",
    shading: { type: docx.ShadingType.CLEAR, fill: "F2F2F2" },
    children: runs,
  });
}

function renderBlockquote(node, ctx) {
  const blocks = [];
  for (const child of node.children || []) {
    if (child.type === "paragraph") {
      blocks.push(
        new docx.Paragraph({
          style: "BodyText",
          indent: { left: 385 },
          children: inlineRuns(child.children, {
            bold: true,
            color: "777777",
          }),
        }),
      );
    } else {
      blocks.push(...renderBlock(child, ctx));
    }
  }
  return blocks;
}

function tableBorders() {
  const side = {
    style: docx.BorderStyle.SINGLE,
    size: 6,
    color: "DDDDDD",
    space: 0,
  };
  return {
    top: side,
    bottom: side,
    left: side,
    right: side,
    insideHorizontal: side,
    insideVertical: side,
  };
}

function renderTable(node) {
  const rows = (node.children || []).map((row, rowIndex) => {
    const isHeader = rowIndex === 0;
    return new docx.TableRow({
      children: row.children.map((cell, cellIndex) => {
        const alignment = alignmentFor(node.align?.[cellIndex]);
        const paragraph = new docx.Paragraph({
          alignment,
          children: inlineRuns(cell.children || [], {
            size: 19,
            bold: isHeader,
          }),
        });
        return new docx.TableCell({
          shading: isHeader ? { fill: "EDEDED" } : undefined,
          children: [paragraph],
        });
      }),
    });
  });

  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
    borders: tableBorders(),
    layout: docx.TableLayoutType.FIXED,
    cellMargin: { top: 0, bottom: 0, left: 0, right: 0 },
    rows,
  });
}

function alignmentFor(value) {
  if (value === "center") {
    return docx.AlignmentType.CENTER;
  }
  if (value === "right") {
    return docx.AlignmentType.RIGHT;
  }
  return docx.AlignmentType.LEFT;
}

function renderList(node, ctx) {
  const blocks = [];
  for (const item of node.children || []) {
    blocks.push(...renderListItem(item, node, ctx));
  }
  return blocks;
}

function renderListItem(item, listNode, ctx) {
  const reference = listNode.ordered
    ? ctx.orderedIds.get(listNode)
    : "unordered-list";
  const level = Math.min(ctx.depth, LIST_LEVELS.length - 1);
  const blocks = [];

  for (const child of item.children || []) {
    if (child.type === "paragraph") {
      blocks.push(
        new docx.Paragraph({
          style: "ListParagraph",
          numbering: { reference, level },
          children: inlineRuns(child.children),
        }),
      );
    } else if (child.type === "list") {
      blocks.push(...renderList(child, { ...ctx, depth: ctx.depth + 1 }));
    } else {
      blocks.push(...renderBlock(child, ctx));
    }
  }

  return blocks;
}

function renderBlock(node, ctx) {
  if (node.type === "heading") {
    return [renderHeading(node)];
  }
  if (node.type === "paragraph") {
    return [renderParagraph(node)];
  }
  if (node.type === "list") {
    return renderList(node, ctx);
  }
  if (node.type === "table") {
    return [renderTable(node)];
  }
  if (node.type === "code") {
    return [renderCodeBlock(node)];
  }
  if (node.type === "blockquote") {
    return renderBlockquote(node, ctx);
  }
  if (node.type === "thematicBreak") {
    return [new docx.Paragraph("")];
  }
  return [];
}

export function buildDocument(markdown) {
  const tree = unified().use(remarkParse).use(remarkGfm).parse(markdown);
  const orderedIds = collectOrderedListIds(tree);
  const numbering = {
    config: [
      bulletNumbering(),
      ...[...orderedIds.values()].map((reference) => orderedNumbering(reference)),
    ],
  };

  const blocks = [];
  for (const child of tree.children || []) {
    blocks.push(...renderBlock(child, { orderedIds, depth: 0 }));
  }

  return new docx.Document({
    styles: STYLES,
    numbering,
    sections: [{ properties: { page: PAGE }, children: blocks }],
  });
}

export async function mdToBuffer(markdown) {
  return docx.Packer.toBuffer(buildDocument(markdown));
}

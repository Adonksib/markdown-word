import { unified } from "unified";
import remarkParse from "remark-parse";
import remarkGfm from "remark-gfm";
import * as docx from "docx";

const LIST_LEVELS = [0, 1, 2, 3, 4, 5];
const HEADING_BY_DEPTH = {
  1: docx.HeadingLevel.HEADING_1,
  2: docx.HeadingLevel.HEADING_2,
  3: docx.HeadingLevel.HEADING_3,
  4: docx.HeadingLevel.HEADING_4,
  5: docx.HeadingLevel.HEADING_5,
  6: docx.HeadingLevel.HEADING_6,
};

const STYLES = {
  paragraphStyles: [
    {
      id: "BlockQuote",
      name: "Block Quote",
      basedOn: "Normal",
      quickFormat: true,
      paragraph: {
        spacing: { before: 120, after: 120 },
        indent: { left: docx.convertInchesToTwip(0.4) },
      },
      run: {
        italics: true,
        color: "404040",
      },
    },
    {
      id: "CodeBlock",
      name: "Code Block",
      basedOn: "Normal",
      quickFormat: true,
      paragraph: {
        spacing: { before: 80, after: 80 },
      },
      run: {
        font: { name: "Consolas" },
        size: 18,
      },
    },
  ],
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
          font: { name: "Consolas" },
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
  const depth = Math.min(Math.max(node.depth, 1), 6);
  return new docx.Paragraph({
    heading: HEADING_BY_DEPTH[depth],
    children: inlineRuns(node.children),
  });
}

function renderParagraph(node, extra = {}) {
  return new docx.Paragraph({
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
      blocks.push(renderParagraph(child, { style: "BlockQuote" }));
    } else {
      blocks.push(...renderBlock(child, ctx));
    }
  }
  return blocks;
}

function renderTable(node) {
  const rows = (node.children || []).map((row, rowIndex) => {
    const isHeader = rowIndex === 0;
    return new docx.TableRow({
      children: row.children.map((cell, cellIndex) => {
        const alignment = alignmentFor(node.align?.[cellIndex]);
        const paragraphs = (cell.children || []).map((child) =>
          new docx.Paragraph({
            alignment,
            children: inlineRuns(child.children || [], {
              bold: isHeader,
            }),
          }),
        );
        return new docx.TableCell({
          shading: isHeader ? { fill: "D9E2F3" } : undefined,
          children: paragraphs.length > 0 ? paragraphs : [new docx.Paragraph("")],
        });
      }),
    });
  });

  return new docx.Table({
    width: { size: 100, type: docx.WidthType.PERCENTAGE },
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
        renderParagraph(child, {
          numbering: { reference, level },
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
    sections: [{ children: blocks }],
  });
}

export async function mdToBuffer(markdown) {
  return docx.Packer.toBuffer(buildDocument(markdown));
}

#!/usr/bin/env node

import { mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { mdToBuffer } from "./converter.js";

const MARKDOWN_EXTENSIONS = new Set([".md", ".markdown"]);

const USAGE = `Markdown 转 Word 命令行工具

用法:
  md2docx convert <输入文件.md> <输出文件.docx>   转换单个 Markdown 文件
  md2docx batch <输入目录> <输出目录>             批量转换目录下的 Markdown 文件

支持: 标题、列表、表格、代码块、引用、粗体、斜体、行内代码、链接。
仅支持 .md 和 .markdown 文件。
`;

class CliError extends Error {
  constructor(message, exitCode = 1) {
    super(message);
    this.exitCode = exitCode;
  }
}

function isMarkdownFile(filePath) {
  return MARKDOWN_EXTENSIONS.has(path.extname(filePath).toLowerCase());
}

async function assertInputFile(filePath) {
  let info;
  try {
    info = await stat(filePath);
  } catch (error) {
    if (error.code === "ENOENT") {
      throw new CliError(`文件不存在: ${filePath}`, 2);
    }
    throw error;
  }
  if (!info.isFile()) {
    throw new CliError(`文件不存在或不是普通文件: ${filePath}`, 2);
  }
}

async function assertInputDir(dirPath) {
  let info;
  try {
    info = await stat(dirPath);
  } catch (error) {
    if (error.code === "ENOENT") {
      throw new CliError(`输入目录不存在: ${dirPath}`, 2);
    }
    throw error;
  }
  if (!info.isDirectory()) {
    throw new CliError(`输入路径不是目录: ${dirPath}`, 2);
  }
}

async function ensureWritableDir(dirPath) {
  try {
    await mkdir(dirPath, { recursive: true });
  } catch (error) {
    throw new CliError(`输出目录不可写: ${dirPath} (${error.message})`, 4);
  }

  const probe = path.join(dirPath, `.write-probe-${process.pid}`);
  try {
    await writeFile(probe, "");
  } catch (error) {
    throw new CliError(`输出目录不可写: ${dirPath} (${error.message})`, 4);
  } finally {
    await rm(probe, { force: true }).catch(() => {});
  }
}

async function walkFiles(dirPath) {
  const entries = await readdir(dirPath, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      const nested = await walkFiles(fullPath);
      files.push(...nested.map((relative) => path.join(entry.name, relative)));
    } else if (entry.isFile()) {
      files.push(entry.name);
    }
  }

  return files;
}

function docxPathFor(relativePath) {
  const parsed = path.parse(relativePath);
  return path.join(parsed.dir, `${parsed.name}.docx`);
}

async function convertSingle(inputPath, outputPath) {
  await assertInputFile(inputPath);
  if (!isMarkdownFile(inputPath)) {
    throw new CliError(
      `格式不支持: ${path.extname(inputPath) || "无扩展名"}，仅支持 .md / .markdown`,
      3,
    );
  }

  const absoluteOutput = path.resolve(outputPath);
  await ensureWritableDir(path.dirname(absoluteOutput));

  const markdown = await readFile(inputPath, "utf8");
  const buffer = await mdToBuffer(markdown);
  await writeFile(absoluteOutput, buffer);

  console.log(`已转换: ${inputPath} -> ${absoluteOutput}`);
}

async function convertBatch(inputDir, outputDir) {
  await assertInputDir(inputDir);
  const absoluteOutputDir = path.resolve(outputDir);
  await ensureWritableDir(absoluteOutputDir);

  const files = await walkFiles(inputDir);
  const markdownFiles = files.filter(isMarkdownFile);
  const skippedCount = files.length - markdownFiles.length;
  const failed = [];
  let converted = 0;

  if (markdownFiles.length === 0) {
    console.log("未找到 Markdown 文件。");
    return;
  }

  for (const relativePath of markdownFiles) {
    const inputPath = path.join(inputDir, relativePath);
    const outputPath = path.join(absoluteOutputDir, docxPathFor(relativePath));
    try {
      await ensureWritableDir(path.dirname(outputPath));
      const markdown = await readFile(inputPath, "utf8");
      const buffer = await mdToBuffer(markdown);
      await writeFile(outputPath, buffer);
      converted += 1;
      console.log(`已转换: ${relativePath}`);
    } catch (error) {
      failed.push({ file: relativePath, message: error.message });
    }
  }

  console.log(
    `完成: 转换 ${converted} 个文件，跳过 ${skippedCount} 个非 Markdown 文件。`,
  );

  if (failed.length > 0) {
    for (const item of failed) {
      console.error(`转换失败: ${item.file} - ${item.message}`);
    }
    process.exitCode = 1;
  }
}

async function main(argv) {
  const [command, ...args] = argv;

  if (!command || command === "help" || command === "--help" || command === "-h") {
    process.stdout.write(USAGE);
    return;
  }

  if (command === "convert") {
    if (args.length !== 2) {
      throw new CliError("convert 需要两个参数: <输入文件.md> <输出文件.docx>");
    }
    await convertSingle(args[0], args[1]);
    return;
  }

  if (command === "batch") {
    if (args.length !== 2) {
      throw new CliError("batch 需要两个参数: <输入目录> <输出目录>");
    }
    await convertBatch(args[0], args[1]);
    return;
  }

  throw new CliError(`未知命令: ${command}\n\n${USAGE}`);
}

main(process.argv.slice(2)).catch((error) => {
  console.error(`错误: ${error.message}`);
  process.exit(error.exitCode ?? 1);
});

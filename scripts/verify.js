import { execFileSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function runCli(args) {
  execFileSync(process.execPath, [path.join(rootDir, "src", "cli.js"), ...args], {
    cwd: rootDir,
    stdio: "inherit",
  });
}

async function documentXml(docxPath) {
  const zip = await JSZip.loadAsync(await readFile(docxPath));
  return zip.file("word/document.xml").async("string");
}

function check(name, condition) {
  if (condition) {
    console.log(`PASS ${name}`);
  } else {
    console.error(`FAIL ${name}`);
    process.exitCode = 1;
  }
}

runCli([
  "convert",
  path.join("examples", "input", "demo.md"),
  path.join("examples", "output", "single", "demo.docx"),
]);
runCli([
  "batch",
  path.join("examples", "input"),
  path.join("examples", "output", "batch"),
]);

const xml = await documentXml(
  path.join(rootDir, "examples", "output", "single", "demo.docx"),
);

check("保留一级标题", xml.includes('w:val="Heading1"'));
check("保留无序/有序列表", xml.includes("w:numPr"));
check("保留表格", xml.includes("<w:tbl>"));
check("保留代码块样式", xml.includes('w:val="CodeBlock"'));
check("保留代码内容", xml.includes("const answer = 42;"));
check("保留引用样式", xml.includes('w:val="BlockQuote"'));

const nestedXml = await documentXml(
  path.join(rootDir, "examples", "output", "batch", "sub", "notes.docx"),
);
check("批量转换递归子目录", nestedXml.includes("子目录文件"));

if (process.exitCode) {
  console.error("验证失败");
} else {
  console.log("验证通过");
}

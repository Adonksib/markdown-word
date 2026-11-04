# Markdown 转 Word

一个最小可用的命令行工具，把 Markdown 文件批量转换为同名 `.docx` 文件。

## 功能

- 单文件转换
- 目录批量转换，递归处理子目录
- 保留标题、无序/有序列表、表格、代码块、引用
- 保留粗体、斜体、删除线、行内代码、链接
- 清晰的错误提示：文件不存在、格式不支持、输出目录不可写

## 技术栈

- [unified](https://unifiedjs.com) + [remark-parse](https://github.com/remarkjs/remark) + [remark-gfm](https://github.com/remarkjs/remark-gfm)：解析 Markdown
- [docx](https://docx.js.org)：生成 Word 文档
- Node.js >= 18

## 安装

```bash
npm install
```

## 使用

### 单文件

```bash
node src/cli.js convert <输入文件.md> <输出文件.docx>
```

示例：

```bash
node src/cli.js convert examples/input/demo.md examples/output/single/demo.docx
```

### 批量

```bash
node src/cli.js batch <输入目录> <输出目录>
```

示例：

```bash
node src/cli.js batch examples/input examples/output/batch
```

批量转换会递归扫描输入目录，保持子目录结构，并把每个
`xxx.md` 或 `xxx.markdown` 转换为同名 `xxx.docx`。非 Markdown 文件会被跳过。

## 支持的元素

| 元素 | 转换结果 |
| --- | --- |
| `#` 到 `######` 标题 | Word Heading 1-6 样式 |
| 无序 / 有序 / 嵌套列表 | Word 列表编号 |
| 管道表格 | Word 表格 |
| 围栏代码块 | 等宽字体、灰色底纹 |
| `>` 引用 | 缩进斜体引用样式 |
| 粗体、斜体、删除线、行内代码、链接 | 对应格式 |

## 错误提示

| 场景 | 退出码 |
| --- | --- |
| 输入文件不存在 | 2 |
| 文件格式不支持（仅 `.md` / `.markdown`） | 3 |
| 输出目录不可写 | 4 |

## 示例

示例文件位于 `examples/input/`：

- `demo.md`：覆盖标题、列表、表格、代码块、引用
- `sub/notes.md`：用于验证递归批量转换

## 验证

运行以下命令，会执行单文件和批量转换，并检查生成的文档中是否保留了
标题、列表、表格、代码块和引用：

```bash
npm run verify
```

## 项目结构

```text
.
├── src/
│   ├── cli.js          # 命令行入口和错误处理
│   └── converter.js    # Markdown AST 到 docx 的转换
├── scripts/
│   └── verify.js       # 验证命令
├── examples/
│   └── input/          # 示例 Markdown 文件
└── package.json
```

## 上传到 GitHub

在 GitHub 新建仓库后，执行：

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <你的仓库地址>
git push -u origin main
```

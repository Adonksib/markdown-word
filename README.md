# Markdown 转 Word（小白教程）

这个工具能把 Markdown（`.md`）文件自动变成 Word（`.docx`）文件。
标题、列表、表格、代码块、引用都会保留下来。

下面每一步都照着做就行。

## 第 0 步：先检查电脑有没有 Node.js

这个工具需要电脑里装了 Node.js。

1. 按键盘上的 `Win` 键，输入 `PowerShell`，回车打开。
2. 在弹出的窗口里输入下面这行，然后回车：

```bash
node -v
```

3. 如果屏幕上出现 `v18`、`v20`、`v22` 这类带数字的结果，就说明装好了，直接跳到第 1 步。
4. 如果提示“不是内部或外部命令”之类的话，去 <https://nodejs.org> 下载 LTS 版本，一路点“下一步”装完，再重新打开 PowerShell，重复第 2 步。

## 第 1 步：进入这个项目的文件夹

在 PowerShell 里输入下面这行，然后回车：

```bash
cd D:\markdown转word
```

`cd` 是“进入文件夹”的意思。

## 第 2 步：安装工具需要的零件（第一次做一次就行）

输入下面这行，回车，等它自己跑完：

```bash
npm install
```

这一步只需要做一次。以后再用这个工具，不用重复做。

## 第 3 步：先转换一个示例文件试试

输入下面这整行，然后回车：

```bash
node src/cli.js convert examples/input/demo.md examples/output/single/demo.docx
```

这句话的意思是：

- `convert`：转换一个文件
- `examples/input/demo.md`：要转换的 Markdown 文件
- `examples/output/single/demo.docx`：生成出来的 Word 文件放在哪里

看到“已转换”三个字，就是成功了。

去打开这个文件看看效果：

```text
D:\markdown转word\examples\output\single\demo.docx
```

## 第 4 步：批量转换一个文件夹里的所有 Markdown 文件

输入下面这行，回车：

```bash
node src/cli.js batch examples/input examples/output/batch
```

意思是：把 `examples/input` 文件夹里所有的 `.md` 文件，全部转成
`.docx`，放到 `examples/output/batch` 文件夹里。

子文件夹也会一起处理，原来的文件夹结构会保留。

结果在这里：

```text
D:\markdown转word\examples\output\batch
```

## 第 5 步：检查工具是否正常工作（可选）

输入下面这行，回车：

```bash
npm run verify
```

如果最后看到“验证通过”，说明一切正常。

## 怎么转换你自己的文件

1. 新建一个文件夹，比如：

```text
D:\我的笔记
```

2. 把你的 `.md` 文件放进去，也可以在里面再建子文件夹。
3. 在 PowerShell 里输入下面这行，回车：

```bash
node src/cli.js batch D:\我的笔记 D:\我的笔记-word
```

4. 转换完成后，打开 `D:\我的笔记-word` 文件夹，里面就是对应的 Word 文件。

## 两条命令怎么改

转换一个文件：

```text
node src/cli.js convert 输入文件.md 输出文件.docx
```

批量转换一个文件夹：

```text
node src/cli.js batch 输入文件夹 输出文件夹
```

把后面两个路径换成你自己的路径就行。路径里有中文或者空格也没关系，
照抄即可。

## 遇到报错怎么办

| 报错里出现的关键词 | 意思 | 解决办法 |
| --- | --- | --- |
| 文件不存在 | 输入的文件路径写错了，或者文件不在那个位置 | 检查路径和文件名是否一致 |
| 格式不支持 | 工具只认 `.md` 和 `.markdown` 结尾的文件 | 把文件后缀改成 `.md` |
| 输出目录不可写 | 输出的文件夹没有写入权限，或者被别的程序占用了 | 换一个输出文件夹试试 |

## 支持哪些内容

| Markdown 写法 | 转换后 |
| --- | --- |
| `#` 到 `######` 开头的标题 | Word 里的 1 到 6 级标题 |
| `-` 或数字开头的列表 | Word 里的项目符号或编号列表 |
| 用 `|` 写的表格 | Word 里的表格 |
| 三个反引号包起来的代码块 | 等宽字体的代码块 |
| `>` 开头的引用 | 缩进的引用样式 |
| 粗体、斜体、删除线、行内代码、链接 | 对应的格式 |

## 想上传到 GitHub

先在 GitHub 网站上新建一个空仓库，然后把下面每行依次输入 PowerShell，
每输一行回车一次：

```bash
git remote add origin 你的仓库地址
git push -u origin main
```

`你的仓库地址` 换成 GitHub 仓库页面上复制出来的地址，例如
`https://github.com/你的用户名/仓库名.git`。

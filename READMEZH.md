# GMath

![GMath 封面图](assets/readme/gmath.png)

GMath 是一个 Microsoft Word 加载项，用来编写并插入可编辑的数学公式。

你可以在任务面板中可视化编辑公式，也可以直接输入 LaTeX。插入到 Word 后，公式是 Word 原生公式，不是图片，因此后续仍然可以在文档中继续编辑。

English documentation: [README.md](README.md).

## 功能

- 在 Word 任务面板中可视化编辑公式
- 支持 LaTeX 输入，并与可视化编辑器同步
- 提供分数、根号、上标、下标、求和、积分、极限、矩阵等快捷按钮
- 支持行内公式和独立成行的显示公式
- 插入 Word 原生可编辑公式

## 使用要求

- Microsoft Word
- Node.js 和 npm
- 首次加载编辑器时需要联网，因为当前版本通过 CDN 加载 MathLive
- 本地旁加载时需要安装受信任的 localhost 开发证书

## 本地运行

在项目目录中运行：

```bash
npm install
npm run dev-certs
npm run serve
```

保持服务运行。加载项地址为：

```text
https://localhost:3000/src/taskpane.html
```

另开一个终端，在 macOS 上旁加载到 Word：

```bash
npm run sideload:mac
```

完全退出并重新打开 Word。随后应能在 **开始** 选项卡中看到 **GMath**。

## 使用方法

1. 打开 Word。
2. 点击 **开始** > **GMath** > **插入公式**。
3. 在面板中可视化输入公式，使用快捷按钮，或直接编辑 LaTeX。
4. 选择是否以独立成行的显示公式插入。
5. 点击 **插入到 Word**。

公式会作为 Word 原生公式插入文档，之后仍可继续编辑。

## 常见问题

如果 Word 中没有出现加载项，请重新运行 `npm run sideload:mac`，然后完全重启 Word。

如果任务面板空白，请确认 `npm run serve` 仍在运行，并确认已通过 `npm run dev-certs` 安装 localhost 证书。

如果 Word 仍显示旧图标或旧界面，请完全退出 Word 后重新打开。Office 可能会缓存加载项清单和图片。

如果公式插入后显示不正确，可以打开任务面板底部的调试区域，查看生成的 MathML 和 OMML。部分复杂 MathML 结构可能暂未支持。

## 第三方软件

GMath 使用 MathLive 提供可视化公式编辑能力。MathLive 采用 MIT License。

## 卸载加载项

在 macOS 上删除旁加载清单：

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

然后重新启动 Word。

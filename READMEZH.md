# GMath

![GMath 封面图](assets/readme/gmath.png)

[English](README.md)

`GMath` 是一个面向 macOS 的 Microsoft Word 公式加载项。它提供任务窗格式的可视化公式编辑、LaTeX 输入、符号速选、图片转公式识别，以及 Word 原生公式插入能力。

插入结果不是图片，而是 Word 原生可编辑公式。

本项目目前主要面向 Microsoft Word for Mac。Windows 旁加载流程还没有打包到此仓库中。

## 功能亮点

- 在 Word 任务窗格中可视化编辑公式。
- 输入 LaTeX，并与可视化编辑器保持同步。
- 通过符号速选插入分数、根号、上下标、括号、矩阵、求和、积分、极限、重音、向量和常用符号。
- 插入行内、行间和右编号 Word 公式。
- 通过 OpenAI 兼容视觉模型接口把公式截图识别为 LaTeX。
- 默认使用 GitHub Pages 托管任务窗格；对于禁止浏览器跨域请求的接口，可临时使用本地代理。

## 环境要求

- macOS
- Microsoft Word for Mac
- Node.js 和 npm

## 安装

在仓库根目录安装依赖：

```bash
npm install
```

将加载项旁加载到 Word for Mac：

```bash
npm run sideload:mac
```

用 `Cmd+Q` 完全退出 Word，然后重新打开。GMath 应出现在 **开始** 选项卡中。

默认旁加载清单会让 Word 打开：

```text
https://ge-shun.github.io/GMath/src/taskpane.html
```

更新仓库后，如果 Word 仍使用旧的清单或任务窗格，请重新运行 `npm run sideload:mac`。

## 使用

1. 打开 Word。
2. 点击 **开始** > **GMath** > **插入公式**。
3. 在任务窗格中使用可视化编辑、符号速选或 LaTeX 源构建公式。
4. 选择 **行内**、**行间** 或 **右编号**。
5. 点击 **插入到 Word**。

## 图片转公式

在任务窗格的 **接口设置** 中填写：

- OpenAI 兼容的 chat completions 地址，例如 `https://api.openai.com/v1/chat/completions`
- API Key
- 支持图片输入的模型名

接口地址和 Key 会保存在任务窗格的本机浏览器存储中。识别时，所选图片和 API Key 会发送给你配置的接口服务商。

对于禁止浏览器跨域请求的接口，启动临时代理：

```bash
npm run dev-certs
npm run serve
```

使用图片识别期间保持该终端运行。代理地址为：

```text
https://localhost:3000/api/ai/chat/completions
```

GMath 不会安装常驻后台服务。

## 开发调试

如果需要让 Word 加载本地任务窗格，而不是 GitHub Pages：

```bash
npm run dev-certs
npm run serve
```

另开一个终端：

```bash
npm run sideload:mac:local
```

完全重启 Word。本地任务窗格地址为：

```text
https://localhost:3000/src/taskpane.html
```

如需切回托管任务窗格：

```bash
npm run sideload:mac
```

## 常见问题

如果 Word 中没有出现 GMath，请重新运行 `npm run sideload:mac`，并用 `Cmd+Q` 完全重启 Word。

如果更新后仍显示旧界面，请完全重启 Word。Office 可能会缓存加载项清单和任务窗格资源。

如果图片识别失败，请确认 API 地址兼容 OpenAI，模型支持图片输入，并在需要 CORS 代理时保持 `npm run serve` 运行。

如果公式插入后显示不正确，可以打开任务窗格底部的调试区域，查看生成的 MathML / OMML。部分复杂结构可能暂未支持。

## 卸载加载项

删除旁加载清单：

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

然后重新启动 Word。

## 第三方软件

GMath 使用 MathLive 提供可视化公式编辑能力。MathLive 采用 MIT License。

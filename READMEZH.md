# GMath

![GMath 封面图](assets/readme/gmath.png)

GMath 是一个 Microsoft Word 加载项，用来编写并插入可编辑的数学公式。

你可以在任务面板里可视化编辑公式，也可以用符号速选，或直接输入 LaTeX。插入到 Word 后，公式是 Word 原生公式，不是图片，因此后续仍然可以在文档中继续编辑。

English documentation: [README.md](README.md).

## 当前平台

GMath 目前按 **macOS 上的 Microsoft Word** 配置。

仓库里的旁加载命令会把 manifest 复制到 Word for Mac 的加载项目录。Windows 旁加载流程目前还没有打包到这个项目里。

## 不开本地服务也能用什么

完成旁加载后，插件主界面默认从 GitHub Pages 加载：

```text
https://ge-shun.github.io/GMath/src/taskpane.html
```

因此，默认模式下打开任务窗格需要能访问 GitHub Pages。它不需要本地后台服务；任务窗格加载完成后，公式编辑和插入在任务窗格与 Word 内完成。

这些基础功能不需要一直运行本地服务：

- 可视化编辑公式
- LaTeX 输入和预览
- 符号速选
- 行内、行间、右编号公式插入
- 插入 Word 原生可编辑公式

只有“图片转公式”在接口不允许浏览器跨域请求时，才可能需要临时启动本地代理。见 [图片转公式](#图片转公式)。

## Mac 安装

需要：

- macOS
- Microsoft Word for Mac
- Node.js 和 npm

在项目目录中先安装依赖：

```bash
npm install
```

然后把加载项旁加载到 Word：

```bash
npm run sideload:mac
```

用 `Cmd+Q` 完全退出 Word，再重新打开。随后应能在 **开始** 选项卡中看到 **GMath**。

如果更新后 Word 仍显示旧界面，请重新运行 `npm run sideload:mac`，再完全退出并重新打开 Word。Office 会缓存加载项清单和任务窗格页面。

## 使用方法

1. 打开 Word。
2. 点击 **开始** > **GMath** > **插入公式**。
3. 在任务面板中可视化输入公式、点击符号速选，或直接编辑 LaTeX 源。
4. 选择 **行内**、**行间** 或 **右编号**。
5. 点击 **插入到 Word**。

插入后的内容是 Word 原生公式，之后仍可在 Word 中继续编辑。

## 图片转公式

图片转公式会把你选择的图片发送到你在 **接口设置** 中填写的 OpenAI 兼容视觉模型接口。

你需要准备：

- API 地址，例如 `https://api.openai.com/v1/chat/completions`，或服务商提供的 `/v1` 地址
- API Key
- 支持图片/视觉输入的模型

API 地址和 Key 只保存在任务窗格使用的本机浏览器存储中。识别时，图片和 Key 会发送给你自己配置的接口服务商。

### 可以直接识别时

如果接口服务商允许浏览器直接请求，不需要启动本地服务。填写 **接口设置** 后，粘贴或选择图片即可识别。

### 出现 `Load failed` 时

很多接口服务商会因为 CORS 禁止浏览器直接请求。此时临时启动本地代理。

第一次使用前先安装本地证书：

```bash
npm run dev-certs
```

然后启动代理：

```bash
npm run serve
```

使用图片识别期间保持这个终端打开。插件会自动回退到：

```text
https://localhost:3000/api/ai/chat/completions
```

识别完可以直接停止服务。GMath 不会安装常驻后台服务。

## 本地运行

普通使用不需要看这一节。

只有在调试本地页面，或想强制 Word 从 `localhost` 加载任务窗格时，才需要本地运行：

第一次使用前先安装本地证书：

```bash
npm run dev-certs
```

终端 1：

```bash
npm run serve
```

终端 2：

```bash
npm run sideload:mac:local
```

然后完全退出并重新打开 Word。此时 Word 会加载：

```text
https://localhost:3000/src/taskpane.html
```

使用本地页面时，需要保持 `npm run serve` 运行。

如果要切回 GitHub Pages 版本：

```bash
npm run sideload:mac
```

然后再次完全重启 Word。

## 常见问题

如果 Word 中没有出现 **GMath**，请运行 `npm run sideload:mac`，然后用 `Cmd+Q` 完全退出 Word 并重新打开。

如果使用默认 GitHub Pages 版本时任务面板空白，请检查网络连接并重启 Word。默认主界面不需要 `npm run serve`。

如果使用 `sideload:mac:local` 时任务面板空白，请确认 `npm run serve` 仍在运行，并确认已通过 `npm run dev-certs` 安装 localhost 证书。

如果图片识别提示 `Load failed`，请临时运行 `npm run serve` 后再试。同时确认 API 地址是 OpenAI 兼容接口，且模型支持图片输入。

如果公式插入后显示不正确，可以打开任务面板底部的调试区域，查看生成的 MathML / OMML。部分复杂结构可能暂未支持。

## 卸载加载项

在 macOS 上删除旁加载清单：

```bash
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml
```

然后重新启动 Word。

## 第三方软件

GMath 使用 MathLive 提供可视化公式编辑能力。MathLive 采用 MIT License。

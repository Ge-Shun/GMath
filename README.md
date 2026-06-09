# GMath — Word 公式加载项

在 Word 中可视化编辑并插入**原生可编辑公式**（OMML），用于替代 MathType。

链路：**MathLive 编辑 → MathML → 自研转换器转 OMML → Flat OPC 包 → Office.js `insertOoxml` 插入**。
插入的是 Word 原生公式，可双击继续编辑，不是图片。

> MathML→OMML 由自研模块 [`src/mathml2omml.js`](src/mathml2omml.js) 完成，**不依赖第三方转换库**，仅用浏览器原生 `DOMParser` + 递归翻译。

## 目录结构

```
manifest.xml        加载项清单（指向 https://localhost:3000）
src/
  taskpane.html     任务面板 UI（内嵌 MathLive 编辑器）
  taskpane.js       核心逻辑：取 MathML → 调转换器 → 包 OOXML → 插入
  mathml2omml.js    自研 MathML→OMML 转换器（无第三方依赖）
package.json        本地开发脚本
```

> 仅 MathLive 编辑器通过 CDN 加载；MathML→OMML 为自研代码。首次使用需联网取 MathLive。

## 运行（macOS）

```bash
cd /Users/geshun/GMath
npm install

# 1) 安装受信任的 localhost 开发证书（Office 加载项要求 HTTPS）
npm run dev-certs

# 2) 启动本地 HTTPS 静态服务（保持此终端开着）
npm run serve
#   访问 https://localhost:3000/taskpane.html 应能看到编辑器

# 3) 新开一个终端，旁加载 manifest 到 Word
npm run sideload:mac

# 4) 完全退出并重新打开 Word，在「开始」选项卡找到 GMath → 点「插入公式」
```

在面板里编辑公式 → 点「插入到 Word」即可。

## 卸载 / 重新加载

- 修改 `src/` 下文件后，在面板里右键 → 重新加载即可，无需重启 Word。
- 卸载：删除 `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`。

## 常见问题

- **插入失败 / 公式错乱**：打开面板底部「查看 OMML / MathML」折叠区，把 MathML 和 OMML 贴出来排查。多为 MathML 中个别标签 mml2omml 暂不支持。
- **面板空白**：多为 HTTPS 证书未被信任。重跑 `npm run dev-certs`，并在浏览器直接打开 `https://localhost:3000/taskpane.html` 确认不报证书错。
- **Windows**：把 `sideload:mac` 换成共享文件夹旁加载或用 `office-addin-debugging`；服务与逻辑通用。

## 后续可扩展

- 公式编号与交叉引用（对标 MathType 的核心增值功能）
- 常用公式收藏 / 模板库
- 读取选中的已有公式回填编辑（OMML→MathML→MathLive）
- 离线打包：把 MathLive 与 mml2omml 改为本地构建产物，去掉 CDN 依赖

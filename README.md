# Word 译文批注审阅工作台（纯前端本地 & GitHub Pages 版）

专为高密度双语译稿批注审阅设计的纯浏览器 `.docx` 工作台。**零后端、零外部 CDN 依赖**，既可直接双击 `index.html` 离线使用，也可一键部署到 GitHub Pages 供团队通过网址打开使用（**所有 `.docx` 解析与导出 100% 在用户浏览器本地内存完成，文稿内容绝不上传服务器**）。

## 核心功能

1. **强联动点击放大定位**：点击文中任意高亮字或角标，右侧对应批注卡自动居中、放大并高亮加框；点击右侧卡片，左侧对应的中文锚点、关联英文词与绿色参照词同步加框微放大。
2. **完整兼容 Word 多轮嵌套批注回复**：无损读取并写回 Word 原生多轮回复结构（`comments.xml` + `commentsExtended.xml` + `commentsIds.xml` + `commentsExtensible.xml`），支持查看各轮作者回复、新增回复、修改、标记已解决。
3. **删除批注自动清底色**：删除某条批注时可一键同步清除其绑定的中英文青／绿高亮底色（支持 `Ctrl+Z` / `Ctrl+Y` 35 步撤销重做）。
4. **双模式修改文字（顶部自由切换）**：
   - **修订模式（Track Changes）**：选中文字改词自动生成标准 `<w:del>` / `<w:ins>` 修订痕迹，并支持对文档已有修订点击「接受／拒绝」；
   - **直接改字模式**：无痕修改定稿文字。

## 部署到 GitHub Pages（仅需 1 分钟）

1. 在 GitHub 新建一个仓库（Public，若为 GitHub Pro/Team 也可 Private），将本文件夹内的 **5 个文件**（`index.html`、`styles.css`、`zip-codec.js`、`docx-engine.js`、`app.js` 及 `.nojekyll`）上传到仓库根目录。
2. 进入仓库 **Settings → Pages**，在 **Build and deployment** 下选择 `Deploy from a branch`，Branch 选 `main` / `/ (root)` 并点 **Save**。
3. 等待约 30 秒，访问生成的 `https://<你的用户名>.github.io/<仓库名>/` 即可直接使用！

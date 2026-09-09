# Markdown 工具箱 — Microsoft Word 加载项

在 Word 与 Markdown 之间双向转换，三个核心功能：

| 功能 | 入口 | 说明 |
| --- | --- | --- |
| **导出 Markdown** | 功能区「导出 Markdown」/ 面板「导出 MD」 | 把 Word 文档转成 `.md`，保留标题、列表、加粗、斜体、删除线、行内代码、表格、引用、超链接 |
| **粘贴 Markdown** | 功能区「粘贴 Markdown」/ 面板「粘贴 MD」 | 复制 Markdown 后一键转为 Word 富文本插入光标处 |
| **导入 Markdown** | 面板「导入 MD」 | 选择或拖入 `.md` 文件，生成排版完整的 Word 文档 |

---

## 快速开始

### 方式 A：双击 exe（推荐，零配置）

下载或构建 `dist\Markdown工具箱.exe`，**双击即可**。程序会自动完成：

1. 生成自签名 HTTPS 证书（纯 Python 实现，无需 openssl）
2. 注册加载项到 Word（写注册表 + 复制清单）
3. 启动本地 HTTPS 服务器
4. 在浏览器打开预览面板

然后在 Word 中：**插入 → 我的加载项 → 共享文件夹 → Markdown 工具箱**。

关闭命令行窗口即停止服务。

**命令行参数：**

```bash
Markdown工具箱.exe                 # 默认端口 3000
Markdown工具箱.exe --port 3001     # 换端口（3000 被占用时）
Markdown工具箱.exe --no-open       # 不自动打开浏览器
Markdown工具箱.exe --no-install    # 跳过 Word 注册
Markdown工具箱.exe --uninstall     # 从 Word 卸载加载项
Markdown工具箱.exe --http          # 仅 HTTP（浏览器预览用）
```

**自行构建 exe：**

```bash
build.bat                          # 一键构建（自动装依赖、生成图标、跑测试、打包）
# 或手动
pip install pyinstaller cryptography pillow
python assets\make-ico.py
pyinstaller --clean --noconfirm build_exe.spec
```

产物：`dist\Markdown工具箱.exe`（约 13 MB，单文件，无需安装 Python）。

### 方式 B：源码运行（开发用）

```bash
cd word-markdown-addin
npm start
```

首次运行会自动生成自签名 HTTPS 证书（需要 `openssl`，Git for Windows 自带）。
服务器地址：`https://localhost:3000`

> 仅想预览界面？`npm run preview` 启动 HTTP 模式，浏览器打开 `http://localhost:3000/taskpane.html`。

### 注册加载项到 Word

exe 会自动完成。若用源码运行，执行：

**PowerShell（无需管理员）：**

```powershell
.\install.ps1
```

脚本会复制清单到本地目录、写入注册表，并把证书加入当前用户受信任根证书。

**注册采用双通道，任一成功即可生效：**

| 通道 | 注册表位置 | 说明 |
| --- | --- | --- |
| **开发者直挂** | `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` 下以**加载项 GUID 为值名**、清单完整路径为数据 | 微软官方工具 `office-addin-dev-settings` 使用的机制，允许本地路径 |
| **共享文件夹** | `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{GUID}` 下 `Id` / `Url` / `Flags`=1 | Word「共享文件夹」列表读取这里；`Url` **必须是 UNC 网络路径**，脚本会复用系统已有的共享自动构造 |

> 早期版本只写了一个名为 `TrustedCatalogs` 的字符串值，Office 并不识别这种写法，因此「共享文件夹」显示为空。现已修正。

### 在 Word 中打开

1. **完全关闭 Word**（检查任务管理器无 `WINWORD.EXE`），重新打开一个文档
2. 「**开始**」选项卡最右侧会出现「Markdown 工具箱」按钮组
3. 若未出现：点击「**插入**」→「**我的加载项**」→「**共享文件夹**」→ 选择「Markdown 工具箱」→「添加」
4. 仍未出现：在「插入」→「我的加载项」→「**上传我的加载项**」中手动选择
   `%LOCALAPPDATA%\MarkdownToolbox\catalog\manifest.xml`

---

## 功能详解

### 功能 1：导出 Markdown

读取文档 OOXML 并离线解析，无需联网。

**支持的 Word 格式 → Markdown 映射：**

| Word | Markdown |
| --- | --- |
| 标题 1–6 级 | `#` … `######` |
| 加粗 / 斜体 / 删除线 | `**` / `*` / `~~` |
| 等宽字体（Consolas 等） | `` `行内代码` `` |
| 无序 / 有序 / 嵌套列表 | `-` / `1.` / 缩进 |
| 表格 | GFM 表格（含对齐方式） |
| 引用样式段落 | `>` |
| 超链接 | `[文字](url)` |
| 图片 | `![图片]` 占位 |
| 上标 / 下标 | `<sup>` / `<sub>` |
| 硬换行 | 行尾两空格 |

**使用方式：**
- 功能区按钮：一键导出，弹窗中下载 `.md` 或复制到剪贴板
- 侧边面板：可选「整篇文档 / 当前选区」，带实时预览与统计

### 功能 2：粘贴 Markdown → Word 富文本

在任意位置复制 Markdown 文本（网页、编辑器、聊天窗口），回到 Word：

- **功能区按钮**：直接读取剪贴板并插入光标处（剪贴板读取被拦截时会弹出手动粘贴窗口）
- **侧边面板**：左侧粘贴/编辑，右侧实时预览 Word 效果，可选插入位置（光标处 / 文档末尾 / 替换全文 / 新建文档）

### 功能 3：导入 .md 文件 → Word 文档

侧边面板「导入 MD」标签页：

1. 拖入或选择 `.md` 文件（支持 `.md` / `.markdown` / `.txt`，UTF-8）
2. 预览渲染结果与文件统计
3. 选择导入方式后点击「生成 Word 文档」

**支持的 Markdown 语法：**

- 标题：ATX（`#`）与 Setext（`===` / `---`）
- 强调：`**粗**`、`*斜*`、`***粗斜***`、`~~删除~~`、`==高亮==`
- 列表：无序、有序、嵌套、任务列表 `- [x]`
- 表格：GFM 表格，支持左/中/右对齐
- 引用块（含嵌套内容）
- 代码：围栏代码块（```）、缩进代码块、行内代码
- 链接、图片、自动链接、水平线
- 转义字符 `\*`

**排版风格预设：**

| 预设 | 正文 | 标题 |
| --- | --- | --- |
| 通用 | Calibri / 微软雅黑 11pt | Calibri Light 蓝色系 |
| 论文 | 宋体 12pt、首行缩进 24pt、1.5 倍行距 | 黑体，一级标题居中 |
| 跟随 Word 样式 | 不写内联样式 | 使用 Word 内置 Heading 1–6 |

---

## 项目结构

```
word-markdown-addin/
├── manifest.xml              # 加载项清单（GUID、功能区、权限）
├── launcher.py               # 单文件启动器（注册 + 证书 + 服务器 + 打开浏览器）
├── build_exe.spec            # PyInstaller 打包配置
├── build.bat                 # 一键构建 exe
├── server.js                 # 零依赖 HTTPS 开发服务器（Node 版）
├── install.ps1               # PowerShell 注册脚本
├── start.bat                 # 双击启动 Node 服务器
├── package.json
├── dist/
│   └── Markdown工具箱.exe     # 打包产物（约 13 MB 单文件）
├── src/
│   ├── taskpane.html/css/js  # 侧边任务面板（三个标签页）
│   ├── commands.html/js      # 功能区命令宿主
│   ├── dialog-export.html    # 导出结果对话框
│   ├── dialog-paste.html     # 手动粘贴对话框
│   ├── dialog-error.html     # 错误提示对话框
│   ├── support.html          # 帮助页（清单 SupportUrl 指向此页）
│   ├── selftest.html         # 浏览器内自检页
│   └── core/
│       ├── zip.js            # 零依赖 ZIP + inflate 解压器
│       ├── markdown.js       # Markdown → Word HTML 转换器
│       ├── docx2md.js        # .docx → Markdown 转换器
│       └── office-api.js     # Office.js 交互封装
├── assets/
│   ├── icon-16/32/80.png     # 功能区图标
│   ├── icon.ico              # exe 图标
│   ├── make-icons.py         # 图标生成脚本
│   └── make-ico.py           # ico 生成脚本
└── test/
    ├── make-fixture.js       # 生成测试用 .docx
    ├── test-markdown.js      # Markdown 转换测试（89 项）
    ├── test-docx.js          # docx 解析测试（33 项）
    └── test-roundtrip.js     # 端到端往返测试（48 项）
└── test-flatopc.js          # Flat OPC 解析测试（19 项）
```

**技术特点：零外部依赖。** 前端 ZIP 解压、inflate、XML 解析、Markdown 解析全部自行实现，`dependencies` 为空，完全离线可用。exe 启动器仅依赖 PyInstaller 与 cryptography（已打包进 exe）。

---

## 测试

```bash
npm test
```

四套测试共 **189 项**，覆盖：

- Markdown 语法解析（标题、强调、列表、表格、引用、代码块、链接、转义、边界情况）
- .docx 解析（ZIP、inflate、样式、编号、表格、行内格式）
- 端到端往返一致性（docx → md → html 信息不丢失）
- **Flat OPC 解析**（`getOoxml()` 真实返回格式，与 ZIP 路径产出逐字节一致）
- 安全过滤（script/iframe 清除、`javascript:` 协议剥离、内联事件移除）

浏览器内自检：启动服务器后访问 `https://localhost:3000/selftest.html`，24 项检查全部通过。

---

## 已知限制

- **图片导出**：Markdown 中仅保留 `![图片]` 占位符，不导出二进制图片（Markdown 本身不支持内嵌图片）
- **复杂表格**：Word 中合并单元格会展开为重复列，Markdown 表格无合并概念
- **公式（OMML）**：暂未转换，导出时忽略
- **剪贴板读取**：受浏览器/系统权限限制，失败时会自动回退到手动粘贴对话框
- **「新建文档」**：需要 WordApi 1.3+，旧版 Word 会自动回退为「插入到文档末尾」

---

## 常见问题

**Q：Word 里找不到加载项（「共享文件夹」显示为空）？**
A：按顺序排查：
1. 确认启动器窗口还开着（加载项指向 `https://localhost:3000`，服务必须运行）。
2. **完全关闭 Word**（任务管理器里确认没有 `WINWORD.EXE`）后重新打开——注册表变更只在 Word 启动时读取。
3. 查看「插入」→「我的加载项」→「共享文件夹」。若仍为空，用「上传我的加载项」手动选
   `%LOCALAPPDATA%\MarkdownToolbox\catalog\manifest.xml`。
4. 想确认注册是否写入，用注册表编辑器查看
   `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`，应有一个值名等于清单 `<Id>` 的字符串项。

**Q：为什么注册表里有两处注册？**
A：Office 有两条互不相通的加载项发现通道——「开发者直挂」和「共享文件夹目录」。脚本两处都写，任一被识别即可，互不冲突。

**Q：浏览器提示证书不安全？**
A：自签名证书的正常现象，点击「继续访问」。exe 首次运行会尝试把证书加入受信任根证书，之后不再提示。

**Q：exe 双击后闪退？**
A：从命令行运行看具体报错：`Markdown工具箱.exe`。常见原因是 3000 端口被占用，改用 `--port 3001`。

**Q：端口 3000 被占用怎么办？**
A：`Markdown工具箱.exe --port 3001`。注意端口变了之后，`manifest.xml` 里的 URL 也要相应修改（或保持默认 3000）。

**Q：导出时没有自动下载？**
A：浏览器可能拦截了下载。弹窗中提供了「下载 .md 文件」按钮，点击即可。

**Q：粘贴 Markdown 后格式不对？**
A：检查插入位置设置。若选择「光标处」，需先在文档中定位光标。

**Q：导出时报 `Failed to execute 'atob' … characters outside of the Latin1 range`？**
A：已在 v1.0.1 修复。原因是 `Body.getOoxml()` 返回的是 **Flat OPC 格式的 XML 字符串**（`<pkg:package>` 结构，内含中文），而不是 base64，早期版本多做了一次 `atob()` 解码。现在 `docx2md` 直接识别并解析 Flat OPC，二进制 `.docx` 与 Flat OPC 两条路径产出完全一致。若仍报错，说明用的是旧版 exe，请重新运行最新构建。

**Q：如何卸载？**
A：`Markdown工具箱.exe --uninstall`（或 `.\install.ps1 -Uninstall`），然后重启 Word。

**Q：exe 需要安装 Python 吗？**
A：不需要。PyInstaller 已把 Python 解释器和所有依赖打包进单个 exe 文件。

---

## 许可

MIT

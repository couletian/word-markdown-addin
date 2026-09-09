# 安全审查报告 · Markdown 工具箱（Word 加载项）

- **审查对象**：`word-markdown-addin` v1.0.1（`launcher.py`、`server.js`、`src/` 全部前端、`install.ps1`、`manifest.xml`、构建脚本）
- **审查日期**：2026-09-09
- **审查方式**：静态代码审计 + 动态实测（本地起服务探测 + Edge headless 真实浏览器执行验证）
- **结论**：**存在 2 个严重漏洞（其中 1 个无需用户交互即可执行）与 1 个高危信息泄露**，建议发布前修复。

---

## 修复状态（v1.0.2，2026-09-09 同日完成）

> **全部 9 项漏洞已修复并回归验证。** 详细修复说明见文末「附录 B：修复记录」。

| 编号 | 等级 | 修复方式 | 回归测试 |
|---|---|---|---|
| V-01 | 严重 | `sanitizeHtml` 重写为**标签/属性白名单**（词法扫描 + 重建） | ✅ 37 项 XSS 向量全部阻断 |
| V-02 | 严重 | `sanitizeUrl` 先**解码 HTML 实体**再去控制字符，最后做协议白名单 | ✅ 18 项实体绕过全部阻断 |
| V-03 | 高危 | 删除 `server.js` 的 `__dirname` 候选路径，改用 `ALLOWED_ROOTS` 白名单 | ✅ `key.pem` 等实测 404 |
| V-04 | 中危 | `safeDecode` 包裹 `decodeURIComponent` + `uncaughtException` 兜底 | ✅ `/%zz` 后服务存活 |
| V-05 | 中危 | `launcher.py` 的 `roots` 移除 `BUNDLE` | ✅ 源码模式实测 404 |
| V-06 | 中危 | CORS 收紧为同源回显 + 新增 `nosniff` | ✅ 实测响应头 |
| V-07 | 低危 | 内联脚本外置 + 严格 CSP（无 `unsafe-inline`）；SRI 经查证对 office.js 不可行 | ✅ 7 个页面 CSP 校验 |
| V-08 | 低危 | `inflateRaw` 增加 200 MB 输出上限 | ✅ 代码断言 |
| V-09 | 低危 | 新增 `.gitignore` 忽略私钥与构建产物 | ✅ 文件断言 |

**验证总量**：Node 测试 339 项（原 189 + 新增 150）全通过；浏览器自检 37 项全通过；真实浏览器攻击复测 `EXECUTED:[]`（零执行）；重建 exe 后内置代码哈希与源码一致且自检 37 项全通过。

---

## 一、风险总览

| 编号 | 漏洞 | 等级 | 触发条件 | 已验证 |
|---|---|---|---|---|
| V-01 | `sanitizeHtml` 可被绕过 → 自触发 XSS（`<svg/onload>`） | **严重** | 预览/插入含恶意 HTML 的 Markdown，**无需点击** | ✅ 浏览器实测执行 |
| V-02 | `sanitizeUrl` 未解码 HTML 实体 → `javascript:` 协议绕过 | **严重** | 用户点击转换出的链接 | ✅ 浏览器实测执行 |
| V-03 | 开发服务器泄露 TLS 私钥与整个项目目录 | **高危** | 访问 `https://localhost:3000/.certs/key.pem` | ✅ 实测 HTTP 200 |
| V-04 | `server.js` 畸形 URL 使进程崩溃 | 中危 | 请求 `/%zz` | ✅ 实测进程退出 |
| V-05 | `launcher.py` 源码模式根目录暴露私钥 | 中危 | 源码方式启动（非 exe） | ✅ 实测 HTTP 200 |
| V-06 | 本地服务无鉴权 + 证书入根 + CORS `*` | 中危 | 长期驻留的本地攻击面 | 代码审计 |
| V-07 | 第三方 CDN 脚本无 SRI | 低危 | office.js 供应链 | 代码审计 |
| V-08 | ZIP 解压无大小上限（解压炸弹） | 低危 | 导入恶意 .docx | 代码审计 |
| V-09 | 无 `.gitignore`，私钥易被误提交 | 低危 | 版本管理 | 代码审计 |

---

## 二、严重漏洞

### V-01 `sanitizeHtml` 正则过滤可绕过 → 自触发 XSS

**位置**：`src/core/markdown.js:100-121`（`sanitizeHtml`），触发点 `src/taskpane.js:139` 与 `:175` 的 `innerHTML` 赋值。

**成因**：`sanitizeHtml` 用正则黑名单过滤，存在两处结构性缺陷：

1. 危险标签正则 `<\s*(script|style|iframe|...)\b[^>]*>` —— **`svg` 不在黑名单内**；
2. 事件属性正则 `\son[a-z]+\s*=` 要求属性前必须是**空白字符**，而 HTML5 允许用 `/` 分隔属性（`<svg/onload=...>`）。

**实测输出**（Node 直接调用 `mdToHtml`）：

```
输入  <svg/onload=window.__p.push("C")>
输出  <svg/onload=window.__p.push("C")>     ← 原样透传
```

**浏览器执行验证**（Edge headless，`innerHTML` 注入后自动触发）：

```
EXECUTED: ["C","M"]     ← C 即 svg/onload，页面加载即执行，无需任何用户交互
```

**影响**：XSS 运行在加载项页面（`https://localhost:3000`）中，该页面持有 `ReadWriteDocument` 权限且可调用 `Word.run()`。攻击者可借此读取/改写用户文档、调用剪贴板接口。数据来源为「用户粘贴的 Markdown」或「导入的 .md 文件」，构成典型的**恶意文档攻击链**（如从网页复制一段 Markdown 后预览即中招）。

**同类未执行向量**（供修复时一并覆盖）：`<img/src=x/onerror=...>`、`<img src=x ONERROR=...>`、`<svg><set attributeName="onload" .../></svg>`（现代浏览器解析差异使其未触发，但不应依赖）。

---

### V-02 `sanitizeUrl` 未解码 HTML 实体 → `javascript:` 绕过

**位置**：`src/core/markdown.js:83-94`（`sanitizeUrl`）与 `:112-119`（`sanitizeHtml` 属性重写）。

**成因**：`sanitizeUrl` 只做「转小写 + 去控制字符」，**不解码 HTML 实体**。而 `sanitizeHtml` 在重写属性时把原值原样拼回引号内，实体被保留，浏览器解析时再解码：

```
输入  <a href="&#x09;javascript:alert(1)">click</a>
输出  <a href="&#x09;javascript:alert(1)">click</a>
```

**浏览器执行验证**：

```
raw=["\tjavascript:alert(1)"]  resolved=[javascript:alert(1)]   ← 实体被解码
window.__pwned = 1                                              ← 点击后 JS 执行
```

**已确认可用的变体**：

| 载荷 | 结果 |
|---|---|
| `&#x09;javascript:`（制表符） | ✅ 执行 |
| `&#x0A;javascript:`（换行） | ✅ 执行 |
| `&#106;avascript:`（十进制 j） | ✅ 执行 |
| `&#x6A;avascript:` | ✅ 执行 |
| `&#0000106;avascript:`（补零） | ✅ 执行 |
| `javascript&#x3A;`（编码冒号） | ✅ 执行 |

**影响**：恶意 .md 中的链接在预览或转换后仍可执行 JS，等同 V-01 的攻击面（需一次点击）。

> 说明：Markdown 原生链接语法 `[x](url)` 路径因 `escapeAttr` 二次转义而安全，**仅 HTML 透传路径**受影响。

---

## 三、高危漏洞

### V-03 开发服务器泄露 TLS 私钥与整个项目目录

**位置**：`server.js:100-123`（`resolveFile`）。

**成因**：`resolveFile` 的候选路径包含 `path.join(__dirname, ...)`，而 `__dirname` 是**项目根目录**。因此除 `src/`、`assets/` 之外，**根目录下所有文件都可被读取**。防目录穿越的检查 `p.startsWith(__dirname)` 恰好为真，等于没有限制。

**实测结果**（`node server.js --http`，`curl`）：

```
/.certs/key.pem        -> HTTP 200  (1708 bytes)   ← TLS 私钥明文
/.certs/cert.pem       -> HTTP 200  (1143 bytes)
/launcher.py           -> HTTP 200  (23331 bytes)
/install.ps1           -> HTTP 200  (8550 bytes)
/package.json          -> HTTP 200  (1179 bytes)
/test/fixture.docx     -> HTTP 200  (2667 bytes)
/../launcher.py        -> HTTP 200  (23331 bytes)
```

`key.pem` 内容确认为 `-----BEGIN PRIVATE KEY-----`。

**影响**：私钥可被本机任意进程读取，进而对 `localhost:3000` 实施中间人解密/伪造。叠加响应头 `Access-Control-Allow-Origin: *`，攻击面进一步扩大（见 V-06）。

**对比**：打包后的 exe 该问题已收敛（仅暴露 `src/`、`assets/`、`manifest.xml`，实测 `launcher.py`/`install.ps1` 均返回 404）。

---

## 四、中危漏洞

### V-04 畸形 URL 使服务器进程崩溃

**位置**：`server.js:101` —— `decodeURIComponent(pathname.split('?')[0])` 未包裹 `try/catch`。

**实测**：

```
请求 /%zz      -> 连接中断，进程退出
请求 /%E0%A4%A -> 连接中断
之后 /taskpane.html -> 502（服务已死）
```

Node 复现：`decodeURIComponent('/%zz')` → `URIError: URI malformed`，未捕获异常终止进程。

**影响**：任何能访问 `localhost:3000` 的页面或本地程序都可让服务停止，用户 Word 加载项随即失效（DoS）。`launcher.py` 的 `_resolve` 已有 `try/except`，不受影响。

### V-05 `launcher.py` 源码模式根目录暴露

**位置**：`launcher.py:297` —— `AddinHandler.roots = [SRC_DIR, ASSETS_DIR, BUNDLE]`。

`BUNDLE` 在非 frozen 模式下即项目根目录。实测源码启动后 `.certs/key.pem`、`launcher.py`、`install.ps1` 均返回 200。exe 模式下 `BUNDLE` 是 `sys._MEIPASS`（只含打包资源），故无此问题。

**建议**：源码模式也去掉 `BUNDLE`，或仅保留 `MANIFEST` 白名单。

### V-06 本地服务无鉴权 + 证书入根 + CORS `*`

- `ensure_certificate()` 调用 `certutil -user -addstore -f Root` 把自签证书加入**当前用户受信任根**（`launcher.py:196-201`）。证书一旦受信任，浏览器对 `https://localhost:3000` 的请求不再告警。
- 服务端对**所有**响应附加 `Access-Control-Allow-Origin: *`。
- 服务无任何鉴权。

**风险**：证书受信任 + CORS 通配 = 理论上任意网页可尝试跨源读取本地文件。当前 Chromium 的 Private Network Access（PNA）机制会拦截「公网 → 本地」的跨源请求（本服务未返回 `Access-Control-Allow-Private-Network`，preflight 也会因缺少 `do_OPTIONS` 而失败），但这属于**依赖浏览器单点防护**，不应作为设计依据。

**建议**：CORS 收紧为 `https://localhost:<port>`（或直接移除，加载项与服务器同源，不需要 CORS）；证书信任改为可选开关。

---

## 五、低危与加固建议

| 编号 | 问题 | 建议 |
|---|---|---|
| V-07 | `office.js` 从 `appsforoffice.microsoft.com` 加载，无 SRI | 加 `integrity` + `crossorigin`；或改为本地打包（离线场景更合适） |
| V-08 | `zip.js` 的 `inflateRaw` 无输出上限，可被解压炸弹耗尽内存 | 加输出大小上限（如 100 MB）并在超限时抛错 |
| V-09 | 无 `.gitignore`，`.certs/key.pem` 存在误提交风险 | 新增 `.gitignore`：`.certs/`、`build/`、`dist/`、`__pycache__/` |
| V-10 | 自签私钥以 `NoEncryption()` 明文落盘 | 至少限制文件权限；长期可考虑 DPAPI 加密 |
| V-11 | `manifest.xml` 权限为 `ReadWriteDocument` | 该权限为三项功能所必需，保留合理，但应在用户文档中明示 |

---

## 六、确认安全的部分

审计中已逐项验证以下设计是稳妥的：

- **对话框页面**（`dialog-export.html` / `dialog-paste.html` / `dialog-error.html`）：全部使用 `textContent` 赋值，无 `innerHTML` 注入点。
- **剪贴板预览**（`taskpane.js:114`）：`textContent`，安全。
- **XML 解析**（`docx2md.js`）：`MiniNode` 解析器显式跳过 `<!DOCTYPE` / `<!ENTITY`（`:174`），不展开外部实体，**无 XXE 风险**。
- **XSS 防护基础**：`escapeHtml` / `escapeAttr` 实现正确；Markdown 原生链接路径经二次转义后安全；`<script>` 标签、`on*` 事件（空格分隔形式）、`iframe`/`object`/`form action` 等经典向量均已被拦截（19/20 项测试通过）。
- **注册表操作**：仅写 `HKEY_CURRENT_USER`，不请求管理员权限，符合最小权限原则。
- **exe 打包**：`--no-install` 等参数齐全，资源暴露面已收敛。
- **无数据外传**：全项目除 `office.js` 外无任何 `fetch`/`XHR`/`sendBeacon`/外链请求，文档内容不会离开本机。

---

## 七、修复优先级建议

1. **P0** —— 重写 `sanitizeHtml`：放弃正则黑名单，改为**标签/属性白名单**方案（可用 `DOMParser` 解析后遍历节点，仅保留 `b/i/u/sup/sub/span/div/p/br/table/thead/tbody/tr/th/td/pre/code/a/img/hr` 等，属性仅保留 `href/src/title/alt/style` 且逐个校验）。
2. **P0** —— `sanitizeUrl` 增加**实体解码**（`&#xNN;` / `&#NNN;` / 命名实体）与「方案内空白」检测，解码后再做协议白名单判断（仅允许 `http/https/mailto/ftp/#` 与 `data:image/*;base64,`）。
3. **P1** —— `server.js`：删除 `path.join(__dirname, ...)` 兜底；`decodeURIComponent` 包 `try/catch`。
4. **P1** —— `launcher.py`：源码模式移除 `BUNDLE`；CORS 收紧。
5. **P2** —— 补 `.gitignore`、ZIP 输出上限、office.js SRI。

---

## 附录：复现命令

```bash
# V-01 / V-02：Node 侧观察过滤结果
node -e "var m=require('./src/core/markdown.js');
  console.log(m.mdToHtml('<svg/onload=alert(1)>'));
  console.log(m.mdToHtml('<a href=\"&#x09;javascript:alert(1)\">x</a>'));"

# V-03 / V-04：开发服务器暴露面与崩溃
node server.js --http --port 3520 &
curl -s -o /dev/null -w '%{http_code}\n' http://127.0.0.1:3520/.certs/key.pem   # 200
curl -s http://127.0.0.1:3520/%zz                                              # 进程退出
```

浏览器侧验证：将 `mdToHtml` 输出注入 `innerHTML`，用 Edge headless `--dump-dom` 读取执行标记（本报告结论均以此法复核）。

---

## 附录 B：修复记录（v1.0.2）

修复日期：2026-09-09 · 修复范围：全部 9 项（V-01 ~ V-09）

### B.1 V-01 / V-02 —— XSS 过滤体系重写

**核心决策：彻底放弃正则黑名单，改为「词法扫描 + 白名单重建」。**

原实现的问题在于「枚举危险模式」在原理上不可穷尽：`<svg/onload=...>` 这类斜杠分隔属性、实体编码的 `javascript:`、控制字符插入的 `java&#x09;script:` 都能绕过。

新实现（`src/core/markdown.js`）：

| 新增部件 | 作用 |
|---|---|
| `decodeEntities()` | 统一解码 `&#xNN;` / `&#NNN;` / 命名实体 / 缺失分号形式，并拒绝代理对；所有校验前先解码，杜绝「编码后绕过」 |
| `CTRL_RE` | 剔除 C0/C1 控制字符与零宽字符，阻断 `java\u0009script:` 类绕过 |
| `SAFE_TAGS` 白名单 | 仅 `p br div span b strong i em u s a img table thead tbody tr th td pre code hr ul ol li blockquote h1-h6` 等保留 |
| `DROP_WHOLE_TAGS` | `script style svg math iframe object form details …` **连内容整体丢弃**（而非只删标签） |
| `SAFE_ATTRS` 白名单 | 仅 `href src alt title style align colspan rowspan` 等；**无任何 `on*`**，且不含 `id`/`class` |
| `URL_ATTRS` | `href`/`src` 走 `sanitizeUrl` 独立校验 |
| `sanitizeStyle()` | 拦截 `expression(` `behavior:` `-moz-binding` `url(` `@import` |
| `sanitizeUrl()` | 先 `decodeEntities` → 去控制字符 → `trim`；`data:image/*;base64,` 单列放行；命中 `ANY_SCHEME_RE` 但不在 `SAFE_SCHEME_RE` 内 → 返回空串 |

`sanitizeHtml` 现在按字符流扫描标签，未在白名单内的标签整体丢弃；文本节点解码实体后再 `escapeHtml` 重新转义，避免「双重解码」引入新注入面。

**回归**：37 项 XSS 向量 + 18 项实体绕过向量全部阻断；`EXECUTED:[]`（零执行）。

### B.2 V-03 / V-05 —— 路径暴露面收敛

- `server.js`：删除 `resolveFile` 中所有 `path.join(__dirname, ...)` 候选路径（正是它能读到 `.certs/key.pem` 的原因），改为 `ALLOWED_ROOTS = [ROOT, ASSETS]` 白名单 + `path.resolve` 后 `path.relative` 越界检查。
- `launcher.py`：`AddinHandler.roots` 移除 `BUNDLE`（源码模式下它等于项目根目录，等同于把整个仓库暴露出去），保留 `[SRC_DIR, ASSETS_DIR]`；`_resolve` 的 `relative_to` 检查保留。

**回归**：`.certs/key.pem`、`launcher.py`、`package.json` 实测均 404。

### B.3 V-04 —— 畸形 URL 崩溃

新增 `safeDecode(pathname)`：`decodeURIComponent` 包 `try/catch`，失败返回 `null` 并响应 400。另加 `process.on('uncaughtException')` / `process.on('unhandledRejection')` 兜底，避免单次畸形请求导致服务静默退出。

**回归**：`/%zz` 后服务存活（旧版此处进程直接退出）。

### B.4 V-06 —— CORS 与响应头收紧

- CORS 由 `Access-Control-Allow-Origin: *` 改为回显 `https://localhost:<port>`，并加 `Vary: Origin`。
- 全站新增 `X-Content-Type-Options: nosniff`（`server.js` 与 `launcher.py` 双端一致）。

### B.5 V-07 —— CSP 与内联脚本外置

**关于 SRI**：经查证（Microsoft 官方文档与社区讨论），`office.js` **无法使用 SRI**——该脚本由 CDN 按 Office 版本动态返回不同内容，且 AppSource 上架强制要求走 CDN，哈希无法固定。因此改用**严格 CSP + 移除内联脚本**的组合，达成等效的「第三方脚本不可篡改扩展」效果。

- 7 个 HTML 页面统一注入 CSP：
  `default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; frame-src 'self' https://appsforoffice.microsoft.com; base-uri 'none'; form-action 'none'`
- 内联 `<script>` 全部外置为独立文件：`src/dialog-export.js`、`src/dialog-paste.js`、`src/dialog-error.js`、`src/selftest.js`，从而**不需要 `'unsafe-inline'`**（`style-src` 仍保留 `'unsafe-inline'`，因导出/预览依赖内联样式）。
- 已验证 CSP 生效后 `Office.onReady` 正常触发、`office.js` 正常加载。

### B.6 V-08 —— ZIP 解压炸弹

`src/core/zip.js` 新增 `MAX_INFLATE_BYTES = 200 * 1024 * 1024`，每块 `inflateRaw` 输出后校验累计大小，超限抛 `ZIP: 解压数据超过安全上限`。

### B.7 V-09 —— 私钥与产物保护

- 新增 `.gitignore`：忽略 `.certs/`、`*.pem`、`*.key`、`build/`、`dist/`、`__pycache__/`、`node_modules/` 与临时文件。
- `launcher.py` 首次生成证书时调用 `_restrict_key_acl()`（`icacls /inheritance:r /grant:r <user>:F`）并 `chmod 600`，限制私钥仅当前用户可读。

### B.8 回归验证矩阵

| 验证项 | 方法 | 结果 |
|---|---|---|
| Node 单元测试 | `npm test` | **339 项全通过**（原 189 + 新增 150） |
| 浏览器自检 | Edge headless `--dump-dom selftest.html` | **37 项全通过** |
| 真实攻击复测 | 注入 sanitize 输出 + 触发 `img.onerror`/`a.click()` | **`EXECUTED:[]` 零执行** |
| exe 内置代码一致性 | `sha256sum` 对比 exe 服务返回文件与源码 | **哈希完全一致** |
| exe 自检 | 重建 exe 后 Edge headless 跑 selftest | **37 项全通过** |
| 路径穿越 | 对 exe 请求 `.certs/key.pem` 等 | **404** |

### B.9 复测命令

```bash
# 全量测试（含安全用例）
npm test                      # 339 项

# 单独跑安全测试
npm run test:security         # 150 项

# 重建 exe 并验证内置代码
python -m PyInstaller --clean --noconfirm build_exe.spec
./dist/Markdown工具箱.exe --http --no-open --no-install --port 3138 &
curl -s http://localhost:3138/core/markdown.js | grep -c SAFE_TAGS   # >0 即含修复
```

### B.10 遗留说明

- **`office.js` 无法加 SRI**：非本项目可控，已用严格 CSP 缓解；如后续 Microsoft 支持 SRI，可再补。
- **`style-src` 保留 `'unsafe-inline'`**：导出/预览功能依赖内联样式注入（`src/core/docx2md.js` 的样式预设）。因 `style-src` 不构成脚本执行面，且 `sanitizeStyle()` 已过滤 CSS 表达式，风险可接受。
- **CSP 的 `script-src` 已去掉 `'unsafe-inline'`**：这是本轮最实质的纵深防御收益。

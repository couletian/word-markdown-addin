/*!
 * test-security.js — 安全回归测试
 * ------------------------------------------------------------------
 * 覆盖 SECURITY-AUDIT.md 中记录的全部漏洞（V-01 ~ V-09），
 * 确保修复不被后续改动悄悄回退。
 *
 * 运行：node test/test-security.js
 */
'use strict';

var fs = require('fs');
var path = require('path');

var md = require('../src/core/markdown.js');

var pass = 0;
var fail = 0;
var failures = [];

function check(name, cond, detail) {
  if (cond) {
    pass++;
    console.log('  \u2713 ' + name);
  } else {
    fail++;
    failures.push(name + (detail ? '  → ' + detail : ''));
    console.log('  \u2717 ' + name + (detail ? '  → ' + detail : ''));
  }
}

function section(title) {
  console.log('\n' + title);
  console.log('-'.repeat(60));
}

/* ------------------------------------------------------------------ *
 * V-01：sanitizeHtml 绕过 → 自触发 XSS
 * ------------------------------------------------------------------ */

section('V-01 sanitizeHtml 标签/事件白名单');

/** 判定清洗结果里是否残留可执行特征 */
function hasExecutable(html) {
  var flags = [];
  if (/<\s*(?:script|svg|math|iframe|object|embed|applet|frame|frameset|link|meta|base|form|video|audio|details|marquee|template|canvas|noscript)\b/i.test(html)) flags.push('危险标签');
  if (/\son[a-z]+\s*=/i.test(html)) flags.push('事件属性');
  if (/javascript\s*:|vbscript\s*:/i.test(html)) flags.push('危险协议');
  if (/\ssrcdoc\s*=/i.test(html)) flags.push('srcdoc');
  if (/expression\s*\(|behavior\s*:|-moz-binding|url\s*\(/i.test(html)) flags.push('危险 CSS');
  if (/xlink\s*:/i.test(html)) flags.push('xlink');
  if (/<(?:script|svg|iframe)/i.test(html)) flags.push('残留标签');
  return flags;
}

var XSS_VECTORS = [
  ['斜杠分隔 onerror', '<img/src=x/onerror=alert(1)>'],
  ['空格 onerror', '<img src=x onerror=alert(1)>'],
  ['大写 ONERROR', '<img src=x ONERROR=alert(1)>'],
  ['svg 斜杠 onload', '<svg/onload=alert(1)>'],
  ['svg 空格 onload', '<svg onload=alert(1)>'],
  ['svg set onload', '<svg><set attributeName="onload" to="alert(1)"/></svg>'],
  ['svg animate onbegin', '<svg><animate onbegin=alert(1) attributeName=x dur=1s/></svg>'],
  ['svg use href', '<svg><use href="data:image/svg+xml,<svg onload=alert(1)>"/></svg>'],
  ['math xlink:href', '<math><a xlink:href="javascript:alert(1)">x</a></math>'],
  ['script 标签', '<script>alert(1)</scr' + 'ipt>'],
  ['script 嵌套拼接', '<scr<script>ipt>alert(1)</scr</script>ipt>'],
  ['iframe srcdoc', '<iframe srcdoc="<script>alert(1)</script>"></iframe>'],
  ['object data', '<object data="javascript:alert(1)"></object>'],
  ['embed src', '<embed src="javascript:alert(1)">'],
  ['form action', '<form action="javascript:alert(1)"><input type=submit></form>'],
  ['details ontoggle', '<details open ontoggle=alert(1)>'],
  ['video src data:text', '<video src="data:text/html,<script>alert(1)</scr' + 'ipt>">'],
  ['audio src', '<audio src="javascript:alert(1)">'],
  ['body onload', '<body onload=alert(1)>'],
  ['style 表达式', '<style>body{background:url(javascript:alert(1))}</style>'],
  ['div style expression', '<div style="width:expression(alert(1))">x</div>'],
  ['div style url', '<div style="background:url(javascript:alert(1))">x</div>'],
  ['img dynsrc', '<img dynsrc="javascript:alert(1)">'],
  ['input autofocus onfocus', '<input autofocus onfocus=alert(1)>'],
  ['marquee onstart', '<marquee onstart=alert(1)>x</marquee>'],
  ['template 内嵌', '<template><script>alert(1)</scr' + 'ipt></template>'],
  ['base href 劫持', '<base href="https://evil.com/">'],
  ['meta refresh', '<meta http-equiv="refresh" content="0;url=javascript:alert(1)">'],
  ['link stylesheet', '<link rel="stylesheet" href="https://evil.com/x.css">'],
  ['注释内藏标签', '<!-- <script>alert(1)</scr' + 'ipt> --><p>ok</p>'],
  ['未闭合 script', '<script>alert(1)'],
  ['DOM clobbering id', '<img id="__proto__" src=x>'],
  ['class 注入', '<div class="evil">x</div>'],
  ['onload 无值', '<img src=x onload>'],
  ['onerror 反引号', '<img src=x onerror=`alert(1)`>'],
  ['tab 分隔属性', '<img\tsrc=x\tonerror=alert(1)>'],
  ['换行分隔属性', '<img\nsrc=x\nonerror=alert(1)>'],
];

XSS_VECTORS.forEach(function (v) {
  var out = md.mdToHtml(v[1], { stylePreset: 'default' });
  var flags = hasExecutable(out);
  check('阻断：' + v[0], flags.length === 0, flags.length ? '残留 ' + flags.join(',') + ' | ' + out.slice(0, 120) : '');
});

section('V-01 安全功能未被误伤');

var KEEP = [
  ['加粗保留', '<strong>粗体</strong>', '粗体'],
  ['斜体保留', '<em>斜体</em>', '斜体'],
  ['下划线保留', '<u>下划线</u>', '下划线'],
  ['段落保留', '<p>段落</p>', '段落'],
  ['换行保留', '<br>', '<br'],
  ['表格保留', '<table><tr><td>格</td></tr></table>', '<table>'],
  ['上标保留', '<sup>2</sup>', '<sup>'],
  ['下标保留', '<sub>2</sub>', '<sub>'],
  ['删除线保留', '<del>删</del>', '<del>'],
  ['高亮保留', '<mark>亮</mark>', '<mark>'],
  ['代码保留', '<code>x=1</code>', '<code>'],
  ['预格式保留', '<pre>code</pre>', '<pre>'],
  ['引用保留', '<blockquote>引</blockquote>', '<blockquote>'],
  ['列表保留', '<ul><li>项</li></ul>', '<li>'],
  ['中文保留', '<p>中文内容</p>', '中文内容'],
  ['实体文字保留', '<p>a &lt; b &amp; c</p>', '&lt;'],
  ['水平线保留', '<hr>', '<hr'],
  ['span 保留', '<span>文字</span>', '文字'],
  ['div 保留', '<div>块</div>', '块'],
  ['font 保留', '<font color="red">红</font>', '红'],
];

KEEP.forEach(function (k) {
  var out = md.mdToHtml(k[1], { stylePreset: 'default' });
  check('保留：' + k[0], out.indexOf(k[2]) >= 0, '输出=' + out.slice(0, 120));
});

section('V-01 安全 URL 与图片');

var SAFE_URLS = [
  ['https 链接', '<a href="https://example.com/a">x</a>', 'href="https://example.com/a"'],
  ['http 链接', '<a href="http://example.com/a">x</a>', 'href="http://example.com/a"'],
  ['相对路径', '<a href="/local/path">x</a>', 'href="/local/path"'],
  ['锚点', '<a href="#sec">x</a>', 'href="#sec"'],
  ['mailto', '<a href="mailto:a@b.com">x</a>', 'mailto:a@b.com'],
  ['https 图片', '<img src="https://example.com/a.png" alt="图">', 'src="https://example.com/a.png"'],
  ['data:image/png', '<img src="data:image/png;base64,iVBORw0KGgo=">', 'data:image/png;base64'],
];

SAFE_URLS.forEach(function (s) {
  var out = md.mdToHtml(s[1], { stylePreset: 'default' });
  check('放行：' + s[0], out.indexOf(s[2]) >= 0, '输出=' + out.slice(0, 140));
});

check('拒绝：data:image/svg+xml（可携带脚本）',
  md.mdToHtml('<img src="data:image/svg+xml;base64,PHN2Zz4=">').indexOf('svg') < 0);
check('拒绝：data:text/html',
  md.mdToHtml('<a href="data:text/html,<script>alert(1)</scr' + 'ipt>">x</a>').indexOf('data:text/html') < 0);
check('拒绝：file://',
  md.mdToHtml('<a href="file:///C:/Windows/win.ini">x</a>').indexOf('file:') < 0);
check('拒绝：blob:',
  md.mdToHtml('<a href="blob:https://x/y">x</a>').indexOf('blob:') < 0);
check('拒绝：about:',
  md.mdToHtml('<a href="about:blank">x</a>').indexOf('about:') < 0);

/* ------------------------------------------------------------------ *
 * V-02：sanitizeUrl 实体解码绕过
 * ------------------------------------------------------------------ */

section('V-02 实体编码绕过 javascript:');

var ENTITY_VECTORS = [
  ['十六进制制表符', '<a href="&#x09;javascript:alert(1)">x</a>'],
  ['十六进制换行', '<a href="&#x0A;javascript:alert(1)">x</a>'],
  ['十六进制回车', '<a href="&#x0D;javascript:alert(1)">x</a>'],
  ['十进制 j', '<a href="&#106;avascript:alert(1)">x</a>'],
  ['十六进制 j', '<a href="&#x6A;avascript:alert(1)">x</a>'],
  ['补零十进制', '<a href="&#0000106;avascript:alert(1)">x</a>'],
  ['方案内制表符', '<a href="java&#x09;script:alert(1)">x</a>'],
  ['方案内换行', '<a href="java&#x0A;script:alert(1)">x</a>'],
  ['编码冒号', '<a href="javascript&#x3A;alert(1)">x</a>'],
  ['无分号实体', '<a href="&#x09javascript:alert(1)">x</a>'],
  ['命名实体 tab', '<a href="&Tab;javascript:alert(1)">x</a>'],
  ['混合大小写', '<a href="JaVaScRiPt:alert(1)">x</a>'],
  ['前导空白', '<a href="   javascript:alert(1)">x</a>'],
  ['空字节', '<a href="\u0000javascript:alert(1)">x</a>'],
  ['U+2028 分隔', '<a href="java\u2028script:alert(1)">x</a>'],
  ['img src 实体', '<img src="&#x09;javascript:alert(1)">'],
  ['img src 编码冒号', '<img src="javascript&#x3A;alert(1)">'],
  ['vbscript 实体', '<a href="&#x09;vbscript:alert(1)">x</a>'],
];

ENTITY_VECTORS.forEach(function (v) {
  var out = md.mdToHtml(v[1], { stylePreset: 'default' });
  var bad = /javascript|vbscript/i.test(out.replace(/&#x?[0-9a-f]+;?/gi, ''));
  check('阻断：' + v[0], !bad, '输出=' + out.slice(0, 140));
});

section('V-02 Markdown 原生链接语法同样安全');

[
  ['原生链接 javascript:', '[x](javascript:alert(1))'],
  ['原生链接实体', '[x](&#x09;javascript:alert(1))'],
  ['原生图片 javascript:', '![x](javascript:alert(1))'],
  ['原生链接 data:text', '[x](data:text/html,<script>alert(1)</scr' + 'ipt>)'],
  ['原生链接 vbscript', '[x](vbscript:alert(1))'],
].forEach(function (v) {
  var out = md.mdToHtml(v[1], { stylePreset: 'default' });
  check('阻断：' + v[0],
    out.indexOf('javascript:') < 0 && out.indexOf('vbscript:') < 0 && out.indexOf('data:text') < 0,
    '输出=' + out.slice(0, 140));
});

check('原生安全链接保留', md.mdToHtml('[x](https://example.com)').indexOf('href="https://example.com"') >= 0);

/* ------------------------------------------------------------------ *
 * V-03 / V-05：服务器不得暴露敏感文件
 * ------------------------------------------------------------------ */

section('V-03 / V-05 服务器暴露面');

var serverSrc = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
var launcherSrc = fs.readFileSync(path.join(__dirname, '..', 'launcher.py'), 'utf8');

check('server.js 不再把 __dirname 作为用户路径的基准',
  !/path\.(?:join|resolve)\(__dirname,\s*(?:clean|rel|pathname|relPath)/.test(serverSrc),
  '仍存在以 __dirname 拼接用户输入路径的写法');

check('server.js 候选路径只用白名单根',
  !/candidates\.push\(/.test(serverSrc),
  '仍存在 candidates.push 的多路径回退');

check('server.js 使用白名单根目录',
  /ALLOWED_ROOTS\s*=\s*\[\s*ROOT\s*,\s*ASSETS\s*\]/.test(serverSrc));

check('server.js 路径校验使用 path.relative',
  /path\.relative\(/.test(serverSrc));

check('launcher.py roots 不含 BUNDLE',
  !/roots\s*=\s*\[[^\]]*BUNDLE/.test(launcherSrc),
  'roots 仍包含 BUNDLE');

check('launcher.py 声明只暴露 src/ 与 assets/',
  /AddinHandler\.roots\s*=\s*\[SRC_DIR,\s*ASSETS_DIR\]/.test(launcherSrc));

/* ------------------------------------------------------------------ *
 * V-04：畸形 URL 不得使进程崩溃
 * ------------------------------------------------------------------ */

section('V-04 畸形 URL 健壮性');

check('server.js 用 safeDecode 包裹 decodeURIComponent',
  /function safeDecode/.test(serverSrc) && /safeDecode\(/.test(serverSrc));

check('server.js 注册 uncaughtException 兜底',
  /uncaughtException/.test(serverSrc));

check('launcher.py 的 _resolve 有异常兜底',
  /except\s*\(ValueError,\s*OSError\)/.test(launcherSrc) || /except ValueError/.test(launcherSrc));

/* ------------------------------------------------------------------ *
 * V-06：CORS 收紧
 * ------------------------------------------------------------------ */

section('V-06 CORS 不再通配');

check('server.js 不再使用 Access-Control-Allow-Origin: *',
  !/Access-Control-Allow-Origin['"]\s*:\s*['"]\*['"]/.test(serverSrc),
  '仍存在通配 CORS');

check('launcher.py 不再使用通配 CORS',
  !/Access-Control-Allow-Origin['"]\s*,\s*['"]\*['"]/.test(launcherSrc));

check('两端都设置 nosniff',
  /X-Content-Type-Options/.test(serverSrc) && /X-Content-Type-Options/.test(launcherSrc));

/* ------------------------------------------------------------------ *
 * V-07：CSP 与脚本外置
 * ------------------------------------------------------------------ */

section('V-07 CSP 与第三方脚本');

var HTML_FILES = [
  'taskpane.html', 'commands.html', 'dialog-export.html',
  'dialog-paste.html', 'dialog-error.html', 'support.html', 'selftest.html'
];

HTML_FILES.forEach(function (name) {
  var html = fs.readFileSync(path.join(__dirname, '..', 'src', name), 'utf8');
  check('CSP 存在：' + name, /Content-Security-Policy/.test(html));
  check('CSP 无 script unsafe-inline：' + name,
    !/script-src[^;]*'unsafe-inline'/.test(html),
    'script-src 仍含 unsafe-inline');
  check('无内联脚本块：' + name,
    !/<script(?![^>]*\bsrc=)[^>]*>[\s\S]*?<\/script>/i.test(html),
    '存在内联 <script> 块');
  check('无内联事件处理器：' + name,
    !/<[^>]+\son(?:click|load|error|change|input|submit|focus|blur)\s*=/i.test(html),
    '存在内联事件处理器');
});

check('CSP 含 object-src none',
  HTML_FILES.every(function (n) {
    return /object-src 'none'/.test(fs.readFileSync(path.join(__dirname, '..', 'src', n), 'utf8'));
  }));
check('CSP 含 base-uri none',
  HTML_FILES.every(function (n) {
    return /base-uri 'none'/.test(fs.readFileSync(path.join(__dirname, '..', 'src', n), 'utf8'));
  }));

/* ------------------------------------------------------------------ *
 * V-08：解压炸弹防护
 * ------------------------------------------------------------------ */

section('V-08 ZIP 解压上限');

var zipSrc = fs.readFileSync(path.join(__dirname, '..', 'src', 'core', 'zip.js'), 'utf8');
check('zip.js 定义 MAX_INFLATE_BYTES', /MAX_INFLATE_BYTES/.test(zipSrc));
check('inflateRaw 含输出上限校验', /解压数据超过安全上限/.test(zipSrc));

/* ------------------------------------------------------------------ *
 * V-09：.gitignore 保护私钥
 * ------------------------------------------------------------------ */

section('V-09 版本管理防护');

var gitignorePath = path.join(__dirname, '..', '.gitignore');
check('.gitignore 存在', fs.existsSync(gitignorePath));
if (fs.existsSync(gitignorePath)) {
  var gi = fs.readFileSync(gitignorePath, 'utf8');
  check('.gitignore 忽略 .certs/', /^\.certs\/?$/m.test(gi));
  check('.gitignore 忽略 *.pem', /\*\.pem/.test(gi));
  check('.gitignore 忽略 build/', /^build\/?$/m.test(gi));
  check('.gitignore 忽略 __pycache__/', /__pycache__/.test(gi));
}

/* ------------------------------------------------------------------ *
 * 回归：确认关键修复代码仍在
 * ------------------------------------------------------------------ */

section('修复代码存在性回归');

var mdSrc = fs.readFileSync(path.join(__dirname, '..', 'src', 'core', 'markdown.js'), 'utf8');
check('markdown.js 有 decodeEntities', /function decodeEntities/.test(mdSrc));
check('markdown.js 有 SAFE_TAGS 白名单', /var SAFE_TAGS\s*=/.test(mdSrc));
check('markdown.js 有 DROP_WHOLE_TAGS', /var DROP_WHOLE_TAGS\s*=/.test(mdSrc));
check('markdown.js 有 SAFE_ATTRS 白名单', /var SAFE_ATTRS\s*=/.test(mdSrc));
check('markdown.js 已移除旧的正则黑名单',
  !/移除危险标签及其内容/.test(mdSrc),
  '仍存在旧的 sanitizeHtml 正则实现');
check('sanitizeUrl 先解码实体',
  /decodeEntities\(url\)/.test(mdSrc));
check('svg 在整段丢弃名单中', /svg:\s*1/.test(mdSrc));
check('math 在整段丢弃名单中', /math:\s*1/.test(mdSrc));

/* ------------------------------------------------------------------ *
 * 汇总
 * ------------------------------------------------------------------ */

console.log('\n' + '='.repeat(60));
console.log('安全回归测试汇总');
console.log('='.repeat(60));
console.log('通过 ' + pass + ' 项，失败 ' + fail + ' 项');
if (fail) {
  console.log('\n失败项：');
  failures.forEach(function (f) { console.log('  - ' + f); });
  process.exit(1);
}
console.log('全部通过 \u2713');

/*!
 * test-markdown.js — markdown.js 转换器单元测试（Node，零依赖）
 * 运行：node test/test-markdown.js
 */
'use strict';

var path = require('path');
var WordMd = { markdown: require(path.join(__dirname, '..', 'src', 'core', 'markdown.js')) };
var md = WordMd.markdown;

var pass = 0, fail = 0;
var failures = [];

function check(name, actual, expectContains) {
  var ok = Array.isArray(expectContains)
    ? expectContains.every(function (s) { return actual.indexOf(s) >= 0; })
    : actual.indexOf(expectContains) >= 0;
  if (ok) {
    pass++;
    console.log('  ✓ ' + name);
  } else {
    fail++;
    failures.push({ name: name, actual: actual, expect: expectContains });
    console.log('  ✗ ' + name);
    console.log('    期望包含: ' + JSON.stringify(expectContains));
    console.log('    实际输出: ' + JSON.stringify(actual.slice(0, 400)));
  }
}

function checkNot(name, actual, expectAbsent) {
  var ok = actual.indexOf(expectAbsent) < 0;
  if (ok) { pass++; console.log('  ✓ ' + name); }
  else {
    fail++;
    failures.push({ name: name, actual: actual, expect: '不应包含 ' + expectAbsent });
    console.log('  ✗ ' + name + ' — 不应包含 ' + JSON.stringify(expectAbsent));
  }
}

function section(title) { console.log('\n' + title); }

/* ============================ 标题 ============================ */
section('标题');
check('ATX 一级标题', md.mdToHtml('# 你好'), ['<h1', '你好', '</h1>']);
check('ATX 六级标题', md.mdToHtml('###### 深标题'), ['<h6', '深标题']);
check('标题内行内格式', md.mdToHtml('## 带 **粗体** 的标题'), ['<h2', '<strong>粗体</strong>']);
check('Setext 标题（===）', md.mdToHtml('大标题\n==='), ['<h1', '大标题']);
check('Setext 标题（---）', md.mdToHtml('小标题\n---'), ['<h2', '小标题']);
check('闭合式 ATX', md.mdToHtml('## 标题 ##'), ['<h2', '标题']);

/* ============================ 段落与换行 ============================ */
section('段落与换行');
check('单段落', md.mdToHtml('普通文字'), ['<p', '普通文字']);
check('两段落分隔', md.mdToHtml('第一段\n\n第二段'), ['第一段', '第二段']);
var twoP = md.mdToHtml('第一段\n\n第二段');
check('两段落生成两个 p', twoP, ['</p>', '<p']);
check('中文软换行不插空格', md.mdToHtml('中文\n换行'), '中文换行');
check('英文软换行插空格', md.mdToHtml('hello\nworld'), 'hello world');
check('硬换行（两空格）', md.mdToHtml('第一行  \n第二行'), '<br />');

/* ============================ 强调 ============================ */
section('强调');
check('加粗', md.mdToHtml('这是 **粗体** 文字'), '<strong>粗体</strong>');
check('下划线加粗', md.mdToHtml('这是 __粗体__ 文字'), '<strong>粗体</strong>');
check('斜体', md.mdToHtml('这是 *斜体* 文字'), '<em>斜体</em>');
check('下划线斜体', md.mdToHtml('这是 _斜体_ 文字'), '<em>斜体</em>');
check('粗斜体', md.mdToHtml('***两者***'), ['<strong><em>', '两者']);
check('删除线', md.mdToHtml('~~删除~~'), '<del>删除</del>');
check('高亮', md.mdToHtml('==高亮=='), '<mark>高亮</mark>');
check('加粗内中文', md.mdToHtml('**中文加粗**'), '<strong>中文加粗</strong>');
checkNot('不应把下划线变量名当斜体', md.mdToHtml('foo_bar_baz'), '<em>');

/* ============================ 行内代码 ============================ */
section('行内代码');
check('行内代码', md.mdToHtml('使用 `code` 示例'), ['<code', 'code</code>']);
check('行内代码含星号', md.mdToHtml('`**not bold**`'), ['<code', '**not bold**']);
check('行内代码不转义标记', mdToPlain(md.mdToHtml('`a_b_c`')), 'a_b_c');

function mdToPlain(html) { return html.replace(/<[^>]+>/g, ''); }

/* ============================ 转义 ============================ */
section('转义');
check('转义星号', md.mdToHtml('\\*不是斜体\\*'), '*不是斜体*');
checkNot('转义后不生成 em', md.mdToHtml('\\*不是斜体\\*'), '<em>');
check('转义反引号', md.mdToHtml('\\`not code\\`'), '`not code`');
check('HTML 实体转义', md.mdToHtml('a < b & c > d'), ['&lt;', '&amp;', '&gt;']);

/* ============================ 列表 ============================ */
section('列表');
check('无序列表', md.mdToHtml('- 项目一\n- 项目二'), ['<ul>', '<li>', '项目一', '项目二', '</ul>']);
check('有序列表', md.mdToHtml('1. 第一\n2. 第二'), ['<ol>', '第一', '第二', '</ol>']);
check('星号列表', md.mdToHtml('* 项目'), ['<ul>', '项目']);
check('加号列表', md.mdToHtml('+ 项目'), ['<ul>', '项目']);
check('列表项内加粗', md.mdToHtml('- **粗** 项'), ['<li>', '<strong>粗</strong>']);
check('嵌套列表', md.mdToHtml('- 外层\n  - 内层'), ['<ul>', '外层', '内层']);
var nested = md.mdToHtml('- 外层\n  - 内层');
check('嵌套列表 ul 数量为 2', String((nested.match(/<ul>/g) || []).length), '2');
check('任务列表未完成', md.mdToHtml('- [ ] 待办'), '\u2610');
check('任务列表已完成', md.mdToHtml('- [x] 完成'), '\u2611');
check('列表项多行合并', md.mdToHtml('- 第一行\n  第二行'), ['<li>', '第一行']);
check('有序列表嵌套', md.mdToHtml('1. 一\n   1. 子项'), ['<ol>', '子项']);
check('列表后接段落', md.mdToHtml('- 项\n\n段落'), ['<ul>', '<p', '段落']);
checkNot('列表标记不误判斜体', md.mdToHtml('*斜体*'), '<ul>');

/* ============================ 表格 ============================ */
section('表格');
var table = md.mdToHtml('| A | B |\n| --- | --- |\n| 1 | 2 |');
check('表格结构', table, ['<table', '<thead>', '<tbody>', '</table>']);
check('表头单元格', table, ['<th', 'A', 'B']);
check('表格数据单元格', table, ['<td', '1', '2']);
check('表格居中对齐', md.mdToHtml('| A |\n| :-: |\n| 1 |'), 'text-align:center');
check('表格右对齐', md.mdToHtml('| A |\n| ---: |\n| 1 |'), 'text-align:right');
check('表格左对齐', md.mdToHtml('| A |\n| :--- |\n| 1 |'), 'text-align:left');
check('表格内加粗', md.mdToHtml('| **A** |\n| --- |\n| 1 |'), '<strong>A</strong>');
check('表格内管道转义', md.mdToHtml('| A |\n| --- |\n| a \\| b |'), 'a | b');
check('无表头分隔的管道不算表格', md.mdToHtml('a | b'), '<p');

/* ============================ 引用块 ============================ */
section('引用块');
check('引用块', md.mdToHtml('> 引用内容'), ['引用内容']);
check('引用块含格式', md.mdToHtml('> **粗** 引用'), '<strong>粗</strong>');
check('多行引用', md.mdToHtml('> 第一行\n> 第二行'), ['第一行', '第二行']);

/* ============================ 代码块 ============================ */
section('代码块');
var fenced = md.mdToHtml('```js\nvar a = 1;\n```');
check('围栏代码块', fenced, ['<pre', 'var a = 1;']);
check('代码块内标记不解析', md.mdToHtml('```\n**not bold**\n```'), ['**not bold**']);
checkNot('代码块内不生成 strong', md.mdToHtml('```\n**x**\n```'), '<strong>');
check('缩进代码块', md.mdToHtml('    code line'), ['<pre', 'code line']);
check('代码块 HTML 转义', md.mdToHtml('```\n<div>\n```'), '&lt;div&gt;');
check('围栏语言标记被忽略', md.mdToHtml('```python\nx=1\n```'), 'x=1');

/* ============================ 链接与图片 ============================ */
section('链接与图片');
check('行内链接', md.mdToHtml('[文字](https://a.com)'), ['<a href="https://a.com"', '文字</a>']);
check('带标题的链接', md.mdToHtml('[文字](https://a.com "标题")'), ['title="标题"', '文字']);
check('自动链接', md.mdToHtml('<https://a.com>'), ['<a href="https://a.com"', 'https://a.com</a>']);
check('邮箱自动链接', md.mdToHtml('<a@b.com>'), ['mailto:a@b.com']);
check('图片', md.mdToHtml('![替代](img.png)'), ['<img src="img.png"', 'alt="替代"']);
check('链接内加粗', md.mdToHtml('[**粗**](https://a.com)'), ['<a href', '<strong>粗</strong>']);

/* ============================ 水平线 ============================ */
section('水平线');
check('三横线', md.mdToHtml('---'), ['style=', 'border-top']);
check('三星号', md.mdToHtml('***'), 'border-top');
checkNot('水平线不生成 hr 标签（default 预设）', md.mdToHtml('---'), '<hr');
check('水平线不误判为列表', md.mdToHtml('---'), ['<p style', '</p>']);

/* ============================ 综合 ============================ */
section('综合场景');
var full = md.mdToHtml([
  '# 文档标题',
  '',
  '这是一段 **加粗** 和 *斜体* 以及 `代码` 的说明。',
  '',
  '## 二级标题',
  '',
  '- 列表一',
  '- 列表二',
  '  - 嵌套项',
  '',
  '| 名称 | 值 |',
  '| --- | ---: |',
  '| 甲 | 1 |',
  '| 乙 | 2 |',
  '',
  '> 引用块内容',
  '',
  '```js',
  'console.log(1);',
  '```',
  '',
  '1. 有序一',
  '2. 有序二'
].join('\n'));

check('综合：标题', full, ['<h1', '<h2']);
check('综合：强调', full, ['<strong>', '<em>', '<code']);
check('综合：列表', full, ['<ul>', '<ol>']);
check('综合：表格', full, ['<table', '<th', '<td']);
check('综合：引用', full, '引用块内容');
check('综合：代码块', full, 'console.log(1);');

/* ============================ 预设 ============================ */
section('样式预设');
check('论文预设正文含宋体', md.mdToHtml('段落', { stylePreset: 'thesis' }), 'SimSun');
check('论文预设标题用黑体', md.mdToHtml('# 标题', { stylePreset: 'thesis' }), 'SimHei');
check('语义预设无内联样式', md.mdToHtml('# 标题', { stylePreset: 'semantic' }), '<h1>');
checkNot('语义预设不含 style', md.mdToHtml('段落', { stylePreset: 'semantic' }), 'style=');
check('默认预设含雅黑', md.mdToHtml('段落'), 'Microsoft YaHei');

/* ============================ 边界情况 ============================ */
section('边界情况');
check('空字符串', md.mdToHtml(''), '');
check('仅空白', md.mdToHtml('   \n\n  '), '');
check('null 输入', md.mdToHtml(null), '');
check('未闭合代码块', md.mdToHtml('```\ncode'), ['<pre', 'code']);
check('未闭合加粗', md.mdToHtml('**未闭合'), '**未闭合');
check('超长单行不崩溃', md.mdToHtml('a'.repeat(50000)).length > 50000 ? 'ok' : 'fail', 'ok');
check('CRLF 归一化', md.mdToHtml('第一段\r\n\r\n第二段'), ['第一段', '第二段']);
check('连续空行', md.mdToHtml('a\n\n\n\n\nb'), ['a', 'b']);
check('表格单列', md.mdToHtml('| A |\n| --- |\n| 1 |'), ['<th', '<td']);

/* ============================ 汇总 ============================ */
console.log('\n' + '='.repeat(60));
console.log('通过 ' + pass + ' 项，失败 ' + fail + ' 项');
if (fail) {
  console.log('\n失败详情：');
  failures.forEach(function (f, i) {
    console.log((i + 1) + '. ' + f.name);
    console.log('   期望: ' + JSON.stringify(f.expect));
    console.log('   实际: ' + JSON.stringify(String(f.actual).slice(0, 300)));
  });
  process.exit(1);
} else {
  console.log('全部通过 ✓');
}

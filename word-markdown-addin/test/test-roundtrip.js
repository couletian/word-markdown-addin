/*!
 * test-roundtrip.js — 端到端往返一致性测试
 * 运行：node test/make-fixture.js && node test/test-roundtrip.js
 *
 * 链路：docx --(功能1 导出)--> markdown --(功能2/3 导入)--> html
 * 验证三大功能串起来后信息不丢失。
 */
'use strict';

var fs = require('fs');
var path = require('path');

global.WordMd = { zip: require(path.join(__dirname, '..', 'src', 'core', 'zip.js')) };
var docx2md = require(path.join(__dirname, '..', 'src', 'core', 'docx2md.js'));
var markdown = require(path.join(__dirname, '..', 'src', 'core', 'markdown.js'));

var pass = 0, fail = 0;
var failures = [];

function check(name, actual, expectContains) {
  var list = Array.isArray(expectContains) ? expectContains : [expectContains];
  var ok = list.every(function (s) { return actual.indexOf(s) >= 0; });
  if (ok) { pass++; console.log('  ✓ ' + name); }
  else {
    fail++;
    failures.push({ name: name, expect: expectContains, actual: actual });
    console.log('  ✗ ' + name + ' 期望: ' + JSON.stringify(expectContains));
  }
}

function section(t) { console.log('\n' + t); }

var fixturePath = path.join(__dirname, 'fixture.docx');
if (!fs.existsSync(fixturePath)) {
  console.error('缺少测试文件，请先运行：node test/make-fixture.js');
  process.exit(1);
}

var buffer = fs.readFileSync(fixturePath);
var bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);

/* ---------- 阶段一：docx → markdown（功能 1） ---------- */
section('阶段一：Word → Markdown（功能 1）');
var md = docx2md.docxToMarkdown(bytes);
check('产出 Markdown 非空', md.length > 100 ? 'ok' : 'fail', 'ok');
check('含一级标题', md, '# 测试文档标题');
check('含表格', md, '| 名称 | 数量 | 备注 |');

/* ---------- 阶段二：markdown → html（功能 2/3） ---------- */
section('阶段二：Markdown → Word HTML（功能 2 / 3）');
var html = markdown.mdToHtml(md, { stylePreset: 'default' });
check('标题转换正确', html, ['<h1', '测试文档标题', '<h2', '<h3']);
check('加粗保留', html, '<strong>加粗文字</strong>');
check('斜体保留', html, '<em>斜体文字</em>');
check('行内代码保留', html, ['<code', '行内代码']);
check('删除线保留', html, '<del>删除线效果</del>');
check('无序列表保留', html, ['<ul>', '无序列表项一', '无序列表项二']);
check('嵌套列表保留', html, '嵌套列表项');
check('有序列表保留', html, ['<ol>', '有序列表项一', '有序列表项二']);
check('表格保留', html, ['<table', '<thead', '<th', '<td', '甲项目', '乙项目']);
check('表格三列', String((html.match(/<th[ >]/g) || []).length), '3');
check('表格两行数据', String((html.match(/<tr>/g) || []).length), '3');
check('引用保留', html, '引用样式的段落');
check('无残留 Markdown 标记', html.indexOf('**') < 0 && html.indexOf('##') < 0 ? 'clean' : 'dirty', 'clean');

/* ---------- 阶段三：语义完整性 ---------- */
section('阶段三：语义完整性核对');
// 提取纯文本对比关键信息
var plain = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
['测试文档标题', '普通正文', '加粗文字', '斜体文字', '行内代码', '删除线效果',
 '二级标题占位', '三级标题占位', '无序列表项一', '有序列表项二', '嵌套列表项',
 '有序列表项一', '有序列表项二', '名称', '数量', '备注', '甲项目', '乙项目',
 '引用样式的段落'].forEach(function (token) {
  check('文本保留：' + token, plain, token);
});

/* ---------- 阶段四：幂等性 ---------- */
section('阶段四：幂等性（Markdown 二次转换不劣化）');
var html2 = markdown.mdToHtml(md, { stylePreset: 'default' });
check('两次转换结果一致', html === html2 ? 'same' : 'diff', 'same');

var mdAgain = markdown.mdToHtml(markdown.mdToHtml('', {}), {});
check('空输入稳定', mdAgain, '');

/* ---------- 阶段五：真实场景 Markdown ---------- */
section('阶段五：复杂 Markdown 场景');
var complex = [
  '# 项目文档',
  '',
  '## 概述',
  '',
  '这是一个包含 **多种格式** 的文档，支持 `代码`、[链接](https://example.com) 和 *强调*。',
  '',
  '### 功能列表',
  '',
  '1. 第一项功能',
  '2. 第二项功能',
  '   - 子项 A',
  '   - 子项 B',
  '3. 第三项功能',
  '',
  '| 参数 | 类型 | 说明 |',
  '| :--- | :---: | ---: |',
  '| `name` | string | 名称 |',
  '| `count` | int | 数量 |',
  '',
  '> 注意：这是一条重要提示。',
  '',
  '```python',
  'def hello():',
  '    print("Hello")',
  '```',
  '',
  '- [x] 已完成任务',
  '- [ ] 待办任务',
  '',
  '---',
  '',
  '结尾段落。'
].join('\n');

var chtml = markdown.mdToHtml(complex, { stylePreset: 'default' });
check('复杂：三级标题', chtml, '<h3');
check('复杂：有序列表', chtml, '<ol>');
check('复杂：嵌套无序列表', String((chtml.match(/<ul>/g) || []).length), '2');
check('复杂：链接', chtml, ['href="https://example.com"', '链接</a>']);
check('复杂：表格对齐', chtml, ['text-align:left', 'text-align:center', 'text-align:right']);
check('复杂：引用', chtml, '重要提示');
check('复杂：代码块', chtml, ['<pre', 'def hello()']);
check('复杂：任务列表勾选', chtml, '\u2611');
check('复杂：任务列表未勾选', chtml, '\u2610');
check('复杂：水平线', chtml, 'border-top');
check('复杂：结尾段落', chtml, '结尾段落');

section('汇总');
console.log('='.repeat(60));
console.log('通过 ' + pass + ' 项，失败 ' + fail + ' 项');
if (fail) {
  failures.slice(0, 5).forEach(function (f, i) {
    console.log((i + 1) + '. ' + f.name + ' 期望: ' + JSON.stringify(f.expect));
  });
  process.exit(1);
} else {
  console.log('全部通过 ✓');
}

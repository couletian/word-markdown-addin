/*!
 * test-docx.js — docx2md 往返测试
 * 运行：node test/make-fixture.js && node test/test-docx.js
 */
'use strict';

var fs = require('fs');
var path = require('path');

var zipMod = require(path.join(__dirname, '..', 'src', 'core', 'zip.js'));
global.WordMd = { zip: zipMod };
var docx2md = require(path.join(__dirname, '..', 'src', 'core', 'docx2md.js'));

var pass = 0, fail = 0;
var failures = [];

function check(name, actual, expectContains) {
  var list = Array.isArray(expectContains) ? expectContains : [expectContains];
  var ok = list.every(function (s) { return actual.indexOf(s) >= 0; });
  if (ok) { pass++; console.log('  ✓ ' + name); }
  else {
    fail++;
    failures.push({ name: name, actual: actual, expect: expectContains });
    console.log('  ✗ ' + name);
    console.log('    期望包含: ' + JSON.stringify(expectContains));
  }
}

function checkNot(name, actual, expectAbsent) {
  var ok = actual.indexOf(expectAbsent) < 0;
  if (ok) { pass++; console.log('  ✓ ' + name); }
  else { fail++; failures.push({ name: name, actual: actual, expect: '不应包含 ' + expectAbsent }); console.log('  ✗ ' + name + ' 不应包含 ' + JSON.stringify(expectAbsent)); }
}

function section(t) { console.log('\n' + t); }

var fixturePath = path.join(__dirname, 'fixture.docx');
if (!fs.existsSync(fixturePath)) {
  console.error('缺少测试文件，请先运行：node test/make-fixture.js');
  process.exit(1);
}

var buffer = fs.readFileSync(fixturePath);
var bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);

section('ZIP 解析');
var files = zipMod.unzip(bytes);
check('解出 document.xml', Object.keys(files).join(','), 'word/document.xml');
check('解出 styles.xml', Object.keys(files).join(','), 'word/styles.xml');
check('解出 numbering.xml', Object.keys(files).join(','), 'word/numbering.xml');
var docText = zipMod.text(files['word/document.xml']);
check('document.xml 内容可读', docText, '测试文档标题');
check('CRC32 计算', String(zipMod.crc32(Buffer.from('hello'))), '907060870');

section('docx → Markdown 转换');
var md = docx2md.docxToMarkdown(bytes);
console.log('\n--- 转换结果 ---');
console.log(md);
console.log('--- 结束 ---\n');

check('一级标题', md, '# 测试文档标题');
check('二级标题', md, '## 二级标题占位');
check('三级标题', md, '### 三级标题占位');
check('普通段落', md, '这是一段普通正文，包含中英文混排内容。');
check('加粗', md, '**加粗文字**');
check('斜体', md, '*斜体文字*');
check('行内代码', md, '`行内代码`');
check('删除线', md, '~~删除线效果~~');
check('无序列表', md, '- 无序列表项一');
check('无序列表第二项', md, '- 无序列表项二');
check('嵌套列表缩进', md, '    - 嵌套列表项');
check('有序列表', md, '1. 有序列表项一');
check('有序列表递增', md, '2. 有序列表项二');
check('表格表头', md, '| 名称 | 数量 | 备注 |');
check('表格分隔行', md, '| --- | --- | --- |');
check('表格数据行', md, '| 甲项目 | 10 | 正常 |');
check('表格第二数据行', md, '| 乙项目 | 20 | 待定 |');
check('引用段落', md, '> 引用样式的段落');

section('边界与健壮性');
checkNot('空标题被跳过', md, '# \n');
checkNot('不产生连续三个空行', md, '\n\n\n');
check('结尾换行', md.slice(-1), '\n');

section('行内格式单元测试');
var I = docx2md._internal;
check('applyEmphasis 加粗', I.applyEmphasis('文字', { b: true }), '**文字**');
check('applyEmphasis 斜体', I.applyEmphasis('文字', { i: true }), '*文字*');
check('applyEmphasis 粗斜体', I.applyEmphasis('文字', { b: true, i: true }), '***文字***');
check('applyEmphasis 删除线', I.applyEmphasis('文字', { strike: true }), '~~文字~~');
check('applyEmphasis 代码', I.applyEmphasis('a_b', { code: true }), '`a_b`');
check('applyEmphasis 转义星号', I.applyEmphasis('a*b', { b: true }), '**a\\*b**');
check('applyEmphasis 空格外移', I.applyEmphasis(' 文字 ', { b: true }), ' **文字** ');

section('汇总');
console.log('='.repeat(60));
console.log('通过 ' + pass + ' 项，失败 ' + fail + ' 项');
if (fail) {
  failures.forEach(function (f, i) {
    console.log((i + 1) + '. ' + f.name + ' 期望: ' + JSON.stringify(f.expect));
  });
  process.exit(1);
} else {
  console.log('全部通过 ✓');
}

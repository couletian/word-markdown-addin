/*!
 * test-flatopc.js — Flat OPC（getOoxml 返回值）解析测试
 * 运行：node test/make-fixture.js && node test/test-flatopc.js
 *
 * 背景：Office.js 的 Body/Range.getOoxml() 返回的不是 base64，
 * 而是 Flat OPC 格式的 XML（<pkg:package> 包裹各 part）。
 * 早期版本误当作 base64 用 atob() 解码，导致
 * "Failed to execute 'atob' … characters outside of the Latin1 range"。
 * 本测试确保 Flat OPC 路径与 ZIP 路径产出完全一致。
 */
'use strict';

var fs = require('fs');
var path = require('path');
var zlib = require('zlib');

var zipMod = require(path.join(__dirname, '..', 'src', 'core', 'zip.js'));
global.WordMd = { zip: zipMod };
var docx2md = require(path.join(__dirname, '..', 'src', 'core', 'docx2md.js'));

var pass = 0, fail = 0;
var failures = [];

function ok(name) { pass++; console.log('  ✓ ' + name); }
function bad(name, detail) {
  fail++;
  failures.push(name);
  console.log('  ✗ ' + name + (detail ? '\n    ' + detail : ''));
}
function check(name, cond, detail) { cond ? ok(name) : bad(name, detail); }
function contains(name, actual, needle) {
  var hit = String(actual).indexOf(needle) >= 0;
  check(name, hit, hit ? '' : '未找到 ' + JSON.stringify(needle));
}
function section(t) { console.log('\n' + t); }

/* ------------------------------------------------------------------ *
 * 把 .docx 二进制转成 Flat OPC XML（模拟 Word getOoxml 的输出）
 * ------------------------------------------------------------------ */

function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function escapeAttr(s) {
  return escapeXml(s).replace(/"/g, '&quot;');
}

/** 需要走 xmlData 的 part（文本类）；其余走 binaryData（base64） */
function isXmlPart(name) {
  return /\.(xml|rels)$/i.test(name) || /\.rels$/i.test(name);
}

function contentTypeFor(name) {
  if (name === '[Content_Types].xml') return 'application/vnd.openxmlformats-officedocument.package+xml';
  if (/\.rels$/i.test(name)) return 'application/vnd.openxmlformats-package.relationships+xml';
  if (/styles\.xml$/i.test(name)) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
  if (/numbering\.xml$/i.test(name)) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
  if (/document\.xml$/i.test(name)) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  if (/\.png$/i.test(name)) return 'image/png';
  return 'application/octet-stream';
}

/** 把 ZIP 里的各 part 组装成 Flat OPC XML 文本 */
function docxToFlatOpc(bytes) {
  var files = zipMod.unzip(bytes);
  var parts = [];

  Object.keys(files).forEach(function (name) {
    var data = files[name];
    var ct = contentTypeFor(name);

    if (isXmlPart(name)) {
      var text = zipMod.text(data);
      // 去掉 XML 声明（Flat OPC 中 part 内嵌，不需要独立声明）
      text = text.replace(/^\uFEFF/, '').replace(/^<\?xml[^>]*\?>\s*/, '');
      parts.push(
        '  <pkg:part pkg:name="/' + name + '" pkg:contentType="' + ct + '" pkg:compression="store">\n' +
        '    <pkg:xmlData>' + text + '</pkg:xmlData>\n' +
        '  </pkg:part>'
      );
    } else {
      parts.push(
        '  <pkg:part pkg:name="/' + name + '" pkg:contentType="' + ct + '" pkg:compression="store">\n' +
        '    <pkg:binaryData>' + Buffer.from(data).toString('base64') + '</pkg:binaryData>\n' +
        '  </pkg:part>'
      );
    }
  });

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">\n' +
    parts.join('\n') + '\n' +
    '</pkg:package>';
}

/* ------------------------------------------------------------------ *
 * 准备数据
 * ------------------------------------------------------------------ */

var fixturePath = path.join(__dirname, 'fixture.docx');
if (!fs.existsSync(fixturePath)) {
  console.error('缺少测试文件，请先运行：node test/make-fixture.js');
  process.exit(1);
}

var buffer = fs.readFileSync(fixturePath);
var bytes = new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);

var mdFromZip = docx2md.docxToMarkdown(bytes);
var flatOpc = docxToFlatOpc(bytes);
var mdFromFlat = docx2md.docxToMarkdown(flatOpc);

section('Flat OPC 识别');
check('docx 二进制不被误判为 Flat OPC',
  docx2md.docxToMarkdown(bytes) === mdFromZip);
check('Flat OPC 字符串被正确识别',
  typeof flatOpc === 'string' && /^<\?xml[\s\S]*<pkg:package/.test(flatOpc),
  '生成的 Flat OPC 前缀：' + flatOpc.slice(0, 60));

section('Flat OPC 转换结果');
contains('一级标题', mdFromFlat, '# 测试文档标题');
contains('普通段落', mdFromFlat, '这是一段普通正文，包含中英文混排内容。');
contains('加粗', mdFromFlat, '**加粗文字**');
contains('斜体', mdFromFlat, '*斜体文字*');
contains('删除线', mdFromFlat, '~~删除线效果~~');
contains('行内代码', mdFromFlat, '`行内代码`');
contains('无序列表', mdFromFlat, '- 无序列表项一');
contains('嵌套列表', mdFromFlat, '    - 嵌套列表项');
contains('有序列表递增', mdFromFlat, '2. 有序列表项二');
contains('表格表头', mdFromFlat, '| 名称 | 数量 | 备注 |');
contains('表格数据行', mdFromFlat, '| 甲项目 | 10 | 正常 |');
contains('引用段落', mdFromFlat, '> 引用样式的段落');

section('两条路径产出必须一致（核心断言）');
check('ZIP 与 Flat OPC 结果完全相同', mdFromZip === mdFromFlat,
  mdFromZip === mdFromFlat ? '' :
    '长度 zip=' + mdFromZip.length + ' flat=' + mdFromFlat.length);

section('回归：不再出现 atob / base64 误用');
var apiSrc = fs.readFileSync(path.join(__dirname, '..', 'src', 'core', 'office-api.js'), 'utf8');
check('office-api.js 的 readAsMarkdown 不再调用 atob',
  !/readAsMarkdown[\s\S]{0,600}?atob\(/.test(apiSrc),
  '仍存在 atob 调用');
check('office-api.js 把 getOoxml 结果直接传给 docxToMarkdown',
  /var ooxml = ooxmlResult\.value;[\s\S]{0,200}?docxToMarkdown\(ooxml/.test(apiSrc),
  '未直接传递 OOXML 字符串');

section('Flat OPC 结构健壮性');
check('空 package 不崩溃', (function () {
  try {
    docx2md.docxToMarkdown(
      '<?xml version="1.0"?><pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>'
    );
    return false; // 应当抛错
  } catch (e) {
    return /缺少 word\/document\.xml/.test(e.message);
  }
})());
check('缺少 pkg 命名空间前缀仍能解析', (function () {
  var broken = flatOpc.replace(/pkg:/g, '');
  try {
    var r = docx2md.docxToMarkdown(broken);
    return r.indexOf('# 测试文档标题') >= 0;
  } catch (e) { return false; }
})());

section('汇总');
console.log('='.repeat(60));
console.log('通过 ' + pass + ' 项，失败 ' + fail + ' 项');
if (fail) {
  failures.forEach(function (f, i) { console.log((i + 1) + '. ' + f); });
  process.exit(1);
} else {
  console.log('全部通过 ✓');
}

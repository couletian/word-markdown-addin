/*!
 * make-fixture.js — 生成测试用 .docx（零依赖，手工构造 OOXML + ZIP）
 * 运行：node test/make-fixture.js
 *
 * 目的：在没有 Word 的环境下验证 docx2md 的解析正确性。
 */
'use strict';

var fs = require('fs');
var path = require('path');
var zlib = require('zlib');

/* ------------------------------------------------------------------ *
 * ZIP 写入（stored + deflate 均支持，此处用 deflate 更接近真实 docx）
 * ------------------------------------------------------------------ */

var CRC_TABLE = (function () {
  var t = new Int32Array(256);
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  var c = -1;
  for (var i = 0; i < buf.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xff];
  return (c ^ -1) >>> 0;
}

function u16(n) { return [n & 0xff, (n >>> 8) & 0xff]; }
function u32(n) { return [n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]; }

function zip(files) {
  var localParts = [];
  var centralParts = [];
  var offset = 0;

  Object.keys(files).forEach(function (name) {
    var nameBuf = Buffer.from(name, 'utf8');
    var raw = Buffer.isBuffer(files[name]) ? files[name] : Buffer.from(files[name], 'utf8');
    var comp = zlib.deflateRawSync(raw, { level: 9 });
    var crc = crc32(raw);

    // 本地文件头
    var local = Buffer.concat([
      Buffer.from(u32(0x04034b50)),
      Buffer.from(u16(20)),           // version needed
      Buffer.from(u16(0x0800)),       // flags: UTF-8
      Buffer.from(u16(8)),            // method: deflate
      Buffer.from(u16(0)),            // time
      Buffer.from(u16(0)),            // date
      Buffer.from(u32(crc)),
      Buffer.from(u32(comp.length)),
      Buffer.from(u32(raw.length)),
      Buffer.from(u16(nameBuf.length)),
      Buffer.from(u16(0)),            // extra len
      nameBuf,
      comp
    ]);
    localParts.push(local);

    // 中央目录项
    var central = Buffer.concat([
      Buffer.from(u32(0x02014b50)),
      Buffer.from(u16(20)),           // version made by
      Buffer.from(u16(20)),           // version needed
      Buffer.from(u16(0x0800)),
      Buffer.from(u16(8)),
      Buffer.from(u16(0)),
      Buffer.from(u16(0)),
      Buffer.from(u32(crc)),
      Buffer.from(u32(comp.length)),
      Buffer.from(u32(raw.length)),
      Buffer.from(u16(nameBuf.length)),
      Buffer.from(u16(0)),            // extra
      Buffer.from(u16(0)),            // comment
      Buffer.from(u16(0)),            // disk
      Buffer.from(u16(0)),            // internal attrs
      Buffer.from(u32(0)),            // external attrs
      Buffer.from(u32(offset)),       // local header offset
      nameBuf
    ]);
    centralParts.push(central);

    offset += local.length;
  });

  var centralBuf = Buffer.concat(centralParts);
  var eocd = Buffer.concat([
    Buffer.from(u32(0x06054b50)),
    Buffer.from(u16(0)),
    Buffer.from(u16(0)),
    Buffer.from(u16(centralParts.length)),
    Buffer.from(u16(centralParts.length)),
    Buffer.from(u32(centralBuf.length)),
    Buffer.from(u32(offset)),
    Buffer.from(u16(0))
  ]);

  return Buffer.concat(localParts.concat([centralBuf, eocd]));
}

/* ------------------------------------------------------------------ *
 * 构造测试文档
 * ------------------------------------------------------------------ */

var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function p(text, styleId) {
  var pPr = styleId ? '<w:pPr><w:pStyle w:val="' + styleId + '"/></w:pPr>' : '';
  return '<w:p>' + pPr + (text ? '<w:r><w:t xml:space="preserve">' + text + '</w:t></w:r>' : '') + '</w:p>';
}

function pRuns(runs, styleId) {
  var pPr = styleId ? '<w:pPr><w:pStyle w:val="' + styleId + '"/></w:pPr>' : '';
  return '<w:p>' + pPr + runs + '</w:p>';
}

function run(text, opts) {
  opts = opts || {};
  var rPr = '';
  var props = [];
  if (opts.b) props.push('<w:b/>');
  if (opts.i) props.push('<w:i/>');
  if (opts.strike) props.push('<w:strike/>');
  if (opts.code) props.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>');
  if (opts.vert) props.push('<w:vertAlign w:val="' + opts.vert + '"/>');
  if (props.length) rPr = '<w:rPr>' + props.join('') + '</w:rPr>';
  return '<w:r>' + rPr + '<w:t xml:space="preserve">' + text + '</w:t></w:r>';
}

function listP(text, numId, ilvl, runs) {
  return '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>' +
    '<w:numPr><w:ilvl w:val="' + ilvl + '"/><w:numId w:val="' + numId + '"/></w:numPr>' +
    '</w:pPr>' + (runs || '<w:r><w:t xml:space="preserve">' + text + '</w:t></w:r>') + '</w:p>';
}

function table(rows) {
  var xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>';
  rows.forEach(function (cells, ri) {
    xml += '<w:tr>';
    cells.forEach(function (cellText) {
      xml += '<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
        '<w:p>' + (ri === 0
          ? '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + cellText + '</w:t></w:r>'
          : '<w:r><w:t xml:space="preserve">' + cellText + '</w:t></w:r>') +
        '</w:p></w:tc>';
    });
    xml += '</w:tr>';
  });
  return xml + '</w:tbl>';
}

var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document ' + NS + '><w:body>' +
  pRuns(run('测试文档标题'), 'Heading1') +
  p('这是一段普通正文，包含中英文混排内容。') +
  pRuns(run('这里有 ') + run('加粗文字', { b: true }) + run(' 和 ') +
    run('斜体文字', { i: true }) + run(' 以及 ') + run('行内代码', { code: true }) + run('。')) +
  pRuns(run('删除线效果', { strike: true })) +
  p('二级标题占位', 'Heading2') +
  p('三级标题占位', 'Heading3') +
  listP('无序列表项一', 1, 0) +
  listP('无序列表项二', 1, 0) +
  listP('嵌套列表项', 1, 1) +
  listP('有序列表项一', 2, 0) +
  listP('有序列表项二', 2, 0) +
  table([
    ['名称', '数量', '备注'],
    ['甲项目', '10', '正常'],
    ['乙项目', '20', '待定']
  ]) +
  p('引用样式的段落', 'Quote') +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>' + // 空标题（应被跳过）
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
  '</w:body></w:document>';

var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:styles ' + NS + '>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>' +
  '<w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>' +
  '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>' +
  '</w:styles>';

var numberingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:numbering ' + NS + '>' +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1">' +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '</w:numbering>';

var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
  '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
  '</Types>';

var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
  '</Relationships>';

var docx = zip({
  '[Content_Types].xml': contentTypes,
  '_rels/.rels': rootRels,
  'word/document.xml': documentXml,
  'word/styles.xml': stylesXml,
  'word/numbering.xml': numberingXml,
  'word/_rels/document.xml.rels': docRels
});

var outPath = path.join(__dirname, 'fixture.docx');
fs.writeFileSync(outPath, docx);
console.log('已生成测试文档：' + outPath + '（' + (docx.length / 1024).toFixed(1) + ' KB）');

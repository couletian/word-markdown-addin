/*!
 * docx2md.js — Word (.docx) → Markdown 转换器
 * ------------------------------------------------------------------
 * 直接解析 OOXML（word/document.xml），保留：
 *   标题、段落、加粗/斜体/删除线、行内代码（等宽字体）、
 *   有序/无序/嵌套列表、表格（含 GFM 对齐）、引用块、
 *   超链接、图片占位、水平线、软/硬换行、上标下标。
 *
 * 依赖：WordMd.zip（zip.js）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./zip.js'));
  } else {
    root.WordMd = root.WordMd || {};
    root.WordMd.docx2md = factory(root.WordMd.zip);
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function (Zip) {
  'use strict';

  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /* ------------------------------------------------------------------ *
   * XML 辅助（用 DOMParser；Node 环境可用内置的轻量解析回退）
   * ------------------------------------------------------------------ */

  /* ------------------------------------------------------------------ *
   * 轻量 XML 解析器（Node 等无 DOMParser 环境的回退实现）
   * 仅实现本模块用到的 DOM 子集，避免引入外部依赖。
   * ------------------------------------------------------------------ */

  function MiniNode(type, name, ns) {
    this.nodeType = type;
    this.localName = name || '';
    this.nodeName = name || '';
    this.namespaceURI = ns || '';
    this.childNodes = [];
    this.attributes = {};
    this.parentNode = null;
    this._text = '';
  }

  MiniNode.prototype.appendChild = function (child) {
    child.parentNode = this;
    this.childNodes.push(child);
    return child;
  };
  MiniNode.prototype.insertBefore = function (child, ref) {
    var idx = this.childNodes.indexOf(ref);
    child.parentNode = this;
    if (idx < 0) this.childNodes.push(child);
    else this.childNodes.splice(idx, 0, child);
    return child;
  };
  MiniNode.prototype.cloneNode = function () {
    var n = new MiniNode(this.nodeType, this.localName, this.namespaceURI);
    n._text = this._text;
    n.attributes = Object.assign({}, this.attributes);
    this.childNodes.forEach(function (c) { n.appendChild(c.cloneNode(true)); });
    return n;
  };
  Object.defineProperty(MiniNode.prototype, 'firstChild', {
    get: function () { return this.childNodes[0] || null; }
  });
  Object.defineProperty(MiniNode.prototype, 'nextSibling', {
    get: function () {
      if (!this.parentNode) return null;
      var sib = this.parentNode.childNodes;
      var idx = sib.indexOf(this);
      return idx >= 0 ? (sib[idx + 1] || null) : null;
    }
  });
  Object.defineProperty(MiniNode.prototype, 'previousSibling', {
    get: function () {
      if (!this.parentNode) return null;
      var sib = this.parentNode.childNodes;
      var idx = sib.indexOf(this);
      return idx > 0 ? sib[idx - 1] : null;
    }
  });
  Object.defineProperty(MiniNode.prototype, 'textContent', {
    get: function () {
      if (this.nodeType === 3) return this._text;
      return this.childNodes.map(function (c) { return c.textContent; }).join('');
    }
  });
  MiniNode.prototype.getAttribute = function (name) {
    return Object.prototype.hasOwnProperty.call(this.attributes, name) ? this.attributes[name] : null;
  };
  MiniNode.prototype.getAttributeNS = function (ns, local) {
    var key = (ns || '') + '|' + local;
    return Object.prototype.hasOwnProperty.call(this.attributes, key) ? this.attributes[key] : null;
  };
  MiniNode.prototype.getElementsByTagNameNS = function (ns, local) {
    var out = [];
    (function walk(node) {
      node.childNodes.forEach(function (c) {
        if (c.nodeType !== 1) return;
        if (c.localName === local && (!ns || c.namespaceURI === ns)) out.push(c);
        walk(c);
      });
    })(this);
    return out;
  };
  MiniNode.prototype.getElementsByTagName = function (name) {
    var out = [];
    var target = String(name).replace(/^.*:/, '');
    (function walk(node) {
      node.childNodes.forEach(function (c) {
        if (c.nodeType !== 1) return;
        if (c.localName === target) out.push(c);
        walk(c);
      });
    })(this);
    return out;
  };

  var ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

  function decodeEntities(s) {
    return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, function (m, ent) {
      if (ent.charAt(0) === '#') {
        var code = ent.charAt(1).toLowerCase() === 'x'
          ? parseInt(ent.slice(2), 16)
          : parseInt(ent.slice(1), 10);
        return isNaN(code) ? m : String.fromCodePoint(code);
      }
      return Object.prototype.hasOwnProperty.call(ENTITIES, ent) ? ENTITIES[ent] : m;
    });
  }

  /** 极简 XML 解析：支持元素、属性、文本、注释、CDATA、自闭合 */
  function miniParseXml(text) {
    var root = new MiniNode(1, '#document', '');
    var stack = [root];
    var i = 0;
    var len = text.length;

    while (i < len) {
      var lt = text.indexOf('<', i);
      if (lt < 0) {
        var tail = text.slice(i);
        if (tail.trim()) {
          var tn = new MiniNode(3, '#text', '');
          tn._text = decodeEntities(tail);
          stack[stack.length - 1].appendChild(tn);
        }
        break;
      }
      if (lt > i) {
        var raw = text.slice(i, lt);
        if (raw.trim()) {
          var tn2 = new MiniNode(3, '#text', '');
          tn2._text = decodeEntities(raw);
          stack[stack.length - 1].appendChild(tn2);
        }
      }

      if (text.startsWith('<!--', lt)) {
        var endC = text.indexOf('-->', lt);
        i = endC < 0 ? len : endC + 3;
        continue;
      }
      if (text.startsWith('<![CDATA[', lt)) {
        var endCd = text.indexOf(']]>', lt);
        var cd = text.slice(lt + 9, endCd < 0 ? len : endCd);
        var cdn = new MiniNode(3, '#text', '');
        cdn._text = cd;
        stack[stack.length - 1].appendChild(cdn);
        i = endCd < 0 ? len : endCd + 3;
        continue;
      }
      if (text.startsWith('<?', lt) || text.startsWith('<!', lt)) {
        var endP = text.indexOf('>', lt);
        i = endP < 0 ? len : endP + 1;
        continue;
      }

      var gt = text.indexOf('>', lt);
      if (gt < 0) break;
      var inner = text.slice(lt + 1, gt);
      var selfClosing = inner.charAt(inner.length - 1) === '/';
      if (selfClosing) inner = inner.slice(0, -1);

      if (inner.charAt(0) === '/') {
        // 结束标签
        if (stack.length > 1) stack.pop();
        i = gt + 1;
        continue;
      }

      // 开始标签：解析标签名与属性
      var nameMatch = /^([^\s/>]+)/.exec(inner);
      var tagName = nameMatch ? nameMatch[1] : '';
      var colon = tagName.indexOf(':');
      var prefix = colon > 0 ? tagName.slice(0, colon) : '';
      var local = colon > 0 ? tagName.slice(colon + 1) : tagName;

      var node = new MiniNode(1, local, '');
      node.nodeName = tagName;
      node._prefix = prefix;

      var attrRe = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
      var am;
      while ((am = attrRe.exec(inner)) !== null) {
        var an = am[1];
        var av = decodeEntities(am[3] != null ? am[3] : am[4]);
        node.attributes[an] = av;
        var ac = an.indexOf(':');
        if (ac > 0) {
          node.attributes[an.slice(0, ac) + '|' + an.slice(ac + 1)] = av;
        }
      }

      stack[stack.length - 1].appendChild(node);
      if (!selfClosing) stack.push(node);
      i = gt + 1;
    }

    // 解析完成后回填命名空间（按 xmlns 声明，仅需 w: 与 r:）
    (function assignNs(node) {
      node.childNodes.forEach(function (c) {
        if (c.nodeType !== 1) return;
        var declared = c.attributes['xmlns:' + c._prefix];
        if (declared) c.namespaceURI = declared;
        else c.namespaceURI = node.namespaceURI;
        assignNs(c);
      });
    })(root);

    return root;
  }

  function parseXml(xmlText) {
    if (typeof DOMParser !== 'undefined') {
      return new DOMParser().parseFromString(xmlText, 'application/xml');
    }
    return miniParseXml(xmlText);
  }

  /** 取元素下所有指定本地名的直接/后代节点 */
  function tags(node, localName) {
    if (!node) return [];
    var list = node.getElementsByTagNameNS
      ? node.getElementsByTagNameNS(W_NS, localName)
      : node.getElementsByTagName('w:' + localName);
    return Array.prototype.slice.call(list);
  }

  function children(node, localName) {
    var out = [];
    if (!node || !node.childNodes) return out;
    for (var i = 0; i < node.childNodes.length; i++) {
      var c = node.childNodes[i];
      if (c.nodeType !== 1) continue;
      var ln = c.localName || String(c.nodeName).replace(/^.*:/, '');
      if (ln === localName) out.push(c);
    }
    return out;
  }

  function attr(node, localName) {
    if (!node) return null;
    if (node.getAttributeNS) {
      var v = node.getAttributeNS(W_NS, localName);
      if (v != null) return v;
    }
    return node.getAttribute ? node.getAttribute('w:' + localName) : null;
  }

  function textOf(node) {
    return node ? (node.textContent || '') : '';
  }

  /* ------------------------------------------------------------------ *
   * 样式表解析：styleId -> {name, outlineLevel}
   * ------------------------------------------------------------------ */

  function parseStyles(xml) {
    var map = {};
    if (!xml) return map;
    var doc = parseXml(xml);
    var styles = tags(doc, 'style');
    styles.forEach(function (st) {
      var id = attr(st, 'styleId');
      if (!id) return;
      var nameEl = children(st, 'name');
      var name = nameEl.length ? attr(nameEl[0], 'val') : '';
      var pPr = children(st, 'pPr')[0];
      var outline = null;
      if (pPr) {
        var o = children(pPr, 'outlineLvl');
        if (o.length) outline = parseInt(attr(o[0], 'val'), 10);
      }
      map[id] = { name: name || '', outline: outline };
    });
    return map;
  }

  /* ------------------------------------------------------------------ *
   * numbering.xml：numId -> {abstractNumId -> {ilvl -> {numFmt, lvlText}}}
   * ------------------------------------------------------------------ */

  function parseNumbering(xml) {
    var abstractMap = {};
    var numMap = {};
    if (!xml) return { numMap: numMap, abstractMap: abstractMap };
    var doc = parseXml(xml);

    tags(doc, 'abstractNum').forEach(function (an) {
      var id = attr(an, 'abstractNumId');
      var levels = {};
      children(an, 'lvl').forEach(function (lvl) {
        var ilvl = attr(lvl, 'ilvl');
        var fmtEl = children(lvl, 'numFmt');
        var txtEl = children(lvl, 'lvlText');
        levels[ilvl] = {
          numFmt: fmtEl.length ? attr(fmtEl[0], 'val') : 'decimal',
          lvlText: txtEl.length ? attr(txtEl[0], 'val') : '%1.'
        };
      });
      abstractMap[id] = levels;
    });

    tags(doc, 'num').forEach(function (n) {
      var numId = attr(n, 'numId');
      var ab = children(n, 'abstractNumId');
      if (ab.length) numMap[numId] = attr(ab[0], 'val');
    });

    return { numMap: numMap, abstractMap: abstractMap };
  }

  /* ------------------------------------------------------------------ *
   * 关系表：rId -> 外部目标
   * ------------------------------------------------------------------ */

  function parseRels(xml) {
    var map = {};
    if (!xml) return map;
    var doc = parseXml(xml);
    var rels = doc.getElementsByTagName('Relationship');
    for (var i = 0; i < rels.length; i++) {
      var r = rels[i];
      var id = r.getAttribute('Id');
      var target = r.getAttribute('Target');
      var mode = r.getAttribute('TargetMode');
      if (id) map[id] = { target: target || '', external: mode === 'External' };
    }
    return map;
  }

  /* ------------------------------------------------------------------ *
   * 行内运行解析
   * ------------------------------------------------------------------ */

  var MONO_RE = /consolas|courier|mono|menlo|monaco|source code|fira code|jetbrains/i;

  function runProps(rPr) {
    var p = { b: false, i: false, strike: false, code: false, sup: false, sub: false };
    if (!rPr) return p;
    if (children(rPr, 'b').length) {
      var bEl = children(rPr, 'b')[0];
      p.b = attr(bEl, 'val') !== '0' && attr(bEl, 'val') !== 'false';
    }
    if (children(rPr, 'i').length) {
      var iEl = children(rPr, 'i')[0];
      p.i = attr(iEl, 'val') !== '0' && attr(iEl, 'val') !== 'false';
    }
    if (children(rPr, 'strike').length || children(rPr, 'dstrike').length) p.strike = true;
    if (children(rPr, 'vertAlign').length) {
      var va = attr(children(rPr, 'vertAlign')[0], 'val');
      if (va === 'superscript') p.sup = true;
      if (va === 'subscript') p.sub = true;
    }
    var fonts = children(rPr, 'rFonts');
    if (fonts.length) {
      var f = fonts[0];
      var name = attr(f, 'ascii') || attr(f, 'hAnsi') || attr(f, 'cs') || '';
      if (MONO_RE.test(name)) p.code = true;
    }
    var shd = children(rPr, 'shd');
    if (shd.length && attr(shd[0], 'fill') && attr(shd[0], 'fill') !== 'auto') {
      // 带底纹的等宽字体，进一步确认是代码
      if (p.code) p.code = true;
    }
    return p;
  }

  function applyEmphasis(text, p) {
    if (!text) return text;
    var out = text;
    if (p.code) {
      // 行内代码：含反引号时用双反引号包裹
      var tick = out.indexOf('`') >= 0 ? '``' : '`';
      var pad = /^`|`$/.test(out) ? ' ' : '';
      return tick + pad + out + pad + tick;
    }
    // 两端为空格时把空格移到标记外，避免 Markdown 渲染异常
    var lead = /^\s/.test(out) ? out.match(/^\s+/)[0] : '';
    var tail = /^\s/.test(out.slice(-1)) || /\s$/.test(out) ? (out.match(/\s+$/) || [''])[0] : '';
    var core = out.slice(lead.length, out.length - tail.length);
    if (!core) return out;
    var esc = core.replace(/([*_~\\`\[\]])/g, '\\$1');
    if (p.b && p.i) core = '***' + esc + '***';
    else if (p.b) core = '**' + esc + '**';
    else if (p.i) core = '*' + esc + '*';
    else core = core;
    if (p.strike) core = '~~' + core + '~~';
    if (p.sup) core = '<sup>' + core + '</sup>';
    if (p.sub) core = '<sub>' + core + '</sub>';
    return lead + core + tail;
  }

  /** 处理一个 w:r 节点 */
  function renderRun(run, ctx) {
    var out = '';
    var rPr = children(run, 'rPr')[0];
    var props = runProps(rPr);
    // 表头单元格：Word 中通常整格加粗，但 Markdown 表头本身已表达语义，无需再包 **
    if (ctx.suppressBold) props.b = false;

    var kids = run.childNodes;
    for (var i = 0; i < kids.length; i++) {
      var c = kids[i];
      if (c.nodeType !== 1) continue;
      var ln = c.localName || String(c.nodeName).replace(/^.*:/, '');

      if (ln === 't') {
        out += applyEmphasis(c.textContent || '', props);
      } else if (ln === 'tab') {
        out += '\t';
      } else if (ln === 'br') {
        var brType = attr(c, 'type');
        out += brType === 'page' ? '\n\n' : '  \n';
      } else if (ln === 'noBreakHyphen') {
        out += '-';
      } else if (ln === 'softHyphen') {
        out += '';
      } else if (ln === 'sym') {
        var ch = attr(c, 'char');
        if (ch) {
          try { out += String.fromCharCode(parseInt(ch, 16)); } catch (e) { /* ignore */ }
        }
      } else if (ln === 'drawing' || ln === 'pict' || ln === 'object') {
        out += ctx.imagePlaceholder();
      }
    }
    return out;
  }

  /** 处理超链接、书签、插入/删除等容器 */
  function renderInlineContainer(node, ctx) {
    var out = '';
    var kids = node.childNodes;
    for (var i = 0; i < kids.length; i++) {
      var c = kids[i];
      if (c.nodeType !== 1) continue;
      var ln = c.localName || String(c.nodeName).replace(/^.*:/, '');
      if (ln === 'r') {
        out += renderRun(c, ctx);
      } else if (ln === 'hyperlink') {
        var rid = attr(c, 'id') || c.getAttributeNS(R_NS, 'id');
        var inner = renderInlineChildren(c, ctx);
        var rel = rid && ctx.rels[rid];
        var href = rel ? rel.target : '';
        var anchor = attr(c, 'anchor');
        if (!href && anchor) href = '#' + anchor;
        if (href && inner.trim()) out += '[' + inner + '](' + href + ')';
        else out += inner;
      } else if (ln === 'ins') {
        out += renderInlineChildren(c, ctx); // 修订：接受插入内容
      } else if (ln === 'del') {
        // 修订：删除内容默认忽略
      } else if (ln === 'smartTag' || ln === 'sdt' || ln === 'sdtContent' || ln === 'bookmarkStart' ||
                 ln === 'bookmarkEnd' || ln === 'proofErr' || ln === 'commentRangeStart' ||
                 ln === 'commentRangeEnd' || ln === 'commentReference') {
        out += renderInlineChildren(c, ctx);
      } else if (ln === 'fldSimple') {
        out += renderInlineChildren(c, ctx);
      }
    }
    return out;
  }

  function renderInlineChildren(node, ctx) {
    var out = '';
    for (var i = 0; i < node.childNodes.length; i++) {
      var c = node.childNodes[i];
      if (c.nodeType !== 1) continue;
      var ln = c.localName || String(c.nodeName).replace(/^.*:/, '');
      if (ln === 'r') out += renderRun(c, ctx);
      else if (ln === 'hyperlink' || ln === 'ins' || ln === 'smartTag' || ln === 'sdt' ||
               ln === 'sdtContent' || ln === 'fldSimple' || ln === 'bookmarkStart' ||
               ln === 'bookmarkEnd' || ln === 'proofErr') {
        out += renderInlineContainer(c, ctx);
      } else if (ln === 't') {
        out += c.textContent || '';
      } else if (ln === 'br') {
        out += '  \n';
      }
    }
    return out;
  }

  /* ------------------------------------------------------------------ *
   * 段落属性解析
   * ------------------------------------------------------------------ */

  function paraInfo(p, ctx) {
    var pPr = children(p, 'pPr')[0] || null;
    var info = {
      styleId: null,
      heading: 0,
      numId: null,
      ilvl: 0,
      quote: false,
      align: null,
      isListItem: false,
      ordered: false,
      numFmt: 'decimal',
      lvlText: '%1.',
      codeBlock: false,
      indent: 0,
      hr: false
    };
    if (!pPr) return info;

    var ps = children(pPr, 'pStyle');
    if (ps.length) info.styleId = attr(ps[0], 'val');

    var ol = children(pPr, 'outlineLvl');
    var outline = ol.length ? parseInt(attr(ol[0], 'val'), 10) : null;

    // 标题识别：样式名 / outlineLvl
    var styleDef = info.styleId && ctx.styles[info.styleId];
    var styleName = styleDef ? styleDef.name : '';
    var headingMatch = /^heading\s*([1-9])$/i.exec(styleName) || /^标题\s*([1-9])$/.exec(styleName);
    if (headingMatch) {
      info.heading = Math.min(6, parseInt(headingMatch[1], 10));
    } else if (outline != null && outline >= 0 && outline <= 5) {
      info.heading = outline + 1;
    } else if (styleDef && styleDef.outline != null && styleDef.outline >= 0 && styleDef.outline <= 5) {
      info.heading = styleDef.outline + 1;
    } else if (/^title$/i.test(styleName) || /^标题$/i.test(styleName)) {
      info.heading = 1;
    } else if (/^subtitle$/i.test(styleName) || /^副标题$/i.test(styleName)) {
      info.heading = 2;
    } else if (/quote|引用|Intense Quote/i.test(styleName)) {
      info.quote = true;
    }

    var numPr = children(pPr, 'numPr')[0];
    if (numPr) {
      var numIdEl = children(numPr, 'numId');
      var ilvlEl = children(numPr, 'ilvl');
      if (numIdEl.length) {
        var nid = attr(numIdEl[0], 'val');
        if (nid && nid !== '0') {
          info.numId = nid;
          info.isListItem = true;
          info.ilvl = ilvlEl.length ? parseInt(attr(ilvlEl[0], 'val'), 10) || 0 : 0;
        }
      }
    }

    var jc = children(pPr, 'jc');
    if (jc.length) info.align = attr(jc[0], 'val');

    var ind = children(pPr, 'ind');
    if (ind.length) info.indent = parseInt(attr(ind[0], 'left') || '0', 10) || 0;

    var shd = children(pPr, 'shd');
    var pbdr = children(pPr, 'pBdr');
    if (shd.length && attr(shd[0], 'fill') && attr(shd[0], 'fill') !== 'auto') {
      // 段落底纹 + 等宽字体 → 代码块
      info.shaded = true;
    }
    if (pbdr.length) {
      var bottom = children(pbdr[0], 'bottom');
      if (bottom.length) {
        var sz = parseInt(attr(bottom[0], 'sz') || '0', 10);
        if (sz > 0 && sz <= 12) info.hr = true; // 细底边框常见于分隔线
      }
    }

    // 编号格式
    if (info.isListItem) {
      var abstractId = ctx.numbering.numMap[info.numId];
      var levels = abstractId != null ? ctx.numbering.abstractMap[abstractId] : null;
      var lvl = levels ? levels[String(info.ilvl)] : null;
      if (lvl) {
        info.numFmt = lvl.numFmt;
        info.lvlText = lvl.lvlText;
      }
      info.ordered = !/bullet|none/.test(info.numFmt);
    }

    return info;
  }

  /* ------------------------------------------------------------------ *
   * 主转换
   * ------------------------------------------------------------------ */

  /* ------------------------------------------------------------------ *
   * Flat OPC 支持
   * ------------------------------------------------------------------
   * Office.js 的 getOoxml() 返回的不是 base64，而是 Flat OPC 格式的 XML：
   *
   *   <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
   *     <pkg:part pkg:name="/word/document.xml" pkg:contentType="...">
   *       <pkg:xmlData><w:document …>…</w:document></pkg:xmlData>
   *     </pkg:part>
   *     <pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">
   *       <pkg:binaryData>iVBORw0…</pkg:binaryData>
   *     </pkg:part>
   *   </pkg:package>
   *
   * 这里把它还原成与 ZIP 解包等价的 { partName: Uint8Array } 结构，
   * 让后续所有转换逻辑完全复用，无需改动。
   * ------------------------------------------------------------------ */

  function escapeXmlText(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function escapeXmlAttr(s) {
    return escapeXmlText(s).replace(/"/g, '&quot;');
  }

  /** 把 MiniNode 子树序列化回 XML 文本（供 parseStyles 等复用） */
  function serializeXml(node) {
    if (!node) return '';
    if (node.nodeType === 3) {
      var t = node._text != null ? node._text : node.textContent;
      return escapeXmlText(t);
    }
    if (node.nodeType !== 1) return '';
    var name = node.nodeName || node.localName;
    var out = '<' + name;
    var attrs = node.attributes || {};
    Object.keys(attrs).forEach(function (k) {
      if (k.indexOf('|') >= 0) return; // 跳过内部别名键
      out += ' ' + k + '="' + escapeXmlAttr(attrs[k]) + '"';
    });
    if (!node.childNodes || node.childNodes.length === 0) return out + '/>';
    out += '>';
    for (var i = 0; i < node.childNodes.length; i++) out += serializeXml(node.childNodes[i]);
    return out + '</' + name + '>';
  }

  function utf8Encode(str) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(str);
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(str, 'utf8'));
    var out = [];
    for (var i = 0; i < str.length; i++) {
      var c = str.charCodeAt(i);
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xc0 | (c >> 6), 0x80 | (c & 0x3f));
      else if (c >= 0xd800 && c <= 0xdbff && i + 1 < str.length) {
        var cp = 0x10000 + ((c - 0xd800) << 10) + (str.charCodeAt(++i) - 0xdc00);
        out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3f), 0x80 | ((cp >> 6) & 0x3f), 0x80 | (cp & 0x3f));
      } else out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 0x3f), 0x80 | (c & 0x3f));
    }
    return new Uint8Array(out);
  }

  function base64Decode(b64) {
    var clean = String(b64).replace(/[^A-Za-z0-9+/=]/g, '');
    if (typeof atob === 'function') {
      var bin = atob(clean);
      var out = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(clean, 'base64'));
    throw new Error('docx2md: 当前环境缺少 base64 解码能力');
  }

  /** Flat OPC XML → { partName: Uint8Array }，键名与 ZIP 一致（无前导 /） */
  function flatOpcToFiles(xmlText) {
    var doc = miniParseXml(xmlText);
    var partNodes = doc.getElementsByTagName('part');
    var files = {};
    for (var i = 0; i < partNodes.length; i++) {
      var p = partNodes[i];
      var rawName = p.getAttribute('pkg:name') || p.getAttribute('name');
      if (!rawName) continue;
      var name = String(rawName).replace(/^\/+/, '');

      var xmlData = null;
      var binData = null;
      for (var j = 0; j < p.childNodes.length; j++) {
        var c = p.childNodes[j];
        if (c.nodeType !== 1) continue;
        if (c.localName === 'xmlData') xmlData = c;
        else if (c.localName === 'binaryData') binData = c;
      }

      if (xmlData) {
        var el = null;
        for (var k = 0; k < xmlData.childNodes.length; k++) {
          if (xmlData.childNodes[k].nodeType === 1) { el = xmlData.childNodes[k]; break; }
        }
        if (el) files[name] = utf8Encode(serializeXml(el));
      } else if (binData) {
        files[name] = base64Decode(binData.textContent || '');
      }
    }
    return files;
  }

  /** 判断输入是否为 Flat OPC XML 文本（按命名空间识别，不依赖 pkg 前缀） */
  function isFlatOpc(input) {
    if (typeof input !== 'string') return false;
    var head = input.slice(0, 600);
    return /schemas\.microsoft\.com\/office\/2006\/xmlPackage/.test(head) ||
      /<(?:pkg:)?package[\s>]/.test(head);
  }

  /**
   * @param {Uint8Array|ArrayBuffer|string} input
   *   - Uint8Array / ArrayBuffer：.docx 二进制
   *   - string：Office.js getOoxml() 返回的 Flat OPC XML
   */
  function docxToMarkdown(input, options) {
    options = options || {};
    var files = isFlatOpc(input) ? flatOpcToFiles(input) : Zip.unzip(input);
    var documentXml = Zip.text(files['word/document.xml']);
    if (!documentXml) throw new Error('docx2md: 缺少 word/document.xml');

    var styles = parseStyles(files['word/styles.xml'] ? Zip.text(files['word/styles.xml']) : null);
    var numbering = parseNumbering(files['word/numbering.xml'] ? Zip.text(files['word/numbering.xml']) : null);
    var rels = parseRels(files['word/_rels/document.xml.rels'] ? Zip.text(files['word/_rels/document.xml.rels']) : null);

    var ctx = {
      styles: styles,
      numbering: numbering,
      rels: rels,
      imagePlaceholder: function () { return options.imagePlaceholder || '![图片]'; }
    };

    var doc = parseXml(documentXml);
    var body = tags(doc, 'body')[0];
    if (!body) throw new Error('docx2md: document.xml 中没有 body');

    var blocks = [];
    var listStack = []; // {numId, level, ordered, counter}

    /** 追加块；kind 用于控制块间连接符（list 块之间用单换行） */
    function push(text, kind, listType) {
      blocks.push({ text: text, kind: kind || 'block', listType: listType || null });
    }

    function currentList(numId, level) {
      // 维护层级栈，保证列表连续性
      while (listStack.length && listStack[listStack.length - 1].level > level) listStack.pop();
      if (listStack.length && listStack[listStack.length - 1].level === level &&
          listStack[listStack.length - 1].numId === numId) {
        return listStack[listStack.length - 1];
      }
      while (listStack.length && listStack[listStack.length - 1].level >= level) listStack.pop();
      var item = { numId: numId, level: level, ordered: false, counter: 0 };
      listStack.push(item);
      return item;
    }

    var node = body.firstChild;
    while (node) {
      if (node.nodeType === 1) {
        var ln = node.localName || String(node.nodeName).replace(/^.*:/, '');

        if (ln === 'p') {
          var info = paraInfo(node, ctx);
          var text = renderInlineContainer(node, ctx);
          var plain = text.replace(/\*\*|__|\*|_|~~|`/g, '').trim();

          if (!plain && !/!\[/.test(text)) {
            listStack.length = 0;
            // 空段落 → 作为块间空行处理（Markdown 用空行分隔，这里跳过）
            node = node.nextSibling;
            continue;
          }

          if (info.hr && !plain) {
            push('---');
            listStack.length = 0;
            node = node.nextSibling;
            continue;
          }

          if (info.heading) {
            listStack.length = 0;
            push('#' + repeat('#', info.heading - 1) + ' ' + text.trim());
          } else if (info.isListItem) {
            var level = info.ilvl;
            var stackItem = currentList(info.numId, level);
            stackItem.ordered = info.ordered;
            var indent = repeat('    ', level);

            var marker;
            if (info.ordered) {
              stackItem.counter++;
              marker = String(stackItem.counter) + '. ';
            } else {
              marker = '- ';
            }
            // 引用式列表也加 > 前缀
            var line = indent + marker + text.trim();
            if (info.quote) line = indent + '> ' + marker + text.trim();
            push(line, 'list', info.ordered ? 'ol' : 'ul');
          } else if (info.quote) {
            listStack.length = 0;
            push('> ' + text.trim());
          } else {
            listStack.length = 0;
            push(text.trim());
          }
        } else if (ln === 'tbl') {
          listStack.length = 0;
          push(renderTable(node, ctx));
        } else if (ln === 'sdt') {
          // 结构化文档标签：递归其内容
          var content = children(node, 'sdtContent')[0];
          if (content) {
            var innerDoc = content;
            var child = innerDoc.firstChild;
            while (child) {
              var next = child.nextSibling;
              body.insertBefore(child.cloneNode(true), node);
              child = next;
            }
            node = node.previousSibling;
            if (!node) {
              node = body.firstChild;
              continue;
            }
          }
        } else if (ln === 'sectPr') {
          // 节属性：忽略
        }
      }
      node = node.nextSibling;
    }

    var md = '';
    for (var bi = 0; bi < blocks.length; bi++) {
      if (bi === 0) { md = blocks[bi].text; continue; }
      var prev = blocks[bi - 1];
      var cur = blocks[bi];
      var sameList = prev.kind === 'list' && cur.kind === 'list' && prev.listType === cur.listType;
      // 同类型列表项之间用单换行；列表类型切换或其他块之间用空行
      md += sameList ? '\n' : '\n\n';
      md += cur.text;
    }
    // 压缩多余空行
    md = md.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+\n/g, '\n').trim();
    return md + '\n';
  }

  function repeat(ch, n) {
    var s = '';
    while (n-- > 0) s += ch;
    return s;
  }

  /* ------------------------------------------------------------------ *
   * 表格
   * ------------------------------------------------------------------ */

  function renderTable(tbl, ctx) {
    var rows = children(tbl, 'tr');
    if (!rows.length) return '';

    var grid = [];
    var aligns = [];
    var maxCols = 0;

    rows.forEach(function (tr) {
      var cells = children(tr, 'tc');
      var row = [];
      cells.forEach(function (tc) {
        var tcPr = children(tc, 'tcPr')[0];
        var span = 1;
        var align = null;
        if (tcPr) {
          var gs = children(tcPr, 'gridSpan');
          if (gs.length) span = parseInt(attr(gs[0], 'val'), 10) || 1;
          var jc = children(tcPr, 'jc');
          if (jc.length) align = attr(jc[0], 'val');
        }
        var cellText = '';
        var ps = children(tc, 'p');
        var parts = [];
        var isHeaderRow = (tr === rows[0]);
        var cellCtx = Object.create(ctx);
        // 表头行整格加粗是 Word 的排版惯例，转 Markdown 时由表头语法表达，避免 ** 冗余
        cellCtx.suppressBold = isHeaderRow;
        ps.forEach(function (p) {
          var t = renderInlineContainer(p, cellCtx).trim();
          if (t) parts.push(t);
        });
        cellText = parts.join('<br />');

        row.push({ text: cellText, span: span, align: align });
        if (align && !aligns[row.length - 1]) aligns[row.length - 1] = align;
      });
      maxCols = Math.max(maxCols, row.reduce(function (a, c) { return a + c.span; }, 0));
      grid.push(row);
    });

    // 补全列数并处理合并单元格（重复内容到被合并的列）
    grid = grid.map(function (row) {
      var out = [];
      row.forEach(function (c) {
        out.push(c.text);
        for (var k = 1; k < c.span; k++) out.push('');
      });
      while (out.length < maxCols) out.push('');
      return out.slice(0, maxCols);
    });

    var header = grid[0];
    var bodyRows = grid.slice(1);
    if (!bodyRows.length) {
      bodyRows = [header.map(function () { return ''; })];
    }

    var sep = [];
    for (var c = 0; c < maxCols; c++) {
      var a = aligns[c];
      if (a === 'center') sep.push(':---:');
      else if (a === 'right' || a === 'end') sep.push('---:');
      else sep.push('---');
    }

    function line(cells) {
      return '| ' + cells.map(function (t) { return String(t).replace(/\|/g, '\\|'); }).join(' | ') + ' |';
    }

    var out = [line(header), line(sep)];
    bodyRows.forEach(function (r) { out.push(line(r)); });
    return out.join('\n');
  }

  return {
    docxToMarkdown: docxToMarkdown,
    _internal: {
      parseStyles: parseStyles,
      parseNumbering: parseNumbering,
      runProps: runProps,
      applyEmphasis: applyEmphasis
    }
  };
});

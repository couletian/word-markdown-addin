/*!
 * markdown.js — 轻量 Markdown → Word 友好 HTML 转换器
 * ------------------------------------------------------------------
 * 无第三方依赖，可在浏览器（Office 插件页面）与 Node（单元测试）中运行。
 *
 * 支持的语法：
 *   块级：ATX/Setext 标题、围栏与缩进代码块、引用块、有序/无序/任务列表（含嵌套）、
 *         GFM 表格（含对齐）、水平线、HTML 块
 *   行内：转义、行内代码、加粗、斜体、粗斜体、删除线、高亮、链接、图片、
 *         自动链接、脚注式引用、硬换行（行尾两空格或反斜杠）
 *
 * 输出为 Word insertHtml 兼容的 HTML 片段，并针对 Word 的内联样式支持做了收敛。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.WordMd = root.WordMd || {};
    root.WordMd.markdown = factory();
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ------------------------------------------------------------------ *
   * 样式预设
   * ------------------------------------------------------------------ */

  var CODE_FONT = "'Consolas','Courier New',monospace";

  var PRESETS = {
    // 默认：现代中文字体，适合日常文档
    default: {
      font: "font-family:'Calibri','Microsoft YaHei',sans-serif;font-size:11pt",
      body: "font-family:'Calibri','Microsoft YaHei',sans-serif;font-size:11pt;line-height:1.6;color:#000000",
      p: 'margin:0 0 8pt 0',
      h1: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:22pt;font-weight:bold;color:#1f3864;margin:18pt 0 8pt 0",
      h2: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:17pt;font-weight:bold;color:#2e5496;margin:16pt 0 7pt 0",
      h3: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:14pt;font-weight:bold;color:#2e5496;margin:14pt 0 6pt 0",
      h4: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:12pt;font-weight:bold;color:#333333;margin:12pt 0 5pt 0",
      h5: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:11pt;font-weight:bold;color:#333333;margin:10pt 0 4pt 0",
      h6: "font-family:'Calibri Light','Microsoft YaHei',sans-serif;font-size:10.5pt;font-weight:bold;color:#666666;margin:10pt 0 4pt 0",
      quote: 'margin:8pt 0 8pt 18pt;padding-left:12pt;border-left:3pt solid #bfbfbf;color:#595959',
      code: "font-family:" + CODE_FONT + ';font-size:10pt;background:#f2f2f2',
      pre: "font-family:" + CODE_FONT + ";font-size:9.5pt;background:#f6f8fa;border:1pt solid #d0d7de;padding:8pt;margin:8pt 0;white-space:pre",
      table: 'border-collapse:collapse;width:100%;margin:8pt 0',
      th: 'border:1pt solid #999999;padding:5pt 8pt;background:#f2f2f2;font-weight:bold',
      td: 'border:1pt solid #999999;padding:5pt 8pt',
      hr: 'border:none;border-top:1pt solid #bfbfbf;margin:12pt 0'
    },

    // 中文学术/论文：宋体正文 + 黑体标题，小四号，1.5 倍行距
    thesis: {
      font: "font-family:'SimSun','宋体',serif;font-size:12pt",
      body: "font-family:'SimSun','宋体',serif;font-size:12pt;line-height:1.5;color:#000000",
      p: 'margin:0 0 6pt 0;text-indent:24pt',
      h1: "font-family:'SimHei','黑体',sans-serif;font-size:18pt;font-weight:bold;color:#000000;margin:18pt 0 10pt 0;text-align:center",
      h2: "font-family:'SimHei','黑体',sans-serif;font-size:15pt;font-weight:bold;color:#000000;margin:16pt 0 8pt 0",
      h3: "font-family:'SimHei','黑体',sans-serif;font-size:13.5pt;font-weight:bold;color:#000000;margin:14pt 0 6pt 0",
      h4: "font-family:'SimHei','黑体',sans-serif;font-size:12pt;font-weight:bold;color:#000000;margin:12pt 0 6pt 0",
      h5: "font-family:'SimHei','黑体',sans-serif;font-size:12pt;font-weight:bold;color:#000000;margin:10pt 0 4pt 0",
      h6: "font-family:'SimHei','黑体',sans-serif;font-size:12pt;font-weight:bold;color:#000000;margin:10pt 0 4pt 0",
      quote: "font-family:'KaiTi','楷体',serif;margin:8pt 0 8pt 24pt;padding-left:12pt;border-left:3pt solid #bfbfbf;color:#404040",
      code: "font-family:" + CODE_FONT + ';font-size:10.5pt',
      pre: "font-family:" + CODE_FONT + ";font-size:10pt;background:#f6f8fa;border:1pt solid #d0d7de;padding:8pt;margin:8pt 0;white-space:pre",
      table: 'border-collapse:collapse;width:100%;margin:8pt 0',
      th: "border:1pt solid #000000;padding:4pt 6pt;background:#f2f2f2;font-family:'SimHei','黑体',sans-serif;font-weight:bold;text-align:center",
      td: 'border:1pt solid #000000;padding:4pt 6pt',
      hr: 'border:none;border-top:1pt solid #000000;margin:12pt 0'
    },

    // 纯语义：不写内联样式，交给 Word 内置样式（Heading 1..6 / 列表 / 表格）
    semantic: null
  };

  var CHECKED = '\u2611';   // ☑
  var UNCHECKED = '\u2610'; // ☐

  /* ------------------------------------------------------------------ *
   * 安全过滤
   * ------------------------------------------------------------------ */

  /** 过滤 URL 协议，阻断 javascript: / vbscript: / data:text 等危险协议 */
  function sanitizeUrl(url) {
    var s = String(url == null ? '' : url).trim();
    if (!s) return '';
    // 去掉控制字符与前后空白后判断协议
    var probe = s.replace(/[\u0000-\u0020]/g, '').toLowerCase();
    if (/^(?:javascript|vbscript|data|file|about|blob):/.test(probe)) {
      // data: 仅允许图片类型
      if (/^data:image\/(?:png|jpe?g|gif|webp|bmp);base64,/i.test(probe)) return s;
      return '';
    }
    return s;
  }

  /**
   * 清洗透传的 HTML 片段：移除危险标签与事件属性
   * 保留 Word 支持的安全标签（br、u、sup、sub、span、div、table 等）
   */
  function sanitizeHtml(html) {
    var out = String(html);
    // 移除危险标签及其内容
    out = out.replace(
      /<\s*(script|style|iframe|object|embed|applet|frame|frameset|link|meta|base)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi,
      ''
    );
    // 移除未闭合的危险单标签
    out = out.replace(/<\s*(script|style|iframe|object|embed|applet|link|meta|base)\b[^>]*\/?>/gi, '');
    // 移除内联事件处理器（onclick=... 等）
    out = out.replace(/\son[a-z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');
    // 移除 javascript:/vbscript: 形式的属性值
    out = out.replace(
      /(\s(?:href|src|action|formaction|xlink:href)\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi,
      function (m, pre, dq, sq, bare) {
        var val = dq != null ? dq : (sq != null ? sq : bare);
        var safe = sanitizeUrl(val);
        return safe ? pre + '"' + safe.replace(/"/g, '&quot;') + '"' : '';
      }
    );
    return out;
  }

  /* ------------------------------------------------------------------ *
   * 基础工具
   * ------------------------------------------------------------------ */

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, '&#39;');
  }

  function repeat(ch, n) {
    var out = '';
    while (n-- > 0) out += ch;
    return out;
  }

  function expandTabs(line) {
    return line.replace(/\t/g, '    ');
  }

  function isBlank(line) {
    return /^\s*$/.test(line);
  }

  function hasCJK(s) {
    return /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/.test(s);
  }

  function styleAttr(preset, key) {
    if (!preset) return '';
    var v = preset[key];
    return v ? ' style="' + v + '"' : '';
  }

  /** 段落样式：继承正文字体，保证 Word 中字体不丢失 */
  function pStyleAttr(preset) {
    if (!preset) return '';
    var parts = [];
    if (preset.font) parts.push(preset.font);
    if (preset.p) parts.push(preset.p);
    return parts.length ? ' style="' + parts.join(';') + '"' : '';
  }

  /* ------------------------------------------------------------------ *
   * 行内解析
   * ------------------------------------------------------------------ */

  var PLACEHOLDER = '\u0000';

  function makeStash() {
    var items = [];
    return {
      hold: function (html) {
        items.push(html);
        return PLACEHOLDER + (items.length - 1) + PLACEHOLDER;
      },
      /** 延迟求值：占位符在最终 restore 时才展开（用于需要二次解析的内容，如链接文字） */
      holdLazy: function (fn) {
        items.push(fn);
        return PLACEHOLDER + (items.length - 1) + PLACEHOLDER;
      },
      restore: function (s) {
        return s.replace(new RegExp(PLACEHOLDER + '(\\d+)' + PLACEHOLDER, 'g'), function (m, i) {
          var item = items[Number(i)];
          return typeof item === 'function' ? item() : item;
        });
      }
    };
  }

  /**
   * 行内 Markdown → HTML
   * @param {string} text
   * @param {object} opts {preset, images:boolean, links:boolean, depth:number}
   */
  function inline(text, opts) {
    opts = opts || {};
    var preset = opts.preset || null;
    var depth = opts.depth || 0;
    var stash = makeStash();
    var s = String(text == null ? '' : text);

    // 1) 反斜杠转义 —— 必须在其余解析之前处理
    s = s.replace(/\\([\\`*_{}\[\]()#+\-.!>~|"'])/g, function (m, ch) {
      return stash.hold(escapeHtml(ch));
    });

    // 2) 图片 ![alt](src "title")：在 HTML 转义前抽取，以便保留原始 alt
    s = s.replace(
      /!\[([^\]]*)\]\(\s*((?:[^()\s]|\([^()\s]*\))+)(?:\s+"([^"]*)")?\s*\)/g,
      function (m, alt, src, title) {
        if (opts.images === false) return stash.hold(escapeHtml(alt || src));
        var safeSrc = sanitizeUrl(src);
        if (!safeSrc) return stash.hold(escapeHtml(alt || src));
        return stash.hold(
          '<img src="' + escapeAttr(safeSrc) + '" alt="' + escapeAttr(alt) + '"' +
          (title ? ' title="' + escapeAttr(title) + '"' : '') + ' />'
        );
      }
    );

    // 3) 链接 [text](href "title")：文字部分延迟二次解析，支持 **[粗](url)** 与 URL 内含括号
    if (opts.links !== false && depth < 4) {
      s = s.replace(
        /\[([^\]]*)\]\(\s*((?:[^()\s]|\([^()\s]*\))+)(?:\s+"([^"]*)")?\s*\)/g,
        function (m, label, href, title) {
          var safeHref = sanitizeUrl(href);
          return stash.holdLazy(function () {
            var inner = inline(label, {
              preset: preset,
              images: false,
              links: false,
              depth: depth + 1
            });
            // 危险协议：只保留文字，不生成链接
            if (!safeHref) return inner;
            return '<a href="' + escapeAttr(safeHref) + '"' +
              (title ? ' title="' + escapeAttr(title) + '"' : '') + '>' +
              inner + '</a>';
          });
        }
      );
    } else if (opts.links === false) {
      // 不解析链接时，仅保留可读文字
      s = s.replace(
        /\[([^\]]*)\]\(\s*((?:[^()\s]|\([^()\s]*\))+)(?:\s+"([^"]*)")?\s*\)/g,
        function (m, label) { return stash.holdLazy(function () { return inline(label, { preset: preset, images: false, links: false, depth: depth + 1 }); }); }
      );
    }

    // 4) 自动链接 <https://...> / <mail@host>
    s = s.replace(/<(https?|ftp|mailto):([^<>\s]+)>/g, function (m, scheme, rest) {
      var url = scheme + ':' + rest;
      return stash.hold('<a href="' + escapeAttr(url) + '">' + escapeHtml(url) + '</a>');
    });    s = s.replace(/<([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})>/g, function (m, mail) {
      return stash.hold('<a href="mailto:' + escapeAttr(mail) + '">' + escapeHtml(mail) + '</a>');
    });

    // 5) HTML 转义（此后 & < > " 均为实体，不会干扰后续标记匹配）
    s = escapeHtml(s);

    // 6) 行内代码（先固化，避免其内部被当作标记解析）
    s = s.replace(/(`+)([\s\S]*?[^`])\1(?!`)/g, function (m, ticks, code) {
      var body = code.replace(/^ /, '').replace(/ $/, '');
      return stash.hold('<code' + styleAttr(preset, 'code') + '>' + body + '</code>');
    });
    s = s.replace(/(`+)(\s*)\1/g, function (m, ticks, sp) {
      return stash.hold('<code' + styleAttr(preset, 'code') + '>' + sp + '</code>');
    });

    // 7) 强调：粗斜体 → 粗体 → 斜体 → 删除线 → 高亮
    s = s.replace(/\*\*\*(?=\S)([\s\S]*?\S)\*\*\*/g, '<strong><em>$1</em></strong>');
    s = s.replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>');
    s = s.replace(/__(?=\S)([\s\S]*?\S)__/g, '<strong>$1</strong>');
    s = s.replace(/(^|[^\w*])\*(?=\S)([^*]*?\S)\*(?!\w)/g, '$1<em>$2</em>');
    s = s.replace(/(^|[^\w_])_(?=\S)([^_]*?\S)_(?!\w)/g, '$1<em>$2</em>');
    s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<del>$1</del>');
    s = s.replace(/==(?=\S)([\s\S]*?\S)==/g, '<mark>$1</mark>');

    return stash.restore(s);
  }

  /**
   * 把段落/列表项内的多行渲染为 HTML，处理软换行与硬换行
   * - 行尾 2 个以上空格 或 反斜杠 → <br />
   * - 中英混排软换行：两侧非 CJK 时补一个空格
   */
  function renderInlineLines(lines, opts) {
    var out = '';
    for (var i = 0; i < lines.length; i++) {
      var raw = lines[i];
      var hard = /(?: {2,}|\\)\s*$/.test(raw);
      var content = hard ? raw.replace(/(?: {2,}|\\)\s*$/, '') : raw;
      var piece = inline(content, opts);
      if (i === 0) { out = piece; continue; }

      var prevRaw = lines[i - 1].replace(/(?: {2,}|\\)\s*$/, '');
      var prevHard = /(?: {2,}|\\)\s*$/.test(lines[i - 1]);
      if (prevHard) {
        out += '<br />' + piece;
        continue;
      }
      var prevCh = prevRaw.charAt(prevRaw.length - 1);
      var curCh = raw.charAt(0);
      if (!prevCh || !curCh) { out += piece; continue; }
      out += (hasCJK(prevCh) || hasCJK(curCh)) ? piece : ' ' + piece;
    }
    return out;
  }

  /* ------------------------------------------------------------------ *
   * 块级解析
   * ------------------------------------------------------------------ */

  var RE_ATX = /^ {0,3}(#{1,6})(\s+(.*?))?\s*#*\s*$/;
  var RE_HR = /^ {0,3}((\*\s*){3,}|(-\s*){3,}|(_\s*){3,})$/;
  var RE_FENCE = /^ {0,3}(`{3,}|~{3,})\s*([^`]*)$/;
  var RE_QUOTE = /^ {0,3}>\s?(.*)$/;
  var RE_LIST = /^(\s*)([-*+]|\d{1,9}[.)])(\s+|$)(.*)$/;
  var RE_SETEXT = /^ {0,3}(=+|-+)\s*$/;

  function matchListLine(line) {
    var m = RE_LIST.exec(line);
    if (!m) return null;
    var indent = expandTabs(m[1]).length;
    var marker = m[2];
    var ordered = /^\d/.test(marker);
    var rest = m[4];
    var checked = null;
    var tm = /^\[([ xX])\]\s+(.*)$/.exec(rest);
    if (tm && !ordered) {
      checked = tm[1].toLowerCase() === 'x';
      rest = tm[2];
    }
    // 无序列表的标记后必须跟空格（避免把 *强调* 误判）
    if (!ordered && !/\s/.test(m[3]) && m[3] !== '') return null;
    return {
      indent: indent,
      ordered: ordered,
      marker: marker,
      content: rest,
      checked: checked
    };
  }

  function splitTableRow(line) {
    var s = line.trim();
    if (s.charAt(0) === '|') s = s.slice(1);
    if (s.charAt(s.length - 1) === '|') s = s.slice(0, -1);
    var cells = [];
    var buf = '';
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === '\\' && s.charAt(i + 1) === '|') {
        buf += '|';
        i++;
      } else if (ch === '|') {
        cells.push(buf);
        buf = '';
      } else {
        buf += ch;
      }
    }
    cells.push(buf);
    return cells.map(function (c) { return c.trim(); });
  }

  function isTableSeparator(line) {
    if (line.indexOf('-') < 0 || line.indexOf('|') < 0) return false;
    var cells = splitTableRow(line);
    if (!cells.length) return false;
    return cells.every(function (c) { return /^:?-{1,}:?$/.test(c); });
  }

  function parseAligns(line) {
    return splitTableRow(line).map(function (c) {
      var left = c.charAt(0) === ':';
      var right = c.charAt(c.length - 1) === ':';
      if (left && right) return 'center';
      if (right) return 'right';
      if (left) return 'left';
      return '';
    });
  }

  /* ---- 列表：先把条目收集成线性数组，再按缩进建树 ---- */

  function collectListEntries(lines, start, baseIndent) {
    var entries = [];
    var i = start;
    var pending = null; // 当前条目（用于续行）

    while (i < lines.length) {
      var line = lines[i];

      if (isBlank(line)) {
        // 空行后若仍是同层或更深层的列表行，则视为条目内空行
        var j = i + 1;
        while (j < lines.length && isBlank(lines[j])) j++;
        if (j < lines.length) {
          var nm = matchListLine(expandTabs(lines[j]));
          if (nm && nm.indent >= baseIndent) {
            if (pending) pending.contentLines.push('');
            i = j;
            continue;
          }
        }
        break;
      }

      var raw = expandTabs(line);
      var m = matchListLine(raw);

      if (m) {
        if (m.indent < baseIndent) break;
        entries.push({
          indent: m.indent,
          ordered: m.ordered,
          marker: m.marker,
          checked: m.checked,
          contentLines: [m.content]
        });
        pending = entries[entries.length - 1];
        i++;
        continue;
      }

      // 非列表行
      var indentLen = raw.length - raw.replace(/^\s+/, '').length;
      var isContinuation = pending &&
        (indentLen > pending.indent || !RE_ATX.test(raw) && !RE_FENCE.test(raw) && !RE_HR.test(raw) && !isBlank(raw));

      if (isContinuation) {
        // 懒续行：直接并入上一个条目
        pending.contentLines.push(raw.replace(/^\s{0,4}/, ''));
        i++;
        continue;
      }
      break;
    }

    return { entries: entries, next: i };
  }

  function buildListTree(entries) {
    var rootItems = [];
    var stack = [];
    entries.forEach(function (e) {
      while (stack.length && e.indent < stack[stack.length - 1].indent) stack.pop();

      var parentItem = null;
      if (stack.length && e.indent === stack[stack.length - 1].indent) {
        parentItem = stack.length > 1 ? stack[stack.length - 2].item : null;
        stack.pop();
      } else if (stack.length) {
        parentItem = stack[stack.length - 1].item;
      }

      var item = {
        ordered: e.ordered,
        checked: e.checked,
        contentLines: e.contentLines,
        children: []
      };
      if (parentItem) parentItem.children.push(item);
      else rootItems.push(item);

      stack.push({ indent: e.indent, item: item });
    });
    return rootItems;
  }

  function renderListGroup(items, opts, depth) {
    var html = '';
    var i = 0;
    while (i < items.length) {
      var type = items[i].ordered ? 'ol' : 'ul';
      var group = [];
      while (i < items.length && (items[i].ordered ? 'ol' : 'ul') === type) {
        group.push(items[i]);
        i++;
      }
      html += '<' + type + '>';
      for (var k = 0; k < group.length; k++) {
        html += renderListItem(group[k], opts, depth);
      }
      html += '</' + type + '>';
    }
    return html;
  }

  function renderListItem(item, opts, depth) {
    var preset = opts.preset || null;
    var lines = item.contentLines.slice();
    // 去掉尾部空行
    while (lines.length && isBlank(lines[lines.length - 1])) lines.pop();

    var prefix = '';
    if (item.checked === true) prefix = CHECKED + ' ';
    else if (item.checked === false) prefix = UNCHECKED + ' ';

    var body = '';
    if (!lines.length) {
      body = '';
    } else if (lines.length === 1) {
      body = prefix + inline(lines[0], opts);
    } else {
      // 多段：首行与后续段落
      var paras = [];
      var buf = [lines[0]];
      for (var n = 1; n < lines.length; n++) {
        if (isBlank(lines[n])) {
          if (buf.length) paras.push(buf);
          buf = [];
        } else {
          buf.push(lines[n]);
        }
      }
      if (buf.length) paras.push(buf);
      paras.forEach(function (p, idx) {
        var pre = idx === 0 ? prefix : '';
        paras[idx] = '<p' + pStyleAttr(preset) + '>' + pre + renderInlineLines(p, opts) + '</p>';
      });
      body = paras.join('');
    }

    var kids = item.children.length
      ? renderListGroup(item.children, opts, (depth || 0) + 1)
      : '';

    return '<li>' + body + kids + '</li>';
  }

  /** 把段落内的多行合并为单行文本（用于不支持换行的场景） */
  function joinParagraphLines(lines) {
    return lines.map(function (l) { return l.replace(/(?: {2,}|\\)\s*$/, ''); }).join('\n');
  }

  /* ---- 主解析循环 ---- */

  function mdToHtml(markdown, options) {
    options = options || {};
    var preset = PRESETS.hasOwnProperty(options.stylePreset)
      ? PRESETS[options.stylePreset]
      : PRESETS['default'];
    var opts = {
      preset: preset,
      images: options.images !== false,
      links: options.links !== false
    };

    var text = String(markdown == null ? '' : markdown).replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ');
    var lines = text.split('\n');
    var out = [];
    var i = 0;

    function para(inner, extra) {
      var cls = extra ? ' class="' + extra + '"' : '';
      return '<p' + cls + pStyleAttr(preset) + '>' + inner + '</p>';
    }

    /** 判断第 i 行是否为 setext 标题下划线（需上一行非空且不是块级起始） */
    function isSetextUnderline(idx) {
      if (idx <= 0 || idx >= lines.length) return false;
      if (!RE_SETEXT.test(lines[idx])) return false;
      var prev = lines[idx - 1];
      if (isBlank(prev)) return false;
      if (RE_ATX.test(prev) || RE_FENCE.test(prev) || RE_QUOTE.test(prev) ||
          matchListLine(expandTabs(prev)) || /^ {4,}/.test(prev)) return false;
      // --- 单独出现时是水平线，不是 setext
      if (/^-+\s*$/.test(lines[idx]) && RE_HR.test(prev)) return false;
      return true;
    }

    while (i < lines.length) {
      var line = lines[i];

      if (isBlank(line)) { i++; continue; }

      // ---- 围栏代码块 ----
      var fm = RE_FENCE.exec(line);
      if (fm) {
        var fence = fm[1];
        var codeLines = [];
        i++;
        while (i < lines.length) {
          var closeRe = new RegExp('^ {0,3}' + fence.charAt(0) + '{' + fence.length + ',}\\s*$');
          if (closeRe.test(lines[i])) { i++; break; }
          codeLines.push(lines[i]);
          i++;
        }
        var codeHtml = escapeHtml(codeLines.join('\n'));
        if (preset) {
          out.push('<pre' + styleAttr(preset, 'pre') + '>' + codeHtml + '</pre>');
        } else {
          out.push('<pre>' + codeHtml + '</pre>');
        }
        continue;
      }

      // ---- 缩进代码块（4 空格） ----
      if (/^ {4,}\S/.test(line) && !matchListLine(line)) {
        var indented = [];
        while (i < lines.length && (/^ {4,}/.test(lines[i]) || isBlank(lines[i]))) {
          if (isBlank(lines[i]) && !/^ {4,}/.test(lines[i + 1] || '')) break;
          indented.push(lines[i].replace(/^ {4}/, ''));
          i++;
        }
        out.push('<pre' + styleAttr(preset, 'pre') + '>' + escapeHtml(indented.join('\n')) + '</pre>');
        continue;
      }

      // ---- ATX 标题 ----
      var hm = RE_ATX.exec(line);
      if (hm) {
        var level = hm[1].length;
        var htext = (hm[3] || '').trim();
        out.push('<h' + level + styleAttr(preset, 'h' + level) + '>' + inline(htext, opts) + '</h' + level + '>');
        i++;
        continue;
      }

      // ---- 水平线 ----
      if (RE_HR.test(line)) {
        if (preset) out.push('<p style="' + preset.hr + '"></p>');
        else out.push('<hr />');
        i++;
        continue;
      }

      // ---- 引用块 ----
      if (RE_QUOTE.test(line)) {
        var quoteLines = [];
        while (i < lines.length && RE_QUOTE.test(lines[i])) {
          quoteLines.push(RE_QUOTE.exec(lines[i])[1]);
          i++;
        }
        // 懒续行
        while (i < lines.length && !isBlank(lines[i]) && !RE_QUOTE.test(lines[i]) &&
               !RE_ATX.test(lines[i]) && !matchListLine(lines[i]) && !RE_FENCE.test(lines[i])) {
          quoteLines.push(lines[i]);
          i++;
        }
        var innerHtml = mdToHtml(quoteLines.join('\n'), {
          stylePreset: 'semantic',
          images: options.images,
          links: options.links
        });
        if (preset) {
          out.push('<div style="' + preset.quote + '">' + innerHtml + '</div>');
        } else {
          out.push('<blockquote>' + innerHtml + '</blockquote>');
        }
        continue;
      }

      // ---- 表格 ----
      if (line.indexOf('|') >= 0 && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
        var header = splitTableRow(line);
        var aligns = parseAligns(lines[i + 1]);
        i += 2;
        var rows = [];
        while (i < lines.length && !isBlank(lines[i]) && lines[i].indexOf('|') >= 0) {
          rows.push(splitTableRow(lines[i]));
          i++;
        }
        out.push(renderTable(header, aligns, rows, opts, preset));
        continue;
      }

      // ---- 列表 ----
      var lm = matchListLine(expandTabs(line));
      if (lm) {
        var collected = collectListEntries(lines, i, lm.indent);
        i = collected.next;
        var tree = buildListTree(collected.entries);
        out.push(renderListGroup(tree, opts, 0));
        continue;
      }

      // ---- HTML 块（经安全清洗后透传；自动链接 <https://…> 不算 HTML 块） ----
      if (/^ {0,3}<[a-zA-Z!/]/.test(line) &&
          !/^ {0,3}<(?:https?|ftp|mailto):[^<>\s]+>\s*$/.test(line) &&
          !/^ {0,3}<[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}>\s*$/.test(line)) {
        var htmlLines = [];
        while (i < lines.length && !isBlank(lines[i])) {
          htmlLines.push(lines[i]);
          i++;
        }
        var cleaned = sanitizeHtml(htmlLines.join('\n')).trim();
        if (cleaned) out.push(cleaned);
        continue;
      }

      // ---- 段落 ----
      var paraLines = [];
      while (i < lines.length && !isBlank(lines[i])) {
        var cur = lines[i];
        if (RE_ATX.test(cur) || RE_FENCE.test(cur) || RE_HR.test(cur) ||
            RE_QUOTE.test(cur) || matchListLine(expandTabs(cur)) ||
            (cur.indexOf('|') >= 0 && isTableSeparator(lines[i + 1] || ''))) {
          break;
        }
        // setext 下划线：作为段落终结符，交给下方标题逻辑处理
        if (paraLines.length === 1 && RE_SETEXT.test(cur)) break;
        paraLines.push(cur);
        i++;
      }
      if (paraLines.length) {
        // Setext 标题：段落仅一行且下一行是 === / ---
        if (paraLines.length === 1 && isSetextUnderline(i)) {
          var lvl = /^=+/.test(lines[i]) ? 1 : 2;
          out.push('<h' + lvl + styleAttr(preset, 'h' + lvl) + '>' +
            inline(paraLines[0], opts) + '</h' + lvl + '>');
          i++;
        } else {
          out.push(para(renderInlineLines(paraLines, opts)));
        }
      } else {
        i++; // 防御：避免死循环
      }
    }

    return out.join('\n');
  }

  function renderTable(header, aligns, rows, opts, preset) {
    var html = '<table' + styleAttr(preset, 'table') + '>';
    if (header.length) {
      html += '<thead><tr>';
      for (var c = 0; c < header.length; c++) {
        var a = aligns[c] || '';
        var alignStyle = a ? ';text-align:' + a : '';
        var thStyle = preset ? ' style="' + preset.th + alignStyle + '"' : (a ? ' style="text-align:' + a + '"' : '');
        html += '<th' + thStyle + '>' + inline(header[c], opts) + '</th>';
      }
      html += '</tr></thead>';
    }
    html += '<tbody>';
    for (var r = 0; r < rows.length; r++) {
      html += '<tr>';
      for (var k = 0; k < header.length; k++) {
        var av = aligns[k] || '';
        var as = av ? ';text-align:' + av : '';
        var tdStyle = preset ? ' style="' + preset.td + as + '"' : (av ? ' style="text-align:' + av + '"' : '');
        html += '<td' + tdStyle + '>' + inline(rows[r][k] == null ? '' : rows[r][k], opts) + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
  }

  return {
    mdToHtml: mdToHtml,
    inline: inline,
    presets: PRESETS,
    _internal: {
      matchListLine: matchListLine,
      splitTableRow: splitTableRow,
      isTableSeparator: isTableSeparator,
      buildListTree: buildListTree,
      joinParagraphLines: joinParagraphLines,
      escapeHtml: escapeHtml
    }
  };
});

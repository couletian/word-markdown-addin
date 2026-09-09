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

  /**
   * 命名实体解码表。
   * 只列出「能拼出危险协议 / 结构」的实体，足以覆盖绕过向量。
   */
  var NAMED_ENTITIES = {
    amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0',
    tab: '\t', newline: '\n', colon: ':', sol: '/', num: '#', period: '.',
    commat: '@', lpar: '(', rpar: ')', semi: ';', equals: '=', quest: '?'
  };

  /** 解码 HTML 实体（十进制 / 十六进制 / 命名；容忍缺失分号，与浏览器一致） */
  function decodeEntities(input) {
    var s = String(input == null ? '' : input);
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#[xX][0-9a-fA-F]+;?|#[0-9]+;?|[a-zA-Z][a-zA-Z0-9]*;?)/g, function (m, body) {
      if (body.charAt(0) === '#') {
        var isHex = body.charAt(1) === 'x' || body.charAt(1) === 'X';
        var code = isHex ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
        if (isNaN(code) || code < 0 || code > 0x10ffff) return m;
        if (code >= 0xd800 && code <= 0xdfff) return m; // 代理区码点非法
        try { return String.fromCodePoint(code); } catch (e) { return m; }
      }
      var key = body.replace(/;$/, '').toLowerCase();
      return Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, key) ? NAMED_ENTITIES[key] : m;
    });
  }

  /**
   * 控制字符 / 不可见字符。
   * 攻击者常用 \t、\n、\0 等把 "javascript:" 拆开以绕过协议匹配。
   */
  var CTRL_RE = /[\u0000-\u0020\u007f\u0080-\u009f\u00a0\u1680\u180e\u2000-\u200f\u2028\u2029\u202f\u205f\u2060\u3000\ufeff]/g;

  var SAFE_SCHEME_RE = /^(?:https?|ftps?|mailto|tel|sms|#|\/\/|\/{1,2}|\.{0,2}\/)/i;
  var DATA_IMG_RE = /^data:image\/(?:png|jpe?g|gif|webp|bmp);base64,[a-z0-9+/=\s]*$/i;
  var ANY_SCHEME_RE = /^[a-z][a-z0-9+.\-]*:/i;

  /**
   * URL 白名单过滤。
   * 顺序很关键：先解码实体、再剥离控制字符，然后才判断协议 —— 否则
   * `&#x09;javascript:` / `java&#x0A;script:` / `javascript&#x3A;` 都会绕过。
   */
  function sanitizeUrl(url) {
    var raw = decodeEntities(url).replace(CTRL_RE, '').trim();
    if (!raw) return '';
    if (DATA_IMG_RE.test(raw)) return raw;   // 仅位图，不含 svg+xml（可携带脚本）
    if (SAFE_SCHEME_RE.test(raw)) return raw;
    // 显式带协议但不属于白名单（javascript:/vbscript:/file:/blob:/about:…）→ 拒绝
    if (ANY_SCHEME_RE.test(raw)) return '';
    return raw; // 相对路径 / 锚点
  }

  /** CSS 白名单过滤：阻断表达式、外链与脚本式 URL */
  var DANGEROUS_CSS_RE = /(?:expression\s*\(|behavior\s*:|-moz-binding|@import|javascript\s*:|vbscript\s*:|url\s*\()/i;

  function sanitizeStyle(value) {
    var s = decodeEntities(value).replace(/[\u0000-\u001f\u007f]/g, '');
    if (DANGEROUS_CSS_RE.test(s)) return '';
    return s;
  }

  /* ---- 白名单定义 ---- */

  /** 允许保留的标签（Word insertHtml 能正确识别的子集） */
  var SAFE_TAGS = {
    p: 1, br: 1, hr: 1, div: 1, span: 1, center: 1,
    b: 1, strong: 1, i: 1, em: 1, u: 1, s: 1, strike: 1, del: 1, ins: 1,
    mark: 1, sub: 1, sup: 1, small: 1, big: 1, font: 1,
    code: 1, pre: 1, blockquote: 1,
    ul: 1, ol: 1, li: 1, dl: 1, dt: 1, dd: 1,
    table: 1, thead: 1, tbody: 1, tfoot: 1, tr: 1, th: 1, td: 1,
    caption: 1, colgroup: 1, col: 1,
    a: 1, img: 1,
    h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, h6: 1
  };

  /** 自闭合标签（无内容） */
  var VOID_TAGS = {
    br: 1, hr: 1, img: 1, col: 1, area: 1, base: 1, input: 1,
    link: 1, meta: 1, param: 1, source: 1, track: 1, wbr: 1, embed: 1
  };

  /**
   * 需要「连内容一起丢弃」的标签。
   * 这些标签内部可藏脚本或可执行属性（如 svg 里的 animate/onload），
   * 保留文字没有意义，必须整段删除。
   */
  var DROP_WHOLE_TAGS = {
    script: 1, style: 1, iframe: 1, object: 1, embed: 1, applet: 1,
    frame: 1, frameset: 1, link: 1, meta: 1, base: 1,
    svg: 1, math: 1, template: 1, noscript: 1, noembed: 1, xmp: 1,
    title: 1, param: 1, source: 1, track: 1, canvas: 1,
    audio: 1, video: 1, form: 1, button: 1, select: 1, textarea: 1,
    input: 1, marquee: 1, details: 1, dialog: 1, slot: 1, portal: 1
  };

  /**
   * 允许保留的属性。
   * 不含任何 on* 事件属性，不含 id/class（防 DOM clobbering 与样式劫持）。
   */
  var SAFE_ATTRS = {
    href: 1, src: 1, alt: 1, title: 1, style: 1,
    align: 1, valign: 1, width: 1, height: 1, border: 1,
    colspan: 1, rowspan: 1, cellpadding: 1, cellspacing: 1,
    start: 1, type: 1, color: 1, face: 1, size: 1
  };

  /** 需要做 URL 校验的属性 */
  var URL_ATTRS = { href: 1, src: 1 };

  /** 找到标签结束的 '>'（跳过引号内的字符） */
  function findTagEnd(s, start) {
    var quote = '';
    for (var k = start + 1; k < s.length; k++) {
      var ch = s.charAt(k);
      if (quote) { if (ch === quote) quote = ''; }
      else if (ch === '"' || ch === "'") quote = ch;
      else if (ch === '>') return k;
    }
    return -1;
  }

  /**
   * 清洗透传的 HTML 片段。
   *
   * 实现方式：词法扫描 + 白名单重建。
   * 只输出 SAFE_TAGS 中的标签、SAFE_ATTRS 中的属性，其余一律丢弃；
   * 文本节点统一「先解码实体、再转义」，避免实体二次编码造成的绕过。
   * 不再使用正则黑名单 —— 那是 V-01 漏洞的根源。
   */
  function sanitizeHtml(html) {
    var src = String(html == null ? '' : html);
    var srcLower = src.toLowerCase();
    var len = src.length;
    var out = [];
    var i = 0;
    var dropStack = []; // 正在丢弃内容的标签名栈

    function pushText(chunk) {
      if (!chunk) return;
      out.push(escapeHtml(decodeEntities(chunk)));
    }

    /** 从 from 起找 name 的配对结束标签，返回其后的下标；找不到返回 -1 */
    function findCloseTag(name, from) {
      var needle = '</' + name;
      var p = srcLower.indexOf(needle, from);
      while (p >= 0) {
        var after = srcLower.charAt(p + needle.length);
        if (after === '' || after === '>' || /\s|\//.test(after)) {
          var gt = src.indexOf('>', p);
          return gt < 0 ? len : gt + 1;
        }
        p = srcLower.indexOf(needle, p + 1);
      }
      return -1;
    }

    while (i < len) {
      var lt = src.indexOf('<', i);
      if (lt < 0) {
        if (!dropStack.length) pushText(src.slice(i));
        break;
      }
      if (lt > i && !dropStack.length) pushText(src.slice(i, lt));

      // 注释 / 文档类型 / 处理指令 / CDATA
      if (src.startsWith('<!--', lt)) {
        var commentEnd = src.indexOf('-->', lt);
        i = commentEnd < 0 ? len : commentEnd + 3;
        continue;
      }
      if (src.startsWith('<!', lt) || src.startsWith('<?', lt)) {
        var declEnd = src.indexOf('>', lt);
        i = declEnd < 0 ? len : declEnd + 1;
        continue;
      }

      var gt = findTagEnd(src, lt);
      if (gt < 0) { // 未闭合的 '<'，按文本处理
        if (!dropStack.length) pushText(src.slice(lt));
        break;
      }

      var inner = src.slice(lt + 1, gt);
      i = gt + 1;

      var closing = false;
      if (inner.charAt(0) === '/') { closing = true; inner = inner.slice(1); }
      var selfClose = /\/\s*$/.test(inner);
      if (selfClose) inner = inner.replace(/\/\s*$/, '');

      var nameMatch = /^([a-zA-Z][^\s/>]*)/.exec(inner);
      var rawName = nameMatch ? nameMatch[1] : '';
      var name = rawName.toLowerCase();
      if (!name) continue;

      // ---------- 结束标签 ----------
      if (closing) {
        if (dropStack.length) {
          var idx = dropStack.lastIndexOf(name);
          if (idx >= 0) dropStack.length = idx;
          continue;
        }
        if (Object.prototype.hasOwnProperty.call(SAFE_TAGS, name) && !VOID_TAGS[name]) {
          out.push('</' + name + '>');
        }
        continue;
      }

      // ---------- 已在丢弃模式：只跟踪嵌套的整段丢弃标签 ----------
      if (dropStack.length) {
        if (DROP_WHOLE_TAGS[name] && !VOID_TAGS[name] && !selfClose) dropStack.push(name);
        continue;
      }

      // ---------- 需要整段丢弃的标签 ----------
      if (DROP_WHOLE_TAGS[name]) {
        if (!selfClose && !VOID_TAGS[name]) {
          var closeAt = findCloseTag(name, i);
          if (closeAt < 0) dropStack.push(name); // 没有结束标签 → 丢弃到文末
          else i = closeAt;
        }
        continue;
      }

      // ---------- 白名单外的标签：丢标签、保留文字 ----------
      if (!Object.prototype.hasOwnProperty.call(SAFE_TAGS, name)) continue;

      // ---------- 白名单标签：重建属性 ----------
      var attrs = '';
      var rest = inner.slice(rawName.length);
      var attrRe = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
      var am;
      while ((am = attrRe.exec(rest)) !== null) {
        var attrName = am[1].toLowerCase();
        if (attrName.indexOf(':') >= 0) continue; // 命名空间属性（xlink:href 等）
        if (!Object.prototype.hasOwnProperty.call(SAFE_ATTRS, attrName)) continue; // 含全部 on*

        var rawVal = am[2] != null ? am[2] : (am[3] != null ? am[3] : (am[4] != null ? am[4] : ''));
        var val = decodeEntities(rawVal);

        if (URL_ATTRS[attrName]) {
          var safe = sanitizeUrl(val);
          if (!safe) continue;
          attrs += ' ' + attrName + '="' + escapeAttr(safe) + '"';
        } else if (attrName === 'style') {
          var style = sanitizeStyle(val);
          if (!style) continue;
          attrs += ' style="' + escapeAttr(style) + '"';
        } else {
          attrs += ' ' + attrName + '="' + escapeAttr(val.replace(/[\u0000-\u001f\u007f]/g, '')) + '"';
        }
      }

      out.push('<' + name + attrs + (VOID_TAGS[name] ? ' />' : '>'));
    }

    return out.join('');
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

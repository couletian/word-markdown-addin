/*!
 * office-api.js — Office.js 交互层
 * ------------------------------------------------------------------
 * 封装三大核心功能，供任务面板（taskpane.js）与功能区命令（commands.js）共用：
 *
 *   1. exportMarkdown()      Word 文档 → .md 文件
 *   2. insertMarkdown()      剪贴板/文本 Markdown → Word 富文本
 *   3. importMarkdownFile()  本地 .md 文件 → 排版完整的 Word 文档
 *
 * 依赖：WordMd.markdown（markdown.js）、WordMd.docx2md（docx2md.js）、WordMd.zip（zip.js）
 */
(function (root) {
  'use strict';

  var WordMd = root.WordMd = root.WordMd || {};

  /* ------------------------------------------------------------------ *
   * 工具
   * ------------------------------------------------------------------ */

  function base64ToUint8Array(base64) {
    var binary = atob(base64);
    var len = binary.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  function uint8ArrayToBase64(bytes) {
    var chunk = 0x8000;
    var parts = [];
    for (var i = 0; i < bytes.length; i += chunk) {
      parts.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
    }
    return btoa(parts.join(''));
  }

  function pad(n) { return n < 10 ? '0' + n : String(n); }

  function timestamp() {
    var d = new Date();
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes());
  }

  /** 生成可下载的 Blob URL（返回 URL 供 UI 渲染下载链接） */
  function createDownloadUrl(text, mime) {
    var blob = new Blob([text], { type: (mime || 'text/markdown') + ';charset=utf-8' });
    return URL.createObjectURL(blob);
  }

  /** 尝试直接触发下载；返回是否成功发起 */
  function triggerDownload(url, filename) {
    try {
      var a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { document.body.removeChild(a); }, 1000);
      return true;
    } catch (e) {
      return false;
    }
  }

  /** 清理历史 blob URL，避免内存泄漏 */
  var liveUrls = [];
  function revokeAll() {
    liveUrls.forEach(function (u) { try { URL.revokeObjectURL(u); } catch (e) { /* ignore */ } });
    liveUrls = [];
  }

  /* ------------------------------------------------------------------ *
   * 1) 导出：Word → Markdown
   * ------------------------------------------------------------------ */

  /**
   * 读取当前文档（或选区）的 OOXML 并转换为 Markdown
   * @param {object} opts { scope: 'document'|'selection' }
   * @returns {Promise<{markdown:string, stats:object}>}
   */
  function readAsMarkdown(opts) {
    opts = opts || {};
    return Word.run(function (context) {
      var target = opts.scope === 'selection'
        ? context.document.getSelection()
        : context.document.body;
      var ooxmlResult = target.getOoxml();
      return context.sync().then(function () {
        // getOoxml() 返回的是 Flat OPC 格式的 XML 字符串（不是 base64），
        // 直接交给 docx2md 解析，切勿再做 base64 解码。
        var ooxml = ooxmlResult.value;
        var markdown = WordMd.docx2md.docxToMarkdown(ooxml, {
          imagePlaceholder: '![图片](图片占位：导出时请手动补全)'
        });
        return { markdown: markdown, stats: computeStats(markdown) };
      });
    });
  }

  function computeStats(md) {
    var lines = md.split('\n');
    return {
      chars: md.length,
      lines: lines.length,
      words: (md.match(/[\u4e00-\u9fa5]|[A-Za-z0-9']+/g) || []).length,
      headings: (md.match(/^#{1,6}\s/gm) || []).length,
      tables: (md.match(/^\|/gm) || []).length,
      links: (md.match(/\[[^\]]*\]\([^)]*\)/g) || []).length,
      images: (md.match(/!\[[^\]]*\]\([^)]*\)/g) || []).length
    };
  }

  /**
   * 导出当前文档为 .md 文件
   * @param {object} opts { scope, filename, autoDownload }
   * @returns {Promise<{markdown, filename, url, stats, downloaded}>}
   */
  function exportMarkdown(opts) {
    opts = opts || {};
    return readAsMarkdown({ scope: opts.scope }).then(function (res) {
      revokeAll();
      var filename = opts.filename || ('document-' + timestamp() + '.md');
      var url = createDownloadUrl(res.markdown, 'text/markdown');
      liveUrls.push(url);
      var downloaded = opts.autoDownload === false ? false : triggerDownload(url, filename);
      return {
        markdown: res.markdown,
        stats: res.stats,
        filename: filename,
        url: url,
        downloaded: downloaded
      };
    });
  }

  /* ------------------------------------------------------------------ *
   * 2) 插入：Markdown → Word 富文本
   * ------------------------------------------------------------------ */

  /**
   * 把 Markdown 转换为 HTML 并插入文档
   * @param {string} markdown
   * @param {object} opts { preset:'default'|'thesis'|'semantic', mode:'cursor'|'end'|'replace'|'new' }
   */
  function insertMarkdown(markdown, opts) {
    opts = opts || {};
    var mode = opts.mode || 'cursor';
    var html = WordMd.markdown.mdToHtml(markdown, {
      stylePreset: opts.preset || 'default',
      images: opts.images !== false,
      links: opts.links !== false
    });
    if (!html.trim()) {
      return Promise.reject(new Error('内容为空，没有可插入的 Markdown'));
    }

    if (mode === 'new') {
      return createNewDocument(html);
    }

    return Word.run(function (context) {
      var range;
      var location;
      if (mode === 'replace') {
        range = context.document.body;
        location = Word.InsertLocation.replace;
      } else if (mode === 'end') {
        range = context.document.body;
        location = Word.InsertLocation.end;
      } else {
        range = context.document.getSelection();
        location = Word.InsertLocation.replace;
      }
      range.insertHtml(html, location);
      return context.sync().then(function () {
        return { inserted: true, mode: mode, htmlLength: html.length };
      });
    });
  }

  /** 新建文档并写入 HTML（WordApi 1.3+） */
  function createNewDocument(html) {
    return Word.run(function (context) {
      if (!context.application.createDocument) {
        // 回退：在当前文档末尾插入，并提示
        context.document.body.insertHtml(html, Word.InsertLocation.end);
        return context.sync().then(function () {
          return { inserted: true, mode: 'end', fallback: true };
        });
      }
      var created = context.application.createDocument();
      return context.sync().then(function () {
        created.body.insertHtml(html, Word.InsertLocation.replace);
        return created.sync();
      }).then(function () {
        created.open();
        return created.sync();
      }).then(function () {
        return { inserted: true, mode: 'new', fallback: false };
      });
    });
  }

  /**
   * 读取剪贴板文本
   * @returns {Promise<string>}
   */
  function readClipboard() {
    if (navigator.clipboard && navigator.clipboard.readText) {
      return navigator.clipboard.readText().catch(function () {
        return Promise.reject(new Error('无法读取剪贴板（可能被系统权限拦截），请在面板的文本框中手动粘贴'));
      });
    }
    return Promise.reject(new Error('当前环境不支持读取剪贴板，请在面板的文本框中手动粘贴'));
  }

  /** 写入剪贴板 */
  function writeClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return Promise.reject(new Error('当前环境不支持写入剪贴板'));
  }

  /* ------------------------------------------------------------------ *
   * 3) 导入：.md 文件 → Word 文档
   * ------------------------------------------------------------------ */

  /**
   * 读取本地 .md 文件内容
   * @param {File} file
   * @returns {Promise<string>}
   */
  function readMarkdownFile(file) {
    return new Promise(function (resolve, reject) {
      if (!file) return reject(new Error('未选择文件'));
      var reader = new FileReader();
      reader.onload = function () {
        var text = String(reader.result || '');
        // 去掉 BOM
        resolve(text.replace(/^\uFEFF/, ''));
      };
      reader.onerror = function () { reject(new Error('文件读取失败：' + (reader.error && reader.error.message))); };
      reader.readAsText(file, 'utf-8');
    });
  }

  /**
   * 导入 .md 文件为 Word 文档
   * @param {File} file
   * @param {object} opts { preset, mode:'new'|'end'|'replace' }
   */
  function importMarkdownFile(file, opts) {
    opts = opts || {};
    return readMarkdownFile(file).then(function (text) {
      return insertMarkdown(text, {
        preset: opts.preset,
        mode: opts.mode || 'new',
        images: opts.images,
        links: opts.links
      }).then(function (res) {
        res.filename = file.name;
        res.markdown = text;
        res.stats = computeStats(text);
        return res;
      });
    });
  }

  /* ------------------------------------------------------------------ *
   * 环境能力检测
   * ------------------------------------------------------------------ */

  function checkRequirements() {
    return new Promise(function (resolve) {
      if (typeof Office === 'undefined' || !Office.onReady) {
        return resolve({ ok: false, reason: '不在 Office 宿主环境中（请通过 Word 加载项打开）' });
      }
      Office.onReady(function (info) {
        if (!info || info.host !== Office.HostType.Word) {
          return resolve({ ok: false, reason: '当前宿主不是 Word' });
        }
        var support = Office.context.requirements.isSetSupported;
        resolve({
          ok: true,
          host: info.host,
          platform: info.platform,
          supports: {
            ooxml: support('WordApi', '1.1'),
            insertHtml: support('WordApi', '1.1'),
            createDocument: support('WordApi', '1.3'),
            search: support('WordApi', '1.4')
          }
        });
      });
    });
  }

  /* ------------------------------------------------------------------ *
   * 导出
   * ------------------------------------------------------------------ */

  WordMd.office = {
    readAsMarkdown: readAsMarkdown,
    exportMarkdown: exportMarkdown,
    insertMarkdown: insertMarkdown,
    importMarkdownFile: importMarkdownFile,
    readMarkdownFile: readMarkdownFile,
    readClipboard: readClipboard,
    writeClipboard: writeClipboard,
    createDownloadUrl: createDownloadUrl,
    triggerDownload: triggerDownload,
    computeStats: computeStats,
    checkRequirements: checkRequirements,
    base64ToUint8Array: base64ToUint8Array,
    uint8ArrayToBase64: uint8ArrayToBase64,
    _internal: { timestamp: timestamp, revokeAll: revokeAll }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

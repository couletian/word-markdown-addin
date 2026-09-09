/*!
 * taskpane.js — 任务面板逻辑
 * ------------------------------------------------------------------
 * 三个标签页对应三大功能：导出 / 粘贴 / 导入
 */
(function () {
  'use strict';

  var $ = function (sel) { return document.querySelector(sel); };
  var $$ = function (sel) { return Array.prototype.slice.call(document.querySelectorAll(sel)); };

  var state = {
    lastMarkdown: '',
    lastUrl: '',
    lastFilename: '',
    preset: 'default',
    scope: 'document',
    pasteMode: 'cursor',
    importMode: 'new',
    previewMode: 'rendered'
  };

  /* ------------------------------------------------------------------ *
   * 通用 UI
   * ------------------------------------------------------------------ */

  function toast(message, type) {
    var el = $('#toast');
    el.textContent = message;
    el.className = 'toast show ' + (type || 'info');
    clearTimeout(el._timer);
    el._timer = setTimeout(function () { el.className = 'toast'; }, 4200);
  }

  function busy(on, label) {
    var overlay = $('#busy');
    overlay.classList.toggle('show', !!on);
    if (label) $('#busy-label').textContent = label;
  }

  function fail(err) {
    console.error(err);
    toast((err && err.message) || String(err), 'error');
  }

  function switchTab(name) {
    $$('.tab').forEach(function (t) {
      t.classList.toggle('active', t.dataset.tab === name);
    });
    $$('.panel').forEach(function (p) {
      p.classList.toggle('active', p.id === 'panel-' + name);
    });
    if (name === 'paste') refreshClipboardPreview();
  }

  function fmtStats(s) {
    if (!s) return '';
    return '字符 ' + s.chars + ' · 行 ' + s.lines + ' · 词 ' + s.words +
      ' · 标题 ' + s.headings + ' · 表格行 ' + s.tables +
      ' · 链接 ' + s.links + ' · 图片 ' + s.images;
  }

  /* ------------------------------------------------------------------ *
   * 功能一：导出 Markdown
   * ------------------------------------------------------------------ */

  function doExport() {
    busy(true, '正在读取文档并转换…');
    WordMd.office.exportMarkdown({
      scope: state.scope,
      autoDownload: true
    }).then(function (res) {
      state.lastMarkdown = res.markdown;
      state.lastUrl = res.url;
      state.lastFilename = res.filename;

      $('#export-result').classList.add('show');
      $('#export-stats').textContent = fmtStats(res.stats);
      $('#export-preview').textContent = res.markdown.slice(0, 20000);

      var link = $('#export-download');
      link.href = res.url;
      link.download = res.filename;
      link.textContent = '⬇ 下载 ' + res.filename;

      if (res.downloaded) {
        toast('已导出 ' + res.filename + '（若浏览器拦截下载，请点击下方链接）', 'success');
      } else {
        toast('转换完成，请点击下方链接保存文件', 'info');
      }
    }).catch(fail).then(function () { busy(false); });
  }

  function doCopyMarkdown() {
    if (!state.lastMarkdown) return toast('请先执行一次导出', 'warn');
    WordMd.office.writeClipboard(state.lastMarkdown)
      .then(function () { toast('Markdown 已复制到剪贴板', 'success'); })
      .catch(fail);
  }

  /* ------------------------------------------------------------------ *
   * 功能二：粘贴 Markdown → Word 富文本
   * ------------------------------------------------------------------ */

  function refreshClipboardPreview() {
    var box = $('#clipboard-status');
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      box.textContent = '当前环境不支持自动读取剪贴板，请直接在下方文本框粘贴';
      box.className = 'hint warn';
      return;
    }
    navigator.clipboard.readText().then(function (text) {
      if (text && text.trim()) {
        box.textContent = '检测到剪贴板内容（' + text.length + ' 字符）：' + text.slice(0, 120).replace(/\s+/g, ' ') + '…';
        box.className = 'hint ok';
        if (!$('#paste-input').value.trim()) {
          $('#paste-input').value = text;
          renderPreview();
        }
      } else {
        box.textContent = '剪贴板为空，请先复制 Markdown 文本';
        box.className = 'hint';
      }
    }).catch(function () {
      box.textContent = '无法自动读取剪贴板（系统权限限制），请在下方文本框手动粘贴';
      box.className = 'hint warn';
    });
  }

  function renderPreview() {
    var md = $('#paste-input').value;
    var target = $('#paste-preview');
    if (!md.trim()) {
      target.innerHTML = '<p class="placeholder">在左侧输入或粘贴 Markdown，这里会实时显示 Word 中的效果</p>';
      return;
    }
    if (state.previewMode === 'rendered') {
      var html = WordMd.markdown.mdToHtml(md, { stylePreset: 'default' });
      target.innerHTML = '<div class="word-preview">' + html + '</div>';
    } else {
      target.textContent = md;
    }
  }

  function doPasteInsert() {
    var md = $('#paste-input').value;
    if (!md.trim()) return toast('请先输入或粘贴 Markdown 内容', 'warn');
    busy(true, '正在转换为 Word 富文本…');
    WordMd.office.insertMarkdown(md, {
      preset: state.preset,
      mode: state.pasteMode
    }).then(function (res) {
      if (res.fallback) {
        toast('当前 Word 版本不支持新建文档，已插入到当前文档末尾', 'warn');
      } else {
        toast('已转换为 Word 富文本并插入', 'success');
      }
    }).catch(fail).then(function () { busy(false); });
  }

  /* ------------------------------------------------------------------ *
   * 功能三：导入 .md 文件
   * ------------------------------------------------------------------ */

  function handleFile(file) {
    if (!file) return;
    if (!/\.(md|markdown|txt)$/i.test(file.name)) {
      return toast('请选择 .md / .markdown / .txt 文件', 'warn');
    }
    WordMd.office.readMarkdownFile(file).then(function (text) {
      state.pendingFile = file;
      state.pendingText = text;
      $('#file-name').textContent = file.name + '（' + (file.size / 1024).toFixed(1) + ' KB）';
      $('#file-stats').textContent = fmtStats(WordMd.office.computeStats(text));
      $('#file-preview').innerHTML = '<div class="word-preview">' +
        WordMd.markdown.mdToHtml(text, { stylePreset: 'default' }) + '</div>';
      $('#file-result').classList.add('show');
      $('#btn-import-now').disabled = false;
      toast('已载入 ' + file.name, 'success');
    }).catch(fail);
  }

  function doImport() {
    if (!state.pendingFile) return toast('请先选择 .md 文件', 'warn');
    busy(true, '正在生成 Word 文档…');
    WordMd.office.importMarkdownFile(state.pendingFile, {
      preset: state.preset,
      mode: state.importMode
    }).then(function (res) {
      if (res.fallback) {
        toast('当前 Word 版本不支持新建文档，已插入到当前文档末尾', 'warn');
      } else {
        toast('已从 ' + res.filename + ' 生成 Word 文档', 'success');
      }
    }).catch(fail).then(function () { busy(false); });
  }

  /* ------------------------------------------------------------------ *
   * 初始化
   * ------------------------------------------------------------------ */

  function bind() {
    $$('.tab').forEach(function (tab) {
      tab.addEventListener('click', function () { switchTab(tab.dataset.tab); });
    });

    // 导出
    $('#btn-export').addEventListener('click', doExport);
    $('#btn-copy-md').addEventListener('click', doCopyMarkdown);
    $$('input[name="scope"]').forEach(function (r) {
      r.addEventListener('change', function () { state.scope = r.value; });
    });

    // 粘贴
    $('#btn-paste-now').addEventListener('click', doPasteInsert);
    $('#btn-read-clip').addEventListener('click', refreshClipboardPreview);
    $('#btn-clear-paste').addEventListener('click', function () {
      $('#paste-input').value = '';
      renderPreview();
    });
    $('#paste-input').addEventListener('input', renderPreview);
    $$('.seg-btn[data-preview]').forEach(function (b) {
      b.addEventListener('click', function () {
        state.previewMode = b.dataset.preview;
        $$('.seg-btn[data-preview]').forEach(function (x) { x.classList.toggle('active', x === b); });
        renderPreview();
      });
    });

    // 导入
    var dropZone = $('#drop-zone');
    var fileInput = $('#file-input');
    $('#btn-pick-file').addEventListener('click', function () { fileInput.click(); });
    dropZone.addEventListener('click', function () { fileInput.click(); });
    fileInput.addEventListener('change', function () { handleFile(fileInput.files[0]); });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dropZone.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragging');
      });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dropZone.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragging');
      });
    });
    dropZone.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      handleFile(f);
    });
    $('#btn-import-now').addEventListener('click', doImport);

    // 样式预设
    $$('input[name="preset"]').forEach(function (r) {
      r.addEventListener('change', function () {
        state.preset = r.value;
        renderPreview();
      });
    });

    // 插入位置
    $$('input[name="paste-mode"]').forEach(function (r) {
      r.addEventListener('change', function () { state.pasteMode = r.value; });
    });
    $$('input[name="import-mode"]').forEach(function (r) {
      r.addEventListener('change', function () { state.importMode = r.value; });
    });
  }

  function init() {
    bind();
    renderPreview();
    WordMd.office.checkRequirements().then(function (info) {
      var bar = $('#env-info');
      if (!info.ok) {
        bar.textContent = '⚠ ' + info.reason;
        bar.className = 'env-bar warn';
        return;
      }
      var caps = [];
      if (info.supports.createDocument) caps.push('新建文档');
      if (info.supports.ooxml) caps.push('OOXML 导出');
      bar.textContent = '✓ Word 已连接 · ' + (info.platform || '') + (caps.length ? ' · ' + caps.join(' / ') : '');
      bar.className = 'env-bar ok';
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();

/*!
 * commands.js — 功能区（Ribbon）命令处理
 * ------------------------------------------------------------------
 * 三个按钮：导出 Markdown / 粘贴 Markdown / 打开面板
 *
 * 注意：ExecuteFunction 命令运行在隐藏的 commands.html 中，
 * 无法使用对话框或下载，因此这里通过 Office 对话框 + 消息通道
 * 把结果交给用户；简单动作（粘贴）则直接执行。
 */
(function () {
  'use strict';

  var DIALOG_WIDTH = 46;
  var DIALOG_HEIGHT = 46;

  /* ------------------------------------------------------------------ *
   * 对话框辅助
   * ------------------------------------------------------------------ */

  function openDialog(url) {
    return new Promise(function (resolve, reject) {
      Office.context.ui.displayDialogAsync(url, {
        height: DIALOG_HEIGHT,
        width: DIALOG_WIDTH,
        displayInIframe: true,
        promptBeforeOpen: false
      }, function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
        } else {
          resolve(result.value);
        }
      });
    });
  }

  /**
   * 打开一个承载结果的对话框：把数据通过 postMessage 传给对话框页面
   * @param {string} page 'dialog-export.html' | 'dialog-paste.html'
   * @param {object} payload
   */
  function openResultDialog(page, payload) {
    var url = location.origin + '/' + page;
    return openDialog(url).then(function (dialog) {
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var msg = {};
        try { msg = JSON.parse(arg.message); } catch (e) { msg = { action: arg.message }; }
        if (msg.action === 'ready') {
          dialog.messageChild(JSON.stringify({ action: 'data', payload: payload }));
        } else if (msg.action === 'close') {
          dialog.close();
        } else if (msg.action === 'insert') {
          // 对话框请求把 Markdown 插入当前文档
          WordMd.office.insertMarkdown(msg.markdown, { preset: msg.preset || 'default', mode: 'cursor' })
            .then(function () {
              dialog.messageChild(JSON.stringify({ action: 'inserted' }));
            })
            .catch(function (err) {
              dialog.messageChild(JSON.stringify({ action: 'error', message: err.message }));
            });
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        dialog.close();
      });
      return dialog;
    });
  }

  /* ------------------------------------------------------------------ *
   * 命令一：导出 Markdown
   * ------------------------------------------------------------------ */

  function exportMarkdown(event) {
    WordMd.office.exportMarkdown({ scope: 'document', autoDownload: false })
      .then(function (res) {
        return openResultDialog('dialog-export.html', {
          markdown: res.markdown,
          filename: res.filename,
          dataUrl: 'data:text/markdown;charset=utf-8,' + encodeURIComponent(res.markdown),
          stats: res.stats
        });
      })
      .catch(function (err) {
        console.error(err);
        return showError('导出失败：' + err.message);
      })
      .then(function () { event.completed(); });
  }

  /* ------------------------------------------------------------------ *
   * 命令二：粘贴 Markdown → Word 富文本
   * ------------------------------------------------------------------ */

  function pasteMarkdown(event) {
    var proceed = function (text) {
      if (!text || !text.trim()) {
        return showError('剪贴板中没有文本内容。请先复制 Markdown 文本，或打开面板手动粘贴。');
      }
      return WordMd.office.insertMarkdown(text, { preset: 'default', mode: 'cursor' });
    };

    WordMd.office.readClipboard()
      .then(proceed)
      .catch(function (err) {
        console.error(err);
        // 读取剪贴板失败 → 打开对话框让用户手动粘贴
        return openResultDialog('dialog-paste.html', { clipboardError: err.message });
      })
      .then(function () { event.completed(); });
  }

  /* ------------------------------------------------------------------ *
   * 命令三：打开面板
   * ------------------------------------------------------------------ */

  function openPane() {
    // ShowTaskpane 由清单直接处理，此处仅作为兜底
  }

  /* ------------------------------------------------------------------ *
   * 错误提示
   * ------------------------------------------------------------------ */

  function showError(message) {
    return openResultDialog('dialog-error.html', { message: message }).then(function (dialog) {
      // 对话框会自行关闭
      return dialog;
    }).catch(function () {
      console.error(message);
    });
  }

  /* ------------------------------------------------------------------ *
   * 注册
   * ------------------------------------------------------------------ */

  Office.onReady(function () {
    window.exportMarkdown = exportMarkdown;
    window.pasteMarkdown = pasteMarkdown;
    window.openPane = openPane;
  });
})();

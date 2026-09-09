    (function () {
      'use strict';
      var payload = null;

      function render(data) {
        payload = data;
        document.getElementById('preview').textContent = data.markdown.slice(0, 100000);
        var s = data.stats || {};
        document.getElementById('stats').textContent =
          '字符 ' + s.chars + ' · 行 ' + s.lines + ' · 词 ' + s.words +
          ' · 标题 ' + s.headings + ' · 表格行 ' + s.tables + ' · 链接 ' + s.links;
        var dl = document.getElementById('download');
        dl.href = data.dataUrl;
        dl.setAttribute('download', data.filename);
      }

      document.getElementById('copy').addEventListener('click', function () {
        if (!payload) return;
        navigator.clipboard.writeText(payload.markdown).then(function () {
          document.getElementById('status').textContent = '已复制';
          setTimeout(function () { document.getElementById('status').textContent = ''; }, 2000);
        });
      });

      document.getElementById('close').addEventListener('click', function () {
        Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
      });

      Office.onReady(function () {
        Office.context.ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          function (arg) {
            var msg = {};
            try { msg = JSON.parse(arg.message); } catch (e) { msg = {}; }
            if (msg.action === 'data') render(msg.payload);
          }
        ).then(function () {
          Office.context.ui.messageParent(JSON.stringify({ action: 'ready' }));
        });
      });
    })();
  

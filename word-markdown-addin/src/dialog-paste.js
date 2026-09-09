    (function () {
      'use strict';

      document.getElementById('convert').addEventListener('click', function () {
        var md = document.getElementById('input').value;
        if (!md.trim()) {
          document.getElementById('status').textContent = '请输入 Markdown 内容';
          document.getElementById('status').style.color = '#a33';
          return;
        }
        Office.context.ui.messageParent(JSON.stringify({
          action: 'insert',
          markdown: md,
          preset: document.getElementById('preset').value
        }));
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
            if (msg.action === 'data' && msg.payload && msg.payload.clipboardError) {
              document.getElementById('note').textContent =
                '无法自动读取剪贴板（' + msg.payload.clipboardError + '），请在此手动粘贴 Markdown。';
            }
            if (msg.action === 'inserted') {
              var st = document.getElementById('status');
              st.textContent = '已插入文档，可继续编辑';
              st.style.color = '#107c10';
            }
            if (msg.action === 'error') {
              var st2 = document.getElementById('status');
              st2.textContent = '插入失败：' + msg.message;
              st2.style.color = '#a33';
            }
          }
        ).then(function () {
          Office.context.ui.messageParent(JSON.stringify({ action: 'ready' }));
        });
      });
    })();
  

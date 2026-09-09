    (function () {
      'use strict';
      document.getElementById('close').addEventListener('click', function () {
        Office.context.ui.messageParent(JSON.stringify({ action: 'close' }));
      });
      Office.onReady(function () {
        Office.context.ui.addHandlerAsync(
          Office.EventType.DialogParentMessageReceived,
          function (arg) {
            var msg = {};
            try { msg = JSON.parse(arg.message); } catch (e) { msg = {}; }
            if (msg.action === 'data' && msg.payload && msg.payload.message) {
              document.getElementById('message').textContent = msg.payload.message;
            }
          }
        ).then(function () {
          Office.context.ui.messageParent(JSON.stringify({ action: 'ready' }));
        });
      });
    })();
  

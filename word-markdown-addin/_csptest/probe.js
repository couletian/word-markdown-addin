setTimeout(function () {
  var el = document.getElementById('r');
  var parts = [];
  parts.push('Office=' + (typeof Office));
  if (typeof Office !== 'undefined') {
    parts.push('onReady=' + (typeof Office.onReady));
    try {
      Office.onReady(function (info) {
        parts.push('onReadyFired host=' + (info && info.host));
        el.textContent = 'RESULT: ' + parts.join(' ');
      });
      setTimeout(function () { el.textContent = 'RESULT: ' + parts.join(' ') + ' (onReady 未回调)'; }, 2500);
    } catch (e) { parts.push('err=' + e.message); el.textContent = 'RESULT: ' + parts.join(' '); }
  } else {
    el.textContent = 'RESULT: ' + parts.join(' ');
  }
}, 3000);

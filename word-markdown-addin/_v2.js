var md = require('./src/core/markdown.js');
var fs = require('fs');
var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><div id="out">' +
  md.mdToHtml('<svg/onload=window.__p.push("C")>', {stylePreset:'default'}) +
  md.mdToHtml('<a href="javascript&#x3A;window.__p.push(\'M\')">M</a>', {stylePreset:'default'}) +
  md.mdToHtml('<a href="https://example.com">safe</a>', {stylePreset:'default'}) +
  '</div><pre id="r">pending</pre><script>' +
  'window.__p=[];' +
  'document.querySelectorAll("#out a").forEach(function(el){try{el.click();}catch(e){}});' +
  'document.getElementById("r").textContent="EXECUTED:"+JSON.stringify(window.__p);' +
  '</'+'script></body></html>';
fs.writeFileSync('./_v2.html', html);
console.log('已写入 _v2.html，长度', html.length);

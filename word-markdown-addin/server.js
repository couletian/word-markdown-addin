/*!
 * server.js — 本地 HTTPS 开发服务器
 * ------------------------------------------------------------------
 * Office 加载项要求通过 HTTPS 加载。本服务器使用自签名证书，
 * 首次运行会自动生成证书（需要 Node 内置 crypto，无外部依赖）。
 *
 * 用法：
 *   node server.js            # 默认 https://localhost:3000
 *   node server.js --port 3001
 *   node server.js --http     # 纯 HTTP（仅用于快速预览，Word 不支持）
 */
'use strict';

var https = require('https');
var http = require('http');
var fs = require('fs');
var path = require('path');
var url = require('url');

var ROOT = path.join(__dirname, 'src');
var ASSETS = path.join(__dirname, 'assets');
var CERT_DIR = path.join(__dirname, '.certs');

var MIME = {
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff2': 'font/woff2',
  '.md': 'text/markdown; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8'
};

function parseArgs() {
  var args = process.argv.slice(2);
  var out = { port: 3000, http: false };
  for (var i = 0; i < args.length; i++) {
    if (args[i] === '--port' && args[i + 1]) { out.port = parseInt(args[i + 1], 10); i++; }
    else if (args[i] === '--http') out.http = true;
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * 自签名证书生成（调用 openssl；不可用时给出明确指引）
 * ------------------------------------------------------------------ */

function ensureCert() {
  var keyPath = path.join(CERT_DIR, 'key.pem');
  var certPath = path.join(CERT_DIR, 'cert.pem');
  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) };
  }

  fs.mkdirSync(CERT_DIR, { recursive: true });
  var { execFileSync } = require('child_process');
  var confPath = path.join(CERT_DIR, 'openssl.cnf');
  fs.writeFileSync(confPath, [
    '[req]',
    'distinguished_name = dn',
    'x509_extensions = v3',
    'prompt = no',
    '[dn]',
    'CN = localhost',
    '[v3]',
    'subjectAltName = DNS:localhost,IP:127.0.0.1',
    'basicConstraints = CA:FALSE',
    'keyUsage = digitalSignature,keyEncipherment',
    'extendedKeyUsage = serverAuth'
  ].join('\n'));

  try {
    execFileSync('openssl', [
      'req', '-x509', '-newkey', 'rsa:2048', '-nodes',
      '-keyout', keyPath, '-out', certPath,
      '-days', '825', '-config', confPath
    ], { stdio: 'ignore' });
    console.log('✓ 已生成自签名证书：' + CERT_DIR);
    return { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) };
  } catch (e) {
    console.error('\n✗ 无法生成自签名证书（未找到 openssl）。\n');
    console.error('请任选一种方式：');
    console.error('  1. 安装 Git for Windows（自带 openssl），然后重新运行本服务器');
    console.error('  2. 使用 Office 官方的 office-addin-dev-certs：');
    console.error('       npx office-addin-dev-certs install');
    console.error('     然后将其生成的证书路径填入本脚本的 CERT_DIR\n');
    process.exit(1);
  }
}

/* ------------------------------------------------------------------ *
 * 静态文件服务
 * ------------------------------------------------------------------ */

/* 只允许暴露这两个目录。曾经把 __dirname（项目根）也列为候选，
   导致 .certs/key.pem、launcher.py 等敏感文件可被直接下载。 */
var ALLOWED_ROOTS = [ROOT, ASSETS];

/** 安全解码 URL 路径；畸形编码返回 null 而不是抛异常终止进程 */
function safeDecode(pathname) {
  try {
    return decodeURIComponent(pathname);
  } catch (e) {
    return null;
  }
}

function resolveFile(pathname) {
  var clean = safeDecode(pathname.split('?')[0]);
  if (clean === null) return null;
  if (clean === '/' || clean === '') clean = '/taskpane.html';

  var rel = clean.replace(/^\//, '');
  for (var i = 0; i < ALLOWED_ROOTS.length; i++) {
    var root = ALLOWED_ROOTS[i];
    var p = path.resolve(root, rel);
    // 防目录穿越：解析后必须仍在允许目录内
    var relCheck = path.relative(path.resolve(root), p);
    if (relCheck.startsWith('..') || path.isAbsolute(relCheck)) continue;
    try {
      if (fs.statSync(p).isFile()) return p;
    } catch (e) { /* 继续尝试 */ }
  }
  return null;
}

function handler(req, res) {
  var pathname = url.parse(req.url).pathname;
  var file = resolveFile(pathname);

  if (!file) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found: ' + pathname);
    return;
  }

  var ext = path.extname(file).toLowerCase();
  var type = MIME[ext] || 'application/octet-stream';
  try {
    var data = fs.readFileSync(file);
    // 同源策略：加载项页面与静态资源同源，无需 CORS；
    // 通配 * 会让任意网页尝试跨源读取本地文件，故收紧为仅本机来源。
    res.writeHead(200, {
      'Content-Type': type,
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': 'https://localhost:' + opts.port,
      'Vary': 'Origin',
      'X-Content-Type-Options': 'nosniff'
    });
    res.end(data);
  } catch (e) {
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('500 读取文件失败：' + e.message);
  }
}

/* ------------------------------------------------------------------ *
 * 启动
 * ------------------------------------------------------------------ */

var opts = parseArgs();

/* 兜底：任何未预期的异常都不应让服务进程直接退出。
   曾经因为 decodeURIComponent 抛 URIError 导致整个服务被杀。 */
process.on('uncaughtException', function (e) {
  console.error('⚠ 已拦截未捕获异常，服务继续运行：' + (e && e.message));
});
process.on('unhandledRejection', function (e) {
  console.error('⚠ 已拦截未处理的 Promise 拒绝：' + (e && (e.message || e)));
});

function start() {
  if (opts.http) {
    http.createServer(handler).listen(opts.port, '127.0.0.1', function () {
      console.log('\n[HTTP 预览模式] http://localhost:' + opts.port + '/taskpane.html');
      console.log('注意：Word 加载项必须使用 HTTPS，此模式仅用于在浏览器中预览界面。\n');
    });
    return;
  }

  var creds = ensureCert();
  https.createServer({ key: creds.key, cert: creds.cert }, handler)
    .listen(opts.port, '127.0.0.1', function () {
      console.log('\n✓ Markdown 工具箱开发服务器已启动');
      console.log('  任务面板：https://localhost:' + opts.port + '/taskpane.html');
      console.log('  命令页：  https://localhost:' + opts.port + '/commands.html');
      console.log('\n首次访问浏览器会提示证书不受信任，点击「继续访问」即可。');
      console.log('按 Ctrl+C 停止服务。\n');
    })
    .on('error', function (e) {
      if (e.code === 'EADDRINUSE') {
        console.error('✗ 端口 ' + opts.port + ' 已被占用，请用 --port 指定其他端口');
      } else {
        console.error('✗ 服务器启动失败：' + e.message);
      }
      process.exit(1);
    });
}

start();

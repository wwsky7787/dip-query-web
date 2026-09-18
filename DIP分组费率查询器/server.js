// 零依赖静态文件服务器，仅供开发预览使用（npm run dev）。
// 最终用户无需此文件：直接双击 index.html 即可在 Chrome 中使用。
const http = require('http');
const fs = require('fs');
const path = require('path');

// 转发命令行 --host / --port 参数
const args = process.argv.slice(2);
function argOf(name, fallback) {
  const i = args.indexOf('--' + name);
  if (i >= 0 && args[i + 1]) return args[i + 1];
  const eq = args.find(a => a.startsWith('--' + name + '='));
  if (eq) return eq.split('=')[1];
  return process.env[name.toUpperCase()] || fallback;
}
const HOST = argOf('host', '127.0.0.1');
const PORT = parseInt(argOf('port', '7100'), 10);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

http.createServer((req, res) => {
  let urlPath = decodeURIComponent(req.url.split('?')[0]);
  if (urlPath === '/') urlPath = '/index.html';
  const file = path.join(__dirname, path.normalize(urlPath));
  if (!file.startsWith(__dirname)) { res.writeHead(403); res.end(); return; }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not Found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream' });
    res.end(data);
  });
}).listen(PORT, HOST, () => {
  console.log(`DIP查询器预览: http://${HOST}:${PORT}/`);
});

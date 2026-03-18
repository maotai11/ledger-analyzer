const http = require('http');
const fs = require('fs');
const path = require('path');
const port = process.env.PORT ? Number(process.env.PORT) : 5178;
const root = process.cwd();

function contentType(p){
  if (p.endsWith('.html')) return 'text/html; charset=utf-8';
  if (p.endsWith('.js')) return 'application/javascript; charset=utf-8';
  if (p.endsWith('.css')) return 'text/css; charset=utf-8';
  if (p.endsWith('.json')) return 'application/json; charset=utf-8';
  return 'application/octet-stream';
}

const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
  const rel = urlPath === '/' ? 'index.html' : urlPath.replace(/^\//,'');
  const filePath = path.join(root, rel);
  fs.readFile(filePath, (err, buf) => {
    if (err) {
      res.statusCode = 404;
      res.end('not found');
      return;
    }
    res.setHeader('Content-Type', contentType(filePath));
    res.end(buf);
  });
});
server.listen(port, () => console.log('server on http://localhost:' + port));

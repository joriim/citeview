// Minimal HTTPS static server for local testing (uses the Office dev certificate).
// Run "npm run certs" once first.
import { createServer } from 'node:https';
import { readFileSync, existsSync, statSync, createReadStream } from 'node:fs';
import { homedir } from 'node:os';
import { join, extname, normalize, resolve } from 'node:path';

const certDir = join(homedir(), '.office-addin-dev-certs');
const root = resolve('.');
const port = Number(process.env.PORT || 3000);
const types = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript',
  '.css': 'text/css', '.png': 'image/png', '.xml': 'application/xml', '.md': 'text/plain; charset=utf-8',
  '.pdf': 'application/pdf', '.json': 'application/json',
};
if (!existsSync(join(certDir, 'localhost.crt'))) {
  console.error('No dev certificate found. Run "npm run certs" first.'); process.exit(1);
}
createServer({ cert: readFileSync(join(certDir, 'localhost.crt')), key: readFileSync(join(certDir, 'localhost.key')) }, (req, res) => {
  const path = normalize(join(root, decodeURIComponent(new URL(req.url, 'https://x').pathname)));
  if (!path.startsWith(root) || !existsSync(path) || statSync(path).isDirectory()) { res.writeHead(404); res.end('Not found'); return; }
  res.writeHead(200, { 'Content-Type': types[extname(path)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
  createReadStream(path).pipe(res);
}).listen(port, () => console.log(`CiteView served at https://localhost:${port}/src/taskpane.html`));

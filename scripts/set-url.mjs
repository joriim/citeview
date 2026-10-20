// Usage: npm run set-url -- https://your.host/citeview
// Writes manifest.hosted.xml with every https://localhost:3000 replaced by your URL.
import { readFileSync, writeFileSync } from 'node:fs';
const url = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\//.test(url)) { console.error('Give an https:// base URL'); process.exit(1); }
const xml = readFileSync('manifest.xml', 'utf8')
  .replaceAll('<AppDomain>https://localhost:3000</AppDomain>', `<AppDomain>${new URL(url).origin}</AppDomain>`)
  .replaceAll('https://localhost:3000', url);
writeFileSync('manifest.hosted.xml', xml);
console.log('Wrote manifest.hosted.xml for', url);

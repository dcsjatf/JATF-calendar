// Renders the current-month poster from booking.html (same drawing code as the portal)
// and writes poster.png only when it changed. Outputs changed=true|false and v=<hash>.
import { chromium } from 'playwright';
import fs from 'node:fs';
import crypto from 'node:crypto';
import path from 'node:path';

const page = 'file://' + path.resolve('booking.html') + '#poster';
const browser = await chromium.launch();
const p = await browser.newPage({ viewport: { width: 1280, height: 900 } });
await p.goto(page, { waitUntil: 'networkidle' });
// wait until the bookings have loaded from the Apps Script API
await p.waitForFunction(() => {
  const s = document.querySelector('#schWrap');
  return s && !s.textContent.includes('Loading');
}, null, { timeout: 60000 });
const data = await p.evaluate(async () => {
  const sel = document.getElementById('pMonth');
  sel.dispatchEvent(new Event('change'));
  await document.fonts.ready;
  await new Promise(r => setTimeout(r, 2500));
  return document.getElementById('poster').toDataURL('image/png');
});
await browser.close();

const buf = Buffer.from(data.split(',')[1], 'base64');
const old = fs.existsSync('poster.png') ? fs.readFileSync('poster.png') : null;
const changed = !old || !old.equals(buf) || !fs.existsSync('posters/poster-' + crypto.createHash('sha1').update(buf).digest('hex').slice(0, 10) + '.png');
if (changed) fs.writeFileSync('poster.png', buf);
const v = crypto.createHash('sha1').update(buf).digest('hex').slice(0, 10);
// Discord caches images by URL: every version also gets its own file name.
fs.mkdirSync('posters', { recursive: true });
const own = `posters/poster-${v}.png`;
if (!fs.existsSync(own)) fs.writeFileSync(own, buf);
// keep only this version and the previous one
const prev = fs.existsSync('posters/latest.txt') ? fs.readFileSync('posters/latest.txt', 'utf8').trim() : '';
fs.readdirSync('posters').filter(f => /^poster-[0-9a-f]+\.png$/.test(f))
  .filter(f => f !== `poster-${v}.png` && f !== `poster-${prev}.png`)
  .forEach(f => fs.unlinkSync('posters/' + f));
if (prev !== v) fs.writeFileSync('posters/latest.txt', v + '\n');
const out = process.env.GITHUB_OUTPUT;
if (out) fs.appendFileSync(out, `changed=${changed}\nv=${v}\n`);
console.log(changed ? `poster.png updated (${v})` : `poster unchanged (${v})`);

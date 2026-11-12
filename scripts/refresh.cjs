// Member Mile daily refresh.
// Builds data-<today>.json and latest.json from:
//   - International: the "Interntional Reporting Summary FY27" Google Sheet, Summary tab (public export link)
//   - US & Canada: the SharePoint "Game Card For Home Page.xlsx" workbook (Event Involvement & Engagement + Membership tabs)
//
// The US & Canada workbook is downloaded from GAME_CARD_URL (a GitHub Actions secret holding a direct
// download link). If it's missing or the download fails, the previous day's US & Canada rows are reused
// and the site shows "North America pending".
//
// Run locally: npm install && node scripts/refresh.cjs

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.resolve(__dirname, '..');
const INTL_SHEET_ID = process.env.INTL_SHEET_ID || '1PakRC0w8AXLYZxw3r2v48VG2lf9igo-cU6wQqmWgiZ8';
const CANADA = ['Canada Pacific Region', 'Lake Ontario Region', 'Montreal Region', 'Northwest Canada Region', 'Red River Region'];

const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/New_York', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const num = (v) => {
  if (isNum(v)) return v;
  const s = String(v ?? '');
  const n = parseFloat(s.replace(/[%,\s]/g, ''));
  return Number.isFinite(n) ? n / (s.includes('%') ? 100 : 1) : 0;
};

async function download(url) {
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText} for ${url.split('?')[0]}`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.slice(0, 2).toString() !== 'PK') throw new Error('Download was not an Excel file (likely a sign-in page)');
  return buf;
}

function rowsOf(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`Missing tab "${name}"`);
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

async function international() {
  const wb = XLSX.read(await download(`https://docs.google.com/spreadsheets/d/${INTL_SHEET_ID}/export?format=xlsx`), { type: 'buffer' });
  const rows = rowsOf(wb, 'Summary');
  const out = [];
  for (const r of rows.slice(2)) {
    const hub = r[0], region = r[1];
    if (hub == null || region == null || String(hub).trim() === '' || String(region).trim() === '') continue;
    if (String(region).trim().toUpperCase() === 'TOTAL') continue;
    const int = (v) => (isNum(v) ? v : 0);
    out.push([
      'INTL',
      String(hub).trim().replace(/^Estern Europe$/, 'Eastern Europe'),
      String(region).trim().replace(/\?+$/, '').trim(),
      int(r[2]), isNum(r[3]) ? r[3] : null, int(r[4]), int(r[5]), int(r[6]), null,
    ]);
  }
  if (!out.length) throw new Error('No international regions found on the Summary tab');
  return out;
}

async function gameCardBuffer() {
  if (process.env.GAME_CARD_URL) return { buf: await download(process.env.GAME_CARD_URL), from: 'the Game Card link' };
  return null;
}

function northAmerica(buf) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const eng = rowsOf(wb, 'Event Involvement & Engagement');
  const mem = rowsOf(wb, 'Membership');
  const header = (eng[1] || []).slice(0, 5).map((x) => String(x ?? '').trim());
  if (header.join('|') !== 'Region|Uninvolved|Involved|Very Involved|Core') throw new Error(`Unexpected engagement header: ${header.join(', ')}`);
  const hasNew = /new/i.test(String((mem[1] || [])[8] ?? ''));
  const totals = {}, newm = {};
  for (const r of mem.slice(2)) {
    if (!r[0] || String(r[0]).includes('TOTAL')) continue;
    const k = String(r[0]).trim();
    totals[k] = num(r[1]);
    newm[k] = hasNew && r[8] != null && r[8] !== '' ? num(r[8]) : null;
  }
  const rows = [], skipped = [];
  for (const r of eng.slice(2)) {
    if (!r[0] || String(r[0]).includes('TOTAL')) continue;
    const n = String(r[0]).trim(), T = totals[n];
    if (T == null) { skipped.push(n); continue; }
    rows.push(['NA', CANADA.includes(n) ? 'Canada' : 'United States', n, T, null,
      Math.round(num(r[4]) * T), Math.round(num(r[3]) * T), Math.round(num(r[2]) * T), Math.round(num(r[1]) * T), newm[n] ?? null]);
  }
  const updated = String(eng[0]?.[0] ?? '').trim();
  return { rows, skipped, updated: updated ? updated[0].toLowerCase() + updated.slice(1) : '', hasNew };
}

function readJSON(p) { try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch { return null; } }

(async () => {
  const notes = [];
  const latest = readJSON(path.join(ROOT, 'latest.json')) || {};
  const prevDate = latest.today && latest.today < today ? latest.today : latest.previous || null;
  const prevData = latest.today ? readJSON(path.join(ROOT, `data-${latest.today}.json`)) : null;

  const intl = await international();

  let na = [], updated = prevData?.updated || '', missing = [];
  try {
    const src = await gameCardBuffer();
    if (!src) throw new Error('the GAME_CARD_URL secret is not set');
    const res = northAmerica(src.buf);
    na = res.rows;
    updated = res.updated ? 'SharePoint sheet ' + res.updated : updated;
    if (res.skipped.length) notes.push(`Skipped (no membership total): ${res.skipped.join(', ')}`);
    if (!res.hasNew) notes.push('New Members column not found on the Membership tab');
    notes.push(`US & Canada from ${src.from}`);
  } catch (e) {
    na = (prevData?.rows || []).filter((r) => r[0] === 'NA');
    missing = ['North America'];
    notes.push(`US & Canada not refreshed (${e.message}); reusing the previous numbers`);
  }

  const data = { date: today, updated, rows: [...intl, ...na] };
  if (missing.length) data.missing = missing;
  fs.writeFileSync(path.join(ROOT, `data-${today}.json`), JSON.stringify(data));
  fs.writeFileSync(path.join(ROOT, 'latest.json'), JSON.stringify({ today, previous: prevDate }));

  const summary = [`Refreshed ${today}: ${intl.length} international regions, ${na.length} US & Canada regions.`, ...notes].join('\n');
  console.log(summary);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, summary.replace(/\n/g, '\n\n') + '\n');
})().catch((e) => { console.error(e); process.exit(1); });

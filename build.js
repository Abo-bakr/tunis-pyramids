#!/usr/bin/env node
// Rebuilds the Tunis Pyramids P&L dashboard (index.html) from the hotel system's
// "Profit and Loss" exports (.xls / .xlsx), one file per month.
//
//   node build.js --local <folder>    read exports from a local folder
//   node build.js                      read exports from Google Drive
//
// Google Drive mode needs two environment variables (set as GitHub Actions secrets):
//   GDRIVE_SERVICE_ACCOUNT_KEY  the service account's JSON key (whole file contents)
//   GDRIVE_FOLDER_ID            the id of the Drive folder that holds the exports
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const ROOT = __dirname;
const MON_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MON_LONG = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON_AR = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
const Q_AR = ['الأول', 'الثاني', 'الثالث', 'الرابع'];

// ---------- reading the exports ----------

function parseExport(buffer, name) {
  const wb = XLSX.read(buffer, {type: 'buffer'});
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header: 1, raw: true, defval: null});
    let date = null;
    for (const r of rows.slice(0, 10)) for (const c of r || []) {
      const m = typeof c === 'string' && /From\s+(\d{4})-(\d{2})-(\d{2})/.exec(c);
      if (m) date = `${m[1]}-${m[2]}`;
    }
    const isPL = rows.some((r) => r && r[3] === 'Account number');
    if (!date || !isPL) continue;
    const num = (v) => (typeof v === 'number' ? v : Number(v) || 0);
    const acc = [], tot = [];
    let budget = 0;
    for (const r of rows) {
      if (!r) continue;
      const k = typeof r[3] === 'string' ? r[3].trim() : (typeof r[3] === 'number' ? String(r[3]) : '');
      if (/^\d{8}$/.test(k)) {
        acc.push({a: k, n: String(r[4] || '').replace(/\s+/g, ' ').trim(), c: num(r[2]), y: num(r[5])});
        budget += Math.abs(num(r[1])) + Math.abs(num(r[6]));
      } else if (/^Total|^G\.O\.P|^N\.O\.P/.test(k)) {
        tot.push({n: k, c: num(r[2]), y: num(r[5])});
      }
    }
    if (!tot.some((t) => t.n === 'Total Revenue')) continue;
    return {month: date, name, acc, tot, budgetEmpty: budget === 0};
  }
  return null;
}

async function readLocal(dir) {
  return fs.readdirSync(dir).filter((f) => /\.xlsx?$/i.test(f)).map((f) => ({
    name: f, modified: fs.statSync(path.join(dir, f)).mtime.toISOString(), buffer: fs.readFileSync(path.join(dir, f)),
  }));
}

async function driveToken(key) {
  const now = Math.floor(Date.now() / 1000);
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const unsigned = b64({alg: 'RS256', typ: 'JWT'}) + '.' + b64({
    iss: key.client_email, scope: 'https://www.googleapis.com/auth/drive.readonly',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  });
  const sig = crypto.createSign('RSA-SHA256').update(unsigned).sign(key.private_key).toString('base64url');
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: {'content-type': 'application/x-www-form-urlencoded'},
    body: new URLSearchParams({grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig}),
  });
  if (!res.ok) throw new Error('Google sign-in failed: ' + res.status + ' ' + await res.text());
  return (await res.json()).access_token;
}

async function readDrive() {
  const raw = process.env.GDRIVE_SERVICE_ACCOUNT_KEY, folder = process.env.GDRIVE_FOLDER_ID;
  if (!raw || !folder) throw new Error('Set GDRIVE_SERVICE_ACCOUNT_KEY and GDRIVE_FOLDER_ID, or use --local <folder>.');
  const token = await driveToken(JSON.parse(raw));
  const H = {authorization: 'Bearer ' + token};
  const files = [];
  let pageToken = '';
  do {
    const q = encodeURIComponent(`'${folder}' in parents and trashed = false`);
    const url = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=nextPageToken,files(id,name,mimeType,modifiedTime)` +
      `&pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true${pageToken ? '&pageToken=' + pageToken : ''}`;
    const res = await fetch(url, {headers: H});
    if (!res.ok) throw new Error('Listing the Drive folder failed: ' + res.status + ' ' + await res.text());
    const j = await res.json();
    files.push(...j.files);
    pageToken = j.nextPageToken || '';
  } while (pageToken);
  const out = [];
  for (const f of files) {
    let url;
    if (f.mimeType === 'application/vnd.google-apps.spreadsheet') {
      url = `https://www.googleapis.com/drive/v3/files/${f.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
    } else if (/\.xlsx?$/i.test(f.name) || /spreadsheet|excel/.test(f.mimeType)) {
      url = `https://www.googleapis.com/drive/v3/files/${f.id}?alt=media&supportsAllDrives=true`;
    } else continue;
    const res = await fetch(url, {headers: H});
    if (!res.ok) { console.warn('Skipped', f.name, res.status); continue; }
    out.push({name: f.name, modified: f.modifiedTime, buffer: Buffer.from(await res.arrayBuffer())});
  }
  return out;
}

// ---------- figures ----------

const REV = {
  'Rooms': ['Total Rooms Revenue', 'Total Room Revenue'],
  'Food': ['Total Food Pakage Revenue', 'Total Food Outlet Revenue'],
  'Beverage': ['Total Beverage Pakage Revenue', 'Total Beverage Outlet Revenue'],
  'Pool & recreation': ['Total Recreation Revenue'],
  'Excursions & other': ['Total Sundry Revenue', 'Total Laundry Revenue', 'Total Rent Revenue'],
};
const COST = {
  'Food cost of sales': ['Total FOOD: CREDIT TO COST OF SALES'],
  'Beverage cost of sales': ['Total BEV: CREDIT COST OF SALES'],
  'Payroll (all departments)': 'PAYROLL',
  'Staff meals': ['Total Emp.Benefit Cost Of Beverage'],
  'Maintenance (R&M)': ['Total R & M Other Expenses'],
  'Admin & general': ['Total A & G Other Expenses'],
  'F&B operating': ['Total F & B Other Expeses'],
  'Rooms operating': ['Total Rooms Other Expeses'],
  'Sales & marketing': ['Total Sales Other Expenses'],
  'Energy': ['Total Energy Cost Expenses'],
  'FF&E & non-operating': ['Total Non Operating Expenses'],
};
// Departments (USALI style): each operated department's own revenue and direct costs.
// Everything else in Total Expenses is undistributed (A&G, sales, HR, security, R&M, energy, non-operating...).
const DEPT = {
  rooms: {rev: REV.Rooms, cost: ['Total Rooms Payroll Expeses', 'Total Rooms Other Expeses', 'Total H.Keeping Payroll Expenses', 'Total H.Keeping Other Expenses']},
  fnb: {rev: [...REV.Food, ...REV.Beverage], cost: ['Total FOOD: CREDIT TO COST OF SALES', 'Total BEV: CREDIT COST OF SALES', 'Total Cost Of Tobaco', 'Total F & B Payroll Expeses', 'Total F & B Other Expeses']},
  other: {rev: [...REV['Pool & recreation'], ...REV['Excursions & other']], cost: ['Total Laundry Cost Of Sales', 'Total Recreation Other Expenses']},
};
const R = (x) => Math.round(x);
const n0 = (v) => R(v).toLocaleString('en-US').replace('-', '−');
const p1 = (v) => (Math.round(v * 10) / 10).toFixed(1);
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c]));
const ym = (y, m) => `${y}-${String(m).padStart(2, '0')}`;
const lab = (id) => { const [y, m] = id.split('-').map(Number); return {en: `${MON_EN[m - 1]} ${y}`, ar: `${MON_AR[m - 1]} ${y}`}; };
const bi = (en, ar) => `<span class="en">${en}</span><span class="ar" lang="ar">${ar}</span>`;

function getter(fn) {
  // fn(totalName) -> amount. Adds group helpers.
  const s = (names) => names.reduce((a, n) => a + fn(n), 0);
  return {g: fn, s};
}

function build(exports) {
  const byMonth = {};
  for (const e of exports) {
    const p = parseExport(e.buffer, e.name);
    if (!p) { console.log('Not a P&L export, skipped:', e.name); continue; }
    p.modified = e.modified;
    if (!byMonth[p.month] || byMonth[p.month].modified < p.modified) byMonth[p.month] = p;
  }
  const months = Object.keys(byMonth).sort();
  if (!months.length) throw new Error('No "Profit and Loss" exports found.');
  const tot = (m, n, k) => { const t = byMonth[m].tot.find((x) => x.n === n); return t ? t[k] : 0; };
  const acc = (m, a, k) => { const x = byMonth[m] && byMonth[m].acc.find((r) => r.a === a); return x ? x[k] : 0; };
  const payrollNames = [...new Set(months.flatMap((m) => byMonth[m].tot.map((t) => t.n)).filter((n) => /Payroll/.test(n)))];
  const otherCost = [...new Set(months.flatMap((m) => byMonth[m].tot.map((t) => t.n)))].filter((n) => /^Total /.test(n) &&
    !/Revenue|Total Expenses|Total Hotel|Total Other deductions|Payroll/.test(n) &&
    !Object.values(COST).flat().includes(n));
  const L = months[months.length - 1];
  const [Y, Lm] = L.split('-').map(Number);
  const prevId = (id) => { const [y, m] = id.split('-').map(Number); return m === 1 ? ym(y - 1, 12) : ym(y, m - 1); };
  const P = byMonth[prevId(L)] ? prevId(L) : null;
  const yearMonths = months.filter((m) => m.startsWith(Y + '-'));
  const first = yearMonths[0];
  const f = Number(first.split('-')[1]);

  // Period definitions: id -> {label, kind, months, cmp, value(totalName)}
  const periods = [];
  const preId = f > 1 ? `${Y}-PRE` : null;
  const preMonths = f - 1;
  if (preId) {
    const range = preMonths === 1 ? {en: `Jan ${Y}`, ar: `يناير ${Y}`} : {en: `Jan–${MON_EN[f - 2]} ${Y}`, ar: `يناير–${MON_AR[f - 2]} ${Y}`};
    const pre = (n) => tot(first, n, 'y') - tot(first, n, 'c');
    periods.push({id: preId, label: range.en + ' (before monthly files)', label_ar: range.ar + ' (قبل الملفات الشهرية)', kind: 'quarter', months: preMonths, cmp: '', v: pre});
    periods.push({id: `${Y}-PRE-avgm`, label: `${range.en} monthly average`, label_ar: `متوسط شهري ${range.ar}`, kind: 'ref', months: 1, cmp: '', v: (n) => pre(n) / preMonths});
    if (preMonths >= 3) periods.push({id: `${Y}-PRE-avgq`, label: `${range.en} quarterly average`, label_ar: `متوسط ربع سنوي ${range.ar}`, kind: 'ref', months: 3, cmp: '', v: (n) => pre(n) * 3 / preMonths});
  }
  for (const m of months) {
    const pm = prevId(m);
    let cmp = byMonth[pm] ? pm : '';
    if (!cmp && m === first && preId) cmp = `${Y}-PRE-avgm`;
    const l = lab(m);
    periods.push({id: m, label: l.en, label_ar: l.ar, kind: 'month', months: 1, cmp, v: (n) => tot(m, n, 'c')});
  }
  let latestQuarter = null;
  const fullQuarters = [];
  for (let q = 1; q <= 4; q++) {
    const qMonths = [0, 1, 2].map((i) => ym(Y, (q - 1) * 3 + 1 + i));
    const have = qMonths.filter((m) => byMonth[m]);
    if (!have.length) continue;
    const full = have.length === 3;
    const id = `${Y}-Q${q}`;
    const names = have.map((m) => MON_EN[Number(m.split('-')[1]) - 1]);
    const namesAr = have.map((m) => MON_AR[Number(m.split('-')[1]) - 1]);
    let cmp = '';
    if (full) {
      const prevQ = fullQuarters[fullQuarters.length - 1];
      if (prevQ && prevQ.q === q - 1) cmp = prevQ.id;
      else if (preId && preMonths >= 3) cmp = `${Y}-PRE-avgq`;
      fullQuarters.push({q, id});
    }
    periods.push({
      id, kind: 'quarter', months: have.length, cmp,
      label: `Q${q} ${Y} (${full ? names[0] + '–' + names[2] : names.join(', ') + ' so far'})`,
      label_ar: `الربع ${Q_AR[q - 1]} ${Y} (${full ? namesAr[0] + '–' + namesAr[2] : namesAr.join('، ') + ' حتى الآن'})`,
      v: (n) => have.reduce((a, m) => a + tot(m, n, 'c'), 0),
    });
    latestQuarter = {id, q, full, label: periods[periods.length - 1].label, label_ar: periods[periods.length - 1].label_ar};
  }
  const ytdId = `${Y}-YTD`;
  periods.push({id: ytdId, kind: 'year', months: Lm, cmp: '', label: `${Y} to date (Jan–${MON_EN[Lm - 1]})`, label_ar: `${Y} حتى تاريخه (يناير–${MON_AR[Lm - 1]})`, v: (n) => tot(L, n, 'y')});

  const periodRows = [], groupRows = [];
  for (const p of periods) {
    const {g, s} = getter(p.v);
    const pay = s(payrollNames);
    const row = {
      period: p.id, label: p.label, label_ar: p.label_ar, kind: p.kind, months: p.months, cmp: p.cmp,
      revenue: R(g('Total Revenue')), expenses: R(g('Total Expenses')), gop: R(g('Total Revenue') - g('Total Expenses')),
      rooms_revenue: R(s(REV.Rooms)), food_revenue: R(s(REV.Food)), beverage_revenue: R(s(REV.Beverage)),
      food_cost: R(s(COST['Food cost of sales'])), beverage_cost: R(s(COST['Beverage cost of sales'])),
      payroll: R(pay), maintenance: R(s(COST['Maintenance (R&M)'])), energy: R(s(COST.Energy)),
    };
    let deptCost = 0;
    for (const [k, d] of Object.entries(DEPT)) {
      const rv = s(d.rev), c = s(d.cost);
      deptCost += c;
      row[k + '_dept_revenue'] = R(rv); row[k + '_dept_cost'] = R(c); row[k + '_profit'] = R(rv - c);
    }
    row.undistributed = R(g('Total Expenses') - deptCost);
    periodRows.push(row);
    if (p.kind === 'ref') continue;
    let rs = 0, cs = 0;
    for (const [grp, names] of Object.entries(REV)) { const v = s(names); rs += v; groupRows.push({period: p.id, section: 'Revenue', group: grp, amount: R(v)}); }
    for (const [grp, names] of Object.entries(COST)) { const v = names === 'PAYROLL' ? pay : s(names); cs += v; groupRows.push({period: p.id, section: 'Cost', group: grp, amount: R(v)}); }
    const other = s(otherCost); cs += other;
    groupRows.push({period: p.id, section: 'Cost', group: 'Other', amount: R(other)});
    const dr = Math.abs(rs - g('Total Revenue')), dc = Math.abs(cs - g('Total Expenses'));
    if (dr > 1 || dc > 1) console.warn(`Warning: groups do not add up for ${p.id} (revenue off by ${R(dr)}, costs off by ${R(dc)})`);
  }

  const summary = months.slice(-12).map((m) => {
    const pr = periodRows.find((r) => r.period === m);
    return {month: lab(m).en, period: m, revenue: pr.revenue, expenses: pr.expenses, gop: pr.gop,
      rooms_profit: pr.rooms_profit, fnb_profit: pr.fnb_profit, other_profit: pr.other_profit, undistributed: pr.undistributed};
  });

  let movers = [];
  if (P) {
    const ids = new Set([...byMonth[L].acc, ...byMonth[P].acc].map((a) => a.a));
    movers = [...ids].filter((a) => a !== '90100062').map((a) => {
      const cur = acc(L, a, 'c'), prev = acc(P, a, 'c');
      const d = (byMonth[L].acc.find((r) => r.a === a) || byMonth[P].acc.find((r) => r.a === a)).n;
      return {account: a, description: d, type: /^[1-5]01|^58|^59|^55160|^40/.test(a) ? 'Revenue' : 'Cost', prev: R(prev), cur: R(cur), change: R(cur - prev)};
    }).sort((x, y) => Math.abs(y.change) - Math.abs(x.change)).slice(0, 15);
  }

  // ---------- generated text ----------
  const ratio = (id) => {
    const r = periodRows.find((x) => x.period === id);
    if (!r) return null;
    return {r, food: r.food_cost / r.food_revenue * 100, bev: r.beverage_cost / r.beverage_revenue * 100,
      payroll: r.payroll / r.revenue * 100, maint: r.maintenance / r.revenue * 100, margin: r.gop / r.revenue * 100};
  };
  const cL = ratio(L), cP = P ? ratio(P) : null;
  const Ln = lab(L), Pn = P ? lab(P) : null;
  const alerts = [], actions = [];
  const add = (sev, icon, en, ar) => alerts.push({sev, icon, en, ar});
  if (cL.r.gop < 0) add('dn', '!', `<b>GOP loss of <span class="tp-num">${n0(cL.r.gop)}</span></b> in ${Ln.en} (<span class="tp-num">${p1(cL.margin)}%</span> margin).`,
    `<b>خسارة تشغيلية <span class="tp-num">${n0(cL.r.gop)}</span></b> في ${Ln.ar} (هامش <span class="tp-num">${p1(cL.margin)}%</span>).`);
  if (cL.food > 35) {
    add('dn', '!', `<b>Food cost at <span class="tp-num">${p1(cL.food)}%</span></b> of food revenue (norm ≤ 35%)${cP ? `, vs <span class="tp-num">${p1(cP.food)}%</span> in ${Pn.en}` : ''}.`,
      `<b>تكلفة الأغذية <span class="tp-num">${p1(cL.food)}%</span></b> من إيرادات الأغذية (المعيار ≤ 35%)${cP ? `، مقابل <span class="tp-num">${p1(cP.food)}%</span> في ${Pn.ar}` : ''}.`);
    actions.push([`Audit ${Ln.en} food cost: stock count, purchase cut-off, transfers, wastage.`, 'Ceiling 35%, reported weekly',
      `مراجعة تكلفة الأغذية لشهر ${Ln.ar}: جرد المخزون، فصل فترات المشتريات، التحويلات، والهدر.`, 'حد أقصى 35%، بتقرير أسبوعي']);
  } else add('up', '✓', `<b>Food cost <span class="tp-num">${p1(cL.food)}%</span></b>, within the 35% norm.`, `<b>تكلفة الأغذية <span class="tp-num">${p1(cL.food)}%</span></b>، ضمن معيار 35%.`);
  if (cL.payroll > 35) {
    const rising = cP && cL.r.payroll > cP.r.payroll && cL.r.revenue < cP.r.revenue;
    add(cL.payroll > 40 ? 'dn' : 'wn', '!', `<b>Payroll <span class="tp-num">${p1(cL.payroll)}%</span> of revenue</b> (<span class="tp-num">${n0(cL.r.payroll)}</span>)${rising ? ', rising while revenue fell' : ''}.`,
      `<b>الرواتب <span class="tp-num">${p1(cL.payroll)}%</span> من الإيرادات</b> (<span class="tp-num">${n0(cL.r.payroll)}</span>)${rising ? '، ترتفع رغم انخفاض الإيرادات' : ''}.`);
    actions.push(['Flex labour to occupancy: leave, extra wages, hiring.', 'Track payroll per occupied room',
      'ربط العمالة بنسبة الإشغال: الإجازات، الأجور الإضافية، والتعيين.', 'متابعة الرواتب لكل غرفة مشغولة']);
  }
  if (cP) {
    const ch = (cL.r.revenue / cP.r.revenue - 1) * 100;
    if (ch <= -15) add('wn', '↓', `<b>Revenue down <span class="tp-num">${p1(-ch)}%</span></b> vs ${Pn.en}, to <span class="tp-num">${n0(cL.r.revenue)}</span>.`,
      `<b>انخفضت الإيرادات <span class="tp-num">${p1(-ch)}%</span></b> مقارنة بـ ${Pn.ar}، إلى <span class="tp-num">${n0(cL.r.revenue)}</span>.`);
    else if (ch >= 10) add('up', '↑', `<b>Revenue up <span class="tp-num">${p1(ch)}%</span></b> vs ${Pn.en}, to <span class="tp-num">${n0(cL.r.revenue)}</span>.`,
      `<b>ارتفعت الإيرادات <span class="tp-num">${p1(ch)}%</span></b> مقارنة بـ ${Pn.ar}، إلى <span class="tp-num">${n0(cL.r.revenue)}</span>.`);
  }
  if (cL.maint > 6) {
    add('wn', '⚙', `<b>Maintenance <span class="tp-num">${n0(cL.r.maintenance)}</span></b>, or <span class="tp-num">${p1(cL.maint)}%</span> of revenue (norm ≤ 6%).`,
      `<b>الصيانة <span class="tp-num">${n0(cL.r.maintenance)}</span></b>، أي <span class="tp-num">${p1(cL.maint)}%</span> من الإيرادات (المعيار ≤ 6%).`);
    actions.push(['Move building and furniture works to the FF&amp;E reserve or capital.', 'Cap routine R&amp;M monthly',
      'نقل أعمال المباني والأثاث إلى مخصص الإحلال والتجديد أو المصروفات الرأسمالية.', 'حد شهري للصيانة الدورية']);
  }
  if (P) {
    const e = acc(L, '85610001', 'c'), eP = acc(P, '85610001', 'c');
    if (eP > 5000 && e < eP * 0.3) {
      add('wn', 'ϟ', `<b>Electricity only <span class="tp-num">${n0(e)}</span></b> vs <span class="tp-num">${n0(eP)}</span> in ${Pn.en}: possibly a missing bill.`,
        `<b>الكهرباء <span class="tp-num">${n0(e)}</span> فقط</b> مقابل <span class="tp-num">${n0(eP)}</span> في ${Pn.ar}: غالباً فاتورة غير مسجلة.`);
      actions.push([`Accrue ${Ln.en} electricity on meter readings or the prior bill.`, 'Accrue utilities every month',
        `تسجيل مستحقات كهرباء ${Ln.ar} بناءً على قراءة العداد أو الفاتورة السابقة.`, 'تسجيل مستحقات المرافق شهرياً']);
    }
  }
  if (cL.bev <= 25) add('up', '✓', `<b>Beverage cost <span class="tp-num">${p1(cL.bev)}%</span></b>, within the 25% norm.`, `<b>تكلفة المشروبات <span class="tp-num">${p1(cL.bev)}%</span></b>، ضمن معيار 25%.`);
  else if (cL.bev > 30) add('wn', '!', `<b>Beverage cost <span class="tp-num">${p1(cL.bev)}%</span></b> of beverage revenue (norm ≤ 25%).`, `<b>تكلفة المشروبات <span class="tp-num">${p1(cL.bev)}%</span></b> من إيرادات المشروبات (المعيار ≤ 25%).`);
  const sevOrder = {dn: 0, wn: 1, up: 2};
  alerts.sort((a, b) => sevOrder[a.sev] - sevOrder[b.sev]);
  const sevLab = {dn: ['Critical', 'حرج'], wn: ['Watch', 'للمتابعة'], up: ['Positive', 'إيجابي']};
  const feed = (lang) => alerts.slice(0, 6).map((a) =>
    `<li><span class="dot ${a.sev}" aria-hidden="true">${a.icon}</span><div>${a[lang]}<small>${sevLab[a.sev][lang === 'en' ? 0 : 1]} · ${Ln[lang]}</small></div></li>`).join('\n     ');

  const checks = [];
  const yMonthsEn = yearMonths.map((m) => MON_LONG[Number(m.split('-')[1]) - 1]).join(', ');
  const yMonthsAr = yearMonths.map((m) => MON_AR[Number(m.split('-')[1]) - 1]).join('، ');
  const r401 = acc(L, '40010001', 'y');
  const r401Quiet = yearMonths.every((m) => Math.abs(acc(m, '40010001', 'c')) < 0.5);
  const ytdGop = periodRows.find((r) => r.period === ytdId).gop;
  let ytdNote = '';
  if (r401 > 0 && r401Quiet) {
    checks.push([`Account 40010001 “Room Revenue” holds <b class="tp-num">${n0(r401)}</b> year-to-date with no postings in ${yMonthsEn}, beside rooms allotments of <span class="tp-num">${n0(acc(L, '10100001', 'y'))}</span> YTD. If it duplicates rooms revenue booked earlier in the year, the YTD GOP of <span class="tp-num">${n0(ytdGop)}</span> is overstated.`,
      `الحساب 40010001 «Room Revenue» يحمل <b class="tp-num">${n0(r401)}</b> منذ بداية العام دون أي قيود في ${yMonthsAr}، إلى جانب إيرادات الغرف المخصصة البالغة <span class="tp-num">${n0(acc(L, '10100001', 'y'))}</span>. إذا كان يكرر إيرادات غرف سُجلت في وقت سابق، فإن الربح التشغيلي منذ بداية العام البالغ <span class="tp-num">${n0(ytdGop)}</span> مُبالغ فيه.`]);
    if (preId) {
      const pr = periods.find((p) => p.id === preId);
      ytdNote = bi(`This period includes ${pr.label.replace(' (before monthly files)', '')}, where account 40010001 “Room Revenue” adds <b class="tp-num">${n0(r401)}</b> to rooms revenue. It may be a duplicate; see the data checks.`,
        `تشمل هذه الفترة ${pr.label_ar.replace(' (قبل الملفات الشهرية)', '')}، حيث يضيف الحساب 40010001 «Room Revenue» مبلغ <b class="tp-num">${n0(r401)}</b> إلى إيرادات الغرف. قد يكون مكرراً؛ راجع مراجعة البيانات.`);
    }
  }
  const bc = acc(L, '20220002', 'c'), by = acc(L, '20220002', 'y');
  if (by > 0 && bc < 0) checks.push([`“BEV: Credit cost of sales” is negative in ${Ln.en} (<span class="tp-num">${n0(bc)}</span>) but positive <span class="tp-num">${n0(by)}</span> year-to-date. A wrong-sign posting earlier in the year would overstate YTD beverage cost.`,
    `حساب «دائن تكلفة مبيعات المشروبات» سالب في ${Ln.ar} (<span class="tp-num">${n0(bc)}</span>) لكنه موجب بقيمة <span class="tp-num">${n0(by)}</span> منذ بداية العام. قيد بإشارة خاطئة في وقت سابق يضخم تكلفة المشروبات منذ بداية العام.`]);
  const py = acc(L, '90100050', 'y');
  if (Math.abs(py) > 0.5) checks.push([`YTD expenses include <span class="tp-num">${n0(py)}</span> of prior-year adjustments (account 90100050).`,
    `تتضمن المصروفات منذ بداية العام <span class="tp-num">${n0(py)}</span> تسويات تخص سنوات سابقة (الحساب 90100050).`]);
  const issues = checks.length;
  if (byMonth[L].budgetEmpty) checks.push(['Budget and last-year columns in the export are empty, so there is no budget comparison.',
    'أعمدة الموازنة والعام السابق فارغة في التصدير، لذا لا توجد مقارنة بالموازنة.']);
  if (issues) actions.push(['Investigate the flagged postings in the data checks and restate YTD GOP.', 'Before the next report',
    'التحقق من القيود المذكورة في مراجعة البيانات وإعادة احتساب الربح منذ بداية العام.', 'قبل التقرير القادم']);
  actions.push(['Load budget and last-year figures; add occupancy, ADR and RevPAR.', 'Variance in every monthly pack',
    'إدخال الموازنة وأرقام العام السابق، وإضافة نسبة الإشغال ومتوسط سعر الغرفة والإيراد لكل غرفة متاحة.', 'تحليل الانحرافات في كل تقرير شهري']);

  const ol = (cls, items) => `<ol class="${cls}">\n     ${items.map((i) => `<li>${i}</li>`).join('\n     ')}\n    </ol>`;
  const checksHtml = checks.length
    ? ol('en', checks.map((c) => c[0])) + '\n    ' + ol('ar" lang="ar', checks.map((c) => c[1]))
    : `<p>${bi('No posting issues found in the latest export.', 'لا توجد ملاحظات على القيود في آخر تصدير.')}</p>`;
  const actionsHtml = `<ol class="tp-time en">\n     ${actions.map((a, i) => `<li><span class="n">${i + 1}</span><div>${a[0]}<small>${a[1]}</small></div></li>`).join('\n     ')}\n    </ol>\n    ` +
    `<ol class="tp-time ar" lang="ar">\n     ${actions.map((a, i) => `<li><span class="n">${i + 1}</span><div>${a[2]}<small>${a[3]}</small></div></li>`).join('\n     ')}\n    </ol>`;
  const qShort = latestQuarter ? {en: latestQuarter.label.replace(/ \(.*\)$/, ''), ar: latestQuarter.label_ar.replace(/ \(.*\)$/, '')} : {en: String(Y), ar: String(Y)};
  const ctaHtml = issues
    ? `<b>${bi('Fix the books first', 'صحّح الدفاتر أولاً')}</b>\n    <p>${bi(`${issues} posting${issues > 1 ? 's' : ''} could change the year-to-date result. Review before reporting ${qShort.en}.`, `${issues} قيود قد تغير نتيجة العام حتى الآن. راجعها قبل إصدار تقرير ${qShort.ar}.`)}</p>`
    : `<b>${bi('Books look consistent', 'الدفاتر متسقة')}</b>\n    <p>${bi('No posting issues found in the latest export.', 'لا توجد ملاحظات على القيود في آخر تصدير.')}</p>`;

  const firstL = lab(months[0]);
  const sameYear = months[0].slice(0, 4) === String(Y);
  const rangeEn = sameYear ? `${MON_EN[Number(months[0].slice(5)) - 1]}–${MON_EN[Lm - 1]} ${Y}` : `${firstL.en} – ${Ln.en}`;
  const rangeAr = sameYear ? `${MON_AR[Number(months[0].slice(5)) - 1]}–${MON_AR[Lm - 1]} ${Y}` : `${firstL.ar} – ${Ln.ar}`;
  const asOf = months.map((m) => byMonth[m].modified).sort().pop();
  const asOfDate = String(asOf).slice(0, 10);
  const fileList = months.map((m) => esc(byMonth[m].name)).join(', ');
  const trendMonths = months.slice(-12);
  const trendEn = trendMonths.length === 1 ? lab(trendMonths[0]).en : `${lab(trendMonths[0]).en} to ${lab(trendMonths[trendMonths.length - 1]).en}`;
  const trendAr = trendMonths.length === 1 ? lab(trendMonths[0]).ar : `${lab(trendMonths[0]).ar} إلى ${lab(trendMonths[trendMonths.length - 1]).ar}`;

  const slots = {
    FOOT: bi(`Source: system P&amp;L exports, ${rangeEn}. Amounts in EGP. Data as of ${asOfDate}.`, `المصدر: قوائم الأرباح والخسائر من النظام، ${rangeAr}. المبالغ بالجنيه المصري. البيانات حتى ${asOfDate}.`),
    H1: bi(`Profit &amp; loss, ${rangeEn}`, `الأرباح والخسائر، ${rangeAr}`),
    PILL: bi(`${qShort.en} · EGP`, `${qShort.ar} · ج.م`),
    YTD_NOTE: ytdNote,
    SCOPE: bi(`Trend lines and the monthly charts show ${trendEn}.`, `خطوط الاتجاه والرسوم الشهرية تعرض ${trendAr}.`),
    YTD_LABEL: bi(`Year to date (Jan–${MON_EN[Lm - 1]}):`, `منذ بداية العام (يناير–${MON_AR[Lm - 1]}):`),
    MOVERS_SUB: P ? bi(`${Ln.en} vs ${Pn.en}, top 15 accounts by size of change. Red hurts profit.`, `${Ln.ar} مقابل ${Pn.ar}، أكبر 15 حساباً من حيث حجم التغير. الأحمر يضر بالربح.`)
      : bi('Needs two consecutive months of exports.', 'يتطلب ملفين لشهرين متتاليين.'),
    PREV_SHORT: P ? bi(MON_EN[Number(P.slice(5)) - 1], MON_AR[Number(P.slice(5)) - 1]) : bi('Prev', 'السابق'),
    CUR_SHORT: bi(MON_EN[Lm - 1], MON_AR[Lm - 1]),
    CHECKS: checksHtml,
    NOTES: `<p class="en">Built automatically from the system exports in the Google Drive folder (${fileList}), current month and YTD columns. Department groups add up to each export's Total Revenue and Total Expenses. GOP is total revenue minus total expenses as reported. Norms are typical ranges for a full-service resort, not this hotel's budget. Currency assumed to be Egyptian pounds.</p>\n    ` +
      `<p class="ar" lang="ar">تُبنى تلقائياً من ملفات النظام في مجلد Google Drive (${fileList})، أعمدة الشهر الحالي ومنذ بداية العام. مجموعات الأقسام تساوي إجمالي الإيرادات وإجمالي المصروفات في كل ملف. الربح التشغيلي هو إجمالي الإيرادات ناقص إجمالي المصروفات كما وردت. المعايير نطاقات معتادة لمنتجع متكامل الخدمات وليست موازنة الفندق. العملة مفترضة بالجنيه المصري.</p>`,
    ALERTS_SCOPE: bi(`Findings for ${Ln.en}`, `ملاحظات ${Ln.ar}`),
    ALERTS: `<ul class="tp-feed en">\n     ${feed('en')}\n    </ul>\n    <ul class="tp-feed ar" lang="ar">\n     ${feed('ar')}\n    </ul>`,
    ACTIONS: actionsHtml,
    CTA: ctaHtml,
    LATEST: L, PREV: P || L, YTD: ytdId,
  };
  const meta = {
    defaults: {year: ytdId, quarter: latestQuarter ? latestQuarter.id : ytdId, month: L},
    preId: preId || '', preMonths: preMonths || 1, ytdNoteFor: ytdNote ? [ytdId, preId] : [],
    rangeEn, asOf: asOfDate, months,
  };
  const src = {summary, groups: groupRows, movers, periods: periodRows};
  return {slots, meta, src, months, latest: L};
}

function render({slots, meta, src}) {
  let page = fs.readFileSync(path.join(__dirname, 'page-template.html'), 'utf8');
  page = page.replace(/\{\{([A-Z0-9_]+)\}\}/g, (m, k) => {
    if (!(k in slots)) throw new Error('No value for slot ' + k);
    return slots[k];
  });
  const shell = fs.readFileSync(path.join(__dirname, 'shell.html'), 'utf8');
  const json = (o) => JSON.stringify(o).replace(/</g, '\\u003c');
  return shell.replace('/*SRC*/', () => json(src)).replace('/*META*/', () => json(meta)).replace('<!--PAGE-->', () => page);
}

(async () => {
  const i = process.argv.indexOf('--local');
  const files = i > 0 ? await readLocal(process.argv[i + 1]) : await readDrive();
  const out = process.argv.includes('--out') ? process.argv[process.argv.indexOf('--out') + 1] : path.join(ROOT, 'index.html');
  const result = build(files);
  const html = render(result);
  const old = fs.existsSync(out) ? fs.readFileSync(out, 'utf8') : '';
  if (old === html) { console.log('No change: dashboard already shows', result.months.join(', ')); return; }
  fs.writeFileSync(out, html);
  console.log('Dashboard rebuilt for', result.months.join(', '), '- latest', result.latest);
})().catch((e) => { console.error(e.message || e); process.exit(1); });

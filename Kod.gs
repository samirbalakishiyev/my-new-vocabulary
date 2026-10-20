/* My New Vocabulary · Google Cədvəl skripti (Apps Script)
   Bu cədvəldə yalnız istifadəçilər (qeydiyyat, son giriş, blok, ödəniş, 👑 premium) və müəllim düzəlişləri saxlanılır.
   Nəticələr cədvələ YAZILMIR. Premium yalnız bu saytındır (English 5–11-dən asılı deyil).
   Köhnə "Qeydiyyat" vərəqindəki istifadəçilər avtomatik "Users" vərəqinə köçürülür; köhnə vərəqlər silinmir.
   Yeniləmək: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (link dəyişmir). */
const KEY = 'BURAYA-OZ-ACARINIZI-YAZIN';

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Users vərəqi (oxunaqlı başlıqlar, tarixlər, blok qutucuqları) ---------- */
const UL = ['Telefon', 'Ad', 'Soyad', 'Rol', 'Qeydiyyat tarixi', 'Son giriş', 'Bloklanıb', 'Sayt', 'Əlaqə saxlanılıb', 'Premium'];
const DF = 'dd.mm.yyyy hh:mm';
const isTrue = x => x === true || /^(true|bəli|beli|yes|1|blok)$/i.test(String(x).trim());
const ms = x => x instanceof Date ? x.getTime() : (+x || '');
const asDate = x => (typeof x === 'number' && x > 1e11) ? new Date(x) : x;

/* başlıq adlarından sütun yerlərini tapır (köhnə və yeni adlar) */
function colMap(hv) {
  const h = hv.map(x => String(x).trim().toLowerCase().replace(/\(.*?\)/g, '').trim());
  const T = [
    k => /^(tel|telefon|nömrə|nomre|phone)/.test(k),
    k => k === 'ad' || k === 'name',
    k => k === 'soyad' || k === 'surname',
    k => k === 'rol' || k === 'role' || k.includes('müəllim'),
    k => k === 'qeyd' || k.startsWith('qeydiyyat') || k === 'at',
    k => k === 'son' || k.startsWith('son giri') || k === 'last',
    k => k === 'blok' || k.startsWith('blok') || k === 'block',
    k => k === 'sayt' || k === 'site',
    k => k === 'elaqe' || k.startsWith('əlaqə') || k.startsWith('elaqe'),
    k => k === 'prem' || k.startsWith('premium')];
  return T.map(f => h.findIndex(f));
}
/* başqa vərəqlərdə qalmış köhnə istifadəçiləri Users-ə köçürür (təkrarlamadan) */
function usersImport(s) {
  const ss = SpreadsheetApp.getActive();
  const have = new Set(s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const add = [];
  ss.getSheets().forEach(sh => {
    const nm = sh.getName();
    if (/^users/i.test(nm) || /config|results/i.test(nm) || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
    const v = sh.getDataRange().getValues(), idx = colMap(v[0]);
    if (idx[0] < 0 || idx[1] < 0) return;
    v.slice(1).forEach(r => {
      const tel = String(r[idx[0]] || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '');
      if (!tel || have.has(tel)) return;
      have.add(tel);
      add.push(idx.map((i, j) => j === 0 ? "'" + tel : i < 0 ? (j === 6 || j === 9 ? false : '') : r[i]));
    });
  });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, UL.length).setValues(add);
}
function usersSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Users');
  if (!s) { s = ss.insertSheet('Users'); s.setFrozenRows(1); }
  const w = Math.max(s.getLastColumn(), 1);
  const hv = s.getRange(1, 1, 1, w).getValues()[0].map(String);
  if (hv.slice(0, UL.length).join('|') !== UL.join('|')) {
    /* köhnə başlıqlar (tel, ad, soyad, rol, qeyd, son, blok, sayt, elaqe) → yeni düzülüş */
    const idx = colMap(hv);
    const n = s.getLastRow() - 1;
    if (n > 0 && !ss.getSheetByName('Users yedək')) s.copyTo(ss).setName('Users yedək');
    if (n > 0 && idx[0] < 0) return s; /* tanınmayan format: heç nəyə toxunma */
    if (n > 0 && idx.some((i, j) => i !== j)) {
      const all = s.getRange(2, 1, n, w).getValues();
      const nv = all.map(r => idx.map((i, j) => i < 0 ? (j === 9 ? false : '') : r[i]));
      s.getRange(2, 1, n, Math.max(w, UL.length)).clearContent();
      s.getRange(2, 1, n, UL.length).setValues(nv);
    }
    s.getRange(1, 1, 1, Math.max(w, UL.length)).clearContent();
    s.getRange(1, 1, 1, UL.length).setValues([UL]).setFontWeight('bold');
  }
  return s;
}
/* köhnə sətirləri düzəldir: uzun rəqəmlər → tarix, TRUE/FALSE → qutucuq */
function usersFix(s) {
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  const rg = s.getRange(2, 1, n, UL.length), v = rg.getValues();
  let ch = false;
  v.forEach(r => { [4, 5, 8].forEach(j => { const d = asDate(r[j]); if (d !== r[j]) { r[j] = d; ch = true; } }); const b = isTrue(r[6]); if (r[6] !== b) { r[6] = b; ch = true; } const p = isTrue(r[9]); if (r[9] !== p) { r[9] = p; ch = true; } });
  if (ch) rg.setValues(v);
  s.getRange(2, 5, n, 2).setNumberFormat(DF);
  s.getRange(2, 9, n, 1).setNumberFormat(DF);
  s.getRange(2, 7, n, 1).insertCheckboxes();
  s.getRange(2, 10, n, 1).insertCheckboxes();
}
function usersRows(s) {
  usersImport(s);
  usersFix(s);
  const v = s.getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '').map(r => ({ tel: String(r[0]), ad: r[1], soyad: r[2], rol: r[3], qeyd: ms(r[4]), son: ms(r[5]), blok: isTrue(r[6]), sayt: r[7], elaqe: ms(r[8]), prem: isTrue(r[9]) }));
}
function userRow(s, tel) {
  const n = s.getLastRow();
  if (n < 2) return 0;
  const v = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(tel)) return i + 2;
  return 0;
}
function userPing(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) {
    s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', new Date(+o.qeyd || now), new Date(now), false, o.sayt || '', '', false]);
    const lr = s.getLastRow();
    s.getRange(lr, 5, 1, 2).setNumberFormat(DF);
    s.getRange(lr, 7).insertCheckboxes();
    s.getRange(lr, 10).insertCheckboxes();
    return { ok: true, blok: false, prem: false };
  }
  const val = s.getRange(r, 1, 1, 10).getValues()[0];
  s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
  s.getRange(r, 6).setValue(new Date(now)).setNumberFormat(DF);
  return { ok: true, blok: isTrue(val[6]), prem: isTrue(val[9]) };
}
function userSet(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) return { ok: false, error: 'tel' };
  if (o.a === 'block') { s.getRange(r, 7).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  if (o.a === 'premium') { s.getRange(r, 10).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  s.getRange(r, 9).setValue(o.v ? new Date(now) : '').setNumberFormat(DF);
  return { ok: true, t: o.v ? now : '' };
}

/* Müəllim düzəlişləri (silinən / düzəldilən tapşırıqlar): Config vərəqinin A1 xanasında bir JSON */
function cfgSheet() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('Config') || ss.insertSheet('Config');
}

/* ✍️ Rəy və təkliflər: "Rəylər" vərəqi */
const FL = ['Tarix', 'Telefon', 'Ad', 'Soyad', 'Rol', 'Rəy'];
function fbSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Rəylər');
  if (!s) { s = ss.insertSheet('Rəylər'); s.appendRow(FL); s.getRange(1, 1, 1, FL.length).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(6, 420); }
  return s;
}
function fbRows() {
  const s = fbSheet();
  if (s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, FL.length).getValues().map(r => ({ at: ms(r[0]), tel: String(r[1]), ad: r[2], soyad: r[3], rol: r[4], x: String(r[5]) }));
}

/* Oxu: ?a=cfg (hamı üçün) · ?a=users&key=… · ?a=plist&key=… (müəllim) */
function doGet(e) {
  const p = e.parameter || {};
  if (p.a === 'cfg') { const v = cfgSheet().getRange('A1').getValue(); let st = null; try { st = v ? JSON.parse(v) : null; } catch (e) {} return out({ state: st }); }
  if (p.key !== KEY) return out({ error: 'key' });
  if (p.a === 'users') return out({ rows: usersRows(usersSheet()) });
  if (p.a === 'fbs') return out({ rows: fbRows() });
  if (p.a === 'plist') return out({ tels: usersRows(usersSheet()).filter(u => u.prem).map(u => u.tel) });
  return out({ error: 'a' });
}

/* Yaz: ping (qeydiyyat + son giriş + blok yoxlaması), pq (premium yoxlaması), block, elaqe (ödəniş), premium, cfg (müəllim) */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents || '{}');
    const now = Date.now();
    if (o.a === 'ping') return out(userPing(o, now));
    if (o.a === 'pq') { /* şagirdin premium yoxlaması (Missions) */
      const s = usersSheet(), r = userRow(s, o.tel);
      return out({ ok: true, prem: r ? isTrue(s.getRange(r, 10).getValue()) : false });
    }
    if (o.a === 'fb') { /* şagirdin rəyi (açar lazım deyil) */
      const x = String(o.x || '').trim().slice(0, 2000);
      if (!x) return out({ ok: false, error: 'x' });
      const s = fbSheet();
      s.appendRow([new Date(+o.at || now), "'" + String(o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', x]);
      s.getRange(s.getLastRow(), 1).setNumberFormat(DF);
      return out({ ok: true });
    }
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'block' || o.a === 'elaqe' || o.a === 'premium') return out(userSet(o, now));
    if (o.a === 'cfg') { cfgSheet().getRange('A1').setNumberFormat('@').setValue(JSON.stringify(o.state || {})); return out({ ok: true }); }
    return out({ ok: false, error: 'a' });
  } finally {
    lock.releaseLock();
  }
}

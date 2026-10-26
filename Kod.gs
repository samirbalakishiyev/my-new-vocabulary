/* My New Vocabulary · Google Cədvəl skripti (Apps Script)
   Bu cədvəldə yalnız istifadəçilər (qeydiyyat, son giriş, blok, ödəniş, 👑 premium) və müəllim düzəlişləri saxlanılır.
   Nəticələr yalnız QRUPA yazılmış şagirdlər üçün yazılır: "Nəticələr" (ümumi) və "Qrup A1" kimi qrup vərəqləri. Premium yalnız bu saytındır (English 5–11-dən asılı deyil).
   Köhnə "Qeydiyyat" vərəqindəki istifadəçilər avtomatik "Users" vərəqinə köçürülür; köhnə vərəqlər silinmir.
   Yeniləmək: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (link dəyişmir). */
const KEY = 'BURAYA-OZ-ACARINIZI-YAZIN';
const ADMIN_PASS = 'BURAYA-PANEL-SIFRESI'; /* 🔐 Müəllim panelinin şifrəsi — özünüz yazın, heç kimə deməyin */
const SKEY = 'BURAYA-LUGET-SIFRESI'; /* 📚 PDF lüğətlərin şifrəsi — heç kimə verməyin */
const SLUG_IDS = ['g5', 'g6', 'g7', 'g8', 'g9', 'g10', 'g11'];
/* "g5,g8" → ['g5','g8'] (köhnə TRUE = hamısı) */
function slugList(v) {
  if (v === true || String(v).toUpperCase() === 'TRUE') return SLUG_IDS.slice();
  return String(v || '').split(/[,\s]+/).filter(x => SLUG_IDS.includes(x));
}
/* hər kitabın öz şifrəsi: HMAC-SHA256(SKEY, kitab) */
function slugKeys(ids) {
  const k = Utilities.base64Decode(SKEY), o = {};
  ids.forEach(id => { o[id] = Utilities.base64Encode(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, Utilities.newBlob(id).getBytes(), k)); });
  return o;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Users vərəqi (oxunaqlı başlıqlar, tarixlər, blok qutucuqları) ---------- */
const UL = ['Telefon', 'Ad', 'Soyad', 'Rol', 'Qeydiyyat tarixi', 'Son giriş', 'Bloklanıb', 'Sayt', 'Əlaqə saxlanılıb', 'Premium', 'PDF lüğət', 'Qrup', 'Profil adı', 'Profil sinfi', 'Düzəldilmiş ad', 'Düzgün nömrə'];
/* 👥 qruplar (müəllim panelində şagirdə verilir) */
const GROUPS = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2', 'N 11', 'N 14', 'N 16.1', 'N 16.2', 'N 18.1', 'N 18.2'];
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
    k => k === 'prem' || k.startsWith('premium'),
    k => k === 'slug' || k.startsWith('pdf'),
    k => k === 'qrup' || k === 'group' || k === 'grp',
    k => k === 'profil adı' || k === 'pn',
    k => k === 'profil sinfi' || k === 'pk',
    k => k.startsWith('düzəldilmiş') || k === 'nm',
    k => k.startsWith('düzgün nömrə') || k === 'tel2'];
  return T.map(f => h.findIndex(f));
}
/* başqa vərəqlərdə qalmış köhnə istifadəçiləri Users-ə köçürür (təkrarlamadan) */
function usersImport(s) {
  const ss = SpreadsheetApp.getActive();
  const have = new Set(s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const add = [];
  ss.getSheets().forEach(sh => {
    const nm = sh.getName();
    if (/^users/i.test(nm) || /config|results|nəticə|qrup|rəy/i.test(nm) || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
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
    if (n > 0) { /* 🛟 sütunları düzməzdən əvvəl HƏR DƏFƏ tarixli nüsxə */ const nm = 'Users yedək ' + Utilities.formatDate(new Date(), 'Asia/Baku', 'dd.MM HH:mm:ss'); if (!ss.getSheetByName(nm)) s.copyTo(ss).setName(nm); }
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
  /* yalnız dəyişən xanaları yazır (bütün cədvəli yenidən yazmır — eyni anda edilən blok/ödəniş itməsin) */
  v.forEach((r, i) => {
    [4, 5, 8].forEach(j => { const d = asDate(r[j]); if (d !== r[j]) s.getRange(i + 2, j + 1).setValue(d); });
    [6, 9].forEach(j => { const b = isTrue(r[j]); if (r[j] !== b && r[j] !== '') s.getRange(i + 2, j + 1).setValue(b); });
  });
  s.getRange(2, 5, n, 2).setNumberFormat(DF);
  s.getRange(2, 9, n, 1).setNumberFormat(DF);
  s.getRange(2, 7, n, 1).insertCheckboxes();
  s.getRange(2, 10, n, 1).insertCheckboxes();
  s.getRange(2, 11, n, 6).removeCheckboxes().setNumberFormat('@');
  /* köhnə qrup adları: L 11 → N 11, L 14 → N 14 */
  const gr = s.getRange(2, 12, n, 1), gv = gr.getValues(); let gch = false;
  gv.forEach(r => { const m = String(r[0]).match(/^L (11|14)$/); if (m) { r[0] = 'N ' + m[1]; gch = true; } });
  if (gch) gr.setValues(gv);
}
function usersRows(s) {
  usersImport(s);
  usersFix(s);
  const v = s.getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '').map(r => ({ tel: String(r[0]), ad: r[1], soyad: r[2], rol: r[3], qeyd: ms(r[4]), son: ms(r[5]), blok: isTrue(r[6]), sayt: r[7], elaqe: ms(r[8]), prem: isTrue(r[9]), slug: slugList(r[10]), grp: String(r[11] || ''), pn: String(r[12] || ''), pk: String(r[13] || ''), nm: String(r[14] || ''), tel2: String(r[15] || '') }));
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
    s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', new Date(+o.qeyd || now), new Date(now), false, o.sayt || '', '', false, '', '', String(o.pn || '').slice(0, 80), String(o.pk || '').slice(0, 40), '', '']);
    s.getRange(s.getLastRow(), 11, 1, 6).setNumberFormat('@');
    const lr = s.getLastRow();
    s.getRange(lr, 5, 1, 2).setNumberFormat(DF);
    s.getRange(lr, 7).insertCheckboxes();
    s.getRange(lr, 10).insertCheckboxes();
    return { ok: true, blok: false, prem: false, grp: '', nm: '' };
  }
  const val = s.getRange(r, 1, 1, 15).getValues()[0];
  s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
  /* 📋 şagirdin "Nəticəni müəllimə göndər" xanasındakı ad və sinif (qrupu tapmaq üçün) */
  if (o.pn !== undefined || o.pk !== undefined) s.getRange(r, 13, 1, 2).setNumberFormat('@').setValues([[String(o.pn || val[12] || '').slice(0, 80), String(o.pk || val[13] || '').slice(0, 40)]]);
  s.getRange(r, 6).setValue(new Date(now)).setNumberFormat(DF);
  return { ok: true, blok: isTrue(val[6]), prem: isTrue(val[9]), grp: String(val[11] || ''), nm: String(val[14] || '') };
}
function userSet(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) return { ok: false, error: 'tel' };
  if (o.a === 'block') { s.getRange(r, 7).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  if (o.a === 'premium') { s.getRange(r, 10).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  if (o.a === 'ren') { /* ✏️ müəllimin düzəltdiyi ad və nömrə ('' = qeydiyyatdakı) — şagirdin girişi dəyişmir */
    const t = String(o.v || '').trim().replace(/\s+/g, ' ').slice(0, 80);
    const t2 = String(o.t2 || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '').slice(0, 9);
    s.getRange(r, 15, 1, 2).setNumberFormat('@').setValues([[t, t2]]);
    return { ok: true, nm: t, tel2: t2 };
  }
  if (o.a === 'grp') { /* 👥 qrup: '' = qrupsuz */
    const g = String(o.v || '');
    if (g && !GROUPS.includes(g)) return { ok: false, error: 'g' };
    s.getRange(r, 12).setNumberFormat('@').setValue(g);
    return { ok: true, grp: g };
  }
  if (o.a === 'slug') { /* 📚 hər kitab ayrıca: "g5,g8" */
    const c = s.getRange(r, 11), set = new Set(slugList(c.getValue()));
    const b = String(o.b || '');
    if (!SLUG_IDS.includes(b)) return { ok: false, error: 'b' };
    if (o.v) set.add(b); else set.delete(b);
    const list = SLUG_IDS.filter(x => set.has(x));
    c.setNumberFormat('@').setValue(list.join(','));
    return { ok: true, slug: list };
  }
  s.getRange(r, 9).setValue(o.v ? new Date(now) : '').setNumberFormat(DF);
  return { ok: true, t: o.v ? now : '' };
}

/* 📊 Şagird nəticələri: "Nəticələr" (hamısı) + "Qrup …" (hər qrup ayrıca). Açar sütunu (telefon|vaxt) təkrarı önləyir. */
const RL = ['Tarix', 'Qrup', 'Telefon', 'Ad', 'Soyad', 'Bölmə', 'Düz', 'Sual', 'Faiz', 'Səhvlər', 'Açar'];
function resSheet(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(RL); s.getRange(1, 1, 1, RL.length).setFontWeight('bold'); s.setFrozenRows(1); s.setColumnWidth(6, 320); s.setColumnWidth(10, 260); }
  return s;
}
function resAdd(o, now) {
  const us = usersSheet(), r = userRow(us, o.tel);
  if (!r) return { ok: true, grp: '' };
  const v = us.getRange(r, 1, 1, 15).getValues()[0], g = String(v[11] || '');
  const nm = String(v[14] || '').trim(), ad = nm ? nm.split(' ')[0] : v[1], soyad = nm ? nm.split(' ').slice(1).join(' ') : v[2];
  if (!g || isTrue(v[6])) return { ok: true, grp: g }; /* qrupsuz və ya bloklanmış: yazılmır */
  const all = resSheet('Nəticələr'), gs = resSheet('Qrup ' + g);
  const list = (Array.isArray(o.list) ? o.list : []).slice(0, 50);
  const rows = [];
  list.forEach(e => {
    const at = +e.at || now, key = String(o.tel) + '|' + at;
    if (all.getLastRow() > 1 && all.getRange(2, 11, all.getLastRow() - 1, 1).createTextFinder(key).matchEntireCell(true).findNext()) return;
    rows.push([new Date(at), g, "'" + String(o.tel), ad, soyad, String(e.sec || '').slice(0, 200), +e.r || 0, +e.n || 0, +e.p || 0, (Array.isArray(e.w) ? e.w : []).slice(0, 30).join(', '), key]);
  });
  if (rows.length) {
    [all, gs].forEach(s => { const lr = s.getLastRow() + 1; s.getRange(lr, 1, rows.length, RL.length).setValues(rows); s.getRange(lr, 1, rows.length, 1).setNumberFormat(DF); });
  }
  return { ok: true, grp: g, added: rows.length };
}
function resRows() {
  const s = SpreadsheetApp.getActive().getSheetByName('Nəticələr');
  if (!s || s.getLastRow() < 2) return [];
  const n = s.getLastRow() - 1, from = Math.max(0, n - 5000);
  return s.getRange(2 + from, 1, n - from, RL.length).getValues().map(r => ({ at: ms(r[0]), grp: String(r[1]), tel: String(r[2]), ad: r[3], soyad: r[4], sec: String(r[5]), r: r[6], n: r[7], p: r[8], w: String(r[9] || '') }));
}

/* Müəllim düzəlişləri (silinən / düzəldilən tapşırıqlar): Config vərəqinin A1 xanasında bir JSON */
function cfgSheet() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName('Config') || ss.insertSheet('Config');
}

/* ✍️ Rəy və təkliflər: "Rəylər" vərəqi. Status: Yeni / Oxundu / Cavablandı / Arxiv.
   "Şagird gizlədib" = şagird öz tərəfində silib (müəllim yenə görür). ID = telefon|vaxt. */
const FL = ['Tarix', 'Telefon', 'Ad', 'Soyad', 'Rol', 'Rəy', 'Status', 'Cavab', 'Cavab tarixi', 'Şagird gizlədib', 'ID'];
const FST = ['Yeni', 'Oxundu', 'Cavablandı', 'Arxiv'];
function fbSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Rəylər');
  if (!s) { s = ss.insertSheet('Rəylər'); s.setFrozenRows(1); s.setColumnWidth(6, 420); s.setColumnWidth(8, 360); }
  const w = Math.max(s.getLastColumn(), 1), hv = s.getRange(1, 1, 1, w).getValues()[0].map(String);
  if (hv.slice(0, FL.length).join('|') !== FL.join('|')) s.getRange(1, 1, 1, FL.length).setValues([FL]).setFontWeight('bold'); /* köhnə 6 sütun eynidir, yenilər sağa əlavə olunur */
  return s;
}
function fbAll() {
  const s = fbSheet();
  if (s.getLastRow() < 2) return { s, v: [] };
  const rg = s.getRange(2, 1, s.getLastRow() - 1, FL.length), v = rg.getValues();
  let ch = false;
  v.forEach(r => { if (!r[10] && r[1] !== '') { r[10] = String(r[1]).replace(/^'/, '') + '|' + ms(r[0]); ch = true; } if (!r[6] && r[1] !== '') { r[6] = 'Yeni'; ch = true; } });
  if (ch) { s.getRange(2, 11, v.length, 1).setNumberFormat('@'); rg.setValues(v); }
  return { s, v };
}
const fbObj = r => ({ id: String(r[10]), at: ms(r[0]), tel: String(r[1]), ad: r[2], soyad: r[3], rol: r[4], x: String(r[5]), st: String(r[6] || 'Yeni'), rep: String(r[7] || ''), rat: ms(r[8]), hid: isTrue(r[9]) });
function fbRows() { return fbAll().v.filter(r => r[1] !== '').map(fbObj); }
/* müəllim: status / cavab / silmə */
function fbSet(o, now) {
  const { s, v } = fbAll(), i = v.findIndex(r => String(r[10]) === String(o.id));
  if (i < 0) return { ok: false, error: 'id' };
  const row = i + 2;
  if (o.del) { s.deleteRow(row); return { ok: true, del: true }; }
  if (typeof o.rep === 'string') {
    const t = o.rep.trim().slice(0, 2000);
    s.getRange(row, 7, 1, 3).setValues([[t ? 'Cavablandı' : 'Oxundu', t, t ? new Date(now) : '']]);
    if (t) s.getRange(row, 9).setNumberFormat(DF);
    return { ok: true, st: t ? 'Cavablandı' : 'Oxundu', rat: t ? now : '' };
  }
  if (FST.includes(o.st)) { s.getRange(row, 7).setValue(o.st); return { ok: true, st: o.st }; }
  return { ok: false, error: 'a' };
}

/* Oxu: ?a=cfg (hamı üçün) · ?a=users&key=… · ?a=plist&key=… (müəllim) */
function doGet(e) {
  const lock = LockService.getScriptLock(); /* oxuma da növbə ilə: eyni anda iki sorğu cədvəli düzəndə qarışıqlıq olmasın */
  lock.waitLock(10000);
  try { return doGet_(e); } finally { lock.releaseLock(); }
}
function doGet_(e) {
  const p = e.parameter || {};
  if (p.a === 'cfg') { const v = cfgSheet().getRange('A1').getValue(); let st = null; try { st = v ? JSON.parse(v) : null; } catch (e) {} return out({ state: st }); }
  if (p.key !== KEY) return out({ error: 'key' });
  if (p.a === 'users') return out({ rows: usersRows(usersSheet()) });
  if (p.a === 'fbs') return out({ rows: fbRows() });
  if (p.a === 'res') return out({ rows: resRows() });
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
    if (o.a === 'tlogin') { /* 🔐 müəllim girişi: 5 səhv cəhddən sonra 15 dəqiqə bağlanır */
      const c = CacheService.getScriptCache(), n = +(c.get('tl_fail') || 0);
      if (ADMIN_PASS === 'BURAYA-PANEL-SIFRESI' || !ADMIN_PASS) return out({ ok: false, error: 'nopass' });
      if (n >= 5) return out({ ok: false, error: 'wait' });
      if (String(o.p || '') !== ADMIN_PASS) { c.put('tl_fail', String(n + 1), 900); return out({ ok: false, error: 'pass', left: 4 - n }); }
      c.remove('tl_fail');
      return out({ ok: true, key: KEY });
    }
    if (o.a === 'ping') return out(userPing(o, now));
    if (o.a === 'res') return out(resAdd(o, now)); /* 📊 nəticə: yalnız qrupa yazılmış şagirdin nəticəsi saxlanılır */
    if (o.a === 'pq') { /* şagirdin premium yoxlaması (Missions) */
      const s = usersSheet(), r = userRow(s, o.tel);
      return out({ ok: true, prem: r ? isTrue(s.getRange(r, 10).getValue()) : false });
    }
    if (o.a === 'fb') { /* şagirdin rəyi (açar lazım deyil) */
      const x = String(o.x || '').trim().slice(0, 2000);
      if (!x) return out({ ok: false, error: 'x' });
      const s = fbSheet(), at = +o.at || now, id = String(o.tel || '') + '|' + at;
      if (s.getLastRow() > 1 && s.getRange(2, 11, s.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext()) return out({ ok: true, id });
      s.appendRow([new Date(at), "'" + String(o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', x, 'Yeni', '', '', false, "'" + id]);
      s.getRange(s.getLastRow(), 1).setNumberFormat(DF);
      return out({ ok: true, id });
    }
    if (o.a === 'fbmine') { /* şagird: öz rəyləri və müəllimin cavabları (gizlətdikləri görünmür) */
      const tel = String(o.tel || '');
      if (!tel) return out({ ok: true, rows: [] });
      return out({ ok: true, rows: fbRows().filter(f => f.tel === tel && !f.hid).map(f => ({ id: f.id, at: f.at, x: f.x, rep: f.rep, rat: f.rat })) });
    }
    if (o.a === 'fbhide') { /* şagird öz tərəfində silir: müəllimdə qalır */
      const { s, v } = fbAll(), i = v.findIndex(r => String(r[10]) === String(o.id) && String(r[1]) === String(o.tel || ''));
      if (i < 0) return out({ ok: false, error: 'id' });
      s.getRange(i + 2, 10).insertCheckboxes().setValue(true);
      return out({ ok: true });
    }
    if (o.a === 'skey') { /* 📚 məktəb lüğətləri: şifrə yalnız kilidi açılmış nömrəyə (və müəllimə) verilir */
      if (o.key === KEY) return out({ ok: true, keys: slugKeys(SLUG_IDS) });
      const s = usersSheet(), r = userRow(s, o.tel);
      if (!r || isTrue(s.getRange(r, 7).getValue())) return out({ ok: true, keys: {} });
      return out({ ok: true, keys: slugKeys(slugList(s.getRange(r, 11).getValue())) });
    }
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'block' || o.a === 'elaqe' || o.a === 'premium' || o.a === 'slug' || o.a === 'grp' || o.a === 'ren') return out(userSet(o, now));
    if (o.a === 'fbset') return out(fbSet(o, now));
    if (o.a === 'cfg') { cfgSheet().getRange('A1').setNumberFormat('@').setValue(JSON.stringify(o.state || {})); return out({ ok: true }); }
    return out({ ok: false, error: 'a' });
  } finally {
    lock.releaseLock();
  }
}

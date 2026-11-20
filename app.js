/* MESco CRM Web — rozszerzenie CRM (Power Apps) o analizy i dashboard na żywo.
   Dane: SharePoint CRM System POC przez Microsoft Graph (logowanie M365 / MSAL). */
(() => {
'use strict';
const CFG = window.CRM_CONFIG;
const G = 'https://graph.microsoft.com/v1.0';

/* ================= ikony (Lucide, jak w Power Apps) ================= */
const I = {
  HR:"<path d='M16 2v2'/><path d='M17.915 22a6 6 0 0 0-12 0'/><path d='M8 2v2'/><circle cx='12' cy='12' r='4'/><rect x='3' y='4' width='18' height='18' rx='2'/>",
  FIRMY:"<path d='M6 22V4a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v18Z'/><path d='M6 12H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h2'/><path d='M18 9h2a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2h-2'/><path d='M10 6h4'/><path d='M10 10h4'/><path d='M10 14h4'/><path d='M10 18h4'/>",
  KONTAKTY:"<path d='M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2'/><circle cx='9' cy='7' r='4'/><path d='M22 21v-2a4 4 0 0 0-3-3.87'/><path d='M16 3.13a4 4 0 0 1 0 7.75'/>",
  KALENDARZ:"<rect width='18' height='18' x='3' y='4' rx='2'/><path d='M16 2v4'/><path d='M8 2v4'/><path d='M3 10h18'/>",
  CENNIK:"<path d='M4 10h12'/><path d='M4 14h9'/><path d='M19 6a7.7 7.7 0 0 0-5.2-2A7.9 7.9 0 0 0 6 12c0 4.4 3.5 8 7.8 8 2 0 3.8-.8 5.2-2'/>",
  SZKOLENIA:"<path d='M21.42 10.922a1 1 0 0 0-.019-1.838L12.83 5.18a2 2 0 0 0-1.66 0L2.6 9.08a1 1 0 0 0 0 1.832l8.57 3.908a2 2 0 0 0 1.66 0z'/><path d='M22 10v6'/><path d='M6 12.5V16a6 3 0 0 0 12 0v-3.5'/>",
  OPPSY:"<polyline points='22 7 13.5 15.5 8.5 10.5 2 17'/><polyline points='16 7 22 7 22 13'/>",
  TECSY:"<path d='M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z'/><path d='m9 12 2 2 4-4'/>",
  PROSPECTING:"<circle cx='12' cy='12' r='10'/><circle cx='12' cy='12' r='6'/><circle cx='12' cy='12' r='2'/>",
  SUPPORTY:"<path d='M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z'/>",
  ZADANIA:"<path d='m3 17 2 2 4-4'/><path d='m3 7 2 2 4-4'/><path d='M13 6h8'/><path d='M13 12h8'/><path d='M13 18h8'/>",
  MARKETING:"<path d='m3 11 18-5v12L3 14v-3z'/><path d='M11.6 16.8a3 3 0 1 1-5.8-1.6'/>",
  DASHBOARD:"<rect width='7' height='9' x='3' y='3' rx='1'/><rect width='7' height='5' x='14' y='3' rx='1'/><rect width='7' height='9' x='14' y='12' rx='1'/><rect width='7' height='5' x='3' y='16' rx='1'/>",
  ADMIN:"<path d='M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z'/><line x1='4' x2='4' y1='22' y2='15'/>",
  USTAWIENIA:"<path d='M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z'/><circle cx='12' cy='12' r='3'/>",
  DEMO:"<path d='M2.586 17.414A2 2 0 0 0 2 18.828V21a1 1 0 0 0 1 1h3a1 1 0 0 0 1-1v-1a1 1 0 0 1 1-1h1a1 1 0 0 0 1-1v-1a1 1 0 0 1 1-1h.172a2 2 0 0 0 1.414-.586l.814-.814a6.5 6.5 0 1 0-4-4z'/><circle cx='16.5' cy='7.5' r='.5' fill='currentColor'/>",
  clock:"<circle cx='12' cy='12' r='10'/><polyline points='12 6 12 12 16 14'/>",
  activity:"<path d='M22 12h-4l-3 9L9 3l-3 9H2'/>",
  bars:"<line x1='12' x2='12' y1='20' y2='10'/><line x1='18' x2='18' y1='20' y2='4'/><line x1='6' x2='6' y1='20' y2='16'/>",
  list:"<path d='M3 12h.01'/><path d='M3 18h.01'/><path d='M3 6h.01'/><path d='M8 12h13'/><path d='M8 18h13'/><path d='M8 6h13'/>",
  phone:"<path d='M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z'/>",
  users:"<path d='M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2'/><circle cx='9' cy='7' r='4'/><path d='M22 21v-2a4 4 0 0 0-3-3.87'/><path d='M16 3.13a4 4 0 0 1 0 7.75'/>",
  box:"<path d='M11 21.73a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73z'/><path d='M12 22V12'/><path d='m3.3 7 7.703 4.734a2 2 0 0 0 1.994 0L20.7 7'/><path d='m7.5 4.27 9 5.15'/>",
  target:"<circle cx='12' cy='12' r='10'/><circle cx='12' cy='12' r='6'/><circle cx='12' cy='12' r='2'/>",
  note:"<path d='M16 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V8Z'/><path d='M15 3v4a2 2 0 0 0 2 2h4'/>",
  sort:"<path d='m3 16 4 4 4-4'/><path d='M7 20V4'/><path d='m21 8-4-4-4 4'/><path d='M17 4v16'/>",
  ext:"<path d='M15 3h6v6'/><path d='M10 14 21 3'/><path d='M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6'/>",
  refresh:"<path d='M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8'/><path d='M21 3v5h-5'/><path d='M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16'/><path d='M8 16H3v5'/>",
  x:"<path d='M18 6 6 18'/><path d='m6 6 12 12'/>",
  alert:"<path d='m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3'/><path d='M12 9v4'/><path d='M12 17h.01'/>",
  euro:"<path d='M4 10h12'/><path d='M4 14h9'/><path d='M19 6a7.7 7.7 0 0 0-5.2-2A7.9 7.9 0 0 0 6 12c0 4.4 3.5 8 7.8 8 2 0 3.8-.8 5.2-2'/>",
  check:"<path d='M20 6 9 17l-5-5'/>",
  tag:"<path d='M12.586 2.586A2 2 0 0 0 11.172 2H4a2 2 0 0 0-2 2v7.172a2 2 0 0 0 .586 1.414l8.704 8.704a2.426 2.426 0 0 0 3.42 0l6.58-6.58a2.426 2.426 0 0 0 0-3.42z'/><circle cx='7.5' cy='7.5' r='.5' fill='currentColor'/>",
};
const svg = (k, color, sw = 2) => `<svg viewBox="0 0 24 24" fill="none" stroke="${color || 'currentColor'}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round">${I[k] || ''}</svg>`;

/* ================= narzędzia ================= */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => n == null || isNaN(n) ? '—' : Math.round(n).toLocaleString('pl-PL');
const eur = n => fmt(n) + ' €';
const fmtK = n => n >= 1e6 ? (n/1e6).toLocaleString('pl-PL',{maximumFractionDigits:2})+' mln' : n >= 1e4 ? Math.round(n/1e3).toLocaleString('pl-PL')+' tys.' : fmt(n);
const sum = (a, f) => a.reduce((s, x) => s + (+f(x) || 0), 0);
const num = v => { if (v == null || v === '') return null; const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/\s/g,'').replace(',','.')); return isNaN(n) ? null : n; };
const toD = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
const ds = d => d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : '';
const dpl = d => d ? `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}` : '—';
const TODAY = (() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); })();
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const daysTo = d => d ? Math.round((d - TODAY) / 864e5) : null;
const initials = s => String(s || '?').split(/\s+/).filter(Boolean).map(w => w[0]).join('').slice(0, 2).toUpperCase();
const lc = s => String(s || '').toLowerCase().trim();
// rok obrotowy Ansys XI–X: Q1 = lis–sty
const fyQ = d => { if (!d) return null; const m = d.getMonth()+1, y = d.getFullYear(); const fy = m >= 11 ? y+1 : y; const q = (m >= 11 || m === 1) ? 1 : m <= 4 ? 2 : m <= 7 ? 3 : 4; return { fy, q, key: fy*10+q, label: `Q${q}-FY${String(fy).slice(2)}` }; };
const fyBase = () => TODAY.getMonth()+1 >= 11 ? TODAY.getFullYear() : TODAY.getFullYear()-1; // jak w Power Apps: b
const qBounds = (b, q) => [[new Date(b,10,1), new Date(b+1,0,31)], [new Date(b+1,1,1), new Date(b+1,3,30)], [new Date(b+1,4,1), new Date(b+1,6,31)], [new Date(b+1,7,1), new Date(b+1,9,31)]][q-1];
const qRange = (fy, q) => { const [a, b] = qBounds(fy-1, q); const mn = ['sty','lut','mar','kwi','maj','cze','lip','sie','wrz','paź','lis','gru']; return `${mn[a.getMonth()]} ${a.getFullYear()} – ${mn[b.getMonth()]} ${b.getFullYear()}`; };
const toast = (msg, kind = '') => { const t = $('#toast'); t.textContent = msg; t.className = 'on ' + kind; clearTimeout(t._h); t._h = setTimeout(() => t.className = '', 2600); };
const tip = $('#tip');
const showTip = (e, html) => { tip.innerHTML = html; tip.style.display = 'block'; tip.style.left = Math.min(e.clientX+14, innerWidth - tip.offsetWidth - 8) + 'px'; tip.style.top = Math.min(e.clientY+14, innerHeight - tip.offsetHeight - 8) + 'px'; };
const hideTip = () => tip.style.display = 'none';
const loader = (msg) => { const l = $('#loader'); if (msg === false) l.classList.add('hide'); else { l.classList.remove('hide'); $('#loadMsg').textContent = msg; } };

/* ================= auth + Graph ================= */
const MOCK = window.CRM_MOCK || null;
let msal, account;
async function token() {
  const req = { scopes: CFG.scopes, account };
  try { return (await msal.acquireTokenSilent(req)).accessToken; }
  catch (e) { await msal.acquireTokenRedirect(req); throw e; }
}
async function gfetch(url, opt = {}) {
  const t = await token();
  const r = await fetch(url.startsWith('http') ? url : G + url, { ...opt, headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json', ...(opt.headers || {}) } });
  if (!r.ok) { const tx = await r.text(); throw new Error(`Graph ${r.status}: ${tx.slice(0, 300)}`); }
  return r.status === 204 ? null : r.json();
}
let SITE = null; const LIST = {};
async function initSite() {
  if (MOCK) return;
  SITE = (await gfetch(`/sites/${CFG.siteHost}:${CFG.sitePath}?$select=id`)).id;
  let url = `/sites/${SITE}/lists?$select=id,displayName,name,system&$top=500`;
  while (url) { const j = await gfetch(url); j.value.forEach(l => { LIST[l.displayName] = l.id; LIST['@' + l.name] = l.id; }); url = j['@odata.nextLink']; }
}
async function items(list, fields) {
  if (MOCK) return (MOCK.lists[list] || []).map(f => ({ ...f }));
  const id = LIST[list]; if (!id) { console.warn('Brak listy', list); return []; }
  let url = `/sites/${SITE}/lists/${id}/items?$expand=fields($select=${fields.join(',')})&$top=999`;
  const out = [];
  while (url) { const j = await gfetch(url); j.value.forEach(it => out.push({ ...it.fields, id: +it.id })); url = j['@odata.nextLink']; }
  return out;
}
async function patchItem(list, id, fields) {
  if (MOCK) { const it = MOCK.lists[list].find(x => x.id === id); Object.assign(it, fields); return; }
  await gfetch(`/sites/${SITE}/lists/${LIST[list]}/items/${id}/fields`, { method: 'PATCH', body: JSON.stringify(fields) });
}
async function createItem(list, fields) {
  if (MOCK) { const id = Math.max(0, ...MOCK.lists[list].map(x => x.id)) + 1; MOCK.lists[list].push({ ...fields, id }); return { id }; }
  const r = await gfetch(`/sites/${SITE}/lists/${LIST[list]}/items`, { method: 'POST', body: JSON.stringify({ fields }) });
  return { id: +r.id };
}
const isoNoon = d => d ? new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12).toISOString() : null;

/* ================= dane ================= */
const DB = { loaded: null };
const ME = { name: '', mail: '', staff: false };
const F = {
  Companies: ['CompaniesGUID','Nazwa','Pe_x0142_nanazwa','Status','Status1','Bran_x017c_a','Creator','TECSh','PreSalesh','WartoscTECS','MarzaTECS','SzkoleniaGratis','SzkolenieGratis','Utworzoneprzez_x002f_OpiekunLookupId','NumerAnsys'],
  Notes: ['Firma','Osoba','TypeLookupId','Rodzaj1','Totatka','Data','Follow_up','FollowUpStatus','Creator','Telefon','LiczbaPolaczen','Opps_ID','IdentyfikatorCompanies','IdentyfikatorContact','Etap','EtapProspecting','DodajFollowUp','Created'],
  Opps: ['Nazwaszansysprzeda_x017c_y','Firma','Status1','Etap1','Etap','Warto_x015b__x0107_wEUR','Spodziewanezamkni_x0119_cie','ClosedDate','Dataotwarcia','Rodzajbiznesu','TypBiznesu','IdentyfikatorCompanies','OPPSGUID','Creator','OpiekunOppsaLookupId'],
  Oferty: ['IdentyfikatorOpps','Suma'],
  Opps_TECS: ['Numer_klienta','Nazwaszansysprzeda_x017c_y_TECS','Pocz_x0105_tek_TECS','Koniec_TECS','OpiekunLookupId','Warto_x015b__x0107_','FirmaID','Etap1','TP','IdentyfikatorCompanies','TECSGUID','Demo','DemoOffice','DemoAdmLookupId','Data_Follow_Up','TECScreator'],
  Oferty_TECS: ['ProducID','Cena_TECS','Liczba_licencji','SUMA','Rabat','IdentyfikatorTECS','TP','TPValue','Demo','FirmaIDofertaTECS','Pocz_x0105_tek_TECS','Koniec_TECS'],
  Cennik_CRM: ['Title','field_2','field_3'],
  Support: ['Data','Czas','IdentyfikatorCompanies'],
  UczestnicySzkolen: ['GUIDFirmy','DataSzkolenia','WartoscDarmowegoSzkolenia'],
  Szkolenia: ['GUIDFirmy','DataSzkolenia','WartoscDarmowegoSzkolenia'],
  KursyWalut: ['Waluta','Kurs'],
  activities: ['Notes_types'],
  MEScoStaff: ['Imi_x0119_iNazwisko','Email'],
};
async function loadAll() {
  loader('Pobieram dane z SharePoint…');
  const names = Object.keys(F);
  const res = {};
  let done = 0;
  await Promise.all(names.map(async n => { res[n] = await items(n, F[n]); loader(`Pobieram dane z SharePoint… ${++done}/${names.length}`); }));
  // użytkownicy (kolumny Osoba)
  let users = [];
  try { users = MOCK ? (MOCK.users || []) : await (async () => { const id = LIST['@users'] || LIST['User Information List']; if (!id) return []; LIST['__users'] = id; return items('__users', ['Title','EMail']); })(); } catch (e) { console.warn(e); }
  build(res, users);
  DB.loaded = new Date();
}
const USERS = {};
function build(R, users) {
  users.forEach(u => USERS[u.id] = u.Title);
  const uname = id => id ? (USERS[+id] || '') : '';
  ME.staff = R.MEScoStaff.some(s => lc(s.Email) === lc(ME.mail));
  DB.kurs = num((R.KursyWalut.find(k => k.Waluta === 'EUR') || {}).Kurs) || 4.3;
  const act = {}; R.activities.forEach(a => act[a.id] = a.Notes_types);
  DB.rodzaje = R.activities.map(a => ({ id: a.id, name: a.Notes_types })).filter(a => a.name);
  // firmy
  DB.comp = R.Companies.map(c => ({
    id: c.id, guid: lc(c.CompaniesGUID), name: (c.Nazwa || c.Pe_x0142_nanazwa || '').trim() || '(bez nazwy)',
    status: c.Status || '', status1: c.Status1 || '', ac: [].concat(c.Bran_x017c_a || []).includes('Academic'), creator: c.Creator || '',
    tecsh: num(c.TECSh) || 0, presh: num(c.PreSalesh) || 0, wartTECS: num(c.WartoscTECS) || 0, marza: num(c.MarzaTECS) || 0,
    szkA: num(c.SzkoleniaGratis) || 0, szkB: num(c.SzkolenieGratis) || 0, opiekun: uname(c.Utworzoneprzez_x002f_OpiekunLookupId) || c.Creator || '',
  }));
  DB.compBy = {}; DB.comp.forEach(c => { if (c.guid) DB.compBy[c.guid] = c; });
  const cname = (guid, fb) => { const c = DB.compBy[lc(guid)]; return c ? c.name : (fb || '').trim() || '(brak firmy)'; };
  const cac = guid => { const c = DB.compBy[lc(guid)]; return c ? c.ac : false; };
  // notatki
  DB.notes = R.Notes.map(n => ({
    id: n.id, firma: (n.Firma || '').trim(), guid: lc(n.IdentyfikatorCompanies), osoba: n.Osoba || '', rodzaj: act[n.TypeLookupId] || n.Rodzaj1 || '', typeId: n.TypeLookupId,
    text: n.Totatka || '', data: toD(n.Data || n.Created), fu: toD(n.Follow_up), fuDone: !!n.FollowUpStatus, creator: n.Creator || '', tel: (n.Telefon || '').trim(),
    calls: num(n.LiczbaPolaczen) || 0, opps: lc(n.Opps_ID), contact: n.IdentyfikatorContact || '', etap: n.Etap || '', etapP: n.EtapProspecting || '', raw: n,
  }));
  const notesBy = {}; DB.notes.forEach(n => { if (n.opps) (notesBy[n.opps] = notesBy[n.opps] || []).push(n); });
  const nInfo = g => { const a = notesBy[g] || []; let last = null; a.forEach(n => { if (n.data && (!last || n.data > last)) last = n.data; }); return { n: a.length, last }; };
  // oferty sprzedaży
  const offBy = {}; R.Oferty.forEach(o => { const g = lc(o.IdentyfikatorOpps); if (!g) return; const x = offBy[g] = offBy[g] || { n: 0, s: 0 }; x.n++; x.s += num(o.Suma) || 0; });
  // szanse sprzedaży
  DB.sales = R.Opps.map(o => {
    const g = lc(o.OPPSGUID), ni = nInfo(g), of = offBy[g] || { n: 0, s: 0 };
    const close = toD(o.Spodziewanezamkni_x0119_cie), closedWin = toD(o.ClosedDate);
    const st = o.Status1 || 'Aktywny';
    const x = { id: o.id, guid: g, name: o.Nazwaszansysprzeda_x017c_y || '', compGuid: lc(o.IdentyfikatorCompanies), comp: cname(o.IdentyfikatorCompanies, o.Firma), ac: cac(o.IdentyfikatorCompanies),
      status1: st, stage: o.Etap1 || o.Etap || '', eur: num(o.Warto_x015b__x0107_wEUR) || 0, close, closedWin, openD: toD(o.Dataotwarcia), rodz: o.Rodzajbiznesu || '', typ: o.TypBiznesu || '',
      creator: o.Creator || '', owner: uname(o.OpiekunOppsaLookupId) || o.Creator || '', nOffers: of.n, offerSum: of.s, nNotes: ni.n, lastNote: ni.last };
    x.grp = st === 'Zamknięty wygrany' ? 'won' : st === 'Zamknięty przegrany' ? 'lost' : 'open';
    x.value = x.eur; x.fq = fyQ(x.grp === 'open' ? close : (closedWin || close)); x.days = daysTo(close);
    x.idle = x.lastNote ? Math.round((TODAY - x.lastNote) / 864e5) : (x.openD ? Math.round((TODAY - x.openD) / 864e5) : null);
    x.flags = [];
    if (x.grp === 'open') {
      if (!close) x.flags.push(['Brak daty zamknięcia', 'c-amb']); else if (x.days < 0) x.flags.push(['Po terminie', 'c-red']);
      if (['Oferta','Rewizja oferty','Decyzja'].includes(x.stage) && !x.nOffers) x.flags.push(['Bez oferty', 'c-red']);
      if (!x.nNotes) x.flags.push(['Bez notatek', 'c-amb']); else if (x.idle > 60) x.flags.push([`Brak notatki ${x.idle} dni`, 'c-amb']);
    }
    // marża / premie wg Power Apps
    const mz = { 'New Business': .35, 'B2TECS': .35, 'New Logo': .45 }[x.rodz] || 0;
    const mzP = { 'New Business': .35, 'B2TECS': .35, 'New Logo': .35 }[x.rodz] || 0;
    const st2 = x.rodz === 'New Logo' ? .05 : .025;
    x.marza = x.eur * mz; x.premQ = x.eur * mzP * st2 * .4; x.premR = x.eur * mzP * st2 * .6;
    return x;
  });
  // cennik
  const price = {}; R.Cennik_CRM.forEach(c => { if (c.Title) price[lc(c.Title)] = [num(c.field_2), num(c.field_3)]; });
  // pozycje ofert TECS
  DB.tecsLines = R.Oferty_TECS.map(l => ({ id: l.id, g: lc(l.IdentyfikatorTECS), prod: l.ProducID || '', qty: num(l.Liczba_licencji), price: num(l.Cena_TECS), sum: num(l.SUMA), disc: num(l.Rabat),
    tp: !!l.TP, tpValue: num(l.TPValue) || 0, demo: !!l.Demo, firma: (l.FirmaIDofertaTECS || '').trim(), start: toD(l.Pocz_x0105_tek_TECS), end: toD(l.Koniec_TECS) }));
  const linesBy = {}; DB.tecsLines.forEach(l => { if (l.g) (linesBy[l.g] = linesBy[l.g] || []).push(l); });
  // TECS
  const DONE = ['Zamówione','Wystawiona faktura','Rezygnacja z TECS'];
  DB.tecs = R.Opps_TECS.map(o => {
    const g = lc(o.TECSGUID), ls = linesBy[g] || [], ni = nInfo(g);
    const has = ls.some(l => l.sum != null), offerSum = sum(ls, l => l.sum);
    const x = { id: o.id, guid: g, compGuid: lc(o.IdentyfikatorCompanies), comp: cname(o.IdentyfikatorCompanies, o.FirmaID), ac: cac(o.IdentyfikatorCompanies),
      owner: uname(o.OpiekunLookupId) || o.TECScreator || '(brak)', stage: o.Etap1 || '', start: toD(o.Pocz_x0105_tek_TECS), end: toD(o.Koniec_TECS),
      wart: num(o.Warto_x015b__x0107_), offerSum: has ? offerSum : null, nOffers: ls.length, nNotes: ni.n, lastNote: ni.last, tp: !!o.TP, name: o.Nazwaszansysprzeda_x017c_y_TECS || '',
      demoFlag: !!o.Demo, demoOffice: !!o.DemoOffice, demoAdm: uname(o.DemoAdmLookupId), follow: toD(o.Data_Follow_Up), nrAnsys: o.Numer_klienta || '', lines: ls };
    x.value = (x.offerSum != null && x.offerSum > 0) ? x.offerSum : (x.wart || 0);
    x.valSrc = (x.offerSum != null && x.offerSum > 0) ? 'oferta' : 'pole';
    x.grp = x.stage === 'Demo' ? 'demo' : x.tp ? 'tp' : ['Przygotowanie oferty','Oferta wysłana'].includes(x.stage) ? 'open' : ['Zamówione','Wystawiona faktura'].includes(x.stage) ? 'won' : x.stage === 'Rezygnacja z TECS' ? 'lost' : 'other';
    x.fq = fyQ(x.end); x.days = daysTo(x.end);
    const active = x.grp === 'open' || (x.grp === 'tp' && !DONE.includes(x.stage));
    x.priceIssues = [];
    if (active && x.days != null && x.days >= 0) ls.forEach(l => {
      const p = price[lc(l.prod)]; const ref = p ? (x.tp ? p[1] : p[0]) : null;
      if (!p) x.priceIssues.push({ ...l, ref, kind: 'brak w cenniku' });
      else if (!l.price) x.priceIssues.push({ ...l, ref, kind: 'brak ceny w ofercie' });
      else if (ref != null && Math.abs(l.price - ref) > .5) x.priceIssues.push({ ...l, ref, diff: l.price - ref, kind: l.price < ref ? 'niższa niż cennik' : 'wyższa niż cennik' });
    });
    x.flags = [];
    if (active) {
      if (x.days != null && x.days < 0) x.flags.push(['Po terminie', 'c-red']);
      if (!x.nOffers) x.flags.push(['Bez oferty', 'c-red']);
      if (!x.nNotes) x.flags.push([x.days != null && x.days <= 90 ? 'Bez notatek — pilne' : 'Bez notatek', x.days != null && x.days <= 90 ? 'c-red' : 'c-amb']);
      if (x.stage === 'Przygotowanie oferty' && x.days != null && x.days >= 0 && x.days <= 90) x.flags.push(['Oferta do wysłania', 'c-amb']);
      if (x.priceIssues.some(p => p.diff != null)) x.flags.push(['Cena ≠ cennik', 'c-amb']);
    }
    return x;
  });
  // support na firmę
  const now = TODAY, mk = d => d.getFullYear()*12 + d.getMonth(), nowM = mk(now);
  DB.sup = {};
  R.Support.forEach(s => { const g = lc(s.IdentyfikatorCompanies), d = toD(s.Data); if (!g || !d) return;
    const x = DB.sup[g] = DB.sup[g] || { n: 0, n90: 0, n365: 0, h365: 0, last: null, m: Array(12).fill(0) };
    const age = (now - d) / 864e5; x.n++; if (age <= 90) x.n90++; if (age <= 365) { x.n365++; x.h365 += num(s.Czas) || 0; }
    if (!x.last || d > x.last) x.last = d; const mi = 11 - (nowM - mk(d)); if (mi >= 0 && mi < 12) x.m[mi]++; });
  DB.supMonths = Array.from({ length: 12 }, (_, i) => { const d = new Date(now.getFullYear(), now.getMonth() - 11 + i, 1); return ['sty','lut','mar','kwi','maj','cze','lip','sie','wrz','paź','lis','gru'][d.getMonth()] + ' ' + String(d.getFullYear()).slice(2); });
  // darmowe szkolenia (12 mies.) na firmę, PLN
  DB.szk = {}; const yAgo = addDays(TODAY, -365);
  [['A', R.UczestnicySzkolen], ['B', R.Szkolenia]].forEach(([k, arr]) => arr.forEach(s => { const g = lc(s.GUIDFirmy), d = toD(s.DataSzkolenia); if (!g || !d || d > TODAY || d < yAgo) return;
    const x = DB.szk[g] = DB.szk[g] || { A: 0, B: 0 }; x[k] += num(s.WartoscDarmowegoSzkolenia) || 0; }));
  // listy pomocnicze
  DB.owners = [...new Set(DB.tecs.map(o => o.owner).concat(DB.sales.map(o => o.owner)).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pl'));
  DB.quarters = [...new Map(DB.tecs.concat(DB.sales).filter(o => o.fq).map(o => [o.fq.key, o.fq])).values()].sort((a, b) => a.key - b.key);
}

/* ================= stan + routing ================= */
const ST = { mod: 'dashboard', tab: 'todo', sort: {}, g: { owner: '', comp: '', q: '', seg: 'all' },
  todo: { plus7: false, desc: false }, akt: { od: new Date(TODAY.getFullYear(), TODAY.getMonth(), 1), do: TODAY, h: '', r: '', f: '' },
  rs: { tryb: 'szanse', od: null, do: null, desc: true, h: '', f: '', r: '', t: '', s: '', to: '', te: '', tf: '', pf: '', pl: '', onlyTP: false },
  hg: { pre: false, all: false, st: '', op: '' }, df: { all: false, op: '', st: '', f: '', lim: 60 }, sstage: '' };
try { const s = JSON.parse(localStorage.getItem('mescocrm') || '{}'); if (s.g) Object.assign(ST.g, s.g); } catch (e) {}
const save = () => { try { localStorage.setItem('mescocrm', JSON.stringify({ g: ST.g })); } catch (e) {} };
const MODS = {
  dashboard: { name: 'Dashboard', icon: 'DASHBOARD', tabs: [['todo','To-Do','clock'],['akt','Aktywności','activity'],['sprz','Raport sprzedaży','OPPSY'],['godz','Godziny TECS / PreSales','bars'],['firmy','Bilans firm','FIRMY']] },
  tecs: { name: 'TECS — odnowienia', icon: 'TECSY', g: true, tabs: [['przeglad','Przegląd','DASHBOARD'],['otwarte','Do zamknięcia','clock'],['zamkniete','Zamknięte','check'],['ceny','Kontrola cen','euro'],['tp','TP','tag'],['bez-oferty','Bez oferty','alert'],['bez-notatek','Bez notatek','note']] },
  oppsy: { name: 'Szanse sprzedaży', icon: 'OPPSY', g: true, tabs: [['s-przeglad','Przegląd','DASHBOARD'],['s-otwarte','W toku','clock'],['s-zamkniete','Wygrane / przegrane','check'],['s-alerty','Alerty','alert']] },
  firmy: { name: 'Klienci — wykorzystanie TECS', icon: 'FIRMY', g: true, noQ: true, tabs: [['klienci','Aktywne TECS i support','SUPPORTY']] },
  demo: { name: 'Kody demo', icon: 'DEMO', g: true, noQ: true, tabs: [['demo-akt','Aktywne i kończące się','clock'],['demo-all','Wszystkie demo','list'],['demo-konw','Konwersja','OPPSY']] },
};
const NAV = [['HR',null],['FIRMY','firmy'],['KONTAKTY',null],['KALENDARZ',null],['CENNIK',null],['SZKOLENIA',null],['OPPSY','oppsy'],['TECSY','tecs'],['DEMO','demo'],['PROSPECTING',null],['SUPPORTY',null],['ZADANIA',null],['MARKETING',null],['DASHBOARD','dashboard'],['ADMIN',null],['USTAWIENIA',null]];
function fromHash() { const [m, t] = location.hash.slice(1).split('/'); if (MODS[m]) { ST.mod = m; ST.tab = MODS[m].tabs.some(x => x[0] === t) ? t : MODS[m].tabs[0][0]; } }
window.addEventListener('hashchange', () => { fromHash(); render(); });
const go = (m, t) => { location.hash = m + '/' + (t || MODS[m].tabs[0][0]); };

/* ================= tabela, wykres ================= */
function table(id, cols, rows, def, opts = {}) {
  if (!rows.length) return `<div class="empty">${opts.empty || 'Brak pozycji dla wybranych filtrów'}</div>`;
  const s = ST.sort[id] || def || { k: 0, asc: true }; const c = cols[s.k] || cols[0]; const get = c.v || (r => r[c.key]);
  const sorted = [...rows].sort((a, b) => { let x = get(a), y = get(b); if (x instanceof Date) x = +x; if (y instanceof Date) y = +y; if (x == null || x === '') return 1; if (y == null || y === '') return -1; const r = typeof x === 'string' ? x.localeCompare(y, 'pl') : x - y; return s.asc ? r : -r; });
  const lim = opts.limit || 500, shown = sorted.slice(0, lim);
  return `<div class="tbl-wrap"><table data-t="${id}"><thead><tr>${cols.map((c, i) => `<th data-i="${i}" class="${c.num ? 'num' : ''} ${i === s.k ? 'sorted' + (s.asc ? ' asc' : '') : ''}">${c.h}</th>`).join('')}</tr></thead><tbody>${
    shown.map(r => `<tr>${cols.map(c => `<td class="${c.num ? 'num' : ''}">${c.f ? c.f(r) : esc((c.v || (x => x[c.key]))(r))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>${sorted.length > lim ? `<div class="desc" style="margin-top:6px">Pokazano ${lim} z ${sorted.length}.</div>` : ''}`;
}
document.addEventListener('click', e => { const th = e.target.closest('th[data-i]'); if (!th) return; const id = th.closest('table').dataset.t, i = +th.dataset.i; const cur = ST.sort[id]; ST.sort[id] = { k: i, asc: cur && cur.k === i ? !cur.asc : false }; render(); });
const kpi = (lab, val, sub, icon, c = '#5B6CDB', bg = '#EEF0FC') => `<div class="kpi">${icon ? `<div class="kic" style="background:${bg}">${svg(icon, c)}</div>` : ''}<div><div class="lab">${lab}</div><div class="val">${val}</div>${sub ? `<div class="s">${sub}</div>` : ''}</div></div>`;
const ptitle = (icon, c, bg, title, sub, right = '') => `<div class="ptitle"><div class="ic" style="background:${bg}">${svg(icon, c)}</div><div><h2>${title}</h2>${sub ? `<div class="sub">${sub}</div>` : ''}</div><div class="right">${right}</div></div>`;
const chips = flags => flags.map(f => `<span class="chip ${f[1]}">${esc(f[0])}</span>`).join('');
const daysCell = d => d == null ? '—' : d < 0 ? `<span class="chip c-red">${-d} dni po</span>` : d <= 90 ? `<span class="chip c-amb">${d} dni</span>` : d + ' dni';
const SPU = 'https://' + CFG.siteHost + CFG.sitePath + '/Lists/';
const tecsLink = o => `<a class="lnk" href="${SPU}Opps_TECS/DispForm.aspx?ID=${o.id}" target="_blank" rel="noopener">#${o.id}</a>`;
const saleLink = o => `<a class="lnk" href="${SPU}Szanse%20Sprzeday/DispForm.aspx?ID=${o.id}" target="_blank" rel="noopener">#${o.id}</a>`;
const compLink = (name, guid) => { const c = DB.compBy[lc(guid)]; return c ? `<a class="lnk" href="${SPU}Companies/DispForm.aspx?ID=${c.id}" target="_blank" rel="noopener">${esc(name)}</a>` : esc(name); };
const stageChip = s => `<span class="chip ${['Zamówione','Wystawiona faktura','Zamknięty wygrany'].includes(s) ? 'c-ok' : ['Rezygnacja z TECS','Zamknięty przegrany'].includes(s) ? 'c-red' : s === 'Demo' ? 'c-vio' : 'c-br'}">${esc(s || '—')}</span>`;

function qChart(list, colors = { won: '#16A34A', open: '#5B6CDB', lost: '#DC2626' }) {
  const ser = [['won','Wygrane',colors.won],['open','Do zamknięcia',colors.open],['lost','Przegrane',colors.lost]];
  const qs = DB.quarters.filter(q => list.some(o => o.fq && o.fq.key === q.key && ['won','open','lost'].includes(o.grp)));
  if (!qs.length) return '<div class="empty">Brak danych</div>';
  const data = qs.map(q => { const o = { q }; ser.forEach(s => { const it = list.filter(x => x.fq && x.fq.key === q.key && x.grp === s[0]); o[s[0]] = sum(it, x => x.value); o[s[0]+'N'] = it.length; }); o.tot = o.won + o.open + o.lost; return o; });
  const max = Math.max(...data.map(d => d.tot), 1), W = Math.max(600, qs.length * 64), H = 250, pl = 58, pb = 38, pt = 8, iw = W - pl - 10, ih = H - pb - pt, bw = Math.min(34, iw / qs.length * .6);
  const nice = v => { const p = Math.pow(10, Math.floor(Math.log10(v))); const n = v / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; };
  const top = nice(max), y = v => pt + ih - v / top * ih, cur = fyQ(TODAY);
  let g = [0,.25,.5,.75,1].map(t => `<line x1="${pl}" x2="${W-10}" y1="${y(t*top)}" y2="${y(t*top)}" stroke="#EEF0F3"/><text x="${pl-6}" y="${y(t*top)+4}" text-anchor="end">${fmtK(t*top)}</text>`).join('');
  data.forEach((d, i) => { const cx = pl + iw * (i + .5) / qs.length, x = cx - bw/2; let acc = 0; const segs = ser.filter(s => d[s[0]] > 0);
    segs.forEach((s, j) => { const v = d[s[0]], y0 = y(acc), h = Math.max(0, y0 - y(acc + v) - (j < segs.length - 1 ? 2 : 0)), r = Math.min(4, h);
      g += j === segs.length - 1 ? `<path d="M${x},${y0} V${y0-h+r} Q${x},${y0-h} ${x+r},${y0-h} H${x+bw-r} Q${x+bw},${y0-h} ${x+bw},${y0-h+r} V${y0} Z" fill="${s[2]}"/>` : `<rect x="${x}" y="${y0-h}" width="${bw}" height="${h}" fill="${s[2]}"/>`; acc += v; });
    const isC = cur && d.q.key === cur.key;
    g += `<text x="${cx}" y="${H-pb+15}" text-anchor="middle" style="${isC ? 'font-weight:700;fill:#1F2937' : ''}">${d.q.label}</text>${isC ? `<text x="${cx}" y="${H-pb+29}" text-anchor="middle" style="font-size:10px">bieżący</text>` : ''}<rect class="qhit" data-i="${i}" x="${cx - iw/qs.length/2}" y="${pt}" width="${iw/qs.length}" height="${ih}" fill="transparent" style="cursor:pointer"/>`; });
  setTimeout(() => document.querySelectorAll('.qhit').forEach(r => { const d = data[+r.dataset.i];
    r.onmousemove = e => showTip(e, `<b>${d.q.label} · ${qRange(d.q.fy, d.q.q)}</b>${ser.map(s => `<div class="r"><span>${s[1]} (${d[s[0]+'N']})</span><span>${fmt(d[s[0]])}</span></div>`).join('')}<div class="r" style="border-top:1px solid #E5E7EB;margin-top:3px;padding-top:3px"><span>Razem</span><span>${fmt(d.tot)}</span></div>`);
    r.onmouseleave = hideTip; r.onclick = () => { ST.g.q = String(d.q.key); save(); hideTip(); render(); }; }), 0);
  return `<div class="legend">${ser.map(s => `<span><i style="background:${s[2]}"></i>${s[1]}</span>`).join('')}</div><div style="overflow-x:auto"><svg width="${W}" height="${H}" role="img" aria-label="Wartość wg kwartału FY">${g}</svg></div>`;
}

/* ================= powłoka ================= */
function renderNav() {
  $('#navItems').innerHTML = NAV.map(([n, m]) => `<a class="nav ${m === ST.mod ? 'on' : ''} ${m ? '' : 'ext'}" ${m ? `href="#${m}"` : (CFG.powerAppsUrl ? `href="${CFG.powerAppsUrl}" target="_blank" rel="noopener"` : '')} title="${m ? n : n + ' — w aplikacji Power Apps'}"><div class="pill"></div>${svg(n, null, 1.9)}<span>${n === 'DEMO' ? 'KODY DEMO' : n}</span></a>`).join('');
}
function gFilters(mod) {
  const M = MODS[mod]; if (!M.g) return '';
  const comps = [...new Set(DB.tecs.map(o => o.comp).concat(DB.sales.map(o => o.comp)))].sort((a, b) => a.localeCompare(b, 'pl'));
  const cq = fyQ(TODAY);
  return `<div class="filters">
    <select id="gOwner"><option value="">Wszyscy opiekunowie</option>${DB.owners.map(o => `<option ${o === ST.g.owner ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>
    <input type="search" id="gComp" list="gCompL" placeholder="Firma…" value="${esc(ST.g.comp)}"><datalist id="gCompL">${comps.map(c => `<option value="${esc(c)}">`).join('')}</datalist>
    ${M.noQ ? '' : `<select id="gQ"><option value="">Wszystkie kwartały FY</option><option value="future" ${ST.g.q === 'future' ? 'selected' : ''}>Od bieżącego kwartału</option>${DB.quarters.map(q => `<option value="${q.key}" ${String(q.key) === ST.g.q ? 'selected' : ''}>${q.label}${cq && q.key === cq.key ? ' (bieżący)' : ''} · ${qRange(q.fy, q.q)}</option>`).join('')}</select>`}
    <div class="seg" id="gSeg">${[['all','Wszystkie'],['com','Komercyjne'],['ac','Academic']].map(([k, l]) => `<button data-v="${k}" class="${ST.g.seg === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <button class="btn gray" id="gReset">Wyczyść filtry</button>
    <span class="sub" style="margin-left:auto;font-size:11.5px;color:#6B7280">Dane: ${DB.loaded ? DB.loaded.toLocaleTimeString('pl-PL', { hour: '2-digit', minute: '2-digit' }) : ''}</span>
    <button class="btn" id="gReload">${svg('refresh')} Odśwież</button></div>`;
}
function bindG() {
  const o = $('#gOwner'); if (!o) return;
  o.onchange = () => { ST.g.owner = o.value; save(); render(); };
  $('#gComp').onchange = e => { ST.g.comp = e.target.value.trim(); save(); render(); };
  if ($('#gQ')) $('#gQ').onchange = e => { ST.g.q = e.target.value; save(); render(); };
  document.querySelectorAll('#gSeg button').forEach(b => b.onclick = () => { ST.g.seg = b.dataset.v; save(); render(); });
  $('#gReset').onclick = () => { ST.g = { owner: '', comp: '', q: '', seg: 'all' }; ST.sstage = ''; save(); render(); };
  $('#gReload').onclick = reload;
}
const gf = (arr, ignoreQ) => arr.filter(o => (!ST.g.owner || o.owner === ST.g.owner) && (!ST.g.comp || lc(o.comp).includes(lc(ST.g.comp)))
  && (ST.g.seg === 'all' || (ST.g.seg === 'ac' ? o.ac : !o.ac))
  && (ignoreQ || !ST.g.q || (ST.g.q === 'future' ? (o.fq && o.fq.key >= fyQ(TODAY).key) : (o.fq && String(o.fq.key) === ST.g.q))));
async function reload() { try { await loadAll(); loader(false); render(); toast('Dane odświeżone', 'ok'); } catch (e) { loader(false); toast('Błąd: ' + e.message, 'err'); } }

function render() {
  hideTip(); renderNav();
  const M = MODS[ST.mod];
  const counts = tabCounts();
  $('#hdr').innerHTML = `<div class="ic">${svg(M.icon, '#5B6CDB')}</div><div><h1>${M.name}</h1><div class="sub">${esc(ME.name)}</div></div>
    <div class="tabs">${M.tabs.map(([k, l, ic]) => `<button class="tab ${ST.tab === k ? 'on' : ''}" data-t="${k}">${svg(ic)}${l}${counts[k] != null ? `<span class="n ${counts[k + '!'] ? 'al' : ''}">${counts[k]}</span>` : ''}</button>`).join('')}</div>${gFilters(ST.mod)}`;
  document.querySelectorAll('.tab').forEach(b => b.onclick = () => go(ST.mod, b.dataset.t));
  bindG();
  const V = VIEWS[ST.tab]; $('#view').innerHTML = V ? V() : '';
  if (V && V.after) V.after();
  (AFTER[ST.tab] || (() => {}))();
  document.querySelectorAll('[data-gocomp]').forEach(a => a.onclick = e => { e.preventDefault(); ST.g.comp = a.dataset.gocomp; save(); go(a.dataset.mod || 'tecs', a.dataset.tab); });
}
function tabCounts() {
  const c = {};
  if (ST.mod === 'dashboard') { const t = todoList(); c.todo = t.length; c['todo!'] = t.some(n => n.fu < TODAY); }
  if (ST.mod === 'tecs') { const L = gf(DB.tecs).filter(o => !['demo','tp'].includes(o.grp)); c.otwarte = L.filter(o => o.grp === 'open').length; c['bez-oferty'] = L.filter(o => o.grp === 'open' && !o.nOffers).length; c['bez-oferty!'] = 1;
    c['bez-notatek'] = L.filter(o => o.grp === 'open' && !o.nNotes && o.days != null && o.days <= 90).length; c['bez-notatek!'] = 1; c.ceny = L.reduce((a, o) => a + o.priceIssues.filter(p => p.diff != null).length, 0); c['ceny!'] = 1; c.tp = gf(DB.tecs).filter(o => o.grp === 'tp').length; }
  if (ST.mod === 'oppsy') { const S = gf(DB.sales); c['s-otwarte'] = S.filter(o => o.grp === 'open').length; c['s-alerty'] = S.filter(o => o.grp === 'open' && o.flags.some(f => f[1] === 'c-red')).length; c['s-alerty!'] = 1; }
  if (ST.mod === 'firmy') { c.klienci = clients().filter(x => x.alert).length; c['klienci!'] = 1; }
  if (ST.mod === 'demo') { c['demo-akt'] = demos().filter(d => d.active).length; c['demo-all'] = demos().length; }
  return c;
}

/* ================= DASHBOARD (jak w Power Apps) ================= */
const VIEWS = {}, AFTER = {};
const todoList = () => { const lim = ST.todo.plus7 ? addDays(TODAY, 7) : TODAY; return DB.notes.filter(n => !n.fuDone && n.fu && n.fu <= lim && n.creator === ME.name).sort((a, b) => ST.todo.desc ? b.fu - a.fu : a.fu - b.fu); };
VIEWS.todo = () => {
  const L = todoList(), zal = L.filter(n => n.fu < TODAY).length, dzis = L.filter(n => +n.fu === +TODAY).length;
  const card = n => { const dd = Math.round((TODAY - n.fu) / 864e5);
    const fuChip = dd > 0 ? `<span class="chip c-red">zaległy ${dd} dni</span>` : dd === 0 ? `<span class="chip c-amb">follow-up dziś</span>` : `<span class="chip c-blue">za ${-dd} dni</span>`;
    return `<div class="fu"><div class="top"><span class="firma lnk" title="Karta firmy">${compLink(n.firma, n.guid)}</span>${fuChip}<button class="btn ok" data-done="${n.id}" title="Oznacz follow-up jako zrobiony">✓ Zrobione</button></div>
      <div><span class="chip c-br">${esc(n.rodzaj || 'notatka')}</span> <span style="font-size:12px;color:#374151">dla: ${esc(n.osoba || '—')}</span></div>
      <div class="txt" title="${esc(n.text)}">${esc(n.text || '—')}</div>
      <div class="bot"><span class="info">Wpis z ${dpl(n.data)} · follow-up ${dpl(n.fu)} · połączeń: ${n.calls}</span>
      ${ME.staff ? `<button class="btn" data-note="${n.id}">${svg('note')} Dodaj notatkę</button>${n.tel ? `<a class="btn" href="tel:${esc(n.tel.replace(/\s/g, ''))}" data-call="${n.id}">${svg('phone')} Zadzwoń</a><a class="btn pri" href="tel:${esc(n.tel.replace(/\s/g, ''))}" data-calln="${n.id}">${svg('phone')} Zadzwoń + notatka</a>` : ''}` : ''}</div></div>`; };
  return `<div class="card panel">${ptitle('clock', '#D97706', '#FEF3E2', 'Twoje follow-upy', `${L.length} do zrobienia`, `<span class="chip c-red">Zaległe: ${zal}</span><span class="chip c-amb">Na dziś: ${dzis}</span>
    <div class="seg" id="tdSeg"><button class="${!ST.todo.plus7 ? 'on' : ''}" data-v="0">Do dziś</button><button class="${ST.todo.plus7 ? 'on' : ''}" data-v="1">+ najbliższe 7 dni</button></div>
    <button class="btn" id="tdSort">${svg('sort')} ${ST.todo.desc ? 'Najnowsze najpierw' : 'Najstarsze najpierw'}</button>`)}
    ${L.length ? `<div class="fu-grid">${L.map(card).join('')}</div>` : `<div class="empty">Brak zaległych follow-upów – wszystko zrobione 🎉</div>`}</div>`;
};
AFTER.todo = () => {
  document.querySelectorAll('#tdSeg button').forEach(b => b.onclick = () => { ST.todo.plus7 = b.dataset.v === '1'; render(); });
  $('#tdSort').onclick = () => { ST.todo.desc = !ST.todo.desc; render(); };
  document.querySelectorAll('[data-done]').forEach(b => b.onclick = async () => { const n = DB.notes.find(x => x.id === +b.dataset.done); b.disabled = true;
    try { await patchItem('Notes', n.id, { FollowUpStatus: true }); n.fuDone = true; toast('Follow-up oznaczony jako zrobiony', 'ok'); render(); } catch (e) { b.disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); } });
  document.querySelectorAll('[data-note]').forEach(b => b.onclick = () => noteModal(DB.notes.find(x => x.id === +b.dataset.note)));
  const call = async (id) => { const n = DB.notes.find(x => x.id === id); const v = Math.min(n.calls + 1, 100); try { await patchItem('Notes', n.id, { LiczbaPolaczen: String(v) }); n.calls = v; } catch (e) { toast('Nie zapisano licznika połączeń', 'err'); } return n; };
  document.querySelectorAll('[data-call]').forEach(a => a.addEventListener('click', () => call(+a.dataset.call).then(() => render())));
  document.querySelectorAll('[data-calln]').forEach(a => a.addEventListener('click', () => call(+a.dataset.calln).then(n => noteModal(n))));
};
/* --- modal notatki po follow-upie --- */
function noteModal(src) {
  const st = { rodzaj: '', dni: 0 };
  const rodz = DB.rodzaje;
  const m = $('#modal'); m.hidden = false;
  const draw = () => {
    const quick = [['Telefon', 'telefon'], ['Oferta', 'oferta'], ['Demo', 'demo']];
    m.innerHTML = `<div class="mbox"><div class="mh"><div class="ic" style="width:36px;height:36px;border-radius:8px;background:#FEF3E2;display:flex;align-items:center;justify-content:center">${svg('note', '#D97706')}</div>
      <div style="flex:1"><div style="font-weight:600;font-size:15px">${esc(src.firma || 'Notatka')}</div><div class="sub" style="font-size:11.5px;color:#6B7280">Nowa notatka po follow-upie · ${esc(src.osoba || '—')}${src.tel ? ' · ' + esc(src.tel) : ''}</div></div>
      <button class="btn gray" id="mClose" title="Zamknij">${svg('x')}</button></div>
      <div class="fgrid"><label>Data wydarzenia<input type="date" id="mData" value="${ds(TODAY)}"></label><label>Osoba<input type="text" id="mOsoba" value="${esc(src.osoba)}"></label>
      <label>Rodzaj kontaktu *<select id="mRodz"><option value="">Wybierz rodzaj</option>${rodz.map(r => `<option value="${r.id}" ${lc(r.name).startsWith(st.rodzaj) && st.rodzaj ? 'selected' : ''}>${esc(r.name)}</option>`).join('')}</select></label></div>
      <label class="fl" style="margin-top:12px">Notatka *<textarea id="mTxt" rows="6" placeholder="Co ustaliliście?"></textarea></label>
      <div class="sec">KOLEJNY FOLLOW-UP</div>
      <div class="fgrid" style="margin-top:6px"><label style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="mFu" ${st.dni ? 'checked' : ''}> ustaw follow-up</label>
      <label>Data follow-upu<input type="date" id="mFuD" value="${st.dni ? ds(addDays(TODAY, st.dni)) : ''}"></label>
      <label style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="mClosePrev" checked> poprzedni follow-up: oznacz jako zrobiony</label></div>
      <div class="desc" style="margin-top:12px">Szybki wybór: rodzaj kontaktu + termin kolejnego follow-upu</div>
      ${quick.map(([l, k]) => `<div class="qrow"><b>${l}</b>${[[7,'7 dni'],[14,'14 dni'],[21,'21 dni'],[28,'28 dni'],[56,'2 miesiące']].map(([d, t]) => `<button class="btn ${st.rodzaj === k && st.dni === d ? 'sel' : ''}" data-q="${k}|${d}">${t}</button>`).join('')}</div>`).join('')}
      <div class="mfoot"><span id="mMiss" style="font-size:12.5px"></span><span style="flex:1"></span><button class="btn gray" id="mCancel" style="height:38px;width:110px;justify-content:center">Anuluj</button><button class="btn pri" id="mSave" style="height:38px;padding:0 20px">Zapisz notatkę</button></div></div>`;
    const txt = m.querySelector('#mTxt'); if (draw._txt) txt.value = draw._txt;
    const check = () => { const miss = []; if (!$('#mRodz').value) miss.push('rodzaj'); if (!txt.value.trim()) miss.push('notatka'); if ($('#mFu').checked && !$('#mFuD').value) miss.push('data follow-upu');
      $('#mMiss').textContent = miss.length ? 'Uzupełnij: ' + miss.join(', ') : 'Gotowe do zapisu'; $('#mMiss').style.color = miss.length ? '#D97706' : '#16A34A'; $('#mSave').disabled = !!miss.length; };
    m.querySelectorAll('input,select,textarea').forEach(el => el.oninput = el.onchange = check); check();
    m.querySelectorAll('[data-q]').forEach(b => b.onclick = () => { const [k, d] = b.dataset.q.split('|'); st.rodzaj = k; st.dni = +d; draw._txt = txt.value; draw(); });
    $('#mClose').onclick = $('#mCancel').onclick = () => { m.hidden = true; m.innerHTML = ''; };
    $('#mSave').onclick = async () => {
      const r = rodz.find(x => String(x.id) === $('#mRodz').value), fuOn = $('#mFu').checked, fuD = fuOn ? toD($('#mFuD').value + 'T12:00:00') : null, d = toD($('#mData').value + 'T12:00:00') || TODAY;
      const f = { Firma: src.firma, IdentyfikatorCompanies: src.raw.IdentyfikatorCompanies || '', IdentyfikatorContact: src.contact, Opps_ID: src.raw.Opps_ID || '', Etap: src.etap, EtapProspecting: src.etapP, Telefon: src.tel,
        Osoba: $('#mOsoba').value, TypeLookupId: String(r.id), Rodzaj1: r.name, Data: isoNoon(d), Totatka: txt.value, DodajFollowUp: fuOn, FollowUpStatus: false, Creator: ME.name };
      if (fuD) f.Follow_up = isoNoon(fuD);
      Object.keys(f).forEach(k => (f[k] === '' || f[k] == null) && delete f[k]);
      $('#mSave').disabled = true;
      try {
        const res = await createItem('Notes', f);
        DB.notes.push({ id: res.id, firma: src.firma, guid: src.guid, osoba: f.Osoba, rodzaj: r.name, typeId: r.id, text: f.Totatka, data: d, fu: fuD, fuDone: false, creator: ME.name, tel: src.tel, calls: 0, opps: src.opps, contact: src.contact, etap: src.etap, etapP: src.etapP, raw: f });
        if ($('#mClosePrev').checked && src.id) { await patchItem('Notes', src.id, { FollowUpStatus: true }); src.fuDone = true; }
        m.hidden = true; m.innerHTML = ''; toast('Notatka zapisana' + (fuD ? ' – follow-up ' + dpl(fuD) : ''), 'ok'); render();
      } catch (e) { $('#mSave').disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); }
    };
  };
  draw();
}
/* --- Aktywności --- */
const AKT_K = [['Wszystkie', null, 'list', '#5B6CDB', '#EEF0FC'], ['Telefony', 'telefon', 'phone', '#2F6FD6', '#E8F0FD'], ['Online', 'spotkanie online', 'DASHBOARD', '#7C3AED', '#F1EAFE'], ['Spotkania F2F', 'spotkanie f2f', 'users', '#16A34A', '#E6F6EE'], ['Oferty', 'oferta', 'box', '#D97706', '#FEF3E2'], ['Demo', 'demo', 'target', '#0D9488', '#E0F5F3']];
const aktBase = () => DB.notes.filter(n => n.data && n.data >= ST.akt.od && n.data <= ST.akt.do);
VIEWS.akt = () => {
  const base = aktBase(), A = ST.akt;
  const L = base.filter(n => (!A.h || n.creator === A.h) && (!A.r || n.rodzaj === A.r) && (!A.f || n.firma === A.f));
  const uniq = k => [...new Set(base.map(n => n[k]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pl'));
  const b = fyBase(), isR = (o, d) => +A.od === +o && +A.do === +d;
  const pm = new Date(TODAY.getFullYear(), TODAY.getMonth() - 1, 1), pmE = new Date(TODAY.getFullYear(), TODAY.getMonth(), 0);
  const quick = [['Ten mies.', new Date(TODAY.getFullYear(), TODAY.getMonth(), 1), TODAY], ['Poprz. mies.', pm, pmE], ...[1,2,3,4].map(q => ['Q' + q, ...qBounds(b, q)]), ['Rok FY', new Date(b, 10, 1), new Date(b+1, 9, 31)]];
  const cnt = p => p == null ? L.length : L.filter(n => lc(n.rodzaj).startsWith(p)).length;
  const groups = [...new Set(L.map(n => n.creator))].map(c => { const x = L.filter(n => n.creator === c); return { c, n: x.length, r: AKT_K.slice(1).map(k => x.filter(n => lc(n.rodzaj).startsWith(k[1])).length) }; }).sort((a, b) => b.n - a.n);
  const cols = [{ h: 'Data', key: 'data', f: n => dpl(n.data) }, { h: 'Handlowiec', key: 'creator' }, { h: 'Rodzaj', key: 'rodzaj', f: n => `<span class="chip c-br">${esc(n.rodzaj || '—')}</span>` },
    { h: 'Firma', key: 'firma', f: n => compLink(n.firma, n.guid) }, { h: 'Notatka', key: 'text', f: n => `<span title="${esc(n.text)}" style="color:#6B7280">${esc(String(n.text).replace(/\n/g, ' ').slice(0, 140))}</span>` }];
  return `<div class="card panel">${ptitle('activity', '#2F6FD6', '#E8F0FD', 'Aktywności', `${dpl(A.od)} – ${dpl(A.do)} · ${L.length} wpisów`)}
    <div class="filters" style="border:0;padding:0;margin-bottom:10px">
    <select id="aH"><option value="">Handlowiec</option>${uniq('creator').map(x => `<option ${x === A.h ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>
    <select id="aR"><option value="">Rodzaj</option>${uniq('rodzaj').map(x => `<option ${x === A.r ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select>
    <input type="search" id="aF" list="aFL" placeholder="Firma" value="${esc(A.f)}"><datalist id="aFL">${uniq('firma').map(x => `<option value="${esc(x)}">`).join('')}</datalist>
    <input type="date" id="aOd" value="${ds(A.od)}"><input type="date" id="aDo" value="${ds(A.do)}">
    ${quick.map(([l, o, d], i) => `<button class="btn ${isR(o, d) ? 'sel' : 'gray'}" data-qa="${i}" title="${l.startsWith('Q') ? 'Kwartał fiskalny ' + l + ' (rok XI–X)' : ''}">${l}</button>`).join('')}</div>
    <div class="kpis" style="grid-template-columns:repeat(6,1fr)">${AKT_K.map(k => kpi(k[0], cnt(k[1]), '', k[2], k[3], k[4])).join('')}</div>
    <div class="grid2" style="grid-template-columns:7fr 3fr"><div>${table('akt', cols, L, { k: 0, asc: false }, { empty: 'Brak aktywności dla wybranych filtrów' })}</div>
    <div style="background:#F8F9FB;border-radius:12px;padding:12px"><div style="font-weight:600;margin-bottom:8px">Podsumowanie wg handlowca</div>
    ${groups.map(g => `<div style="background:#fff;border:1px solid #E5E7EB;border-radius:8px;padding:8px 10px;margin-bottom:6px;display:flex;gap:10px;align-items:center"><div class="av" style="width:32px;height:32px;flex-basis:32px">${initials(g.c)}</div>
      <div style="flex:1"><div style="display:flex;justify-content:space-between"><b style="font-size:13px">${esc(g.c || '—')}</b><b style="color:#5B6CDB;font-size:15px">${g.n}</b></div><div style="font-size:11px;color:#6B7280">Tel ${g.r[0]} · Online ${g.r[1]} · F2F ${g.r[2]} · Oferty ${g.r[3]} · Demo ${g.r[4]}</div></div></div>`).join('') || '<div class="empty">—</div>'}</div></div></div>`;
};
AFTER.akt = () => {
  const A = ST.akt; $('#aH').onchange = e => { A.h = e.target.value; render(); }; $('#aR').onchange = e => { A.r = e.target.value; render(); }; $('#aF').onchange = e => { A.f = e.target.value; render(); };
  $('#aOd').onchange = e => { A.od = toD(e.target.value + 'T12:00') || A.od; render(); }; $('#aDo').onchange = e => { A.do = toD(e.target.value + 'T12:00') || A.do; render(); };
  const b = fyBase(), pm = new Date(TODAY.getFullYear(), TODAY.getMonth() - 1, 1), pmE = new Date(TODAY.getFullYear(), TODAY.getMonth(), 0);
  const q = [[new Date(TODAY.getFullYear(), TODAY.getMonth(), 1), TODAY], [pm, pmE], ...[1,2,3,4].map(i => qBounds(b, i)), [new Date(b, 10, 1), new Date(b+1, 9, 31)]];
  document.querySelectorAll('[data-qa]').forEach(x => x.onclick = () => { [A.od, A.do] = q[+x.dataset.qa]; render(); });
};
/* --- Raport sprzedaży --- */
VIEWS.sprz = () => {
  const R = ST.rs, b = fyBase();
  const inR = d => (!R.od || (d && d >= R.od)) && (!R.do || (d && d <= R.do));
  const qBtns = R.tryb === 'tp' ? '' : `<input type="date" id="rOd" value="${ds(R.od)}"><input type="date" id="rDo" value="${ds(R.do)}">${[1,2,3,4].map(q => { const [o, d] = qBounds(b, q); return `<button class="btn ${R.od && +R.od === +o && +R.do === +d ? 'sel' : 'gray'}" data-rq="${q}">Q${q}</button>`; }).join('')}<button class="btn ${!R.od && !R.do ? 'sel' : 'gray'}" data-rq="0">Wszystko</button><button class="btn" id="rSort">${svg('sort')} ${R.desc ? 'Najnowsze najpierw' : 'Najstarsze najpierw'}</button>`;
  const head = ptitle('OPPSY', '#16A34A', '#E6F6EE', 'Raport sprzedaży', { tecs: 'TECSy wg końca okresu', tp: 'Licencje w TECS / Tech Pool' }[R.tryb] || 'Szanse wg daty zamknięcia',
    `<div class="seg" id="rTryb">${[['szanse','Szanse'],['tecs','TECSy'],['tp','Licencje Tech Pool']].map(([k, l]) => `<button class="${R.tryb === k ? 'on' : ''}" data-v="${k}">${l}</button>`).join('')}</div>${qBtns}`);
  const sel = (id, ph, vals, cur) => `<select id="${id}"><option value="">${ph}</option>${vals.map(v => `<option ${v === cur ? 'selected' : ''}>${esc(v)}</option>`).join('')}</select>`;
  const uniq = (a, f) => [...new Set(a.map(f).filter(Boolean))].sort((x, y) => x.localeCompare(y, 'pl'));
  if (R.tryb === 'szanse') {
    const base = DB.sales.filter(o => (!R.od && !R.do) || inR(o.closedWin));
    const L = base.filter(o => (!R.h || o.creator === R.h) && (!R.f || o.comp === R.f) && (!R.r || o.rodz === R.r) && (!R.t || o.typ === R.t) && (!R.s || o.status1 === R.s));
    const cols = [{ h: 'Firma', key: 'comp', f: o => `${saleLink(o)} <b>${esc(o.comp)}</b>` }, { h: 'Rodzaj', key: 'rodz' }, { h: 'Typ', key: 'typ' }, { h: 'Wartość', key: 'eur', num: 1, f: o => `<b>${eur(o.eur)}</b>` }, { h: 'Marża MESco', key: 'marza', num: 1, f: o => `<span style="color:#5B6CDB">${eur(o.marza)}</span>` },
      { h: 'Premia Q', key: 'premQ', num: 1, f: o => eur(o.premQ) }, { h: 'Premia R', key: 'premR', num: 1, f: o => eur(o.premR) }, { h: 'Zamknięcie', key: 'closedWin', f: o => dpl(o.closedWin) },
      { h: 'Status', key: 'status1', f: o => `<b style="color:${o.status1 === 'Zamknięty wygrany' ? '#16A34A' : o.status1 === 'Zamknięty przegrany' ? '#DC2626' : '#5B6CDB'}">${esc(o.status1)}</b>` }, { h: 'Handlowiec', key: 'creator' }];
    return `<div class="card panel">${head}<div class="filters" style="border:0;padding:0;margin-bottom:10px">${sel('rH', 'Handlowiec', uniq(base, o => o.creator), R.h)}${sel('rF', 'Firma', uniq(base, o => o.comp), R.f)}${sel('rR', 'Rodzaj biznesu', uniq(base, o => o.rodz), R.r)}${sel('rT', 'Typ biznesu', uniq(base, o => o.typ), R.t)}${sel('rS', 'Status', uniq(base, o => o.status1), R.s)}</div>
      <div class="kpis" style="grid-template-columns:repeat(5,1fr)">${kpi('Liczba szans', L.length)}${kpi('Wartość', `<span style="color:#16A34A">${eur(sum(L, o => o.eur))}</span>`)}${kpi('Marża MESco', `<span style="color:#5B6CDB">${eur(sum(L, o => o.marza))}</span>`)}${kpi('Do premii Q', `<span style="color:#D97706">${eur(sum(L, o => o.premQ))}</span>`)}${kpi('Do premii R', `<span style="color:#7C3AED">${eur(sum(L, o => o.premR))}</span>`)}</div>
      ${table('rsz', cols, L, { k: 7, asc: !R.desc })}</div>`;
  }
  if (R.tryb === 'tecs') {
    const base = DB.tecs.filter(o => inR(o.end));
    const L = base.filter(o => (!R.to || o.owner === R.to) && (!R.te || o.stage === R.te) && (!R.tf || o.comp === R.tf));
    const W = o => o.wart || 0;
    const cols = [{ h: 'Firma', key: 'comp', f: o => `${tecsLink(o)} <b title="${esc(o.name)}">${esc(o.comp)}</b>` }, { h: 'Początek', key: 'start', f: o => dpl(o.start) }, { h: 'Koniec', key: 'end', f: o => dpl(o.end) }, { h: 'Wartość', v: W, num: 1, f: o => `<b>${eur(W(o))}</b>` },
      { h: 'NTA 67%', v: W, num: 1, f: o => `<span style="color:#5B6CDB">${eur(W(o) * .67)}</span>` }, { h: 'NTM 33%', v: W, num: 1, f: o => eur(W(o) * .33) }, { h: 'Etap', key: 'stage', f: o => stageChip(o.stage) }, { h: 'Opiekun', key: 'owner' }];
    return `<div class="card panel">${head}<div class="filters" style="border:0;padding:0;margin-bottom:10px">${sel('rTo', 'Opiekun', uniq(base, o => o.owner), R.to)}${sel('rTe', 'Etap', uniq(base, o => o.stage), R.te)}${sel('rTf', 'Firma', uniq(base, o => o.comp), R.tf)}</div>
      <div class="kpis" style="grid-template-columns:repeat(4,1fr)">${kpi('Liczba TECSów', L.length)}${kpi('Wartość TECS', `<span style="color:#16A34A">${eur(sum(L, W))}</span>`)}${kpi('Wartość NTA (67%)', `<span style="color:#5B6CDB">${eur(sum(L, W) * .67)}</span>`)}${kpi('Wartość NTM (33%)', `<span style="color:#D97706">${eur(sum(L, W) * .33)}</span>`)}</div>
      ${table('rte', cols, L, { k: 2, asc: !R.desc })}</div>`;
  }
  const L = DB.tecsLines.filter(l => (!R.pf || l.firma === R.pf) && (!R.pl || l.prod === R.pl) && (!R.onlyTP || l.tp));
  const cols = [{ h: 'Firma', key: 'firma', f: l => `<b>${esc(l.firma)}</b>` }, { h: 'Licencja / produkt', key: 'prod' }, { h: 'Początek', key: 'start', f: l => dpl(l.start) }, { h: 'Koniec', key: 'end', f: l => dpl(l.end) },
    { h: 'Tech Pool', v: l => l.tp ? 1 : 0, f: l => l.tp ? '<span class="chip c-vio">Tech Pool</span>' : '' }, { h: 'Wartość TP', key: 'tpValue', num: 1, f: l => `<b>${eur(l.tpValue)}</b>` }];
  return `<div class="card panel">${head}<div class="filters" style="border:0;padding:0;margin-bottom:10px">${sel('rPf', 'Firma', uniq(DB.tecsLines, l => l.firma), R.pf)}${sel('rPl', 'Licencja', uniq(DB.tecsLines, l => l.prod), R.pl)}
    <div class="seg" id="rTP"><button class="${!R.onlyTP ? 'on' : ''}" data-v="0">Wszystkie</button><button class="${R.onlyTP ? 'on' : ''}" data-v="1">Tylko Tech Pool</button></div></div>
    <div class="kpis" style="grid-template-columns:1fr 1fr">${kpi('Pozycji licencji', L.length)}${kpi('Wartość Tech Pool', `<span style="color:#16A34A">${eur(sum(L, l => l.tpValue))}</span>`)}</div>${table('rtp', cols, L, { k: 0, asc: true })}</div>`;
};
AFTER.sprz = () => {
  const R = ST.rs, on = (id, k) => { const el = $('#' + id); if (el) el.onchange = e => { R[k] = e.target.value; render(); }; };
  document.querySelectorAll('#rTryb button').forEach(b => b.onclick = () => { R.tryb = b.dataset.v; render(); });
  [['rH','h'],['rF','f'],['rR','r'],['rT','t'],['rS','s'],['rTo','to'],['rTe','te'],['rTf','tf'],['rPf','pf'],['rPl','pl']].forEach(([i, k]) => on(i, k));
  if ($('#rOd')) { $('#rOd').onchange = e => { R.od = toD(e.target.value + 'T12:00'); render(); }; $('#rDo').onchange = e => { R.do = toD(e.target.value + 'T12:00'); render(); }; }
  document.querySelectorAll('[data-rq]').forEach(b => b.onclick = () => { const q = +b.dataset.rq; if (!q) { R.od = R.do = null; } else [R.od, R.do] = qBounds(fyBase(), q); render(); });
  if ($('#rSort')) $('#rSort').onclick = () => { R.desc = !R.desc; delete ST.sort.rsz; delete ST.sort.rte; render(); };
  document.querySelectorAll('#rTP button').forEach(b => b.onclick = () => { R.onlyTP = b.dataset.v === '1'; render(); });
};
/* --- Godziny --- */
VIEWS.godz = () => {
  const H = ST.hg, base = DB.comp;
  const L = base.filter(c => (H.pre ? c.presh : c.tecsh) > 0 && (H.all ? (c.name !== 'MESco' && (!H.op || c.creator === H.op)) : c.creator === ME.name) && (!H.st || c.status1 === H.st))
    .map(c => ({ ...c, godz: H.pre ? c.presh : c.tecsh })).sort((a, b) => b.godz - a.godz);
  const max = Math.max(1, ...L.map(c => c.godz)), color = H.pre ? '#7C3AED' : '#5B6CDB';
  const uniq = f => [...new Set(base.map(f).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pl'));
  return `<div class="card panel">${ptitle('bars', '#7C3AED', '#F1EAFE', 'Zużycie godzin', `${H.pre ? 'PreSales' : 'TECS'} · razem ${fmt(sum(L, c => c.godz))} h · ${L.length} firm`,
    `<div class="seg" id="hP"><button class="${!H.pre ? 'on' : ''}" data-v="0">TECS</button><button class="${H.pre ? 'on' : ''}" data-v="1">PreSales</button></div>
     <div class="seg" id="hW"><button class="${!H.all ? 'on' : ''}" data-v="0">Moje</button><button class="${H.all ? 'on' : ''}" data-v="1">Wszystkie</button></div>
     <select id="hSt"><option value="">Status klienta</option>${uniq(c => c.status1).map(s => `<option ${s === H.st ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>
     ${H.all ? `<select id="hOp"><option value="">Opiekun</option>${uniq(c => c.creator).map(s => `<option ${s === H.op ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>` : ''}`)}
    ${L.length ? L.map(c => `<div class="hrow" title="TECS: ${c.tecsh} h · PreSales: ${c.presh} h"><div class="n">${compLink(c.name, c.guid)}</div><div class="track"><div style="width:${Math.max(.6, c.godz / max * 100)}%;background:${color}"></div></div><b style="font-size:12.5px">${c.godz.toLocaleString('pl-PL')} h</b></div>`).join('') : '<div class="empty">Brak godzin dla wybranych filtrów</div>'}</div>`;
};
AFTER.godz = () => { const H = ST.hg; document.querySelectorAll('#hP button').forEach(b => b.onclick = () => { H.pre = b.dataset.v === '1'; render(); }); document.querySelectorAll('#hW button').forEach(b => b.onclick = () => { H.all = b.dataset.v === '1'; render(); });
  $('#hSt').onchange = e => { H.st = e.target.value; render(); }; if ($('#hOp')) $('#hOp').onchange = e => { H.op = e.target.value; render(); }; };
/* --- Bilans firm --- */
const zuzycie = c => c.marza > 0 ? Math.round((c.wartTECS + c.szkA + c.szkB) / c.marza * 100) : 0;
VIEWS.firmy = () => {
  const D = ST.df, base = D.op ? DB.comp.filter(c => c.creator === D.op) : D.all ? DB.comp.filter(c => c.name !== 'MESco') : DB.comp.filter(c => c.creator === ME.name);
  const L = base.filter(c => (!D.st || c.status1 === D.st) && (!D.f || c.name === D.f)).sort((a, b) => a.name.localeCompare(b.name, 'pl'));
  const uniq = f => [...new Set(DB.comp.map(f).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pl'));
  const salesBy = {}; DB.sales.forEach(s => { (salesBy[s.compGuid] = salesBy[s.compGuid] || []).push(s); });
  const notesBy = {}; DB.notes.forEach(n => { if (n.guid) (notesBy[n.guid] = notesBy[n.guid] || []).push(n); });
  const tile = (lab, val, sub, c) => `<div class="kpi strip" style="--c:${c};flex-direction:column;align-items:flex-start;gap:0;padding-left:16px"><div class="lab">${lab}</div><div class="val" style="color:${c};font-size:16px">${val}</div><div class="s">${sub}</div></div>`;
  const card = c => { const ss = salesBy[c.guid] || [], ns = notesBy[c.guid] || [], z = zuzycie(c), bil = c.marza - c.szkA - c.szkB - c.wartTECS;
    const g = st => ss.filter(s => s.status1 === st), cn = p => ns.filter(n => lc(n.rodzaj).startsWith(p)).length;
    return `<div class="bil"><div class="h"><div class="av" title="${esc(c.opiekun)}">${initials(c.opiekun)}</div><div style="flex:1;min-width:260px"><div><span style="font-size:14px">${compLink(c.name, c.guid)}</span> ${c.status1 ? `<span class="chip c-mut">${esc(c.status1)}</span>` : ''}</div>
      <div style="font-size:12px;color:#6B7280">TECS: ${c.tecsh} h · PreSales: ${c.presh} h</div><div style="font-size:12px;color:#374151;margin-top:2px">Aktywności: 📞 ${cn('telefon')}  🖥 ${cn('spotkanie online')}  🤝 ${cn('spotkanie f2f')}  🖱 ${cn('demo')}  💰 ${cn('oferta')}</div></div>
      <div style="width:250px"><div style="text-align:right;font-weight:600;font-size:12.5px;color:${bil < 0 ? '#DC2626' : '#16A34A'}">Poziom zużycia TECSu: ${z} %</div><div class="track" style="margin-top:6px"><div style="width:${Math.max(2, Math.min(z, 100))}%;background:${z > 100 ? '#DC2626' : z > 75 ? '#D97706' : '#16A34A'}"></div></div></div>
      ${ME.staff ? `<button class="btn" data-przelicz="${c.id}">${svg('refresh')} Przelicz</button>` : ''}<button class="btn" data-gocomp="${esc(c.name)}" data-mod="firmy" data-tab="klienci">${svg('FIRMY')} Karta firmy</button></div>
      <div class="tiles">${tile('Aktywne oppsy', eur(sum(g('Aktywny'), s => s.eur)), g('Aktywny').length + ' szt.', '#5B6CDB')}${tile('Wygrane oppsy', eur(sum(g('Zamknięty wygrany'), s => s.eur)), g('Zamknięty wygrany').length + ' szt.', '#16A34A')}
      ${tile('Przegrane oppsy', eur(sum(g('Zamknięty przegrany'), s => s.eur)), g('Zamknięty przegrany').length + ' szt.', '#DC2626')}${tile('Koszt TECSu', eur(c.wartTECS), `(TECS + PreSales) × ${CFG.stawkaGodzinowa} €`, '#DC2626')}
      ${tile('Szkolenia gratis', eur(c.szkA + c.szkB), 'ostatnie 12 mies.', '#DC2626')}${tile('Marża MESco za TECS', eur(c.marza), '33% aktywnych TECS', '#16A34A')}${tile('Bilans', eur(bil), bil < 0 ? 'strata' : 'zysk', bil < 0 ? '#DC2626' : '#16A34A')}</div></div>`; };
  return `<div class="card panel">${ptitle('FIRMY', '#0D9488', '#E0F5F3', 'Bilans firm', `${L.length} firm · koszty TECS i szkoleń vs marża MESco`,
    `<div class="seg" id="fW"><button class="${!D.all ? 'on' : ''}" data-v="0">Moje</button><button class="${D.all ? 'on' : ''}" data-v="1">Wszystkie</button></div>
     <select id="fOp"><option value="">Opiekun</option>${uniq(c => c.creator).map(s => `<option ${s === D.op ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>
     <select id="fSt"><option value="">Status</option>${uniq(c => c.status1).map(s => `<option ${s === D.st ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>
     <input type="search" id="fF" list="fFL" placeholder="Firma" value="${esc(D.f)}"><datalist id="fFL">${uniq(c => c.name).map(s => `<option value="${esc(s)}">`).join('')}</datalist>`)}
    ${L.length ? L.slice(0, D.lim).map(card).join('') + (L.length > D.lim ? `<div style="text-align:center"><button class="btn" id="fMore">Pokaż więcej (${L.length - D.lim})</button></div>` : '') : '<div class="empty">Brak firm dla wybranych filtrów</div>'}</div>`;
};
AFTER.firmy = () => {
  const D = ST.df; document.querySelectorAll('#fW button').forEach(b => b.onclick = () => { D.all = b.dataset.v === '1'; D.lim = 60; render(); });
  $('#fOp').onchange = e => { D.op = e.target.value; render(); }; $('#fSt').onchange = e => { D.st = e.target.value; render(); }; $('#fF').onchange = e => { D.f = e.target.value; render(); };
  if ($('#fMore')) $('#fMore').onclick = () => { D.lim += 60; render(); };
  document.querySelectorAll('[data-przelicz]').forEach(b => b.onclick = async () => {
    const c = DB.comp.find(x => x.id === +b.dataset.przelicz); b.disabled = true;
    const szk = DB.szk[c.guid] || { A: 0, B: 0 };
    const marzaSuma = sum(DB.tecs.filter(t => t.compGuid === c.guid && t.end && t.end >= TODAY), t => t.wart);
    const v = { WartoscTECS: (c.tecsh + c.presh) * CFG.stawkaGodzinowa, MarzaTECS: marzaSuma * .33, SzkoleniaGratis: DB.kurs > 0 ? szk.A / DB.kurs : 0, SzkolenieGratis: DB.kurs > 0 ? szk.B / DB.kurs : 0 };
    try { await patchItem('Companies', c.id, Object.fromEntries(Object.entries(v).map(([k, x]) => [k, String(Math.round(x * 100) / 100)])));
      c.wartTECS = v.WartoscTECS; c.marza = v.MarzaTECS; c.szkA = v.SzkoleniaGratis; c.szkB = v.SzkolenieGratis; toast('Przeliczono bilans firmy ' + c.name, 'ok'); render(); }
    catch (e) { b.disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); }
  });
};

/* ================= TECS ================= */
const tCols = () => [{ h: 'Szansa', key: 'id', f: tecsLink }, { h: 'Firma', key: 'comp', f: o => `<span data-gocomp="${esc(o.comp)}" data-mod="tecs" data-tab="otwarte" class="lnk" style="font-weight:500">${esc(o.comp)}</span>` }, { h: 'Opiekun', key: 'owner' }, { h: 'Etap', key: 'stage', f: o => stageChip(o.stage) },
  { h: 'Koniec TECS', key: 'end', f: o => `<span style="white-space:nowrap">${dpl(o.end)} <span style="color:#9CA3AF">${o.fq ? o.fq.label : ''}</span></span>` },
  { h: 'Wartość', key: 'value', num: 1, f: o => fmt(o.value) + (o.valSrc === 'pole' && o.value ? ' <span title="z pola Wartość — brak sumy oferty" style="color:#9CA3AF">*</span>' : '') }];
const tL = () => gf(DB.tecs).filter(o => !['demo','tp'].includes(o.grp));
VIEWS.przeglad = () => {
  const L = tL(), op = L.filter(o => o.grp === 'open'), won = L.filter(o => o.grp === 'won'), lost = L.filter(o => o.grp === 'lost'), late = op.filter(o => o.days != null && o.days < 0), q90 = op.filter(o => o.days != null && o.days >= 0 && o.days <= 90);
  const cl = won.length + lost.length, owners = [...new Set(L.map(o => o.owner))];
  const orows = owners.map(w => { const x = L.filter(o => o.owner === w), o1 = x.filter(o => o.grp === 'open'), w1 = x.filter(o => o.grp === 'won'), l1 = x.filter(o => o.grp === 'lost');
    return { w, n: o1.length, v: sum(o1, o => o.value), wv: sum(w1, o => o.value), wn: w1.length, lv: sum(l1, o => o.value), ln: l1.length, rate: (w1.length + l1.length) ? w1.length / (w1.length + l1.length) : null,
      noOff: o1.filter(o => !o.nOffers).length, noN: o1.filter(o => !o.nNotes && o.days != null && o.days <= 90).length, late: o1.filter(o => o.days != null && o.days < 0).length }; }).filter(r => r.n + r.wn + r.ln);
  const mx = Math.max(1, ...orows.map(r => r.v)), al = (n, c = 'c-red') => n ? `<span class="chip ${c}">${n}</span>` : '0';
  const ocols = [{ h: 'Opiekun', key: 'w' }, { h: 'Otwarte', key: 'n', num: 1 }, { h: 'Wartość otwartych', key: 'v', num: 1, f: r => `${fmt(r.v)} <span class="bar" style="width:${Math.round(r.v / mx * 80)}px"></span>` }, { h: 'Wygrane', key: 'wv', num: 1, f: r => `${fmt(r.wv)} <span style="color:#9CA3AF">(${r.wn})</span>` },
    { h: 'Przegrane', key: 'lv', num: 1, f: r => `${fmt(r.lv)} <span style="color:#9CA3AF">(${r.ln})</span>` }, { h: 'Win rate', key: 'rate', num: 1, f: r => r.rate == null ? '—' : Math.round(r.rate * 100) + '%' }, { h: 'Po terminie', key: 'late', num: 1, f: r => al(r.late) }, { h: 'Bez oferty', key: 'noOff', num: 1, f: r => al(r.noOff) }, { h: 'Bez notatek ≤3 mies.', key: 'noN', num: 1, f: r => al(r.noN, 'c-amb') }];
  const comps = [...new Set(op.map(o => o.comp))].map(c => { const x = op.filter(o => o.comp === c); return { c, n: x.length, v: sum(x, o => o.value), next: Math.min(...x.map(o => o.days ?? 9999)), fl: x.reduce((a, o) => a + o.flags.filter(f => f[1] === 'c-red').length, 0) }; });
  return `<div class="card panel"><div class="kpis">${kpi('Do zamknięcia', fmt(sum(op, o => o.value)), `${op.length} szans · ${q90.length} kończy się ≤90 dni`, 'clock')}${kpi('Wygrane', fmt(sum(won, o => o.value)), `${won.length} (Zamówione + Faktura)`, 'check', '#16A34A', '#E6F6EE')}
    ${kpi('Przegrane (Rezygnacja)', fmt(sum(lost, o => o.value)), `${lost.length} szans`, 'x', '#DC2626', '#FEE2E2')}${kpi('Win rate', cl ? Math.round(won.length / cl * 100) + '%' : '—', 'liczbowo', 'target', '#0D9488', '#E0F5F3')}${kpi('Po terminie, wciąż otwarte', late.length, fmt(sum(late, o => o.value)) + ' wartości', 'alert', '#D97706', '#FEF3E2')}</div>
    ${ptitle('bars', '#5B6CDB', '#EEF0FC', 'Wartość TECS wg kwartału FY Ansys', 'Kwartał wg Koniec_TECS. Kliknij słupek, żeby filtrować. Bez TP i Demo.')}${qChart(gf(DB.tecs, true).filter(o => !['demo','tp'].includes(o.grp)))}</div>
    <div class="grid2"><div class="card panel">${ptitle('users', '#5B6CDB', '#EEF0FC', 'Opiekunowie', '')}${table('town', ocols, orows, { k: 2, asc: false })}</div>
    <div class="card panel">${ptitle('FIRMY', '#0D9488', '#E0F5F3', 'Firmy z otwartymi TECS', 'Kliknij firmę, żeby zobaczyć jej szanse')}${table('tcmp', [{ h: 'Firma', key: 'c', f: r => `<span class="lnk" data-gocomp="${esc(r.c)}" data-mod="tecs" data-tab="otwarte">${esc(r.c)}</span>` }, { h: 'Szanse', key: 'n', num: 1 }, { h: 'Wartość', key: 'v', num: 1, f: r => fmt(r.v) }, { h: 'Najbliższy koniec', key: 'next', num: 1, f: r => r.next === 9999 ? '—' : daysCell(r.next) }, { h: 'Alerty', key: 'fl', num: 1, f: r => al(r.fl) }], comps, { k: 2, asc: false })}</div></div>`;
};
VIEWS.otwarte = () => { const op = tL().filter(o => o.grp === 'open'), by = s => op.filter(o => o.stage === s);
  return `<div class="card panel"><div class="kpis">${kpi('Przygotowanie oferty', fmt(sum(by('Przygotowanie oferty'), o => o.value)), by('Przygotowanie oferty').length + ' szans', 'note')}${kpi('Oferta wysłana', fmt(sum(by('Oferta wysłana'), o => o.value)), by('Oferta wysłana').length + ' szans', 'box', '#D97706', '#FEF3E2')}
    ${kpi('Oferta do wysłania', op.filter(o => o.flags.some(f => f[0] === 'Oferta do wysłania')).length, 'Przygotowanie oferty, ≤90 dni do końca', 'clock', '#D97706', '#FEF3E2')}${kpi('Po terminie', op.filter(o => o.days != null && o.days < 0).length, 'koniec minął, etap otwarty', 'alert', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('clock', '#5B6CDB', '#EEF0FC', 'Szanse TECS do zamknięcia', 'Etapy: Przygotowanie oferty, Oferta wysłana · * wartość z pola Wartość (brak sumy w ofercie)')}
    ${table('topen', [...tCols(), { h: 'Do końca', key: 'days', num: 1, f: o => daysCell(o.days) }, { h: 'Oferty', key: 'nOffers', num: 1 }, { h: 'Notatki', key: 'nNotes', num: 1, f: o => o.nNotes ? `${o.nNotes} <span style="color:#9CA3AF">${dpl(o.lastNote)}</span>` : '0' }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], op, { k: 6, asc: true })}</div>`; };
VIEWS.zamkniete = () => { const L = tL(), won = L.filter(o => o.grp === 'won'), lost = L.filter(o => o.grp === 'lost'), z = won.filter(o => o.stage === 'Zamówione'), f = won.filter(o => o.stage === 'Wystawiona faktura');
  const cols = [...tCols(), { h: 'Początek', key: 'start', f: o => dpl(o.start) }];
  return `<div class="card panel"><div class="kpis">${kpi('Zamówione', fmt(sum(z, o => o.value)), z.length + ' — czeka na fakturę', 'box', '#16A34A', '#E6F6EE')}${kpi('Wystawiona faktura', fmt(sum(f, o => o.value)), f.length + ' szans', 'check', '#16A34A', '#E6F6EE')}${kpi('Rezygnacja z TECS', fmt(sum(lost, o => o.value)), lost.length + ' szans', 'x', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('x', '#DC2626', '#FEE2E2', 'Przegrane — Rezygnacja z TECS', '')}${table('tlost', cols, lost, { k: 4, asc: false })}</div><div class="card panel">${ptitle('check', '#16A34A', '#E6F6EE', 'Wygrane — Zamówione i Wystawiona faktura', '')}${table('twon', cols, won, { k: 4, asc: false })}</div>`; };
const priceCols = refLab => [{ h: 'Szansa', v: r => r.o.id, f: r => tecsLink(r.o) }, { h: 'Firma', v: r => r.o.comp }, { h: 'Opiekun', v: r => r.o.owner }, { h: 'Koniec TECS', v: r => r.o.end, f: r => dpl(r.o.end) }, { h: 'Produkt', key: 'prod' }, { h: 'Szt.', key: 'qty', num: 1 },
  { h: 'Cena w ofercie', key: 'price', num: 1, f: r => fmt(r.price) }, { h: refLab, key: 'ref', num: 1, f: r => fmt(r.ref) },
  { h: 'Różnica', key: 'diff', num: 1, f: r => r.diff == null ? `<span class="chip c-mut">${esc(r.kind)}</span>` : `<span class="chip ${r.diff < 0 ? 'c-amb' : 'c-blue'}">${r.diff > 0 ? '+' : ''}${fmt(r.diff)} (${r.ref ? (r.diff / r.ref * 100).toFixed(1).replace('.', ',') : '—'}%)</span>` }, { h: 'Rabat', key: 'disc', num: 1, f: r => r.disc ? r.disc + '%' : '—' }];
VIEWS.ceny = () => { const rows = []; tL().forEach(o => o.priceIssues.forEach(p => rows.push({ o, ...p }))); const diff = rows.filter(r => r.diff != null), rest = rows.filter(r => r.diff == null);
  return `<div class="card panel"><div class="kpis">${kpi('Pozycje z inną ceną', diff.length, `w ${new Set(diff.map(r => r.o.id)).size} szansach · ${fmt(sum(diff, r => r.diff * (r.qty || 1)))} różnicy`, 'euro', '#D97706', '#FEF3E2')}${kpi('Bez ceny w ofercie', rest.filter(r => r.kind === 'brak ceny w ofercie').length, 'Cena_TECS pusta lub 0', 'alert')}${kpi('Produkt spoza cennika', rest.filter(r => r.kind === 'brak w cenniku').length, 'nazwa nie pasuje do Cennik_CRM', 'tag')}</div>
    ${ptitle('euro', '#D97706', '#FEF3E2', 'Ceny niezgodne z bieżącym cennikiem', 'Otwarte TECS z końcem w przyszłości (bez TP). Cena_TECS przed rabatem vs TECS Fee.')}${table('pdiff', priceCols('TECS Fee'), diff, { k: 8, asc: true })}</div>
    <div class="card panel">${ptitle('alert', '#6B7280', '#F3F4F6', 'Pozycje bez ceny lub spoza cennika', '')}${table('prest', priceCols('TECS Fee'), rest, { k: 3, asc: true })}</div>`; };
VIEWS.tp = () => { const DONE = ['Zamówione','Wystawiona faktura','Rezygnacja z TECS'], x = gf(DB.tecs).filter(o => o.grp === 'tp'), act = x.filter(o => !DONE.includes(o.stage));
  x.forEach(o => { o._link = DB.tecs.find(p => p.compGuid === o.compGuid && p.id !== o.id && !['tp','demo'].includes(p.grp) && p.end && o.end && Math.abs(p.end - o.end) <= 62 * 864e5) || null; });
  const rows = []; x.forEach(o => o.priceIssues.forEach(p => rows.push({ o, ...p })));
  return `<div class="card panel"><div class="kpis">${kpi('Szanse TP', x.length, 'osobna kategoria — nie są odnawiane', 'tag', '#7C3AED', '#F1EAFE')}${kpi('W toku', act.length, fmt(sum(act, o => o.value)) + ' wartości', 'clock')}${kpi('Bez powiązanej szansy', x.filter(o => !o._link).length, 'ta sama firma, koniec ±2 mies.', 'alert', '#D97706', '#FEF3E2')}</div>
    ${ptitle('tag', '#7C3AED', '#F1EAFE', 'Szanse TP (TP = tak)', 'Nie wchodzą do lejka odnowień. Powiązanie dopasowane automatycznie — docelowo kolumna z ID powiązanej szansy.')}
    ${table('tp', [...tCols(), { h: 'Do końca', key: 'days', num: 1, f: o => DONE.includes(o.stage) ? '—' : daysCell(o.days) }, { h: 'Powiązana szansa', v: o => o._link ? o._link.id : null, f: o => o._link ? `${tecsLink(o._link)} ${stageChip(o._link.stage)}` : '<span class="chip c-amb">nie znaleziono</span>' }, { h: 'Oferty', key: 'nOffers', num: 1 }, { h: 'Notatki', key: 'nNotes', num: 1 }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], x, { k: 4, asc: true })}</div>
    <div class="card panel">${ptitle('euro', '#7C3AED', '#F1EAFE', 'Ceny TP vs cennik (Annual Lease)', 'TP w toku z końcem w przyszłości')}${table('tpc', priceCols('Annual Lease'), rows, { k: 3, asc: true })}</div>`; };
VIEWS['bez-oferty'] = () => { const x = tL().filter(o => o.grp === 'open' && !o.nOffers);
  return `<div class="card panel"><div class="kpis">${kpi('Otwarte TECS bez oferty', x.length, fmt(sum(x, o => o.value)) + ' wg pola Wartość', 'alert', '#DC2626', '#FEE2E2')}${kpi('Kończące się ≤90 dni', x.filter(o => o.days != null && o.days >= 0 && o.days <= 90).length, '', 'clock', '#D97706', '#FEF3E2')}${kpi('Po terminie', x.filter(o => o.days != null && o.days < 0).length, '', 'x', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('alert', '#DC2626', '#FEE2E2', 'TECS bez żadnej pozycji w Oferty_TECS', 'Powiązanie: Oferty_TECS.IdentyfikatorTECS = Opps_TECS.TECSGUID')}${table('nooff', [...tCols(), { h: 'Do końca', key: 'days', num: 1, f: o => daysCell(o.days) }, { h: 'Notatki', key: 'nNotes', num: 1 }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], x, { k: 6, asc: true })}</div>`; };
VIEWS['bez-notatek'] = () => { const x = tL().filter(o => o.grp === 'open' && !o.nNotes), crit = x.filter(o => o.days != null && o.days <= 90);
  const cols = [...tCols(), { h: 'Do końca', key: 'days', num: 1, f: o => daysCell(o.days) }, { h: 'Follow-up', key: 'follow', f: o => dpl(o.follow) }, { h: 'Oferty', key: 'nOffers', num: 1 }];
  return `<div class="card panel"><div class="kpis">${kpi('Otwarte TECS bez notatek', x.length, '', 'note')}${kpi('Pilne: ≤3 miesiące lub po terminie', crit.length, 'kontakt w pierwszej kolejności', 'alert', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('alert', '#DC2626', '#FEE2E2', 'Pilne — koniec ≤90 dni lub już minął', 'Powiązanie: Notes.Opps_ID = Opps_TECS.TECSGUID')}${table('nn1', cols, crit, { k: 6, asc: true })}</div><div class="card panel">${ptitle('note', '#5B6CDB', '#EEF0FC', 'Pozostałe bez notatek', '')}${table('nn2', cols, x.filter(o => !crit.includes(o)), { k: 6, asc: true })}</div>`; };

/* ================= SZANSE SPRZEDAŻY ================= */
const STAGES = ['Pierwszy kontakt','Discovery','Spotkanie techniczne','Demo/Benchmark','Oferta','Rewizja oferty','Decyzja','Wdrożenie/Realizacja'];
const sCols = () => [{ h: 'Szansa', key: 'id', f: o => `${saleLink(o)}<div style="font-size:11px;color:#9CA3AF">${esc(o.name)}</div>` }, { h: 'Firma', key: 'comp', f: o => `<span class="lnk" style="font-weight:500" data-gocomp="${esc(o.comp)}" data-mod="oppsy" data-tab="s-otwarte">${esc(o.comp)}</span>` }, { h: 'Opiekun', key: 'owner' },
  { h: 'Etap', v: o => STAGES.indexOf(o.stage), f: o => `<span class="chip ${o.grp === 'won' ? 'c-ok' : o.grp === 'lost' ? 'c-red' : 'c-br'}">${esc(o.stage || '—')}</span>` }, { h: 'Rodzaj', key: 'rodz' },
  { h: 'Spodziewane zamknięcie', v: o => o.close, f: o => o.close ? `<span style="white-space:nowrap">${dpl(o.close)} <span style="color:#9CA3AF">${o.fq ? o.fq.label : ''}</span></span>` : '<span class="chip c-amb">brak</span>' },
  { h: 'Wartość EUR', key: 'eur', num: 1, f: o => fmt(o.eur) }, { h: 'Suma ofert', key: 'offerSum', num: 1, f: o => o.nOffers ? `${fmt(o.offerSum)} <span style="color:#9CA3AF">(${o.nOffers})</span>` : '<span style="color:#9CA3AF">—</span>' },
  { h: 'Notatki', key: 'nNotes', num: 1, f: o => o.nNotes ? `${o.nNotes} <span style="color:#9CA3AF">${dpl(o.lastNote)}</span>` : '0' }];
VIEWS['s-przeglad'] = () => {
  const S = gf(DB.sales), op = S.filter(o => o.grp === 'open'), won = S.filter(o => o.grp === 'won'), lost = S.filter(o => o.grp === 'lost'), cl = won.length + lost.length;
  const fr = STAGES.map(s => { const x = op.filter(o => o.stage === s); return { s, n: x.length, v: sum(x, o => o.value), noOff: x.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, noN: x.filter(o => !o.nNotes).length }; });
  const mx = Math.max(1, ...fr.map(r => r.v));
  const owners = [...new Set(S.map(o => o.owner))].map(w => { const x = S.filter(o => o.owner === w), o1 = x.filter(o => o.grp === 'open'), w1 = x.filter(o => o.grp === 'won'), l1 = x.filter(o => o.grp === 'lost');
    return { w, n: o1.length, v: sum(o1, o => o.value), wv: sum(w1, o => o.value), wn: w1.length, ln: l1.length, rate: (w1.length + l1.length) ? w1.length / (w1.length + l1.length) : null, late: o1.filter(o => o.days != null && o.days < 0).length, noOff: o1.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, idle: o1.filter(o => o.flags.some(f => /notat/.test(f[0]))).length }; });
  const al = (n, c = 'c-red') => n ? `<span class="chip ${c}">${n}</span>` : '0';
  const rz = [...new Set(op.map(o => o.rodz || '—'))].map(r => { const x = op.filter(o => (o.rodz || '—') === r); return { r, n: x.length, v: sum(x, o => o.value), m: sum(x, o => o.marza) }; });
  return `<div class="card panel"><div class="kpis">${kpi('W toku (Aktywny)', eur(sum(op, o => o.value)), op.length + ' szans', 'clock')}${kpi('Zamknięty wygrany', eur(sum(won, o => o.value)), won.length + ' szans', 'check', '#16A34A', '#E6F6EE')}${kpi('Zamknięty przegrany', eur(sum(lost, o => o.value)), lost.length + ' szans', 'x', '#DC2626', '#FEE2E2')}
    ${kpi('Win rate', cl ? Math.round(won.length / cl * 100) + '%' : '—', cl ? 'wartościowo ' + Math.round(sum(won, o => o.value) / Math.max(1, sum(won, o => o.value) + sum(lost, o => o.value)) * 100) + '%' : '', 'target', '#0D9488', '#E0F5F3')}${kpi('Po terminie zamknięcia', op.filter(o => o.days != null && o.days < 0).length, op.filter(o => !o.close).length + ' bez daty', 'alert', '#D97706', '#FEF3E2')}${kpi('Etap oferty bez oferty', op.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, 'Oferta / Rewizja / Decyzja', 'box', '#DC2626', '#FEE2E2')}</div>
    <div class="grid2"><div>${ptitle('OPPSY', '#5B6CDB', '#EEF0FC', 'Lejek w toku wg etapu (Etap1)', 'Status1 = Aktywny. Kliknij etap, żeby zobaczyć listę.')}
    <table><thead><tr><th>Etap</th><th class="num">Szanse</th><th>Wartość EUR</th><th class="num">Bez oferty</th><th class="num">Bez notatek</th></tr></thead><tbody>${fr.map(r => `<tr data-stage="${esc(r.s)}" style="cursor:pointer"><td>${esc(r.s)}</td><td class="num">${r.n}</td><td class="num" style="text-align:left"><span class="bar" style="width:${Math.round(r.v / mx * 200)}px"></span> ${fmt(r.v)}</td><td class="num">${al(r.noOff)}</td><td class="num">${al(r.noN, 'c-amb')}</td></tr>`).join('')}</tbody></table></div>
    <div>${ptitle('euro', '#16A34A', '#E6F6EE', 'Rodzaj biznesu — w toku', 'z marżą MESco wg Twoich stawek')}${table('srz', [{ h: 'Rodzaj', key: 'r' }, { h: 'Szanse', key: 'n', num: 1 }, { h: 'Wartość EUR', key: 'v', num: 1, f: r => fmt(r.v) }, { h: 'Marża MESco', key: 'm', num: 1, f: r => fmt(r.m) }], rz, { k: 2, asc: false })}</div></div></div>
    <div class="card panel">${ptitle('bars', '#5B6CDB', '#EEF0FC', 'Wartość EUR wg kwartału FY Ansys', 'W toku wg Spodziewanego zamknięcia, zamknięte wg daty zamknięcia (ClosedDateWin). Kliknij słupek, żeby filtrować.')}${qChart(gf(DB.sales, true))}</div>
    <div class="card panel">${ptitle('users', '#5B6CDB', '#EEF0FC', 'Opiekunowie', '')}${table('sown', [{ h: 'Opiekun', key: 'w' }, { h: 'W toku', key: 'n', num: 1 }, { h: 'Wartość w toku', key: 'v', num: 1, f: r => eur(r.v) }, { h: 'Wygrane', key: 'wv', num: 1, f: r => `${eur(r.wv)} <span style="color:#9CA3AF">(${r.wn})</span>` }, { h: 'Win rate', key: 'rate', num: 1, f: r => r.rate == null ? '—' : Math.round(r.rate * 100) + '%' }, { h: 'Po terminie', key: 'late', num: 1, f: r => al(r.late) }, { h: 'Bez oferty', key: 'noOff', num: 1, f: r => al(r.noOff) }, { h: 'Bez notatek / >60 dni', key: 'idle', num: 1, f: r => al(r.idle, 'c-amb') }], owners, { k: 2, asc: false })}</div>`;
};
AFTER['s-przeglad'] = () => document.querySelectorAll('tr[data-stage]').forEach(tr => tr.onclick = () => { ST.sstage = tr.dataset.stage; go('oppsy', 's-otwarte'); });
VIEWS['s-otwarte'] = () => { let op = gf(DB.sales).filter(o => o.grp === 'open'); const all = op; if (ST.sstage) op = op.filter(o => o.stage === ST.sstage);
  return `<div class="card panel">${ptitle('clock', '#5B6CDB', '#EEF0FC', 'Szanse sprzedaży w toku', 'Status1 = Aktywny', `<div class="seg" id="sStg" style="flex-wrap:wrap"><button class="${!ST.sstage ? 'on' : ''}" data-v="">Wszystkie</button>${STAGES.map(s => `<button class="${ST.sstage === s ? 'on' : ''}" data-v="${esc(s)}">${esc(s)} (${all.filter(o => o.stage === s).length})</button>`).join('')}</div>`)}
    <div class="kpis">${kpi('Szanse', op.length, '', 'list')}${kpi('Wartość EUR', eur(sum(op, o => o.value)), '', 'euro', '#16A34A', '#E6F6EE')}${kpi('Marża MESco', eur(sum(op, o => o.marza)), '', 'euro')}${kpi('Suma ofert', fmt(sum(op, o => o.offerSum)), op.filter(o => o.nOffers).length + ' z ofertą', 'box', '#D97706', '#FEF3E2')}</div>
    ${table('sopen', [...sCols(), { h: 'Do zamknięcia', key: 'days', num: 1, f: o => o.days == null ? '—' : o.days < 0 ? `<span class="chip c-red">${-o.days} dni po</span>` : o.days + ' dni' }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], op, { k: 5, asc: true })}</div>`; };
AFTER['s-otwarte'] = () => document.querySelectorAll('#sStg button').forEach(b => b.onclick = () => { ST.sstage = b.dataset.v; render(); });
VIEWS['s-zamkniete'] = () => { const S = gf(DB.sales), won = S.filter(o => o.grp === 'won'), lost = S.filter(o => o.grp === 'lost');
  const cols = [...sCols(), { h: 'Data zamknięcia', key: 'closedWin', f: o => dpl(o.closedWin) }];
  return `<div class="card panel"><div class="kpis">${kpi('Zamknięty wygrany', eur(sum(won, o => o.value)), won.length + ' szans', 'check', '#16A34A', '#E6F6EE')}${kpi('Marża MESco z wygranych', eur(sum(won, o => o.marza)), '', 'euro')}${kpi('Premia Q + R', eur(sum(won, o => o.premQ + o.premR)), '', 'target', '#7C3AED', '#F1EAFE')}${kpi('Zamknięty przegrany', eur(sum(lost, o => o.value)), lost.length + ' szans', 'x', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('check', '#16A34A', '#E6F6EE', 'Wygrane — Status1 = Zamknięty wygrany', '')}${table('swon', cols, won, { k: 9, asc: false })}</div><div class="card panel">${ptitle('x', '#DC2626', '#FEE2E2', 'Przegrane — Status1 = Zamknięty przegrany', '')}${table('slost', cols, lost, { k: 9, asc: false })}</div>`; };
VIEWS['s-alerty'] = () => { const op = gf(DB.sales).filter(o => o.grp === 'open'), late = op.filter(o => o.days != null && o.days < 0), noOff = op.filter(o => o.flags.some(f => f[0] === 'Bez oferty')), idle = op.filter(o => o.flags.some(f => /notat/.test(f[0]))), noD = op.filter(o => !o.close);
  const cols = [...sCols(), { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], colsI = [...sCols(), { h: 'Dni ciszy', key: 'idle', num: 1 }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }];
  return `<div class="card panel"><div class="kpis">${kpi('Po terminie zamknięcia', late.length, eur(sum(late, o => o.value)), 'alert', '#DC2626', '#FEE2E2')}${kpi('Etap oferty bez oferty', noOff.length, eur(sum(noOff, o => o.value)), 'box', '#DC2626', '#FEE2E2')}${kpi('Bez notatek lub >60 dni ciszy', idle.length, '', 'note', '#D97706', '#FEF3E2')}${kpi('Bez daty zamknięcia', noD.length, '', 'KALENDARZ', '#D97706', '#FEF3E2')}</div>
    ${ptitle('box', '#DC2626', '#FEE2E2', 'Na etapie Oferta / Rewizja / Decyzja bez oferty', 'Powiązanie: Oferty.IdentyfikatorOpps = OPPSGUID')}${table('sa1', cols, noOff, { k: 6, asc: false })}</div>
    <div class="card panel">${ptitle('alert', '#DC2626', '#FEE2E2', 'Po spodziewanej dacie zamknięcia', '')}${table('sa2', cols, late, { k: 5, asc: true })}</div>
    <div class="card panel">${ptitle('note', '#D97706', '#FEF3E2', 'Bez notatek lub ostatnia notatka starsza niż 60 dni', 'Powiązanie: Notes.Opps_ID = OPPSGUID')}${table('sa3', colsI, idle, { k: 9, asc: false })}</div>`; };

/* ================= KLIENCI — wykorzystanie TECS ================= */
function clients() {
  const act = DB.tecs.filter(o => o.grp !== 'demo' && o.stage !== 'Rezygnacja z TECS' && o.start && o.end && o.start <= TODAY && o.end >= TODAY && o.compGuid);
  const by = {}; act.forEach(o => (by[o.compGuid] = by[o.compGuid] || []).push(o));
  return Object.entries(by).map(([g, x]) => { const s = DB.sup[g] || { n: 0, n90: 0, n365: 0, h365: 0, last: null, m: Array(12).fill(0) }, val = sum(x, o => o.value);
    return { guid: g, comp: x[0].comp, ac: x[0].ac, owner: [...new Set(x.map(o => o.owner))].join(', '), owners: x.map(o => o.owner), n: x.length, tp: x.some(o => o.tp), val, next: Math.min(...x.map(o => o.days)),
      s, ago: s.last ? Math.round((TODAY - s.last) / 864e5) : null, alert: s.n90 === 0, h10: val ? s.h365 / (val / 1e4) : null }; });
}
const spark = m => { const mx = Math.max(...m, 1); return `<svg class="spark" width="96" height="22">${m.map((v, i) => `<rect class="${v ? '' : 'z'}" x="${i * 8}" y="${v ? 22 - Math.max(2, v / mx * 20) : 20}" width="6" height="${v ? Math.max(2, v / mx * 20) : 2}" rx="1"><title>${DB.supMonths[i]}: ${v}</title></rect>`).join('')}</svg>`; };
VIEWS.klienci = () => {
  const K = clients().filter(c => (!ST.g.owner || c.owners.includes(ST.g.owner)) && (!ST.g.comp || lc(c.comp).includes(lc(ST.g.comp))) && (ST.g.seg === 'all' || (ST.g.seg === 'ac' ? c.ac : !c.ac)));
  const al = K.filter(c => c.alert), soon = al.filter(c => c.next <= 180);
  const cols = [{ h: 'Firma', key: 'comp', f: c => `${compLink(c.comp, c.guid)}${c.ac ? ' <span class="chip c-mut">Academic</span>' : ''}${c.tp ? ' <span class="chip c-vio">TP</span>' : ''}` }, { h: 'Opiekun TECS', key: 'owner' }, { h: 'Aktywne TECS', key: 'n', num: 1 }, { h: 'Wartość TECS', key: 'val', num: 1, f: c => fmt(c.val) },
    { h: 'Najbliższy koniec', key: 'next', num: 1, f: c => c.next <= 90 ? `<span class="chip c-amb">${c.next} dni</span>` : c.next + ' dni' }, { h: 'Wpisy 90 dni', v: c => c.s.n90, num: 1, f: c => c.s.n90 || '<span class="chip c-red">0</span>' }, { h: 'Wpisy 12 mies.', v: c => c.s.n365, num: 1, f: c => c.s.n365 },
    { h: 'Godziny 12 mies.', v: c => c.s.h365, num: 1, f: c => fmt(c.s.h365) }, { h: 'h / 10 tys. TECS', key: 'h10', num: 1, f: c => c.h10 == null ? '—' : c.h10.toFixed(1).replace('.', ',') },
    { h: 'Ostatni wpis', key: 'ago', num: 1, f: c => c.s.last ? `${dpl(c.s.last)} <span style="color:#9CA3AF">(${c.ago} dni)</span>` : '<span class="chip c-red">nigdy</span>' }, { h: `Support ${DB.supMonths[0]} → ${DB.supMonths[11]}`, v: c => c.s.n365, f: c => spark(c.s.m) },
    { h: 'Alert', v: c => c.alert ? (c.next <= 180 ? 2 : 1) : 0, f: c => c.alert ? `<span class="chip c-red">90 dni bez supportu</span>${c.next <= 180 ? '<br><span class="chip c-amb">odnowienie ≤180 dni</span>' : ''}` : '' }];
  return `<div class="card panel"><div class="kpis">${kpi('Klienci z aktywnym TECS', K.length, fmt(sum(K, c => c.val)) + ' wartości TECS', 'TECSY')}${kpi('Bez wpisu w Support od 90 dni', al.length, fmt(sum(al, c => c.val)) + ' wartości TECS', 'alert', '#DC2626', '#FEE2E2')}
    ${kpi('W tym odnowienie ≤180 dni', soon.length, 'najwyższe ryzyko rezygnacji', 'clock', '#D97706', '#FEF3E2')}${kpi('Wpisy support 12 mies.', fmt(sum(K, c => c.s.n365)), fmt(sum(K, c => c.s.h365)) + ' godzin', 'SUPPORTY', '#0D9488', '#E0F5F3')}</div>
    ${ptitle('SUPPORTY', '#0D9488', '#E0F5F3', 'Wykorzystanie TECS na kliencie', 'Aktywny TECS = dziś między Początek_TECS a Koniec_TECS (bez Demo i Rezygnacji, TP razem). Support powiązany po IdentyfikatorCompanies.')}
    ${table('kl', cols, K, { k: 11, asc: false })}</div>`;
};

/* ================= KODY DEMO ================= */
function demos() {
  return gf(DB.tecs, true).filter(o => o.grp === 'demo').map(d => {
    const after = DB.tecs.concat(DB.sales).filter(p => (p.compGuid === d.compGuid && d.compGuid) && p !== d && p.grp !== 'demo' && ((p.start || p.openD || p.close || p.end) && d.start && (p.start || p.openD || p.close || p.end) >= d.start));
    const next = after.sort((a, b) => (a.start || a.openD || a.close || a.end) - (b.start || b.openD || b.close || b.end))[0] || null;
    const prods = d.lines.map(l => l.prod).filter(Boolean);
    return { ...d, planned: !!(d.start && d.start > TODAY), active: d.days != null && d.days >= 0 && !(d.start && d.start > TODAY), ending: d.days != null && d.days >= 0 && d.days <= 7, next, nextKind: next ? (next.status1 !== undefined ? 'szansa' : 'TECS') : '', prods, len: d.start && d.end ? Math.round((d.end - d.start) / 864e5) : null };
  });
}
const dCols = () => [{ h: 'Demo', key: 'id', f: tecsLink }, { h: 'Firma', key: 'comp', f: d => compLink(d.comp, d.compGuid) }, { h: 'Opiekun', key: 'owner' }, { h: 'Produkty', v: d => d.prods.join(', '), f: d => d.prods.length ? d.prods.map(p => `<span class="chip c-mut">${esc(p.replace(/^Ansys /, ''))}</span>`).join('') : '<span style="color:#9CA3AF">—</span>' },
  { h: 'Start', key: 'start', f: d => dpl(d.start) }, { h: 'Koniec', key: 'end', f: d => dpl(d.end) }, { h: 'Dni', key: 'len', num: 1 },
  { h: 'Status', key: 'days', num: 1, f: d => d.planned ? `<span class="chip c-vio">zaplanowane od ${dpl(d.start)}</span>` : d.days == null ? '—' : d.days < 0 ? `<span class="chip c-mut">zakończone ${-d.days} dni temu</span>` : d.days <= 7 ? `<span class="chip c-amb">kończy się za ${d.days} dni</span>` : `<span class="chip c-blue">aktywne · ${d.days} dni</span>` },
  { h: 'Office', v: d => d.demoOffice ? 1 : 0, f: d => d.demoOffice ? '<span class="chip c-br">w biurze</span>' : '' }, { h: 'Kody wydał', key: 'demoAdm' },
  { h: 'Co dalej', v: d => d.next ? 1 : 0, f: d => d.next ? `<span class="chip c-ok">${d.nextKind} #${d.next.id}</span> <span style="font-size:11px;color:#6B7280">${esc(d.next.stage || d.next.status1 || '')}</span>` : (d.active ? '<span style="color:#9CA3AF">w trakcie</span>' : '<span class="chip c-amb">brak szansy</span>') },
  { h: 'Notatki', key: 'nNotes', num: 1 }];
VIEWS['demo-akt'] = () => { const D = demos(), act = D.filter(d => d.active || d.planned), end7 = act.filter(d => d.ending), endedNo = D.filter(d => !d.active && !d.next && d.days != null && d.days >= -30);
  return `<div class="card panel"><div class="kpis">${kpi('Aktywne i zaplanowane', act.length, D.filter(d => d.planned).length + ' zaplanowanych', 'DEMO', '#0D9488', '#E0F5F3')}${kpi('Kończą się w ciągu 7 dni', end7.length, 'czas na follow-up', 'clock', '#D97706', '#FEF3E2')}${kpi('Zakończone (30 dni) bez szansy', endedNo.length, 'brak szansy/TECS po demo', 'alert', '#DC2626', '#FEE2E2')}${kpi('Aktywne bez notatki', act.filter(d => !d.nNotes).length, '', 'note', '#D97706', '#FEF3E2')}</div>
    ${ptitle('DEMO', '#0D9488', '#E0F5F3', 'Aktywne kody demo', 'Opps_TECS z Etap1 = Demo, produkty z Oferty_TECS')}${table('dact', dCols(), act, { k: 7, asc: true }, { empty: 'Brak aktywnych kodów demo' })}</div>
    <div class="card panel">${ptitle('alert', '#DC2626', '#FEE2E2', 'Zakończone w ostatnich 30 dniach bez kolejnej szansy', 'Szansa sprzedaży lub TECS tej samej firmy rozpoczęte po starcie demo')}${table('dend', dCols(), endedNo, { k: 7, asc: false }, { empty: 'Wszystkie zakończone dema mają kolejną szansę 👍' })}</div>`; };
VIEWS['demo-all'] = () => { const D = demos(); return `<div class="card panel">${ptitle('list', '#5B6CDB', '#EEF0FC', 'Wszystkie dema', D.length + ' pozycji')}${table('dall', dCols(), D, { k: 4, asc: false })}</div>`; };
VIEWS['demo-konw'] = () => { const D = demos().filter(d => !d.active && !d.planned);
  const by = [...new Set(D.map(d => d.owner))].map(w => { const x = D.filter(d => d.owner === w), c = x.filter(d => d.next); return { w, n: x.length, c: c.length, r: x.length ? c.length / x.length : 0, avg: x.length ? sum(x, d => d.len || 0) / x.length : 0 }; });
  const prod = {}; D.forEach(d => d.prods.forEach(p => { const k = p.replace(/^Ansys /, ''); const x = prod[k] = prod[k] || { p: k, n: 0, c: 0 }; x.n++; if (d.next) x.c++; }));
  const conv = D.filter(d => d.next).length;
  return `<div class="card panel"><div class="kpis">${kpi('Zakończone dema', D.length, '', 'DEMO', '#0D9488', '#E0F5F3')}${kpi('Z kolejną szansą', conv, D.length ? Math.round(conv / D.length * 100) + '% konwersji' : '', 'OPPSY', '#16A34A', '#E6F6EE')}${kpi('Średnia długość demo', fmt(sum(D, d => d.len || 0) / Math.max(1, D.length)) + ' dni', '', 'clock')}</div>
    <div class="grid2"><div>${ptitle('users', '#5B6CDB', '#EEF0FC', 'Konwersja wg opiekuna', '')}${table('dko', [{ h: 'Opiekun', key: 'w' }, { h: 'Dema', key: 'n', num: 1 }, { h: 'Z szansą', key: 'c', num: 1 }, { h: 'Konwersja', key: 'r', num: 1, f: r => `${Math.round(r.r * 100)}% <span class="bar" style="width:${Math.round(r.r * 60)}px;background:#16A34A"></span>` }, { h: 'Śr. dni', key: 'avg', num: 1, f: r => fmt(r.avg) }], by, { k: 1, asc: false })}</div>
    <div>${ptitle('box', '#D97706', '#FEF3E2', 'Konwersja wg produktu', '')}${table('dkp', [{ h: 'Produkt', key: 'p' }, { h: 'Dema', key: 'n', num: 1 }, { h: 'Z szansą', key: 'c', num: 1 }, { h: 'Konwersja', v: r => r.c / r.n, num: 1, f: r => Math.round(r.c / r.n * 100) + '%' }], Object.values(prod), { k: 1, asc: false })}</div></div></div>`; };

/* ================= start ================= */
function clock() { const d = new Date(); $('#uClock').textContent = d.toLocaleTimeString('pl-PL', { hour: '2-digit', minute: '2-digit' }) + '  ·  ' + d.toLocaleDateString('pl-PL', { weekday: 'short', day: 'numeric', month: 'short' }); }
async function start() {
  clock(); setInterval(clock, 15000);
  try {
    if (MOCK) { ME.name = MOCK.me.displayName; ME.mail = MOCK.me.mail; }
    else {
      msal = new window.msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + CFG.tenantId, redirectUri: location.origin + location.pathname }, cache: { cacheLocation: 'localStorage' } });
      const r = await msal.handleRedirectPromise();
      account = (r && r.account) || msal.getAllAccounts()[0];
      if (!account) { loader('Przekierowuję do logowania Microsoft 365…'); await msal.loginRedirect({ scopes: CFG.scopes }); return; }
      msal.setActiveAccount(account);
      const me = await gfetch('/me?$select=displayName,mail,userPrincipalName'); ME.name = me.displayName; ME.mail = me.mail || me.userPrincipalName;
      try { const t = await token(); const p = await fetch(G + '/me/photos/64x64/$value', { headers: { Authorization: 'Bearer ' + t } }); if (p.ok) { const u = URL.createObjectURL(await p.blob()); $('#uAvatar').innerHTML = `<img src="${u}" alt="">`; } } catch (e) {}
    }
    $('#uName').textContent = ME.name; if (!$('#uAvatar img')) $('#uAvatar').textContent = initials(ME.name);
    await initSite(); await loadAll(); loader(false); fromHash(); render();
  } catch (e) {
    console.error(e); loader(false);
    $('#view').innerHTML = `<div class="card login"><h2>Nie udało się wczytać danych</h2><p style="color:#6B7280">${esc(e.message)}</p><button class="btn pri" onclick="location.reload()">Spróbuj ponownie</button></div>`;
  }
}
$('.brand').onclick = () => go('dashboard');
window.CRM = { DB, ST, render };
start();
})();

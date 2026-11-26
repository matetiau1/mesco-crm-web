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
  chevL:"<path d='m15 18-6-6 6-6'/>",
  chevR:"<path d='m9 18 6-6-6-6'/>",
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
const STAGES = ['Pierwszy kontakt','Discovery','Oferta','Spotkanie techniczne','Demo/Benchmark','Decyzja','Rewizja oferty','Wdrożenie/Realizacja'];
const STAGE_SHORT = ['Pierwszy kontakt','Discovery','Oferta','Spotk. tech.','Demo/Benchmark','Decyzja','Rewizja oferty','Wdrożenie'];
// [pole, etykieta, max] — jak w Power Apps: 0 → częściowo (8) → pełna
const MEDDIC = [['meddicM','Metric',15],['meddicE','Economic Buyer',25],['meddicDc','Decision Criteria',15],['meddicDp','Decision Process',15],['meddicI','Identify Pain',15],['meddicC','Champion',15]];
const medColor = v => v >= 70 ? ['#16A34A','#E6F6EE','c-ok'] : v >= 40 ? ['#D97706','#FEF3E2','c-amb'] : ['#DC2626','#FEE2E2','c-red'];
const DB = { loaded: null };
const ME = { name: '', mail: '', staff: false };
const F = {
  Companies: ['CompaniesGUID','Nazwa','Pe_x0142_nanazwa','Status','Status1','Bran_x017c_a','Creator','TECSh','PreSalesh','WartoscTECS','MarzaTECS','SzkoleniaGratis','SzkolenieGratis','Utworzoneprzez_x002f_OpiekunLookupId','NumerAnsys'],
  Notes: ['Firma','Osoba','TypeLookupId','Rodzaj1','Totatka','Data','Follow_up','FollowUpStatus','Creator','Telefon','LiczbaPolaczen','Opps_ID','IdentyfikatorCompanies','IdentyfikatorContact','Etap','EtapProspecting','DodajFollowUp','Created'],
  Opps: ['Nazwaszansysprzeda_x017c_y','Firma','Status1','Etap1','Etap','Warto_x015b__x0107_wEUR','Spodziewanezamkni_x0119_cie','ClosedDate','Dataotwarcia','Rodzajbiznesu','TypBiznesu','IdentyfikatorCompanies','OPPSGUID','Creator','OpiekunOppsaLookupId',
    'meddicM','meddicE','meddicDc','meddicDp','meddicI','meddicC','FTDate','TerminDR','DISCH','DISB','DISP','OFP','OFSF','OFDate','QuoteNumber','QuoteNumberLink','TMFU','TMEB','OFDP','TMDC','TMM','DBDE','DBDate','DSM','DEB','RTL','RWH','RCR','WRDL','WROF','WRO','LinkSalesForce','DR_Odrzucony','Konkurencja','BrakDofinansowania','Zamkni_x0119_tyOppsNotes','Notatki'],
  Oferty: ['IdentyfikatorOpps','Suma','ProductID','Liczba_licencji','Paidup','TECS','Lease','Lease_Days','AEC','ALH','Rabat','PaidUP0','LeaseDays','AECYN','ALHYN'],
  Opps_TECS: ['Numer_klienta','Nazwaszansysprzeda_x017c_y_TECS','Pocz_x0105_tek_TECS','Koniec_TECS','OpiekunLookupId','Warto_x015b__x0107_','FirmaID','Etap1','TP','IdentyfikatorCompanies','TECSGUID','Demo','DemoOffice','DemoAdmLookupId','Data_Follow_Up','TECScreator'],
  Oferty_TECS: ['ProducID','Cena_TECS','Liczba_licencji','SUMA','Rabat','IdentyfikatorTECS','TP','TPValue','Demo','FirmaIDofertaTECS','Pocz_x0105_tek_TECS','Koniec_TECS'],
  Cennik_CRM: ['Title','field_2','field_3'],
  Support: ['Data','Czas','IdentyfikatorCompanies'],
  UczestnicySzkolen: ['GUIDFirmy','DataSzkolenia','WartoscDarmowegoSzkolenia'],
  Szkolenia: ['GUIDFirmy','DataSzkolenia','WartoscDarmowegoSzkolenia'],
  KursyWalut: ['Waluta','Kurs'],
  activities: ['Notes_types'],
  MEScoStaff: ['Imi_x0119_iNazwisko','Email'],
  MarketingKontakty: ['Imie','Nazwisko','Email','Telefon','Firma','Zrodlo','GUIDWebinar','GUIDKontaktu','Obecny','KonwertLead','FollowUp'],
  Contacts: ['Email','ContactGUID','IdentyfikatorCompanies','FirmaID','Imi_x0119_','Nazwisko','Stanowisko'],
  Webinary: ['Tytul','GUIDWebinar','DataWebinaru','Status','GrupaDocelowa'],
  Leady: ['Firma','Imi_x0119_inazwisko','Data','ZrodloLookupId','OpiekunLookupId','OstatniaDataFollowUp','FollowUp','JakTrafil','Efekt','Status','Wartosc','Komentarz'],
  Source: ['Title'],
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
    id: c.id, guid: lc(c.CompaniesGUID), guidRaw: c.CompaniesGUID || '', name: (c.Nazwa || c.Pe_x0142_nanazwa || '').trim() || '(bez nazwy)',
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
  DB.offers = {}; R.Oferty.forEach(o => { const g = lc(o.IdentyfikatorOpps); if (g) (DB.offers[g] = DB.offers[g] || []).push(o); });
  const offBy = {}; R.Oferty.forEach(o => { const g = lc(o.IdentyfikatorOpps); if (!g) return; const x = offBy[g] = offBy[g] || { n: 0, s: 0 }; x.n++; x.s += num(o.Suma) || 0; });
  // szanse sprzedaży
  const mkSale = o => {
    const g = lc(o.OPPSGUID), ni = nInfo(g), of = offBy[g] || { n: 0, s: 0 };
    const close = toD(o.Spodziewanezamkni_x0119_cie), closedWin = toD(o.ClosedDate);
    const st = o.Status1 || 'Aktywny';
    const x = { id: o.id, guid: g, name: o.Nazwaszansysprzeda_x017c_y || '', compGuid: lc(o.IdentyfikatorCompanies), comp: cname(o.IdentyfikatorCompanies, o.Firma), ac: cac(o.IdentyfikatorCompanies),
      status1: st, stage: o.Etap1 || o.Etap || '', eur: num(o.Warto_x015b__x0107_wEUR) || 0, close, closedWin, openD: toD(o.Dataotwarcia), rodz: o.Rodzajbiznesu || '', typ: o.TypBiznesu || '',
      creator: o.Creator || '', owner: uname(o.OpiekunOppsaLookupId) || o.Creator || '', nOffers: of.n, offerSum: of.s, nNotes: ni.n, lastNote: ni.last, raw: o };
    x.med = MEDDIC.map(m => num(o[m[0]]) || 0); x.meddic = x.med.reduce((a, b) => a + b, 0);
    x.grp = st === 'Zamknięty wygrany' ? 'won' : st === 'Zamknięty przegrany' ? 'lost' : 'open';
    x.value = x.eur; x.fq = fyQ(x.grp === 'open' ? close : (closedWin || close)); x.days = daysTo(close);
    x.idle = x.lastNote ? Math.round((TODAY - x.lastNote) / 864e5) : (x.openD ? Math.round((TODAY - x.openD) / 864e5) : null);
    x.flags = [];
    if (x.grp === 'open') {
      if (!close) x.flags.push(['Brak daty zamknięcia', 'c-amb']); else if (x.days < 0) x.flags.push(['Po terminie', 'c-red']);
      if (['Oferta','Rewizja oferty','Decyzja'].includes(x.stage) && !x.nOffers) x.flags.push(['Bez oferty', 'c-red']);
      if (STAGES.indexOf(x.stage) >= 5 && x.meddic < 40) x.flags.push([`MEDDIC ${x.meddic}%`, 'c-amb']);
      if (!x.nNotes) x.flags.push(['Bez notatek', 'c-amb']); else if (x.idle > 60) x.flags.push([`Brak notatki ${x.idle} dni`, 'c-amb']);
    }
    // marża / premie wg Power Apps
    const mz = { 'New Business': .35, 'B2TECS': .35, 'New Logo': .45 }[x.rodz] || 0;
    const mzP = { 'New Business': .35, 'B2TECS': .35, 'New Logo': .35 }[x.rodz] || 0;
    const st2 = x.rodz === 'New Logo' ? .05 : .025;
    x.marza = x.eur * mz; x.premQ = x.eur * mzP * st2 * .4; x.premR = x.eur * mzP * st2 * .6;
    return x;
  };
  DB.mkSale = mkSale; DB.sales = R.Opps.map(mkSale);
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
      raw: o, demoFlag: !!o.Demo, demoOffice: !!o.DemoOffice, demoAdm: uname(o.DemoAdmLookupId), follow: toD(o.Data_Follow_Up), nrAnsys: o.Numer_klienta || '', lines: ls };
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
  // marketing: kontakty CRM po e-mailu, webinary, zapisy
  DB.contacts = R.Contacts.map(c => ({ id: c.id, guid: c.ContactGUID || '', mail: lc(c.Email), compGuid: lc(c.IdentyfikatorCompanies), firma: c.FirmaID || '', name: ((c.Imi_x0119_ || '') + ' ' + (c.Nazwisko || '')).trim(), stan: c.Stanowisko || '' }));
  DB.contactByMail = {}; DB.contactByGuid = {}; DB.contacts.forEach(c => { if (c.mail) DB.contactByMail[c.mail] = c; if (c.guid) DB.contactByGuid[lc(c.guid)] = c; });
  DB.webinars = R.Webinary.map(w => ({ id: w.id, guid: lc(w.GUIDWebinar), title: (w.Tytul || '').trim() || '(bez tytułu)', date: toD(w.DataWebinaru), status: w.Status || '', grupa: w.GrupaDocelowa || '' }));
  DB.webBy = {}; DB.webinars.forEach(w => DB.webBy[w.guid] = w);
  DB.mkt = R.MarketingKontakty.map(m => { const mail = lc(m.Email), c = (m.GUIDKontaktu && DB.contactByGuid[lc(m.GUIDKontaktu)]) || DB.contactByMail[mail] || null;
    return { id: m.id, mail, name: ((m.Imie || '') + ' ' + (m.Nazwisko || '')).replace(/\s+/g, ' ').trim(), tel: m.Telefon || '', firmaTxt: (m.Firma || '').trim(), zrodlo: (m.Zrodlo || '').trim(), web: DB.webBy[lc(m.GUIDWebinar)] || { guid: lc(m.GUIDWebinar), title: '(nieznane wydarzenie)', date: null },
      obecny: m.Obecny === 'Tak' ? true : m.Obecny === 'Nie' ? false : null, guidK: m.GUIDKontaktu || '', contact: c, lead: !!m.KonwertLead }; });
  // leady
  const src = {}; (R.Source || []).forEach(x => src[x.id] = x.Title || '');
  const notesByName = {}; DB.notes.forEach(n => { if (n.firma && n.data) (notesByName[lc(n.firma)] = notesByName[lc(n.firma)] || []).push(n.data); });
  DB.leads = (R.Leady || []).map(l => { const data = toD(l.Data), fu = toD(l.OstatniaDataFollowUp), firma = (l.Firma || '').trim();
    const nd = (notesByName[lc(firma)] || []).filter(d => !data || d >= data);
    const last = [data, fu, ...nd].filter(Boolean).sort((a, b) => b - a)[0] || null;
    return { id: l.id, firma, osoba: (l.Imi_x0119_inazwisko || '').trim(), data, fu, status: l.Status || '', efekt: l.Efekt || '', jak: l.JakTrafil || '', zrodlo: src[+l.ZrodloLookupId] || '',
      owner: uname(l.OpiekunLookupId), wart: num(l.Wartosc), last, idle: last ? Math.round((TODAY - last) / 864e5) : null, raw: l }; });
  // ostatni kontakt (dowolna notatka) per firma
  DB.lastByComp = {}; DB.notes.forEach(n => { if (!n.data) return; [n.guid, 'n:' + lc(n.firma)].forEach(k => { if (k && k !== 'n:' && (!DB.lastByComp[k] || n.data > DB.lastByComp[k])) DB.lastByComp[k] = n.data; }); });
  DB.quarters = [...new Map(DB.tecs.concat(DB.sales).filter(o => o.fq).map(o => [o.fq.key, o.fq])).values()].sort((a, b) => a.key - b.key);
}

/* ================= stan + routing ================= */
const ST = { mod: 'dashboard', tab: 'todo', sort: {}, g: { owner: '', comp: '', q: '', seg: 'all' },
  todo: { plus7: false, desc: false }, plan: { who: 'me', ciche: 14 }, sp: { who: 'me', seg: 'klient', dni: 365 }, akt: { od: new Date(TODAY.getFullYear(), TODAY.getMonth(), 1), do: TODAY, h: '', r: '', f: '' },
  rs: { tryb: 'szanse', od: null, do: null, desc: true, h: '', f: '', r: '', t: '', s: '', to: '', te: '', tf: '', pf: '', pl: '', onlyTP: false },
  hg: { pre: false, all: false, st: '', op: '' }, mk: { q: '', crm: '', min: 1, att: '', web: '' }, df: { all: false, op: '', st: '', f: '', lim: 60 }, sstage: '' };
try { const s = JSON.parse(localStorage.getItem('mescocrm') || '{}'); if (s.g) Object.assign(ST.g, s.g); } catch (e) {}
const save = () => { try { localStorage.setItem('mescocrm', JSON.stringify({ g: ST.g })); } catch (e) {} };
const MODS = {
  dashboard: { name: 'Dashboard', icon: 'DASHBOARD', tabs: [['todo','To-Do','clock'],['plan','Plan na tydzień','KALENDARZ'],['spotk','Bez spotkań 12 mies.','users'],['akt','Aktywności','activity'],['sprz','Raport sprzedaży','OPPSY'],['godz','Godziny TECS / PreSales','bars'],['firmy','Bilans firm','FIRMY']] },
  tecs: { name: 'TECS — odnowienia', icon: 'TECSY', g: true, tabs: [['przeglad','Przegląd','DASHBOARD'],['otwarte','Do zamknięcia','clock'],['zamkniete','Zamknięte','check'],['ceny','Kontrola cen','euro'],['tp','TP','tag'],['bez-oferty','Bez oferty','alert'],['bez-notatek','Bez notatek','note']] },
  oppsy: { name: 'Szanse sprzedaży', icon: 'OPPSY', g: true, tabs: [['s-przeglad','Przegląd','DASHBOARD'],['s-otwarte','W toku','clock'],['s-zamkniete','Wygrane / przegrane','check'],['s-alerty','Alerty','alert']] },
  firmy: { name: 'Klienci — wykorzystanie TECS', icon: 'FIRMY', g: true, noQ: true, tabs: [['klienci','Aktywne TECS i support','SUPPORTY']] },
  marketing: { name: 'Marketing — webinary i wydarzenia', icon: 'MARKETING', tabs: [['m-osoby','Uczestnicy','users'],['m-firmy','Firmy','FIRMY'],['m-web','Wydarzenia','KALENDARZ']] },
  demo: { name: 'Kody demo', icon: 'DEMO', g: true, noQ: true, tabs: [['demo-akt','Aktywne i kończące się','clock'],['demo-all','Wszystkie demo','list'],['demo-konw','Konwersja','OPPSY']] },
};
const NAV = [['HR',null],['FIRMY','firmy'],['KONTAKTY',null],['KALENDARZ',null],['CENNIK',null],['SZKOLENIA',null],['OPPSY','oppsy'],['TECSY','tecs'],['DEMO','demo'],['PROSPECTING',null],['SUPPORTY',null],['ZADANIA',null],['MARKETING','marketing'],['DASHBOARD','dashboard'],['ADMIN',null],['USTAWIENIA',null]];
function fromHash() { const [m, t, id] = location.hash.slice(1).split('/').map(decodeURIComponent);
  if (m === 'oppsy' && t === 'opp' && id) { ST.mod = m; ST.tab = 'opp'; ST.oppId = +id; return; }
  if (m === 'oppsy' && t === 's-otwarte') ST.sstage = id || '';
  if (MODS[m]) { ST.mod = m; ST.tab = MODS[m].tabs.some(x => x[0] === t) ? t : MODS[m].tabs[0][0]; } }
window.addEventListener('hashchange', () => { fromHash(); render(); });
const go = (m, t, x) => { location.hash = m + '/' + (t || MODS[m].tabs[0][0]) + (x ? '/' + encodeURIComponent(x) : ''); };
// powrót: poprzedni ekran aplikacji (jak Back() w Power Apps), a gdy go nie ma — lista szans
let navDepth = 0; window.addEventListener('hashchange', () => { navDepth++; window.scrollTo(0, 0); });
const goBack = fb => { if (navDepth > 0) { navDepth -= 2; history.back(); } else fb(); };

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
  const M = MODS[mod]; if (!M.g || ST.tab === 'opp') return '';
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
  if (ST.mod === 'dashboard') { const t = todoList(); c.todo = t.length; c['todo!'] = t.some(n => n.fu < TODAY); const P = planData(); c.plan = P.total; c['plan!'] = P.pilne > 0; const Q = spData(); c.spotk = Q.bez.length; c['spotk!'] = Q.bez.some(x => x.tecs); }
  if (ST.mod === 'tecs') { const L = gf(DB.tecs).filter(o => !['demo','tp'].includes(o.grp)); c.otwarte = L.filter(o => o.grp === 'open').length; c['bez-oferty'] = L.filter(o => o.grp === 'open' && !o.nOffers).length; c['bez-oferty!'] = 1;
    c['bez-notatek'] = L.filter(o => o.grp === 'open' && !o.nNotes && o.days != null && o.days <= 90).length; c['bez-notatek!'] = 1; c.ceny = L.reduce((a, o) => a + o.priceIssues.filter(p => p.diff != null).length, 0); c['ceny!'] = 1; c.tp = gf(DB.tecs).filter(o => o.grp === 'tp').length; }
  if (ST.mod === 'oppsy') { const S = gf(DB.sales); c['s-otwarte'] = S.filter(o => o.grp === 'open').length; c['s-alerty'] = S.filter(o => o.grp === 'open' && o.flags.some(f => f[1] === 'c-red')).length; c['s-alerty!'] = 1; }
  if (ST.mod === 'firmy') { c.klienci = clients().filter(x => x.alert).length; c['klienci!'] = 1; }
  if (ST.mod === 'marketing') { const P = mktPeople(); c['m-osoby'] = P.length; c['m-web'] = DB.webinars.filter(w => DB.mkt.some(m => m.web === w)).length; }
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

/* --- Plan na tydzień: zadania wyliczone z danych --- */
const weekEnd = () => addDays(TODAY, (7 - TODAY.getDay()) % 7); // najbliższa niedziela
function planData() {
  const P = ST.plan, who = P.who === 'me' ? ME.name : P.who;
  const mine = o => !who || o.owner === who;
  const opps = DB.sales.filter(o => o.grp === 'open' && mine(o) && o.close && o.days <= 30).sort((a, b) => a.days - b.days);
  const tecs = DB.tecs.filter(o => o.grp === 'open' && mine(o) && o.end && o.days <= 90).sort((a, b) => a.days - b.days);
  const fu = DB.notes.filter(n => !n.fuDone && n.fu && n.fu < TODAY && (!who || n.creator === who)).sort((a, b) => a.fu - b.fu);
  // „bez odpowiedzi”: oferta wysłana, a od ostatniego kontaktu (notatki lub daty oferty) minęło ≥ N dni
  const since = (last, sent) => { const d = [last, sent].filter(Boolean).sort((a, b) => b - a)[0]; return d ? Math.round((TODAY - d) / 864e5) : null; };
  const offS = DB.sales.filter(o => o.grp === 'open' && mine(o) && (o.raw.OFP || (['Oferta','Rewizja oferty','Decyzja'].includes(o.stage) && o.nOffers)))
    .map(o => ({ kind: 'opp', o, sent: toD(o.raw.OFDate), cisza: since(o.lastNote, toD(o.raw.OFDate)) }));
  const offT = DB.tecs.filter(o => o.stage === 'Oferta wysłana' && mine(o) && !(o.days != null && o.days < -60))
    .map(o => ({ kind: 'tecs', o, sent: null, cisza: since(o.lastNote, null) }));
  const offers = offS.concat(offT).filter(x => x.cisza == null || x.cisza >= P.ciche).sort((a, b) => (b.cisza ?? 9999) - (a.cisza ?? 9999));
  const we = weekEnd(), inWeek = d => d && d >= TODAY && d <= we;
  const pilne = opps.filter(o => o.days < 0).length + tecs.filter(o => o.days < 0).length + fu.length;
  // „Do zrobienia” — szybkie listy
  const openS = DB.sales.filter(o => o.grp === 'open' && mine(o));
  openS.forEach(o => { const d = [o.lastNote, DB.lastByComp[o.compGuid]].filter(Boolean).sort((a, b) => b - a)[0] || o.openD; o.mv = d ? Math.round((TODAY - d) / 864e5) : null; });
  const stale = openS.filter(o => o.mv != null && o.mv > 30);
  const overdue = openS.filter(o => o.close && o.days < 0);
  const kl = DB.comp.filter(c => c.status1 === 'Klient' && (!who || c.opiekun === who)).map(c => { const d = [DB.lastByComp[c.guid], DB.lastByComp['n:' + lc(c.name)]].filter(Boolean).sort((a, b) => b - a)[0] || null;
    return { c, last: d, ago: d ? Math.round((TODAY - d) / 864e5) : null }; });
  const zan = kl.filter(x => x.ago == null || x.ago > 90);
  const leadsOpen = DB.leads.filter(l => ['Nowy','Otwarty'].includes(l.status) && (!who || l.owner === who));
  const leadsCold = leadsOpen.filter(l => l.idle == null || l.idle > 14);
  const tecsFu = DB.tecs.filter(o => o.follow && o.follow < TODAY && !['Zamówione','Wystawiona faktura','Rezygnacja z TECS'].includes(o.stage) && mine(o));
  return { who, opps, tecs, fu, offers, we, inWeek, pilne, total: opps.length + tecs.length + fu.length + offers.length,
    openS, stale, overdue, kl, zan, leadsOpen, leadsCold, tecsFu };
}
const dueCell = (d, days, inWeek) => !d ? '<span class="chip c-amb">brak daty</span>' : `<span style="white-space:nowrap">${dpl(d)}</span> ${days < 0 ? `<span class="chip c-red">${-days} dni po</span>` : days === 0 ? '<span class="chip c-red">dziś</span>' : inWeek ? `<span class="chip c-amb">ten tydzień</span>` : `<span class="chip c-mut">za ${days} dni</span>`}`;
const actBtns = (kind, id) => ME.staff ? `<button class="btn" data-pnote="${kind}|${id}" title="Dodaj notatkę">${svg('note')}</button>` : '';
VIEWS.plan = () => {
  const P = planData(), S = ST.plan;
  const lastC = d => d ? `${dpl(d)} <span style="color:#9CA3AF">(${Math.round((TODAY - d) / 864e5)} dni temu)</span>` : '<span class="chip c-amb">brak</span>';
  const opC = [{ h: 'Szansa', key: 'name', f: o => `${oppLink(o)}<div style="font-size:11px;color:#9CA3AF">${esc(o.comp)}</div>` }, { h: 'Etap', v: o => STAGES.indexOf(o.stage), f: o => `<span class="chip c-br">${esc(o.stage || '—')}</span>` },
    { h: 'Zamknięcie', key: 'days', f: o => dueCell(o.close, o.days, P.inWeek(o.close)) }, { h: 'Wartość', key: 'eur', num: 1, f: o => eur(o.eur) }, { h: 'MEDDIC', key: 'meddic', num: 1, f: medChip },
    { h: 'Ostatni kontakt', v: o => o.lastNote, f: o => lastC(o.lastNote) }, { h: 'Co zrobić', v: o => o.flags.length, f: o => chips(o.flags) || (o.days < 0 ? '' : '<span style="color:#9CA3AF">potwierdź termin / decyzję</span>') }, { h: '', f: o => actBtns('opp', o.id) }];
  if (!P.who) opC.splice(2, 0, { h: 'Opiekun', key: 'owner' });
  const tcC = [{ h: 'TECS', key: 'comp', f: o => `${compLink(o.comp, o.compGuid)}<div style="font-size:11px;color:#9CA3AF">${tecsLink(o)} ${esc(o.name)}</div>` }, { h: 'Etap', key: 'stage', f: o => stageChip(o.stage) },
    { h: 'Koniec TECS', key: 'days', f: o => dueCell(o.end, o.days, P.inWeek(o.end)) }, { h: 'Wartość', key: 'value', num: 1, f: o => eur(o.value) }, { h: 'Oferty', key: 'nOffers', num: 1 },
    { h: 'Ostatni kontakt', v: o => o.lastNote, f: o => lastC(o.lastNote) }, { h: 'Co zrobić', v: o => o.flags.length, f: o => chips(o.flags) || '<span style="color:#9CA3AF">wyślij / domknij odnowienie</span>' }, { h: '', f: o => actBtns('tecs', o.id) }];
  if (!P.who) tcC.splice(2, 0, { h: 'Opiekun', key: 'owner' });
  const fuC = [{ h: 'Follow-up', key: 'fu', f: n => `${dpl(n.fu)} <span class="chip c-red">${Math.round((TODAY - n.fu) / 864e5)} dni po</span>` }, { h: 'Firma', key: 'firma', f: n => compLink(n.firma, n.guid) }, { h: 'Osoba', key: 'osoba' },
    { h: 'Rodzaj', key: 'rodzaj', f: n => `<span class="chip c-br">${esc(n.rodzaj || '—')}</span>` }, { h: 'Notatka', key: 'text', f: n => `<span title="${esc(n.text)}" style="color:#6B7280">${esc(String(n.text).replace(/\n/g, ' ').slice(0, 110))}</span>` },
    { h: '', f: n => ME.staff ? `<span style="white-space:nowrap"><button class="btn ok" data-pdone="${n.id}" title="Oznacz jako zrobiony">✓</button> <button class="btn" data-pfu="${n.id}" title="Notatka + kolejny follow-up">${svg('note')}</button></span>` : '' }];
  if (!P.who) fuC.splice(1, 0, { h: 'Kto', key: 'creator' });
  const ofC = [{ h: 'Co', v: x => x.kind, f: x => x.kind === 'opp' ? '<span class="chip c-br">Szansa</span>' : '<span class="chip c-vio">TECS</span>' },
    { h: 'Szansa / firma', v: x => x.o.comp, f: x => x.kind === 'opp' ? `${oppLink(x.o)}<div style="font-size:11px;color:#9CA3AF">${esc(x.o.comp)}</div>` : `${compLink(x.o.comp, x.o.compGuid)}<div style="font-size:11px;color:#9CA3AF">${tecsLink(x.o)} ${esc(x.o.name)}</div>` },
    { h: 'Etap', v: x => x.o.stage, f: x => stageChip(x.o.stage) }, { h: 'Wartość', v: x => x.o.value, num: 1, f: x => eur(x.o.value) }, { h: 'Oferta wysłana', v: x => x.sent, f: x => x.sent ? dpl(x.sent) : '<span style="color:#9CA3AF">—</span>' },
    { h: 'Ostatni kontakt', v: x => x.o.lastNote, f: x => lastC(x.o.lastNote) }, { h: 'Cisza', v: x => x.cisza ?? 9999, num: 1, f: x => x.cisza == null ? '<span class="chip c-red">brak kontaktu</span>' : `<span class="chip ${x.cisza >= 30 ? 'c-red' : 'c-amb'}">${x.cisza} dni</span>` }, { h: '', f: x => actBtns(x.kind, x.o.id) }];
  if (!P.who) ofC.splice(2, 0, { h: 'Opiekun', v: x => x.o.owner, f: x => esc(x.o.owner) });
  const sec = (icon, c, bg, t, sub, body) => `<div class="card panel">${ptitle(icon, c, bg, t, sub)}${body}</div>`;
  const owners = DB.owners;
  return `<div class="card panel">${ptitle('KALENDARZ', '#5B6CDB', '#EEF0FC', 'Plan na tydzień', `${dpl(TODAY)} – ${dpl(P.we)} · zadania wyliczone automatycznie z szans, TECS i notatek`,
    `<select id="plWho"><option value="me" ${S.who === 'me' ? 'selected' : ''}>Moje (${esc(ME.name)})</option><option value="" ${S.who === '' ? 'selected' : ''}>Cały zespół</option>${owners.filter(o => o !== ME.name).map(o => `<option ${S.who === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`)}
    <div class="kpis">${kpi('Oppsy do zamknięcia ≤30 dni', P.opps.length, `${eur(sum(P.opps, o => o.eur))} · ${P.opps.filter(o => o.days < 0).length} po terminie`, 'OPPSY', '#5B6CDB', '#EEF0FC')}
      ${kpi('TECSy do zrobienia (≤3 mies.)', P.tecs.length, `${eur(sum(P.tecs, o => o.value))} · ${P.tecs.filter(o => o.days < 0).length} po terminie`, 'TECSY', '#0D9488', '#E0F5F3')}
      ${kpi('Zaległe follow-upy', P.fu.length, P.fu.length ? 'najstarszy z ' + dpl(P.fu[0].fu) : 'wszystko zrobione', 'clock', '#DC2626', '#FEE2E2')}
      ${kpi('Oferty bez odpowiedzi', P.offers.length, `cisza ≥ ${S.ciche} dni`, 'box', '#D97706', '#FEF3E2')}</div></div>
  ${doZrobienia(P)}
  ${sec('OPPSY', '#5B6CDB', '#EEF0FC', 'Szanse sprzedaży — zamknięcie w ciągu 30 dni', 'Status1 = Aktywny, Spodziewane zamknięcie ≤ dziś + 30 dni (także te po terminie)', table('pl1', opC, P.opps, { k: P.who ? 2 : 3, asc: true }, { empty: 'Brak szans z zamknięciem w najbliższych 30 dniach' }))}
  ${sec('TECSY', '#0D9488', '#E0F5F3', 'TECSy do zrobienia — koniec w ciągu 3 miesięcy', 'Etap: Przygotowanie oferty / Oferta wysłana, Koniec TECS ≤ dziś + 90 dni', table('pl2', tcC, P.tecs, { k: P.who ? 2 : 3, asc: true }, { empty: 'Brak TECS do odnowienia w najbliższych 3 miesiącach' }))}
  ${sec('clock', '#DC2626', '#FEE2E2', 'Zaległe follow-upy', 'Follow-up przed dzisiejszą datą, nieoznaczony jako zrobiony', table('pl3', fuC, P.fu, { k: 0, asc: true }, { empty: 'Brak zaległych follow-upów 🎉' }))}
  ${sec('box', '#D97706', '#FEF3E2', 'Wysłane oferty bez odpowiedzi', `Szanse z „oferta wysłana” (OFP) lub na etapie Oferta/Rewizja/Decyzja z ofertą + TECS „Oferta wysłana”, bez notatki od ≥ <select id="plCisza" style="height:24px;font-size:11.5px">${[7, 14, 21, 30].map(d => `<option ${S.ciche === d ? 'selected' : ''}>${d}</option>`).join('')}</select> dni`, table('pl4', ofC, P.offers, { k: P.who ? 6 : 7, asc: false }, { empty: 'Brak ofert czekających na odpowiedź' }))}`;
};
/* --- „Do zrobienia”: 5 krótkich list (maks. 15 pozycji) --- */
const vK = v => v ? fmtK(v) + ' €' : '<span style="color:#9CA3AF">–</span>';
const agoD = d => d == null ? '<span class="chip c-red">nigdy</span>' : `<span class="mono">${d} d</span>`;
function doZrobienia(P) {
  const card = (icon, c, bg, t, sub, id, cols, rows, def, empty) => `<div class="card panel dz">${ptitle(icon, c, bg, t, sub)}${table(id, cols, rows, def, { limit: 15, empty })}</div>`;
  const firmaO = o => `<a class="lnk" href="#oppsy/opp/${o.id}" title="${esc(o.name)}">${esc(o.comp)}</a>`;
  const stC = [{ h: 'Firma', v: o => o.comp, f: firmaO }, { h: 'Wartość', key: 'eur', num: 1, f: o => vK(o.eur) }, { h: 'Etap', key: 'stage', f: o => esc(o.stage || '—') }, { h: 'Bez ruchu', key: 'mv', num: 1, f: o => agoD(o.mv) }];
  const ovC = [{ h: 'Firma', v: o => o.comp, f: firmaO }, { h: 'Wartość', key: 'eur', num: 1, f: o => vK(o.eur) }, { h: 'Etap', key: 'stage', f: o => esc(o.stage || '—') }, { h: 'Termin', key: 'close', num: 1, f: o => `<span class="mono">${ds(o.close)}</span>` }];
  const zC = [{ h: 'Klient', v: x => x.c.name, f: x => compLink(x.c.name, x.c.guid) }, { h: 'Ostatni kontakt', v: x => x.ago ?? 99999, num: 1, f: x => agoD(x.ago) }];
  const lC = [{ h: 'Firma', key: 'firma', f: l => `<a class="lnk" href="${SPU}Leady/DispForm.aspx?ID=${l.id}" target="_blank" rel="noopener">${esc(l.firma || '—')}</a>${l.osoba ? `<div style="font-size:11px;color:#9CA3AF">${esc(l.osoba)}</div>` : ''}` },
    { h: 'Status', key: 'status' }, { h: 'Bez reakcji', v: l => l.idle ?? 99999, num: 1, f: l => l.idle == null ? '<span class="chip c-red">brak daty</span>' : `<span class="mono ${l.idle > 14 ? 'red' : ''}">${l.idle} d</span>` },
    { h: 'Efekt / źródło', v: l => l.efekt || l.zrodlo || l.jak, f: l => esc([l.efekt, l.zrodlo || l.jak].filter(Boolean).join(' · ') || '—') }, { h: 'Wartość', key: 'wart', num: 1, f: l => l.wart ? fmtK(l.wart) : '<span style="color:#9CA3AF">–</span>' }];
  const tC = [{ h: 'Firma', key: 'comp', f: o => compLink(o.comp, o.compGuid) }, { h: 'Wartość', key: 'value', num: 1, f: o => vK(o.value) }, { h: 'Etap', key: 'stage', f: o => esc(o.stage || '—') }, { h: 'Follow-up', key: 'follow', num: 1, f: o => `<span class="mono">${ds(o.follow)}</span>` }];
  const who = P.who ? '' : ' (cały zespół)';
  return `<div class="dz-h"><h2>Do zrobienia${who}</h2></div><div class="dzgrid">
    ${card('OPPSY', '#D97706', '#FEF3E2', 'Zastałe szanse', `${P.stale.length} z ${P.openS.length} bez ruchu od ponad 30 dni · ${fmtK(sum(P.stale, o => o.eur))} €`, 'dz1', stC, P.stale, { k: 1, asc: false }, 'Wszystkie szanse mają ruch w ostatnich 30 dniach')}
    ${card('KALENDARZ', '#DC2626', '#FEE2E2', 'Szanse po terminie', `${P.overdue.length} z datą zamknięcia w przeszłości · ${fmtK(sum(P.overdue, o => o.eur))} €`, 'dz2', ovC, P.overdue, { k: 1, asc: false }, 'Brak szans po terminie')}
    ${card('users', '#5B6CDB', '#EEF0FC', 'Zaniedbani klienci', `${P.zan.length} z ${P.kl.length} klientów bez kontaktu od ponad 90 dni`, 'dz3', zC, P.zan, { k: 1, asc: false }, 'Każdy klient miał kontakt w ostatnich 90 dniach')}
    ${card('phone', '#0D9488', '#E0F5F3', 'Leady bez reakcji', `${P.leadsCold.length} z ${P.leadsOpen.length} otwartych leadów bez reakcji od ponad 14 dni`, 'dz4', lC, P.leadsOpen, { k: 2, asc: false }, 'Brak otwartych leadów')}
    ${card('TECSY', '#7C3AED', '#F1EAFE', 'TECS: zaległy follow-up', `${P.tecsFu.length} pozycji z minionym terminem follow-up`, 'dz5', tC, P.tecsFu, { k: 3, asc: true }, 'Brak zaległych follow-upów TECS')}
  </div><div class="desc" style="margin:6px 2px 14px">Listy pokazują do 15 pozycji, domyślnie od największej wartości lub najdłuższego braku kontaktu (kliknij nagłówek, by zmienić sortowanie). Zaległe follow-upy z notatek: ${P.fu.length}.</div>`;
}
AFTER.plan = () => {
  $('#plWho').onchange = e => { ST.plan.who = e.target.value; render(); };
  $('#plCisza').onchange = e => { ST.plan.ciche = +e.target.value; render(); };
  document.querySelectorAll('[data-pdone]').forEach(b => b.onclick = async () => { const n = DB.notes.find(x => x.id === +b.dataset.pdone); b.disabled = true;
    try { await patchItem('Notes', n.id, { FollowUpStatus: true }); n.fuDone = true; toast('Follow-up oznaczony jako zrobiony', 'ok'); render(); } catch (e) { b.disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); } });
  document.querySelectorAll('[data-pfu]').forEach(b => b.onclick = () => noteModal(DB.notes.find(x => x.id === +b.dataset.pfu)));
  document.querySelectorAll('[data-pnote]').forEach(b => b.onclick = () => { const [k, id] = b.dataset.pnote.split('|');
    if (k === 'opp') { const x = DB.sales.find(o => o.id === +id); noteModal(oppNoteSrc(x)); }
    else { const o = DB.tecs.find(t => t.id === +id); noteModal({ firma: o.comp, guid: o.compGuid, osoba: '', tel: '', contact: '', etap: o.stage, etapP: '', opps: o.guid, oppName: o.name || 'TECS', raw: { IdentyfikatorCompanies: o.raw.IdentyfikatorCompanies || '', Opps_ID: o.raw.TECSGUID || '' } }); } });
};

/* --- Klienci bez spotkania online / F2F w ostatnim roku --- */
const isMeet = n => /^spotkanie/i.test(n.rodzaj || '');
function spData() {
  const S = ST.sp, who = S.who === 'me' ? ME.name : S.who, od = addDays(TODAY, -S.dni);
  const byG = {}, byN = {}; DB.notes.forEach(n => { if (n.guid) (byG[n.guid] = byG[n.guid] || []).push(n); else if (n.firma) (byN[lc(n.firma)] = byN[lc(n.firma)] || []).push(n); });
  const L = DB.comp.filter(c => {
    const tecs = DB.tecs.filter(o => o.compGuid === c.guid && o.grp !== 'demo' && o.stage !== 'Rezygnacja z TECS' && o.start && o.end && o.start <= TODAY && o.end >= TODAY);
    c._tecs = tecs; c._opps = DB.sales.filter(o => o.compGuid === c.guid && o.grp === 'open');
    const segOk = S.seg === 'klient' ? c.status1 === 'Klient' : S.seg === 'tecs' ? tecs.length > 0 : c.status1 !== 'Prospect';
    return segOk && (!who || c.opiekun === who);
  }).map(c => {
    const ns = (byG[c.guid] || []).concat(byN[lc(c.name)] || []).filter(n => n.data).sort((a, b) => b.data - a.data);
    const meets = ns.filter(isMeet), lastM = meets[0] || null, last = ns[0] || null;
    return { c, tecs: c._tecs, tecsVal: sum(c._tecs, o => o.value), tecsEnd: c._tecs.length ? new Date(Math.min(...c._tecs.map(o => +o.end))) : null, opps: c._opps, oppsVal: sum(c._opps, o => o.eur),
      lastM, lastMd: lastM ? lastM.data : null, last, meet12: meets.filter(n => n.data >= od).length, nMeets: meets.length, nF2F: meets.filter(n => /f2f/i.test(n.rodzaj)).length };
  });
  const bez = L.filter(x => !x.meet12);
  return { L, bez, od, who };
}
VIEWS.spotk = () => {
  const Q = spData(), S = ST.sp, B = Q.bez;
  const ago = d => d ? Math.round((TODAY - d) / 864e5) : null;
  const cols = [{ h: 'Firma', v: x => x.c.name, f: x => `${compLink(x.c.name, x.c.guid)}${x.c.ac ? ' <span class="chip c-mut">Academic</span>' : ''}<div style="font-size:11px;color:#9CA3AF">${esc(x.c.status1)}</div>` },
    { h: 'Opiekun', v: x => x.c.opiekun, f: x => esc(x.c.opiekun || '—') },
    { h: 'Ostatnie spotkanie', v: x => x.lastMd || new Date(0), f: x => x.lastM ? `${dpl(x.lastMd)} <span class="chip c-red">${ago(x.lastMd)} dni temu</span><div style="font-size:11px;color:#9CA3AF">${esc(x.lastM.rodzaj)}${x.lastM.osoba ? ' · ' + esc(x.lastM.osoba) : ''}</div>` : '<span class="chip c-red">nigdy (brak w CRM)</span>' },
    { h: 'Spotkania łącznie', key: 'nMeets', num: 1, f: x => x.nMeets ? `${x.nMeets} <span style="color:#9CA3AF">(${x.nF2F} F2F · ${x.nMeets - x.nF2F} online)</span>` : '0' },
    { h: 'Aktywny TECS', key: 'tecsVal', num: 1, f: x => x.tecs.length ? `${eur(x.tecsVal)}<div style="font-size:11px;color:#9CA3AF">do ${dpl(x.tecsEnd)}</div>` : '<span style="color:#9CA3AF">—</span>' },
    { h: 'Otwarte szanse', key: 'oppsVal', num: 1, f: x => x.opps.length ? `${eur(x.oppsVal)} <span style="color:#9CA3AF">(${x.opps.length})</span>` : '<span style="color:#9CA3AF">—</span>' },
    { h: '', f: x => ME.staff ? `<button class="btn" data-spn="${x.c.id}" title="Dodaj notatkę / umów spotkanie">${svg('note')}</button>` : '' }];
  if (Q.who) cols.splice(1, 1);
  const segs = [['klient', 'Status: Klient'], ['tecs', 'Z aktywnym TECS'], ['all', 'Wszyscy poza Prospect']];
  return `<div class="card panel">${ptitle('users', '#DC2626', '#FEE2E2', 'Klienci bez spotkania online ani F2F', `Liczą się wyłącznie notatki „spotkanie online” i „spotkanie F2F” (telefony, maile, wysyłki itp. pomijane) · brak spotkania od ${dpl(Q.od)}`,
    `<div class="seg" id="spSeg">${segs.map(([k, l]) => `<button data-v="${k}" class="${S.seg === k ? 'on' : ''}">${l}</button>`).join('')}</div>
     <select id="spDni">${[[180, '6 mies.'], [365, '12 mies.'], [730, '24 mies.']].map(([d, l]) => `<option value="${d}" ${S.dni === d ? 'selected' : ''}>${l}</option>`).join('')}</select>
     <select id="spWho"><option value="me" ${S.who === 'me' ? 'selected' : ''}>Moi klienci (${esc(ME.name)})</option><option value="" ${S.who === '' ? 'selected' : ''}>Wszyscy opiekunowie</option>${[...new Set(DB.comp.map(c => c.opiekun).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'pl')).filter(o => o !== ME.name).map(o => `<option ${S.who === o ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`)}
    <div class="kpis">${kpi('Klienci w zestawieniu', Q.L.length, `${Q.L.length - B.length} ze spotkaniem w okresie`, 'FIRMY', '#0D9488', '#E0F5F3')}
      ${kpi('Bez spotkania', B.length, Q.L.length ? Math.round(B.length / Q.L.length * 100) + '% klientów' : '', 'alert', '#DC2626', '#FEE2E2')}
      ${kpi('…w tym z aktywnym TECS', B.filter(x => x.tecs.length).length, eur(sum(B, x => x.tecsVal)) + ' TECS bez spotkania', 'TECSY', '#7C3AED', '#F1EAFE')}
      ${kpi('Nigdy nie było spotkania', B.filter(x => !x.nMeets).length, 'brak notatki spotkanie online / F2F w CRM', 'x', '#D97706', '#FEF3E2')}</div>
    ${table('spt', cols, B, { k: Q.who ? 3 : 4, asc: false }, { empty: 'Każdy klient miał spotkanie w tym okresie 🎉' })}</div>`;
};
AFTER.spotk = () => {
  document.querySelectorAll('#spSeg button').forEach(b => b.onclick = () => { ST.sp.seg = b.dataset.v; render(); });
  $('#spDni').onchange = e => { ST.sp.dni = +e.target.value; render(); };
  $('#spWho').onchange = e => { ST.sp.who = e.target.value; render(); };
  document.querySelectorAll('[data-spn]').forEach(b => b.onclick = () => { const c = DB.comp.find(x => x.id === +b.dataset.spn);
    noteModal({ firma: c.name, guid: c.guid, osoba: '', tel: '', contact: '', etap: '', etapP: '', opps: '', oppName: '', raw: { IdentyfikatorCompanies: c.guidRaw || '', Opps_ID: '' } }); });
};
/* --- modal notatki po follow-upie --- */
function noteModal(src) {
  const st = { rodzaj: '', dni: 0 };
  const rodz = DB.rodzaje;
  const m = $('#modal'); m.hidden = false;
  const draw = () => {
    const quick = [['Telefon', 'telefon'], ['Oferta', 'oferta'], ['Demo', 'demo']];
    m.innerHTML = `<div class="mbox"><div class="mh"><div class="ic" style="width:36px;height:36px;border-radius:8px;background:#FEF3E2;display:flex;align-items:center;justify-content:center">${svg('note', '#D97706')}</div>
      <div style="flex:1"><div style="font-weight:600;font-size:15px">${esc(src.firma || 'Notatka')}</div><div class="sub" style="font-size:11.5px;color:#6B7280">${src.id ? 'Nowa notatka po follow-upie' : (src.oppName ? 'Nowa notatka do szansy ' + esc(src.oppName) : 'Nowa notatka')} · ${esc(src.osoba || '—')}${src.tel ? ' · ' + esc(src.tel) : ''}</div></div>
      <button class="btn gray" id="mClose" title="Zamknij">${svg('x')}</button></div>
      <div class="fgrid"><label>Data wydarzenia<input type="date" id="mData" value="${ds(TODAY)}"></label><label>Osoba<input type="text" id="mOsoba" value="${esc(src.osoba)}"></label>
      <label>Rodzaj kontaktu *<select id="mRodz"><option value="">Wybierz rodzaj</option>${rodz.map(r => `<option value="${r.id}" ${lc(r.name).startsWith(st.rodzaj) && st.rodzaj ? 'selected' : ''}>${esc(r.name)}</option>`).join('')}</select></label></div>
      <label class="fl" style="margin-top:12px">Notatka *<textarea id="mTxt" rows="6" placeholder="Co ustaliliście?"></textarea></label>
      <div class="sec">KOLEJNY FOLLOW-UP</div>
      <div class="fgrid" style="margin-top:6px"><label style="flex-direction:row;align-items:center;gap:8px"><input type="checkbox" id="mFu" ${st.dni ? 'checked' : ''}> ustaw follow-up</label>
      <label>Data follow-upu<input type="date" id="mFuD" value="${st.dni ? ds(addDays(TODAY, st.dni)) : ''}"></label>
      <label style="flex-direction:row;align-items:center;gap:8px;${src.id ? '' : 'visibility:hidden'}"><input type="checkbox" id="mClosePrev" ${src.id ? 'checked' : ''}> poprzedni follow-up: oznacz jako zrobiony</label></div>
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
const oppLink = o => `<a class="lnk" href="#oppsy/opp/${o.id}" title="Otwórz szczegóły szansy">${esc(o.name || '#' + o.id)}</a>`;
const medChip = o => `<span class="chip ${medColor(o.meddic)[2]}" title="MEDDIC: ${MEDDIC.map((m, i) => m[1] + ' ' + o.med[i] + '/' + m[2]).join(', ')}">${o.meddic}%</span>`;
const sCols = () => [{ h: 'Szansa', key: 'name', f: o => `${oppLink(o)}<div style="font-size:11px;color:#9CA3AF">#${o.id}</div>` }, { h: 'Firma', key: 'comp', f: o => `<span class="lnk" style="font-weight:500" data-gocomp="${esc(o.comp)}" data-mod="oppsy" data-tab="s-otwarte">${esc(o.comp)}</span>` }, { h: 'Opiekun', key: 'owner' },
  { h: 'Etap', v: o => STAGES.indexOf(o.stage), f: o => `<span class="chip ${o.grp === 'won' ? 'c-ok' : o.grp === 'lost' ? 'c-red' : 'c-br'}">${esc(o.stage || '—')}</span>` }, { h: 'Rodzaj', key: 'rodz' },
  { h: 'Spodziewane zamknięcie', v: o => o.close, f: o => o.close ? `<span style="white-space:nowrap">${dpl(o.close)} <span style="color:#9CA3AF">${o.fq ? o.fq.label : ''}</span></span>` : '<span class="chip c-amb">brak</span>' },
  { h: 'Wartość EUR', key: 'eur', num: 1, f: o => fmt(o.eur) }, { h: 'Suma ofert', key: 'offerSum', num: 1, f: o => o.nOffers ? `${fmt(o.offerSum)} <span style="color:#9CA3AF">(${o.nOffers})</span>` : '<span style="color:#9CA3AF">—</span>' },
  { h: 'Notatki', key: 'nNotes', num: 1, f: o => o.nNotes ? `${o.nNotes} <span style="color:#9CA3AF">${dpl(o.lastNote)}</span>` : '0' }, { h: 'MEDDIC', key: 'meddic', num: 1, f: medChip }];
VIEWS['s-przeglad'] = () => {
  const S = gf(DB.sales), op = S.filter(o => o.grp === 'open'), won = S.filter(o => o.grp === 'won'), lost = S.filter(o => o.grp === 'lost'), cl = won.length + lost.length;
  const fr = STAGES.map(s => { const x = op.filter(o => o.stage === s); return { s, n: x.length, v: sum(x, o => o.value), noOff: x.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, noN: x.filter(o => !o.nNotes).length, med: x.length ? Math.round(sum(x, o => o.meddic) / x.length) : null }; });
  const mx = Math.max(1, ...fr.map(r => r.v));
  const owners = [...new Set(S.map(o => o.owner))].map(w => { const x = S.filter(o => o.owner === w), o1 = x.filter(o => o.grp === 'open'), w1 = x.filter(o => o.grp === 'won'), l1 = x.filter(o => o.grp === 'lost');
    return { w, n: o1.length, v: sum(o1, o => o.value), wv: sum(w1, o => o.value), wn: w1.length, ln: l1.length, rate: (w1.length + l1.length) ? w1.length / (w1.length + l1.length) : null, late: o1.filter(o => o.days != null && o.days < 0).length, noOff: o1.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, idle: o1.filter(o => o.flags.some(f => /notat/.test(f[0]))).length }; });
  const al = (n, c = 'c-red') => n ? `<span class="chip ${c}">${n}</span>` : '0';
  const rz = [...new Set(op.map(o => o.rodz || '—'))].map(r => { const x = op.filter(o => (o.rodz || '—') === r); return { r, n: x.length, v: sum(x, o => o.value), m: sum(x, o => o.marza) }; });
  return `<div class="card panel"><div class="kpis">${kpi('W toku (Aktywny)', eur(sum(op, o => o.value)), op.length + ' szans', 'clock')}${kpi('Zamknięty wygrany', eur(sum(won, o => o.value)), won.length + ' szans', 'check', '#16A34A', '#E6F6EE')}${kpi('Zamknięty przegrany', eur(sum(lost, o => o.value)), lost.length + ' szans', 'x', '#DC2626', '#FEE2E2')}
    ${kpi('Win rate', cl ? Math.round(won.length / cl * 100) + '%' : '—', cl ? 'wartościowo ' + Math.round(sum(won, o => o.value) / Math.max(1, sum(won, o => o.value) + sum(lost, o => o.value)) * 100) + '%' : '', 'target', '#0D9488', '#E0F5F3')}${kpi('Po terminie zamknięcia', op.filter(o => o.days != null && o.days < 0).length, op.filter(o => !o.close).length + ' bez daty', 'alert', '#D97706', '#FEF3E2')}${kpi('Etap oferty bez oferty', op.filter(o => o.flags.some(f => f[0] === 'Bez oferty')).length, 'Oferta / Rewizja / Decyzja', 'box', '#DC2626', '#FEE2E2')}${kpi('Średni MEDDIC w toku', op.length ? Math.round(sum(op, o => o.meddic) / op.length) + '%' : '—', op.filter(o => o.meddic < 40).length + ' szans poniżej 40%', 'target', '#7C3AED', '#F1EAFE')}</div>
    <div class="grid2"><div>${ptitle('OPPSY', '#5B6CDB', '#EEF0FC', 'Lejek w toku wg etapu (Etap1)', 'Status1 = Aktywny. Kliknij etap, żeby zobaczyć listę.')}
    <table><thead><tr><th>Etap</th><th class="num">Szanse</th><th>Wartość EUR</th><th class="num">Bez oferty</th><th class="num">Bez notatek</th><th class="num">Śr. MEDDIC</th></tr></thead><tbody>${fr.map(r => `<tr data-stage="${esc(r.s)}" style="cursor:pointer"><td>${esc(r.s)}</td><td class="num">${r.n}</td><td class="num" style="text-align:left"><span class="bar" style="width:${Math.round(r.v / mx * 200)}px"></span> ${fmt(r.v)}</td><td class="num">${al(r.noOff)}</td><td class="num">${al(r.noN, 'c-amb')}</td><td class="num">${r.med == null ? '—' : `<span class="chip ${medColor(r.med)[2]}">${r.med}%</span>`}</td></tr>`).join('')}</tbody></table></div>
    <div>${ptitle('euro', '#16A34A', '#E6F6EE', 'Rodzaj biznesu — w toku', 'z marżą MESco wg Twoich stawek')}${table('srz', [{ h: 'Rodzaj', key: 'r' }, { h: 'Szanse', key: 'n', num: 1 }, { h: 'Wartość EUR', key: 'v', num: 1, f: r => fmt(r.v) }, { h: 'Marża MESco', key: 'm', num: 1, f: r => fmt(r.m) }], rz, { k: 2, asc: false })}</div></div></div>
    <div class="card panel">${ptitle('bars', '#5B6CDB', '#EEF0FC', 'Wartość EUR wg kwartału FY Ansys', 'W toku wg Spodziewanego zamknięcia, zamknięte wg daty zamknięcia (ClosedDateWin). Kliknij słupek, żeby filtrować.')}${qChart(gf(DB.sales, true))}</div>
    <div class="card panel">${ptitle('users', '#5B6CDB', '#EEF0FC', 'Opiekunowie', '')}${table('sown', [{ h: 'Opiekun', key: 'w' }, { h: 'W toku', key: 'n', num: 1 }, { h: 'Wartość w toku', key: 'v', num: 1, f: r => eur(r.v) }, { h: 'Wygrane', key: 'wv', num: 1, f: r => `${eur(r.wv)} <span style="color:#9CA3AF">(${r.wn})</span>` }, { h: 'Win rate', key: 'rate', num: 1, f: r => r.rate == null ? '—' : Math.round(r.rate * 100) + '%' }, { h: 'Po terminie', key: 'late', num: 1, f: r => al(r.late) }, { h: 'Bez oferty', key: 'noOff', num: 1, f: r => al(r.noOff) }, { h: 'Bez notatek / >60 dni', key: 'idle', num: 1, f: r => al(r.idle, 'c-amb') }], owners, { k: 2, asc: false })}</div>`;
};
AFTER['s-przeglad'] = () => document.querySelectorAll('tr[data-stage]').forEach(tr => tr.onclick = () => go('oppsy', 's-otwarte', tr.dataset.stage));
VIEWS['s-otwarte'] = () => { let op = gf(DB.sales).filter(o => o.grp === 'open'); const all = op; if (ST.sstage) op = op.filter(o => o.stage === ST.sstage);
  const si = STAGES.indexOf(ST.sstage);
  return `${ST.sstage ? `<div class="card backbar">${backBtn('stgBack', 'Wróć do przeglądu')}<div><div class="bt">Etap: ${esc(ST.sstage)}</div><div class="bs">${op.length} szans w toku · ${eur(sum(op, o => o.value))}</div></div>
    <div class="steps mini">${STAGES.map((s, i) => `<span class="${i === si ? 'cur' : i < si ? 'done' : ''}">${i < si ? '✓' : i + 1}. ${esc(STAGE_SHORT[i])}</span>`).join('')}</div></div>` : ''}<div class="card panel">${ptitle('clock', '#5B6CDB', '#EEF0FC', 'Szanse sprzedaży w toku', 'Status1 = Aktywny', `<div class="seg" id="sStg" style="flex-wrap:wrap"><button class="${!ST.sstage ? 'on' : ''}" data-v="">Wszystkie</button>${STAGES.map(s => `<button class="${ST.sstage === s ? 'on' : ''}" data-v="${esc(s)}">${esc(s)} (${all.filter(o => o.stage === s).length})</button>`).join('')}</div>`)}
    <div class="kpis">${kpi('Szanse', op.length, '', 'list')}${kpi('Wartość EUR', eur(sum(op, o => o.value)), '', 'euro', '#16A34A', '#E6F6EE')}${kpi('Marża MESco', eur(sum(op, o => o.marza)), '', 'euro')}${kpi('Suma ofert', fmt(sum(op, o => o.offerSum)), op.filter(o => o.nOffers).length + ' z ofertą', 'box', '#D97706', '#FEF3E2')}</div>
    ${table('sopen', [...sCols(), { h: 'Do zamknięcia', key: 'days', num: 1, f: o => o.days == null ? '—' : o.days < 0 ? `<span class="chip c-red">${-o.days} dni po</span>` : o.days + ' dni' }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], op, { k: 5, asc: true })}</div>`; };
AFTER['s-otwarte'] = () => { document.querySelectorAll('#sStg button').forEach(b => b.onclick = () => { history.replaceState(null, '', '#oppsy/s-otwarte' + (b.dataset.v ? '/' + encodeURIComponent(b.dataset.v) : '')); ST.sstage = b.dataset.v; render(); });
  const bk = $('#stgBack'); if (bk) bk.onclick = () => goBack(() => go('oppsy', 's-przeglad')); };
VIEWS['s-zamkniete'] = () => { const S = gf(DB.sales), won = S.filter(o => o.grp === 'won'), lost = S.filter(o => o.grp === 'lost');
  const cols = [...sCols(), { h: 'Data zamknięcia', key: 'closedWin', f: o => dpl(o.closedWin) }];
  return `<div class="card panel"><div class="kpis">${kpi('Zamknięty wygrany', eur(sum(won, o => o.value)), won.length + ' szans', 'check', '#16A34A', '#E6F6EE')}${kpi('Marża MESco z wygranych', eur(sum(won, o => o.marza)), '', 'euro')}${kpi('Premia Q + R', eur(sum(won, o => o.premQ + o.premR)), '', 'target', '#7C3AED', '#F1EAFE')}${kpi('Zamknięty przegrany', eur(sum(lost, o => o.value)), lost.length + ' szans', 'x', '#DC2626', '#FEE2E2')}</div>
    ${ptitle('check', '#16A34A', '#E6F6EE', 'Wygrane — Status1 = Zamknięty wygrany', '')}${table('swon', cols, won, { k: 10, asc: false })}</div><div class="card panel">${ptitle('x', '#DC2626', '#FEE2E2', 'Przegrane — Status1 = Zamknięty przegrany', '')}${table('slost', cols, lost, { k: 10, asc: false })}</div>`; };
VIEWS['s-alerty'] = () => { const op = gf(DB.sales).filter(o => o.grp === 'open'), late = op.filter(o => o.days != null && o.days < 0), noOff = op.filter(o => o.flags.some(f => f[0] === 'Bez oferty')), idle = op.filter(o => o.flags.some(f => /notat/.test(f[0]))), noD = op.filter(o => !o.close);
  const cols = [...sCols(), { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }], colsI = [...sCols(), { h: 'Dni ciszy', key: 'idle', num: 1 }, { h: 'Alerty', v: o => o.flags.length, f: o => chips(o.flags) }];
  return `<div class="card panel"><div class="kpis">${kpi('Po terminie zamknięcia', late.length, eur(sum(late, o => o.value)), 'alert', '#DC2626', '#FEE2E2')}${kpi('Etap oferty bez oferty', noOff.length, eur(sum(noOff, o => o.value)), 'box', '#DC2626', '#FEE2E2')}${kpi('Bez notatek lub >60 dni ciszy', idle.length, '', 'note', '#D97706', '#FEF3E2')}${kpi('Bez daty zamknięcia', noD.length, '', 'KALENDARZ', '#D97706', '#FEF3E2')}</div>
    ${ptitle('box', '#DC2626', '#FEE2E2', 'Na etapie Oferta / Rewizja / Decyzja bez oferty', 'Powiązanie: Oferty.IdentyfikatorOpps = OPPSGUID')}${table('sa1', cols, noOff, { k: 6, asc: false })}</div>
    <div class="card panel">${ptitle('alert', '#DC2626', '#FEE2E2', 'Po spodziewanej dacie zamknięcia', '')}${table('sa2', cols, late, { k: 5, asc: true })}</div>
    <div class="card panel">${ptitle('note', '#D97706', '#FEF3E2', 'Bez notatek lub ostatnia notatka starsza niż 60 dni', 'Powiązanie: Notes.Opps_ID = OPPSGUID')}${table('sa3', colsI, idle, { k: 10, asc: false })}</div>`; };

/* ================= SZCZEGÓŁY SZANSY (jak DetailScreen w Power Apps) ================= */
const backBtn = (id, t = 'Wróć') => `<button class="backb" id="${id}" title="${t}">${svg('chevL', '#374151')}</button>`;
const STF = [
  [['FTDate','Data spotkania','date'], ['TerminDR','Termin DR z SF','date']],
  [['DISCH','Champion','text'], ['DISB','Budżet (EUR)','num'], ['DISP','Pain','area',2]],
  [['OFP','Czy oferta została wysłana?','bool'], ['OFSF','Czy zrobiłeś quote w SF?','bool'], ['QuoteNumber','Numer quota','text'], ['OFDate','Data spotkania','date'], ['QuoteNumberLink','Link do quota','url',2]],
  [['TMFU','Data follow-up','date'], ['TMEB','Economic Buyer – Eb','text'], ['OFDP','Proces decyzji – Dp','area'], ['TMDC','Kryteria – Dc','area'], ['TMM','Metryki – M','area',2]],
  [['DBDE','Czy będzie Demo/Benchmark?','bool'], ['DBDate','Data demo / benchmarku','date']],
  [['DSM','Spotkanie podsumowujące','date'], ['DEB','Potwierdzenie Eb','area',2]],
  [['RTL','Czy potrzebne są szkolenia?','area',2], ['RWH','Czy potrzebny jest sprzęt?','area',2], ['RCR','Czy będzie zmiana konfiguracji?','area',2]],
  [['WRDL','Data deadline','date'], ['WROF','Czy jest Order Form od klienta?','bool'], ['WRO','Zamówienie od klienta','area',2]],
];
const oppNoteSrc = x => ({ firma: x.comp, guid: x.compGuid, osoba: x.raw.DISCH || '', tel: '', contact: '', etap: x.stage, etapP: '', opps: x.guid, oppName: x.name, raw: { IdentyfikatorCompanies: x.raw.IdentyfikatorCompanies || '', Opps_ID: x.raw.OPPSGUID || '' } });
const urlOf = v => !v ? '' : typeof v === 'string' ? v : (v.Url || v.url || '');
const curSale = () => DB.sales.find(o => o.id === ST.oppId);
async function saveSale(x, fields, msg) {
  await patchItem('Opps', x.id, fields);
  Object.assign(x.raw, fields);
  const nx = DB.mkSale(x.raw); nx.nNotes = x.nNotes; nx.lastNote = x.lastNote;
  DB.sales[DB.sales.indexOf(x)] = nx;
  if (msg) toast(msg, 'ok');
  return nx;
}
function openModal(html, w = 700) { const m = $('#modal'); m.hidden = false; m.innerHTML = `<div class="mbox" style="width:${w}px">${html}</div>`;
  const close = () => { m.hidden = true; m.innerHTML = ''; }; m.querySelectorAll('[data-close]').forEach(b => b.onclick = close); return { m, close }; }
const mHead = (icon, c, bg, t, sub) => `<div class="mh"><div style="width:36px;height:36px;border-radius:8px;background:${bg};display:flex;align-items:center;justify-content:center">${svg(icon, c)}</div>
  <div style="flex:1"><div style="font-weight:600;font-size:15px">${t}</div><div style="font-size:11.5px;color:#6B7280">${sub}</div></div><button class="btn gray" data-close title="Zamknij">${svg('x')}</button></div>`;

VIEWS.opp = () => {
  const x = curSale();
  if (!x) return `<div class="card panel"><div class="empty">Nie znaleziono szansy #${esc(ST.oppId)} — mogła zostać usunięta.<br><br><button class="btn" id="dsBack0">Wróć do listy</button></div></div>`;
  const r = x.raw, si = STAGES.indexOf(x.stage);
  if (!ST.ds || ST.ds.id !== x.id) ST.ds = { id: x.id, tab: Math.max(0, si) };
  const tab = ST.ds.tab, mc = medColor(x.meddic);
  const notes = DB.notes.filter(n => n.opps === x.guid).sort((a, b) => (b.data || 0) - (a.data || 0));
  const stCol = { 'Aktywny': ['#5B6CDB', '#EEF0FC'], 'Zamknięty wygrany': ['#16A34A', '#E6F6EE'], 'Zamknięty przegrany': ['#DC2626', '#FEE2E2'] }[x.status1] || ['#6B7280', '#F3F4F6'];
  const sub = [x.rodz || '—', x.typ || '—', r.QuoteNumber ? 'Quote: ' + r.QuoteNumber : '', r.DISCH ? 'Champion: ' + r.DISCH : ''].filter(Boolean).map(esc).join('  ·  ');
  const sbtn = (st, l, c, bg, act) => `<button class="sbtn" data-st="${st}" style="--c:${c};--bg:${bg}" ${x.status1 === st ? 'data-on' : ''}>${x.status1 === st ? '✓ ' : ''}${l}</button>`;
  const field = ([k, l, t, span]) => { const v = r[k];
    const inp = t === 'date' ? `<input type="date" data-f="${k}" data-t="date" value="${ds(toD(v))}">`
      : t === 'bool' ? `<button class="sw ${v ? 'on' : ''}" data-f="${k}" data-t="bool" type="button"><i></i><span>${v ? 'Tak' : 'Nie'}</span></button>`
      : t === 'area' ? `<textarea data-f="${k}" data-t="text" rows="${span ? 5 : 4}">${esc(v || '')}</textarea>`
      : t === 'url' ? (urlOf(v) ? `<a class="lnk" href="${esc(urlOf(v))}" target="_blank" rel="noopener" style="line-height:34px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(urlOf(v))}</a>` : `<span style="line-height:34px;color:#9CA3AF">brak — link dodasz w SharePoint („Edytuj w SharePoint”)</span>`)
      : `<input type="text" data-f="${k}" data-t="${t}" value="${esc(v ?? '')}">`;
    return `<label class="${span ? 'span2' : ''}">${l}${inp}</label>`; };
  return `<div class="det">
  <div class="card dhead">${backBtn('dsBack')}
    <div class="dt"><div class="t" title="${esc(x.comp + ' — ' + x.name)}">${compLink(x.comp, x.compGuid)}  —  ${esc(x.name)}</div><div class="s">${sub}</div>
      <div class="pills"><span class="pill" style="background:#F3F4F6;color:#374151">Etap: ${esc(x.stage || '—')}</span><span class="pill" style="background:${stCol[1]};color:${stCol[0]}">${esc(x.status1)}</span><span class="pill" style="background:#E6F6EE;color:#16A34A">${eur(x.eur)}</span>${x.flags.map(f => `<span class="pill chip ${f[1]}">${esc(f[0])}</span>`).join('')}</div></div>
    <div class="dr"><div class="sbtns">${sbtn('Aktywny', 'Aktywny', '#5B6CDB', '#EEF0FC')}${sbtn('Zamknięty wygrany', 'Wygrany', '#16A34A', '#E6F6EE')}${sbtn('Zamknięty przegrany', 'Przegrany', '#DC2626', '#FEE2E2')}</div>
      <div class="own"><div class="av">${initials(x.owner)}</div><div><div style="font-size:10.5px;color:#9CA3AF">Opiekun</div><div style="font-weight:600;font-size:12.5px">${esc(x.owner || '—')}</div></div><span style="flex:1"></span>
      <button class="btn" id="dsFirma" title="Wszystkie szanse tej firmy">Firma</button>${urlOf(r.LinkSalesForce) ? `<a class="btn" href="${esc(urlOf(r.LinkSalesForce))}" target="_blank" rel="noopener">SalesForce</a>` : ''}</div></div></div>
  <div class="card steps">${STAGES.map((s, i) => `<button data-stg="${i}" class="${i === si ? 'cur' : i < si ? 'done' : ''}" title="${esc(s)}">${i < si ? '✓' : i + 1 + '.'}  ${esc(STAGE_SHORT[i])}</button>`).join('')}</div>
  <div class="dgrid"><div>
    <div class="dk"><div class="card kc"><div class="kic" style="background:#E6F6EE">${svg('OPPSY', '#16A34A')}</div><div><div class="lab">Wartość oppsa</div><div class="val">${eur(x.eur)}</div>${x.nOffers && Math.round(x.offerSum) !== Math.round(x.eur) ? `<div class="s" style="color:#D97706">suma pozycji ofert: ${eur(x.offerSum)}</div>` : ''}</div></div>
      <div class="card kc"><div class="kic" style="background:#F1EAFE">${svg('target', '#7C3AED')}</div><div style="flex:1"><div class="lab">MEDDIC score</div><div class="val">${x.meddic} %</div><div class="track" style="height:6px"><div style="width:${Math.min(100, x.meddic)}%;background:${mc[0]}"></div></div></div></div></div>
    <div class="card dc"><div class="sec" style="margin:0 0 8px">MEDDIC — kliknij, aby zmienić ocenę (0 → częściowo → pełna)</div>
      <div class="med">${MEDDIC.map((m, i) => { const v = x.med[i], c = v >= m[2] ? 'full' : v > 0 ? 'half' : 'zero'; return `<button data-med="${i}" class="${c}"><b>${m[1]}</b>${v} / ${m[2]}</button>`; }).join('')}</div></div>
    <div class="card dc">${ptitle('note', '#5B6CDB', '#EEF0FC', `Notatki (${notes.length})`, '', `<button class="btn pri" id="dsNote">+ Dodaj notatkę</button>`)}
      <div class="dnotes">${notes.length ? notes.map(n => `<div class="dn" data-n="${n.id}"><div><div class="h">${dpl(n.data)}  ·  ${esc(n.osoba || '—')}  ·  ${esc(n.rodzaj)}</div><div class="x">${esc(String(n.text).replace(/\n/g, ' '))}</div></div>
        ${n.fu ? `<span class="chip ${n.fuDone ? 'c-mut' : n.fu < TODAY ? 'c-red' : 'c-amb'}">FollUp ${dpl(n.fu).slice(0, 5)}${n.fuDone ? ' ✓' : ''}</span>` : ''}${svg('chevR', '#9CA3AF')}</div>`).join('') : '<div class="empty" style="padding:20px">Brak notatek</div>'}</div></div>
  </div><div class="card dc etc">
    ${ptitle('list', '#0D9488', '#E0F5F3', 'Dane etapu: ' + esc(STAGES[tab]))}
    <div class="etabs">${STAGES.map((s, i) => `<button data-et="${i}" class="${i === tab ? 'on' : ''}" title="${esc(s)}">${i === si ? '● ' : ''}${esc(['Kontakt','Discovery','Oferta','Spotk. tech.','Demo','Decyzja','Rewizja','Wdrożenie'][i])}</button>`).join('')}</div>
    <div class="fgrid ef">${STF[tab].map(field).join('')}</div>
    <div class="efoot"><span class="desc" style="margin:0">Zapisuje tylko pola widocznej zakładki. Kropka ● oznacza bieżący etap.</span><button class="btn pri" id="dsSaveEt" style="height:34px;padding:0 20px">Zapisz etap</button></div>
  </div></div>
  <div class="card dact"><button class="btn" id="dsOf">${svg('box')} Produkty i ceny (${x.nOffers})</button><button class="btn" id="dsRecalc">${svg('euro')} Przelicz wartość</button>
    <a class="btn" href="${SPU}Szanse%20Sprzeday/EditForm.aspx?ID=${x.id}" target="_blank" rel="noopener">${svg('ext')} Edytuj w SharePoint (załączniki, opiekun)</a><button class="btn pri" id="dsNote2">${svg('note')} Dodaj notatkę</button></div></div>`;
};
AFTER.opp = () => {
  const back = () => goBack(() => go('oppsy', 's-otwarte'));
  if ($('#dsBack0')) { $('#dsBack0').onclick = back; return; }
  let x = curSale(); const r = x.raw;
  $('#dsBack').onclick = back;
  const run = async (btn, fn) => { if (btn) btn.disabled = true; try { x = await fn(); render(); } catch (e) { if (btn) btn.disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); } };
  $('#dsFirma').onclick = () => { ST.g.comp = x.comp; save(); go('oppsy', 's-otwarte'); };
  // etapy
  document.querySelectorAll('[data-stg]').forEach(b => b.onclick = () => { const i = +b.dataset.stg; if (STAGES[i] === x.stage) { ST.ds.tab = i; return render(); }
    run(b, async () => { const nx = await saveSale(x, { Etap1: STAGES[i] }, 'Etap: ' + STAGES[i]); ST.ds.tab = i; return nx; }); });
  document.querySelectorAll('[data-et]').forEach(b => b.onclick = () => { ST.ds.tab = +b.dataset.et; render(); });
  document.querySelectorAll('.ef .sw').forEach(b => b.onclick = () => { b.classList.toggle('on'); b.querySelector('span').textContent = b.classList.contains('on') ? 'Tak' : 'Nie'; });
  $('#dsSaveEt').onclick = e => { const f = {};
    document.querySelectorAll('.ef [data-f]').forEach(el => { const k = el.dataset.f, t = el.dataset.t;
      f[k] = t === 'bool' ? el.classList.contains('on') : t === 'date' ? (el.value ? isoNoon(toD(el.value + 'T12:00:00')) : null) : t === 'num' ? num(el.value) : el.value; });
    run(e.currentTarget, () => saveSale(x, f, 'Zapisano dane etapu: ' + STAGES[ST.ds.tab])); };
  // MEDDIC
  document.querySelectorAll('[data-med]').forEach(b => b.onclick = () => { const m = MEDDIC[+b.dataset.med], v = String(r[m[0]] || '0'), nv = v === '0' ? '8' : v === '8' ? String(m[2]) : '0';
    run(b, () => saveSale(x, { [m[0]]: nv }, `${m[1]}: ${nv} / ${m[2]}`)); });
  // status
  document.querySelectorAll('[data-st]').forEach(b => b.onclick = () => {
    const st = b.dataset.st;
    if (st === 'Aktywny') return run(b, () => saveSale(x, { Status1: 'Aktywny' }, 'Status: Aktywny'));
    if (st === 'Zamknięty wygrany') { const { m, close } = openModal(`${mHead('check', '#16A34A', '#E6F6EE', 'Gratulacje, wygrana! 🥳', 'Zamykasz oppsa jako wygranego — podaj datę zamknięcia')}
      <div class="fgrid"><label>Data zamknięcia *<input type="date" id="wD" value="${ds(x.closedWin || TODAY)}"></label></div>
      <div class="desc" style="margin-top:12px">Nowy TECS po wygranej założysz w Power Apps (formularz TECS_newOpps).</div>
      <div class="mfoot"><span style="flex:1"></span><button class="btn gray" data-close style="height:36px">Anuluj</button><button class="btn pri" id="wS" style="height:36px;background:#16A34A;border-color:#16A34A">Zapisz</button></div>`, 600);
      $('#wS').onclick = () => { const d = $('#wD').value; if (!d) return toast('Podaj datę zamknięcia', 'err'); $('#wS').disabled = true;
        saveSale(x, { Status1: 'Zamknięty wygrany', ClosedDate: isoNoon(toD(d + 'T12:00:00')) }, 'Opps zamknięty jako wygrany').then(() => { close(); render(); }).catch(e => { $('#wS').disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); }); };
      return; }
    const p = { DR_Odrzucony: !!r.DR_Odrzucony, Konkurencja: !!r.Konkurencja, BrakDofinansowania: !!r.BrakDofinansowania };
    const { close } = openModal(`${mHead('x', '#DC2626', '#FEE2E2', 'Zamknięcie jako przegrany', 'Podaj powody, dla których zamykasz oppsa')}
      <div class="sec">POWODY</div><div class="lostp">${[['DR_Odrzucony','Odrzucony DR'],['Konkurencja','Konkurencja'],['BrakDofinansowania','Brak dofinansowania']].map(([k, l]) => `<button data-p="${k}" class="${p[k] ? 'on' : ''}">${p[k] ? '✓  ' : ''}${l}</button>`).join('')}</div>
      <label class="fl" style="margin-top:12px">Notatka<textarea id="lN" rows="8" placeholder="Co zdecydowało? Czego się nauczyliśmy?">${esc(r.Zamkni_x0119_tyOppsNotes || '')}</textarea></label>
      <div class="mfoot"><span style="flex:1"></span><button class="btn gray" data-close style="height:36px">Anuluj</button><button class="btn pri" id="lS" style="height:36px;background:#DC2626;border-color:#DC2626">Zamknij oppsa</button></div>`, 800);
    document.querySelectorAll('[data-p]').forEach(b2 => b2.onclick = () => { const k = b2.dataset.p; p[k] = !p[k]; b2.classList.toggle('on', p[k]); b2.textContent = (p[k] ? '✓  ' : '') + b2.textContent.replace('✓  ', ''); });
    $('#lS').onclick = () => { $('#lS').disabled = true; saveSale(x, { Status1: 'Zamknięty przegrany', ...p, Zamkni_x0119_tyOppsNotes: $('#lN').value }, 'Opps zamknięty jako przegrany').then(() => { close(); render(); }).catch(e => { $('#lS').disabled = false; toast('Błąd zapisu: ' + e.message, 'err'); }); };
  });
  // notatki
  const addNote = () => noteModal(oppNoteSrc(x));
  $('#dsNote').onclick = $('#dsNote2').onclick = addNote;
  document.querySelectorAll('[data-n]').forEach(d => d.onclick = () => { const n = DB.notes.find(z => z.id === +d.dataset.n);
    openModal(`${mHead('note', '#5B6CDB', '#EEF0FC', `${dpl(n.data)}  ·  ${esc(n.osoba || '—')}  ·  ${esc(n.rodzaj)}`, `Etap: ${esc(n.etap || '—')}  ·  dodał(a): ${esc(n.creator || '—')}${n.fu ? '  ·  follow-up: ' + dpl(n.fu) : ''}`)}
      <div class="ntext">${esc(n.text || '')}</div><div class="mfoot">${n.tel ? `<a class="btn" href="tel:${esc(n.tel.replace(/\s/g, ''))}">${svg('phone')} Zadzwoń</a>` : ''}<a class="btn" href="${SPU}Notes/EditForm.aspx?ID=${n.id}" target="_blank" rel="noopener">${svg('ext')} Edytuj w SharePoint</a><span style="flex:1"></span><button class="btn pri" data-close style="height:36px;padding:0 20px">Zamknij</button></div>`, 900); });
  // oferty
  const recalc = btn => run(btn, () => saveSale(x, { Warto_x015b__x0107_wEUR: Math.round(sum(DB.offers[x.guid] || [], o => num(o.Suma)) * 100) / 100 }, 'Przeliczono wartość oppsa'));
  $('#dsRecalc').onclick = e => recalc(e.currentTarget);
  $('#dsOf').onclick = () => { const L = DB.offers[x.guid] || [], s0 = sum(L, o => num(o.Suma)), diff = Math.round(s0) !== Math.round(x.eur);
    const typ = o => o.PaidUP0 ? 'Paid-up' : o.LeaseDays ? `Lease ${num(o.Lease_Days) || ''} dni` : o.AECYN ? 'AEC' : o.ALHYN ? 'ALH' : '—';
    const f2 = v => (num(v) || 0).toLocaleString('pl-PL', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const { close } = openModal(`${mHead('box', '#0D9488', '#E0F5F3', 'Produkty i ceny', `${esc(x.comp)}  —  ${esc(x.name)}  ·  ${esc(x.typ)}`)}
      ${L.length ? `<div class="tbl-wrap" style="margin-top:12px"><table><thead><tr><th>Licencja</th><th class="num">Ilość</th><th>Rodzaj</th><th class="num">Paid-up</th><th class="num">TECS</th><th class="num">Lease</th><th class="num">AEC / ALH</th><th class="num">Rabat</th><th class="num">Suma</th></tr></thead><tbody>
      ${L.map(o => `<tr><td style="font-weight:600">${esc(o.ProductID || '—')}</td><td class="num">${num(o.Liczba_licencji) ?? ''}</td><td>${typ(o)}</td><td class="num">${fmt(num(o.Paidup) || 0)} €</td><td class="num">${fmt(num(o.TECS) || 0)} €</td><td class="num">${fmt(num(o.Lease) || 0)} €</td><td class="num">${fmt((num(o.AEC) || 0) + (num(o.ALH) || 0))} €</td><td class="num">${num(o.Rabat) || 0}%</td><td class="num" style="font-weight:600">${f2(o.Suma)} €</td></tr>`).join('')}</tbody></table></div>`
        : '<div class="empty">Brak produktów — dodasz je w Power Apps („Dodaj produkt”)</div>'}
      <div class="mfoot"><span class="desc" style="margin:0;max-width:420px">${diff ? 'Pomarańczowa wartość = suma pozycji różni się od wartości oppsa. Kliknij „Przelicz wartość”.' : ''}</span><span style="flex:1"></span>
      <div style="text-align:right"><div style="font-size:11px;color:#6B7280">Suma pozycji</div><div style="font-weight:600;font-size:15px">${f2(s0)} €</div></div>
      <div style="text-align:right;margin-left:18px"><div style="font-size:11px;color:#6B7280">Wartość oppsa w systemie</div><div style="font-weight:600;font-size:15px;color:${diff ? '#D97706' : '#16A34A'}">${f2(x.eur)} €</div></div>
      <button class="btn" id="ofRe" style="height:36px;margin-left:12px">${svg('euro')} Przelicz wartość</button></div>`, 1200);
    $('#ofRe').onclick = e => { close(); recalc(e.currentTarget); }; };
};


/* ================= MARKETING — webinary / wydarzenia ================= */
const pct = (a, b) => b ? Math.round(a / b * 100) + '%' : '—';
const attChip = (a, r) => { if (!r) return '—'; const p = a / r; return `<span class="chip ${p >= .75 ? 'c-ok' : p >= .4 ? 'c-amb' : 'c-red'}">${Math.round(p * 100)}%</span>`; };
const compOf = m => { const c = m.contact; if (c) { const co = DB.compBy[c.compGuid]; return { name: co ? co.name : (c.firma || m.firmaTxt || '—'), guid: c.compGuid, crm: !!co }; } return { name: m.firmaTxt || '—', guid: '', crm: false }; };
function mktPeople() {
  const by = {};
  DB.mkt.forEach(m => { const k = m.mail || 'id' + m.id; const p = by[k] = by[k] || { mail: m.mail, name: m.name, tel: m.tel, rows: [], contact: null, comp: null };
    p.rows.push(m); if (m.contact) p.contact = m.contact; if (!p.name && m.name) p.name = m.name; });
  return Object.values(by).map(p => { p.comp = compOf(p.rows.find(r => r.contact) || p.rows[0]); p.reg = p.rows.length; p.att = p.rows.filter(r => r.obecny === true).length; p.abs = p.rows.filter(r => r.obecny === false).length;
    p.rows.sort((a, b) => (b.web.date || 0) - (a.web.date || 0)); p.last = p.rows[0].web.date; p.lastAtt = (p.rows.find(r => r.obecny) || {}).web; p.crm = !!p.contact;
    const g = p.comp.guid; p.openOpps = g ? DB.sales.filter(o => o.compGuid === g && o.grp === 'open').length : 0; p.tecs = g ? DB.tecs.some(o => o.compGuid === g && o.grp !== 'demo' && o.start && o.end && o.start <= TODAY && o.end >= TODAY) : false; return p; });
}
const mFilt = P => { const F = ST.mk; return P.filter(p => (!F.crm || (F.crm === 'crm' ? p.crm : !p.crm)) && p.reg >= F.min && (!F.att || (F.att === 'never' ? p.att === 0 : p.att > 0))
  && (!F.web || p.rows.some(r => r.web.guid === F.web)) && (!F.q || lc(p.name + ' ' + p.mail + ' ' + p.comp.name).includes(lc(F.q)))); };
const evChips = p => p.rows.map(r => `<span class="chip ${r.obecny ? 'c-ok' : r.obecny === false ? 'c-red' : 'c-mut'}" title="${esc(r.web.title)} · ${dpl(r.web.date)} · ${r.obecny ? 'obecny' : r.obecny === false ? 'nieobecny' : 'brak informacji'}">${esc(r.web.title.length > 22 ? r.web.title.slice(0, 21) + '…' : r.web.title)}</span>`).join('');
const mkFilters = (withWeb = true) => { const F = ST.mk; const ws = DB.webinars.filter(w => DB.mkt.some(m => m.web === w)).sort((a, b) => (b.date || 0) - (a.date || 0));
  return `<div class="filters" style="border:0;padding:0;margin-bottom:10px"><input type="search" id="mkQ" placeholder="Szukaj osoby / firmy / e-mail…" value="${esc(F.q)}" style="width:260px">
  <div class="seg" id="mkCrm">${[['','Wszyscy'],['crm','W CRM'],['out','Spoza CRM']].map(([k, l]) => `<button data-v="${k}" class="${F.crm === k ? 'on' : ''}">${l}</button>`).join('')}</div>
  <div class="seg" id="mkAtt">${[['','Wszyscy'],['yes','Byli min. raz'],['never','Zapisani, nigdy nie byli']].map(([k, l]) => `<button data-v="${k}" class="${F.att === k ? 'on' : ''}">${l}</button>`).join('')}</div>
  <select id="mkMin">${[1, 2, 3, 5].map(n => `<option value="${n}" ${F.min === n ? 'selected' : ''}>${n === 1 ? 'Dowolna liczba zapisów' : 'min. ' + n + ' zapisy'}</option>`).join('')}</select>
  ${withWeb ? `<select id="mkWeb" style="max-width:300px"><option value="">Wszystkie wydarzenia</option>${ws.map(w => `<option value="${w.guid}" ${F.web === w.guid ? 'selected' : ''}>${dpl(w.date)} · ${esc(w.title)}</option>`).join('')}</select>` : ''}
  <button class="btn gray" id="mkReset">Wyczyść</button>${ME.staff ? `<button class="btn" id="mkMatch" style="margin-left:auto" title="Uzupełnia GUIDKontaktu w MarketingKontakty, gdy e-mail pasuje do kontaktu w Contacts">${svg('refresh')} Dopasuj kontakty po e-mailu</button>` : ''}</div>`; };
const bindMk = () => { const F = ST.mk;
  if ($('#mkQ')) $('#mkQ').onchange = e => { F.q = e.target.value.trim(); render(); };
  document.querySelectorAll('#mkCrm button').forEach(b => b.onclick = () => { F.crm = b.dataset.v; render(); });
  document.querySelectorAll('#mkAtt button').forEach(b => b.onclick = () => { F.att = b.dataset.v; render(); });
  if ($('#mkMin')) $('#mkMin').onchange = e => { F.min = +e.target.value; render(); };
  if ($('#mkWeb')) $('#mkWeb').onchange = e => { F.web = e.target.value; render(); };
  if ($('#mkReset')) $('#mkReset').onclick = () => { ST.mk = { q: '', crm: '', min: 1, att: '', web: '' }; render(); };
  if ($('#mkMatch')) $('#mkMatch').onclick = async e => { const b = e.currentTarget, todo = DB.mkt.filter(m => !m.guidK && m.mail && DB.contactByMail[m.mail]);
    if (!todo.length) return toast('Wszystkie pasujące zapisy mają już GUIDKontaktu', 'ok');
    b.disabled = true; let ok = 0;
    for (const m of todo) { const c = DB.contactByMail[m.mail]; try { await patchItem('MarketingKontakty', m.id, { GUIDKontaktu: c.guid }); m.guidK = c.guid; m.contact = c; ok++; } catch (er) { console.warn(er); } }
    toast(`Dopasowano ${ok} z ${todo.length} zapisów`, ok === todo.length ? 'ok' : 'err'); render(); };
  document.querySelectorAll('tr').forEach(tr => { const a = tr.querySelector('[data-mweb]'); if (a) a.onclick = () => { ST.mk.web = a.dataset.mweb; go('marketing', 'm-osoby'); }; });
};
VIEWS['m-osoby'] = () => { const all = mktPeople(), P = mFilt(all), regs = DB.mkt.length, att = DB.mkt.filter(m => m.obecny).length;
  const cols = [{ h: 'Uczestnik', key: 'name', f: p => `<b>${esc(p.name || '—')}</b><div style="font-size:11px;color:#9CA3AF">${esc(p.mail)}${p.contact && p.contact.stan ? ' · ' + esc(p.contact.stan) : ''}</div>` },
    { h: 'Firma', v: p => p.comp.name, f: p => (p.comp.crm ? compLink(p.comp.name, p.comp.guid) : esc(p.comp.name)) + (p.tecs ? ' <span class="chip c-vio" title="Aktywny TECS">TECS</span>' : '') + (p.openOpps ? ` <span class="chip c-br" title="Otwarte szanse sprzedaży">${p.openOpps} opps</span>` : '') },
    { h: 'CRM', v: p => p.crm ? 1 : 0, f: p => p.crm ? '<span class="chip c-ok">w CRM</span>' : '<span class="chip c-amb">brak</span>' },
    { h: 'Zapisy', key: 'reg', num: 1 }, { h: 'Obecny', key: 'att', num: 1 }, { h: 'Nieobecny', key: 'abs', num: 1 }, { h: 'Frekwencja', v: p => p.att / p.reg, num: 1, f: p => attChip(p.att, p.reg) },
    { h: 'Ostatni zapis', v: p => p.last, f: p => dpl(p.last) }, { h: 'Wydarzenia (zielone = był)', v: p => p.reg, f: evChips }];
  return `<div class="card panel">${ptitle('users', '#5B6CDB', '#EEF0FC', 'Uczestnicy wydarzeń', 'MarketingKontakty, osoba = adres e-mail · dopasowanie do Contacts po e-mailu (GUIDKontaktu)')}
    <div class="kpis">${kpi('Osoby', all.length, `${all.filter(p => p.crm).length} w CRM · ${all.filter(p => !p.crm).length} spoza CRM`, 'users')}${kpi('Zapisy', regs, `${pct(att, regs)} frekwencji`, 'list', '#0D9488', '#E0F5F3')}
    ${kpi('Wracają (≥2 zapisy)', all.filter(p => p.reg > 1).length, `${all.filter(p => p.att > 1).length} było ≥2 razy`, 'refresh', '#16A34A', '#E6F6EE')}${kpi('Zapisani, ani razu nie byli', all.filter(p => !p.att).length, 'do follow-upu z nagraniem', 'alert', '#D97706', '#FEF3E2')}
    ${kpi('Z firm z otwartą szansą / TECS', all.filter(p => p.openOpps || p.tecs).length, 'ciepłe kontakty dla handlowców', 'target', '#7C3AED', '#F1EAFE')}</div>
    ${mkFilters()}${table('mko', cols, P, { k: 3, asc: false })}</div>`; };
AFTER['m-osoby'] = bindMk;
VIEWS['m-firmy'] = () => { const P = mFilt(mktPeople()), by = {};
  P.forEach(p => { const k = p.comp.crm ? 'g:' + p.comp.guid : 't:' + lc(p.comp.name); const x = by[k] = by[k] || { name: p.comp.name, guid: p.comp.guid, crm: p.comp.crm, ppl: [], reg: 0, att: 0, ev: new Set(), openOpps: p.openOpps, tecs: p.tecs, last: null };
    x.ppl.push(p); x.reg += p.reg; x.att += p.att; p.rows.forEach(r => { x.ev.add(r.web.guid); if (r.web.date && (!x.last || r.web.date > x.last)) x.last = r.web.date; }); });
  const L = Object.values(by), ac = g => (DB.compBy[g] || {}).ac;
  const cols = [{ h: 'Firma', key: 'name', f: c => (c.crm ? compLink(c.name, c.guid) : esc(c.name)) + (c.crm ? '' : ' <span class="chip c-amb" title="Brak dopasowania do firmy w CRM">spoza CRM</span>') + (c.crm && ac(c.guid) ? ' <span class="chip c-mut">Academic</span>' : '') },
    { h: 'Osoby', v: c => c.ppl.length, num: 1 }, { h: 'Wydarzenia', v: c => c.ev.size, num: 1 }, { h: 'Zapisy', key: 'reg', num: 1 }, { h: 'Obecności', key: 'att', num: 1 }, { h: 'Frekwencja', v: c => c.att / c.reg, num: 1, f: c => attChip(c.att, c.reg) },
    { h: 'Ostatnio', v: c => c.last, f: c => dpl(c.last) }, { h: 'Sprzedaż', v: c => c.openOpps + (c.tecs ? 100 : 0), f: c => (c.tecs ? '<span class="chip c-vio">aktywny TECS</span>' : '') + (c.openOpps ? `<span class="chip c-br">${c.openOpps} otwarte szanse</span>` : '') || '<span style="color:#9CA3AF">—</span>' },
    { h: 'Uczestnicy', v: c => c.ppl.length, f: c => c.ppl.sort((a, b) => b.att - a.att).map(p => `<span title="${esc(p.mail)}">${esc(p.name || p.mail)} <span style="color:#9CA3AF">${p.att}/${p.reg}</span></span>`).join(', ') }];
  return `<div class="card panel">${ptitle('FIRMY', '#0D9488', '#E0F5F3', 'Firmy na wydarzeniach', 'Firma z CRM przez kontakt (Contacts → IdentyfikatorCompanies); osoby spoza CRM grupowane po polu „Firma” z formularza zapisu')}
    <div class="kpis">${kpi('Firmy z CRM', L.filter(c => c.crm).length, `${L.filter(c => !c.crm).length} nazw spoza CRM`, 'FIRMY', '#0D9488', '#E0F5F3')}${kpi('Firmy z ≥2 wydarzeń', L.filter(c => c.ev.size > 1).length, '', 'refresh', '#16A34A', '#E6F6EE')}
    ${kpi('Klienci z aktywnym TECS', L.filter(c => c.tecs).length, 'szansa na upsell / szkolenia', 'TECSY', '#7C3AED', '#F1EAFE')}${kpi('Firmy z otwartą szansą', L.filter(c => c.openOpps).length, '', 'OPPSY')}</div>
    ${mkFilters()}${table('mkf', cols, L, { k: 3, asc: false })}</div>`; };
AFTER['m-firmy'] = bindMk;
VIEWS['m-web'] = () => { const L = DB.webinars.map(w => { const r = DB.mkt.filter(m => m.web === w); if (!r.length) return null;
    const first = r.filter(m => !DB.mkt.some(o => o.mail === m.mail && o.web.date && w.date && o.web.date < w.date)).length;
    return { w, reg: r.length, att: r.filter(m => m.obecny).length, crm: r.filter(m => m.contact).length, nowi: first, ret: r.length - first, warm: r.filter(m => { const g = m.contact && m.contact.compGuid; return g && DB.sales.some(o => o.compGuid === g && o.grp === 'open'); }).length }; }).filter(Boolean);
  const cols = [{ h: 'Data', v: x => x.w.date, f: x => dpl(x.w.date) }, { h: 'Wydarzenie', v: x => x.w.title, f: x => `<span class="lnk" data-mweb="${x.w.guid}" title="Pokaż uczestników">${esc(x.w.title)}</span>${x.w.grupa ? ` <span class="chip c-mut">${esc(x.w.grupa)}</span>` : ''}` },
    { h: 'Zapisani', key: 'reg', num: 1 }, { h: 'Obecni', key: 'att', num: 1 }, { h: 'Frekwencja', v: x => x.att / x.reg, num: 1, f: x => attChip(x.att, x.reg) }, { h: 'W CRM', key: 'crm', num: 1, f: x => `${x.crm} <span style="color:#9CA3AF">(${pct(x.crm, x.reg)})</span>` },
    { h: 'Nowe osoby', key: 'nowi', num: 1 }, { h: 'Powracające', key: 'ret', num: 1 }, { h: 'Z firm z otwartą szansą', key: 'warm', num: 1 }];
  const regs = sum(L, x => x.reg), att = sum(L, x => x.att);
  return `<div class="card panel">${ptitle('KALENDARZ', '#D97706', '#FEF3E2', 'Wydarzenia (lista Webinary)', 'Kliknij tytuł, żeby zobaczyć uczestników')}
    <div class="kpis">${kpi('Wydarzenia', L.length, '', 'KALENDARZ', '#D97706', '#FEF3E2')}${kpi('Zapisy', regs, `średnio ${Math.round(regs / Math.max(1, L.length))} na wydarzenie`, 'list')}${kpi('Obecni', att, pct(att, regs) + ' frekwencji', 'check', '#16A34A', '#E6F6EE')}</div>
    ${table('mkw', cols, L, { k: 0, asc: false })}</div>`; };
AFTER['m-web'] = bindMk;

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

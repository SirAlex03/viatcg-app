/**
 * V.I.A. TCG — Tracker investimento carte e prodotti · versione 3.2
 * Backend Google Apps Script. I dati vivono nel Google Sheet indicato in CONFIG.SPREADSHEET_ID.
 */
const CONFIG = {
  PIN_INIZIALE: '',                  // PIN di partenza (impostato solo nell'editor di Apps Script, mai qui)
  SOCI: ['Alex', 'Vitto', 'Ivo'],    // quote uguali: 1/3 a testa
  NOME_APP: 'V.I.A. TCG',
  BACKUP_DA_TENERE: 12,              // backup settimanali conservati
  SPREADSHEET_ID: '1dtjNeu2U6Cx1hSZ0OzFPOhz_jPpZNztVaDpZIHGRMiQ',  // foglio "V.I.A. TCG – dati"
  TZ: 'Europe/Rome'
};
const VERSIONE_DATI = '3';

/* ------------------------------------------------------------------ */
/* Struttura dei fogli                                                 */
/* ------------------------------------------------------------------ */
const COMMON = ['CreatoDa', 'CreatoIl', 'ModificatoDa', 'ModificatoIl', 'Archiviato'];
const SCHEMA = {
  Prodotti: ['ID', 'DataAcquisto', 'Nome', 'Categoria', 'CanaleAcquisto', 'PrezzoAcquisto', 'PagatoDa',
    'DataVendita', 'CanaleVendita', 'PrezzoVendita', 'IncassatoDa', 'FotoId', 'Note'].concat(COMMON).concat([
    'LottoID', 'Origine', 'ParentID', 'Condizione', 'Lingua', 'Aperto', 'DataApertura',
    'GradStato', 'GradCasa', 'GradVoto', 'GradCert', 'GradDataInvio', 'GradDataRientro', 'GradCosto', 'GradPagatoDa',
    'ScambioOutID', 'ScambioInID', 'InArrivo', 'DataArrivo', 'DaIncassare', 'DataIncasso', 'ImportoIncassato']),
  PrimaNota: ['ID', 'Data', 'Tipo', 'Descrizione', 'Importo', 'Socio', 'Categoria', 'RifProdotto', 'RifPrestito',
    'NumRata', 'QuotaCapitale', 'QuotaInteressi', 'Auto'].concat(COMMON).concat([
    'PagatoDa', 'RifLotto', 'RifOrigine', 'RifScambio', 'Allegati', 'DaIncassare', 'DataIncasso']),
  Prestiti: ['ID', 'Descrizione', 'Banca', 'Importo', 'TassoAnnuo', 'NumRate', 'RataFissa', 'DataErogazione',
    'DataPrimaRata', 'Note'].concat(COMMON),
  Canali: ['ID', 'Nome', 'Tipo', 'Uso', 'Note'].concat(COMMON).concat(['IncassoDifferito']),
  Scambi: ['ID', 'Data', 'Controparte', 'Conguaglio', 'PagatoDa', 'Note'].concat(COMMON),
  Log: ['Timestamp', 'Utente', 'Azione', 'Foglio', 'ID', 'Dettaglio', 'Cassa', 'Crediti'],
  Soci: ['Socio', 'PIN', 'DaCambiare', 'UltimaModifica', 'TentativiErrati', 'BloccatoFino', 'NovitaLette']
};
const DATE_COLS = {
  Prodotti: ['DataAcquisto', 'DataVendita', 'DataApertura', 'GradDataInvio', 'GradDataRientro', 'DataArrivo', 'DataIncasso'],
  PrimaNota: ['Data', 'DataIncasso'], Prestiti: ['DataErogazione', 'DataPrimaRata'], Scambi: ['Data']
};
const NUM_COLS = {
  Prodotti: ['PrezzoAcquisto', 'PrezzoVendita', 'GradCosto', 'ImportoIncassato'],
  PrimaNota: ['Importo', 'NumRata', 'QuotaCapitale', 'QuotaInteressi'],
  Prestiti: ['Importo', 'TassoAnnuo', 'NumRate', 'RataFissa'], Scambi: ['Conguaglio'], Log: ['Cassa', 'Crediti']
};
const BOOL_COLS = ['Archiviato', 'Auto', 'Aperto', 'DaCambiare', 'InArrivo', 'DaIncassare', 'IncassoDifferito'];
const TEXT_COLS = ['NovitaLette', 'ID', 'LottoID', 'ParentID', 'RifProdotto', 'RifPrestito', 'RifLotto', 'RifOrigine', 'RifScambio',
  'ScambioOutID', 'ScambioInID', 'GradVoto', 'GradCert', 'Allegati', 'PIN', 'BloccatoFino', 'UltimaModifica'];

const SIGILLATI = ['Box', 'Display', 'Booster / bustine', 'Starter deck', 'Collection / premium'];
const CONDIZIONI = ['M', 'NM', 'EX', 'GD', 'LP', 'PL', 'PO'];
const GRAD_STATI = ['DA_INVIARE', 'INVIATA', 'GRADATA'];
const TIPI_MANUALI = ['VERSAMENTO', 'PRELIEVO', 'SPESA', 'ENTRATA'];
const TIPI_CANALE = ['Negozio', 'Sito / piattaforma', 'Privato', 'Fiera / evento', 'Altro'];
const USI_CANALE = ['ACQUISTO', 'VENDITA', 'ENTRAMBI'];
function pagatori_() { return ['CASSA', 'TUTTI'].concat(CONFIG.SOCI); }
function incassatori_() { return ['CASSA'].concat(CONFIG.SOCI); }

/* ------------------------------------------------------------------ */
/* Web app                                                             */
/* ------------------------------------------------------------------ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.NOME_APP)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Accesso via fetch (servirà per l'interfaccia installabile come app). */
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { /* richiesta non valida */ }
  return ContentService.createTextOutput(JSON.stringify(api(req))).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('V.I.A. TCG')
    .addItem('Esegui setup / aggiornamento', 'setup')
    .addItem('Fai un backup adesso', 'backupSettimanale')
    .addToUi();
}

/** Unico punto di ingresso chiamato dal sito. */
function api(req) {
  DB_ = {};
  try {
    req = req || {};
    const a = auth_(req);
    if (!a.ok) return a;
    const user = a.user;
    const p = req.payload || {};
    if (req.action === 'cambiaPin') return withLock_(function () { cambiaPin_(user, req.pin, p); flush_(); return { ok: true }; });
    if (a.daCambiare) return { ok: false, code: 'CAMBIA_PIN', error: 'Al primo accesso devi scegliere il tuo PIN personale' };
    if (req.action === 'novitaLette') return withLock_(function () { aggiornaSocio_(socio_(user), { NovitaLette: String(p.versione || '') }); flush_(); return { ok: true }; });

    switch (req.action) {
      case 'ping': return { ok: true, soci: CONFIG.SOCI };
      case 'getRev': return { ok: true, rev: getRev_() };
      case 'getAll': return { ok: true, data: getAll_(user) };
      case 'attivita': return { ok: true, attivita: attivita_(p) };
      case 'getFoto': return { ok: true, dataUrl: getFile_(p.id, 'FOTO_FOLDER_ID') };
      case 'getAllegato': return { ok: true, dataUrl: getFile_(p.id, 'RICEVUTE_FOLDER_ID') };
      case 'uploadFoto': return { ok: true, id: uploadFoto_(p, user) };
      case 'uploadAllegato': return { ok: true, allegato: uploadAllegato_(p, user) };
    }

    return withLock_(function () {
      let result = null;
      const cassa0 = cassa_(), crediti0 = crediti_();
      switch (req.action) {
        case 'saveProdotto': result = saveProdotto_(p, user); break;
        case 'modificaMultipla': result = modificaMultipla_(p, user); break;
        case 'vendiMultipla': result = vendiMultipla_(p, user); break;
        case 'archiviaMultipla': result = archiviaMultipla_(p, user); break;
        case 'arrivo': result = arrivo_(p, user); break;
        case 'incasso': result = incasso_(p, user); break;
        case 'apri': result = apri_(p, user); break;
        case 'annullaApertura': result = annullaApertura_(p, user); break;
        case 'gradazione': result = gradazione_(p, user); break;
        case 'scambio': result = scambio_(p, user); break;
        case 'annullaScambio': result = annullaScambio_(p, user); break;
        case 'saveMovimento': result = saveMovimento_(p, user); break;
        case 'setAllegati': result = setAllegati_(p, user); break;
        case 'savePrestito': result = savePrestito_(p, user); break;
        case 'pagaRata': result = pagaRata_(p, user); break;
        case 'saveCanale': result = saveCanale_(p, user); break;
        case 'archivia': result = setArchiviato_(p.foglio, p.id, true, user); break;
        case 'ripristina': result = setArchiviato_(p.foglio, p.id, false, user); break;
        default: throw new Error('Azione sconosciuta: ' + req.action);
      }
      segnaCassa_(cassa0, crediti0);
      flush_();
      bumpRev_();
      return { ok: true, result: result, data: getAll_(user) };
    });
  } catch (e) {
    DB_ = {};
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/** Liquidità in cassa (stessa regola del report: le vendite da incassare non contano). */
function cassa_() {
  const S = { VERSAMENTO: 1, PRELIEVO: -1, ACQUISTO: -1, VENDITA: 1, SPESA: -1, ENTRATA: 1, GRADING: -1,
    CONGUAGLIO_OUT: -1, CONGUAGLIO_IN: 1, EROGAZIONE: 1, RATA: -1 };
  return r2_(rows_('PrimaNota').reduce(function (a, m) {
    if (!isAct_(m) || !S[m.Tipo] || (m.Tipo === 'VENDITA' && m.DaIncassare === true)) return a;
    return a + S[m.Tipo] * (Number(m.Importo) || 0);
  }, 0));
}

/** Scrive sulle attività registrate in questa operazione di quanto è cambiata la cassa. */
function crediti_() {
  return r2_(rows_('PrimaNota').reduce(function (a, m) {
    return (isAct_(m) && m.Tipo === 'VENDITA' && m.DaIncassare === true) ? a + (Number(m.Importo) || 0) : a;
  }, 0));
}
function segnaCassa_(cassa0, crediti0) {
  const t = DB_.Log; if (!t || !t.added.length) return;
  const d = r2_(cassa_() - cassa0), c = r2_(crediti_() - crediti0), last = t.added[t.added.length - 1];
  last.Cassa = Math.abs(d) < 0.005 ? '' : d;
  last.Crediti = Math.abs(c) < 0.005 ? '' : c;
}

/**
 * v3: ricostruisce l'effetto su cassa e "da incassare" delle attività registrate prima della v3.2,
 * abbinando a ogni attività i movimenti creati o archiviati nello stesso momento dallo stesso socio.
 * Le modifiche di importo fatte in passato non sono ricostruibili e restano senza cifra.
 */
function migraV3_() {
  const SG = { VERSAMENTO: 1, PRELIEVO: -1, ACQUISTO: -1, VENDITA: 1, SPESA: -1, ENTRATA: 1, GRADING: -1,
    CONGUAGLIO_OUT: -1, CONGUAGLIO_IN: 1, EROGAZIONE: 1, RATA: -1 };
  const t = tbl_('Log'), L = t.rows;
  const sec = function (s) { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/);
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]) / 1000 : null; };
  const idx = L.map(function (g, i) { return { i: i, s: sec(g.Timestamp), u: g.Utente }; }).filter(function (x) { return x.s !== null; });
  const eff = L.map(function () { return { c: 0, k: 0, n: 0 }; });
  const assegna = function (quando, chi, c, k) {
    const s = sec(quando); if (s === null) return;
    let best = null;
    idx.forEach(function (x) { if (x.u !== chi) return; const d = x.s - s; if (d >= -3 && d <= 5 && (!best || Math.abs(d) < Math.abs(best.d))) best = { i: x.i, d: d }; });
    if (best) { eff[best.i].c += c; eff[best.i].k += k; eff[best.i].n++; }
  };
  rows_('PrimaNota').forEach(function (m) {
    const sg = SG[m.Tipo]; if (!sg) return;
    const imp = Number(m.Importo) || 0;
    const nuovaPendente = m.Tipo === 'VENDITA' && m.DaIncassare === true && Math.abs((sec(m.ModificatoIl) || 0) - (sec(m.CreatoIl) || 0)) <= 5;
    assegna(m.CreatoIl, m.CreatoDa, nuovaPendente ? 0 : sg * imp, nuovaPendente ? imp : 0);
    if (m.Archiviato === true) {
      const pend = m.Tipo === 'VENDITA' && m.DaIncassare === true;
      assegna(m.ModificatoIl, m.ModificatoDa, pend ? 0 : -sg * imp, pend ? -imp : 0);
    }
  });
  // vendite spostate tra "incassato" e "da incassare"
  const vend = {};
  rows_('PrimaNota').forEach(function (m) { if (isAct_(m) && m.Tipo === 'VENDITA' && m.RifProdotto) vend[m.RifProdotto] = Number(m.Importo) || 0; });
  L.forEach(function (g, i) {
    const imp = vend[g.ID]; if (imp == null) return;
    if (/^Da incassare 1 vendit/.test(g.Azione)) { eff[i].c -= imp; eff[i].k += imp; eff[i].n++; }
    if (/^Incassate 1 vendit/.test(g.Azione)) { eff[i].c += imp; eff[i].k -= imp; eff[i].n++; }
  });
  let n = 0;
  L.forEach(function (g, i) {
    if (!eff[i].n || (g.Cassa !== '' && g.Cassa != null) || (g.Crediti !== '' && g.Crediti != null)) return;
    const c = r2_(eff[i].c), k = r2_(eff[i].k);
    if (Math.abs(c) < 0.005 && Math.abs(k) < 0.005) return;
    g.Cassa = Math.abs(c) < 0.005 ? '' : c; g.Crediti = Math.abs(k) < 0.005 ? '' : k;
    if (g._row) t.dirty[g._row] = g;
    n++;
  });
  return ['Attività passate con importo ricostruito: ' + n];
}

/** Attività più vecchie, a pagine: offset = quante ne hai già, n = quante ancora. */
function attivita_(p) {
  const L = rows_('Log'), off = Math.max(0, Number(p.offset) || 0), n = Math.min(100, Math.max(1, Number(p.n) || 20));
  const fine = L.length - off, inizio = Math.max(0, fine - n);
  return { righe: L.slice(inizio, fine).reverse().map(pulisci_), altre: inizio > 0 };
}
function pulisci_(r) { const o = {}; Object.keys(r).forEach(function (k) { if (k !== '_row') o[k] = r[k]; }); return o; }

function getAll_(user) {
  const clean = function (rows) {
    return rows.map(function (r) { const o = {}; Object.keys(r).forEach(function (k) { if (k !== '_row') o[k] = r[k]; }); return o; });
  };
  return {
    rev: getRev_(),
    soci: CONFIG.SOCI,
    prodotti: clean(rows_('Prodotti')),
    primaNota: clean(rows_('PrimaNota')),
    prestiti: clean(rows_('Prestiti')),
    canali: clean(rows_('Canali')),
    scambi: clean(rows_('Scambi')),
    log: clean(rows_('Log').slice(-25).reverse()),
    logTotale: rows_('Log').length,
    io: { novita: user && socio_(user) ? String(socio_(user).NovitaLette || '') : '' },
    oggi: today_()
  };
}

/* ------------------------------------------------------------------ */
/* Archivio dati in memoria (una lettura per foglio, una scrittura)    */
/* ------------------------------------------------------------------ */
let DB_ = {};

function tbl_(name) {
  if (DB_[name]) return DB_[name];
  const sh = sheet_(name);
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  const vals = (lr > 0 && lc > 0) ? sh.getRange(1, 1, lr, lc).getValues() : [[]];
  const headers = (vals[0] || []).map(String);
  const t = { name: name, sh: sh, headers: headers, rows: [], byId: {}, dirty: {}, added: [], raw: {}, lastRow: lr };
  const dc = DATE_COLS[name] || [], nc = NUM_COLS[name] || [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[0] === '' || r[0] === null) { t.raw[i + 1] = r; continue; }
    const o = {};
    headers.forEach(function (h, j) { if (h) o[h] = normCell_(r[j], h, dc, nc); });
    o._row = i + 1;
    t.rows.push(o);
    if (o.ID) t.byId[o.ID] = o;
  }
  DB_[name] = t;
  return t;
}

function normCell_(v, h, dc, nc) {
  if (isDate_(v)) return Utilities.formatDate(v, CONFIG.TZ, dc.indexOf(h) >= 0 ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm:ss');
  if (v === null || v === undefined) return '';
  if (dc.indexOf(h) >= 0) return String(v).slice(0, 10);
  if (nc.indexOf(h) >= 0 && typeof v === 'string' && v !== '') {
    const n = Number(v.replace(/\s/g, '').replace(',', '.'));
    return isNaN(n) ? v : n;
  }
  if (BOOL_COLS.indexOf(h) >= 0) {
    if (v === 'TRUE' || v === 'true') return true;
    if (v === 'FALSE' || v === 'false') return false;
  }
  if (TEXT_COLS.indexOf(h) >= 0 && typeof v === 'number') return String(v);
  return v;
}

function rows_(name) { return tbl_(name).rows; }
function get_(name, id) { return id ? (tbl_(name).byId[id] || null) : null; }
function isAct_(o) { return !!o && o.Archiviato !== true; }

/** Inserisce o aggiorna (per ID) una riga nella memoria; la scrittura avviene in flush_(). */
function put_(name, rec, user) {
  const t = tbl_(name);
  const now = nowStr_();
  let o = rec.ID ? t.byId[rec.ID] : null;
  const isNew = !o;
  if (isNew) {
    o = {};
    t.headers.forEach(function (h) { if (h) o[h] = ''; });
    o.CreatoDa = user; o.CreatoIl = now; o.Archiviato = false;
    t.rows.push(o); t.added.push(o);
  }
  Object.keys(rec).forEach(function (k) { o[k] = rec[k]; });
  if (o.ID) t.byId[o.ID] = o;
  o.ModificatoDa = user; o.ModificatoIl = now;
  if (!isNew && o._row) t.dirty[o._row] = o;
  return o;
}

/** Aggiorna solo se qualcosa cambia davvero. */
function patch_(name, id, fields, user) {
  const o = get_(name, id);
  if (!o) return null;
  const changed = Object.keys(fields).some(function (k) { return String(o[k]) !== String(fields[k]); });
  if (!changed) return o;
  return put_(name, Object.assign({ ID: id }, fields), user);
}

function log_(user, azione, foglio, id, dettaglio) {
  const t = tbl_('Log');
  const o = { Timestamp: nowStr_(), Utente: user, Azione: azione, Foglio: foglio, ID: id,
    Dettaglio: String(dettaglio || '').slice(0, 300) };
  t.rows.push(o); t.added.push(o);
}

function flush_() {
  Object.keys(DB_).forEach(function (name) {
    const t = DB_[name];
    const H = t.headers;
    const toRow = function (o) { return H.map(function (h) { return h ? cellOut_(o[h]) : ''; }); };
    const dirty = Object.keys(t.dirty).map(function (k) { return t.dirty[k]; });
    if (dirty.length > 15) {
      const map = {};
      t.rows.forEach(function (o) { if (o._row) map[o._row] = o; });
      const block = [];
      for (let r = 2; r <= t.lastRow; r++) {
        if (map[r]) block.push(toRow(map[r]));
        else if (t.raw[r]) block.push(H.map(function (h, j) { return t.raw[r][j] === undefined ? '' : t.raw[r][j]; }));
        else block.push(H.map(function () { return ''; }));
      }
      if (block.length) t.sh.getRange(2, 1, block.length, H.length).setValues(block);
    } else {
      dirty.forEach(function (o) { t.sh.getRange(o._row, 1, 1, H.length).setValues([toRow(o)]); });
    }
    if (t.added.length) {
      const start = t.lastRow + 1, n = t.added.length;
      ensureRows_(t.sh, start + n - 1, name);
      t.sh.getRange(start, 1, n, H.length).setValues(t.added.map(toRow));
      t.added.forEach(function (o, i) { o._row = start + i; });
      t.lastRow += n;
    }
    t.dirty = {}; t.added = [];
  });
}

function cellOut_(v) {
  if (v === undefined || v === null) return '';
  return safe_(v);
}

function ensureRows_(sh, needed, name) {
  const max = sh.getMaxRows();
  if (needed > max) {
    sh.insertRowsAfter(max, needed - max + 200);
    applyFormats_(sh, name);
  }
}

function applyFormats_(sh, name) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  (DATE_COLS[name] || []).concat(TEXT_COLS).forEach(function (h) {
    const col = headers.indexOf(h) + 1;
    if (col > 0) sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  });
}

/* ------------------------------------------------------------------ */
/* Prodotti                                                            */
/* ------------------------------------------------------------------ */
function isSealed_(cat) { return SIGILLATI.indexOf(cat) >= 0; }
function isSold_(p) { return p.PrezzoVendita !== '' && p.PrezzoVendita !== null && p.PrezzoVendita !== undefined; }
function inMagazzino_(p) { return isAct_(p) && !isSold_(p) && !p.ScambioOutID && p.Aperto !== true; }
function figli_(parentId) {
  return rows_('Prodotti').filter(function (c) { return isAct_(c) && c.ParentID === parentId; });
}
function err_(msg) { throw new Error(msg); }

function saveProdotto_(p, user) {
  const ex = p.ID ? get_('Prodotti', p.ID) : null;
  if (p.ID && !ex) err_('Prodotto non trovato');
  if (ex && !isAct_(ex)) err_('Il prodotto è stato eliminato: ripristinalo dall\'Archivio');
  const origine = ex ? (ex.Origine || 'ACQUISTO') : (p.Origine === 'SBUSTATA' ? 'SBUSTATA' : 'ACQUISTO');

  const nome = str_(p.Nome);
  if (!nome) err_('Il nome del prodotto è obbligatorio');
  const cat = str_(p.Categoria);
  if (!cat) err_('Seleziona la categoria');
  if (ex && ex.Aperto === true && !isSealed_(cat)) err_('Il prodotto è stato aperto: deve restare un prodotto sigillato');
  const cond = str_(p.Condizione);
  if (cond && CONDIZIONI.indexOf(cond) < 0) err_('Condizione non valida');
  const base = { Nome: nome, Categoria: cat, Condizione: isSealed_(cat) ? '' : cond,
    Lingua: str_(p.Lingua), Note: str_(p.Note), FotoId: str_(p.FotoId) };

  let acq = {};
  let parent = null;
  if (origine === 'ACQUISTO') {
    const data = date_(p.DataAcquisto);
    if (!data) err_('Inserisci la data di acquisto');
    const canale = str_(p.CanaleAcquisto);
    if (!canale) err_('Seleziona il canale di acquisto');
    const pag = str_(p.PagatoDa);
    if (pagatori_().indexOf(pag) < 0) err_('Indica chi ha pagato');
    const prezzo = num_(p.PrezzoAcquisto, 'Prezzo di acquisto');
    if (prezzo === '' || prezzo < 0) err_('Inserisci il prezzo di acquisto (anche 0)');
    acq = { DataAcquisto: data, CanaleAcquisto: canale, PagatoDa: pag, PrezzoAcquisto: prezzo };
  } else if (origine === 'SBUSTATA') {
    parent = get_('Prodotti', ex ? ex.ParentID : str_(p.ParentID));
    if (!parent || !isAct_(parent)) err_('Seleziona il prodotto sigillato da cui è stata sbustata');
    if (!isSealed_(parent.Categoria)) err_('Il prodotto di provenienza non è un prodotto sigillato');
    if (isSold_(parent) || parent.ScambioOutID) err_('Il prodotto di provenienza è stato venduto o scambiato');
    const data = date_(p.DataAcquisto) || parent.DataApertura;
    if (!data) err_('Inserisci la data di apertura');
    if (data < parent.DataAcquisto) err_('La data di apertura è precedente all\'acquisto del prodotto sigillato');
    acq = { DataAcquisto: data, ParentID: parent.ID, CanaleAcquisto: '', PagatoDa: '' };
  } else {
    const data = date_(p.DataAcquisto);
    if (!data) err_('Inserisci la data dello scambio');
    const v = num_(p.PrezzoAcquisto, 'Valore');
    if (v === '' || v < 0) err_('Inserisci il valore attribuito al prodotto (anche 0)');
    acq = { DataAcquisto: data, PrezzoAcquisto: v };
  }

  const prezzoV = num_(p.PrezzoVendita, 'Prezzo di vendita');
  let vend = { DataVendita: '', CanaleVendita: '', PrezzoVendita: '', IncassatoDa: '' };
  if (prezzoV !== '') {
    if (prezzoV < 0) err_('Prezzo di vendita non valido');
    if (ex && ex.Aperto === true) err_('Questo prodotto è stato aperto: si vendono le carte trovate, non il prodotto');
    if (ex && ex.ScambioOutID) err_('Questo prodotto è già stato dato in uno scambio');
    const dv = date_(p.DataVendita);
    if (!dv) err_('Inserisci la data di vendita');
    const cv = str_(p.CanaleVendita);
    if (!cv) err_('Seleziona il canale di vendita');
    const inc = str_(p.IncassatoDa);
    if (incassatori_().indexOf(inc) < 0) err_('Indica chi ha incassato la vendita');
    if (dv < acq.DataAcquisto) err_('La data di vendita è precedente alla data di acquisto');
    vend = { DataVendita: dv, CanaleVendita: cv, PrezzoVendita: prezzoV, IncassatoDa: inc };
    if (Object.prototype.hasOwnProperty.call(p, 'DaIncassare')) Object.assign(vend, statoIncasso_(p, [prezzoV], dv)[0]);
    else if (ex && isSold_(ex)) Object.assign(vend, ex.DaIncassare === true ? statoIncasso_({ DaIncassare: true }, [prezzoV], dv)[0]
      : statoIncasso_({ DaIncassare: false, DataIncasso: ex.DataIncasso || dv, ImportoIncassato: ex.ImportoIncassato === '' ? '' : Math.min(Number(ex.ImportoIncassato), prezzoV) }, [prezzoV], dv)[0]);
  }
  if (prezzoV === '') Object.assign(vend, { DaIncassare: '', DataIncasso: '', ImportoIncassato: '' });

  const ids = [];
  if (!ex) {
    const q = parseInt(p.Quantita || 1, 10);
    if (!(q >= 1 && q <= 200)) err_('Quantità non valida (da 1 a 200)');
    if (prezzoV !== '' && q > 1) err_('Con più pezzi registra prima l\'acquisto, poi vendi i singoli pezzi');
    const prezzi = [];
    if (origine === 'ACQUISTO') {
      if (p.PrezzoModo === 'totale' && q > 1) {
        const unit = Math.floor(acq.PrezzoAcquisto * 100 / q) / 100;
        for (let i = 0; i < q; i++) prezzi.push(i < q - 1 ? unit : r2_(acq.PrezzoAcquisto - unit * (q - 1)));
      } else {
        for (let i = 0; i < q; i++) prezzi.push(acq.PrezzoAcquisto);
      }
    }
    const grad = gradAcquisto_(p, cat);
    const lotto = newId_('T');
    for (let i = 0; i < q; i++) {
      const rec = Object.assign({ ID: newId_('P'), LottoID: lotto, Origine: origine, Aperto: false }, base, acq, vend, grad);
      if (origine === 'ACQUISTO' && p.InArrivo === true) { rec.InArrivo = true; rec.DataArrivo = ''; }
      rec.PrezzoAcquisto = origine === 'ACQUISTO' ? prezzi[i] : 0;
      put_('Prodotti', rec, user);
      ids.push(rec.ID);
    }
    if (origine === 'ACQUISTO') syncLotto_(lotto, user, p.AllegatiAcquisto);
    if (origine === 'SBUSTATA') { apriSeServe_(parent, acq.DataAcquisto, user); ricalcolaQuote_(parent.ID, user); }
    ids.forEach(function (id) { syncVendita_(get_('Prodotti', id), user); });
    log_(user, 'Nuovo prodotto' + (q > 1 ? ' ×' + q : ''), 'Prodotti', ids[0], nome);
  } else {
    const upd = Object.assign({ ID: ex.ID }, base, acq, vend);
    if (Object.prototype.hasOwnProperty.call(p, 'InArrivo')) Object.assign(upd, statoArrivo_(ex, p.InArrivo !== true, p.DataArrivo));
    if (origine === 'SBUSTATA') delete upd.PrezzoAcquisto;
    put_('Prodotti', upd, user);
    ids.push(ex.ID);
    if (origine === 'ACQUISTO') {
      if (p.ApplicaLotto && ex.LottoID) {
        rows_('Prodotti').forEach(function (x) {
          if (x.ID !== ex.ID && isAct_(x) && x.LottoID === ex.LottoID && (x.Origine || 'ACQUISTO') === 'ACQUISTO') {
            patch_('Prodotti', x.ID, { Nome: nome, Categoria: cat, DataAcquisto: acq.DataAcquisto,
              CanaleAcquisto: acq.CanaleAcquisto, PagatoDa: acq.PagatoDa }, user);
          }
        });
      }
      syncLotto_(ex.LottoID, user, p.AllegatiAcquisto);
    }
    if (origine === 'SBUSTATA') ricalcolaQuote_(ex.ParentID, user);
    if (ex.Aperto === true) ricalcolaQuote_(ex.ID, user);
    syncVendita_(get_('Prodotti', ex.ID), user);
    log_(user, 'Modifica prodotto', 'Prodotti', ex.ID, nome);
  }
  if (acq.CanaleAcquisto) ensureCanale_(acq.CanaleAcquisto, 'ACQUISTO', user);
  if (vend.CanaleVendita) ensureCanale_(vend.CanaleVendita, 'VENDITA', user);
  return { ids: ids };
}

function gradAcquisto_(p, cat) {
  const casa = str_(p.GradCasa);
  if (!casa || isSealed_(cat)) return {};
  const voto = str_(p.GradVoto);
  if (!voto) err_('Inserisci il voto della carta gradata');
  return { GradStato: 'GRADATA', GradCasa: casa, GradVoto: voto, GradCert: str_(p.GradCert), GradCosto: '' };
}

/** Un solo movimento di acquisto per lotto, con importo = somma dei pezzi attivi. */
function syncLotto_(lotto, user, allegati) {
  if (!lotto) return;
  const pezzi = rows_('Prodotti').filter(function (x) {
    return isAct_(x) && x.LottoID === lotto && (x.Origine || 'ACQUISTO') === 'ACQUISTO';
  });
  if (!pezzi.length) {
    autoMov_('ACQUISTO', 'RifLotto', lotto, null, user);
    syncVersAuto_('L:' + lotto, null, user);
    return;
  }
  const f = pezzi[0];
  const tot = r2_(pezzi.reduce(function (s, x) { return s + (Number(x.PrezzoAcquisto) || 0); }, 0));
  const data = pezzi.map(function (x) { return x.DataAcquisto; }).sort()[0];
  const descr = 'Acquisto: ' + f.Nome + (pezzi.length > 1 ? ' ×' + pezzi.length : '') + (f.CanaleAcquisto ? ' (' + f.CanaleAcquisto + ')' : '');
  const fields = { Data: data, Descrizione: descr, Importo: tot, PagatoDa: f.PagatoDa || 'CASSA', Socio: '',
    RifProdotto: pezzi.length === 1 ? f.ID : '' };
  if (allegati) fields.Allegati = allegatiJson_(allegati);
  autoMov_('ACQUISTO', 'RifLotto', lotto, fields, user);
  syncVersAuto_('L:' + lotto, { pagatoDa: f.PagatoDa, importo: tot, data: data, descr: descr }, user);
}

function syncVendita_(x, user) {
  if (!x) return;
  if (!isAct_(x) || !isSold_(x)) {
    autoMov_('VENDITA', 'RifProdotto', x.ID, null, user);
    autoMov_('SPESA', 'RifProdotto', x.ID, null, user);
    return;
  }
  const daInc = x.DaIncassare === true;
  autoMov_('VENDITA', 'RifProdotto', x.ID, {
    Data: x.DataVendita, Importo: Number(x.PrezzoVendita),
    Descrizione: 'Vendita: ' + x.Nome + (x.CanaleVendita ? ' (' + x.CanaleVendita + ')' : ''),
    Socio: x.IncassatoDa === 'CASSA' ? '' : x.IncassatoDa,
    DaIncassare: daInc, DataIncasso: daInc ? '' : (x.DataIncasso || x.DataVendita)
  }, user);
  const comm = (!daInc && x.ImportoIncassato !== '' && x.ImportoIncassato != null) ? r2_(Number(x.PrezzoVendita) - Number(x.ImportoIncassato)) : 0;
  autoMov_('SPESA', 'RifProdotto', x.ID, comm > 0 ? {
    Data: x.DataIncasso || x.DataVendita, Importo: comm, Categoria: 'Commissioni piattaforme', PagatoDa: 'CASSA', Socio: '',
    Descrizione: 'Commissioni: ' + x.Nome + (x.CanaleVendita ? ' (' + x.CanaleVendita + ')' : '')
  } : null, user);
}

/* ------------------------------------------------------------------ */
/* Arrivo della merce e incasso delle vendite                          */
/* ------------------------------------------------------------------ */
/** Campi di incasso per una o più vendite. prezzi = prezzi di vendita dei pezzi, nello stesso ordine. */
function statoIncasso_(p, prezzi, dataVendita) {
  if (p.DaIncassare === true) return prezzi.map(function () { return { DaIncassare: true, DataIncasso: '', ImportoIncassato: '' }; });
  const di = date_(p.DataIncasso) || dataVendita;
  if (dataVendita && di < dataVendita) err_('La data di incasso è precedente alla data di vendita');
  const totV = r2_(prezzi.reduce(function (a, b) { return a + Number(b); }, 0));
  const ric = num_(p.ImportoIncassato, 'Importo incassato');
  if (ric !== '' && ric < 0) err_('Importo incassato non valido');
  if (ric !== '' && ric > totV + 0.001) err_('L\'importo incassato supera il prezzo di vendita: se hai incassato di più, aumenta il prezzo');
  let quote;
  if (ric === '' || prezzi.length === 1 || totV === 0) {
    quote = prezzi.length === 1 ? [ric === '' ? Number(prezzi[0]) : ric] : (ric === '' ? prezzi.map(Number) : dividi_(ric, prezzi.length));
  } else {
    // ripartisce l'incassato in proporzione ai prezzi, resto sull'ultimo
    let acc = 0;
    quote = prezzi.map(function (v, i) {
      if (i === prezzi.length - 1) return r2_(ric - acc);
      const q = Math.floor(ric * Number(v) / totV * 100) / 100; acc = r2_(acc + q); return q;
    });
  }
  return quote.map(function (q) { return { DaIncassare: false, DataIncasso: di, ImportoIncassato: r2_(q) }; });
}

function statoArrivo_(x, arrivato, data) {
  if (arrivato) return { InArrivo: false, DataArrivo: date_(data) || x.DataArrivo || '' };
  if (x.Aperto === true) err_('"' + x.Nome + '" è già stato aperto: non può essere ancora in arrivo');
  if (x.GradStato === 'INVIATA' || x.GradStato === 'GRADATA') err_('"' + x.Nome + '" risulta già inviata a gradare');
  return { InArrivo: true, DataArrivo: '' };
}

/** Segna più pezzi come arrivati (o di nuovo in arrivo). */
function arrivo_(p, user) {
  const pezzi = pezziSelezionati_(p.ids);
  const arrivato = p.arrivato !== false;
  const data = arrivato ? (date_(p.data) || today_()) : '';
  pezzi.forEach(function (x) {
    if (arrivato && data < x.DataAcquisto) err_('La data di arrivo è precedente all\'acquisto di "' + x.Nome + '"');
    put_('Prodotti', Object.assign({ ID: x.ID }, statoArrivo_(x, arrivato, data)), user);
  });
  log_(user, (arrivato ? 'Arrivati ' : 'In arrivo ') + pezzi.length + ' pezzi', 'Prodotti', pezzi[0].ID, pezzi[0].Nome + (pezzi.length > 1 ? ' e altri' : ''));
  return { n: pezzi.length };
}

/** Segna più vendite come incassate (con importo realmente ricevuto) o di nuovo da incassare. */
function incasso_(p, user) {
  const pezzi = pezziSelezionati_(p.ids);
  pezzi.forEach(function (x) { if (!isSold_(x)) err_('"' + x.Nome + '" non risulta venduto'); });
  const incassato = p.incassato !== false;
  const maxDV = pezzi.map(function (x) { return x.DataVendita; }).sort().pop();
  const stati = statoIncasso_({ DaIncassare: !incassato, DataIncasso: p.data || today_(), ImportoIncassato: p.importo },
    pezzi.map(function (x) { return Number(x.PrezzoVendita); }), maxDV);
  pezzi.forEach(function (x, i) {
    put_('Prodotti', Object.assign({ ID: x.ID }, stati[i]), user);
    syncVendita_(get_('Prodotti', x.ID), user);
  });
  log_(user, (incassato ? 'Incassate ' : 'Da incassare ') + pezzi.length + ' vendite', 'Prodotti', pezzi[0].ID, pezzi[0].Nome + (pezzi.length > 1 ? ' e altri' : ''));
  return { n: pezzi.length };
}

/* ------------------------------------------------------------------ */
/* Apertura dei prodotti sigillati                                     */
/* ------------------------------------------------------------------ */
function apriSeServe_(parent, data, user) {
  if (parent.Aperto === true) return;
  if (!inMagazzino_(parent)) err_('Il prodotto sigillato non è disponibile in magazzino');
  if (parent.InArrivo === true) err_('"' + parent.Nome + '" non è ancora arrivato: segnalo come arrivato prima di aprirlo');
  put_('Prodotti', { ID: parent.ID, Aperto: true, DataApertura: data }, user);
}

/** Divide il costo del sigillato in parti uguali tra le carte trovate (resto sull'ultima). */
function ricalcolaQuote_(parentId, user) {
  const parent = get_('Prodotti', parentId);
  if (!parent) return;
  const kids = figli_(parentId).slice().sort(function (a, b) {
    return String(a.CreatoIl + a.ID) < String(b.CreatoIl + b.ID) ? -1 : 1;
  });
  if (!kids.length) return;
  const base = Number(parent.PrezzoAcquisto) || 0, n = kids.length;
  const q = Math.floor(base * 100 / n) / 100;
  kids.forEach(function (k, i) {
    const v = i < n - 1 ? q : r2_(base - q * (n - 1));
    if (Number(k.PrezzoAcquisto) !== v) patch_('Prodotti', k.ID, { PrezzoAcquisto: v }, user);
  });
}

function apri_(p, user) {
  const parent = get_('Prodotti', p.id);
  if (!parent || !isAct_(parent)) err_('Prodotto non trovato');
  if (!isSealed_(parent.Categoria)) err_('Si possono aprire solo i prodotti sigillati');
  const data = date_(p.data) || parent.DataApertura;
  if (!data) err_('Inserisci la data di apertura');
  if (data < parent.DataAcquisto) err_('La data di apertura è precedente all\'acquisto');
  const carte = (p.carte || []).filter(function (c) { return str_(c.Nome); });
  if (parent.Aperto === true && !carte.length) err_('Aggiungi almeno una carta');
  apriSeServe_(parent, data, user);
  let tot = 0;
  carte.forEach(function (c) {
    const q = parseInt(c.Quantita || 1, 10);
    if (!(q >= 1 && q <= 200)) err_('Quantità non valida per "' + c.Nome + '"');
    const cat = str_(c.Categoria) || 'Carta singola';
    const cond = str_(c.Condizione);
    if (cond && CONDIZIONI.indexOf(cond) < 0) err_('Condizione non valida');
    const lotto = newId_('T');
    for (let i = 0; i < q; i++) {
      put_('Prodotti', { ID: newId_('P'), LottoID: lotto, Origine: 'SBUSTATA', ParentID: parent.ID, DataAcquisto: data,
        Nome: str_(c.Nome), Categoria: cat, Condizione: isSealed_(cat) ? '' : cond, Lingua: str_(c.Lingua),
        PrezzoAcquisto: 0, PagatoDa: '', CanaleAcquisto: '', Aperto: false, FotoId: str_(c.FotoId), Note: str_(c.Note) }, user);
      tot++;
    }
  });
  ricalcolaQuote_(parent.ID, user);
  log_(user, 'Apertura', 'Prodotti', parent.ID, parent.Nome + (tot ? ' · ' + tot + ' carte' : ''));
  return { id: parent.ID };
}

function annullaApertura_(p, user) {
  const parent = get_('Prodotti', p.id);
  if (!parent || parent.Aperto !== true) err_('Il prodotto non risulta aperto');
  if (figli_(parent.ID).length) err_('Elimina prima le carte trovate in questo prodotto');
  put_('Prodotti', { ID: parent.ID, Aperto: false, DataApertura: '' }, user);
  log_(user, 'Apertura annullata', 'Prodotti', parent.ID, parent.Nome);
  return { id: parent.ID };
}

/* ------------------------------------------------------------------ */
/* Gradazione                                                          */
/* ------------------------------------------------------------------ */
function gradazione_(p, user) {
  const x = get_('Prodotti', p.id);
  if (!x || !isAct_(x)) err_('Prodotto non trovato');
  if (isSealed_(x.Categoria)) err_('La gradazione si applica alle carte, non ai prodotti sigillati');
  const stato = str_(p.GradStato);
  if (!stato) {
    put_('Prodotti', { ID: x.ID, GradStato: '', GradCasa: '', GradVoto: '', GradCert: '', GradDataInvio: '',
      GradDataRientro: '', GradCosto: '', GradPagatoDa: '' }, user);
    syncGrading_(x, user);
    log_(user, 'Gradazione rimossa', 'Prodotti', x.ID, x.Nome);
    return { id: x.ID };
  }
  if (GRAD_STATI.indexOf(stato) < 0) err_('Stato della gradazione non valido');
  if (x.InArrivo === true && stato !== 'DA_INVIARE') err_('"' + x.Nome + '" non è ancora arrivato: segnalo come arrivato prima di inviarlo a gradare');
  const casa = str_(p.GradCasa);
  if (!casa) err_('Seleziona la casa di gradazione');
  const voto = str_(p.GradVoto);
  if (stato === 'GRADATA' && !voto) err_('Inserisci il voto ottenuto');
  const costo = num_(p.GradCosto, 'Costo della gradazione');
  if (costo !== '' && costo < 0) err_('Costo non valido');
  const dInvio = date_(p.GradDataInvio), dRientro = date_(p.GradDataRientro);
  const pag = str_(p.GradPagatoDa);
  if (costo > 0) {
    if (!dInvio) err_('Inserisci la data di invio (è la data del costo)');
    if (pagatori_().indexOf(pag) < 0) err_('Indica chi ha pagato la gradazione');
  }
  if (dInvio && dRientro && dRientro < dInvio) err_('La data di rientro è precedente all\'invio');
  put_('Prodotti', { ID: x.ID, GradStato: stato, GradCasa: casa, GradVoto: voto, GradCert: str_(p.GradCert),
    GradDataInvio: dInvio, GradDataRientro: dRientro, GradCosto: costo === '' ? '' : costo,
    GradPagatoDa: costo > 0 ? pag : '' }, user);
  syncGrading_(get_('Prodotti', x.ID), user, p.Allegati);
  log_(user, 'Gradazione', 'Prodotti', x.ID, x.Nome + ' · ' + casa + (voto ? ' ' + voto : ''));
  return { id: x.ID };
}

function syncGrading_(x, user, allegati) {
  const costo = (isAct_(x) && x.GradStato) ? (Number(x.GradCosto) || 0) : 0;
  if (!(costo > 0)) {
    autoMov_('GRADING', 'RifProdotto', x.ID, null, user);
    syncVersAuto_('G:' + x.ID, null, user);
    return;
  }
  const descr = 'Gradazione: ' + x.Nome + ' (' + x.GradCasa + ')';
  const fields = { Data: x.GradDataInvio, Importo: costo, Descrizione: descr, PagatoDa: x.GradPagatoDa, Socio: '' };
  if (allegati) fields.Allegati = allegatiJson_(allegati);
  autoMov_('GRADING', 'RifProdotto', x.ID, fields, user);
  syncVersAuto_('G:' + x.ID, { pagatoDa: x.GradPagatoDa, importo: costo, data: x.GradDataInvio, descr: descr }, user);
}

/* ------------------------------------------------------------------ */
/* Scambi                                                              */
/* ------------------------------------------------------------------ */
function scambio_(p, user) {
  const data = date_(p.Data);
  if (!data) err_('Inserisci la data dello scambio');
  const contro = str_(p.Controparte);
  if (!contro) err_('Indica con chi è avvenuto lo scambio');
  const dati = (p.Dati || []).filter(function (v, i, a) { return v && a.indexOf(v) === i; });
  if (!dati.length) err_('Seleziona almeno un prodotto dato');
  dati.forEach(function (id) {
    const x = get_('Prodotti', id);
    if (!x || !inMagazzino_(x)) err_('"' + (x ? x.Nome : id) + '" non è disponibile in magazzino');
    if (x.DataAcquisto > data) err_('"' + x.Nome + '" risulta acquistato dopo la data dello scambio');
  });
  const ric = (p.Ricevuti || []).filter(function (r) { return str_(r.Nome); });
  if (!ric.length) err_('Inserisci almeno un prodotto ricevuto');
  ric.forEach(function (r) {
    if (!str_(r.Categoria)) err_('Seleziona la categoria di "' + r.Nome + '"');
    const q = parseInt(r.Quantita || 1, 10);
    if (!(q >= 1 && q <= 200)) err_('Quantità non valida per "' + r.Nome + '"');
    const v = num_(r.Valore, 'Valore');
    if (v === '' || v < 0) err_('Inserisci il valore di "' + r.Nome + '" (anche 0)');
    const cond = str_(r.Condizione);
    if (cond && CONDIZIONI.indexOf(cond) < 0) err_('Condizione non valida');
  });
  const tipoC = str_(p.ConguaglioTipo) || 'NESSUNO';
  let cong = num_(p.Conguaglio, 'Conguaglio');
  cong = cong === '' ? 0 : cong;
  if (cong < 0) err_('Conguaglio non valido');
  let signed = 0, pag = '';
  if (tipoC === 'PAGATO' && cong > 0) {
    signed = cong;
    pag = str_(p.PagatoDa);
    if (pagatori_().indexOf(pag) < 0) err_('Indica chi ha pagato il conguaglio');
  } else if (tipoC === 'RICEVUTO' && cong > 0) {
    signed = -cong;
  }
  const id = newId_('S');
  const s = put_('Scambi', { ID: id, Data: data, Controparte: contro, Conguaglio: signed, PagatoDa: pag, Note: str_(p.Note) }, user);
  dati.forEach(function (pid) { put_('Prodotti', { ID: pid, ScambioOutID: id }, user); });
  ric.forEach(function (r) {
    const q = parseInt(r.Quantita || 1, 10), cat = str_(r.Categoria);
    const lotto = newId_('T');
    for (let i = 0; i < q; i++) {
      put_('Prodotti', { ID: newId_('P'), LottoID: lotto, Origine: 'SCAMBIO', ScambioInID: id, DataAcquisto: data,
        Nome: str_(r.Nome), Categoria: cat, Condizione: isSealed_(cat) ? '' : str_(r.Condizione), Lingua: str_(r.Lingua),
        CanaleAcquisto: contro, PrezzoAcquisto: num_(r.Valore), PagatoDa: '', Aperto: false, FotoId: str_(r.FotoId) }, user);
    }
  });
  syncScambio_(s, user, p.Allegati);
  ensureCanale_(contro, 'ENTRAMBI', user);
  log_(user, 'Scambio', 'Scambi', id, 'con ' + contro + ' · dati ' + dati.length + ' · ricevuti ' + ric.length);
  return { id: id };
}

function syncScambio_(s, user, allegati) {
  const c = isAct_(s) ? (Number(s.Conguaglio) || 0) : 0;
  const tipo = c > 0 ? 'CONGUAGLIO_OUT' : 'CONGUAGLIO_IN';
  const altro = c > 0 ? 'CONGUAGLIO_IN' : 'CONGUAGLIO_OUT';
  autoMov_(altro, 'RifScambio', s.ID, null, user);
  if (!c) {
    autoMov_(tipo, 'RifScambio', s.ID, null, user);
    syncVersAuto_('S:' + s.ID, null, user);
    return;
  }
  const descr = 'Conguaglio scambio con ' + s.Controparte;
  const fields = { Data: s.Data, Importo: Math.abs(c), Descrizione: descr, PagatoDa: c > 0 ? s.PagatoDa : '', Socio: '' };
  if (allegati) fields.Allegati = allegatiJson_(allegati);
  autoMov_(tipo, 'RifScambio', s.ID, fields, user);
  syncVersAuto_('S:' + s.ID, c > 0 ? { pagatoDa: s.PagatoDa, importo: c, data: s.Data, descr: descr } : null, user);
}

function annullaScambio_(p, user) {
  const s = get_('Scambi', p.id);
  if (!s || !isAct_(s)) err_('Scambio non trovato');
  const ricevuti = rows_('Prodotti').filter(function (x) { return isAct_(x) && x.ScambioInID === s.ID; });
  ricevuti.forEach(function (x) {
    if (!inMagazzino_(x)) err_('"' + x.Nome + '" ricevuto nello scambio è già stato venduto, aperto o scambiato');
  });
  ricevuti.forEach(function (x) {
    put_('Prodotti', { ID: x.ID, Archiviato: true }, user);
    syncGrading_(get_('Prodotti', x.ID), user);
  });
  rows_('Prodotti').forEach(function (x) {
    if (x.ScambioOutID === s.ID) put_('Prodotti', { ID: x.ID, ScambioOutID: '' }, user);
  });
  put_('Scambi', { ID: s.ID, Archiviato: true }, user);
  syncScambio_(get_('Scambi', s.ID), user);
  log_(user, 'Scambio annullato', 'Scambi', s.ID, 'con ' + s.Controparte);
  return { id: s.ID };
}

/* ------------------------------------------------------------------ */
/* Prima nota (movimenti manuali e allegati)                           */
/* ------------------------------------------------------------------ */
function saveMovimento_(p, user) {
  const tipo = str_(p.Tipo);
  if (TIPI_MANUALI.indexOf(tipo) < 0) err_('Tipo di movimento non valido');
  if (p.ID) {
    const ex = get_('PrimaNota', p.ID);
    if (!ex) err_('Movimento non trovato');
    if (ex.Auto === true) err_('Questo movimento è automatico: modificalo dal prodotto, dallo scambio o dal prestito collegato');
  }
  const data = date_(p.Data);
  if (!data) err_('Inserisci la data');
  const importo = num_(p.Importo, 'Importo');
  if (importo === '' || importo <= 0) err_('Inserisci un importo maggiore di zero');
  let socio = '', pag = '';
  if (tipo === 'VERSAMENTO' || tipo === 'PRELIEVO') {
    socio = str_(p.Socio);
    if (['TUTTI'].concat(CONFIG.SOCI).indexOf(socio) < 0) err_('Seleziona il socio');
  }
  if (tipo === 'SPESA') {
    pag = str_(p.PagatoDa);
    if (pagatori_().indexOf(pag) < 0) err_('Indica chi ha pagato la spesa');
  }
  const rec = {
    ID: p.ID || newId_('M'), Data: data, Tipo: tipo, Descrizione: str_(p.Descrizione), Importo: importo,
    Socio: socio, PagatoDa: pag, Categoria: (tipo === 'SPESA' || tipo === 'ENTRATA') ? str_(p.Categoria) : '',
    RifProdotto: '', RifPrestito: '', RifLotto: '', RifOrigine: '', RifScambio: '',
    NumRata: '', QuotaCapitale: '', QuotaInteressi: '', Auto: false
  };
  if (p.Allegati) rec.Allegati = allegatiJson_(p.Allegati);
  const isNew = !p.ID;
  put_('PrimaNota', rec, user);
  syncVersAuto_('M:' + rec.ID, tipo === 'SPESA'
    ? { pagatoDa: pag, importo: importo, data: data, descr: 'Spesa: ' + (rec.Descrizione || rec.Categoria || '') } : null, user);
  log_(user, (isNew ? 'Nuovo movimento ' : 'Modifica movimento ') + tipo, 'PrimaNota', rec.ID, rec.Descrizione + ' € ' + importo);
  return { id: rec.ID };
}

function setAllegati_(p, user) {
  const m = get_('PrimaNota', p.movId);
  if (!m) err_('Movimento non trovato');
  put_('PrimaNota', { ID: m.ID, Allegati: allegatiJson_(p.allegati || []) }, user);
  log_(user, 'Allegati aggiornati', 'PrimaNota', m.ID, m.Descrizione);
  return { id: m.ID };
}

function allegatiJson_(list) {
  const clean = (list || []).filter(function (a) { return a && a.id; }).slice(0, 20).map(function (a) {
    return { id: String(a.id), nome: String(a.nome || 'allegato').slice(0, 80), tipo: String(a.tipo || '').slice(0, 40) };
  });
  return clean.length ? JSON.stringify(clean) : '';
}

/** Movimento automatico unico per (tipo, riferimento): crea, aggiorna o archivia. */
function autoMov_(tipo, refKey, refId, fields, user) {
  const movs = rows_('PrimaNota').filter(function (m) {
    return isAct_(m) && m.Auto === true && m.Tipo === tipo && m[refKey] === refId;
  });
  if (!fields) {
    movs.forEach(function (m) { put_('PrimaNota', { ID: m.ID, Archiviato: true }, user); });
    return null;
  }
  movs.slice(1).forEach(function (m) { put_('PrimaNota', { ID: m.ID, Archiviato: true }, user); });
  const rec = Object.assign({ Tipo: tipo, Auto: true }, fields);
  rec[refKey] = refId;
  if (movs[0]) {
    const o = movs[0];
    const changed = Object.keys(rec).some(function (k) { return String(o[k]) !== String(rec[k]); });
    if (!changed) return o;
    rec.ID = o.ID;
  } else {
    rec.ID = newId_('M');
  }
  return put_('PrimaNota', rec, user);
}

/** Quando paga un socio (o tutti e 3) di tasca propria, registra il versamento corrispondente. */
function syncVersAuto_(key, info, user) {
  const need = info && info.importo > 0 && info.pagatoDa && info.pagatoDa !== 'CASSA';
  autoMov_('VERSAMENTO', 'RifOrigine', key, need ? {
    Data: info.data, Importo: r2_(info.importo), Socio: info.pagatoDa, PagatoDa: '',
    Descrizione: 'Pagato di tasca propria · ' + info.descr
  } : null, user);
}

/* ------------------------------------------------------------------ */
/* Prestiti                                                            */
/* ------------------------------------------------------------------ */
function savePrestito_(p, user) {
  const importo = num_(p.Importo, 'Importo');
  if (importo === '' || importo <= 0) err_('Importo del prestito non valido');
  const tasso = num_(p.TassoAnnuo, 'Tasso');
  if (tasso === '' || tasso < 0) err_('Tasso annuo non valido (usa 0 se senza interessi)');
  const n = parseInt(p.NumRate, 10);
  if (!(n > 0 && n <= 600)) err_('Numero di rate non valido');
  const dataE = date_(p.DataErogazione);
  if (!dataE) err_('Inserisci la data di erogazione');
  const dataR = date_(p.DataPrimaRata);
  if (!dataR) err_('Inserisci la data della prima rata');
  const rataFissa = num_(p.RataFissa, 'Rata');
  if (p.ID) {
    const ex = get_('Prestiti', p.ID);
    if (!ex) err_('Prestito non trovato');
    if (rateAttive_(p.ID).length && (Number(ex.Importo) !== importo || Number(ex.TassoAnnuo) !== tasso ||
        Number(ex.NumRate) !== n || String(ex.RataFissa) !== String(rataFissa) || ex.DataPrimaRata !== dataR)) {
      err_('Ci sono già rate pagate: per cambiare importo, tasso, rate o date annulla prima i pagamenti delle rate');
    }
  }
  const rec = { ID: p.ID || newId_('L'), Descrizione: str_(p.Descrizione) || 'Prestito', Banca: str_(p.Banca),
    Importo: importo, TassoAnnuo: tasso, NumRate: n, RataFissa: rataFissa, DataErogazione: dataE,
    DataPrimaRata: dataR, Note: str_(p.Note) };
  const isNew = !p.ID;
  put_('Prestiti', rec, user);
  syncErogazione_(rec, user);
  log_(user, isNew ? 'Nuovo prestito' : 'Modifica prestito', 'Prestiti', rec.ID, rec.Descrizione + ' € ' + importo);
  return { id: rec.ID };
}

function syncErogazione_(l, user) {
  autoMov_('EROGAZIONE', 'RifPrestito', l.ID, isAct_(get_('Prestiti', l.ID)) ? {
    Data: l.DataErogazione, Importo: Number(l.Importo), Socio: '',
    Descrizione: 'Erogazione prestito: ' + l.Descrizione + (l.Banca ? ' (' + l.Banca + ')' : '')
  } : null, user);
}

/** Piano di ammortamento alla francese (rata costante, tasso fisso). */
function piano_(l) {
  const P = Number(l.Importo), n = parseInt(l.NumRate, 10), i = Number(l.TassoAnnuo) / 100 / 12;
  let rata = (l.RataFissa !== '' && l.RataFissa != null && Number(l.RataFissa) > 0)
    ? Number(l.RataFissa)
    : (i === 0 ? P / n : P * i / (1 - Math.pow(1 + i, -n)));
  rata = r2_(rata);
  const out = [];
  let res = P;
  for (let k = 1; k <= n; k++) {
    const interessi = r2_(res * i);
    let capitale = k === n ? r2_(res) : r2_(rata - interessi);
    if (capitale > res) capitale = r2_(res);
    res = r2_(res - capitale);
    out.push({ n: k, scadenza: addMonths_(l.DataPrimaRata, k - 1), rata: r2_(capitale + interessi),
      capitale: capitale, interessi: interessi, residuo: res });
    if (res <= 0) break;
  }
  return out;
}

function pagaRata_(p, user) {
  const l = get_('Prestiti', p.prestitoId);
  if (!l || !isAct_(l)) err_('Prestito non trovato');
  const n = parseInt(p.n, 10);
  const r = piano_(l).filter(function (x) { return x.n === n; })[0];
  if (!r) err_('Rata inesistente');
  if (rateAttive_(l.ID).some(function (m) { return Number(m.NumRata) === n; })) err_('Rata già segnata come pagata');
  const rec = { ID: newId_('M'), Data: date_(p.data) || r.scadenza, Tipo: 'RATA',
    Descrizione: 'Rata ' + n + '/' + l.NumRate + ' ' + l.Descrizione + (l.Banca ? ' (' + l.Banca + ')' : ''),
    Importo: r.rata, Socio: '', Categoria: '', RifPrestito: l.ID, NumRata: n,
    QuotaCapitale: r.capitale, QuotaInteressi: r.interessi, Auto: true };
  put_('PrimaNota', rec, user);
  log_(user, 'Rata pagata', 'PrimaNota', rec.ID, rec.Descrizione + ' € ' + rec.Importo);
  return { id: rec.ID };
}

function rateAttive_(prestitoId) {
  return rows_('PrimaNota').filter(function (m) {
    return m.Tipo === 'RATA' && m.RifPrestito === prestitoId && isAct_(m);
  });
}

/* ------------------------------------------------------------------ */
/* Anagrafica canali                                                   */
/* ------------------------------------------------------------------ */
function saveCanale_(p, user) {
  const nome = str_(p.Nome);
  if (!nome) err_('Inserisci il nome del canale');
  const tipo = str_(p.Tipo) || 'Altro';
  if (TIPI_CANALE.indexOf(tipo) < 0) err_('Tipo di canale non valido');
  const uso = str_(p.Uso) || 'ENTRAMBI';
  if (USI_CANALE.indexOf(uso) < 0) err_('Uso del canale non valido');
  const dup = rows_('Canali').filter(function (c) { return c.ID !== p.ID && String(c.Nome).toLowerCase() === nome.toLowerCase(); })[0];
  if (dup) err_('Esiste già un canale chiamato "' + dup.Nome + '"' + (isAct_(dup) ? '' : ' (disattivato: riattivalo)'));
  let nomePrima = '';
  if (p.ID) {
    const old = get_('Canali', p.ID);
    if (!old) err_('Canale non trovato');
    nomePrima = old.Nome;
  }
  const c = put_('Canali', { ID: p.ID || newId_('C'), Nome: nome, Tipo: tipo, Uso: uso, Note: str_(p.Note), IncassoDifferito: p.IncassoDifferito === true }, user);
  if (nomePrima && nomePrima !== nome) rinominaCanale_(nomePrima, nome, user);
  log_(user, p.ID ? 'Modifica canale' : 'Nuovo canale', 'Canali', c.ID, nome);
  return { id: c.ID };
}

function rinominaCanale_(vecchio, nuovo, user) {
  rows_('Prodotti').forEach(function (x) {
    const f = {};
    if (x.CanaleAcquisto === vecchio) f.CanaleAcquisto = nuovo;
    if (x.CanaleVendita === vecchio) f.CanaleVendita = nuovo;
    if (Object.keys(f).length) patch_('Prodotti', x.ID, f, user);
  });
  rows_('Scambi').forEach(function (s) { if (s.Controparte === vecchio) patch_('Scambi', s.ID, { Controparte: nuovo }, user); });
}

function ensureCanale_(nome, uso, user) {
  if (!nome) return;
  const c = rows_('Canali').filter(function (x) { return String(x.Nome).toLowerCase() === String(nome).toLowerCase(); })[0];
  if (!c) { put_('Canali', { ID: newId_('C'), Nome: nome, Tipo: 'Altro', Uso: uso || 'ENTRAMBI', Note: '' }, user); return; }
  const f = {};
  if (!isAct_(c)) f.Archiviato = false;
  if (uso && c.Uso !== uso && c.Uso !== 'ENTRAMBI') f.Uso = 'ENTRAMBI';
  if (Object.keys(f).length) put_('Canali', Object.assign({ ID: c.ID }, f), user);
}

/* ------------------------------------------------------------------ */
/* Operazioni su più prodotti insieme (lotti, selezioni)               */
/* ------------------------------------------------------------------ */
function pezziSelezionati_(ids) {
  const list = (ids || []).filter(function (v, i, a) { return v && a.indexOf(v) === i; });
  if (!list.length) err_('Nessun prodotto selezionato');
  if (list.length > 500) err_('Troppi prodotti selezionati');
  return list.map(function (id) {
    const x = get_('Prodotti', id);
    if (!x || !isAct_(x)) err_('Uno dei prodotti selezionati non esiste più: aggiorna la pagina');
    return x;
  });
}

function dividi_(tot, n) {
  const unit = Math.floor(tot * 100 / n) / 100, out = [];
  for (let i = 0; i < n; i++) out.push(i < n - 1 ? unit : r2_(tot - unit * (n - 1)));
  return out;
}

/** Modifica gli stessi campi su più pezzi. Si cambiano solo i campi presenti in p.fields. */
function modificaMultipla_(p, user) {
  const pezzi = pezziSelezionati_(p.ids);
  const f = p.fields || {};
  const has = function (k) { return Object.prototype.hasOwnProperty.call(f, k); };
  const comuni = {}, acquisto = {};
  if (has('Nome')) { comuni.Nome = str_(f.Nome); if (!comuni.Nome) err_('Il nome non può essere vuoto'); }
  if (has('Categoria')) {
    comuni.Categoria = str_(f.Categoria);
    if (!comuni.Categoria) err_('La categoria non può essere vuota');
    pezzi.forEach(function (x) { if (x.Aperto === true && !isSealed_(comuni.Categoria)) err_('"' + x.Nome + '" è stato aperto: deve restare un prodotto sigillato'); });
  }
  if (has('Condizione')) { comuni.Condizione = str_(f.Condizione); if (comuni.Condizione && CONDIZIONI.indexOf(comuni.Condizione) < 0) err_('Condizione non valida'); }
  if (has('Lingua')) comuni.Lingua = str_(f.Lingua);
  if (has('Note')) comuni.Note = str_(f.Note);
  if (has('FotoId')) comuni.FotoId = str_(f.FotoId);
  if (has('DataAcquisto')) { acquisto.DataAcquisto = date_(f.DataAcquisto); if (!acquisto.DataAcquisto) err_('La data di acquisto non può essere vuota'); }
  if (has('CanaleAcquisto')) { acquisto.CanaleAcquisto = str_(f.CanaleAcquisto); if (!acquisto.CanaleAcquisto) err_('Il canale di acquisto non può essere vuoto'); }
  if (has('PagatoDa')) { acquisto.PagatoDa = str_(f.PagatoDa); if (pagatori_().indexOf(acquisto.PagatoDa) < 0) err_('Indica chi ha pagato'); }
  const acqPezzi = pezzi.filter(function (x) { return (x.Origine || 'ACQUISTO') === 'ACQUISTO'; });
  let prezzi = null;
  const pr = p.prezzo || {};
  if (pr.modo === 'unit' || pr.modo === 'totale') {
    const v = num_(pr.valore, 'Prezzo');
    if (v === '' || v < 0) err_('Inserisci un prezzo valido (anche 0)');
    if (!acqPezzi.length) err_('Il prezzo si può cambiare solo sui pezzi acquistati (le carte sbustate prendono il costo dal box)');
    prezzi = pr.modo === 'totale' ? dividi_(v, acqPezzi.length) : acqPezzi.map(function () { return v; });
  }
  if (!Object.keys(comuni).length && !Object.keys(acquisto).length && !prezzi && !(p.allegati && p.allegati.length)) err_('Non hai cambiato nessun campo');
  if (acquisto.DataAcquisto) acqPezzi.forEach(function (x) {
    if (isSold_(x) && x.DataVendita < acquisto.DataAcquisto) err_('"' + x.Nome + '" risulta venduto prima della nuova data di acquisto');
  });
  const lotti = {}, parents = {};
  pezzi.forEach(function (x) {
    const upd = Object.assign({ ID: x.ID }, comuni);
    if (comuni.Categoria && isSealed_(comuni.Categoria)) upd.Condizione = '';
    if ((x.Origine || 'ACQUISTO') === 'ACQUISTO') {
      Object.assign(upd, acquisto);
      const i = acqPezzi.indexOf(x);
      if (prezzi && i >= 0) upd.PrezzoAcquisto = prezzi[i];
      if (x.LottoID) lotti[x.LottoID] = 1;
    }
    if (x.ParentID) parents[x.ParentID] = 1;
    put_('Prodotti', upd, user);
  });
  Object.keys(lotti).forEach(function (l) {
    let al = null;
    if (p.allegati && p.allegati.length) {
      const m = rows_('PrimaNota').filter(function (mm) { return isAct_(mm) && mm.Auto === true && mm.Tipo === 'ACQUISTO' && mm.RifLotto === l; })[0];
      let prima = [];
      try { prima = m && m.Allegati ? JSON.parse(m.Allegati) : []; } catch (e) { prima = []; }
      al = prima.concat(p.allegati);
    }
    syncLotto_(l, user, al);
  });
  pezzi.forEach(function (x) {
    const y = get_('Prodotti', x.ID);
    if (y.Aperto === true) ricalcolaQuote_(y.ID, user);
    syncVendita_(y, user);
    if (y.GradStato) syncGrading_(y, user);
  });
  if (acquisto.CanaleAcquisto) ensureCanale_(acquisto.CanaleAcquisto, 'ACQUISTO', user);
  log_(user, 'Modifica di ' + pezzi.length + ' pezzi', 'Prodotti', pezzi[0].ID, pezzi[0].Nome + (pezzi.length > 1 ? ' e altri' : ''));
  return { n: pezzi.length };
}

/** Vende più pezzi insieme: stessa data, canale e incasso; prezzo per pezzo o totale diviso. */
function vendiMultipla_(p, user) {
  const pezzi = pezziSelezionati_(p.ids);
  pezzi.forEach(function (x) { if (!inMagazzino_(x)) err_('"' + x.Nome + '" non è in magazzino (venduto, aperto o scambiato)'); });
  const dv = date_(p.DataVendita);
  if (!dv) err_('Inserisci la data di vendita');
  const cv = str_(p.CanaleVendita);
  if (!cv) err_('Seleziona il canale di vendita');
  const inc = str_(p.IncassatoDa);
  if (incassatori_().indexOf(inc) < 0) err_('Indica chi ha incassato la vendita');
  const v = num_(p.valore, 'Prezzo di vendita');
  if (v === '' || v < 0) err_('Inserisci il prezzo di vendita');
  pezzi.forEach(function (x) { if (dv < x.DataAcquisto) err_('"' + x.Nome + '" risulta acquistato dopo la data di vendita'); });
  const prezzi = p.modo === 'totale' ? dividi_(v, pezzi.length) : pezzi.map(function () { return v; });
  const inc2 = statoIncasso_({ DaIncassare: p.DaIncassare === true, DataIncasso: p.DataIncasso, ImportoIncassato: p.ImportoIncassato }, prezzi, dv);
  pezzi.forEach(function (x, i) {
    put_('Prodotti', Object.assign({ ID: x.ID, DataVendita: dv, CanaleVendita: cv, PrezzoVendita: prezzi[i], IncassatoDa: inc }, inc2[i]), user);
    syncVendita_(get_('Prodotti', x.ID), user);
  });
  ensureCanale_(cv, 'VENDITA', user);
  log_(user, 'Vendita di ' + pezzi.length + ' pezzi', 'Prodotti', pezzi[0].ID, pezzi[0].Nome + ' · € ' + r2_(prezzi.reduce(function (a, b) { return a + b; }, 0)));
  return { n: pezzi.length };
}

function archiviaMultipla_(p, user) {
  const pezzi = pezziSelezionati_(p.ids);
  const ids = pezzi.map(function (x) { return x.ID; });
  // prima le carte sbustate, poi i sigillati da cui provengono
  pezzi.sort(function (a, b) { return (b.ParentID && ids.indexOf(b.ParentID) >= 0 ? 1 : 0) - (a.ParentID && ids.indexOf(a.ParentID) >= 0 ? 1 : 0); });
  pezzi.forEach(function (x) { setArchiviato_('Prodotti', x.ID, true, user, true); });
  log_(user, 'Eliminati ' + pezzi.length + ' pezzi', 'Prodotti', pezzi[0].ID, pezzi[0].Nome + (pezzi.length > 1 ? ' e altri' : ''));
  return { n: pezzi.length };
}

/* ------------------------------------------------------------------ */
/* Accesso: PIN personale per socio (foglio "Soci")                    */
/* ------------------------------------------------------------------ */
function socio_(nome) {
  return rows_('Soci').filter(function (s) { return s.Socio === nome; })[0] || null;
}

function aggiornaSocio_(s, fields) {
  Object.keys(fields).forEach(function (k) { s[k] = fields[k]; });
  tbl_('Soci').dirty[s._row] = s;
}

function auth_(req) {
  const nome = String(req.user || '');
  if (CONFIG.SOCI.indexOf(nome) < 0) return { ok: false, code: 'PIN', error: 'Seleziona chi sei' };
  const s = socio_(nome);
  if (!s) return { ok: false, error: 'Socio non configurato: esegui setup dall\'editor di Apps Script' };
  const now = nowStr_();
  if (s.BloccatoFino && String(s.BloccatoFino) > now) {
    return { ok: false, code: 'BLOCCATO', error: 'Troppi PIN errati: accesso bloccato fino alle ' + String(s.BloccatoFino).slice(11, 16) };
  }
  if (String(req.pin || '') !== String(s.PIN)) {
    const n = (Number(s.TentativiErrati) || 0) + 1;
    if (n >= 5) {
      const fino = Utilities.formatDate(new Date(Date.now() + 15 * 60 * 1000), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
      aggiornaSocio_(s, { TentativiErrati: 0, BloccatoFino: fino });
      log_(nome, 'Accesso bloccato per 15 minuti', 'Soci', '', '5 PIN errati');
      flush_();
      return { ok: false, code: 'BLOCCATO', error: 'Troppi PIN errati: accesso bloccato per 15 minuti' };
    }
    aggiornaSocio_(s, { TentativiErrati: n });
    flush_();
    return { ok: false, code: 'PIN', error: 'PIN errato (tentativi rimasti: ' + (5 - n) + ')' };
  }
  if (Number(s.TentativiErrati) || s.BloccatoFino) { aggiornaSocio_(s, { TentativiErrati: 0, BloccatoFino: '' }); flush_(); }
  return { ok: true, user: nome, daCambiare: s.DaCambiare === true };
}

function cambiaPin_(user, vecchio, p) {
  const s = socio_(user);
  const nuovo = String(p.nuovo || '').trim();
  if (!/^\d{4,8}$/.test(nuovo)) err_('Il nuovo PIN deve avere da 4 a 8 cifre');
  if (nuovo === String(vecchio)) err_('Il nuovo PIN deve essere diverso da quello attuale');
  if (nuovo === String(CONFIG.PIN_INIZIALE)) err_('Scegli un PIN diverso da quello iniziale');
  aggiornaSocio_(s, { PIN: nuovo, DaCambiare: false, UltimaModifica: nowStr_(), TentativiErrati: 0, BloccatoFino: '' });
  log_(user, 'PIN personale cambiato', 'Soci', '', '');
}

function seedSoci_() {
  const t = tbl_('Soci');
  if (!CONFIG.PIN_INIZIALE && CONFIG.SOCI.some(function (n) { return !socio_(n); })) err_('Imposta CONFIG.PIN_INIZIALE nell\'editor prima di eseguire setup');
  CONFIG.SOCI.forEach(function (nome) {
    if (socio_(nome)) return;
    const o = { Socio: nome, PIN: String(CONFIG.PIN_INIZIALE), DaCambiare: true, UltimaModifica: nowStr_(), TentativiErrati: 0, BloccatoFino: '' };
    t.rows.push(o); t.added.push(o);
  });
  flush_();
}

/* ------------------------------------------------------------------ */
/* Archiviazione (niente viene mai cancellato davvero)                 */
/* ------------------------------------------------------------------ */
function setArchiviato_(foglio, id, value, user, silent) {
  if (['Prodotti', 'PrimaNota', 'Prestiti', 'Canali'].indexOf(foglio) < 0) err_('Operazione non valida');
  const rec = get_(foglio, id);
  if (!rec) err_('Elemento non trovato');

  if (foglio === 'Prodotti') {
    if (value) {
      if (figli_(id).length) err_('Elimina prima le carte trovate in questo prodotto');
      if (rec.ScambioOutID || rec.ScambioInID) err_('Il prodotto fa parte di uno scambio: annulla lo scambio dal menu Scambi');
    } else {
      if (rec.ParentID && !isAct_(get_('Prodotti', rec.ParentID))) err_('Ripristina prima il prodotto sigillato da cui proviene');
      if (rec.ScambioInID && !isAct_(get_('Scambi', rec.ScambioInID))) err_('Lo scambio di provenienza è stato annullato');
    }
    put_('Prodotti', { ID: id, Archiviato: value }, user);
    const x = get_('Prodotti', id);
    if ((x.Origine || 'ACQUISTO') === 'ACQUISTO') syncLotto_(x.LottoID, user);
    if (x.Origine === 'SBUSTATA') ricalcolaQuote_(x.ParentID, user);
    syncVendita_(x, user);
    syncGrading_(x, user);
  } else if (foglio === 'PrimaNota') {
    if (rec.Auto === true && rec.Tipo !== 'RATA') {
      err_('Movimento automatico: si gestisce dal prodotto, dallo scambio o dal prestito collegato');
    }
    if (!value && rec.Tipo === 'RATA' && rateAttive_(rec.RifPrestito).some(function (m) { return Number(m.NumRata) === Number(rec.NumRata); })) {
      err_('Questa rata risulta già pagata');
    }
    put_('PrimaNota', { ID: id, Archiviato: value }, user);
    if (rec.Auto !== true) {
      const m = get_('PrimaNota', id);
      syncVersAuto_('M:' + id, (!value && m.Tipo === 'SPESA')
        ? { pagatoDa: m.PagatoDa, importo: Number(m.Importo), data: m.Data, descr: 'Spesa: ' + (m.Descrizione || m.Categoria || '') } : null, user);
    }
  } else if (foglio === 'Prestiti') {
    if (value && rateAttive_(id).length) err_('Ci sono rate pagate: annullale prima di eliminare il prestito');
    put_('Prestiti', { ID: id, Archiviato: value }, user);
    syncErogazione_(get_('Prestiti', id), user);
  } else {
    put_('Canali', { ID: id, Archiviato: value }, user);
  }
  if (!silent) log_(user, value ? 'Archiviato' : 'Ripristinato', foglio, id, rec.Nome || rec.Descrizione || '');
  return { id: id };
}

/* ------------------------------------------------------------------ */
/* File: foto prodotti (visibili con link) e ricevute (private)        */
/* ------------------------------------------------------------------ */
function uploadFoto_(p, user) {
  const m = String(p.dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) err_('Formato immagine non valido');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 8 * 1024 * 1024) err_('Immagine troppo grande');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const name = 'foto_' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss') + '.' + ext;
  const file = folder_('FOTO_FOLDER_ID', CONFIG.NOME_APP + ' - Foto').createFile(Utilities.newBlob(bytes, m[1], name));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* si userà il caricamento tramite script */ }
  return file.getId();
}

function uploadAllegato_(p, user) {
  const m = String(p.dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp)|application\/pdf);base64,(.+)$/);
  if (!m) err_('Sono accettate solo immagini (JPG, PNG) o PDF');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 15 * 1024 * 1024) err_('File troppo grande (massimo 15 MB)');
  const ext = m[1] === 'application/pdf' ? 'pdf' : m[1].split('/')[1].replace('jpeg', 'jpg');
  const base = str_(p.nome).replace(/[^\w.\- ]+/g, '').slice(0, 60) || 'ricevuta';
  const name = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmmss') + '_' + base.replace(/\.(pdf|jpe?g|png|webp)$/i, '') + '.' + ext;
  const file = folder_('RICEVUTE_FOLDER_ID', CONFIG.NOME_APP + ' - Ricevute').createFile(Utilities.newBlob(bytes, m[1], name));
  return { id: file.getId(), nome: name, tipo: m[1] };
}

function getFile_(id, propKey) {
  const folderId = PropertiesService.getScriptProperties().getProperty(propKey);
  const file = DriveApp.getFileById(String(id));
  const parents = file.getParents();
  let ok = false;
  while (parents.hasNext()) { if (parents.next().getId() === folderId) ok = true; }
  if (!ok) err_('File non accessibile');
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/* ------------------------------------------------------------------ */
/* Setup, migrazione, backup                                           */
/* ------------------------------------------------------------------ */
function setup() {
  DB_ = {};
  const ss = ss_();
  try { ss.setSpreadsheetTimeZone(CONFIG.TZ); } catch (e) { /* ignorato */ }
  Object.keys(SCHEMA).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SCHEMA[name];
    const current = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    const merged = current.filter(function (h) { return h; });
    headers.forEach(function (h) { if (merged.indexOf(h) < 0) merged.push(h); });
    sh.getRange(1, 1, 1, merged.length).setValues([merged])
      .setFontWeight('bold').setBackground('#1b1f27').setFontColor('#e8b04a');
    sh.setFrozenRows(1);
    applyFormats_(sh, name);
  });
  ss.getSheets().forEach(function (sh) {
    if (!SCHEMA[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  folder_('FOTO_FOLDER_ID', CONFIG.NOME_APP + ' - Foto');
  folder_('BACKUP_FOLDER_ID', CONFIG.NOME_APP + ' - Backup');
  folder_('RICEVUTE_FOLDER_ID', CONFIG.NOME_APP + ' - Ricevute');
  const has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'backupSettimanale'; });
  if (!has) ScriptApp.newTrigger('backupSettimanale').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();

  seedSoci_();
  const props = PropertiesService.getScriptProperties();
  let report = [];
  const vd = props.getProperty('VERSIONE_DATI') || '1';
  if (vd !== VERSIONE_DATI) {
    if (vd < '2') report = report.concat(migraV2_());
    if (vd < '3') report = report.concat(migraV3_());
    flush_();
    props.setProperty('VERSIONE_DATI', VERSIONE_DATI);
  }
  DB_ = {};
  bumpRev_();
  const msg = 'Setup completato' + (report.length ? '\n' + report.join('\n') : '');
  console.log(msg);
  return msg;
}

function migraV2_() {
  const U = 'Sistema';
  const report = [];
  const P = rows_('Prodotti');
  P.forEach(function (x) {
    if (isAct_(x) && (x.DataAcquisto === '2026-11-08' || x.DataAcquisto === '2026-11-09')) {
      report.push('DA VERIFICARE · data di acquisto ' + x.DataAcquisto + ': ' + x.Nome + ' [' + x.ID + ']');
    }
  });
  // Campi nuovi sui prodotti esistenti e raggruppamento in lotti dei pezzi identici inseriti uno alla volta
  const gruppi = {};
  P.forEach(function (x) {
    const f = {};
    if (!x.Origine) f.Origine = 'ACQUISTO';
    if (!x.PagatoDa) f.PagatoDa = 'CASSA';
    if (x.Aperto === '') f.Aperto = false;
    if (Object.keys(f).length) patch_('Prodotti', x.ID, f, U);
    if (!x.LottoID) {
      const k = [x.Nome, x.DataAcquisto, x.CanaleAcquisto, x.PrezzoAcquisto, x.PagatoDa, isAct_(x)].join('|');
      (gruppi[k] = gruppi[k] || []).push(x);
    }
  });
  let lottiUniti = 0;
  Object.keys(gruppi).forEach(function (k) {
    const lotto = newId_('T');
    gruppi[k].forEach(function (x) { patch_('Prodotti', x.ID, { LottoID: lotto }, U); });
    if (gruppi[k].length > 1) lottiUniti++;
  });
  rows_('PrimaNota').forEach(function (m) {
    if (m.Tipo === 'ACQUISTO' && m.Auto === true && !m.RifLotto && m.RifProdotto) {
      const x = get_('Prodotti', m.RifProdotto);
      if (x) patch_('PrimaNota', m.ID, { RifLotto: x.LottoID }, U);
    }
  });
  const lotti = {};
  P.forEach(function (x) { if (isAct_(x) && x.LottoID) lotti[x.LottoID] = 1; });
  Object.keys(lotti).forEach(function (l) { syncLotto_(l, U); });
  if (lottiUniti) report.push('Pezzi identici raggruppati in ' + lottiUniti + ' lotti (un solo movimento di acquisto per lotto)');
  // Capitale iniziale versato da tutti e 3 i soci
  rows_('PrimaNota').forEach(function (m) {
    if (isAct_(m) && m.Tipo === 'VERSAMENTO' && m.Auto !== true && /capitale iniziale/i.test(String(m.Descrizione)) && m.Socio !== 'TUTTI') {
      patch_('PrimaNota', m.ID, { Socio: 'TUTTI' }, U);
      report.push('Versamento "' + m.Descrizione + '" (' + m.Importo + ' €) ora attribuito a tutti e 3 i soci');
    }
  });
  // Anagrafica canali dai canali già usati
  P.forEach(function (x) {
    ensureCanale_(x.CanaleAcquisto, 'ACQUISTO', U);
    ensureCanale_(x.CanaleVendita, 'VENDITA', U);
  });
  // Riscrive tutte le righe: le date diventano testo AAAA-MM-GG (niente più problemi di fuso orario)
  ['Prodotti', 'PrimaNota', 'Prestiti'].forEach(function (n) {
    const t = tbl_(n);
    t.rows.forEach(function (o) { if (o._row) t.dirty[o._row] = o; });
  });
  log_(U, 'Aggiornamento dati alla versione 2', '', '', report.join(' | '));
  return report;
}

function backupSettimanale() {
  const ss = ss_();
  const folder = folder_('BACKUP_FOLDER_ID', CONFIG.NOME_APP + ' - Backup');
  const name = ss.getName() + ' - backup ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH.mm');
  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(CONFIG.BACKUP_DA_TENERE).forEach(function (f) { f.setTrashed(true); });
}

/* ------------------------------------------------------------------ */
/* Utilità                                                             */
/* ------------------------------------------------------------------ */
function ss_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) err_('Foglio "' + name + '" mancante: esegui la funzione setup dall\'editor di Apps Script');
  return sh;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) err_('Il sito è occupato da un altro salvataggio, riprova tra qualche secondo');
  try { return fn(); } finally { lock.releaseLock(); }
}

function getRev_() { return Number(PropertiesService.getScriptProperties().getProperty('REV') || 0); }
function bumpRev_() { PropertiesService.getScriptProperties().setProperty('REV', String(getRev_() + 1)); }

function folder_(propKey, name) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(propKey);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* ricrea */ } }
  const f = DriveApp.createFolder(name);
  props.setProperty(propKey, f.getId());
  return f;
}

let ID_SEQ_ = 0;
function newId_(prefix) {
  ID_SEQ_ = (ID_SEQ_ + 1) % 1296;
  return prefix + Date.now().toString(36).toUpperCase() + ('0' + ID_SEQ_.toString(36)).slice(-2).toUpperCase() +
    Math.random().toString(36).slice(2, 5).toUpperCase();
}
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
function today_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'); }
function str_(v) { return v == null ? '' : String(v).trim().slice(0, 500); }
function r2_(x) { return Math.round((Number(x) + Number.EPSILON) * 100) / 100; }
function num_(v, label) {
  if (v === '' || v == null) return '';
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  if (isNaN(n)) err_((label || 'Valore') + ' non è un numero valido');
  return r2_(n);
}
function date_(v) {
  if (!v) return '';
  const s = String(v).slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) err_('Data non valida: ' + v);
  return s;
}
function addMonths_(iso, k) {
  const y = Number(iso.slice(0, 4)), m = Number(iso.slice(5, 7)) - 1, d = Number(iso.slice(8, 10));
  const tm = m + k, ty = y + Math.floor(tm / 12), mm = ((tm % 12) + 12) % 12;
  const last = new Date(Date.UTC(ty, mm + 1, 0)).getUTCDate();
  const dd = Math.min(d, last);
  return ty + '-' + ('0' + (mm + 1)).slice(-2) + '-' + ('0' + dd).slice(-2);
}
/** Evita che un testo che inizia con = + - @ venga interpretato come formula. */
function safe_(v) {
  if (typeof v === 'string' && /^([=+@]|-(?!\d))/.test(v)) return "'" + v;
  return v;
}

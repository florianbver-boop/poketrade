/**
 * PokeTrade — liaison entre l'app et ce Google Sheet.
 * À coller dans : Extensions > Apps Script (remplace tout le contenu), puis Déployer > Nouveau déploiement > Application Web.
 * Les onglets Bilan, Cartes et Frais sont créés automatiquement au premier appel.
 */
const CARTES = ['ID', 'Achetée le', 'Type', 'Nom', 'Numéro', 'Série', 'Langue', 'Gradation', 'Note', 'Prix achat', 'Coût gradation', 'Prix vente', 'Vendue le', 'Bénéfice', 'Statut'];
const CK = ['id', 'buyDate', 'type', 'name', 'number', 'set', 'lang', 'gradeCompany', 'grade', 'buyPrice', 'gradeCost', 'sellPrice', 'sellDate'];
const ITEMS = ['ID', 'Acheté le', 'Catégorie', 'Nom', 'Série', 'Langue', 'Quantité', 'Prix achat', 'Prix vente', 'Vendu le', 'Bénéfice', 'Statut'];
const IK = ['id', 'buyDate', 'cat', 'name', 'set', 'lang', 'qty', 'buyPrice', 'sellPrice', 'sellDate'];
const BILAN_V = '3';
const FRAIS = ['ID', 'Nom', 'Catégorie', 'Montant', 'Fréquence', 'Début', 'Fin', 'Total payé'];
const FK = ['id', 'name', 'cat', 'amount', 'freq', 'start', 'end'];
const MONEY = '#,##0.00 "€"';
const PROFIT = '#,##0.00 "€";[Red]-#,##0.00 "€"';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function header_(sh, cols) {
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1a1d24').setFontColor('#ffcb05');
  sh.setFrozenRows(1);
  sh.hideColumns(1);
}

function setup() {
  const ss = ss_();
  if (!ss.getSheetByName('Cartes')) {
    const first = ss.getSheets()[0];
    const c = (ss.getSheets().length === 1 && first.getLastRow() === 0) ? first.setName('Cartes') : ss.insertSheet('Cartes');
    ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach(l => c.getRange(l + ':' + l).setNumberFormat('@'));
    c.getRange('J:L').setNumberFormat(MONEY);
    c.getRange('N:N').setNumberFormat(PROFIT);
    c.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    c.getRange('M:M').setNumberFormat('dd/mm/yyyy');
    header_(c, CARTES);
    c.setColumnWidth(4, 220);
  }
  if (!ss.getSheetByName('Frais')) {
    const f = ss.insertSheet('Frais');
    ['B', 'C', 'E'].forEach(l => f.getRange(l + ':' + l).setNumberFormat('@'));
    f.getRange('D:D').setNumberFormat(MONEY);
    f.getRange('H:H').setNumberFormat(MONEY);
    f.getRange('F:G').setNumberFormat('dd/mm/yyyy');
    header_(f, FRAIS);
    f.setColumnWidth(2, 200);
  }
  if (!ss.getSheetByName('Items')) {
    const it = ss.insertSheet('Items');
    ['C', 'D', 'E', 'F'].forEach(l => it.getRange(l + ':' + l).setNumberFormat('@'));
    it.getRange('H:I').setNumberFormat(MONEY);
    it.getRange('K:K').setNumberFormat(PROFIT);
    it.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    it.getRange('J:J').setNumberFormat('dd/mm/yyyy');
    header_(it, ITEMS);
    it.setColumnWidth(4, 220);
  }
  let b = ss.getSheetByName('Bilan');
  if (!b) b = ss.insertSheet('Bilan', 0);
  if (PropertiesService.getScriptProperties().getProperty('BILAN_V') !== BILAN_V) buildBilan_(b);
  repair_();
}

const BILAN = [
  ['Bilan PokeTrade', ''],
  ['', ''],
  ['Cartes achetées', '=COUNTA(Cartes!A2:A)'],
  ['Items achetés (lignes)', '=COUNTA(Items!A2:A)'],
  ['Objets vendus', '=COUNT(Cartes!L2:L)+COUNT(Items!I2:I)'],
  ['Total achats (cartes + items)', '=SUM(Cartes!J2:J)+SUM(Items!H2:H)'],
  ['Total gradation', '=SUM(Cartes!K2:K)'],
  ['Total ventes', '=SUM(Cartes!L2:L)+SUM(Items!I2:I)'],
  ['Coût des objets vendus', '=SUMIFS(Cartes!J2:J,Cartes!L2:L,"<>")+SUMIFS(Cartes!K2:K,Cartes!L2:L,"<>")+SUMIFS(Items!H2:H,Items!I2:I,"<>")'],
  ['Marge sur les ventes', '=B8-B9'],
  ['Frais annexes (bot, cookgroup, proxies…)', '=SUM(Frais!H2:H)'],
  ['BÉNÉFICE NET', '=B10-B11'],
  ['ROI net', '=IFERROR(B12/(B9+B11),0)'],
  ['Stock non vendu (coût)', '=B6+B7-B9'],
  ['Trésorerie (ventes − toutes les dépenses)', '=B8-B6-B7-B11'],
  ['', ''],
  ['Marge sur les cartes', '=SUM(Cartes!N2:N)'],
  ['Marge sur les items', '=SUM(Items!K2:K)']
];

function buildBilan_(b) {
  b.clear();
  b.getRange(1, 1, BILAN.length, 1).setValues(BILAN.map(r => [r[0]]));
  bilanFormulas_(b);
  b.getRange('A1').setFontSize(16).setFontWeight('bold');
  b.getRange('B6:B12').setNumberFormat(PROFIT);
  b.getRange('B14:B15').setNumberFormat(PROFIT);
  b.getRange('B17:B18').setNumberFormat(PROFIT);
  b.getRange('B13').setNumberFormat('0.0%');
  b.getRange('A12:B12').setFontWeight('bold').setFontSize(13).setBackground('#fff4c2');
  b.setColumnWidth(1, 330);
  b.setColumnWidth(2, 140);
  PropertiesService.getScriptProperties().setProperty('BILAN_V', BILAN_V);
}

/* Séparateur des formules : "," ou ";" selon la langue du Sheet (détecté une fois). */
function sep_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SEP');
  if (s) return s;
  const c = ss_().getSheetByName('Bilan').getRange('D1');
  c.setFormula('=IF(1=1,1,0)');
  SpreadsheetApp.flush();
  s = c.getDisplayValue() === '1' ? ',' : ';';
  c.clearContent();
  props.setProperty('SEP', s);
  return s;
}
function F_(range, f) { range.setFormula(sep_() === ',' ? f : f.replace(/,/g, ';')); }

function bilanFormulas_(b) {
  BILAN.forEach((r, i) => { if (r[1]) F_(b.getRange(i + 1, 2), r[1]); });
}

/* Répare les formules si elles affichent une erreur (ex. ancienne version du script). */
function repair_() {
  const ss = ss_();
  const bad = rg => /ERROR|ERREUR/i.test(rg.getDisplayValue());
  const b = ss.getSheetByName('Bilan'), c = ss.getSheetByName('Cartes'), f = ss.getSheetByName('Frais'), it = ss.getSheetByName('Items');
  const broken = bad(b.getRange('B9')) || (c.getLastRow() > 1 && bad(c.getRange(2, 15))) || (f.getLastRow() > 1 && bad(f.getRange(2, 8)));
  if (!broken) return;
  PropertiesService.getScriptProperties().deleteProperty('SEP');
  bilanFormulas_(b);
  for (let r = 2; r <= c.getLastRow(); r++) rowFormulas_('Cartes', c, r);
  for (let r = 2; r <= f.getLastRow(); r++) rowFormulas_('Frais', f, r);
  for (let r = 2; r <= it.getLastRow(); r++) rowFormulas_('Items', it, r);
}

function read_() {
  setup();
  const ss = ss_();
  const tz = ss.getSpreadsheetTimeZone();
  const rd = (name, keys) => {
    const sh = ss.getSheetByName(name);
    const n = sh.getLastRow() - 1;
    if (n < 1) return [];
    return sh.getRange(2, 1, n, keys.length).getValues().filter(r => r[0] !== '').map(r => {
      const o = {};
      keys.forEach((k, i) => { let v = r[i]; if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd'); o[k] = v; });
      return o;
    });
  };
  return { ok: true, v: 3, url: ss.getUrl(), cartes: rd('Cartes', CK), items: rd('Items', IK), frais: rd('Frais', FK) };
}

function cell_(k, v) {
  if (v === undefined || v === null || v === '') return '';
  if (/Date$/.test(k) || k === 'start' || k === 'end') {
    const p = String(v).split('-').map(Number);
    return p.length === 3 ? new Date(p[0], p[1] - 1, p[2], 12) : '';
  }
  if (['buyPrice', 'gradeCost', 'sellPrice', 'amount', 'qty'].indexOf(k) >= 0) return Number(v);
  return String(v);
}

function rowOf_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function upsert_(name, keys, data) {
  const sh = ss_().getSheetByName(name);
  let r = rowOf_(sh, data.id);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, keys.length).setValues([keys.map(k => cell_(k, data[k]))]);
  rowFormulas_(name, sh, r);
}

function rowFormulas_(name, sh, r) {
  if (name === 'Cartes') {
    F_(sh.getRange(r, 14), '=IF(L' + r + '="","",L' + r + '-J' + r + '-K' + r + ')');
    F_(sh.getRange(r, 15), '=IF(L' + r + '="","En stock","Vendue")');
  } else if (name === 'Items') {
    F_(sh.getRange(r, 11), '=IF(I' + r + '="","",I' + r + '-H' + r + ')');
    F_(sh.getRange(r, 12), '=IF(I' + r + '="","En stock","Vendu")');
  } else {
    const fin = 'IF(G' + r + '="",TODAY(),MIN(G' + r + ',TODAY()))';
    F_(sh.getRange(r, 8), '=IFERROR(IF(E' + r + '="Ponctuel",D' + r + ',(IF(E' + r + '="Annuel",DATEDIF(F' + r + ',' + fin + ',"Y"),DATEDIF(F' + r + ',' + fin + ',"M"))+1)*D' + r + '),0)');
  }
}

function del_(name, id) {
  const sh = ss_().getSheetByName(name);
  const r = rowOf_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try { return out_(read_()); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    setup();
    if (req.action === 'saveCard') upsert_('Cartes', CK, req.data);
    else if (req.action === 'deleteCard') del_('Cartes', req.id);
    else if (req.action === 'saveItem') upsert_('Items', IK, req.data);
    else if (req.action === 'deleteItem') del_('Items', req.id);
    else if (req.action === 'saveCost') upsert_('Frais', FK, req.data);
    else if (req.action === 'deleteCost') del_('Frais', req.id);
    else throw new Error('Action inconnue : ' + req.action);
    SpreadsheetApp.flush();
    return out_(read_());
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

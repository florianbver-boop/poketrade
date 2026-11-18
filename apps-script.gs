/**
 * PokeTrade — liaison entre l'app et ce Google Sheet.
 * À coller dans : Extensions > Apps Script (remplace tout le contenu), puis Déployer > Nouveau déploiement > Application Web.
 * Les onglets Bilan, Cartes et Frais sont créés automatiquement au premier appel.
 */
const CARTES = ['ID', 'Achetée le', 'Type', 'Nom', 'Numéro', 'Série', 'Langue', 'Gradation', 'Note', 'Prix achat', 'Coût gradation', 'Prix vente', 'Vendue le', 'Bénéfice', 'Statut'];
const CK = ['id', 'buyDate', 'type', 'name', 'number', 'set', 'lang', 'gradeCompany', 'grade', 'buyPrice', 'gradeCost', 'sellPrice', 'sellDate'];
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
  if (!ss.getSheetByName('Bilan')) {
    const b = ss.insertSheet('Bilan', 0);
    const rows = [
      ['Bilan PokeTrade', ''],
      ['', ''],
      ['Cartes achetées', '=COUNTA(Cartes!A2:A)'],
      ['Cartes vendues', '=COUNT(Cartes!L2:L)'],
      ['Total achats', '=SUM(Cartes!J2:J)'],
      ['Total gradation', '=SUM(Cartes!K2:K)'],
      ['Total ventes', '=SUM(Cartes!L2:L)'],
      ['Coût des cartes vendues (achat + gradation)', '=SUMIFS(Cartes!J2:J,Cartes!L2:L,"<>")+SUMIFS(Cartes!K2:K,Cartes!L2:L,"<>")'],
      ['Marge sur les ventes', '=B7-B8'],
      ['Frais annexes (bot, cookgroup, proxies…)', '=SUM(Frais!H2:H)'],
      ['BÉNÉFICE NET', '=B9-B10'],
      ['ROI net', '=IFERROR(B11/(B8+B10),0)'],
      ['Stock non vendu (coût)', '=B5+B6-B8'],
      ['Trésorerie (ventes − toutes les dépenses)', '=B7-B5-B6-B10']
    ];
    b.getRange(1, 1, rows.length, 2).setValues(rows);
    b.getRange('A1').setFontSize(16).setFontWeight('bold');
    b.getRange('B5:B11').setNumberFormat(PROFIT);
    b.getRange('B13:B14').setNumberFormat(PROFIT);
    b.getRange('B12').setNumberFormat('0.0%');
    b.getRange('A11:B11').setFontWeight('bold').setFontSize(13).setBackground('#fff4c2');
    b.setColumnWidth(1, 330);
    b.setColumnWidth(2, 140);
  }
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
  return { ok: true, url: ss.getUrl(), cartes: rd('Cartes', CK), frais: rd('Frais', FK) };
}

function cell_(k, v) {
  if (v === undefined || v === null || v === '') return '';
  if (/Date$/.test(k) || k === 'start' || k === 'end') {
    const p = String(v).split('-').map(Number);
    return p.length === 3 ? new Date(p[0], p[1] - 1, p[2], 12) : '';
  }
  if (['buyPrice', 'gradeCost', 'sellPrice', 'amount'].indexOf(k) >= 0) return Number(v);
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
  if (name === 'Cartes') {
    sh.getRange(r, 14, 1, 2).setFormulas([[
      '=IF(L' + r + '="","",L' + r + '-J' + r + '-K' + r + ')',
      '=IF(L' + r + '="","En stock","Vendue")'
    ]]);
  } else {
    const fin = 'IF(G' + r + '="",TODAY(),MIN(G' + r + ',TODAY()))';
    sh.getRange(r, 8).setFormula('=IFERROR(IF(E' + r + '="Ponctuel",D' + r + ',(IF(E' + r + '="Annuel",DATEDIF(F' + r + ',' + fin + ',"Y"),DATEDIF(F' + r + ',' + fin + ',"M"))+1)*D' + r + '),0)');
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
    else if (req.action === 'saveCost') upsert_('Frais', FK, req.data);
    else if (req.action === 'deleteCost') del_('Frais', req.id);
    SpreadsheetApp.flush();
    return out_(read_());
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

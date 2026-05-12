/**
 * Code.gs — Backend Google Apps Script pour Kermesse 2026
 *
 * DÉPLOIEMENT :
 * 1. Ouvre un Google Sheet neuf (ou existant)
 * 2. Extensions → Apps Script → colle ce code
 * 3. Lance setup() UNE SEULE FOIS pour créer les feuilles
 * 4. Déployer → Nouveau déploiement
 *    - Type : Application Web
 *    - Exécuter en tant que : Moi
 *    - Accès : Tout le monde (anonyme)
 * 5. Copie l'URL obtenue → colle-la dans CFG.GAS_URL dans index.html
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SH = {
  STANDS:  'Stands',
  INSCRIP: 'Inscriptions_Stands',
  GATEAUX: 'Inscriptions_Gateaux',
};

/* ======================================================
   doGet — toutes les opérations via JSONP (GET)
====================================================== */
function doGet(e) {
  const p = e.parameter || {};
  let data;

  switch (p.action || 'grid') {
    case 'grid':               data = getGrid();                        break;
    case 'lookup':             data = lookupTel(p.telephone);           break;
    case 'inscription_stand':  data = inscriptionStand(p);              break;
    case 'inscription_gateau': data = inscriptionGateau(p);             break;
    case 'delete':             data = deleteRow(p.id, p.telephone);     break;
    default:                   data = { ok: false, error: 'unknown' };
  }

  // JSONP : le front passe ?callback=__gcbXXX
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(data);
}

/* ======================================================
   doPost — inscriptions & suppression (URLSearchParams)
====================================================== */
function doPost(e) {
  const p = e.parameter || {};

  switch (p.action) {
    case 'inscription_stand':  return json(inscriptionStand(p));
    case 'inscription_gateau': return json(inscriptionGateau(p));
    case 'delete':             return json(deleteRow(p.id, p.telephone));
    default:                   return json({ ok: false, error: 'unknown_action' });
  }
}

/* ======================================================
   getGrid — avec cache 30 s (CacheService)
====================================================== */
function getGrid() {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get('grid');
  if (hit) {
    try { return JSON.parse(hit); } catch {}
  }
  const data = buildGrid();
  try { cache.put('grid', JSON.stringify(data), 30); } catch {}
  return data;
}

function invalidateGrid() {
  try { CacheService.getScriptCache().remove('grid'); } catch {}
}

function buildGrid() {
  const shStands  = SS.getSheetByName(SH.STANDS);
  const shInscrip = SS.getSheetByName(SH.INSCRIP);
  if (!shStands) return { ok: false, error: 'sheet_missing' };

  const standsRows = shStands.getDataRange().getValues();   // [id, nom, cap_16h45, cap_17h30, cap_18h15]
  const inscRows   = shInscrip ? shInscrip.getDataRange().getValues() : [];

  const CRENEAUX = ['16h45', '17h30', '18h15'];

  // Map standId+créneau → liste bénévoles
  const inscMap = {};
  for (let i = 1; i < inscRows.length; i++) {
    const [rowId, , prenom, nom, , prenomEnfant, classe, standId, creneau] = inscRows[i];
    if (!standId || !creneau) continue;
    const key = `${standId}__${creneau}`;
    if (!inscMap[key]) inscMap[key] = [];
    inscMap[key].push({ p: prenom, n: nom, e: prenomEnfant || '', cl: classe || '' });
  }

  let totalPlaces = 0, totalInscrits = 0, standsIncomplets = 0;
  const stands = [];

  for (let i = 1; i < standsRows.length; i++) {
    const [id, nom, cap1, cap2, cap3] = standsRows[i];
    if (!id || !nom) continue;

    const caps = [cap1, cap2, cap3].map(c => parseInt(c) || 0);
    const slots = CRENEAUX.map((cr, ci) => {
      const cap = caps[ci];
      const b   = inscMap[`${id}__${cr}`] || [];
      return { c: cr, cap, ins: b.length, b };
    });

    const sp = slots.reduce((s, sl) => s + sl.cap, 0);
    const si = slots.reduce((s, sl) => s + sl.ins, 0);
    totalPlaces   += sp;
    totalInscrits += si;
    if (si < sp) standsIncomplets++;

    stands.push({ id: String(id), nom: String(nom), slots });
  }

  return {
    ok: true,
    stats: { inscrits: totalInscrits, total_places: totalPlaces, stands_incomplets: standsIncomplets },
    stands,
  };
}

/* ======================================================
   inscriptionStand
====================================================== */
function inscriptionStand(p) {
  const sh = SS.getSheetByName(SH.INSCRIP);
  if (!sh) return { ok: false, error: 'sheet_missing' };

  // Vérifie qu'il reste de la place
  const grid  = getGrid();
  const stand = grid.stands.find(s => s.id === p.stand_id);
  if (!stand) return { ok: false, error: 'stand_not_found' };
  const slot = stand.slots.find(s => s.c === p.creneau);
  if (!slot) return { ok: false, error: 'slot_not_found' };
  if (slot.ins >= slot.cap) return { ok: false, error: 'full', message: 'Ce créneau est complet.' };

  const id = Utilities.getUuid();
  sh.appendRow([
    id,
    new Date(),
    p.prenom        || '',
    p.nom           || '',
    p.telephone     || '',
    p.prenom_enfant || '',
    p.classe        || '',
    p.stand_id      || '',
    p.creneau       || '',
    p.stand_nom     || '',
  ]);

  invalidateGrid();
  return { ok: true, wrote: true, id };
}

/* ======================================================
   inscriptionGateau
====================================================== */
function inscriptionGateau(p) {
  const sh = SS.getSheetByName(SH.GATEAUX);
  if (!sh) return { ok: false, error: 'sheet_missing' };

  const id = Utilities.getUuid();
  sh.appendRow([
    id,
    new Date(),
    p.prenom_nom    || '',
    p.telephone     || '',
    p.prenom_enfant || '',
    p.type_gateau   || '',
    p.nom_gateau    || '',
    p.parts         || '',
    p.depot         || '',
  ]);

  invalidateGrid();
  return { ok: true, wrote: true, id };
}

/* ======================================================
   lookupTel — trouve toutes les inscriptions d'un parent
====================================================== */
function lookupTel(telephone) {
  if (!telephone) return { ok: false };
  const tel = String(telephone).replace(/\s/g, '');
  const inscriptions = [];

  const shInscrip = SS.getSheetByName(SH.INSCRIP);
  if (shInscrip) {
    const rows = shInscrip.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const [rowId, , prenom, nom, phone, , , standId, creneau, standNom] = rows[i];
      if (String(phone).replace(/\s/g, '') === tel) {
        inscriptions.push({ id: rowId, type: 'stand', stand_nom: standNom || standId, creneau, prenom, nom });
      }
    }
  }

  const shGateaux = SS.getSheetByName(SH.GATEAUX);
  if (shGateaux) {
    const rows = shGateaux.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const [rowId, , prenomNom, phone, , typeGateau, nomGateau, , depot] = rows[i];
      if (String(phone).replace(/\s/g, '') === tel) {
        inscriptions.push({ id: rowId, type: 'gateau', nom_gateau: nomGateau, type_gateau: typeGateau, depot });
      }
    }
  }

  return { ok: true, inscriptions };
}

/* ======================================================
   deleteRow — supprime une inscription (vérifie le tel)
====================================================== */
function deleteRow(id, telephone) {
  if (!id || !telephone) return { ok: false };
  const tel = String(telephone).replace(/\s/g, '');

  function tryDelete(sheetName, telCol) {
    const sh = SS.getSheetByName(sheetName);
    if (!sh) return false;
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) &&
          String(rows[i][telCol]).replace(/\s/g, '') === tel) {
        sh.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  }

  // col 4 = téléphone dans Inscriptions_Stands ; col 3 dans Inscriptions_Gateaux
  if (tryDelete(SH.INSCRIP, 4)) { invalidateGrid(); return { ok: true, wrote: true }; }
  if (tryDelete(SH.GATEAUX, 3)) { invalidateGrid(); return { ok: true, wrote: true }; }
  return { ok: false, error: 'not_found' };
}

/* ======================================================
   json — helper Content Service
====================================================== */
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================================================
   setup — à lancer UNE SEULE FOIS depuis l'éditeur Apps Script
====================================================== */
function setup() {
  // Crée les feuilles manquantes
  Object.values(SH).forEach(name => {
    if (!SS.getSheetByName(name)) SS.insertSheet(name);
  });

  // Feuille Stands — en-têtes + données initiales
  const shStands = SS.getSheetByName(SH.STANDS);
  if (shStands.getLastRow() === 0) {
    shStands.appendRow(['id', 'nom', 'cap_16h45', 'cap_17h30', 'cap_18h15']);
    [
      ['s0',  'Caisse — Vente tickets',          0, 0, 0],
      ['s1',  'Buvette',                          4, 4, 4],
      ['s2',  'Chamboule-tout',                   2, 2, 2],
      ['s3',  'Lancé tête de clown',              2, 2, 2],
      ['s4',  'Pêche aux canards',                2, 2, 2],
      ['s5',  'Maquillage',                       3, 3, 3],
      ['s6',  'Stand créatif — clowns & masques', 2, 2, 2],
      ['s7',  "Toucher à l'aveugle",              0, 0, 0],
      ['s8',  'Grenouille tonneau',               2, 2, 2],
      ['s9',  'Jeu de massacre',                  2, 2, 2],
      ['s10', "Tir à l'arc — initiation",         2, 2, 2],
      ['s11', 'Course en sac',                    2, 2, 2],
      ['s12', 'Tombola — tirage',                 2, 2, 2],
      ['s13', 'Accueil & orientation',            3, 3, 3],
      ['s14', 'Vente de glaces',                  2, 2, 2],
    ].forEach(row => shStands.appendRow(row));
    Logger.log('Feuille Stands créée avec %s stands.', shStands.getLastRow() - 1);
  }

  // Feuille Inscriptions_Stands — en-têtes
  const shInscrip = SS.getSheetByName(SH.INSCRIP);
  if (shInscrip.getLastRow() === 0) {
    shInscrip.appendRow(['id', 'timestamp', 'prenom', 'nom', 'telephone', 'prenom_enfant', 'classe', 'stand_id', 'creneau', 'stand_nom']);
  }

  // Feuille Inscriptions_Gateaux — en-têtes
  const shGateaux = SS.getSheetByName(SH.GATEAUX);
  if (shGateaux.getLastRow() === 0) {
    shGateaux.appendRow(['id', 'timestamp', 'prenom_nom', 'telephone', 'prenom_enfant', 'type_gateau', 'nom_gateau', 'parts', 'depot']);
  }

  Logger.log('✅ Setup terminé ! Tu peux maintenant déployer comme Web App.');
}

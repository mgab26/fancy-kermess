/**
 * Code.gs — Backend Google Apps Script pour Kermesse 2026
 *
 * ARCHITECTURE (Option B) :
 * - Ce script est associé au Google Sheet "Kermesse Couteron Bénévoles"
 * - buildGrid() lit les inscriptions directes depuis "Inscriptions bénévoles"
 *   ET les inscriptions soumises via l'app depuis "Inscriptions_Stands"
 * - Les écritures de l'app (inscription/suppression) vont dans Inscriptions_Stands
 * - Les inscriptions directes dans le sheet sont visibles mais non-supprimables via l'app
 *
 * DÉPLOIEMENT :
 * 1. Ouvre le Google Sheet "Kermesse Couteron Bénévoles"
 * 2. Extensions → Apps Script → colle ce code
 * 3. Lance setup() UNE SEULE FOIS pour créer les feuilles internes
 * 4. Déployer → Nouveau déploiement
 *    - Type : Application Web
 *    - Exécuter en tant que : Moi
 *    - Accès : Tout le monde (anonyme)
 * 5. Copie l'URL → CFG.GAS_URL dans index.html
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SH = {
  BENEV:   'Inscriptions bénévoles',  // feuille parent — lecture seule
  INSCRIP: 'Inscriptions_Stands',     // inscriptions app — lecture/écriture
  GATEAUX: 'Inscriptions_Gateaux',    // gâteaux app — lecture/écriture
};

/**
 * Définition des stands : id, nom affiché, lignes dans "Inscriptions bénévoles" (1-indexed).
 * La capacité par créneau = nombre de lignes du stand.
 * Pour ajuster la capacité d'un stand, ajoute ou retire des numéros de lignes.
 */
const STAND_DEFS = [
  { id: 's0',  nom: 'Entrée - Caisse - Vente tickets', rows: [7, 8] },
  { id: 's1',  nom: 'Buvette',                          rows: [9, 10, 11, 12, 13] },
  { id: 's2',  nom: 'Chamboule tout',                   rows: [19, 20] },
  { id: 's3',  nom: 'Lancé tête de clown',              rows: [21] },
  { id: 's4',  nom: 'Pêche au canard',                  rows: [22] },
  { id: 's5',  nom: 'Maquillage',                       rows: [23, 24, 25] },
  { id: 's6',  nom: 'Stand créatif - clowns et masques', rows: [26, 27, 28] },
  { id: 's7',  nom: "Toucher à l'aveugle (x2)",         rows: [29] },
  { id: 's8',  nom: 'Jeu en bois - Grenouille tonneau', rows: [30] },
  { id: 's9',  nom: 'Jeu en bois - Le piège',           rows: [31] },
  { id: 's10', nom: 'Jeu en bois - La Table élastique', rows: [32] },
  { id: 's11', nom: 'Jeu en bois - Puissance 4 Géant',  rows: [33] },
  { id: 's12', nom: 'Jeu en bois - Jeu des bâtonnets',  rows: [34] },
  { id: 's13', nom: 'Jeu en bois - Parcours élec spirale', rows: [35] },
  { id: 's14', nom: 'Jeu en bois - Jeu équilibre',      rows: [36] },
  { id: 's15', nom: 'Jeu en bois - Aerobille',          rows: [37] },
  { id: 's16', nom: 'Jeu en bois - Billard carrousel',  rows: [38] },
  { id: 's17', nom: 'Jeu en bois - Cornhole',           rows: [39] },
  { id: 's18', nom: 'Jeu en bois - La roulette',        rows: [40] },
  { id: 's19', nom: 'Stand Tatouage',                   rows: [41, 42] },
  { id: 's20', nom: 'Activités diverses cirque',        rows: [43, 44, 45] },
];

// Mapping créneaux → colonnes dans la feuille (0-indexed)
const SLOT_COLS = [
  { creneau: '16h45', colNom: 2, colEnfant: 3 },
  { creneau: '17h30', colNom: 4, colEnfant: 5 },
  { creneau: '18h15', colNom: 6, colEnfant: 7 },
];

/* ======================================================
   doGet — toutes les opérations via JSONP (GET)
====================================================== */
function doGet(e) {
  const p = e.parameter || {};
  let data;

  switch (p.action || 'grid') {
    case 'grid':               data = getGrid();                    break;
    case 'lookup':             data = lookupTel(p.telephone);       break;
    case 'inscription_stand':  data = inscriptionStand(p);          break;
    case 'inscription_gateau': data = inscriptionGateau(p);         break;
    case 'delete':             data = deleteRow(p.id, p.telephone); break;
    default:                   data = { ok: false, error: 'unknown' };
  }

  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(data);
}

/* ======================================================
   doPost — inscriptions & suppression
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
   getGrid — avec cache 30 s
====================================================== */
function getGrid() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('grid');
  if (hit) { try { return JSON.parse(hit); } catch {} }
  const data = buildGrid();
  try { cache.put('grid', JSON.stringify(data), 30); } catch {}
  return data;
}

function invalidateGrid() {
  try { CacheService.getScriptCache().remove('grid'); } catch {}
}

/* ======================================================
   buildGrid — fusionne inscriptions sheet + inscriptions app
====================================================== */
function buildGrid() {
  const sheetMap = readSheetSignups();
  const appMap   = readAppSignups();

  let totalPlaces = 0, totalInscrits = 0, standsIncomplets = 0;
  const stands = [];

  for (const def of STAND_DEFS) {
    const cap = def.rows.length; // capacité par créneau = nb de lignes

    const slots = SLOT_COLS.map(({ creneau }) => {
      const key  = `${def.id}__${creneau}`;
      const b    = [...(sheetMap[key] || []), ...(appMap[key] || [])];
      return { c: creneau, cap, ins: b.length, b };
    });

    const sp = slots.reduce((s, sl) => s + sl.cap, 0);
    const si = slots.reduce((s, sl) => s + sl.ins, 0);
    totalPlaces   += sp;
    totalInscrits += si;
    if (si < sp) standsIncomplets++;

    stands.push({ id: def.id, nom: def.nom, slots });
  }

  return {
    ok: true,
    stats: { inscrits: totalInscrits, total_places: totalPlaces, stands_incomplets: standsIncomplets },
    stands,
  };
}

/* ======================================================
   readSheetSignups — parse "Inscriptions bénévoles" (lecture seule)
====================================================== */
function readSheetSignups() {
  const sh = SS.getSheetByName(SH.BENEV);
  if (!sh) return {};

  const allValues = sh.getDataRange().getValues();
  const result = {};

  for (const def of STAND_DEFS) {
    for (const { creneau, colNom, colEnfant } of SLOT_COLS) {
      const key = `${def.id}__${creneau}`;
      result[key] = [];

      for (const rowNum of def.rows) {
        const row = allValues[rowNum - 1];
        if (!row) continue;

        const nomContact = String(row[colNom]    || '').trim();
        const enfantInfo  = String(row[colEnfant] || '').trim();

        if (nomContact) {
          result[key].push({ p: nomContact, n: '', e: enfantInfo, cl: '', src: 'sheet' });
        }
      }
    }
  }

  return result;
}

/* ======================================================
   readAppSignups — lit Inscriptions_Stands
====================================================== */
function readAppSignups() {
  const sh = SS.getSheetByName(SH.INSCRIP);
  if (!sh) return {};

  const rows = sh.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < rows.length; i++) {
    const [rowId, , prenom, nom, , prenomEnfant, classe, standId, creneau] = rows[i];
    if (!standId || !creneau) continue;
    const key = `${standId}__${creneau}`;
    if (!result[key]) result[key] = [];
    result[key].push({ id: rowId, p: prenom, n: nom, e: prenomEnfant || '', cl: classe || '', src: 'app' });
  }

  return result;
}

/* ======================================================
   inscriptionStand
====================================================== */
function inscriptionStand(p) {
  const sh = SS.getSheetByName(SH.INSCRIP);
  if (!sh) return { ok: false, error: 'sheet_missing' };

  const grid  = getGrid();
  const stand = grid.stands.find(s => s.id === p.stand_id);
  if (!stand) return { ok: false, error: 'stand_not_found' };
  const slot = stand.slots.find(s => s.c === p.creneau);
  if (!slot)  return { ok: false, error: 'slot_not_found' };
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
   lookupTel — cherche dans les inscriptions app
   (les inscriptions directes dans le sheet n'ont pas de tel structuré)
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
   deleteRow — supprime uniquement les inscriptions app
   (les inscriptions directes dans le sheet ne sont pas supprimables via l'app)
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

  if (tryDelete(SH.INSCRIP, 4)) { invalidateGrid(); return { ok: true, wrote: true }; }
  if (tryDelete(SH.GATEAUX, 3)) { invalidateGrid(); return { ok: true, wrote: true }; }
  return { ok: false, error: 'not_found' };
}

/* ======================================================
   json helper
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
  [SH.INSCRIP, SH.GATEAUX].forEach(name => {
    if (!SS.getSheetByName(name)) SS.insertSheet(name);
  });

  const shInscrip = SS.getSheetByName(SH.INSCRIP);
  if (shInscrip.getLastRow() === 0) {
    shInscrip.appendRow(['id', 'timestamp', 'prenom', 'nom', 'telephone',
                         'prenom_enfant', 'classe', 'stand_id', 'creneau', 'stand_nom']);
  }

  const shGateaux = SS.getSheetByName(SH.GATEAUX);
  if (shGateaux.getLastRow() === 0) {
    shGateaux.appendRow(['id', 'timestamp', 'prenom_nom', 'telephone',
                         'prenom_enfant', 'type_gateau', 'nom_gateau', 'parts', 'depot']);
  }

  Logger.log('✅ Setup OK. Déploie maintenant comme Web App.');
}

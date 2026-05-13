/**
 * Code.gs — Backend Google Apps Script pour Kermesse 2026
 *
 * ARCHITECTURE :
 * - Ce script est associé au Google Sheet "Kermesse Couteron Bénévoles"
 * - Toutes les inscriptions stand (app + directes) vivent dans "Inscriptions bénévoles"
 * - L'app écrit directement dans les cellules du sheet (Nom + Contact + Enfant et classe)
 * - Seules les inscriptions avec un n° de téléphone dans la colonne Contact peuvent
 *   être supprimées via l'app
 * - Les gâteaux sont stockés dans "Inscriptions_Gateaux" (onglet séparé)
 *
 * DÉPLOIEMENT :
 * 1. Ouvre le Google Sheet "Kermesse Couteron Bénévoles"
 * 2. Extensions → Apps Script → colle ce code
 * 3. Lance setup() UNE SEULE FOIS pour créer l'onglet Inscriptions_Gateaux
 * 4. Déployer → Nouveau déploiement
 *    - Type : Application Web
 *    - Exécuter en tant que : Moi
 *    - Accès : Tout le monde (anonyme)
 * 5. Copie l'URL → CFG.GAS_URL dans index.html
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SH = {
  BENEV:   'Inscriptions bénévoles',  // inscriptions stands — lecture/écriture
  GATEAUX: 'Inscriptions_Gateaux',   // gâteaux — lecture/écriture
};

/**
 * Définition des stands : id, nom affiché, lignes dans "Inscriptions bénévoles" (1-indexed).
 * La capacité par créneau = nombre de lignes du stand.
 */
const STAND_DEFS = [
  { id: 's0',  nom: 'Entrée - Caisse - Vente tickets',       rows: [7, 8] },
  { id: 's1',  nom: 'Buvette',                               rows: [9, 10, 11, 12, 13] },
  { id: 's2',  nom: 'Chamboule tout',                        rows: [19, 20] },
  { id: 's3',  nom: 'Lancé tête de clown',                   rows: [21] },
  { id: 's4',  nom: 'Pêche au canard',                       rows: [22] },
  { id: 's5',  nom: 'Maquillage',                            rows: [23, 24, 25] },
  { id: 's6',  nom: 'Stand créatif - clowns et masques',     rows: [26, 27, 28] },
  { id: 's7',  nom: "Toucher à l'aveugle (x2)",              rows: [29] },
  { id: 's8',  nom: 'Jeu en bois - Grenouille tonneau',      rows: [30] },
  { id: 's9',  nom: 'Jeu en bois - Le piège',                rows: [31] },
  { id: 's10', nom: 'Jeu en bois - La Table élastique',      rows: [32] },
  { id: 's11', nom: 'Jeu en bois - Puissance 4 Géant',       rows: [33] },
  { id: 's12', nom: 'Jeu en bois - Jeu des bâtonnets',       rows: [34] },
  { id: 's13', nom: 'Jeu en bois - Parcours élec spirale',   rows: [35] },
  { id: 's14', nom: 'Jeu en bois - Jeu équilibre',           rows: [36] },
  { id: 's15', nom: 'Jeu en bois - Aerobille',               rows: [37] },
  { id: 's16', nom: 'Jeu en bois - Billard carrousel',       rows: [38] },
  { id: 's17', nom: 'Jeu en bois - Cornhole',                rows: [39] },
  { id: 's18', nom: 'Jeu en bois - La roulette',             rows: [40] },
  { id: 's19', nom: 'Stand Tatouage',                        rows: [41, 42] },
  { id: 's20', nom: 'Activités diverses cirque',             rows: [43, 44, 45] },
];

// Colonnes dans "Inscriptions bénévoles" (0-indexed)
// Structure par créneau : Nom | Contact | Enfant et classe
const SLOT_COLS = [
  { creneau: '16h45', colNom: 2, colContact: 3, colEnfant: 4 },
  { creneau: '17h30', colNom: 5, colContact: 6, colEnfant: 7 },
  { creneau: '18h15', colNom: 8, colContact: 9, colEnfant: 10 },
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
   buildGrid — lit directement "Inscriptions bénévoles"
====================================================== */
function buildGrid() {
  const sh = SS.getSheetByName(SH.BENEV);
  const allValues = sh ? sh.getDataRange().getValues() : [];

  let totalPlaces = 0, totalInscrits = 0, standsIncomplets = 0;
  const stands = [];

  for (const def of STAND_DEFS) {
    const cap = def.rows.length;

    const slots = SLOT_COLS.map(({ creneau, colNom, colContact, colEnfant }) => {
      const b = [];
      for (const rowNum of def.rows) {
        const row = allValues[rowNum - 1];
        if (!row) continue;
        const nomVal     = String(row[colNom]     || '').trim();
        const contactVal = String(row[colContact] || '').trim();
        const enfantVal  = String(row[colEnfant]  || '').trim();
        if (nomVal) {
          const id = contactVal ? `${def.id}__${creneau}__${rowNum}` : null;
          b.push({ id, p: nomVal, n: '', e: enfantVal, tel: contactVal });
        }
      }
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
   inscriptionStand — écrit directement dans "Inscriptions bénévoles"
====================================================== */
function inscriptionStand(p) {
  const sh = SS.getSheetByName(SH.BENEV);
  if (!sh) return { ok: false, error: 'sheet_missing' };

  const def = STAND_DEFS.find(d => d.id === p.stand_id);
  if (!def) return { ok: false, error: 'stand_not_found' };

  const slotCols = SLOT_COLS.find(s => s.creneau === p.creneau);
  if (!slotCols) return { ok: false, error: 'slot_not_found' };

  const allValues = sh.getDataRange().getValues();

  // Trouver la première ligne vide du stand pour ce créneau
  for (const rowNum of def.rows) {
    const row = allValues[rowNum - 1];
    if (!row) continue;
    const nomVal = String(row[slotCols.colNom] || '').trim();
    if (!nomVal) {
      const nomComplet = [p.prenom, p.nom].filter(Boolean).join(' ');
      const enfantInfo = [p.prenom_enfant, p.classe].filter(Boolean).join(' · ');

      sh.getRange(rowNum, slotCols.colNom     + 1).setValue(nomComplet);
      sh.getRange(rowNum, slotCols.colContact + 1).setValue(p.telephone || '');
      sh.getRange(rowNum, slotCols.colEnfant  + 1).setValue(enfantInfo);

      const id = `${p.stand_id}__${p.creneau}__${rowNum}`;
      invalidateGrid();
      return { ok: true, wrote: true, id };
    }
  }

  return { ok: false, error: 'full', message: 'Ce créneau est complet.' };
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

  return { ok: true, wrote: true, id };
}

/* ======================================================
   lookupTel — cherche dans "Inscriptions bénévoles" + gâteaux
====================================================== */
function lookupTel(telephone) {
  if (!telephone) return { ok: false };
  const tel = String(telephone).replace(/\s/g, '');
  const inscriptions = [];

  const sh = SS.getSheetByName(SH.BENEV);
  if (sh) {
    const allValues = sh.getDataRange().getValues();
    for (const def of STAND_DEFS) {
      for (const { creneau, colNom, colContact } of SLOT_COLS) {
        for (const rowNum of def.rows) {
          const row = allValues[rowNum - 1];
          if (!row) continue;
          const contactVal = String(row[colContact] || '').replace(/\s/g, '');
          if (contactVal === tel) {
            const nomVal = String(row[colNom] || '').trim();
            const id = `${def.id}__${creneau}__${rowNum}`;
            inscriptions.push({ id, type: 'stand', stand_nom: def.nom, creneau, prenom: nomVal, nom: '' });
          }
        }
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
   deleteRow — efface les cellules dans "Inscriptions bénévoles"
   (uniquement si le Contact correspond au numéro fourni)
====================================================== */
function deleteRow(id, telephone) {
  if (!id || !telephone) return { ok: false };
  const tel = String(telephone).replace(/\s/g, '');

  // Format ID pour les stands : "standId__creneau__rowNum"
  const parts = String(id).split('__');
  if (parts.length === 3) {
    const [, creneau, rowNumStr] = parts;
    const rowNum = parseInt(rowNumStr, 10);
    const slotCols = SLOT_COLS.find(s => s.creneau === creneau);

    if (slotCols && rowNum) {
      const sh = SS.getSheetByName(SH.BENEV);
      if (sh) {
        const contactVal = String(
          sh.getRange(rowNum, slotCols.colContact + 1).getValue() || ''
        ).replace(/\s/g, '');

        if (contactVal === tel) {
          sh.getRange(rowNum, slotCols.colNom     + 1).clearContent();
          sh.getRange(rowNum, slotCols.colContact + 1).clearContent();
          sh.getRange(rowNum, slotCols.colEnfant  + 1).clearContent();
          invalidateGrid();
          return { ok: true, wrote: true };
        }
      }
    }
  }

  // Suppression gâteau (UUID dans Inscriptions_Gateaux)
  const shGateaux = SS.getSheetByName(SH.GATEAUX);
  if (shGateaux) {
    const rows = shGateaux.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) &&
          String(rows[i][3]).replace(/\s/g, '') === tel) {
        shGateaux.deleteRow(i + 1);
        return { ok: true, wrote: true };
      }
    }
  }

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
  if (!SS.getSheetByName(SH.GATEAUX)) SS.insertSheet(SH.GATEAUX);

  const shGateaux = SS.getSheetByName(SH.GATEAUX);
  if (shGateaux.getLastRow() === 0) {
    shGateaux.appendRow(['id', 'timestamp', 'prenom_nom', 'telephone',
                         'prenom_enfant', 'type_gateau', 'nom_gateau', 'parts', 'depot']);
  }

  Logger.log('✅ Setup OK. Déploie maintenant comme Web App.');
}

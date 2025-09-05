// === CONFIGURATION ===
const PAPPERS_API_KEY = 'YOURAPIKEY';

// === TRAITE UNE SEULE LIGNE ===
function processSingleRow(row) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const siret = sheet.getRange(row, 2).getValue(); // Colonne B - SIRET

  if (!siret) {
    Logger.log(`Pas de SIRET en ligne ${row}`);
    return;
  }

  // Clean the SIRET (remove spaces)
  const cleanSiret = siret.toString().replace(/\s/g, '');
  
  // Basic validation - check if it's mostly digits
  if (!/^\d+$/.test(cleanSiret) || cleanSiret.length < 11) {
    sheet.getRange(row, 4).setValue('SIRET invalide');
    Logger.log(`SIRET invalide ligne ${row}: ${cleanSiret}`);
    return;
  }

  const fullUrl = `https://api.pappers.fr/v2/entreprise?siret=${cleanSiret}`;

  const options = {
    method: 'get',
    headers: {
      'api-key': PAPPERS_API_KEY
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const json = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      sheet.getRange(row, 4).setValue(`Erreur API: ${response.getResponseCode()}`);
      Logger.log(`Erreur API ligne ${row}: ${response.getResponseCode()}`);
      return;
    }

    if (!json || json.erreur) {
      sheet.getRange(row, 4).setValue('Entreprise non trouvée');
      Logger.log(`Entreprise non trouvée ligne ${row}`);
      return;
    }

    // Extract specific data
    const extractedSiret = json.etablissement && json.etablissement.siret ? json.etablissement.siret : 'SIRET non trouvé';
    const companyName = json.nom_entreprise || 'Nom non trouvé';

    // Write results to columns D and E
    sheet.getRange(row, 4).setValue(extractedSiret); // Column D
    sheet.getRange(row, 5).setValue(companyName); // Column E
    
    Logger.log(`✅ Résultat écrit ligne ${row}: ${companyName}`);
    
  } catch (e) {
    Logger.log(`❌ Erreur ligne ${row}: ${e.message}`);
    sheet.getRange(row, 4).setValue(`Erreur: ${e.message}`);
  }
}

// === TEST MANUEL ===
function testRow2() {
  processSingleRow(2);
}

function testRow3() {
  processSingleRow(3);
}

// === PLAGE FIXE ===
function processRowRange(startRow, endRow) {
  Logger.log(`Traitement des lignes ${startRow} à ${endRow}`);
  for (let row = startRow; row <= endRow; row++) {
    processSingleRow(row);
    Utilities.sleep(300); // éviter rate-limit
  }
  Logger.log('Traitement terminé');
}

// === LIGNES SÉLECTIONNÉES ===
function processSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    Logger.log('Aucune sélection trouvée.');
    return;
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const endRow = startRow + numRows - 1;

  Logger.log(`Traitement des lignes sélectionnées ${startRow} à ${endRow}`);
  processRowRange(startRow, endRow);
}

// === TOUTES LES LIGNES AVEC DONNÉES ===
function processAllSIRETs() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  Logger.log(`Traitement de toutes les lignes (2 à ${lastRow})`);
  processRowRange(2, lastRow);
}

// === LIGNES MARQUÉES (par ex. avec "SIRET" en colonne quelconque) ===
function processMarkedRows(markerColumn = 6) { // Column F by default
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const marker = sheet.getRange(row, markerColumn).getValue();
    if (marker && marker.toString().toUpperCase() === 'SIRET') {
      processSingleRow(row);
      sheet.getRange(row, markerColumn).setValue('✔');
      Utilities.sleep(300);
    }
  }

  Logger.log('Traitement des lignes marquées terminé.');
}

// === UTILITAIRES ===
function processFirst10Rows() {
  processRowRange(2, 11);
}

function processRows3to4() {
  processRowRange(3, 4);
}

function setupHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1).setValue('company_id');
  sheet.getRange(1, 2).setValue('siret');
  sheet.getRange(1, 3).setValue('user');
  sheet.getRange(1, 4).setValue('siret_extracted');
  sheet.getRange(1, 5).setValue('company_name');
}

// === AUTO-TRIGGER (optional) ===
function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  
  // Check if the edited cell is in column B (column 2)
  if (range.getColumn() === 2) {
    const row = range.getRow();
    if (row > 1) { // Skip header row
      Logger.log(`Auto-processing row ${row} after edit`);
      processSingleRow(row);
    }
  }
}

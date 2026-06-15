/**
 * Fichier : Code.gs
 * Point d'entrée et API de la Web App.
 */

const SPREADSHEET_ID = '1NnsosT5vWlKtIhoXxzUeKZfE47Zy6Thahem3Ox5dwcM';
const SHEET_NAME = 'Content';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Formation OPTIMA - Portail Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================================================
// GESTION DU CONTENU (CMS GSHEET)
// =========================================================================

function initContentSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  const defaultContent = [
    ['Key', 'Value'],
    ['app_title', 'OPTIMA S&OP'],
    ['app_subtitle', 'Parcours Manager'],
    // ... (tronqué ici pour la lisibilité, gardez votre fonction d'origine si vous le souhaitez)
  ];
  
  sheet.getRange(1, 1, defaultContent.length, 2).setValues(defaultContent);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E2E8F0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
}

function getContent() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const content = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      content[data[i][0]] = data[i][1];
    }
  }
  return content;
}

function saveContent(updates) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const keysMap = {};
  
  for (let i = 1; i < data.length; i++) {
    keysMap[data[i][0]] = i + 1; 
  }
  
  for (let key in updates) {
    if (keysMap[key]) {
      sheet.getRange(keysMap[key], 2).setValue(updates[key]);
    } else {
      sheet.appendRow([key, updates[key]]);
    }
  }
  return true;
}

// =========================================================================
// GESTION DES RÉFÉRENTIELS (JOB CODES & BU)
// =========================================================================

function getJobCodes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Codes métiers');
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 
  
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const jobCodes = [];
  
  for (let i = 0; i < data.length; i++) {
    const code = data[i][0];       
    const englishJob = data[i][1]; 
    const frenchJob = data[i][2];  
    const family = data[i][3];     
    
    if (!code || code.toString().trim() === '') continue; 
    
    jobCodes.push({
      code: String(code).trim(),
      name: String(frenchJob || englishJob).trim(), 
      family: String(family || 'Autres').trim()
    });
  }
  
  return jobCodes;
}

/**
 * NOUVEAU : Récupère les données BU & Départements
 */
function getBUData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('BU&Département');
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 
  
  // Lecture des colonnes A à D
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const result = [];
  
  for (let i = 0; i < data.length; i++) {
    // Sécurisation totale des variables
    const colA = String(data[i][0] || '').trim(); // Code Département
    const colB = String(data[i][1] || '').trim(); // Département
    const colC = String(data[i][2] || '').trim(); // Code BU
    const colD = String(data[i][3] || '').trim(); // BU
    
    // On ne garde la ligne que si au moins l'une des colonnes contient du texte
    // Cela élimine les "lignes fantômes" générées par l'IMPORTRANGE
    if (colA !== '' || colB !== '' || colC !== '' || colD !== '') {
      result.push({
        codeDept: colA,
        dept: colB,
        codeBu: colC,
        bu: colD
      });
    }
  }
  
  return result;
}
// =========================================================================
// GESTION DE LA PROGRESSION UTILISATEUR ET DE L'XP
// =========================================================================

/**
 * MODIFIÉ : Enregistre une action spécifique dans la Matrice_Suivi de manière dynamique
 * en ajoutant des colonnes si de nouvelles clés apparaissent.
 */
function logUserProgress(email, taskKey, dateStr) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Log dans l'historique complet (Suivi_Formation)
  let trackingSheet = ss.getSheetByName('Suivi_Formation');
  if (!trackingSheet) {
    trackingSheet = ss.insertSheet('Suivi_Formation');
    trackingSheet.appendRow(['Date et Heure', 'Email', 'Tâche', 'Détails']);
    trackingSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f4f6");
  }
  trackingSheet.appendRow([new Date(), email, taskKey, dateStr]);

  // 2. Log dans la Matrice de progression (Matrice_Suivi)
  let sheet = ss.getSheetByName('Matrice_Suivi');
  if (!sheet) {
    sheet = ss.insertSheet('Matrice_Suivi');
    sheet.appendRow(['Email']); 
    sheet.getRange(1, 1, 1, 1).setFontWeight("bold").setBackground("#f3f4f6");
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Trouve la colonne de cette tâche, ou crée-la si elle n'existe pas
  let colIndex = headers.indexOf(taskKey);
  if (colIndex === -1) {
    colIndex = headers.length;
    sheet.getRange(1, colIndex + 1).setValue(taskKey);
    sheet.getRange(1, colIndex + 1).setFontWeight("bold").setBackground("#f3f4f6");
  }

  // Trouve l'utilisateur, ou crée-le
  let userRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      userRow = i + 1;
      break;
    }
  }

  if (userRow === -1) {
    sheet.appendRow([email]);
    userRow = sheet.getLastRow();
  }

  // Écrit la date de validation
  sheet.getRange(userRow, colIndex + 1).setValue(dateStr);
}

/**
 * MODIFIÉ : Récupère la progression sous forme de dictionnaire d'objets pour React
 */
function getUserProgress(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Matrice_Suivi');
  
  let progressData = {};
  if (!sheet) return progressData;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      for (let j = 1; j < headers.length; j++) {
        if (data[i][j] && data[i][j] !== '') {
          // On associe la clé d'en-tête à la valeur de la cellule (date)
          progressData[headers[j]] = String(data[i][j]);
        }
      }
      break;
    }
  }
  
  return progressData;
}

/**
 * NOUVEAU : Réinitialise la progression de l'utilisateur (Bouton Reset)
 */
function resetUserProgress(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Matrice_Suivi');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      // Efface toutes les colonnes après l'Email (Colonne 1) pour cet utilisateur
      const numCols = sheet.getLastColumn();
      if(numCols > 1) {
        sheet.getRange(i + 1, 2, 1, numCols - 1).clearContent();
      }
      break;
    }
  }
}

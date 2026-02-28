// ==========================================
// WILDU BULK WA - CORE ENGINE (FASE 6.0 - LUCCHETTO GEO & MACROAREE)
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Wildu WA Bulk')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initApp() {
  var props = PropertiesService.getUserProperties();
  var dbId = props.getProperty('BULK_DB_ID');
  var ss; var isNew = false;
  
  if (dbId) { try { ss = SpreadsheetApp.openById(dbId); } catch(e) { dbId = null; } }

  if (!dbId) {
    isNew = true;
    ss = SpreadsheetApp.create("Wildu_Bulk_Master_DB");
    props.setProperty('BULK_DB_ID', ss.getId());
    
    var configTab = ss.getSheets()[0]; configTab.setName("CONFIG");
    configTab.appendRow(["CHIAVE", "VALORE"]); configTab.appendRow(["DELAY_MIN", "30"]); configTab.appendRow(["DELAY_MAX", "80"]); configTab.appendRow(["PAUSA_BLOCCO", "25"]); configTab.appendRow(["MINUTI_PAUSA", "15"]); configTab.appendRow(["LIMITE_GIORNO", "80"]);
    configTab.getRange("A1:B1").setFontWeight("bold").setBackground("#d0e0e3");

    var sourcesTab = ss.insertSheet("SOURCES");
    sourcesTab.appendRow(["ID_FILE_META", "NOME_FOGLIO", "COL_NOME", "COL_TEL", "COL_TAGS", "RIGA_START", "ETICHETTA_SORGENTE", "COL_EMAIL", "COL_REGIONE", "SKIP_GEO"]);
    sourcesTab.getRange("A1:J1").setFontWeight("bold").setBackground("#fff2cc");

    var shadowTab = ss.insertSheet("SHADOW_LEADS");
    shadowTab.appendRow(["NOME", "TELEFONO", "SORGENTE", "STATO_INVIO", "DATA_IMPORT", "NOTE", "FUNNEL", "TAGS", "EMAIL", "REGIONE"]);
    shadowTab.getRange("A1:J1").setFontWeight("bold").setBackground("#d9ead3");
    
    var dictTab = ss.insertSheet("DICT_GEO"); dictTab.appendRow(["LOCALITA", "REGIONE"]); dictTab.getRange("A1:B1").setFontWeight("bold").setBackground("#fce5cd");
  }

  checkAndUpgradeSchema(ss);
  return { success: true, dbUrl: ss.getUrl(), dbId: ss.getId() };
}

function checkAndUpgradeSchema(ss) {
  var shadowTab = ss.getSheetByName("SHADOW_LEADS");
  if (shadowTab) {
    var lastCol = shadowTab.getLastColumn();
    if (lastCol > 0) {
      var headers = shadowTab.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf("FUNNEL") === -1) { lastCol++; shadowTab.getRange(1, lastCol).setValue("FUNNEL").setFontWeight("bold").setBackground("#d9ead3"); }
      if (headers.indexOf("TAGS") === -1) { lastCol++; shadowTab.getRange(1, lastCol).setValue("TAGS").setFontWeight("bold").setBackground("#d9ead3"); }
      if (headers.indexOf("EMAIL") === -1) { lastCol++; shadowTab.getRange(1, lastCol).setValue("EMAIL").setFontWeight("bold").setBackground("#d9ead3"); }
      if (headers.indexOf("REGIONE") === -1) { lastCol++; shadowTab.getRange(1, lastCol).setValue("REGIONE").setFontWeight("bold").setBackground("#d9ead3"); }
    }
  }
  var sourcesTab = ss.getSheetByName("SOURCES");
  if (sourcesTab) {
    var expectedHeaders = ["ID_FILE_META", "NOME_FOGLIO", "COL_NOME", "COL_TEL", "COL_TAGS", "RIGA_START", "ETICHETTA_SORGENTE", "COL_EMAIL", "COL_REGIONE", "SKIP_GEO", "COL_ETA", "COL_DESIDERI"];
    var lastCol = sourcesTab.getLastColumn();
    if (lastCol === 0) {
      sourcesTab.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]).setFontWeight("bold").setBackground("#fff2cc");
    } else {
      var currentHeaders = sourcesTab.getRange(1, 1, 1, lastCol).getValues()[0];
      expectedHeaders.forEach(function(h) {
        if (currentHeaders.indexOf(h) === -1) {
          lastCol++;
          // Forza la conversione a stringa per evitare errori di tipo se l'header è un array
          sourcesTab.getRange(1, lastCol).setValue(String(h)).setFontWeight("bold").setBackground("#fff2cc");
          currentHeaders.push(h);
        }
      });
    }
  }
  if (!ss.getSheetByName("DICT_GEO")) { var dictTab = ss.insertSheet("DICT_GEO"); dictTab.appendRow(["LOCALITA", "REGIONE"]); dictTab.getRange("A1:B1").setFontWeight("bold").setBackground("#fce5cd"); }
  
  // AGGIORNAMENTO SCHEMA: FOGLIO CODA DI INVIO
  if (!ss.getSheetByName("QUEUE_SENDER")) { 
    var queueTab = ss.insertSheet("QUEUE_SENDER"); 
    queueTab.appendRow(["TELEFONO", "NOME", "MESSAGGIO", "DATA_PROGRAMMATA", "STATO_CODA", "IMMAGINE"]); 
    queueTab.getRange("A1:F1").setFontWeight("bold").setBackground("#f4cccc"); 
  }

  // AGGIORNAMENTO SCHEMA: COLONNA ETA' E DESIDERI NEL DB OMBRA
  var shadowTab = ss.getSheetByName("SHADOW_LEADS");
  if (shadowTab) {
    var header = shadowTab.getRange(1, 1, 1, shadowTab.getLastColumn()).getValues()[0];
    if (header.indexOf("ANNO_NASCITA") === -1) {
      shadowTab.getRange(1, shadowTab.getLastColumn() + 1).setValue("ANNO_NASCITA").setFontWeight("bold").setBackground("#d9ead3");
    }
    if (header.indexOf("DESIDERI") === -1) {
      shadowTab.getRange(1, shadowTab.getLastColumn() + 1).setValue("DESIDERI").setFontWeight("bold").setBackground("#d9ead3");
    }
    if (header.indexOf("STORICO_DESIDERI") === -1) {
      shadowTab.getRange(1, shadowTab.getLastColumn() + 1).setValue("STORICO_DESIDERI").setFontWeight("bold").setBackground("#d9ead3");
    }
  }
  
  // AGGIORNAMENTO SCHEMA: MACROAREE MULTIPLE
  var macroTab = ss.getSheetByName("MACROAREE");
  if (!macroTab) {
    macroTab = ss.insertSheet("MACROAREE"); macroTab.appendRow(["MACROAREA", "REGIONI"]); macroTab.getRange("A1:B1").setFontWeight("bold").setBackground("#cfe2f3");
    var defaultMacro = [
      ["NORD", "EMILIA-ROMAGNA, FRIULI-VENEZIA GIULIA, LIGURIA, LOMBARDIA, PIEMONTE, TRENTINO-ALTO ADIGE, VALLE D'AOSTA, VENETO"],
      ["CENTRO", "LAZIO, MARCHE, TOSCANA, UMBRIA"],
      ["SUD", "ABRUZZO, BASILICATA, CALABRIA, CAMPANIA, MOLISE, PUGLIA, SARDEGNA, SICILIA"],
      ["VICINI WILDU", "TOSCANA"]
    ];
    macroTab.getRange(2, 1, defaultMacro.length, 2).setValues(defaultMacro);
  } else {
    // Pialla il vecchio schema a 20 righe per evitare conflitti e imposta il nuovo
    if (macroTab.getRange("A1").getValue() === "REGIONE") {
      macroTab.clear(); macroTab.appendRow(["MACROAREA", "REGIONI"]); macroTab.getRange("A1:B1").setFontWeight("bold").setBackground("#cfe2f3");
      var defaultMacro = [ ["NORD", "EMILIA-ROMAGNA, FRIULI-VENEZIA GIULIA, LIGURIA, LOMBARDIA, PIEMONTE, TRENTINO-ALTO ADIGE, VALLE D'AOSTA, VENETO"], ["CENTRO", "LAZIO, MARCHE, TOSCANA, UMBRIA"], ["SUD", "ABRUZZO, BASILICATA, CALABRIA, CAMPANIA, MOLISE, PUGLIA, SARDEGNA, SICILIA"], ["VICINI WILDU", "TOSCANA"] ];
      macroTab.getRange(2, 1, defaultMacro.length, 2).setValues(defaultMacro);
    }
  }
}

const GEO_TRANSLATIONS = {
  "FLORENCE": "FIRENZE", "VENICE": "VENEZIA", "MILAN": "MILANO", "ROME": "ROMA", 
  "NAPLES": "NAPOLI", "TURIN": "TORINO", "GENOA": "GENOVA", "PADUA": "PADOVA", 
  "SYRACUSE": "SIRACUSA", "MANTUA": "MANTOVA", "LEGHORN": "LIVORNO",
  "APULIA": "PUGLIA", "SICILY": "SICILIA", "TUSCANY": "TOSCANA", 
  "SARDINIA": "SARDEGNA", "LOMBARDY": "LOMBARDIA", "PIEDMONT": "PIEMONTE", 
  "LATIUM": "LAZIO", "AOSTA VALLEY": "VALLE D'AOSTA", "TRENTINO SOUTH TYROL": "TRENTINO-ALTO ADIGE",
  "EMILIA ROMAGNA": "EMILIA-ROMAGNA", "EMIGLIA": "EMILIA-ROMAGNA", 
  "TRENTINO": "TRENTINO-ALTO ADIGE", "ALTO ADIGE": "TRENTINO-ALTO ADIGE", 
  "FRIULI": "FRIULI-VENEZIA GIULIA", "VAL D AOSTA": "VALLE D'AOSTA", "VAL DAOSTA": "VALLE D'AOSTA"
};

function extractSheetId(input) { if (!input) return null; var match = input.match(/\/d\/([a-zA-Z0-9-_]+)/); return match ? match[1] : input; }

function fetchExternalSheetInfo(urlOrId) {
  try {
    var id = extractSheetId(urlOrId); if (!id) throw new Error("Link o ID non valido.");
    var ss = SpreadsheetApp.openById(id); var sheets = ss.getSheets(); var sheetNames = []; var headersMap = {};
    sheets.forEach(function(s) {
      var name = s.getName(); sheetNames.push(name); var lastCol = s.getLastColumn(); var lastRow = s.getLastRow();
      var rowsToFetch = Math.min(lastRow, 10);
      if (lastCol > 0 && rowsToFetch > 0) {
        var rawData = s.getRange(1, 1, rowsToFetch, lastCol).getValues();
        var stringData = rawData.map(function(row) { return row.map(function(cell) { return cell ? String(cell).trim() : ""; }); });
        headersMap[name] = stringData;
      } else { headersMap[name] = []; }
    });
    return { success: true, id: id, name: ss.getName(), tabs: sheetNames, headersMap: headersMap };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// SALVATAGGIO CONFIGURAZIONI SORGENTI
// ==========================================
function saveNewSource(cfg) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
    var ss = SpreadsheetApp.openById(dbId);
    var tab = ss.getSheetByName("SOURCES");
    
    // Sicurezza: Crea l'intestazione della colonna Età (Colonna K) se non esiste
    var headers = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0];
    if (headers.indexOf("COL_ETA") === -1) {
      tab.getRange(1, headers.length + 1).setValue("COL_ETA").setFontWeight("bold").setBackground("#d9ead3");
    }

    var skipStr = cfg.skipGeo ? "SI" : "NO";
    var newRow = [
      cfg.id, 
      cfg.tabName, 
      cfg.colNomeString, 
      cfg.colTelString, 
      cfg.colTagsString, 
      cfg.headerRow, 
      cfg.label, 
      cfg.colEmailString, 
      cfg.colRegioneString, 
      skipStr, 
      cfg.colEtaString || "",
      cfg.colDesideriString || ""  // <-- AGGIUNTO IL PONTE PER I DESIDERI
    ];
    
    if(cfg.editIndex !== "") {
      var rowIdx = parseInt(cfg.editIndex) + 2;
      tab.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);
    } else {
      tab.appendRow(newRow);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function removeSourceAt(rowIndex) { try { var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); SpreadsheetApp.openById(dbId).getSheetByName("SOURCES").deleteRow(rowIndex + 2); return { success: true }; } catch(e) { return { success: false, error: e.message }; } }

function getSavedSources() { var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); if (!dbId) return []; return SpreadsheetApp.openById(dbId).getSheetByName("SOURCES").getDataRange().getValues().slice(1); }

function downloadComuniItaliani() {
  try {
    var url = "https://raw.githubusercontent.com/matteocontrini/comuni-json/master/comuni.json"; var response = UrlFetchApp.fetch(url); var comuniData = JSON.parse(response.getContentText());
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); var dictTab = SpreadsheetApp.openById(dbId).getSheetByName("DICT_GEO");
    dictTab.clear(); var outputData = [["LOCALITA", "REGIONE"]];
    var regioni = ["Abruzzo","Basilicata","Calabria","Campania","Emilia-Romagna","Friuli-Venezia Giulia","Lazio","Liguria","Lombardia","Marche","Molise","Piemonte","Puglia","Sardegna","Sicilia","Toscana","Trentino-Alto Adige","Umbria","Valle d'Aosta","Veneto"];
    regioni.forEach(function(r) { outputData.push([r.toUpperCase(), r]); });
    comuniData.forEach(function(c) {
      var nome = String(c.nome).toUpperCase(); var regione = c.regione.nome; outputData.push([nome, regione]);
      if(c.sigla) { outputData.push([String(c.sigla).toUpperCase(), regione]); }
    });
    dictTab.getRange(1, 1, outputData.length, 2).setValues(outputData);
    return { success: true, count: outputData.length - 1 };
  } catch(e) { return { success: false, error: e.message }; }
}

// --- MOTORE 1: SINCRO-SCAN (CON FILTRO SELETTIVO E TURBO SMART SCAN) ---
function runSincroScan(targetLabel, isSmartMode) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); 
    
    // Tracciamento per il report finale
    var sourcesProcessed = [];
    var isSelective = (targetLabel && targetLabel !== "ALL");
    var smart = !!isSmartMode; // Forza a booleano 
    var ss = SpreadsheetApp.openById(dbId);
    var shadowTab = ss.getSheetByName("SHADOW_LEADS"); 
    var sourcesTab = ss.getSheetByName("SOURCES"); 
    var shadowData = shadowTab.getDataRange().getValues(); 
    var shadowHeaders = shadowData[0];
    
    var iNome = shadowHeaders.indexOf("NOME"), iTel = shadowHeaders.indexOf("TELEFONO"), iSorg = shadowHeaders.indexOf("SORGENTE"), iStato = shadowHeaders.indexOf("STATO_INVIO"), iData = shadowHeaders.indexOf("DATA_IMPORT"), iNote = shadowHeaders.indexOf("NOTE"), iFunnel = shadowHeaders.indexOf("FUNNEL"), iTags = shadowHeaders.indexOf("TAGS"), iEmail = shadowHeaders.indexOf("EMAIL"), iRegione = shadowHeaders.indexOf("REGIONE");
    var iAnnoNascita = shadowHeaders.indexOf("ANNO_NASCITA");
    var iDesideri = shadowHeaders.indexOf("DESIDERI");
    var iStorico = shadowHeaders.indexOf("STORICO_DESIDERI");
    
    var leadsMap = {}; var leadsOrder = [];
    
    // Inizializzazione Log Array
    var logDetails = [];

    for(var i = 1; i < shadowData.length; i++) { 
      var phoneKey = String(shadowData[i][iTel]).replace(/[^0-9+]/g, ''); if(!phoneKey) continue; 
      var row = shadowData[i];
      var regField = String(row[iRegione]);
      
      if (regField && regField !== "N/D" && regField.indexOf("(Tentativo") === -1 && regField.indexOf("(Fallito") === -1) {
        var cleanEx = regField; var tagEx = "";
        if(regField.indexOf("(Auto:") > -1) { var spl = regField.split("(Auto:"); cleanEx = spl[0].trim(); tagEx = " (Auto:" + spl[1]; }
        var upEx = cleanEx.toUpperCase();
        if(upEx.indexOf("EMILIA") > -1) cleanEx = "Emilia-Romagna";
        else if(upEx.indexOf("ADIGE") > -1 || upEx.indexOf("TRENTINO") > -1) cleanEx = "Trentino-Alto Adige";
        else if(upEx.indexOf("FRIULI") > -1) cleanEx = "Friuli-Venezia Giulia";
        else if(upEx.indexOf("AOSTA") > -1) cleanEx = "Valle d'Aosta";
        row[iRegione] = cleanEx + tagEx;
      }
      leadsMap[phoneKey] = row; leadsOrder.push(phoneKey); 
    }

    var sourcesList = sourcesTab.getDataRange().getValues(); 
    var countNew = 0, countUpdated = 0; 
    var today = new Date().toLocaleDateString('it-IT');
    var authGeoKeywords = ["dove abita", "regione", "città", "citta"];

    for(var s = 1; s < sourcesList.length; s++) {
      var srcId = sourcesList[s][0], srcTabName = sourcesList[s][1], strNome = sourcesList[s][2], strTel = sourcesList[s][3], strTagsArr = sourcesList[s][4] ? sourcesList[s][4].split(',') : [], headRow = parseInt(sourcesList[s][5]) - 1, srcLabel = sourcesList[s][6], strEmail = sourcesList[s][7], strRegioneMapped = sourcesList[s][8], skipGeo = sourcesList[s][9] === "SI", strEtaMapped = sourcesList[s][10], strDesideriMapped = sourcesList[s][11];

      // 🛡️ ROAD 2.5: Se è un sync selettivo, controlla se la sorgente è tra quelle scelte
      if (isSelective) {
        var targets = targetLabel.split(',').map(function(t) { return t.trim(); });
        if (targets.indexOf(srcLabel) === -1) continue;
      }
      
      sourcesProcessed.push(srcLabel);

      try {
        var extSS = SpreadsheetApp.openById(srcId); var extTab = extSS.getSheetByName(srcTabName); var extData = extTab.getDataRange().getValues(); var extHeaders = extData[headRow];
        var extIdxNome = extHeaders.indexOf(strNome), extIdxTel = extHeaders.indexOf(strTel), extIdxEmail = strEmail ? extHeaders.indexOf(strEmail) : -1, extIdxRegioneMapped = strRegioneMapped ? extHeaders.indexOf(strRegioneMapped) : -1;
        var extIdxEtaMapped = strEtaMapped ? extHeaders.indexOf(strEtaMapped) : -1; 
        var extIdxDesideriMapped = strDesideriMapped ? extHeaders.indexOf(strDesideriMapped) : -1; 
        var extIdxTags = strTagsArr.map(function(t) { return extHeaders.indexOf(t.trim()); }).filter(function(idx) { return idx > -1; });
        var extIdxAuthGeo = []; extHeaders.forEach(function(h, idx) { var cleanHeader = String(h).toLowerCase(); for(var k=0; k < authGeoKeywords.length; k++) { if(cleanHeader.indexOf(authGeoKeywords[k]) > -1) { extIdxAuthGeo.push(idx); break; } } });
        
        if(extIdxNome === -1 || extIdxTel === -1) continue;

        for(var r = headRow + 1; r < extData.length; r++) {
          var rowRaw = extData[r]; var rawTel = String(rowRaw[extIdxTel]).replace(/[^0-9+]/g, ''); if(!rawTel) continue;

          // 🚀 TURBO SMART SCAN: Se il lead esiste già, salta tutta la logica di calcolo/update
          if (smart && leadsMap[rawTel]) continue;
          var rawNome = String(rowRaw[extIdxNome]).trim(), rawEmail = extIdxEmail > -1 ? String(rowRaw[extIdxEmail]).trim() : "";
          
          var rawEtaVal = extIdxEtaMapped > -1 ? String(rowRaw[extIdxEtaMapped]).trim() : "";
          var calculatedBirthYear = rawEtaVal ? parseBirthYear(rawEtaVal) : "";
          var rawDesideriVal = extIdxDesideriMapped > -1 ? String(rowRaw[extIdxDesideriMapped]).trim().toUpperCase() : "";

          var rawTags = []; extIdxTags.forEach(function(idx) { if(rowRaw[idx]) { rawTags.push(String(rowRaw[idx]).trim()); } });
          var authGeoInputs = []; extIdxAuthGeo.forEach(function(idx) { if(rowRaw[idx]) { authGeoInputs.push(String(rowRaw[idx]).trim()); } });
          var rawAuthGeoText = authGeoInputs.join(" "); var rawRegioneMappedVal = extIdxRegioneMapped > -1 ? String(rowRaw[extIdxRegioneMapped]).trim() : "";

          var isMappedValidRegion = false;
          if(rawRegioneMappedVal) {
            var upMap = rawRegioneMappedVal.toUpperCase();
            if(upMap.indexOf("EMILIA") > -1) { rawRegioneMappedVal = "Emilia-Romagna"; isMappedValidRegion = true; }
            else if(upMap.indexOf("TRENTINO") > -1 || upMap.indexOf("ADIGE") > -1) { rawRegioneMappedVal = "Trentino-Alto Adige"; isMappedValidRegion = true; }
            else if(upMap.indexOf("FRIULI") > -1) { rawRegioneMappedVal = "Friuli-Venezia Giulia"; isMappedValidRegion = true; }
            else if(upMap.indexOf("AOSTA") > -1) { rawRegioneMappedVal = "Valle d'Aosta"; isMappedValidRegion = true; }
            else {
              var validR = ["Abruzzo","Basilicata","Calabria","Campania","Lazio","Liguria","Lombardia","Marche","Molise","Piemonte","Puglia","Sardegna","Sicilia","Toscana","Umbria","Veneto"];
              for(var v=0; v<validR.length; v++) { if(upMap === validR[v].toUpperCase()) { rawRegioneMappedVal = validR[v]; isMappedValidRegion = true; break; } }
            }
          }

          var existingRow = leadsMap[rawTel] || null; 
          var existingReg = existingRow ? String(existingRow[iRegione]) : "N/D";
          if (existingReg === "") existingReg = "N/D";
          
          // --- SISTEMA ANTICORPI: GERARCHIA REGIONI ---
          var existingScore = 0;
          if (existingReg !== "N/D" && existingReg.indexOf("(Auto") === -1 && existingReg.indexOf("(Tentativo") === -1 && existingReg.indexOf("(Fallito") === -1 && existingReg.indexOf("(Mappato)") === -1) {
            existingScore = 4; // Qualità Massima: Inserita Manualmente
          } else if (existingReg.indexOf("(Mappato)") > -1) {
            existingScore = 3; // Qualità Alta: Trovata in un modulo Meta
          } else if (existingReg.indexOf("(Auto:") > -1) {
            existingScore = 2; // Qualità Media: Risolta dal Bot Geo
          } else if (existingReg.indexOf("(Tentativo:") > -1 || existingReg.indexOf("(Fallito:") > -1) {
            existingScore = 1; // Qualità Bassa: Parola in attesa di calcolo
          }
          
          var newRegione = "N/D";
          var newScore = 0;
          
          if (rawRegioneMappedVal && isMappedValidRegion) {
            newRegione = rawRegioneMappedVal + " (Mappato)";
            newScore = 3;
          } else if (rawRegioneMappedVal && !isMappedValidRegion) {
            newRegione = "N/D (Tentativo: " + rawRegioneMappedVal + ")";
            newScore = 1;
          } else if (rawAuthGeoText && !skipGeo) {
            newRegione = "N/D (Tentativo: " + rawAuthGeoText + ")";
            newScore = 1;
          }
          
          // Decide chi vince: aggiorna SOLO se il nuovo dato è qualitativamente MIGLIORE (o se prima non c'era nulla)
          var finalRegione = existingReg;
          if (newScore > existingScore) {
            finalRegione = newRegione;
          }
          // ---------------------------------------------

          if(existingRow) {
            var changed = false;
            var logsRiga = []; // Tiene traccia di cosa cambia su QUESTA riga

            if(existingRow[iSorg].indexOf(srcLabel) === -1) { 
              existingRow[iSorg] += " | " + srcLabel; 
              logsRiga.push("+ Nuova Sorgente (" + srcLabel + ")");
              
              var wasArchived = false;
              if(existingRow[iTags] && existingRow[iTags].indexOf("[ARCHIVIATO]") > -1) {
                wasArchived = true;
                existingRow[iTags] = existingRow[iTags].replace(/\[ARCHIVIATO\]/g, "").replace(/,\s*,/g, ",").replace(/(^,)|(,$)/g, "").trim();
              }
              
              var currentFunnel = String(existingRow[iFunnel]).toUpperCase().trim();
              var vecchiStati = ["CHIUSO", "NON RISPONDE"]; 
              
              if(wasArchived || vecchiStati.indexOf(currentFunnel) > -1 || currentFunnel === "") {
                existingRow[iFunnel] = "RITORNO";
                logsRiga.push("+ Risveglio / Funnel in RITORNO");
              }
              changed = true; 
            }
            if(!existingRow[iEmail] && rawEmail) { existingRow[iEmail] = rawEmail; logsRiga.push("+ Email aggiunta"); changed = true; }
            if(existingRow[iRegione] !== finalRegione) { existingRow[iRegione] = finalRegione; logsRiga.push("+ Update Regione"); changed = true; }
            
            if(iAnnoNascita > -1 && !existingRow[iAnnoNascita] && calculatedBirthYear) {
              existingRow[iAnnoNascita] = calculatedBirthYear;
              logsRiga.push("+ Anno Nascita");
              changed = true;
            }
            if(iDesideri > -1 && !existingRow[iDesideri] && rawDesideriVal) {
              existingRow[iDesideri] = rawDesideriVal;
              logsRiga.push("+ Desideri");
              changed = true;
            }
            
            // Fix controllo Tag per evitare i falsi positivi
            var curTagsStr = existingRow[iTags] ? String(existingRow[iTags]) : "";
            var curTags = curTagsStr ? curTagsStr.split(",").map(function(t) { return t.trim().toUpperCase() }).filter(function(t) { return t; }) : [];
            var tagsAggiunti = [];

            rawTags.forEach(function(rawT) {
              if(rawT) {
                var subTags = String(rawT).split(",").map(function(st) { return st.trim(); });
                subTags.forEach(function(t) {
                  if(t && curTags.indexOf(t.toUpperCase()) === -1) {
                    curTags.push(t.toUpperCase()); 
                    tagsAggiunti.push(t);
                  }
                });
              }
            });
            
            if (tagsAggiunti.length > 0) {
              existingRow[iTags] = curTags.join(", ");
              logsRiga.push("+ Tags: " + tagsAggiunti.join(", "));
              changed = true;
            }

            if(changed) {
              countUpdated++;
              logDetails.push("🔄 AGGIORNATO: " + rawNome + " (" + rawTel + ") -> " + logsRiga.join(" | "));
            }
            
          } else {
            // Nuova riga
            var nr = new Array(shadowHeaders.length).fill("");
            nr[iNome] = rawNome; nr[iTel] = rawTel; nr[iSorg] = srcLabel; nr[iStato] = "DA CONTATTARE"; nr[iData] = today; nr[iNote] = ""; nr[iFunnel] = "NUOVO"; nr[iTags] = rawTags.join(", "); nr[iEmail] = rawEmail; nr[iRegione] = finalRegione;
            if (iAnnoNascita > -1) nr[iAnnoNascita] = calculatedBirthYear; 
            if (iDesideri > -1) nr[iDesideri] = rawDesideriVal;
            if (iStorico > -1) nr[iStorico] = "";
            leadsMap[rawTel] = nr; leadsOrder.push(rawTel); 
            countNew++;
            logDetails.push("🟢 NUOVO: " + rawNome + " (" + rawTel + ") dalla sorgente " + srcLabel);
          }
        }
      } catch(e) { 
        logDetails.push("❌ ERRORE LETTURA FOGLIO: " + srcLabel + " -> " + e.message);
        continue; 
      }
    }

    var outData = [shadowHeaders]; leadsOrder.forEach(function(p) { outData.push(leadsMap[p]); });
    shadowTab.clearContents().getRange(1, 1, outData.length, outData[0].length).setValues(outData);
    
    // Ritorna il report log al frontend con info sorgenti e modalità
    var modeTag = smart ? "🚀 TURBO (Solo Nuovi)" : "🧬 PROFONDO (Aggiorna Esistenti)";
    var summarySorgenti = (isSelective ? "🎯 SYNC MIRATO" : "🌍 SYNC TOTALE") + " | " + modeTag;
    
    logDetails.unshift("========================================");
    logDetails.unshift(summarySorgenti);
    logDetails.unshift("========================================");
    
    return { success: true, newCount: countNew, updatedCount: countUpdated, logDetails: logDetails };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

// --- MOTORE 2: ANALISI GEO BATCH (VERSIONE SMART & RE-TRY) ---
function runGeoAnalysisBatch() {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); 
    var ss = SpreadsheetApp.openById(dbId);
    var shadowTab = ss.getSheetByName("SHADOW_LEADS"); 
    var dictTab = ss.getSheetByName("DICT_GEO");
    
    var data = shadowTab.getDataRange().getValues(); 
    var headers = data[0]; 
    
    // Indici per leggere e aggiornare (Aggiunti Nome e Tel per il Log)
    var iRegione = headers.indexOf("REGIONE");
    var iNome = headers.indexOf("NOME");
    var iTel = headers.indexOf("TELEFONO");
    
    var geoDict = []; 
    if(dictTab && dictTab.getLastRow() > 1) { 
      geoDict = dictTab.getDataRange().getValues().slice(1); 
      // Ordina il dizionario dalla parola più lunga alla più corta
      geoDict.sort(function(a, b) { return b[0].length - a[0].length; }); 
    }

    var processedInThisBatch = 0; 
    var maxBatchSize = 400; 
    var remaining = 0; 
    var updates = [];
    
    // --- VARIABILI PER IL REPORT E I LOG ---
    var logDetails = [];
    var updatedCount = 0;

    for(var i = 1; i < data.length; i++) {
      var reg = String(data[i][iRegione]);
      
      // Cattura nome e telefono per il report
      var nomeLead = iNome > -1 ? String(data[i][iNome]) : "Sconosciuto";
      var telLead = iTel > -1 ? String(data[i][iTel]) : "";

      // 🧠 LOGICA SMART: Prova ad analizzare sia i "Tentativi" che i "Falliti" precedenti
      if(reg.indexOf("N/D (Tentativo:") === 0 || reg.indexOf("N/D (Fallito:") === 0) {
        if(processedInThisBatch < maxBatchSize) {
          
          // Usa la funzione centralizzata di estrazione parola (Road 2)
          var rawText = extractGeoKeyword(reg); 
          
          if(rawText) {
            var cleanGeo = rawText; 
            
            // Applica le traduzioni manuali (se presenti)
            if(typeof GEO_TRANSLATIONS !== 'undefined') {
              for(var eng in GEO_TRANSLATIONS) { 
                if(cleanGeo.indexOf(eng) > -1) { cleanGeo += " " + GEO_TRANSLATIONS[eng]; } 
              }
            }
            
            var found = false; 
            var finalReg = "";
            
            for(var d = 0; d < geoDict.length; d++) {
              var localita = geoDict[d][0]; 
              var regex = new RegExp("\\b" + localita + "\\b", "gi");
              
              if(regex.test(cleanGeo)) { 
                finalReg = geoDict[d][1] + " (Auto: " + rawText + ")"; 
                found = true; 
                
                // 🟢 SUCCESSO: Registra nel Log
                updatedCount++;
                logDetails.push("🟢 ASSEGNATO: " + nomeLead + " (" + telLead + ") | Testo: '" + rawText + "' -> " + geoDict[d][1]);
                break; 
              }
            }
            
            if(!found) { 
              finalReg = "N/D (Fallito: " + rawText + ")"; 
              // 🟡 FALLIMENTO: Registra nel Log
              logDetails.push("🟡 FALLITO: " + nomeLead + " (" + telLead + ") | Nessun match nel DB per: '" + rawText + "'");
            } 
            
            // Aggiunge l'aggiornamento alla lista (Solo se è cambiato qualcosa)
            if (reg !== finalReg) {
              updates.push({row: i + 1, col: iRegione + 1, val: finalReg});
            }
          }
          processedInThisBatch++;
        } else { 
          remaining++; 
        }
      }
    }
    
    // Scrive tutti gli aggiornamenti sul foglio in colpo solo
    if(updates.length > 0) { 
      updates.forEach(function(u) { shadowTab.getRange(u.row, u.col).setValue(u.val); }); 
    }
    
    return { 
      success: true, 
      processed: processedInThisBatch, 
      remaining: remaining,
      logDetails: logDetails,
      updatedCount: updatedCount
    };
    
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function updateLeadFull(phoneKey, fieldName, newValue) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); var shadowTab = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
    var data = shadowTab.getDataRange().getValues(); var headers = data[0]; var colIdx = headers.indexOf(fieldName); var telIdx = headers.indexOf("TELEFONO");
    if(colIdx === -1 || telIdx === -1) throw new Error("Colonna non trovata");
    for(var i = 1; i < data.length; i++) { if(String(data[i][telIdx]) === String(phoneKey)) { shadowTab.getRange(i + 1, colIdx + 1).setValue(newValue); return { success: true }; } }
    return { success: false, error: "Contatto non trovato" };
  } catch(e) { return { success: false, error: e.message }; }
}

// --- ROAD 1 + ROAD 2: TURBO-SAVE CON AUTO-APPRENDIMENTO ---
function updateLeadTurbo(phoneKey, payload) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); 
    var shadowTab = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
    var data = shadowTab.getDataRange().getValues(); 
    var headers = data[0]; 
    var telIdx = headers.indexOf("TELEFONO");
    var regIdx = headers.indexOf("REGIONE"); // Aggiunto per leggere il vecchio valore
    
    if(telIdx === -1) throw new Error("Colonna TELEFONO non trovata");
    
    var learningQueue = []; // Coda per memorizzare le nuove parole
    
    for(var i = 1; i < data.length; i++) { 
      if(String(data[i][telIdx]) === String(phoneKey)) { 
        
        // 🧠 INNESCO GEO-INTELLIGENCE: Intercetta il dato prima di sovrascriverlo
        if (payload["REGIONE"] && regIdx > -1) {
          var oldReg = String(data[i][regIdx]);
          var newReg = payload["REGIONE"];
          var keyword = extractGeoKeyword(oldReg); // Estrae la parola (se era un tentativo)
          if (keyword && newReg && newReg !== "N/D") {
            learningQueue.push({keyword: keyword, regione: newReg});
          }
        }

        // Apre la valigetta e aggiorna solo i campi presenti
        for (var fieldName in payload) {
          var colIdx = headers.indexOf(fieldName);
          if (colIdx !== -1) {
            shadowTab.getRange(i + 1, colIdx + 1).setValue(payload[fieldName]);
          }
        }
        
        // Se ha scoperto qualcosa, lo manda al dizionario prima di chiudere
        if (learningQueue.length > 0) {
          processGeoLearning(learningQueue, dbId);
        }
        
        return { success: true }; 
      } 
    }
    return { success: false, error: "Contatto non trovato" };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function getConfigData() { var id = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); if(!id) return []; return SpreadsheetApp.openById(id).getSheetByName("CONFIG").getDataRange().getValues().slice(1); }
// --- TURBO-SAVE CONFIGURAZIONI (Ottimizzato) ---
function updateAppConfigs(newConfigs) { 
  try { 
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); 
    var configTab = SpreadsheetApp.openById(dbId).getSheetByName("CONFIG"); 
    var data = configTab.getDataRange().getValues(); 
    var changed = false;

    // 1. Lavora tutto istantaneamente nella memoria RAM (Array)
    for (var i = 1; i < data.length; i++) { 
      var key = data[i][0]; 
      if (newConfigs[key] !== undefined) { 
        data[i][1] = newConfigs[key]; // Aggiorna il valore nell'array
        delete newConfigs[key]; 
        changed = true;
      } 
    } 

    // 2. Unica scrittura massiva per aggiornare le vecchie configurazioni
    if (changed) {
      configTab.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    // 3. Gestione nuove chiavi (nel caso in futuro ne aggiungessimo altre)
    var newRows = [];
    for (var newKey in newConfigs) { 
      if(newConfigs[newKey] !== undefined) {
        newRows.push([newKey, newConfigs[newKey]]); 
      }
    } 
    if (newRows.length > 0) {
      configTab.getRange(configTab.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    }

    return { success: true }; 
  } catch (e) { 
    return { success: false, error: e.message }; 
  } 
}

function getMacroareeConfig() { var id = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); if(!id) return []; return SpreadsheetApp.openById(id).getSheetByName("MACROAREE").getDataRange().getValues().slice(1); }
function updateMacroareeConfig(macroList) { try { var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); var macroTab = SpreadsheetApp.openById(dbId).getSheetByName("MACROAREE"); macroTab.clearContents(); macroTab.appendRow(["MACROAREA", "REGIONI"]); if(macroList.length > 0) { macroTab.getRange(2, 1, macroList.length, 2).setValues(macroList); } return { success: true }; } catch (e) { return { success: false, error: e.message }; } }

function getShadowLeads() { var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); if (!dbId) return []; return SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS").getDataRange().getDisplayValues(); }

// ==========================================
// MODULO 2.0: CAMPAGNE, SEGMENTAZIONE E CODE
// ==========================================

function calculateCampaignTarget(filters) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID'); 
    var ss = SpreadsheetApp.openById(dbId);
    var shadowTab = ss.getSheetByName("SHADOW_LEADS");
    var data = shadowTab.getDataRange().getValues();
    var headers = data[0];
    
    // Novità: Legge le soglie d'età e le Macroaree dal foglio CONFIG/MACROAREE
    var th1 = 17, th2 = 30, th3 = 45, th4 = 54;
    var macroareeMap = {};
    var configTab = ss.getSheetByName("CONFIG");
    if(configTab) {
      var confData = configTab.getDataRange().getValues();
      for(var c=0; c<confData.length; c++){
        if(confData[c][0] === "AGE_TH_1" && confData[c][1]) th1 = parseInt(confData[c][1]);
        if(confData[c][0] === "AGE_TH_2" && confData[c][1]) th2 = parseInt(confData[c][1]);
        if(confData[c][0] === "AGE_TH_3" && confData[c][1]) th3 = parseInt(confData[c][1]);
        if(confData[c][0] === "AGE_TH_4" && confData[c][1]) th4 = parseInt(confData[c][1]);
      }
    }
    var macroTab = ss.getSheetByName("MACROAREE");
    if(macroTab) {
      var mData = macroTab.getDataRange().getValues();
      for(var m=1; m<mData.length; m++) { macroareeMap[mData[m][0]] = String(mData[m][1]).toUpperCase(); }
    }
    
    var iTel = headers.indexOf("TELEFONO"), iNome = headers.indexOf("NOME"), iReg = headers.indexOf("REGIONE");
    var iTags = headers.indexOf("TAGS"), iStato = headers.indexOf("STATO_INVIO"), iAnno = headers.indexOf("ANNO_NASCITA");
    
    var validLeads = [];
    var oggi = new Date();
    var currentYear = oggi.getFullYear(); 
    
    var calderone = filters.calderone || [];
    var requisito = filters.requisito || [];
    var esclusione = filters.esclusione || [];
    var scudoGiorni = parseInt(filters.scudoGiorni) || 0;
    var soloNuovi = filters.soloNuovi || false;
    
    for(var i = 1; i < data.length; i++) {
      var row = data[i];
      var tel = String(row[iTel]); if(!tel) continue;
      
      var reg = String(row[iReg]).toUpperCase().trim();
      if (reg.indexOf("(AUTO:") > -1) reg = reg.split("(AUTO:")[0].trim();
      if (reg === "") reg = "N/D";
      var tags = String(row[iTags]).toUpperCase();
      
      // --- 🛡️ SCUDO ANTI-ARCHIVIATI ---
      // Se il lead ha il tag archiviato, salta al prossimo. Niente coda WA!
      if (tags.indexOf("[ARCHIVIATO]") > -1) continue;
      // ---------------------------------
      
      var stato = String(row[iStato]).toUpperCase();

      // INIEZIONE MACROAREA (Server-side) - FIX EFFETTO SPUGNA
      var macroList = [];
      if (reg !== "N/D") {
        for (var macroKey in macroareeMap) { if (macroareeMap[macroKey].indexOf(reg) > -1) macroList.push(macroKey); }
      }
      var leadMacros = macroList.join(" ");
      
      var anno = iAnno > -1 ? row[iAnno] : "";
      var ageGroup = "N/D";
      if (anno && !isNaN(anno)) {
        var currentAge = currentYear - parseInt(anno);
        if(currentAge <= th1) ageGroup = "CHILD";
        else if(currentAge <= th2) ageGroup = "YOUNG";
        else if(currentAge <= th3) ageGroup = "ADULT";
        else if(currentAge <= th4) ageGroup = "MAN";
        else ageGroup = "OLD";
      }
      
      // La stringa combinata PERFETTA (stessa dell'HTML)
      var combinedData = reg + " " + tags + " AGE_" + ageGroup + " " + leadMacros; 
      
      // --- 1. SCUDO TEMPORALE E SEGMENTO SMART ---
      var isNuovo = stato === "DA CONTATTARE" || stato === "";
      if (soloNuovi && !isNuovo) continue; 
      
      if (!isNuovo && scudoGiorni > 0 && stato.indexOf("CONTATTATO IL") > -1) {
        var dataMatch = stato.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (dataMatch) {
          var dataContatto = new Date(dataMatch[3], parseInt(dataMatch[2])-1, dataMatch[1]);
          var diffGiorni = Math.floor((oggi - dataContatto) / (1000 * 60 * 60 * 24));
          if (diffGiorni <= scudoGiorni) continue;
        }
      }
      
      // --- 2. IL CALDERONE (OR) ---
      var inCalderone = false;
      if (calderone.length === 0 || calderone.indexOf("TUTTI") > -1) {
        inCalderone = true;
      } else {
        for(var c=0; c<calderone.length; c++) {
          if(combinedData.indexOf(calderone[c].toUpperCase()) > -1) { inCalderone = true; break; }
        }
      }
      if (!inCalderone) continue; 
      
      // --- 3. REQUISITO RIGIDO (AND) ---
      var haRequisiti = true;
      if (requisito.length > 0) {
        for(var r=0; r<requisito.length; r++) {
          if(combinedData.indexOf(requisito[r].toUpperCase()) === -1) { haRequisiti = false; break; }
        }
      }
      if (!haRequisiti) continue; 
      
      // --- 4. ESCLUSIONE (NOT) ---
      var daEscludere = false;
      if (esclusione.length > 0) {
        for(var e=0; e<esclusione.length; e++) {
          if(combinedData.indexOf(esclusione[e].toUpperCase()) > -1) { daEscludere = true; break; }
        }
      }
      if (daEscludere) continue; 
      
      validLeads.push({ telefono: tel, nome: String(row[iNome]) });
    }
    
    return { success: true, count: validLeads.length, leads: validLeads };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// ==========================================
// UPLOAD IMMAGINI SU GOOGLE DRIVE
// ==========================================
function uploadImageToDrive(dataURI, filename) {
  try {
    var folderName = "WA_BULK_IMAGES";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var type = dataURI.split(';')[0].split(':')[1];
    var base64 = dataURI.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), type, filename);
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Genera il link di download diretto per WhatsApp
    var url = "https://drive.google.com/uc?export=download&id=" + file.getId();
    return { success: true, url: url };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// GENERATORE CODA (SPINTAX & SCHEDULING)
// ==========================================
function buildCampaignQueue(leads, msgBlocks, desTattoo) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
    var ss = SpreadsheetApp.openById(dbId);
    var queueTab = ss.getSheetByName("QUEUE_SENDER");
    var configTab = ss.getSheetByName("CONFIG");

    // Assicurati che ci sia la colonna IMMAGINE nell'intestazione
    if (queueTab.getRange("F1").getValue() !== "IMMAGINE") {
      queueTab.getRange("F1").setValue("IMMAGINE").setFontWeight("bold").setBackground("#f4cccc");
    }

    // Legge il Limite Giornaliero per calcolare i giorni
    var limit = 80;
    var confData = configTab.getDataRange().getValues();
    for(var c=0; c<confData.length; c++) { if(confData[c][0] === "LIMITE_GIORNO" && confData[c][1]) limit = parseInt(confData[c][1]); }

    var outData = [];
    var today = new Date(); // Oggi

    for(var i = 0; i < leads.length; i++) {
      var lead = leads[i];
      
      // MOTORE SPINTAX: Pesca a caso una frase per ogni blocco
      var hook = msgBlocks.hooks.length > 0 ? msgBlocks.hooks[Math.floor(Math.random() * msgBlocks.hooks.length)] : "";
      var middle = msgBlocks.middles.length > 0 ? msgBlocks.middles[Math.floor(Math.random() * msgBlocks.middles.length)] : "";
      var ending = msgBlocks.endings.length > 0 ? msgBlocks.endings[Math.floor(Math.random() * msgBlocks.endings.length)] : "";

      // Assembla il messaggio saltando le parti vuote
      var msgParts = [];
      if(hook) msgParts.push(hook);
      if(msgBlocks.body) msgParts.push(msgBlocks.body);
      if(middle) msgParts.push(middle);
      if(ending) msgParts.push(ending);

      var finalMsg = msgParts.join("\n\n");
      
      // Magia: Sostituisce il tag {NOME} con il vero nome del contatto
      finalMsg = finalMsg.replace(/\{NOME\}/gi, lead.nome);

      // CALCOLO DATA ANTI-BAN: Spalma i messaggi nei giorni successivi
      var dayOffset = Math.floor(i / limit); // Es. se limite è 80, il lead 81 avrà offset 1 (Domani)
      var sendDate = new Date(today.getTime() + (dayOffset * 24 * 60 * 60 * 1000));
      var dateStr = Utilities.formatDate(sendDate, "Europe/Rome", "yyyy-MM-dd");

      // Aggiunge la riga alla coda
      outData.push([lead.telefono, lead.nome, finalMsg, dateStr, "DA INVIARE", msgBlocks.image || ""]);
    }

    // Scrive tutto in un colpo solo per massima velocità
    if(outData.length > 0) {
      queueTab.getRange(queueTab.getLastRow() + 1, 1, outData.length, 6).setValues(outData);
      
      // IL TATUAGGIO STORICO DESIDERI
      if (desTattoo) {
        var sData = ss.getSheetByName("SHADOW_LEADS").getDataRange().getValues();
        var hStorico = sData[0].indexOf("STORICO_DESIDERI");
        var hTel = sData[0].indexOf("TELEFONO");
        if (hStorico > -1) {
          var leadPhones = leads.map(function(l) { return String(l.telefono); });
          var tattooStr = "[TENTATIVO " + desTattoo.toUpperCase() + ": " + today.toLocaleDateString('it-IT') + "]";
          for(var r=1; r<sData.length; r++) {
             if (leadPhones.indexOf(String(sData[r][hTel])) > -1) {
                var oldVal = sData[r][hStorico] ? sData[r][hStorico] + " | " : "";
                ss.getSheetByName("SHADOW_LEADS").getRange(r+1, hStorico+1).setValue(oldVal + tattooStr);
             }
          }
        }
      }
    }

    return { success: true, count: outData.length };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// ELIMINAZIONE MANUALE LEAD (UNSUBSCRIBE/TEST)
// ==========================================
function deleteLeadManual(telefono) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
    var sheet = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(telefono)) { // <-- ECCO IL FIX: [1] è la colonna TELEFONO!
        sheet.deleteRow(i + 1); 
        return { success: true };
      }
    }
    return { success: false, error: "Contatto non trovato nel database." };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// PARSER INTELLIGENTE ETA E DATA DI NASCITA
// ==========================================
function parseBirthYear(rawValue) {
  if (!rawValue) return "";
  var val = String(rawValue).trim();
  if (val === "") return "";

  var currentYear = new Date().getFullYear();

  // --- DOGANA PRIORITARIA: ECCEZIONI STATICHE (Trattamento Fasce Testuali) ---
  // Gestisce i 400 lead con fasce pre-impostate evitando calcoli errati o scarti.
  if (val.indexOf("29 -") > -1 || val.indexOf("29-") > -1) return 1999; 
  if (val.indexOf("30/51") > -1) return 1986;
  if (val.indexOf("52/66") > -1) return 1966;
  if (val.indexOf("66+") > -1 || val.indexOf("66 +") > -1) return 1955;

  // PRE-FILTRO: L'intuizione del "900" (Auto-correzione refusi)
  // Resta intatto: se non è una delle fasce sopra, controlla se manca l'1 davanti (es. 929)
  var typoMatch = val.match(/\b(9\d{2})\b/);
  if (typoMatch) {
    val = val.replace(/\b9\d{2}\b/, "1" + typoMatch[1]);
  }

  // CASO 1: Anno di nascita esplicito a 4 cifre
  var yearMatch = val.match(/\b(19[0-9]{2}|20[0-9]{2})\b/);
  if (yearMatch) {
    var parsedYear = parseInt(yearMatch[1]);
    // FIX: Scarta il futuro (es. 2050) e scarta roba troppo vecchia (prima del 1900)
    if (parsedYear >= 1900 && parsedYear <= currentYear) {
      return parsedYear;
    }
  }

  // CASO 2: Età pura dichiarata (da 1 a 3 cifre isolate)
  // Resta intatto: calcola l'anno partendo dall'età (es. "33" -> 1993)
  var ageMatch = val.match(/\b(\d{1,3})\b/);
  if (ageMatch) {
    var num = parseInt(ageMatch[1]);
    if (num >= 14 && num <= 115) { 
      return currentYear - num;
    }
  }

  // Restituisce vuoto se non corrisponde a nessuno dei criteri sopra
  return "";
}


function RIPARA_SOURCES() {
  var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
  var tab = SpreadsheetApp.openById(dbId).getSheetByName("SOURCES");
  
  // Scrive le 12 intestazioni esatte e definitive
  var correctHeaders = [
    "ID_FILE_META", "NOME_FOGLIO", "COL_NOME", "COL_TEL", "COL_TAGS", 
    "RIGA_START", "ETICHETTA_SORGENTE", "COL_EMAIL", "COL_REGIONE", 
    "SKIP_GEO", "COL_ETA", "COL_DESIDERI"
  ];
  
  // Le stampa brutalmente sulla riga 1, colorandole
  tab.getRange(1, 1, 1, correctHeaders.length)
     .setValues([correctHeaders])
     .setFontWeight("bold")
     .setBackground("#fff2cc");
}

function PulisciTagDuplicati() {
  var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
  var sheet = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iTags = headers.indexOf("TAGS");
  
  var updates = [];
  for (var i = 1; i < data.length; i++) {
    var tagStr = String(data[i][iTags] || "");
    if (tagStr) {
      // Divide per virgola, toglie spazi, mette in maiuscolo e rimuove i duplicati veri
      var cleanArray = tagStr.split(",").map(t => t.trim()).filter((v, i, a) => v && a.indexOf(v) === i);
      var newStr = cleanArray.join(", ");
      if (newStr !== tagStr) {
        sheet.getRange(i + 1, iTags + 1).setValue(newStr);
      }
    }
  }
  Logger.log("Pulizia completata!");
}

function runSourceCleanup(sourceIndex, mode) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
    var ss = SpreadsheetApp.openById(dbId);
    var sourcesTab = ss.getSheetByName("SOURCES");
    var shadowTab = ss.getSheetByName("SHADOW_LEADS");
    
    // 1. Recupera configurazione sorgente
    var sourceCfg = sourcesTab.getRange(sourceIndex + 2, 1, 1, 12).getValues()[0];
    var srcId = sourceCfg[0], srcTabName = sourceCfg[1], srcTelCol = sourceCfg[3], 
        srcTagsArr = sourceCfg[4] ? sourceCfg[4].split(',') : [], headRow = parseInt(sourceCfg[5]) - 1, 
        srcLabel = sourceCfg[6];

    // 2. Legge il Foglio Meta originale (Mappa Storica)
    var extSS = SpreadsheetApp.openById(srcId);
    var extTab = extSS.getSheetByName(srcTabName);
    var extData = extTab.getDataRange().getValues();
    var extHeaders = extData[headRow];
    
    // Mappatura colonne esterne
    var extIdxTel = extHeaders.indexOf(srcTelCol);
    var extIdxTags = srcTagsArr.map(h => extHeaders.indexOf(h.trim())).filter(i => i > -1);

    // 3. Crea Mappa dei dati da rimuovere (Telefono -> Lista Valori da cancellare)
    var removalMap = {};
    for(var r = headRow + 1; r < extData.length; r++) {
      var tel = String(extData[r][extIdxTel]).replace(/[^0-9+]/g, '');
      if(!tel) continue;
      var valuesToFind = [];
      extIdxTags.forEach(idx => { if(extData[r][idx]) valuesToFind.push(String(extData[r][idx]).trim().toUpperCase()); });
      removalMap[tel] = valuesToFind;
    }

    // 4. Scansione Database Ombra e Pulizia
    var shadowData = shadowTab.getDataRange().getValues();
    var shadowHeaders = shadowData[0];
    var iSorg = shadowHeaders.indexOf("SORGENTE"), iTags = shadowHeaders.indexOf("TAGS"), 
        iReg = shadowHeaders.indexOf("REGIONE"), iTel = shadowHeaders.indexOf("TELEFONO"),
        iEmail = shadowHeaders.indexOf("EMAIL"), iEta = shadowHeaders.indexOf("ANNO_NASCITA"),
        iDes = shadowHeaders.indexOf("DESIDERI");

    var countCleaned = 0;
    for (var i = 1; i < shadowData.length; i++) {
      var row = shadowData[i];
      var currentSorg = String(row[iSorg]);
      
      // Filtro: Agiamo solo se il lead appartiene a questa sorgente
      if (currentSorg.indexOf(srcLabel) > -1) {
        var phoneKey = String(row[iTel]).replace(/[^0-9+]/g, '');
        var changed = false;

        // --- PULIZIA TAGS (Strada B) ---
        if ((mode === 'TAGS' || mode === 'ALL') && removalMap[phoneKey]) {
          var leadTags = row[iTags] ? row[iTags].split(',').map(t => t.trim()) : [];
          var tagsToRemove = removalMap[phoneKey];
          
          var newTags = leadTags.filter(t => tagsToRemove.indexOf(t.toUpperCase()) === -1);
          if (newTags.length !== leadTags.length) {
            row[iTags] = newTags.join(', ');
            changed = true;
          }
        }

        // --- PULIZIA REGIONE/EMAIL/ETA (Solo se con firma Bot) ---
        if (mode === 'GEO' || mode === 'ALL') {
          // Regione: solo se contiene (Auto: o (Tentativo:
          var regVal = String(row[iReg]);
          if (mode === 'GEO' || mode === 'ALL') {
          var regVal = String(row[iReg]);
          if (regVal.indexOf('(Auto:') > -1 || regVal.indexOf('(Tentativo:') > -1 || regVal.indexOf('(Mappato)') > -1 || regVal === "N/D") {
            row[iReg] = ""; changed = true;
          }
            row[iReg] = ""; changed = true;
          }
          // Email: se vogliamo resettarla (qui decidiamo se è rischioso, per ora la lasciamo se non specificato)
          // Anno Nascita e Desideri (Zona Volatile)
          row[iEta] = ""; row[iDes] = ""; changed = true;
        }

        if (changed) {
          shadowTab.getRange(i + 1, 1, 1, shadowHeaders.length).setValues([row]);
          countCleaned++;
        }
      }
    }
    return { success: true, count: countCleaned };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


function ApplicaTatuaggioMappatoRetroattivo() {
  var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
  var sheet = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iReg = headers.indexOf("REGIONE");

  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var reg = String(data[i][iReg]).trim();
    
    // Se la cella non è vuota, non è N/D e non ha già parentesi (Auto, Mappato, Tentativo)...
    if (reg !== "" && reg !== "N/D" && reg.indexOf("(") === -1) {
      // Aggiunge il tatuaggio di Meta
      sheet.getRange(i + 1, iReg + 1).setValue(reg + " (Mappato)");
      count++;
    }
  }
  Logger.log("Operazione completata! Tatuaggio (Mappato) applicato a " + count + " lead.");
}

// ==========================================
// WILDU CRM - MOTORE AZIONI MASSIVE CON GEO-LEARNING
// ==========================================
function processBulkAction(phones, actionType, newValue) {
  try {
    var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
    var ss = SpreadsheetApp.openById(dbId);
    var shadowTab = ss.getSheetByName("SHADOW_LEADS");
    var data = shadowTab.getDataRange().getValues();
    var headers = data[0];
    
    var iTel = headers.indexOf("TELEFONO");
    var iFunnel = headers.indexOf("FUNNEL");
    var iReg = headers.indexOf("REGIONE");
    var iTags = headers.indexOf("TAGS");
    
    if(iTel === -1) throw new Error("Colonna TELEFONO non trovata nel Foglio Ombra.");

    var count = 0;
    var learningQueue = []; // 🧠 Coda per l'apprendimento massivo
    
    // --- SCENARIO 1: ELIMINAZIONE DEFINITIVA (Metodo Array Veloce) ---
    if (actionType === 'DELETE') {
      var newData = [headers]; // Inizializza con le intestazioni
      for(var i = 1; i < data.length; i++) {
        var t = String(data[i][iTel]).replace(/[^0-9+]/g, '');
        if (phones.indexOf(t) === -1) {
          newData.push(data[i]); // Mantieni solo chi NON è nella lista
        } else {
          count++; // Lead eliminato
        }
      }
      // Sovrascrive il foglio con l'array pulito (Fulmineo e sicuro)
      shadowTab.clearContents().getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      return { success: true, count: count };
    }

    // --- SCENARIO 2: AGGIORNAMENTI (Funnel, Regione, Archiviazione) ---
    var changed = false;
    for(var i = 1; i < data.length; i++) {
      var t = String(data[i][iTel]).replace(/[^0-9+]/g, '');
      
      if (phones.indexOf(t) > -1) { // Se il telefono è tra quelli selezionati
        
        if (actionType === 'UPDATE_FUNNEL' && iFunnel > -1) {
          data[i][iFunnel] = newValue;
          changed = true; count++;
        } 
        else if (actionType === 'UPDATE_REGIONE' && iReg > -1) {
          // 🧠 INNESCO GEO-INTELLIGENCE (BULK): Legge prima di sovrascrivere
          var oldReg = String(data[i][iReg]);
          var keyword = extractGeoKeyword(oldReg);
          if (keyword && newValue && newValue !== "N/D") {
            learningQueue.push({keyword: keyword, regione: newValue});
          }
          
          data[i][iReg] = newValue;
          changed = true; count++;
        }
        else if (actionType === 'ADD_TAGS' && iTags > -1) {
          // Separa i tag vecchi e i nuovi, mettendoli in ordine
          var currentTagsArr = data[i][iTags] ? String(data[i][iTags]).split(',').map(function(t){return t.trim().toUpperCase();}).filter(function(t){return t;}) : [];
          var newTagsArr = newValue.split(',').map(function(t){return t.trim().toUpperCase();}).filter(function(t){return t;});
          
          // Unisce senza creare cloni (se uno ha già VIP, non lo mette due volte)
          newTagsArr.forEach(function(nt) {
            if (currentTagsArr.indexOf(nt) === -1) currentTagsArr.push(nt);
          });
          
          data[i][iTags] = currentTagsArr.join(', ');
          changed = true; count++;
        }
        else if (actionType === 'REMOVE_TAGS' && iTags > -1) {
          if (data[i][iTags]) {
            var currentTagsArr = String(data[i][iTags]).split(',').map(function(t){return t.trim().toUpperCase();}).filter(function(t){return t;});
            var tagsToRemoveArr = newValue.split(',').map(function(t){return t.trim().toUpperCase();}).filter(function(t){return t;});
            
            // Filtra e tiene solo i tag che NON sono nella lista di quelli da rimuovere
            var finalTagsArr = currentTagsArr.filter(function(t) {
              return tagsToRemoveArr.indexOf(t) === -1;
            });
            
            data[i][iTags] = finalTagsArr.join(', ');
            changed = true; count++;
          }
        }
        else if (actionType === 'ARCHIVE' && iTags > -1) {
          var currentTags = String(data[i][iTags]).trim();
          if (currentTags.indexOf("[ARCHIVIATO]") === -1) {
            data[i][iTags] = currentTags ? currentTags + ", [ARCHIVIATO]" : "[ARCHIVIATO]";
            changed = true; count++;
          }
        }
        else if (actionType === 'RESTORE' && iTags > -1) {
          var currentTags = String(data[i][iTags]);
          if (currentTags.indexOf("[ARCHIVIATO]") > -1) {
            // 1. Rimuove il tag e pulisce eventuali virgole orfane
            data[i][iTags] = currentTags.replace(/\[ARCHIVIATO\]/g, "").replace(/,\s*,/g, ",").replace(/(^,)|(,$)/g, "").trim();
            
            // 2. Forza lo status Funnel su "RITORNO"
            if (iFunnel > -1) {
              data[i][iFunnel] = "RITORNO";
            }
            
            changed = true; count++;
          }
        }
      }
    }

    // Scrive le modifiche in un colpo solo (Batch)
    if (changed) {
      shadowTab.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
    
    // 🧠 Avvia il salvataggio in dizionario per le parole raccolte (Se ce ne sono)
    if (learningQueue.length > 0) {
      processGeoLearning(learningQueue, dbId);
    }

    return { success: true, count: count };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// UTILITY: PULIZIA RETROATTIVA ANNI IMPOSSIBILI
// ==========================================
function PulisciAnniFuturi() {
  var dbId = PropertiesService.getUserProperties().getProperty('BULK_DB_ID');
  var sheet = SpreadsheetApp.openById(dbId).getSheetByName("SHADOW_LEADS");
  var data = sheet.getDataRange().getValues();
  var iAnno = data[0].indexOf("ANNO_NASCITA");
  var currentY = new Date().getFullYear();
  var count = 0;
  
  for(var i=1; i<data.length; i++) {
     var y = parseInt(data[i][iAnno]);
     if(!isNaN(y) && (y > currentY || y < 1900)) {
        sheet.getRange(i+1, iAnno+1).setValue(""); // Svuota la cella corrotta
        count++;
     }
  }
  Logger.log("Pulizia completata: rimossi " + count + " anni anomali dal database.");
}

// ==========================================
// INTELLIGENZA GEO-BOT (AUTO-APPRENDIMENTO) - V2 CORAZZATA
// ==========================================
function extractGeoKeyword(regStr) {
  if (!regStr) return null;
  
  // 1. Cerca la parola dopo i due punti e si ferma alla parentesi chiusa
  var match = String(regStr).match(/(?:Tentativo|Fallito):\s*([^)]+)/i);
  
  if (match && match[1]) {
    var keyword = match[1].toUpperCase();
    
    // 🛡️ TRITACARNE: Rimuove attivamente "N/D", "FALLITO" e "TENTATIVO" se sono finiti dentro per errore
    keyword = keyword.replace(/N\/D/g, "")
                     .replace(/FALLITO/g, "")
                     .replace(/TENTATIVO/g, "");
                     
    // 2. Pulisce simboli strani tenendo solo lettere, numeri e accenti
    keyword = keyword.replace(/[^A-Z0-9\s\u00C0-\u017F]/g, "").trim();
    
    // 3. Rimuove eventuali doppi spazi creati dalla pulizia
    keyword = keyword.replace(/\s+/g, " ");
    
    if (keyword.length > 0) {
      return keyword;
    }
  }
  return null;
}

function processGeoLearning(learningQueue, dbId) {
  if (!learningQueue || learningQueue.length === 0) return;
  
  try {
    var ss = SpreadsheetApp.openById(dbId);
    var dictTab = ss.getSheetByName("DICT_GEO");
    if (!dictTab) return;
    
    var existingData = dictTab.getDataRange().getValues();
    var existingKeywords = {};
    
    // Crea una mappa veloce delle parole che il bot sa già per non duplicarle
    for(var i=1; i<existingData.length; i++) {
      if(existingData[i][0]) {
        existingKeywords[String(existingData[i][0]).toUpperCase()] = true;
      }
    }
    
    var newRows = [];
    learningQueue.forEach(function(item) {
      var kw = item.keyword;
      var reg = item.regione;
      
      // Salva solo se la parola è nuova e la regione è valida
      if (kw && reg && reg !== "N/D" && reg !== "" && !existingKeywords[kw]) {
        newRows.push([kw, reg]);
        existingKeywords[kw] = true; 
      }
    });
    
    // Scrittura massiva sul foglio DICT_GEO
    if (newRows.length > 0) {
      dictTab.getRange(dictTab.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
      Logger.log("🧠 Geo-Bot ha imparato " + newRows.length + " nuove località.");
    }
  } catch(e) {
    Logger.log("Errore Geo-Learning: " + e.message);
  }
}



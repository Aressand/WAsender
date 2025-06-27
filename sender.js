/**
 * Debug: mostra cosa vede il sistema
 */
function debugDatiClienti() {
  console.log('🔍 DEBUG DATI CLIENTI');
  console.log('===================');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    console.log('❌ Foglio "Clienti" non trovato!');
    console.log('Fogli disponibili:');
    SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(s => {
      console.log(`  - "${s.getName()}"`);
    });
    return;
  }
  
  console.log('✅ Foglio "Clienti" trovato');
  
  const data = sheet.getDataRange().getValues();
  console.log(`Righe totali: ${data.length}`);
  
  // Mostra headers
  console.log('\n📋 HEADERS (riga 1):');
  data[0].forEach((header, index) => {
    console.log(`  Colonna ${String.fromCharCode(65 + index)} (${index}): "${header}"`);
  });
  
  // Mostra prime 5 righe di dati
  console.log('\n📊 PRIME 5 RIGHE DI DATI:');
  for (let i = 1; i < Math.min(6, data.length); i++) {
    console.log(`\nRiga ${i + 1}:`);
    console.log(`  Nome: "${data[i][0]}"`);
    console.log(`  Telefono: "${data[i][2]}"`);
    console.log(`  Template (col I/8): "${data[i][8]}"`);
    console.log(`  Stato WhatsApp (col J/9): "${data[i][9]}" (tipo: ${typeof data[i][9]})`);
    console.log(`  Lunghezza stato: ${data[i][9] ? data[i][9].length : 0}`);
    
    // Controlla se c'è match
    if (data[i][9] === 'Da Inviare') {
      console.log('  ✅ MATCH: Da Inviare');
    } else {
      console.log('  ❌ NO MATCH');
      // Mostra caratteri nascosti
      if (data[i][9]) {
        console.log(`  Codici caratteri: ${Array.from(String(data[i][9])).map(c => c.charCodeAt(0)).join(', ')}`);
      }
    }
  }
  
  // Conta stati
  console.log('\n📈 CONTEGGIO STATI (colonna J):');
  const stati = {};
  for (let i = 1; i < data.length; i++) {
    const stato = data[i][9];  // Colonna J
    stati[stato] = (stati[stato] || 0) + 1;
  }
  Object.entries(stati).forEach(([stato, count]) => {
    console.log(`  "${stato}": ${count}`);
  });
}

// ====================================
// NOTE FINALI E ISTRUZIONI
// ====================================

/**
 * ISTRUZIONI PER LA CONFIGURAZIONE:
 * 
 * STRUTTURA FOGLIO "Clienti":
 * A: Nome
 * B: Cognome
 * C: Telefono
 * D: (vuota)
 * E: Data Chiamata
 * F: Esito
 * G: PDV
 * H: Operatore
 * I: Template (numero 1-5)
 * J: Stato WhatsApp ("Da Inviare", "Inviato", "Errore")
 * K: Data Invio
 * L: Contatto Processato
 * 
 * 1. PRIMA ESECUZIONE:
 *    - Esegui "setupIniziale()" per aggiungere la colonna "Contatto Processato"
 *    - Inserisci la tua API Key WASender in CONFIG.WASENDER_API_KEY
 * 
 * 2. GOOGLE CONTACTS (opzionale):
 *    - Menu Apps Script → Servizi (icona +)
 *    - Cerca "People API" 
 *    - Clicca "Aggiungi" e autorizza
 * 
 * 3. TRIGGER AUTOMATICO:
 *    - Menu Apps Script → Trigger (icona orologio)
 *    - "+ Aggiungi trigger"
 *    - Funzione: processaClientiCompleto
 *    - Tipo: Timer → Ogni 10 minuti
 * 
 * 4. LIMITI (opzionale):
 *    - CONFIG.LIMITE_GIORNALIERO = 50 (per max 50 msg/giorno)
 *    - CONFIG.ORA_INIZIO / ORA_FINE per orario operativo
 * 
 * 5. USO QUOTIDIANO:
 *    - Operatori inseriscono dati con stato "Da Inviare"
 *    - Sistema elabora automaticamente ogni 10 minuti (9-15)
 *    - Menu → "Elabora ora" per invio immediato
 *    - Menu → "Report giornaliero" per statistiche
 */

// Fine del codice// ====================================
// SETUP E MIGRAZIONE
// ====================================

/**
 * Setup iniziale - aggiunge colonna "Contatto Processato"
 */
function setupIniziale() {
  console.log('🔧 SETUP INIZIALE');
  console.log('================');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    console.log('❌ Foglio "Clienti" non trovato!');
    return;
  }
  
  // Verifica se la colonna L esiste già
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (headers[11] === 'Contatto Processato') {
    console.log('✅ Colonna "Contatto Processato" già presente');
    return;
  }
  
  // Aggiungi header in colonna L
  console.log('📋 Aggiunta colonna "Contatto Processato"...');
  sheet.getRange(1, 12).setValue('Contatto Processato');
  sheet.getRange(1, 12).setFontWeight('bold').setBackground('#f0f0f0');
  
  // Imposta larghezza colonna
  sheet.setColumnWidth(12, 130);
  
  console.log('✅ Setup completato!');
  console.log('');
  console.log('📝 PROSSIMI PASSI:');
  console.log('1. Configura CONFIG.WASENDER_API_KEY nel codice');
  console.log('2. Se vuoi creare contatti automaticamente:');
  console.log('   - Vai su Servizi → Aggiungi "People API"');
  console.log('3. Imposta trigger ogni 10 minuti per "processaClientiCompleto"');
  console.log('4. (Opzionale) Imposta CONFIG.LIMITE_GIORNALIERO se necessario');
}

/**
 * Reset colonna contatto processato
 */
function resetContattoProcessato() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // Pulisci colonna L (tranne header)
    sheet.getRange(2, 12, lastRow - 1, 1).clearContent();
    console.log(`✅ Reset completato per ${lastRow - 1} righe`);
  }
}// ====================================
// GESTIONE CONTATTI GOOGLE
// ====================================

/**
 * Gestisce creazione/ricerca contatto Google
 * @returns {string} 'CREATO' | 'ESISTENTE' | 'ERRORE'
 */
function gestisciContattoGoogle(datiCliente) {
  try {
    const telefono = formattaNumeroItalia(datiCliente.telefono);
    
    // Cerca contatto esistente
    const risultatiRicerca = People.People.searchContacts({
      query: telefono,
      readMask: 'names,phoneNumbers'
    });
    
    if (risultatiRicerca.results && risultatiRicerca.results.length > 0) {
      // Contatto già esiste
      return 'ESISTENTE';
    }
    
    // Crea nuovo contatto
    const nuovoContatto = {
      names: [{
        givenName: datiCliente.nome || '',
        familyName: datiCliente.cognome || ''
      }],
      phoneNumbers: [{
        value: telefono,
        type: 'mobile'
      }]
    };
    
    // Aggiungi PDV se presente
    if (datiCliente.pdv) {
      nuovoContatto.organizations = [{
        name: datiCliente.pdv,
        title: 'Cliente'
      }];
    }
    
    // Crea il contatto
    People.People.createContact({
      resource: nuovoContatto
    });
    
    return 'CREATO';
    
  } catch (error) {
    console.error('Errore gestione contatto:', error);
    return 'ERRORE';
  }
}

/**
 * Verifica se Google People API è disponibile
 */
function verificaGooglePeopleAPI() {
  try {
    People.People.searchContacts({
      query: 'test',
      readMask: 'names',
      pageSize: 1
    });
    console.log('✅ Google People API disponibile');
    return true;
  } catch (error) {
    console.log('⚠️ Google People API non configurata o non disponibile');
    console.log('   I contatti non verranno creati automaticamente');
    return false;
  }
}// ====================================
// CONFIGURAZIONE SISTEMA
// ====================================
const CONFIG = {
  WASENDER_API_KEY: 'INSERISCI-LA-TUA-API-KEY-QUI',
  WASENDER_URL: 'https://wasenderapi.com/api/send-message',
  DELAY_TRA_INVII: 30000, // 30 secondi tra un invio e l'altro
  MAX_INVII_PER_SESSIONE: 20, // Massimo messaggi per sessione
  ORARIO_INIZIO: 9, // Ora inizio invii automatici
  ORARIO_FINE: 19   // Ora fine invii automatici
};

// ====================================
// FUNZIONI CORE
// ====================================

/**
 * Formatta numero italiano aggiungendo +39
 */
function formattaNumeroItalia(telefono) {
  if (!telefono) return null;
  
  let numero = telefono.toString().trim();
  
  // Se già ha il +39, restituiscilo così com'è
  if (numero.startsWith('+39')) {
    return numero;
  }
  
  // Altrimenti aggiungi +39
  return '+39' + numero;
}

/**
 * Ottiene template dal foglio Template
 */
function getTemplate(templateId) {
  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Template');
  
  if (!templateSheet) {
    throw new Error('Foglio "Template" non trovato!');
  }
  
  const templates = templateSheet.getDataRange().getValues();
  
  // Cerca il template per ID (assumendo che l'ID sia nella colonna A)
  for (let i = 1; i < templates.length; i++) {
    if (templates[i][0] == templateId) { 
      return {
        messaggio: templates[i][2] || '', // Colonna C = Messaggio
        immagineUrl: templates[i][3] || null  // Colonna D = URL Immagine (opzionale)
      };
    }
  }
  
  // Template di default se non trova l'ID
  return {
    messaggio: 'Ciao [nome], ti abbiamo chiamato. Quando possiamo richiamarti?',
    immagineUrl: null
  };
}

/**
 * Sostituisce gli shortcode nel messaggio
 */
function sostituisciShortcode(messaggio, datiCliente) {
  let messaggioFinale = messaggio;
  
  const sostituzioni = {
    '[nome]': datiCliente.nome || '',
    '[cognome]': datiCliente.cognome || '',
    '[pdv]': datiCliente.pdv || '',
    '[operatore]': datiCliente.operatore || '',
    '[esito]': datiCliente.esito || '',
    '[data]': Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy'),
    '[ora]': Utilities.formatDate(new Date(), 'GMT+1', 'HH:mm')
  };
  
  for (const [placeholder, valore] of Object.entries(sostituzioni)) {
    const regex = new RegExp(placeholder.replace(/[[\]]/g, '\\$&'), 'g');
    messaggioFinale = messaggioFinale.replace(regex, valore);
  }
  
  return messaggioFinale;
}

/**
 * Invia messaggio WhatsApp con eventuale immagine
 */
function inviaWhatsApp(telefono, messaggio, immagineUrl = null) {
  const numeroFormattato = formattaNumeroItalia(telefono);
  
  if (!numeroFormattato) {
    throw new Error('Numero di telefono non valido');
  }
  
  const payload = {
    'to': numeroFormattato,
    'text': messaggio
  };
  
  // Aggiungi immagine se presente
  if (immagineUrl && immagineUrl.trim() !== '') {
    payload.imageUrl = immagineUrl.trim();
  }
  
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + CONFIG.WASENDER_API_KEY,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(CONFIG.WASENDER_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error(`Errore API (${responseCode}): ${responseText}`);
  }
  
  const jsonResponse = JSON.parse(responseText);
  if (jsonResponse.success === false) {
    throw new Error(`Errore WASender: ${jsonResponse.message || 'Errore sconosciuto'}`);
  }
  
  return responseText;
}

// ====================================
// FUNZIONE PRINCIPALE
// ====================================

// ====================================
// PROCESSO UNIFICATO
// ====================================

/**
 * Processo principale unificato: crea contatti e invia messaggi
 */
function processaClientiCompleto() {
  // Controllo orario operativo
  if (!isOrarioOperativo()) {
    console.log('⏰ Fuori orario operativo (9-15). Elaborazione annullata.');
    return;
  }
  
  // Controllo limite giornaliero
  const messaggiOggi = contaMessaggiInviatiOggi();
  if (CONFIG.LIMITE_GIORNALIERO && messaggiOggi >= CONFIG.LIMITE_GIORNALIERO) {
    console.log(`🛑 Raggiunto limite giornaliero (${CONFIG.LIMITE_GIORNALIERO} messaggi)`);
    logOperazione(`Limite giornaliero raggiunto: ${messaggiOggi}/${CONFIG.LIMITE_GIORNALIERO}`, 'INFO');
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    throw new Error('Foglio "Clienti" non trovato!');
  }
  
  const data = sheet.getDataRange().getValues();
  let contattiProcessati = 0;
  let contattiCreati = 0;
  let contattiEsistenti = 0;
  let errori = 0;
  
  console.log('🚀 Inizio elaborazione completa...');
  console.log(`📅 ${Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss")}`);
  
  // Debug: mostra quante righe stiamo analizzando
  console.log(`📊 Analizzando ${data.length - 1} righe di dati...`);
  
  // Calcola limite effettivo per questa sessione
  let limiteSessione = CONFIG.MAX_INVII_PER_SESSIONE;
  if (CONFIG.LIMITE_GIORNALIERO) {
    const rimanenti = CONFIG.LIMITE_GIORNALIERO - messaggiOggi;
    limiteSessione = Math.min(limiteSessione, rimanenti);
  }
  
  // Processa righe
  for (let i = 1; i < data.length; i++) {
    if (contattiProcessati >= limiteSessione) {
      console.log('⚠️ Raggiunto limite per questa sessione');
      break;
    }
    
    const row = data[i];
    const datiCliente = {
      nome: row[0],           // A - Nome
      cognome: row[1],        // B - Cognome  
      telefono: row[2],       // C - Telefono
      // row[3] è vuota       // D - Vuota
      dataChiamata: row[4],   // E - Data Chiamata
      esito: row[5],          // F - Esito
      pdv: row[6],            // G - PDV
      operatore: row[7],      // H - Operatore
      templateId: row[8],     // I - Template
      statoWhatsApp: row[9],  // J - Stato WhatsApp
      dataInvio: row[10],     // K - Data Invio
      contattoProcessato: row[11] // L - Contatto Processato
    };
    
    // Debug per prima riga
    if (i === 1) {
      console.log(`\n🔍 Debug prima riga:`);
      console.log(`  Nome: "${datiCliente.nome}"`);
      console.log(`  Stato (col J): "${datiCliente.statoWhatsApp}"`);
      console.log(`  Telefono: "${datiCliente.telefono}"`);
      console.log(`  Template (col I): "${datiCliente.templateId}"`);
      console.log(`  Test: stato === 'Da Inviare'? ${datiCliente.statoWhatsApp === 'Da Inviare'}`);
    }
    
    // Processa solo se "Da Inviare" e con telefono
    if (datiCliente.statoWhatsApp === 'Da Inviare' && datiCliente.telefono) {
      try {
        console.log(`\n📱 Elaboro ${datiCliente.nome} ${datiCliente.cognome}...`);
        
        // FASE 1: Gestione contatto Google (se abilitato)
        if (typeof People !== 'undefined') {
          try {
            const risultatoContatto = gestisciContattoGoogle(datiCliente);
            if (risultatoContatto === 'CREATO') {
              contattiCreati++;
              console.log('   ✅ Contatto creato');
              // Attendi sincronizzazione
              Utilities.sleep(CONFIG.DELAY_DOPO_CONTATTO);
            } else {
              contattiEsistenti++;
              console.log('   ℹ️ Contatto già esistente');
            }
            
            // Aggiorna stato contatto
            sheet.getRange(i + 1, 12).setValue('Si');
            
          } catch (e) {
            console.log('   ⚠️ Errore contatto Google:', e.toString());
            sheet.getRange(i + 1, 12).setValue('Errore');
            // Continua comunque con l'invio
          }
        }
        
        // FASE 2: Invio WhatsApp
        const template = getTemplate(datiCliente.templateId || '1');
        const messaggioPersonalizzato = sostituisciShortcode(template.messaggio, datiCliente);
        const immagineFinale = template.immagineUrl;
        
        console.log('   📤 Invio WhatsApp...');
        inviaWhatsApp(datiCliente.telefono, messaggioPersonalizzato, immagineFinale);
        
        // Aggiorna stato
        sheet.getRange(i + 1, 10).setValue('Inviato');  // Colonna J
        sheet.getRange(i + 1, 11).setValue(new Date()); // Colonna K
        
        contattiProcessati++;
        console.log('   ✅ Completato!');
        
        // Pausa tra invii
        if (contattiProcessati < limiteSessione && i < data.length - 1) {
          Utilities.sleep(CONFIG.DELAY_TRA_INVII);
        }
        
      } catch (error) {
        console.error(`   ❌ Errore: ${error.toString()}`);
        
        sheet.getRange(i + 1, 10).setValue('Errore');  // Colonna J
        logErrore(
          formattaNumeroItalia(datiCliente.telefono),
          error.toString(),
          `${datiCliente.nome} ${datiCliente.cognome}`
        );
        
        errori++;
      }
    }
  }
  
  // Riepilogo
  console.log('\n📊 RIEPILOGO ELABORAZIONE');
  console.log('========================');
  console.log(`✅ Messaggi inviati: ${contattiProcessati}`);
  console.log(`📇 Contatti creati: ${contattiCreati}`);
  console.log(`ℹ️ Contatti esistenti: ${contattiEsistenti}`);
  console.log(`❌ Errori: ${errori}`);
  console.log(`📅 Completato: ${new Date().toLocaleString('it-IT')}`);
  
  if (CONFIG.LIMITE_GIORNALIERO) {
    console.log(`📈 Totale oggi: ${messaggiOggi + contattiProcessati}/${CONFIG.LIMITE_GIORNALIERO}`);
  }
  
  logOperazione(
    `Elaborazione: ${contattiProcessati} inviati, ${contattiCreati} contatti creati, ${errori} errori`,
    'INFO'
  );
}

/**
 * Controlla se siamo in orario operativo
 */
function isOrarioOperativo() {
  // Forza timezone italiano
  const now = new Date();
  const oraItalia = Utilities.formatDate(now, "Europe/Rome", "HH");
  const oraCorrente = parseInt(oraItalia);
  
  console.log(`Controllo orario: ${oraCorrente}:00 (Roma)`);
  
  return oraCorrente >= CONFIG.ORA_INIZIO && oraCorrente < CONFIG.ORA_FINE;
}

/**
 * Conta messaggi inviati oggi
 */
function contaMessaggiInviatiOggi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  const oggi = Utilities.formatDate(new Date(), "Europe/Rome", "yyyy-MM-dd");
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const dataInvio = data[i][10]; // Colonna K - Data Invio
    if (dataInvio) {
      const dataInvioStr = Utilities.formatDate(new Date(dataInvio), "Europe/Rome", "yyyy-MM-dd");
      if (dataInvioStr === oggi) {
        count++;
      }
    }
  }
  
  return count;
}

/**
 * Versione con controlli orario per automazione
 */
function processaNuoviClientiConControlli() {
  // Verifica orario
  const ora = new Date().getHours();
  if (ora < CONFIG.ORARIO_INIZIO || ora >= CONFIG.ORARIO_FINE) {
    console.log(`⏰ Fuori orario lavorativo (${CONFIG.ORARIO_INIZIO}-${CONFIG.ORARIO_FINE}). Elaborazione annullata.`);
    return;
  }
  
  // Esegui elaborazione
  processaNuoviClienti();
}

// ====================================
// FUNZIONI UTILITA'
// ====================================

/**
 * Crea/aggiorna foglio di log
 */
function logOperazione(messaggio, tipo = 'INFO', telefono = '', dettagliErrore = '') {
  let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  // Se non esiste il foglio Log, crealo con nuova struttura
  if (!logSheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.insertSheet('Log');
    logSheet.getRange(1, 1, 1, 5)
      .setValues([['Data/Ora', 'Tipo', 'Messaggio', 'Telefono', 'Dettagli Errore']])
      .setFontWeight('bold')
      .setBackground('#f0f0f0');
    
    // Imposta larghezza colonne
    logSheet.setColumnWidth(1, 150); // Data/Ora
    logSheet.setColumnWidth(2, 80);  // Tipo
    logSheet.setColumnWidth(3, 250); // Messaggio
    logSheet.setColumnWidth(4, 120); // Telefono
    logSheet.setColumnWidth(5, 300); // Dettagli Errore
  }
  
  // Aggiungi riga di log
  logSheet.appendRow([new Date(), tipo, messaggio, telefono, dettagliErrore]);
  
  // Se è un errore, colora la riga
  if (tipo === 'ERRORE') {
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1, 1, 5).setBackground('#ffe6e6');
  }
}

/**
 * Log specifico per errori WhatsApp
 */
function logErrore(telefono, errore, nomeCliente = '') {
  const messaggioErrore = `Errore invio WhatsApp a ${nomeCliente}`;
  logOperazione(messaggioErrore, 'ERRORE', telefono, errore);
}

/**
 * Mostra dashboard con statistiche
 */
function creaDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  const data = sheet.getDataRange().getValues();
  
  let stats = {
    totale: 0,
    daInviare: 0,
    inviati: 0,
    errori: 0,
    inviatOggi: 0
  };
  
  const oggi = new Date().toDateString();
  
  // Analizza dati (salta header)
  for (let i = 1; i < data.length; i++) {
    if (data[i][2]) { // Se c'è un telefono
      stats.totale++;
      
      const stato = data[i][8]; // Colonna I - Stato WhatsApp
      const dataInvio = data[i][9]; // Colonna J - Data Invio
      
      switch(stato) {
        case 'Da Inviare': stats.daInviare++; break;
        case 'Inviato': stats.inviati++; break;
        case 'Errore': stats.errori++; break;
      }
      
      if (dataInvio && new Date(dataInvio).toDateString() === oggi) {
        stats.inviatOggi++;
      }
    }
  }
  
  // Mostra risultati
  console.log('');
  console.log('📊 DASHBOARD INVII WHATSAPP');
  console.log('================================');
  console.log(`📱 Totale contatti: ${stats.totale}`);
  console.log(`📤 Da inviare: ${stats.daInviare}`);
  console.log(`✅ Inviati totali: ${stats.inviati}`);
  console.log(`❌ Errori: ${stats.errori}`);
  console.log(`📅 Inviati oggi: ${stats.inviatOggi}`);
  console.log('================================');
  console.log('');
  
  return stats;
}

/**
 * Resetta errori per nuovo tentativo
 */
function riprovaErrori() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  const data = sheet.getDataRange().getValues();
  let resettati = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === 'Errore') { // Stato = Errore
      sheet.getRange(i + 1, 9).setValue('Da Inviare'); // Resetta a "Da Inviare"
      resettati++;
    }
  }
  
  console.log(`🔄 Resettati ${resettati} errori per nuovo tentativo`);
  logOperazione(`Resettati ${resettati} errori per nuovo tentativo`, 'INFO');
  
  return resettati;
}

// ====================================
// MENU E INTERFACCIA
// ====================================

/**
 * Crea menu personalizzato all'apertura
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Conta messaggi in coda
  const inCoda = contaMessaggiInCoda();
  const menuTitle = inCoda > 0 ? `📱 WhatsApp (${inCoda} in coda)` : '📱 WhatsApp Call Center';
  
  ui.createMenu(menuTitle)
    .addItem('📤 Elabora ora', 'processaClientiCompleto')
    .addItem('📊 Mostra dashboard', 'creaDashboard')
    .addItem('🔄 Riprova errori', 'riprovaErrori')
    .addSeparator()
    .addItem('❌ Mostra ultimi errori', 'mostraUltimiErrori')
    .addItem('🧹 Pulisci log vecchi', 'pulisciLogVecchi')
    .addItem('📈 Report giornaliero', 'reportGiornaliero')
    .addSeparator()
    .addItem('🧪 Test invio singolo', 'testInvioSingolo')
    .addItem('⚙️ Verifica configurazione', 'verificaConfigurazione')
    .addItem('🕐 Test orario', 'testOrario')
    .addItem('🔍 Debug dati clienti', 'debugDatiClienti')
    .addItem('📋 Verifica struttura fogli', 'verificaStruttura')
    .addSeparator()
    .addItem('🔄 Reset elaborazione bloccata', 'resetElaborazioneInCorso')
    .addToUi();
}

/**
 * Conta messaggi in coda
 */
function contaMessaggiInCoda() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Da Inviare') { // Colonna J - Stato WhatsApp
      count++;
    }
  }
  
  return count;
}

/**
 * Report giornaliero dettagliato
 */
function reportGiornaliero() {
  console.log('📊 REPORT GIORNALIERO');
  console.log('===================');
  console.log(`Data: ${Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy")}`);
  console.log(`Orario operativo: ${CONFIG.ORA_INIZIO}:00 - ${CONFIG.ORA_FINE}:00`);
  
  const stats = creaDashboard();
  
  console.log(`\n📈 STATISTICHE:`);
  console.log(`Messaggi inviati oggi: ${stats.inviatOggi}`);
  console.log(`In coda: ${stats.daInviare}`);
  console.log(`Errori: ${stats.errori}`);
  
  if (CONFIG.LIMITE_GIORNALIERO) {
    const percentuale = Math.round((stats.inviatOggi / CONFIG.LIMITE_GIORNALIERO) * 100);
    console.log(`\n📊 Limite giornaliero: ${stats.inviatOggi}/${CONFIG.LIMITE_GIORNALIERO} (${percentuale}%)`);
    
    if (stats.inviatOggi >= CONFIG.LIMITE_GIORNALIERO) {
      console.log('⚠️ LIMITE GIORNALIERO RAGGIUNTO!');
    }
  }
  
  // Verifica orario
  if (!isOrarioOperativo()) {
    console.log('\n⏰ FUORI ORARIO OPERATIVO');
    console.log('I messaggi in coda verranno elaborati domani');
  }
  
  // Statistiche per PDV
  console.log('\n📍 MESSAGGI PER PDV:');
  const perPdv = getStatistichePerPDV();
  Object.entries(perPdv).forEach(([pdv, count]) => {
    console.log(`${pdv}: ${count} messaggi`);
  });
}

/**
 * Ottieni statistiche per PDV
 */
function getStatistichePerPDV() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const oggi = Utilities.formatDate(new Date(), "Europe/Rome", "yyyy-MM-dd");
  const stats = {};
  
  for (let i = 1; i < data.length; i++) {
    const dataInvio = data[i][10];  // Colonna K - Data Invio
    const pdv = data[i][6] || 'Non specificato';  // Colonna G - PDV
    
    if (dataInvio) {
      const dataInvioStr = Utilities.formatDate(new Date(dataInvio), "Europe/Rome", "yyyy-MM-dd");
      if (dataInvioStr === oggi) {
        stats[pdv] = (stats[pdv] || 0) + 1;
      }
    }
  }
  
  return stats;
}

// ====================================
// FUNZIONI GESTIONE LOG ED ERRORI
// ====================================

/**
 * Mostra ultimi errori dal log
 */
function mostraUltimiErrori() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  if (!logSheet) {
    console.log('📋 Nessun log presente ancora');
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const errori = [];
  
  // Raccogli solo gli errori (salta header)
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === 'ERRORE') {
      errori.push({
        data: data[i][0],
        messaggio: data[i][2],
        telefono: data[i][3],
        dettaglio: data[i][4]
      });
    }
  }
  
  if (errori.length === 0) {
    console.log('✅ Nessun errore trovato nel log!');
    return;
  }
  
  // Mostra ultimi 10 errori
  console.log('❌ ULTIMI ERRORI WHATSAPP');
  console.log('=========================');
  
  const ultimi10 = errori.slice(-10).reverse();
  
  ultimi10.forEach((errore, index) => {
    console.log(`\n${index + 1}. ${new Date(errore.data).toLocaleString('it-IT')}`);
    console.log(`   📱 Telefono: ${errore.telefono}`);
    console.log(`   ❌ Errore: ${errore.dettaglio}`);
    console.log(`   👤 ${errore.messaggio}`);
  });
  
  console.log(`\n📊 Totale errori nel log: ${errori.length}`);
}

/**
 * Pulisce log più vecchi di 30 giorni
 */
function pulisciLogVecchi() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  if (!logSheet) {
    console.log('📋 Nessun log da pulire');
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const oggi = new Date();
  const trentaGiorniFa = new Date(oggi.getTime() - 30 * 24 * 60 * 60 * 1000);
  
  let righeEliminate = 0;
  
  // Parti dal fondo per non alterare gli indici mentre elimini
  for (let i = data.length - 1; i > 0; i--) {
    const dataRiga = new Date(data[i][0]);
    
    if (dataRiga < trentaGiorniFa) {
      logSheet.deleteRow(i + 1);
      righeEliminate++;
    }
  }
  
  console.log(`🧹 Eliminate ${righeEliminate} righe di log più vecchie di 30 giorni`);
  logOperazione(`Pulizia log: eliminate ${righeEliminate} righe vecchie`, 'INFO');
}

/**
 * Report errori per numero
 */
function reportErroriPerNumero() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  if (!logSheet) {
    console.log('📋 Nessun log presente');
    return;
  }
  
  const data = logSheet.getDataRange().getValues();
  const erroriPerNumero = {};
  
  // Conta errori per numero
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === 'ERRORE' && data[i][3]) {
      const numero = data[i][3];
      erroriPerNumero[numero] = (erroriPerNumero[numero] || 0) + 1;
    }
  }
  
  // Ordina per numero di errori
  const numeriOrdinati = Object.entries(erroriPerNumero)
    .sort((a, b) => b[1] - a[1]);
  
  console.log('📊 NUMERI CON PIÙ ERRORI');
  console.log('=======================');
  
  numeriOrdinati.slice(0, 10).forEach(([numero, count]) => {
    console.log(`${numero}: ${count} errori`);
  });
}

// ====================================
// FUNZIONI DI TEST E VERIFICA
// ====================================

/**
 * Test invio singolo
 */
function testInvioSingolo() {
  const NUMERO_TEST = '3401234567'; // ← INSERISCI IL TUO NUMERO
  
  console.log('=== TEST INVIO SINGOLO ===');
  
  const messaggioTest = 'Test sistema WhatsApp Call Center - Messaggio di prova! ✅';
  
  try {
    console.log('Numero test:', NUMERO_TEST);
    console.log('Numero formattato:', formattaNumeroItalia(NUMERO_TEST));
    
    const risultato = inviaWhatsApp(NUMERO_TEST, messaggioTest);
    console.log('✅ Invio riuscito!');
    console.log('Risposta:', risultato);
  } catch (error) {
    console.error('❌ Errore:', error.toString());
  }
}

/**
 * Verifica configurazione sistema
 */
function verificaConfigurazione() {
  console.log('');
  console.log('⚙️ VERIFICA CONFIGURAZIONE');
  console.log('==========================');
  console.log('✓ API Key configurata:', CONFIG.WASENDER_API_KEY !== 'INSERISCI-LA-TUA-API-KEY-QUI' ? 'SI' : '❌ NO - INSERIRE API KEY!');
  console.log('✓ URL API:', CONFIG.WASENDER_URL);
  console.log('✓ Delay tra invii:', CONFIG.DELAY_TRA_INVII/1000, 'secondi');
  console.log('✓ Max invii per sessione:', CONFIG.MAX_INVII_PER_SESSIONE);
  console.log('✓ Orario invii:', `${CONFIG.ORARIO_INIZIO}:00 - ${CONFIG.ORARIO_FINE}:00`);
  console.log('');
}

/**
 * Verifica struttura fogli
 */
function verificaStruttura() {
  console.log('');
  console.log('📋 VERIFICA STRUTTURA FOGLI');
  console.log('===========================');
  
  // Verifica foglio Clienti
  const clientiSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (clientiSheet) {
    console.log('✅ Foglio "Clienti" presente');
    const headers = clientiSheet.getRange(1, 1, 1, 12).getValues()[0];
    console.log('   Colonne:', headers.filter((h, i) => h || i === 3).map((h, i) => h || '(vuota)').join(', '));
    
    // Verifica headers critici
    if (headers[9] !== 'Stato WhatsApp') {
      console.log('⚠️ Colonna "Stato WhatsApp" non in posizione J (9)');
    }
    if (headers[11] !== 'Contatto Processato') {
      console.log('⚠️ Colonna "Contatto Processato" non in posizione L (11)');
    }
  } else {
    console.log('❌ Foglio "Clienti" MANCANTE!');
  }
  
  // Verifica foglio Template
  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Template');
  if (templateSheet) {
    console.log('✅ Foglio "Template" presente');
    const numTemplates = templateSheet.getLastRow() - 1;
    console.log(`   Template disponibili: ${numTemplates}`);
  } else {
    console.log('❌ Foglio "Template" MANCANTE!');
  }
  
  // Verifica foglio Log
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  if (logSheet) {
    console.log('✅ Foglio "Log" presente');
    const logData = logSheet.getDataRange().getValues();
    const errori = logData.filter(row => row[1] === 'ERRORE').length;
    console.log(`   Totale log: ${logData.length - 1}`);
    console.log(`   Errori registrati: ${errori}`);
  } else {
    console.log('⚠️ Foglio "Log" non presente (verrà creato automaticamente)');
  }
  
  console.log('');
}

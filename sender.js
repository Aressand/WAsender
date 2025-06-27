// ====================================
// WHATSAPP CALL CENTER - VERSIONE OTTIMIZZATA
// Sistema anti-duplicati integrato
// ====================================

// ====================================
// CONFIGURAZIONE SISTEMA
// ====================================
const CONFIG = {
  WASENDER_API_KEY: '',
  WASENDER_URL: 'https://wasenderapi.com/api/send-message',
  DELAY_TRA_INVII: 30000, // 30 secondi tra un invio e l'altro
  DELAY_DOPO_CONTATTO: 1000, // 1 secondo dopo creazione contatto
  MAX_INVII_PER_SESSIONE: 20, // Massimo messaggi per sessione
  LIMITE_GIORNALIERO: null, // null = nessun limite, oppure numero (es: 50)
  ORARIO_INIZIO: 9, // Ora inizio invii automatici
  ORARIO_FINE: 19, // Ora fine invii automatici
  MINUTI_SOGLIA_DUPLICATI: 30, // Considera duplicato se inviato negli ultimi 30 minuti
  AUTO_PULIZIA_STATI_BLOCCATI: true // Pulisci automaticamente stati "In Corso" vecchi
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
  
  if (numero.startsWith('+39')) {
    return numero;
  }
  
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
  
  for (let i = 1; i < templates.length; i++) {
    if (templates[i][0] == templateId) { 
      return {
        messaggio: templates[i][2] || '',
        immagineUrl: templates[i][3] || null
      };
    }
  }
  
  // Template di default
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
// SISTEMA ANTI-DUPLICATI
// ====================================

/**
 * Acquisisce lock per evitare esecuzioni multiple
 */
function acquisciLock() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Rilascia il lock
 */
function rilasciaLock() {
  try {
    const lock = LockService.getScriptLock();
    lock.releaseLock();
  } catch (e) {
    // Ignora errori di rilascio
  }
}

/**
 * Verifica se un messaggio è stato inviato di recente allo stesso numero
 */
function isMessaggioInviatoRecentemente(telefono) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return false;
  
  const numeroFormattato = formattaNumeroItalia(telefono);
  const now = new Date();
  const soglia = new Date(now.getTime() - CONFIG.MINUTI_SOGLIA_DUPLICATI * 60 * 1000);
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const telefonoRiga = formattaNumeroItalia(data[i][2]);
    const dataInvio = data[i][9];
    const stato = data[i][8];
    
    if (telefonoRiga === numeroFormattato && 
        dataInvio && 
        stato === 'Inviato' &&
        new Date(dataInvio) > soglia) {
      return true;
    }
  }
  
  return false;
}

/**
 * Pulizia stati "In Corso" rimasti bloccati
 */
function pulisciStatiInCorso() {
  if (!CONFIG.AUTO_PULIZIA_STATI_BLOCCATI) return 0;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let statiPuliti = 0;
  const now = new Date();
  const soglia = new Date(now.getTime() - 10 * 60 * 1000); // 10 minuti fa
  
  for (let i = 1; i < data.length; i++) {
    const stato = data[i][9];
    const dataInvio = data[i][10];
    
    if (stato === 'In Corso') {
      if (!dataInvio || new Date(dataInvio) < soglia) {
        sheet.getRange(i + 1, 9).setValue('Da Inviare');
        statiPuliti++;
      }
    }
  }
  
  if (statiPuliti > 0) {
    logOperazione(`Ripuliti ${statiPuliti} stati "In Corso" bloccati`, 'INFO');
  }
  
  return statiPuliti;
}

// ====================================
// GESTIONE CONTATTI GOOGLE
// ====================================

/**
 * Gestisce creazione/ricerca contatto Google
 */
function gestisciContattoGoogle(datiCliente) {
  try {
    const telefono = formattaNumeroItalia(datiCliente.telefono);
    
    const risultatiRicerca = People.People.searchContacts({
      query: telefono,
      readMask: 'names,phoneNumbers'
    });
    
    if (risultatiRicerca.results && risultatiRicerca.results.length > 0) {
      return 'ESISTENTE';
    }
    
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
    
    if (datiCliente.pdv) {
      nuovoContatto.organizations = [{
        name: datiCliente.pdv,
        title: 'Cliente'
      }];
    }
    
    People.People.createContact({
      resource: nuovoContatto
    });
    
    return 'CREATO';
    
  } catch (error) {
    return 'ERRORE';
  }
}

// ====================================
// FUNZIONE PRINCIPALE ANTI-DUPLICATI
// ====================================

/**
 * Processo principale con protezione anti-duplicati
 */
function processaClienti() {
  // Controllo lock per evitare esecuzioni multiple
  if (!acquisciLock()) {
    return;
  }

  try {
    // Pulizia automatica stati bloccati
    pulisciStatiInCorso();
    
    // Controllo orario operativo
    if (!isOrarioOperativo()) {
      return;
    }
    
    // Controllo limite giornaliero
    const messaggiOggi = contaMessaggiInviatiOggi();
    if (CONFIG.LIMITE_GIORNALIERO && messaggiOggi >= CONFIG.LIMITE_GIORNALIERO) {
      logOperazione(`Raggiunto limite giornaliero: ${messaggiOggi}/${CONFIG.LIMITE_GIORNALIERO}`, 'INFO');
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
    if (!sheet) {
      throw new Error('Foglio "Clienti" non trovato!');
    }
    
    // Sincronizza dati
    SpreadsheetApp.flush();
    
    const data = sheet.getDataRange().getValues();
    let contattiProcessati = 0;
    let contattiCreati = 0;
    let contattiEsistenti = 0;
    let errori = 0;
    let duplicatiEvitati = 0;
    
    // Calcola limite effettivo per questa sessione
    let limiteSessione = CONFIG.MAX_INVII_PER_SESSIONE;
    if (CONFIG.LIMITE_GIORNALIERO) {
      const rimanenti = CONFIG.LIMITE_GIORNALIERO - messaggiOggi;
      limiteSessione = Math.min(limiteSessione, rimanenti);
    }
    
    // Processa righe
    for (let i = 1; i < data.length; i++) {
      if (contattiProcessati >= limiteSessione) {
        break;
      }
      
      const row = data[i];
      const datiCliente = {
        nome: row[0],
        cognome: row[1],
        telefono: row[2],
        dataChiamata: row[3],
        esito: row[4],
        pdv: row[5],
        operatore: row[6],
        templateId: row[7],
        statoWhatsApp: row[8],
        dataInvio: row[9],
        contattoProcessato: row[10]
      };
      
      // Processa solo se "Da Inviare" e con telefono
      if (datiCliente.statoWhatsApp === 'Da Inviare' && datiCliente.telefono) {
        
        // Controllo anti-duplicati
        if (isMessaggioInviatoRecentemente(datiCliente.telefono)) {
          duplicatiEvitati++;
          sheet.getRange(i + 1, 9).setValue('Duplicato Evitato');
          SpreadsheetApp.flush();
          continue;
        }
        
        // Doppio controllo dello stato
        const statoCorrente = sheet.getRange(i + 1, 9).getValue();
        if (statoCorrente !== 'Da Inviare') {
          continue;
        }
        
        try {
          // Cambia stato PRIMA dell'invio
          sheet.getRange(i + 1, 9).setValue('In Corso');
          sheet.getRange(i + 1, 10).setValue(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss");
          SpreadsheetApp.flush();
          
          // Gestione contatto Google (se disponibile)
          if (typeof People !== 'undefined') {
            try {
              const risultatoContatto = gestisciContattoGoogle(datiCliente);
              if (risultatoContatto === 'CREATO') {
                contattiCreati++;
                Utilities.sleep(CONFIG.DELAY_DOPO_CONTATTO);
              } else {
                contattiEsistenti++;
              }
              
              sheet.getRange(i + 1, 11).setValue('Si');
              
            } catch (e) {
              sheet.getRange(i + 1, 11).setValue('Errore');
            }
          }
          
          // Invio WhatsApp
          const template = getTemplate(datiCliente.templateId || '1');
          const messaggioPersonalizzato = sostituisciShortcode(template.messaggio, datiCliente);
          const immagineFinale = template.immagineUrl;
          
          // Controllo finale prima dell'invio
          const ultimoControlloStato = sheet.getRange(i + 1, 9).getValue();
          if (ultimoControlloStato !== 'In Corso') {
            continue;
          }
          
          inviaWhatsApp(datiCliente.telefono, messaggioPersonalizzato, immagineFinale);
          
          // Aggiorna stato dopo invio riuscito
          sheet.getRange(i + 1, 9).setValue('Inviato');
          sheet.getRange(i + 1, 10).setValue(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss");
          SpreadsheetApp.flush();
          
          logInvioRiuscito(datiCliente.telefono, `${datiCliente.nome} ${datiCliente.cognome}`);
          
          contattiProcessati++;
          
          // Pausa tra invii
          if (contattiProcessati < limiteSessione && i < data.length - 1) {
            Utilities.sleep(CONFIG.DELAY_TRA_INVII);
          }
          
        } catch (error) {
          sheet.getRange(i + 1, 9).setValue('Errore');
          sheet.getRange(i + 1, 10).setValue(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss");
          SpreadsheetApp.flush();
          
          logErrore(
            formattaNumeroItalia(datiCliente.telefono),
            error.toString(),
            `${datiCliente.nome} ${datiCliente.cognome}`
          );
          
          errori++;
        }
      }
    }
    
    // Log riepilogo se ci sono stati invii
    if (contattiProcessati > 0 || duplicatiEvitati > 0 || errori > 0) {
      logOperazione(
        `Elaborazione: ${contattiProcessati} inviati, ${duplicatiEvitati} duplicati evitati, ${errori} errori`,
        'INFO'
      );
    }
    
  } finally {
    rilasciaLock();
  }
}

// ====================================
// FUNZIONI UTILITA'
// ====================================

/**
 * Controlla se siamo in orario operativo
 */
function isOrarioOperativo() {
  const now = new Date();
  const oraItalia = Utilities.formatDate(now, "Europe/Rome", "HH");
  const oraCorrente = parseInt(oraItalia);
  
  return oraCorrente >= CONFIG.ORARIO_INIZIO && oraCorrente < CONFIG.ORARIO_FINE;
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
    const dataInvio = data[i][9];
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
 * Conta messaggi in coda
 */
function contaMessaggiInCoda() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Da Inviare') {
      count++;
    }
  }
  
  return count;
}

// ====================================
// LOGGING
// ====================================

/**
 * Crea/aggiorna foglio di log
 */
function logOperazione(messaggio, tipo = 'INFO', telefono = '', dettagliErrore = '') {
  let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  if (!logSheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = ss.insertSheet('Log');
    logSheet.getRange(1, 1, 1, 5)
      .setValues([['Data/Ora', 'Tipo', 'Messaggio', 'Telefono', 'Dettagli']])
      .setFontWeight('bold')
      .setBackground('#f0f0f0');
    
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(2, 80);
    logSheet.setColumnWidth(3, 250);
    logSheet.setColumnWidth(4, 120);
    logSheet.setColumnWidth(5, 300);
  }
  
  logSheet.appendRow([new Date(), tipo, messaggio, telefono, dettagliErrore]);
  
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
 * Log specifico per invii riusciti
 */
function logInvioRiuscito(telefono, nomeCliente) {
  logOperazione(`Messaggio inviato a ${nomeCliente}`, 'INVIO', telefono);
}

// ====================================
// DASHBOARD E REPORT
// ====================================

/**
 * Dashboard con statistiche
 */
function mostraDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  const data = sheet.getDataRange().getValues();
  
  let stats = {
    totale: 0,
    daInviare: 0,
    inviati: 0,
    errori: 0,
    inCorso: 0,
    inviatOggi: 0
  };
  
  const oggi = new Date().toDateString();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2]) {
      stats.totale++;
      
      const stato = data[i][9];
      const dataInvio = data[i][10];
      
      switch(stato) {
        case 'Da Inviare': stats.daInviare++; break;
        case 'Inviato': stats.inviati++; break;
        case 'Errore': stats.errori++; break;
        case 'In Corso': stats.inCorso++; break;
      }
      
      if (dataInvio && new Date(dataInvio).toDateString() === oggi) {
        stats.inviatOggi++;
      }
    }
  }
  
  console.log('📊 DASHBOARD WHATSAPP CALL CENTER');
  console.log('=================================');
  console.log(`📱 Totale contatti: ${stats.totale}`);
  console.log(`📤 Da inviare: ${stats.daInviare}`);
  console.log(`✅ Inviati totali: ${stats.inviati}`);
  console.log(`⏳ In corso: ${stats.inCorso}`);
  console.log(`❌ Errori: ${stats.errori}`);
  console.log(`📅 Inviati oggi: ${stats.inviatOggi}`);
  
  if (CONFIG.LIMITE_GIORNALIERO) {
    const percentuale = Math.round((stats.inviatOggi / CONFIG.LIMITE_GIORNALIERO) * 100);
    console.log(`📈 Limite giornaliero: ${stats.inviatOggi}/${CONFIG.LIMITE_GIORNALIERO} (${percentuale}%)`);
  }
  
  console.log('=================================');
  
  return stats;
}

/**
 * Report giornaliero
 */
function reportGiornaliero() {
  console.log('📊 REPORT GIORNALIERO');
  console.log('===================');
  console.log(`Data: ${Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy")}`);
  
  const stats = mostraDashboard();
  
  if (!isOrarioOperativo()) {
    console.log('\n⏰ FUORI ORARIO OPERATIVO');
  }
  
  return stats;
}

// ====================================
// FUNZIONI DI GESTIONE
// ====================================

/**
 * Resetta errori per nuovo tentativo
 */
function riprovaErrori() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  const data = sheet.getDataRange().getValues();
  let resettati = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'Errore') {
      sheet.getRange(i + 1, 9).setValue('Da Inviare');
      resettati++;
    }
  }
  
  console.log(`🔄 Resettati ${resettati} errori per nuovo tentativo`);
  logOperazione(`Resettati ${resettati} errori per nuovo tentativo`, 'INFO');
  
  return resettati;
}

/**
 * Reset manuale stati "In Corso"
 */
function resetStatiInCorso() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  const data = sheet.getDataRange().getValues();
  let resetCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === 'In Corso') {
      sheet.getRange(i + 1, 9).setValue('Da Inviare');
      resetCount++;
    }
  }
  
  console.log(`🔄 Reset ${resetCount} stati "In Corso"`);
  logOperazione(`Reset manuale: ${resetCount} stati resettati da "In Corso"`, 'INFO');
  
  return resetCount;
}

/**
 * Setup iniziale
 */
function setupIniziale() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    console.log('❌ Foglio "Clienti" non trovato!');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (headers[11] === 'Contatto Processato') {
    console.log('✅ Colonna "Contatto Processato" già presente');
    return;
  }
  
  sheet.getRange(1, 11).setValue('Contatto Processato');
  sheet.getRange(1, 11).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.setColumnWidth(12, 130);
  
  console.log('✅ Setup completato!');
  console.log('📝 Configura CONFIG.WASENDER_API_KEY e imposta trigger automatico');
}

/**
 * Test invio singolo
 */
function testInvioSingolo() {
  const NUMERO_TEST = '3408076933'; // ← INSERISCI IL TUO NUMERO
  
  console.log('=== TEST INVIO SINGOLO ===');
  
  const messaggioTest = 'Test sistema WhatsApp Call Center - Anti-duplicati attivo! ✅';
  
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

// ====================================
// MENU PRINCIPALE
// ====================================

/**
 * Crea menu personalizzato
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  const inCoda = contaMessaggiInCoda();
  const menuTitle = inCoda > 0 ? `📱 WhatsApp (${inCoda} in coda)` : '📱 WhatsApp Call Center';
  
  ui.createMenu(menuTitle)
    .addItem('📤 Elabora ora', 'processaClienti')
    .addItem('📊 Dashboard', 'mostraDashboard')
    .addItem('📈 Report giornaliero', 'reportGiornaliero')
    .addSeparator()
    .addItem('🔄 Riprova errori', 'riprovaErrori')
    .addItem('🔄 Reset stati "In Corso"', 'resetStatiInCorso')
    .addSeparator()
    .addItem('🧪 Test invio singolo', 'testInvioSingolo')
    .addItem('⚙️ Setup iniziale', 'setupIniziale')
    .addToUi();
}

// ====================================
// ISTRUZIONI FINALI
// ====================================

/*
ISTRUZIONI PER LA CONFIGURAZIONE:

STRUTTURA FOGLIO "Clienti":
A: Nome | B: Cognome | C: Telefono | D: Data Chiamata
E: Esito | F: PDV | G: Operatore | H: Template | I: Stato WhatsApp
J: Data Invio | K: Contatto Processato

PASSI PER L'INSTALLAZIONE:
1. Esegui "setupIniziale()" 
2. Configura CONFIG.WASENDER_API_KEY
3. (Opzionale) Abilita Google People API per contatti automatici
4. Imposta trigger automatico ogni 10 minuti per "processaClienti"
5. Configura CONFIG.LIMITE_GIORNALIERO se necessario

CARATTERISTICHE ANTI-DUPLICATI:
- Lock per evitare esecuzioni multiple
- Controllo messaggi inviati negli ultimi 30 minuti
- Stato "In Corso" per prevenire doppi invii
- Pulizia automatica stati bloccati
- Sincronizzazione forzata dei dati

USO QUOTIDIANO:
- Operatori inseriscono dati con stato "Da Inviare"
- Sistema elabora automaticamente ogni 10 minuti
- Menu → "Elabora ora" per invio immediato
- Menu → "Dashboard" per statistiche
*/

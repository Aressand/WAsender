/**
 * Sistema WhatsApp Call Center con Google Sheets
 * Versione: 2.0
 * 
 * Funzionalità:
 * - Invio automatico messaggi WhatsApp tramite WASender
 * - Gestione contatti Google (opzionale)
 * - Sistema anti-duplicati
 * - Elaborazione con limiti orari e giornalieri
 * - Report e statistiche
 */

// ====================================
// CONFIGURAZIONE
// ====================================
const CONFIG = {
  // WASender API
  WASENDER_API_KEY: '',
  WASENDER_DEVICE_ID: '39',
  
  // Limiti invio
  LIMITE_GIORNALIERO: 100,
  MAX_INVII_PER_SESSIONE: 20,
  DELAY_TRA_INVII: 30000, // 30 secondi
  
  // Orario operativo
  ORA_INIZIO: 9,
  ORA_FINE: 19,
  
  // Google Contacts
  USA_GOOGLE_CONTACTS: false,
  GRUPPO_CONTATTI: 'Call Center WhatsApp'
};

// ====================================
// MENU PERSONALIZZATO
// ====================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 WhatsApp Call Center')
    .addItem('📤 Elabora ora', 'processaClientiCompleto')
    .addItem('📊 Report giornaliero', 'mostraReportGiornaliero')
    .addItem('📈 Statistiche complete', 'mostraStatistiche')
    .addSeparator()
    .addItem('🔧 Setup iniziale', 'setupIniziale')
    .addItem('🧪 Test invio singolo', 'testInvioSingolo')
    .addItem('⚙️ Verifica configurazione', 'verificaConfigurazione')
    .addSeparator()
    .addItem('🔄 Reset elaborazione bloccata', 'resetElaborazioneInCorso')
    .addToUi();
}

// ====================================
// FUNZIONI ANTI-DUPLICATI
// ====================================

/**
 * Controlla se ci sono altre esecuzioni in corso
 */
function isElaborazioneInCorso() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastRun = scriptProperties.getProperty('ELABORAZIONE_IN_CORSO');
  
  if (lastRun) {
    const lastRunTime = new Date(lastRun);
    const now = new Date();
    const diffMinutes = (now - lastRunTime) / (1000 * 60);
    
    // Se l'ultima esecuzione è stata meno di 5 minuti fa, blocca
    if (diffMinutes < 5) {
      console.log(`⚠️ Elaborazione già in corso (iniziata ${Math.floor(diffMinutes)} minuti fa)`);
      return true;
    }
  }
  
  return false;
}

/**
 * Imposta flag elaborazione in corso
 */
function setElaborazioneInCorso(inCorso = true) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (inCorso) {
    scriptProperties.setProperty('ELABORAZIONE_IN_CORSO', new Date().toISOString());
  } else {
    scriptProperties.deleteProperty('ELABORAZIONE_IN_CORSO');
  }
}

/**
 * Resetta il flag di elaborazione in corso
 */
function resetElaborazioneInCorso() {
  setElaborazioneInCorso(false);
  console.log('✅ Flag elaborazione resettato');
  SpreadsheetApp.getUi().alert('Flag elaborazione resettato. Puoi ora eseguire una nuova elaborazione.');
}

// ====================================
// FUNZIONI PRINCIPALI
// ====================================

/**
 * Processa clienti con stato "Da Inviare" con protezione anti-duplicati
 */
function processaClientiCompleto() {
  // Controlla se c'è già un'elaborazione in corso
  if (isElaborazioneInCorso()) {
    console.log('❌ Elaborazione già in corso, annullo questa esecuzione');
    logOperazione('WARNING', 'Tentativo di avviare elaborazione mentre una è già in corso');
    return;
  }
  
  // Imposta flag elaborazione in corso
  setElaborazioneInCorso(true);
  
  try {
    // Controlla orario operativo
    const controlloOrario = controllaOrarioOperativo();
    console.log(`Controllo orario: ${controlloOrario.oraRoma} (Roma)`);
    console.log(`Deve essere tra ${CONFIG.ORA_INIZIO}:00 e ${CONFIG.ORA_FINE}:00`);
    console.log(`Risultato: ${controlloOrario.operativo ? 'OPERATIVO' : 'FUORI ORARIO'}`);
    
    if (!controlloOrario.operativo) {
      const messaggio = `Elaborazione non eseguita: fuori orario operativo (${controlloOrario.oraRoma})`;
      console.log(messaggio);
      logOperazione('INFO', messaggio);
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
    
    if (!sheet) {
      throw new Error('Foglio "Clienti" non trovato!');
    }
    
    const data = sheet.getDataRange().getValues();
    
    console.log('🚀 Inizio elaborazione completa...');
    console.log(`📅 ${Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy HH:mm:ss")}`);
    console.log(`📊 Analizzando ${data.length - 1} righe di dati...`);
    
    // Calcola limite effettivo per questa sessione
    const inviiOggi = contaInviiOggi();
    const limiteRimanente = CONFIG.LIMITE_GIORNALIERO - inviiOggi;
    const limiteSessione = Math.min(CONFIG.MAX_INVII_PER_SESSIONE, limiteRimanente);
    
    console.log(`📊 Invii oggi: ${inviiOggi}/${CONFIG.LIMITE_GIORNALIERO}`);
    console.log(`📊 Limite per questa sessione: ${limiteSessione}`);
    
    if (limiteSessione <= 0) {
      const messaggio = 'Limite giornaliero raggiunto!';
      console.log(`⚠️ ${messaggio}`);
      logOperazione('INFO', messaggio);
      return;
    }
    
    // Contatori
    let contattiProcessati = 0;
    let messaggiInviati = 0;
    let contattiCreati = 0;
    let contattiEsistenti = 0;
    let errori = 0;
    
    // Set per tracciare numeri già processati in questa sessione
    const numeriProcessati = new Set();
    
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
        const numeroFormattato = formattaNumeroItalia(datiCliente.telefono);
        
        // Controlla se già processato in questa sessione
        if (numeriProcessati.has(numeroFormattato)) {
          console.log(`⚠️ Numero ${numeroFormattato} già processato in questa sessione, salto`);
          continue;
        }
        
        console.log(`\n📱 Processando ${datiCliente.nome} ${datiCliente.cognome} - ${numeroFormattato}`);
        
        try {
          // 1. Crea/Aggiorna contatto se abilitato
          if (CONFIG.USA_GOOGLE_CONTACTS && !datiCliente.contattoProcessato) {
            try {
              const risultatoContatto = creaOAggiornaContatto(datiCliente);
              if (risultatoContatto.creato) {
                contattiCreati++;
              } else {
                contattiEsistenti++;
              }
              
              // Marca come processato
              sheet.getRange(i + 1, 12).setValue('SI');
              
              console.log(`   ✅ Contatto ${risultatoContatto.creato ? 'creato' : 'già esistente'}`);
            } catch (error) {
              console.log(`   ⚠️ Errore contatto: ${error.toString()}`);
            }
          }
          
          // 2. Prepara e invia messaggio
          const template = getTemplate(datiCliente.templateId);
          const messaggio = personalizzaMessaggio(template.messaggio, datiCliente);
          
          // Usa immagine dal template (non c'è colonna immagineUrl)
          const immagineFinale = template.immagineUrl;
          
          console.log(`   📝 Template: ${template.nome}`);
          console.log(`   💬 Messaggio: ${messaggio.substring(0, 50)}...`);
          
          const risultatoInvio = inviaMessaggioWhatsApp(
            numeroFormattato,
            messaggio,
            immagineFinale
          );
          
          if (risultatoInvio.success) {
            console.log(`   ✅ Messaggio inviato con successo`);
            
            // Aggiorna stato a "Inviato"
            sheet.getRange(i + 1, 10).setValue('Inviato');  // Colonna J
            sheet.getRange(i + 1, 11).setValue(new Date()); // Colonna K
            
            // Forza il salvataggio immediato
            SpreadsheetApp.flush();
            
            // Aggiungi a numeri processati
            numeriProcessati.add(numeroFormattato);
            
            contattiProcessati++;
            messaggiInviati++;
            
            // Delay tra invii
            if (contattiProcessati < limiteSessione) {
              console.log(`   ⏳ Attendo ${CONFIG.DELAY_TRA_INVII/1000} secondi...`);
              Utilities.sleep(CONFIG.DELAY_TRA_INVII);
            }
          } else {
            throw new Error(risultatoInvio.error || 'Errore sconosciuto');
          }
          
        } catch (error) {
          errori++;
          
          // Se lo stato è già "Errore", mantienilo ma logga il nuovo tentativo
          const statoAttuale = sheet.getRange(i + 1, 10).getValue();
          if (statoAttuale !== 'Errore') {
            sheet.getRange(i + 1, 10).setValue('Errore');  // Colonna J
          }
          
          // Mantieni stato "Errore" ma logga nuovo tentativo
          console.error(`❌ Ancora errore per ${datiCliente.nome}:`, error.toString());
        }
      }
    }
    
    // Risultati
    const risultato = {
      inviati: messaggiInviati,
      contattiCreati: contattiCreati,
      contattiEsistenti: contattiEsistenti,
      errori: errori
    };
    
    // Log riepilogo
    const messaggio = `Elaborazione: ${risultato.inviati} inviati, ${risultato.contattiCreati} contatti creati, ${risultato.errori} errori.`;
    logOperazione('INFO', messaggio);
    
    console.log('\n📊 RIEPILOGO ELABORAZIONE');
    console.log('========================');
    console.log(`✅ Messaggi inviati: ${risultato.inviati}`);
    console.log(`📇 Contatti creati: ${risultato.contattiCreati}`);
    console.log(`ℹ️ Contatti esistenti: ${risultato.contattiEsistenti}`);
    console.log(`❌ Errori: ${risultato.errori}`);
    console.log(`📅 Completato: ${Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM/yyyy, HH:mm:ss")}`);
    
  } catch (error) {
    console.error('❌ Errore durante elaborazione:', error);
    logOperazione('ERROR', `Errore elaborazione: ${error.toString()}`);
  } finally {
    // Rimuovi sempre il flag elaborazione in corso
    setElaborazioneInCorso(false);
  }
}

// ====================================
// FUNZIONI WASENDER
// ====================================

/**
 * Invia messaggio WhatsApp tramite WASender
 */
function inviaMessaggioWhatsApp(numero, messaggio, immagineUrl = null) {
  try {
    const payload = {
      api_key: CONFIG.WASENDER_API_KEY,
      device: CONFIG.WASENDER_DEVICE_ID,
      number: numero,
      message: messaggio
    };
    
    // Aggiungi immagine se presente
    if (immagineUrl) {
      payload.image_url = immagineUrl;
    }
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.wasender.net/send', options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status === true || result.status === 'success') {
      return { success: true };
    } else {
      return { 
        success: false, 
        error: result.message || 'Errore sconosciuto' 
      };
    }
    
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

// ====================================
// FUNZIONI GOOGLE CONTACTS
// ====================================

/**
 * Crea o aggiorna contatto in Google Contacts
 */
function creaOAggiornaContatto(datiCliente) {
  if (!CONFIG.USA_GOOGLE_CONTACTS) {
    return { creato: false, esistente: false };
  }
  
  try {
    const numeroFormattato = formattaNumeroItalia(datiCliente.telefono);
    const nomeCompleto = `${datiCliente.nome} ${datiCliente.cognome}`.trim();
    
    // Cerca contatto esistente per numero
    const contattiEsistenti = People.People.searchContacts({
      query: numeroFormattato,
      readMask: 'names,phoneNumbers'
    });
    
    if (contattiEsistenti.results && contattiEsistenti.results.length > 0) {
      // Contatto già esistente
      return { creato: false, esistente: true };
    }
    
    // Crea nuovo contatto
    const nuovoContatto = {
      names: [{
        givenName: datiCliente.nome,
        familyName: datiCliente.cognome
      }],
      phoneNumbers: [{
        value: numeroFormattato,
        type: 'mobile'
      }],
      memberships: [{
        contactGroupMembership: {
          contactGroupResourceName: getOrCreateContactGroup()
        }
      }]
    };
    
    // Aggiungi note se presenti dati extra
    const note = [];
    if (datiCliente.pdv) note.push(`PDV: ${datiCliente.pdv}`);
    if (datiCliente.operatore) note.push(`Operatore: ${datiCliente.operatore}`);
    if (datiCliente.dataChiamata) note.push(`Data chiamata: ${datiCliente.dataChiamata}`);
    if (datiCliente.esito) note.push(`Esito: ${datiCliente.esito}`);
    
    if (note.length > 0) {
      nuovoContatto.biographies = [{
        value: note.join('\n'),
        contentType: 'TEXT_PLAIN'
      }];
    }
    
    People.People.createContact(nuovoContatto);
    
    return { creato: true, esistente: false };
    
  } catch (error) {
    console.error('Errore gestione contatto:', error);
    throw error;
  }
}

/**
 * Ottieni o crea gruppo contatti
 */
function getOrCreateContactGroup() {
  try {
    // Cerca gruppo esistente
    const gruppi = People.ContactGroups.list({
      pageSize: 100
    });
    
    if (gruppi.contactGroups) {
      const gruppoEsistente = gruppi.contactGroups.find(
        g => g.name === CONFIG.GRUPPO_CONTATTI
      );
      
      if (gruppoEsistente) {
        return gruppoEsistente.resourceName;
      }
    }
    
    // Crea nuovo gruppo
    const nuovoGruppo = People.ContactGroups.create({
      contactGroup: {
        name: CONFIG.GRUPPO_CONTATTI
      }
    });
    
    return nuovoGruppo.resourceName;
    
  } catch (error) {
    console.error('Errore gestione gruppo contatti:', error);
    throw error;
  }
}

// ====================================
// FUNZIONI TEMPLATE
// ====================================

/**
 * Ottieni template per ID
 */
function getTemplate(templateId = '1') {
  // Converti in stringa se è un numero
  const idTemplate = String(templateId || '1');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Template');
  
  if (!sheet) {
    throw new Error('Foglio Template non trovato!');
  }
  
  const data = sheet.getDataRange().getValues();
  
  // Cerca template per ID (colonna A)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === idTemplate) {
      return {
        id: data[i][0],
        nome: data[i][1],
        messaggio: data[i][2],
        immagineUrl: data[i][3] || ''
      };
    }
  }
  
  // Se non trovato, usa il primo template
  console.log(`⚠️ Template ${idTemplate} non trovato, uso il template 1`);
  return {
    id: '1',
    nome: 'Default',
    messaggio: data[1] && data[1][2] ? data[1][2] : 'Ciao {nome}, grazie per averci contattato!',
    immagineUrl: data[1] && data[1][3] ? data[1][3] : ''
  };
}

/**
 * Personalizza messaggio con dati cliente
 */
function personalizzaMessaggio(messaggio, datiCliente) {
  return messaggio
    .replace(/{nome}/g, datiCliente.nome || '')
    .replace(/{cognome}/g, datiCliente.cognome || '')
    .replace(/{nomeCompleto}/g, `${datiCliente.nome} ${datiCliente.cognome}`.trim())
    .replace(/{pdv}/g, datiCliente.pdv || '')
    .replace(/{operatore}/g, datiCliente.operatore || '')
    .replace(/{dataChiamata}/g, datiCliente.dataChiamata || '');
}

// ====================================
// FUNZIONI UTILITÀ
// ====================================

/**
 * Formatta numero per l'Italia
 */
function formattaNumeroItalia(numero) {
  // Rimuovi spazi e caratteri non numerici
  let pulito = String(numero).replace(/\D/g, '');
  
  // Se inizia con 39, è già formato internazionale
  if (pulito.startsWith('39')) {
    return pulito;
  }
  
  // Se inizia con +39, rimuovi il +
  if (pulito.startsWith('+39')) {
    return pulito.substring(1);
  }
  
  // Altrimenti aggiungi prefisso Italia
  return '39' + pulito;
}

/**
 * Controlla orario operativo
 */
function controllaOrarioOperativo() {
  const now = new Date();
  const romaTime = new Date(now.toLocaleString("en-US", {timeZone: "Europe/Rome"}));
  const ora = romaTime.getHours();
  const minuti = romaTime.getMinutes();
  const oraFormattata = `${ora}:${minuti.toString().padStart(2, '0')}`;
  
  const operativo = ora >= CONFIG.ORA_INIZIO && ora < CONFIG.ORA_FINE;
  
  return {
    operativo: operativo,
    ora: ora,
    oraRoma: oraFormattata
  };
}

/**
 * Conta invii di oggi
 */
function contaInviiOggi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  if (!sheet) return 0;
  
  const data = sheet.getDataRange().getValues();
  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);
  
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const dataInvio = data[i][10]; // Colonna K - Data Invio
    if (dataInvio) {
      const dataInvioObj = new Date(dataInvio);
      dataInvioObj.setHours(0, 0, 0, 0);
      
      if (dataInvioObj.getTime() === oggi.getTime()) {
        count++;
      }
    }
  }
  
  return count;
}

/**
 * Conta messaggi da inviare
 */
function contaDaInviare() {
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

// ====================================
// FUNZIONI LOG
// ====================================

/**
 * Log operazione
 */
function logOperazione(tipo, messaggio) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  
  if (!sheet) {
    console.log('Foglio Log non trovato, lo creo...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.insertSheet('Log');
    logSheet.getRange(1, 1, 1, 3).setValues([['Data/Ora', 'Tipo', 'Messaggio']]);
    logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  logSheet.appendRow([
    new Date(),
    tipo,
    messaggio
  ]);
}

// ====================================
// FUNZIONI REPORT
// ====================================

/**
 * Mostra report giornaliero
 */
function mostraReportGiornaliero() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Foglio "Clienti" non trovato!');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);
  
  let inviatiOggi = 0;
  let perPDV = {};
  
  for (let i = 1; i < data.length; i++) {
    const dataInvio = data[i][10];  // Colonna K - Data Invio
    const pdv = data[i][6] || 'Non specificato';  // Colonna G - PDV
    
    if (dataInvio) {
      const dataInvioObj = new Date(dataInvio);
      dataInvioObj.setHours(0, 0, 0, 0);
      
      if (dataInvioObj.getTime() === oggi.getTime()) {
        inviatiOggi++;
        perPDV[pdv] = (perPDV[pdv] || 0) + 1;
      }
    }
  }
  
  let report = `📊 REPORT GIORNALIERO - ${Utilities.formatDate(oggi, "Europe/Rome", "dd/MM/yyyy")}\n\n`;
  report += `📨 Messaggi inviati oggi: ${inviatiOggi}\n`;
  report += `📍 Messaggi da inviare: ${contaDaInviare()}\n\n`;
  
  if (Object.keys(perPDV).length > 0) {
    report += `📍 INVII PER PDV:\n`;
    for (let pdv in perPDV) {
      report += `   • ${pdv}: ${perPDV[pdv]}\n`;
    }
  }
  
  report += `\n⏰ Orario operativo: ${CONFIG.ORA_INIZIO}:00 - ${CONFIG.ORA_FINE}:00\n`;
  report += `📊 Limite giornaliero: ${CONFIG.LIMITE_GIORNALIERO}`;
  
  SpreadsheetApp.getUi().alert('Report Giornaliero', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Mostra statistiche complete
 */
function mostraStatistiche() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Foglio "Clienti" non trovato!');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  // Inizializza contatori
  const stati = {};
  let totaleInviati = 0;
  let contattiProcessati = 0;
  
  // Calcola statistiche (salta header)
  for (let i = 1; i < data.length; i++) {
    const statoWhatsApp = data[i][9];   // Colonna J
    const dataInvio = data[i][10];      // Colonna K
    const contattoProcessato = data[i][11]; // Colonna L
    
    stati[statoWhatsApp] = (stati[statoWhatsApp] || 0) + 1;
    
    if (dataInvio) {
      totaleInviati++;
    }
    
    if (contattoProcessato === 'SI') {
      contattiProcessati++;
    }
  }
  
  // Messaggi per operatore
  const perOperatore = {};
  for (let i = 1; i < data.length; i++) {
    const operatore = data[i][7];  // Colonna H - Operatore
    const stato = data[i][9];      // Colonna J - Stato WhatsApp
    if (operatore && stato === 'Inviato') {
      perOperatore[operatore] = (perOperatore[operatore] || 0) + 1;
    }
  }
  
  let stats = `📊 STATISTICHE COMPLETE\n\n`;
  stats += `📋 TOTALE RECORD: ${data.length - 1}\n\n`;
  
  stats += `📨 STATI WHATSAPP:\n`;
  for (let stato in stati) {
    if (stato) {
      stats += `   • ${stato}: ${stati[stato]}\n`;
    }
  }
  
  stats += `\n📤 Totale messaggi inviati: ${totaleInviati}\n`;
  stats += `📇 Contatti processati: ${contattiProcessati}\n`;
  
  if (Object.keys(perOperatore).length > 0) {
    stats += `\n👥 INVII PER OPERATORE:\n`;
    for (let op in perOperatore) {
      stats += `   • ${op}: ${perOperatore[op]}\n`;
    }
  }
  
  stats += `\n⚙️ CONFIGURAZIONE:\n`;
  stats += `   • Limite giornaliero: ${CONFIG.LIMITE_GIORNALIERO}\n`;
  stats += `   • Invii per sessione: ${CONFIG.MAX_INVII_PER_SESSIONE}\n`;
  stats += `   • Delay tra invii: ${CONFIG.DELAY_TRA_INVII/1000} secondi\n`;
  stats += `   • Google Contacts: ${CONFIG.USA_GOOGLE_CONTACTS ? 'Attivo' : 'Disattivo'}`;
  
  SpreadsheetApp.getUi().alert('Statistiche Complete', stats, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ====================================
// FUNZIONI TEST E SETUP
// ====================================

/**
 * Test invio singolo messaggio
 */
function testInvioSingolo() {
  const ui = SpreadsheetApp.getUi();
  
  const numeroResponse = ui.prompt(
    'Test Invio', 
    'Inserisci numero di telefono (es: 3401234567):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (numeroResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const numero = numeroResponse.getResponseText();
  
  if (!numero) {
    ui.alert('Numero non valido!');
    return;
  }
  
  const nomeResponse = ui.prompt(
    'Test Invio', 
    'Inserisci nome destinatario:', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (nomeResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const nome = nomeResponse.getResponseText() || 'Test';
  
  try {
    ui.alert('Invio in corso...');
    
    const template = getTemplate('1');
    const messaggio = personalizzaMessaggio(template.messaggio, {
      nome: nome,
      cognome: '',
      pdv: 'Test PDV',
      operatore: 'Test'
    });
    
    const numeroFormattato = formattaNumeroItalia(numero);
    const risultato = inviaMessaggioWhatsApp(numeroFormattato, messaggio);
    
    if (risultato.success) {
      ui.alert('✅ Successo!', `Messaggio inviato a ${numeroFormattato}`, ui.ButtonSet.OK);
      logOperazione('TEST', `Test invio riuscito a ${numeroFormattato}`);
    } else {
      ui.alert('❌ Errore', `Errore: ${risultato.error}`, ui.ButtonSet.OK);
      logOperazione('ERROR', `Test invio fallito: ${risultato.error}`);
    }
    
  } catch (error) {
    ui.alert('❌ Errore', error.toString(), ui.ButtonSet.OK);
    logOperazione('ERROR', `Test invio errore: ${error.toString()}`);
  }
}

/**
 * Verifica configurazione
 */
function verificaConfigurazione() {
  console.log('🔧 VERIFICA CONFIGURAZIONE');
  console.log('========================');
  
  // Verifica API Key
  if (!CONFIG.WASENDER_API_KEY || CONFIG.WASENDER_API_KEY === 'TUA_API_KEY_QUI') {
    console.log('❌ API Key WASender NON configurata!');
    console.log('   Inserisci la tua API key in CONFIG.WASENDER_API_KEY');
  } else {
    console.log('✅ API Key WASender configurata');
    console.log(`   Primi caratteri: ${CONFIG.WASENDER_API_KEY.substring(0, 5)}...`);
  }
  
  // Verifica Device ID
  console.log(`\n📱 Device ID: ${CONFIG.WASENDER_DEVICE_ID}`);
  
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
    console.log('\n✅ Foglio "Template" presente');
    const templates = templateSheet.getDataRange().getValues();
    console.log(`   Template disponibili: ${templates.length - 1}`);
  } else {
    console.log('\n❌ Foglio "Template" MANCANTE!');
  }
  
  // Verifica Google Contacts
  if (CONFIG.USA_GOOGLE_CONTACTS) {
    try {
      const test = People.People.searchContacts({
        query: 'test',
        pageSize: 1
      });
      console.log('\n✅ Google Contacts API funzionante');
    } catch (error) {
      console.log('\n❌ Google Contacts API non configurato!');
      console.log('   Vai su Apps Script → Servizi → Aggiungi "People API"');
    }
  } else {
    console.log('\n⚠️ Google Contacts disabilitato');
  }
  
  // Verifica orario
  const orario = controllaOrarioOperativo();
  console.log(`\n⏰ Orario attuale: ${orario.oraRoma} (Roma)`);
  console.log(`   Orario operativo: ${CONFIG.ORA_INIZIO}:00 - ${CONFIG.ORA_FINE}:00`);
  console.log(`   Stato: ${orario.operativo ? '✅ OPERATIVO' : '❌ FUORI ORARIO'}`);
  
  // Verifica limiti
  console.log(`\n📊 LIMITI:`);
  console.log(`   Limite giornaliero: ${CONFIG.LIMITE_GIORNALIERO}`);
  console.log(`   Invii oggi: ${contaInviiOggi()}`);
  console.log(`   Da inviare: ${contaDaInviare()}`);
  
  console.log('\n✅ Verifica completata!');
}

/**
 * Setup iniziale - Aggiunge colonna per tracciare contatti processati
 */
function setupIniziale() {
  try {
    console.log('🔧 Avvio setup iniziale...');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clienti');
    
    if (!sheet) {
      console.log('❌ Foglio "Clienti" non trovato!');
      console.log('Crea un foglio chiamato "Clienti" e riprova.');
      return;
    }
    
    console.log('✅ Foglio "Clienti" trovato');
    
    // Verifica se la colonna L esiste già
    const headers = sheet.getRange(1, 1, 1, 12).getValues()[0];
    console.log('Headers attuali:', headers);
    
    if (!headers[11] || headers[11] !== 'Contatto Processato') {
      // Aggiungi header in colonna L
      sheet.getRange(1, 12).setValue('Contatto Processato');
      sheet.getRange(1, 12).setFontWeight('bold');
      
      console.log('✅ Setup completato!');
      console.log('Aggiunta colonna "Contatto Processato" in posizione L (colonna 12)');
      
      // Mostra alert solo se eseguito dal menu
      if (typeof ScriptApp !== 'undefined') {
        SpreadsheetApp.getUi().alert(
          'Setup completato!', 
          'Aggiunta colonna "Contatto Processato" in posizione L',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } else {
      console.log('ℹ️ Setup già completato');
      console.log('La colonna "Contatto Processato" è già presente in posizione L');
      
      // Mostra alert solo se eseguito dal menu
      if (typeof ScriptApp !== 'undefined') {
        SpreadsheetApp.getUi().alert(
          'Setup già completato', 
          'La colonna "Contatto Processato" è già presente',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
    
    console.log('🏁 Fine setup iniziale');
    
  } catch (error) {
    console.error('❌ Errore durante setup:', error);
    console.error('Stack trace:', error.stack);
  }
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
 *    - Sistema elabora automaticamente ogni 10 minuti (9-19)
 *    - Menu → "Elabora ora" per invio immediato
 *    - Menu → "Report giornaliero" per statistiche
 */

// ── SPETTACOLI APP SCRIPT v22 ──

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(err) {
    var chi = props.getProperty('lockOwner') || 'un collaboratore';
    return ContentService.createTextOutput(
      JSON.stringify({result:'locked', lockedBy: chi})
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'ai' || body.action === 'ai_pdf') {
      var apiKey = props.getProperty('anthropicKey') || '';
      if(!apiKey){ lock.releaseLock(); return resp('missing_api_key'); }

      var messages;
      if (body.action === 'ai_pdf' && body.pdf_b64) {
        // Analisi PDF con documento allegato
        messages = [{
          role: 'user',
          content: [
            {
              type: 'document',
              source: { type: 'base64', media_type: 'application/pdf', data: body.pdf_b64 }
            },
            { type: 'text', text: body.prompt }
          ]
        }];
      } else if (body.action === 'ai_pdf' && body.img_b64 && (body.img_type||'').startsWith('image/')) {
        // Analisi immagine (JPEG/PNG)
        messages = [{
          role: 'user',
          content: [
            {
              type: 'image',
              source: { type: 'base64', media_type: body.img_type || 'image/jpeg', data: body.img_b64 }
            },
            { type: 'text', text: body.prompt }
          ]
        }];
      } else if (body.action === 'ai_pdf' && body.img_b64) {
        // Analisi Excel/CSV come documento
        var excelMediaType = body.img_type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        messages = [{
          role: 'user',
          content: [
            {
              type: 'document',
              source: { type: 'base64', media_type: excelMediaType, data: body.img_b64 }
            },
            { type: 'text', text: body.prompt }
          ]
        }];
      } else {
        messages = [{role:'user', content: body.prompt}];
      }

      var aiRes = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'pdfs-2024-09-25'
        },
        payload: JSON.stringify({
          model: 'claude-opus-4-8',
          max_tokens: 16000,
          messages: messages
        }),
        muteHttpExceptions: true
      });
      var aiJson = JSON.parse(aiRes.getContentText());
      var text = '';
      if(aiJson.content && aiJson.content[0]){
        text = aiJson.content[0].text || '';
      } else if(aiJson.error){
        // Errore API Anthropic — restituisce dettaglio
        lock.releaseLock();
        return ContentService.createTextOutput(JSON.stringify({result:'', error: aiJson.error.type+': '+aiJson.error.message})).setMimeType(ContentService.MimeType.JSON);
      }
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({result: text})).setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'syncConfig') {
      props.setProperty('teatriDB', JSON.stringify(body.teatriDB||{}));
      if(body.manifCustom) props.setProperty('manifCustom', JSON.stringify(body.manifCustom));
      if(body.dittaDB) props.setProperty('dittaDB', JSON.stringify(body.dittaDB));
      if(body.collaboratori) props.setProperty('collaboratori', JSON.stringify(body.collaboratori));
      if(body.spazioInfo) props.setProperty('spazioInfo', JSON.stringify(body.spazioInfo));
      if(body.aziendeDB) props.setProperty('aziendeDB', JSON.stringify(body.aziendeDB));
      lock.releaseLock();
      return resp('config ok');
    }

    if (body.action === 'sendEmail') {
      var to      = body.to || '';
      var subject = body.subject || '';
      var message = body.message || '';
      if(!to || !subject || !message){ lock.releaseLock(); return resp('missing fields'); }
      MailApp.sendEmail({
        to: to,
        subject: subject,
        body: message,
        htmlBody: message.split(String.fromCharCode(10)).join('<br>')
      });
      lock.releaseLock();
      return resp('sent');
    }
    if (body.action === 'applicaAnagrafica') {
      var ra; try { ra = applicaAnagrafica(body.remaps); } catch(errA) { ra = { error: errA.message }; }
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify(ra)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'uploadFile') {
      lock.releaseLock();
      var ru; try { ru = caricaDocumento(body.cartella, body.nome, body.mime, body.b64); } catch(errU) { ru = { error: errU.message }; }
      return ContentService.createTextOutput(JSON.stringify(ru)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'leggiScheda') {
      lock.releaseLock();   // la lettura dura anche un minuto: non bloccare le sync dei colleghi
      var rs; try { rs = leggiSchedaThread(body.link, body.fonte); } catch(errL) { rs = { error: errL.message }; }
      return ContentService.createTextOutput(JSON.stringify(rs)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'updateProposte') {
      var shU = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Proposte');
      if(!shU){ lock.releaseLock(); return resp('no sheet'); }
      var rU = shU.getDataRange().getValues();
      var hU = rU[0].map(function(h){ return String(h).trim().toUpperCase(); });
      var cId = hU.indexOf('ID'), cSt = hU.indexOf('STATO');
      var mappa = body.updates || {};
      var n = 0;
      for(var i=1;i<rU.length;i++){
        var id = String(rU[i][cId]);
        if(mappa[id]){ shU.getRange(i+1, cSt+1).setValue(mappa[id]); n++; }
      }
      lock.releaseLock();
      return resp('updated ' + n);
    }
    if (body.action === 'syncCollab') {
      if(body.collaboratori) props.setProperty('collaboratori', JSON.stringify(body.collaboratori));
      lock.releaseLock();
      return resp('collab ok');
    }

    if (body.action !== 'sync') { lock.releaseLock(); return resp('unknown action'); }

    var autore = body.autore || 'Sconosciuto';
    var ts = body.timestamp || new Date().getTime();
    props.setProperty('lockOwner', autore);
    props.setProperty('lastModified', String(ts));
    props.setProperty('lastAuthor', autore);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Spettacoli') || ss.insertSheet('Spettacoli');
    var header = ['ID','DATA','CITTA','TEATRO','SPETTACOLO','MANIFESTAZIONE',
      'N.MACC','TAR.MACC','SUB.MACC','N.ELETT','TAR.ELETT','SUB.ELETT',
      'N.FACC','TAR.FACC','SUB.FACC','N.FONI','TAR.FONI','SUB.FONI',
      'N.SART','TAR.SART','SUB.SART','TOTALE','NOTE','FATTURATO','PREMONTAGGIO','COLLAB_JSON','REPLICA_OF','MAILS_JSON','CONTATTI_JSON'];
    sh.clearContents();
    sh.getRange(1,1,1,header.length).setValues([header]);
    var rows = body.data.map(function(r,i) {
      return [r.id||i+1,r.data,r.citta,r.teatro,r.spettacolo,r.manifestazione,
        r.n_macc,r.tariffa_macc,r.sub_macc,
        r.n_elett,r.tariffa_elett,r.sub_elett,
        r.n_facc,r.tariffa_facc,r.sub_facc,
        r.n_foni||0,r.tariffa_foni||0,r.sub_foni||0,
        r.n_sart||0,r.tariffa_sart||0,r.sub_sart||0,
        r.totale,r.note,r.fatturato,r.premontaggio||'NO',r.collab_json||'',r.replica_of||'',r.mails_json||'',r.contatti_json||''];
    });
    if(rows.length) sh.getRange(2,1,rows.length,header.length).setValues(rows);
    return resp('ok: '+rows.length+' righe');

  } catch(err) {
    return resp('error: '+err.message);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = e.parameter.action;
  var props = PropertiesService.getScriptProperties();

  if(action === 'status') {
    var payload = JSON.stringify({
      lastModified: parseInt(props.getProperty('lastModified')||'0'),
      lastAuthor: props.getProperty('lastAuthor')||''
    });
    var cb = e.parameter.callback;
    if(cb) return ContentService.createTextOutput(cb+'('+payload+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  }
  if(action === 'auditAnagrafica') {
    try { return ContentService.createTextOutput(JSON.stringify(auditAnagrafica())).setMimeType(ContentService.MimeType.JSON); }
    catch(err) { return ContentService.createTextOutput(JSON.stringify({error: err.message})).setMimeType(ContentService.MimeType.JSON); }
  }
  if(action === 'rubrica') {
    try { return ContentService.createTextOutput(JSON.stringify({rubrica: leggiRubrica(e.parameter.refresh === '1')})).setMimeType(ContentService.MimeType.JSON); }
    catch(err) { return ContentService.createTextOutput(JSON.stringify({error: err.message})).setMimeType(ContentService.MimeType.JSON); }
  }
  if(action === 'proposte') {
    try {
      var shP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Proposte');
      if(!shP) return ContentService.createTextOutput(JSON.stringify({proposte:[]})).setMimeType(ContentService.MimeType.JSON);
      var rP = shP.getDataRange().getValues();
      var hP = rP[0].map(function(h){ return String(h).trim().toLowerCase(); });
      var out = [];
      for(var i=1;i<rP.length;i++){
        var o = {};
        hP.forEach(function(h,k){
          var v = rP[i][k];
          if(v instanceof Date) v = Utilities.formatDate(v,'GMT+1','yyyy-MM-dd');
          o[h] = v;
        });
        if(String(o.stato).toLowerCase()==='pending') out.push(o);
      }
      return ContentService.createTextOutput(JSON.stringify({proposte:out})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({error:err.message})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  if(action === 'pull') {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName('Spettacoli');
      if(!sh) {
        return ContentService.createTextOutput(JSON.stringify({data:[],lastAuthor:'',teatriDB:{},manifCustom:[],dittaDB:[],collaboratori:[]})).setMimeType(ContentService.MimeType.JSON);
      }
      var rows = sh.getDataRange().getValues();
      var header = rows[0];
      var data = rows.slice(1).map(function(r) {
        var obj = {};
        header.forEach(function(h,i){
          var val = r[i];
          if(val instanceof Date) {
            val = val.getFullYear()+'-'+String(val.getMonth()+1).padStart(2,'0')+'-'+String(val.getDate()).padStart(2,'0');
          }
          obj[h.toString().trim().toLowerCase().replace(/\./g,'_')] = val;
        });
        return obj;
      }).filter(function(r){ return r['data'] && r['spettacolo']; });

      var teatriDBraw = props.getProperty('teatriDB');
      var teatriDBobj = teatriDBraw ? JSON.parse(teatriDBraw) : {};
      var manifRaw = props.getProperty('manifCustom');
      var manifCustom = manifRaw ? JSON.parse(manifRaw) : [];

      var result = JSON.stringify({
        data: data,
        lastAuthor: props.getProperty('lastAuthor')||'',
        lastModified: parseInt(props.getProperty('lastModified')||'0'),
        teatriDB: teatriDBobj,
        manifCustom: manifCustom,
        dittaDB: JSON.parse(props.getProperty('dittaDB')||'[]'),
        collaboratori: JSON.parse(props.getProperty('collaboratori')||'[]'),
        spazioInfo: JSON.parse(props.getProperty('spazioInfo')||'{}'),
        aziendeDB: JSON.parse(props.getProperty('aziendeDB')||'[]'),
        anagraficaVer: parseInt(props.getProperty('anagraficaVer')||'0'),
        anagraficaRemaps: JSON.parse(props.getProperty('anagraficaRemaps')||'[]')
      });
      return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({error: err.message})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({result:'unknown action'})).setMimeType(ContentService.MimeType.JSON);
}

function resp(msg) {
  return ContentService.createTextOutput(
    JSON.stringify({result:msg})
  ).setMimeType(ContentService.MimeType.JSON);
}

function test() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Spettacoli') || ss.insertSheet('Spettacoli');
  sh.getRange(1,1).setValue('TEST OK');
}
function testLetturaWord() {
  var threads = GmailApp.search('from:(r.cantarini@amat.marche.it OR m.osimani@amat.marche.it OR d.sepe@amat.marche.it) has:attachment', 0, 20);
  for (var t = 0; t < threads.length; t++) {
    var msgs = threads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var all = msgs[m].getAttachments();
      for (var i = 0; i < all.length; i++) {
        var a = all[i];
        var nome = a.getName();
        if (!/\.docx?$/i.test(nome)) continue;
        Logger.log('=== ' + msgs[m].getSubject() + ' → ' + nome);
          var file = Drive.Files.create(
          {name: 'tmp_gf', mimeType: 'application/vnd.google-apps.document'},
          a.copyBlob()
        );
        Logger.log(DocumentApp.openById(file.id).getBody().getText().substring(0, 3000));
        DriveApp.getFileById(file.id).setTrashed(true);
      }
    }
  }
}
function testEstrazione() {
  var props = PropertiesService.getScriptProperties();
  var KEY = props.getProperty('anthropicKey');

  // Costruisce l anagrafica citta/teatri dalle Properties
  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var anagrafica = Object.keys(tdb).sort().map(function(c) {
    return c + ': ' + tdb[c].join(', ');
  }).join('\n');
  if (!anagrafica) anagrafica = '(vuota)';
  Logger.log('--- ANAGRAFICA ---\n' + anagrafica + '\n');

  var threads = GmailApp.search('from:(r.cantarini@amat.marche.it OR m.osimani@amat.marche.it OR d.sepe@amat.marche.it) has:attachment', 0, 20);

  for (var t = 0; t < threads.length; t++) {
    var msgs = threads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var all = msgs[m].getAttachments();
      for (var i = 0; i < all.length; i++) {
        var a = all[i];
        if (!/\.docx?$/i.test(a.getName())) continue;

        var file = Drive.Files.create(
          {name: 'tmp_gf', mimeType: 'application/vnd.google-apps.document'}, a.copyBlob());
        var testo = DocumentApp.openById(file.id).getBody().getText();
        DriveApp.getFileById(file.id).setTrashed(true);

        Logger.log('===== ' + a.getName());
        Logger.log(estraiCalendario(testo, KEY, anagrafica));
      }
    }
  }
}

function estraiCalendario(testo, KEY, anagrafica) {
  var prompt =
'Analizza questo documento e rispondi SOLO con JSON, senza markdown ne commenti.\n\n' +
'ANAGRAFICA ESISTENTE (citta e spazi gia in archivio):\n' + anagrafica + '\n\n' +
'PASSO 1 - Classifica il documento nel campo "tipo":\n' +
'  "calendario"  = calendario di stagione, festival o rassegna con date e spettacoli\n' +
'  "scheda"      = scheda tecnica di uno spettacolo (audio, luci, palco)\n' +
'  "altro"       = qualsiasi altra cosa (normative, piani di emergenza, ecc.)\n' +
'Se tipo non e "calendario", rispondi solo {"tipo":"..."} e fermati.\n\n' +
'PASSO 2 - Riconosci il FORMATO del calendario:\n' +
'  A) STAGIONE TEATRALE: tutti gli spettacoli sono nello stesso teatro, che e\n' +
'     nominato nelle prime righe del documento.\n' +
'  B) FESTIVAL / RASSEGNA: gli spettacoli si tengono in luoghi diversi (chiese,\n' +
'     piazze, cortili, teatri vari), indicati serata per serata nel testo.\n\n' +
'PASSO 3 - Estrai:\n' +
'{"tipo":"calendario","formato":"stagione|festival","citta":"...",\n' +
' "spazio":"...","manifestazione_generale":"...","stagione":"26/27",\n' +
' "spettacoli":[{"data":"AAAA-MM-GG","titolo":"...","spazio":"...",\n' +
'   "produzione":"...","manifestazione":"...","repliche":["AAAA-MM-GG"],\n' +
'   "premontaggio":"AAAA-MM-GG|null","note":"..."}]}\n\n' +
'REGOLE SUGLI SPAZI (le piu importanti):\n' +
'- Se formato = "stagione": metti il teatro nel campo "spazio" generale e lascia\n' +
'  VUOTO il campo "spazio" dentro ogni spettacolo.\n' +
'- Se formato = "festival": lascia VUOTO il campo "spazio" generale e metti in\n' +
'  OGNI spettacolo il suo luogo specifico, ricavandolo dal testo della serata\n' +
'  (es. "CHIESA DI SANT AGOSTINO", "PIAZZA UMBERTO I").\n' +
'- Il nome di un festival o di una rassegna (es. ME.MO.FEST, BUONA ESTATE,\n' +
'  MIRALTEATRO, TAU) NON e uno spazio: va in "manifestazione", MAI in "spazio".\n' +
'- Se per uno spettacolo non riesci a determinare il luogo, lascia "spazio" vuoto.\n' +
'  NON inventare e non usare il nome del festival come ripiego.\n' +
'- Gli spazi vanno sempre in MAIUSCOLO. Se corrispondono a uno gia in anagrafica,\n' +
'  usa ESATTAMENTE la grafia dell anagrafica.\n\n' +
'ALTRE REGOLE:\n' +
'- CITTA: sempre MAIUSCOLO, grafia dell anagrafica se gia presente.\n' +
'- ANNO: se le date non hanno anno e c e una stagione tipo 26/27, allora mesi da\n' +
'  SETTEMBRE a DICEMBRE = 2026 e mesi da GENNAIO ad AGOSTO = 2027. Ignora la\n' +
'  posizione nel documento. Se l anno e scritto esplicitamente, usa quello.\n' +
'- TITOLO: riportalo ESATTAMENTE come stampato nel documento, articoli compresi.\n' +
'  Non aggiungere e non togliere "IL", "LA", "LO".\n' +
'- "4 e 5 novembre" = un solo spettacolo con data del 4 e repliche ["...-11-05"].\n' +
'- Due recite lo STESSO giorno (es. "ore 17.30 e ore 19") NON sono repliche:\n' +
'  restano un solo spettacolo, scrivi gli orari nelle note.\n' +
'- "teatro a disposizione dal 12 ottobre" = premontaggio "2026-10-12".\n' +
'- manifestazione = sezione o rassegna a cui appartiene lo spettacolo. Per una\n' +
'  stagione senza intestazione di sezione usa "STAGIONE DI PROSA".\n' +
'- IGNORA: abbonamenti, prezzi, riduzioni, trailer, link, sinossi lunghe.\n\n' +
'DOCUMENTO:\n' + testo;

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {'Content-Type':'application/json', 'x-api-key': KEY, 'anthropic-version':'2023-06-01'},
    payload: JSON.stringify({model:'claude-opus-4-8', max_tokens:16000,
      messages:[{role:'user', content: prompt}]}),
    muteHttpExceptions: true
  });
  var j = JSON.parse(res.getContentText());
  if (j.error) return 'ERRORE API: ' + j.error.message;
  return j.content[0].text;
}
function diagnosiAnagrafica() {
  var props = PropertiesService.getScriptProperties();
  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');

  // Conta gli spettacoli per coppia citta|teatro
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spettacoli');
  var rows = sh.getDataRange().getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iC = hdr.indexOf('CITTA'), iT = hdr.indexOf('TEATRO');
  var conta = {};
  for (var r = 1; r < rows.length; r++) {
    var k = String(rows[r][iC]).trim().toUpperCase() + '|' + String(rows[r][iT]).trim().toUpperCase();
    conta[k] = (conta[k] || 0) + 1;
  }

  var NON_TEATRI = /^(AUDITORIUM|CHIESA|SALA|PINACOTECA|VILLA|ROTONDA|RIDOTTO|CINETEATRO|ARENA|PALAZZO|CORTILE|PIAZZA)/;

  function canonico(nome) {
    var n = nome.trim().toUpperCase().replace(/\s+/g, ' ');
    if (NON_TEATRI.test(n)) return n;              // non e un teatro: lascia com e
    n = n.replace(/^TEATRO\s+/, '');                // togli TEATRO iniziale
    return 'TEATRO ' + n;                           // e rimettilo in forma canonica
  }

  Logger.log('=== PROPOSTA NORMALIZZAZIONE ===\n');
  Object.keys(tdb).sort().forEach(function(citta) {
    var visti = {};
    var righe = [];
    tdb[citta].forEach(function(teatro) {
      var can = canonico(teatro);
      var n = conta[citta.toUpperCase() + '|' + teatro.toUpperCase()] || 0;
      righe.push('     "' + teatro + '"  (' + n + ' spett.)  ->  "' + can + '"');
      visti[can] = (visti[can] || 0) + 1;
    });
    var dup = Object.keys(visti).filter(function(k){ return visti[k] > 1; });
    Logger.log(citta + (dup.length ? '   <<< DOPPIONI: ' + dup.join(' / ') : ''));
    righe.forEach(function(r){ Logger.log(r); });
  });

  // Citta simili tra loro
  Logger.log('\n=== CITTA SOSPETTE ===');
  var cs = Object.keys(tdb);
  for (var i = 0; i < cs.length; i++) {
    for (var j = i + 1; j < cs.length; j++) {
      var a = cs[i].replace(/[^A-Z]/gi, '').toUpperCase();
      var b = cs[j].replace(/[^A-Z]/gi, '').toUpperCase();
      if (a.indexOf(b) === 0 || b.indexOf(a) === 0) Logger.log(cs[i] + '  <->  ' + cs[j]);
    }
  }

  // Spazi generici da rinominare a mano
  Logger.log('\n=== NOMI GENERICI (serve il nome vero) ===');
  Object.keys(tdb).forEach(function(c) {
    tdb[c].forEach(function(t) {
      if (/^(TEATRO|TEATRO COMUNALE|COMUNALE)$/i.test(t.trim()))
        Logger.log(c + ': "' + t + '"');
    });
  });
}
var DRY_RUN = false;   // metti false SOLO dopo aver controllato il log

function normalizzaAnagrafica() {
  var props = PropertiesService.getScriptProperties();

  // Casi speciali: chiave "CITTA ORIGINALE|TEATRO ORIGINALE"
  var OVERRIDE = {
    "ASCOLI PICENO|SAN PIETRO IN CASTELLO": "SAN PIETRO IN CASTELLO",
    "ASCOLI PICENO|SANT'ANDREA":            "SANT'ANDREA",
    "ASCOLI PICENO|FILARMONICI":            "TEATRO DEI FILARMONICI",
    "FANO|FORTUNA":                         "TEATRO DELLA FORTUNA",
    "FANO|TEATRO DELLA FORTUNA":            "TEATRO DELLA FORTUNA",
    "CORINALDO|TEATRO":                     "TEATRO GOLDONI",
    "SAN SEVERINO|FERONIA":                 "TEATRO FERONIA",
    "SAN SEVERINO MARCHE|TEATRO":           "TEATRO FERONIA",
    "SANTELPIDIO A MARE|CICCONI":           "TEATRO CICCONI",
    "SANT'ELPIDIO A MARE|TEATRO":           "TEATRO CICCONI"
  };
  var CITTA_MAP = {
    "SAN SEVERINO":       "SAN SEVERINO MARCHE",
    "SANTELPIDIO A MARE": "SANT'ELPIDIO A MARE",
    "PORTO SANTELPIDIO":  "PORTO SANT'ELPIDIO"
  };
  var NON_TEATRI = /^(AUDITORIUM|CHIESA|SALA|PINACOTECA|VILLA|ROTONDA|RIDOTTO|ARENA|PALAZZO|CORTILE|PIAZZA)/;

  function normCitta(c) {
    var n = String(c).trim().toUpperCase().replace(/\s+/g,' ');
    return CITTA_MAP[n] || n;
  }
  function normTeatro(citta, teatro) {
    var c = String(citta).trim().toUpperCase().replace(/\s+/g,' ');
    var t = String(teatro).trim().toUpperCase().replace(/\s+/g,' ');
    if (!t) return '';
    var ov = OVERRIDE[c + '|' + t];
    if (ov) return ov;
    if (NON_TEATRI.test(t)) return t;
    return 'TEATRO ' + t.replace(/^TEATRO\s+/, '');
  }

  // ---- 1. teatriDB ----
  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var nuovoTdb = {};
  Logger.log('=== TEATRI ===');
  Object.keys(tdb).forEach(function(citta) {
    var cN = normCitta(citta);
    tdb[citta].forEach(function(teatro) {
      var tN = normTeatro(citta, teatro);
      if (!nuovoTdb[cN]) nuovoTdb[cN] = [];
      if (nuovoTdb[cN].indexOf(tN) < 0) nuovoTdb[cN].push(tN);
      if (cN !== citta.toUpperCase() || tN !== teatro.toUpperCase())
        Logger.log(citta + ' | ' + teatro + '   ->   ' + cN + ' | ' + tN);
    });
  });
  Object.keys(nuovoTdb).forEach(function(c){ nuovoTdb[c].sort(); });

  // ---- 2. foglio Spettacoli ----
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spettacoli');
  var rng = sh.getDataRange(), rows = rng.getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iC = hdr.indexOf('CITTA'), iT = hdr.indexOf('TEATRO');
  var modificate = 0;
  for (var r = 1; r < rows.length; r++) {
    var oc = rows[r][iC], ot = rows[r][iT];
    if (!oc) continue;
    var nc = normCitta(oc), nt = normTeatro(oc, ot);
    if (nc !== String(oc).trim() || nt !== String(ot).trim()) {
      rows[r][iC] = nc; rows[r][iT] = nt; modificate++;
    }
  }
  Logger.log('\nRighe foglio da modificare: ' + modificate + ' su ' + (rows.length-1));

  // ---- 3. spazioInfo (chiavi "CITTA||TEATRO") ----
  var si = JSON.parse(props.getProperty('spazioInfo') || '{}');
  var nuovoSi = {}, siMod = 0;
  Logger.log('\n=== SCHEDE SPAZIO ===');
  Object.keys(si).forEach(function(k) {
    var p = k.split('||');
    var nk = normCitta(p[0]) + '||' + normTeatro(p[0], p[1] || '');
    nuovoSi[nk] = si[k];
    if (nk !== k) { Logger.log(k + '   ->   ' + nk); siMod++; }
  });
  if (!siMod) Logger.log('(nessuna da rimappare)');

  if (DRY_RUN) { Logger.log('\n*** DRY RUN — non ho scritto niente ***'); return; }

  props.setProperty('teatriDB', JSON.stringify(nuovoTdb));
  props.setProperty('spazioInfo', JSON.stringify(nuovoSi));
  rng.setValues(rows);
  Logger.log('\n*** SCRITTO: ' + modificate + ' righe, ' + siMod + ' schede spazio ***');
}
function diagnosiManifestazioni() {
  var props = PropertiesService.getScriptProperties();
  var manif = JSON.parse(props.getProperty('manifCustom') || '[]');

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spettacoli');
  var rows = sh.getDataRange().getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iM = hdr.indexOf('MANIFESTAZIONE');

  var conta = {};
  for (var r = 1; r < rows.length; r++) {
    var m = String(rows[r][iM] || '').trim();
    if (!m) continue;
    conta[m] = (conta[m] || 0) + 1;
  }

  Logger.log('=== USATE NEGLI SPETTACOLI ===');
  Object.keys(conta).sort(function(a,b){ return conta[b]-conta[a]; }).forEach(function(m) {
    var inLista = manif.indexOf(m) >= 0 ? '' : '   [NON in manifCustom]';
    Logger.log('  ' + conta[m] + '\t"' + m + '"' + inLista);
  });

  Logger.log('\n=== IN ANAGRAFICA MA MAI USATE ===');
  manif.forEach(function(m) { if (!conta[m]) Logger.log('  "' + m + '"'); });

  Logger.log('\n=== POSSIBILI DOPPIONI ===');
  var tutte = Object.keys(conta).concat(manif.filter(function(m){ return !conta[m]; }));
  var norm = function(s){ return s.toUpperCase().replace(/[^A-Z0-9]/g,''); };
  for (var i = 0; i < tutte.length; i++) {
    for (var j = i+1; j < tutte.length; j++) {
      var a = norm(tutte[i]), b = norm(tutte[j]);
      if (a === b || a.indexOf(b) === 0 || b.indexOf(a) === 0)
        Logger.log('  "' + tutte[i] + '"  <->  "' + tutte[j] + '"');
    }
  }
}
var DRY_MANIF = false;   // metti false dopo aver controllato il log

function normalizzaManifestazioni() {
  var props = PropertiesService.getScriptProperties();

  var MAP = {
    "STAGIONE  DI PROSA":  "STAGIONE DI PROSA",
    "STAGIONE TEATRALE":   "STAGIONE DI PROSA",
    "STAGIONE DANZA":      "STAGIONE DI DANZA",
    "STAGIONE RAGAZZI":    "STAGIONE TEATRO RAGAZZI",
    "TEATRO RAGAZZI":      "STAGIONE TEATRO RAGAZZI",
    "VILLAINVITA 2026":    "VILLAINVITA"
  };

  var CANONICHE = [
    "STAGIONE DI PROSA", "STAGIONE DI DANZA", "STAGIONE TEATRO RAGAZZI",
    "FUORI ABBONAMENTO", "GESTIONE", "TEATRO OLTRE", "PLAYLIST",
    "VILLAINVITA", "LORETO IN SCENA", "SINFONIE D'AUTUNNO",
    "RESIDENZE MULTIDISCIPLINARI", "ENERGIE VIVE", "SCENE D'ARTISTA",
    "CIVITANOVA DANZA", "DANZA KID"
  ];

  function norm(m) {
    var n = String(m || '').trim().toUpperCase().replace(/\s+/g, ' ');
    if (!n) return '';
    return MAP[String(m).trim().toUpperCase()] || MAP[n] || n;
  }

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spettacoli');
  var rng = sh.getDataRange(), rows = rng.getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iM = hdr.indexOf('MANIFESTAZIONE');

  var mod = 0, riepilogo = {};
  for (var r = 1; r < rows.length; r++) {
    var o = rows[r][iM];
    if (!o) continue;
    var n = norm(o);
    if (n !== String(o).trim()) {
      var k = '"' + String(o).trim() + '"  ->  "' + n + '"';
      riepilogo[k] = (riepilogo[k] || 0) + 1;
      rows[r][iM] = n;
      mod++;
    }
  }

  Logger.log('=== TRASFORMAZIONI ===');
  Object.keys(riepilogo).sort().forEach(function(k){ Logger.log('  ' + riepilogo[k] + '\t' + k); });
  Logger.log('\nRighe da modificare: ' + mod);
  Logger.log('\n=== manifCustom che verra scritta ===');
  CANONICHE.slice().sort().forEach(function(m){ Logger.log('  ' + m); });

  if (DRY_MANIF) { Logger.log('\n*** DRY RUN — non ho scritto niente ***'); return; }

  props.setProperty('manifCustom', JSON.stringify(CANONICHE.slice().sort()));
  rng.setValues(rows);
  Logger.log('\n*** SCRITTO: ' + mod + ' righe + manifCustom (' + CANONICHE.length + ' voci) ***');
}
var DRY_SCAN = false;   // false solo dopo aver controllato il log

var DRY_SCAN = false;
var DA_DATA  = '2026/01/01';
var MAX_MIN  = 4.5;

function scansionaMail() {
  var t0 = Date.now();
  var props = PropertiesService.getScriptProperties();
  var KEY = props.getProperty('anthropicKey');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function pulisci(s) {
    return String(s || '').replace(/[\u2018\u2019\u02BC]/g, "'")
      .replace(/[\u2013\u2014]/g, '-').replace(/\s+/g, ' ').trim();
  }
  function chiave(data, citta, titolo) {
    return String(data) + '|' + pulisci(citta).toUpperCase() + '|' + pulisci(titolo).toUpperCase();
  }

  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var anagrafica = Object.keys(tdb).sort().map(function(c) {
    return c + ': ' + tdb[c].join(', ');
  }).join('\n');

  var shS = ss.getSheetByName('Spettacoli');
  var rowsS = shS.getDataRange().getValues();
  var hS = rowsS[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iD = hS.indexOf('DATA'), iC = hS.indexOf('CITTA'), iSp = hS.indexOf('SPETTACOLO');
  var esistenti = {};
  for (var r = 1; r < rowsS.length; r++) {
    var d = rowsS[r][iD];
    if (d instanceof Date) d = Utilities.formatDate(d, 'GMT+1', 'yyyy-MM-dd');
    esistenti[chiave(d, rowsS[r][iC], rowsS[r][iSp])] = true;
  }

  var shP = _preparaProposte(ss);
  if (false) {
    shP = ss.insertSheet('Proposte');
    shP.getRange(1,1,1,14).setValues([['ID','TS','TIPO','STATO','FONTE',
      'CITTA','SPAZIO','DATA','TITOLO','PRODUZIONE','MANIFESTAZIONE',
      'REPLICHE','PREMONTAGGIO','NOTE']]);
  } else {
    var rP = shP.getDataRange().getValues();
    var hP = rP[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var pC = hP.indexOf('CITTA'), pD = hP.indexOf('DATA'), pT = hP.indexOf('TITOLO');
    for (var q = 1; q < rP.length; q++) {
      var dp = rP[q][pD];
      if (dp instanceof Date) dp = Utilities.formatDate(dp, 'GMT+1', 'yyyy-MM-dd');
      esistenti[chiave(dp, rP[q][pC], rP[q][pT])] = true;
    }
  }
  Logger.log('Chiavi note: ' + Object.keys(esistenti).length);

  var label = GmailApp.getUserLabelByName('GF-vista') || GmailApp.createLabel('GF-vista');
  var threads = GmailApp.search('from:(r.cantarini@amat.marche.it OR m.osimani@amat.marche.it OR d.sepe@amat.marche.it) has:attachment after:' + DA_DATA + ' -label:GF-vista', 0, 300);
  Logger.log('Thread da esaminare: ' + threads.length + '\n');

  var scritte = 0, dup = 0, fatti = 0, senzaSpazio = 0, interrotto = false;

  for (var t = 0; t < threads.length; t++) {
    if ((Date.now() - t0) / 60000 > MAX_MIN) { interrotto = true; break; }

    var nuove = [];
    var msgs = threads[t].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var all = msgs[m].getAttachments();
      for (var i = 0; i < all.length; i++) {
        var a = all[i];
        if (!/\.docx?$/i.test(a.getName())) continue;

        var file = Drive.Files.create(
          {name:'tmp_gf', mimeType:'application/vnd.google-apps.document'}, a.copyBlob());
        var testo = DocumentApp.openById(file.id).getBody().getText();
        DriveApp.getFileById(file.id).setTrashed(true);
        fatti++;

        var out = estraiCalendario(testo, KEY, anagrafica);
        var j;
        try { j = JSON.parse(out.replace(/```json|```/g,'').trim()); }
        catch(e) { Logger.log('!! JSON non valido: ' + a.getName()); continue; }
        if (j.tipo !== 'calendario') { Logger.log('- ' + a.getName() + ' -> ' + j.tipo); continue; }

        Logger.log('+ ' + a.getName() + ' [' + (j.formato||'?') + '] -> ' +
                   j.citta + ' : ' + j.spettacoli.length);

        j.spettacoli.forEach(function(sp) {
          var k = chiave(sp.data, j.citta, sp.titolo);
          if (esistenti[k]) { dup++; return; }
          esistenti[k] = true;
          var spazio = pulisci(sp.spazio || j.spazio || '').toUpperCase();
          if (!spazio) senzaSpazio++;
          var manif = sp.manifestazione || j.manifestazione_generale || '';
          nuove.push([
            'P' + Date.now() + '_' + Math.floor(Math.random()*10000),
            new Date(), 'nuovo_spettacolo', 'pending',
            msgs[m].getSubject() + ' | ' + a.getName(),
            pulisci(j.citta).toUpperCase(), spazio, sp.data, pulisci(sp.titolo),
            sp.produzione || '', manif, (sp.repliche || []).join(', '),
            sp.premontaggio || '', sp.note || '', threads[t].getPermalink()
          ]);
        });
      }
    }

    if (!DRY_SCAN) {
      if (nuove.length) {
        shP.getRange(shP.getLastRow()+1, 1, nuove.length, 15).setValues(nuove);
        scritte += nuove.length;
      }
      threads[t].addLabel(label);
    } else { scritte += nuove.length; }
  }

  Logger.log('\n=== RISULTATO ===');
  Logger.log('Word elaborati: ' + fatti);
  Logger.log('Proposte scritte: ' + scritte + '  |  gia note: ' + dup + '  |  senza spazio: ' + senzaSpazio);
  Logger.log('Thread rimasti: ' + (threads.length - t));
  Logger.log(interrotto ? '\n>>> Fermato per tempo. RILANCIA per continuare.' : '\n*** Arretrato completato ***');
}
function scansionaSettimanale() {
  var out = [];
  try { scansionaMail(); out.push('Allegati: scansione eseguita'); }
  catch(e){ out.push('Errore allegati: ' + e.message); }
  try { out.push(analizzaCorpoMail()); }
  catch(e){ out.push('Errore corpo mail: ' + e.message); }
  try { out.push(scansionaSchede()); }
  catch(e){ out.push('Errore schede: ' + e.message); }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shP = ss.getSheetByName('Proposte');
  var pend = 0;
  if (shP) {
    var rP = shP.getDataRange().getValues();
    var iSt = rP[0].map(function(h){ return String(h).trim().toUpperCase(); }).indexOf('STATO');
    for (var i=1;i<rP.length;i++) if (String(rP[i][iSt]).toLowerCase()==='pending') pend++;
  }
  if (pend > 0) {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'GammaFucker — ' + pend + ' proposte da controllare',
      body: 'Scansione settimanale completata.\n\n' + out.join('\n') +
            '\n\nIn attesa di approvazione: ' + pend +
            '\n\nApri l app, tab Proposte.'
    });
  }
}

function analizzaCorpoMail() {
  var props = PropertiesService.getScriptProperties();
  var KEY = props.getProperty('anthropicKey');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var anagrafica = Object.keys(tdb).sort().map(function(c){
    return c + ': ' + tdb[c].join(', ');
  }).join('\n');

  var shS = ss.getSheetByName('Spettacoli');
  var rS = shS.getDataRange().getValues();
  var hS = rS[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iD=hS.indexOf('DATA'), iC=hS.indexOf('CITTA'), iT=hS.indexOf('SPETTACOLO');
  var oggi = new Date(); var limite = new Date(oggi.getTime() + 365*86400000);
  var elenco = [];
  for (var r=1;r<rS.length;r++){
    var d = rS[r][iD];
    if (d instanceof Date) { if (d < oggi || d > limite) continue;
      d = Utilities.formatDate(d,'GMT+1','yyyy-MM-dd'); }
    else { if (!d || String(d) < Utilities.formatDate(oggi,'GMT+1','yyyy-MM-dd')) continue; }
    elenco.push(d + ' | ' + rS[r][iC] + ' | ' + rS[r][iT]);
  }

  var lbl = GmailApp.getUserLabelByName('GF-corpo') || GmailApp.createLabel('GF-corpo');
  var threads = GmailApp.search('from:(r.cantarini@amat.marche.it OR m.osimani@amat.marche.it OR d.sepe@amat.marche.it) newer_than:60d -label:GF-corpo', 0, 30);
  if (!threads.length) return 'Corpo mail: nessuna nuova';

  var mails = [];
  threads.forEach(function(t){
    var ms = t.getMessages();
    var m = ms[ms.length-1];
    mails.push({ og: m.getSubject(),
                 link: t.getPermalink(),
                 tx: m.getPlainBody().substring(0,1200).replace(/\s+/g,' ') });
  });

  var prompt =
'Sei un assistente che smista mail di programmazione teatrale.\n' +
'Rispondi SOLO con un array JSON, senza markdown.\n\n' +
'CITTA E SPAZI IN ARCHIVIO:\n' + anagrafica + '\n\n' +
'SPETTACOLI GIA IN CALENDARIO (data | citta | titolo):\n' + elenco.join('\n') + '\n\n' +
'Per OGNI mail restituisci un oggetto:\n' +
'{"i":<indice>,"tipo":"...","urgenza":"alta|media|bassa","riassunto":"...",\n' +
' "data":"AAAA-MM-GG|null","citta":"...|null","spazio":"...|null","riferimento":"...|null"}\n\n' +
'TIPI:\n' +
' "nuovo_evento"     = annuncia uno spettacolo NON presente nell elenco sopra\n' +
' "modifica_evento"  = cambia data, orario o spazio di uno gia in elenco\n' +
' "adempimento"      = richiede un azione (convocazione VVF, libretto ZTL, permessi)\n' +
' "info"             = riferisce a uno spettacolo in elenco senza chiedere nulla\n' +
' "irrilevante"      = non riguarda programmazione o spazi\n\n' +
'REGOLE:\n' +
'- "riferimento": se la mail parla di uno spettacolo dell elenco, copia qui la sua\n' +
'  riga esatta. Altrimenti null.\n' +
'- "riassunto": una frase, max 20 parole, che dica cosa devo fare o sapere.\n' +
'- urgenza alta = c e una scadenza o un adempimento; bassa = solo informativa.\n' +
'- Sii selettivo: se non riguarda spettacoli, spazi o adempimenti, e "irrilevante".\n\n' +
'MAIL:\n' +
mails.map(function(m,i){ return '['+i+'] OGGETTO: '+m.og+'\nTESTO: '+m.tx; }).join('\n\n');

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:'post',
    headers:{'Content-Type':'application/json','x-api-key':KEY,'anthropic-version':'2023-06-01'},
    payload: JSON.stringify({model:'claude-opus-4-8', max_tokens:8000,
      messages:[{role:'user', content:prompt}]}),
    muteHttpExceptions:true
  });
  var jr = JSON.parse(res.getContentText());
  if (jr.error) return 'Corpo mail — errore API: ' + jr.error.message;

  var arr;
  try { arr = JSON.parse(jr.content[0].text.replace(/```json|```/g,'').trim()); }
  catch(e){ return 'Corpo mail: JSON non valido'; }

  var shP = _preparaProposte(ss);
  var righe = [];
  arr.forEach(function(o){
    if (o.tipo === 'irrilevante' || o.tipo === 'info') return;
    var m = mails[o.i] || {og:'', link:''};
    righe.push(['S'+Date.now()+'_'+Math.floor(Math.random()*10000), new Date(),
      'segnalazione', 'pending', m.og,
      o.citta || '', o.spazio || '', o.data || '', o.tipo.toUpperCase().replace('_',' '),
      '', o.urgenza || '', '', '',
      o.riassunto + (o.riferimento ? '  [rif: ' + o.riferimento + ']' : ''),
      m.link || '']);
  });
  if (righe.length) shP.getRange(shP.getLastRow()+1,1,righe.length,15).setValues(righe);
  threads.forEach(function(t){ t.addLabel(lbl); });

  return 'Corpo mail: ' + mails.length + ' esaminate, ' + righe.length + ' segnalazioni';
}

// Crea il foglio Proposte se manca e garantisce la colonna LINK
function _preparaProposte(ss) {
  var sh = ss.getSheetByName('Proposte');
  var intest = ['ID','TS','TIPO','STATO','FONTE','CITTA','SPAZIO','DATA','TITOLO',
                'PRODUZIONE','MANIFESTAZIONE','REPLICHE','PREMONTAGGIO','NOTE','LINK','DETTAGLI_JSON','FILE_URL'];
  if (!sh) {
    sh = ss.insertSheet('Proposte');
    sh.getRange(1,1,1,intest.length).setValues([intest]);
    return sh;
  }
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
            .map(function(x){ return String(x).trim().toUpperCase(); });
  ['LINK','DETTAGLI_JSON','FILE_URL'].forEach(function(c){
    if (h.indexOf(c) < 0) { sh.getRange(1, h.length+1).setValue(c); h.push(c); }
  });
  return sh;
}
function testCorpoMail() {
  Logger.log(analizzaCorpoMail());
}
// ════════════════════ SCHEDE TECNICHE ════════════════════
// Cerca mail (qualsiasi mittente) con allegati PDF/Word che sembrano schede
// tecniche, le fa leggere a Claude e scrive una proposta 'scheda_tecnica'
// nel foglio Proposte con i dettagli in DETTAGLI_JSON. Nessuna scrittura
// sugli spettacoli: l'avvallo avviene nell'app.
var SCHEDE_GIORNI   = 120;    // finestra di ricerca mail (le schede arrivano mesi prima)
var SCHEDE_MAX_MIN  = 3.5;    // guardia tempo: oltre non si iniziano nuove letture (ogni PDF ~30-60 s)
var SCHEDE_MAX_LETTURE = 8;   // allegati letti per esecuzione; il resto al giro dopo
var SCHEDE_MAX_MB   = 15;     // allegati più grandi vengono saltati
var SCHEDE_CARTELLA = 'GammaFucker Schede';

function _ctxSchede() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var futuri = _spettacoliFuturi(ss);
  var shP = _preparaProposte(ss);
  var hP = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var col = function(n){ return hP.indexOf(n); };
  // già in Proposte (qualsiasi stato): stessa mail + stesso allegato non si rilegge
  var giaLetti = {};
  var rP = shP.getDataRange().getValues();
  for (var q = 1; q < rP.length; q++) if (rP[q][col('TIPO')] === 'scheda_tecnica')
    giaLetti[rP[q][col('LINK')] + '|' + rP[q][col('FONTE')]] = { riga: q + 1, stato: String(rP[q][col('STATO')] || '').toLowerCase() };
  return {
    KEY: props.getProperty('anthropicKey'), t0: Date.now(),
    elenco: futuri.map(function(s){ return s.riga; }),
    titoliNorm: futuri.map(function(s){ return _normTitoloGS(s.titolo); }).filter(function(t){ return t.length >= 5; }),
    shP: shP, hP: hP, col: col, giaLetti: giaLetti,
    lbl: GmailApp.getUserLabelByName('GF-scheda') || GmailApp.createLabel('GF-scheda'),
    esaminati: 0, scritte: 0, scartate: 0, errori: 0, giaSaltati: 0, ripristinate: 0, interrotto: false
  };
}

function scansionaSchede() {
  var ctx = _ctxSchede();
  if (!ctx.KEY) return 'Schede: anthropicKey mancante';
  var threads = GmailApp.search('has:attachment newer_than:' + SCHEDE_GIORNI + 'd -label:GF-scheda (filename:pdf OR filename:docx OR filename:doc)', 0, 100);
  Logger.log('Schede: ' + threads.length + ' thread con allegati da filtrare, ' + ctx.elenco.length + ' spettacoli futuri');
  for (var t = 0; t < threads.length && !ctx.interrotto; t++) _elaboraThreadSchede(threads[t], ctx, false);
  var out = 'Schede: ' + ctx.esaminati + ' allegati letti, ' + ctx.scritte + ' proposte, ' + ctx.scartate + ' non schede, ' + ctx.errori + ' errori (verranno ritentati)' +
            (ctx.interrotto ? ' — FERMATO (tempo o tetto letture): rilancia per continuare' : ' — completato');
  Logger.log(out);
  return out;
}

// Lettura forzata di un singolo thread (dall'app: pulsante sulla segnalazione).
// Ignora il filtro "sembra scheda" e l'etichetta: legge tutti i PDF/Word del thread.
function leggiSchedaThread(link, fonte) {
  var th = null;
  var id = String(link || '').split(/[\/#]/).pop();
  if (id) { try { th = GmailApp.getThreadById(id); } catch(e) {} }
  if (!th && fonte) {
    // segnalazioni vecchie senza link: cerca per oggetto (senza prefissi I:/R:/Fwd:)
    var ogg = String(fonte).split(' | ')[0].replace(/^\s*((I|R|RE|FW|FWD|TR|Fwd|Re|I):\s*)+/i, '').replace(/"/g, '').trim();
    if (ogg.length > 6) {
      var found = GmailApp.search('subject:"' + ogg.substring(0, 80) + '" has:attachment newer_than:365d', 0, 3);
      if (found.length) th = found[0];
    }
  }
  if (!th) return { error: 'mail non trovata (né link né oggetto corrispondente)' };
  var ctx = _ctxSchede();
  if (!ctx.KEY) return { error: 'anthropicKey mancante' };
  _elaboraThreadSchede(th, ctx, true);
  return { letti: ctx.esaminati, proposte: ctx.scritte, non_schede: ctx.scartate, errori: ctx.errori, gia_lette: ctx.giaSaltati, ripristinate: ctx.ripristinate };
}

// Un thread: filtra gli allegati, li fa leggere, scrive le righe. Etichetta il thread
// solo se tutti i candidati sono stati elaborati senza errori.
function _elaboraThreadSchede(thread, ctx, forza) {
  var msgs = thread.getMessages();
  var candidati = [];
  for (var m = 0; m < msgs.length; m++) {
    var atts = msgs[m].getAttachments();
    var oggetto = msgs[m].getSubject() || '';
    for (var i = 0; i < atts.length; i++) {
      var a = atts[i];
      var nome = a.getName() || '';
      if (!/\.(pdf|docx?)$/i.test(nome)) continue;
      if (a.getSize() > SCHEDE_MAX_MB * 1024 * 1024) { Logger.log('  salto (troppo grande): ' + nome); continue; }
      if (forza || _sembraScheda(oggetto, nome, ctx.titoliNorm)) candidati.push({msg: msgs[m], att: a, oggetto: oggetto, nome: nome});
    }
  }
  if (!candidati.length) return;   // non etichettato: se lo spettacolo entra in calendario dopo, verrà riesaminato

  var link = thread.getPermalink();
  var righe = [];
  var completo = true, conErrori = false;
  for (var c = 0; c < candidati.length; c++) {
    var k = candidati[c];
    var gia = ctx.giaLetti[link + '|' + k.oggetto + ' | ' + k.nome];
    if (gia) {
      // già letta: in lettura forzata, se era stata scartata/approvata la rimetto in attesa senza rileggerla
      if (forza && gia.stato !== 'pending') { ctx.shP.getRange(gia.riga, ctx.col('STATO') + 1).setValue('pending'); gia.stato = 'pending'; ctx.ripristinate++; }
      else ctx.giaSaltati++;
      continue;
    }
    if (!forza && ((Date.now() - ctx.t0) / 60000 > SCHEDE_MAX_MIN || ctx.esaminati >= SCHEDE_MAX_LETTURE)) { ctx.interrotto = true; completo = false; break; }
    ctx.esaminati++;
    Logger.log('  leggo [' + ctx.esaminati + '] ' + k.nome + '  (' + Math.round((Date.now()-ctx.t0)/1000) + 's)');
    var j = _analizzaScheda(ctx.KEY, ctx.elenco, k);
    if (!j) { ctx.errori++; conErrori = true; continue; }          // errore API/JSON: NON etichettare, si ritenta
    if (!j.e_scheda_tecnica) { ctx.scartate++; continue; }

    var match = _parseRiga(j.match);
    var fileUrl = '';
    try { fileUrl = _salvaAllegato(k.att, (match ? match.titolo : j.titolo) || 'scheda'); } catch(e) { Logger.log('Drive: ' + e.message); }

    var dettagli = {
      titolo: j.titolo || '', compagnia: j.compagnia || '',
      match: j.match || null, match_confidenza: j.match_confidenza || 'bassa',
      orari: j.orari || {}, durata: j.durata || null,
      tecnici: j.tecnici || {}, tecnici_note: j.tecnici_note || '',
      esigenze: j.esigenze || [], confidenza: j.confidenza || {},
      riassunto: j.riassunto || '', mittente: k.msg.getFrom(), data_mail: Utilities.formatDate(k.msg.getDate(), 'Europe/Rome', 'yyyy-MM-dd'),
      contatti: _contattiScheda(k.msg, j.contatti)
    };
    var col = ctx.col, riga = [];
    riga[col('ID')] = 'T' + Date.now() + '_' + Math.floor(Math.random()*10000);
    riga[col('TS')] = new Date();
    riga[col('TIPO')] = 'scheda_tecnica';
    riga[col('STATO')] = 'pending';
    riga[col('FONTE')] = k.oggetto + ' | ' + k.nome;
    riga[col('CITTA')] = match ? match.citta : '';
    riga[col('SPAZIO')] = match ? match.teatro : '';
    riga[col('DATA')] = match ? match.data : '';
    riga[col('TITOLO')] = match ? match.titolo : (j.titolo || '');
    riga[col('PRODUZIONE')] = j.compagnia || '';
    riga[col('MANIFESTAZIONE')] = '';
    riga[col('REPLICHE')] = '';
    riga[col('PREMONTAGGIO')] = '';
    riga[col('NOTE')] = j.riassunto || '';
    riga[col('LINK')] = link;
    riga[col('DETTAGLI_JSON')] = JSON.stringify(dettagli);
    riga[col('FILE_URL')] = fileUrl;
    for (var z = 0; z < ctx.hP.length; z++) if (riga[z] === undefined) riga[z] = '';
    righe.push(riga);
    ctx.giaLetti[link + '|' + k.oggetto + ' | ' + k.nome] = { riga: -1, stato: 'pending' };
  }
  if (righe.length) { ctx.shP.getRange(ctx.shP.getLastRow()+1, 1, righe.length, ctx.hP.length).setValues(righe); ctx.scritte += righe.length; }
  if (completo && !conErrori) thread.addLabel(ctx.lbl);
}

// Toglie l'etichetta GF-scheda da tutti i thread: al prossimo giro vengono riesaminati.
// Le schede già trasformate in proposta non vengono rilette (dedupe su mail+allegato).
function azzeraEtichettaSchede() {
  var lbl = GmailApp.getUserLabelByName('GF-scheda');
  if (!lbl) return 'Etichetta assente';
  var th = lbl.getThreads(0, 500), n = 0;
  th.forEach(function(t){ t.removeLabel(lbl); n++; });
  Logger.log('Rimossa etichetta da ' + n + ' thread'); return n;
}

// Contatti della compagnia: mittente della mail (o mittente originale se inoltrata
// da un collega AMAT) + referenti trovati nella scheda. Dedupe per email.
function _contattiScheda(msg, daScheda) {
  var out = [], visti = {};
  function add(nome, ruolo, email, tel) {
    email = String(email || '').trim().toLowerCase();
    tel = String(tel || '').trim();
    if (!email && !tel) return;
    if (email && visti[email]) return;
    if (email) visti[email] = true;
    out.push({ nome: String(nome || '').replace(/"/g, '').trim(), ruolo: ruolo || '', email: email, telefono: tel });
  }
  function parseFrom(str) {
    var m = String(str || '').match(/^\s*"?([^"<]*)"?\s*<([^>]+)>/);
    if (m) return { nome: m[1].trim(), email: m[2].trim() };
    var e = String(str || '').match(/[\w.+-]+@[\w.-]+\.\w+/);
    return e ? { nome: '', email: e[0] } : null;
  }
  var from = parseFrom(msg.getFrom());
  var interno = from && /@amat\.marche\.it$/i.test(from.email);
  if (interno) {
    // mail inoltrata: cerco l'intestazione "Da:"/"From:" del messaggio originale nel corpo
    var corpo = (msg.getPlainBody() || '').substring(0, 4000);
    var mm = corpo.match(/(?:^|\n)\s*(?:Da|From|De)\s*:\s*(.+)/i);
    var orig = mm ? parseFrom(mm[1]) : null;
    if (orig && !/@amat\.marche\.it$/i.test(orig.email)) add(orig.nome, 'mittente', orig.email, '');
  } else if (from) add(from.nome, 'mittente', from.email, '');
  (daScheda || []).forEach(function(c){ if (c) add(c.nome, c.ruolo, c.email, c.telefono); });
  return out;
}

// Spettacoli da oggi a +12 mesi, con la riga "data | citta | teatro | titolo" usata nel prompt
function _spettacoliFuturi(ss) {
  var sh = ss.getSheetByName('Spettacoli');
  if (!sh) return [];
  var rows = sh.getDataRange().getValues();
  var h = rows[0].map(function(x){ return String(x).trim().toUpperCase(); });
  var iD = h.indexOf('DATA'), iC = h.indexOf('CITTA'), iT = h.indexOf('TEATRO'), iS = h.indexOf('SPETTACOLO');
  var oggi = Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd');
  var lim = Utilities.formatDate(new Date(Date.now() + 365*86400000), 'Europe/Rome', 'yyyy-MM-dd');
  var out = [];
  for (var r = 1; r < rows.length; r++) {
    var d = rows[r][iD];
    if (d instanceof Date) d = Utilities.formatDate(d, 'Europe/Rome', 'yyyy-MM-dd');
    d = String(d || '');
    if (!d || d < oggi || d > lim) continue;
    var o = { data: d, citta: String(rows[r][iC]||''), teatro: String(rows[r][iT]||''), titolo: String(rows[r][iS]||'') };
    o.riga = o.data + ' | ' + o.citta + ' | ' + o.teatro + ' | ' + o.titolo;
    out.push(o);
  }
  return out;
}

function _normTitoloGS(t) {
  return String(t || '').replace(/[‘’ʼ]/g, "'").replace(/[–—]/g, '-')
    .toUpperCase().replace(/^(IL|LO|LA|I|GLI|LE|L')\s+/, '').replace(/[^A-Z0-9 ]/g, '').replace(/\s+/g, ' ').trim();
}

function _sembraScheda(oggetto, nome, titoliNorm) {
  var txt = _normTitoloGS(oggetto + ' ' + nome.replace(/\.[a-z0-9]+$/i, ''));
  if (/SCHEDA TECNICA|SCHEDA TEC|RIDER|TECHNICAL|SCH TECNICA|SCHEDATECNICA/.test(txt)) return true;
  for (var i = 0; i < titoliNorm.length; i++) if (txt.indexOf(titoliNorm[i]) >= 0) return true;
  return false;
}

function _parseRiga(riga) {
  if (!riga) return null;
  var p = String(riga).split('|').map(function(x){ return x.trim(); });
  if (p.length < 4 || !/^\d{4}-\d{2}-\d{2}$/.test(p[0])) return null;
  return { data: p[0], citta: p[1], teatro: p[2], titolo: p.slice(3).join(' | ') };
}

function _salvaAllegato(att, titolo) {
  var it = DriveApp.getFoldersByName(SCHEDE_CARTELLA);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(SCHEDE_CARTELLA);
  var nome = (titolo || 'scheda').replace(/[\\/:*?"<>|]/g, ' ').trim().substring(0, 60) + ' - ' + att.getName();
  var f = folder.createFile(att.copyBlob()).setName(nome);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return f.getUrl();
}

// Manda l'allegato a Claude: PDF come documento, Word convertito in testo
function _analizzaScheda(KEY, elenco, k) {
  var prompt =
'Sei l assistente tecnico di AMAT Marche (circuito teatrale). Ti passo un allegato ricevuto via mail.\n' +
'Oggetto mail: ' + k.oggetto + '\nNome file: ' + k.nome + '\n\n' +
'SPETTACOLI IN CALENDARIO (data | citta | teatro | titolo):\n' + (elenco.length ? elenco.join('\n') : '(nessuno)') + '\n\n' +
'Rispondi SOLO con un oggetto JSON, senza markdown:\n' +
'{"e_scheda_tecnica": true|false,\n' +
' "titolo": "titolo dello spettacolo|null", "compagnia": "compagnia o produzione|null",\n' +
' "match": "riga ESATTA dell elenco a cui si riferisce|null", "match_confidenza": "alta|media|bassa",\n' +
' "orari": {"arrivo":"HH:MM|null","scarico":"HH:MM|null","montaggio":"HH:MM|null","prove":"HH:MM|null","spettacolo":"HH:MM|null","smontaggio":"HH:MM|null"},\n' +
' "durata": "durata spettacolo|null",\n' +
' "tecnici": {"macchinisti":n|null,"elettricisti":n|null,"fonici":n|null,"sarte":n|null,"facchini":n|null},\n' +
' "tecnici_note": "dettagli sul personale richiesto al teatro (turni, orari, qualifiche)|",\n' +
' "esigenze": ["richieste particolari: potenza elettrica, dimensioni minime palco, carichi, graticcio, ballerine, fumo, ecc."],\n' +
' "confidenza": {"orari":"alta|media|bassa","tecnici":"alta|media|bassa"},\n' +
' "contatti": [{"nome":"...","ruolo":"direttore tecnico|direttore di scena|produzione|tour manager|altro","email":"...|null","telefono":"...|null"}],\n' +
' "riassunto": "una frase, max 25 parole"}\n\n' +
'REGOLE:\n' +
'- e_scheda_tecnica = true solo se il documento e una scheda tecnica / rider di uno spettacolo.\n' +
'- "tecnici" = personale che la compagnia CHIEDE AL TEATRO (non quello che porta con se). Se non specificato, null.\n' +
'- Se la scheda distingue montaggio e spettacolo, riporta i numeri del montaggio in tecnici e i turni in tecnici_note.\n' +
'- "match": scegli la riga dell elenco solo se titolo (anche parziale) e compatibile; se ci sono piu date dello stesso titolo scegli la prima futura e dillo in riassunto. Se nessuna riga e compatibile, null.\n' +
'- Orari nel formato HH:MM; se e indicata solo una durata scrivila in durata.\n' +
'- "contatti": le persone della compagnia indicate nel documento (referenti tecnici, produzione) con email/telefono se presenti; [] se nessuna.\n' +
'- Non inventare: campo non presente = null.';

  var content = [];
  if (/\.pdf$/i.test(k.nome)) {
    content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: Utilities.base64Encode(k.att.getBytes()) } });
  } else {
    var testo = '';
    try {
      var file = Drive.Files.create({ name: 'tmp_gf_scheda', mimeType: 'application/vnd.google-apps.document' }, k.att.copyBlob());
      testo = DocumentApp.openById(file.id).getBody().getText();
      DriveApp.getFileById(file.id).setTrashed(true);
    } catch(e) { Logger.log('Word non leggibile: ' + k.nome + ' ' + e.message); return null; }
    if (!testo.trim()) return null;
    content.push({ type: 'text', text: 'TESTO DEL DOCUMENTO:\n' + testo.substring(0, 60000) });
  }
  content.push({ type: 'text', text: prompt });

  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'x-api-key': KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: 'claude-opus-4-8', max_tokens: 3000, messages: [{ role: 'user', content: content }] }),
    muteHttpExceptions: true
  });
  var jr; try { jr = JSON.parse(res.getContentText()); } catch(e) { return null; }
  if (jr.error) { Logger.log('API: ' + jr.error.message); return null; }
  try { return JSON.parse(jr.content[0].text.replace(/```json|```/g, '').trim()); }
  catch(e) { Logger.log('JSON scheda non valido: ' + k.nome); return null; }
}

function testSchede() { Logger.log(scansionaSchede()); }

// ════════════════════ RUBRICA (Contatti Google) ════════════════════
// Legge i contatti dell'account del foglio tramite People API (servizio avanzato
// "People" in appsscript.json). Cache 6 ore; ?refresh=1 la salta.
function leggiRubrica(refresh) {
  var cache = CacheService.getScriptCache();
  if (!refresh) {
    var parts = [], i = 0, chunk;
    while ((chunk = cache.get('rubrica_' + i)) !== null) { parts.push(chunk); i++; }
    if (parts.length) return JSON.parse(parts.join(''));
  }
  var out = [], token = null;
  do {
    var res = People.People.Connections.list('people/me', {
      personFields: 'names,emailAddresses,phoneNumbers,organizations',
      pageSize: 1000, pageToken: token, sortOrder: 'FIRST_NAME_ASCENDING'
    });
    (res.connections || []).forEach(function(p) {
      var nome = (p.names && p.names[0] && p.names[0].displayName) || '';
      var emails = (p.emailAddresses || []).map(function(x){ return String(x.value || '').trim(); }).filter(Boolean);
      var tels = (p.phoneNumbers || []).map(function(x){ return String(x.value || '').trim(); }).filter(Boolean);
      var org = (p.organizations && p.organizations[0] && p.organizations[0].name) || '';
      if (!emails.length && !tels.length) return;
      out.push({ nome: String(nome || org || emails[0] || tels[0] || ''), email: emails[0] || '', tel: tels[0] || '', emails: emails, tels: tels, org: org });
    });
    token = res.nextPageToken;
  } while (token);
  // "Altri contatti": persone a cui hai scritto senza salvarle in rubrica
  try {
    var tok2 = null;
    do {
      var r2 = People.OtherContacts.list({ readMask: 'names,emailAddresses,phoneNumbers', pageSize: 1000, pageToken: tok2 });
      (r2.otherContacts || []).forEach(function(p) {
        var nome = (p.names && p.names[0] && p.names[0].displayName) || '';
        var emails = (p.emailAddresses || []).map(function(x){ return String(x.value || '').trim(); }).filter(Boolean);
        if (!emails.length) return;
        if (out.some(function(o){ return o.emails.indexOf(emails[0]) >= 0; })) return;
        out.push({ nome: String(nome || emails[0]), email: emails[0], tel: '', emails: emails, tels: [], org: '', altro: true });
      });
      tok2 = r2.nextPageToken;
    } while (tok2);
  } catch(e) { Logger.log('OtherContacts: ' + e.message); }
  out.sort(function(a, b){ return String(a.nome || '').localeCompare(String(b.nome || ''), 'it'); });
  // cache a pezzi (limite 100 KB per chiave)
  var json = JSON.stringify(out), n = 0;
  for (var k = 0; k < json.length; k += 90000) cache.put('rubrica_' + (n++), json.substring(k, k + 90000), 21600);
  return out;
}
function testRubrica() { var r = leggiRubrica(true); Logger.log(r.length + ' contatti; primo: ' + JSON.stringify(r[0])); }

// ════════════════════ DOCUMENTI (upload su Drive) ════════════════════
// Salva un file in "GammaFucker Documenti/<cartella>" con link di sola lettura
// per chi ha il link. Sostituisce il copia-incolla dei link Dropbox.
var DOC_CARTELLA = 'GammaFucker Documenti';
function caricaDocumento(cartella, nome, mime, b64) {
  if (!nome || !b64) return { error: 'file mancante' };
  var it = DriveApp.getFoldersByName(DOC_CARTELLA);
  var root = it.hasNext() ? it.next() : DriveApp.createFolder(DOC_CARTELLA);
  var sub = String(cartella || 'Varie').replace(/[\\/:*?"<>|]/g, ' ').trim().substring(0, 80) || 'Varie';
  var it2 = root.getFoldersByName(sub);
  var folder = it2.hasNext() ? it2.next() : root.createFolder(sub);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', nome);
  var f = folder.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return { url: f.getUrl(), id: f.getId(), nome: f.getName(), cartella: sub };
}

// ════════════════════ ANAGRAFICA: verifica e unificazione ════════════════════
// Diagnosi (GET ?action=auditAnagrafica) e applicazione (POST applicaAnagrafica).
// Ogni rinomina applicata viene registrata in un "giornale" (anagraficaRemaps, con
// versione crescente) che il pull consegna alle app: ogni dispositivo la applica ai
// propri dati locali, così la normalizzazione non viene annullata dai merge local-wins.
var NON_TEATRI_RE = /^(AUDITORIUM|CHIESA|SALA|PINACOTECA|VILLA|ROTONDA|RIDOTTO|CINETEATRO|CINEMA|ARENA|PALAZZO|CORTILE|PIAZZA|PALAS|PALASPORT|CASTELLO|ABBAZIA|MUSEO|BIBLIOTECA|CENTRO|SPAZIO|GIARDIN)/;
function _canonTeatro(t) {
  var n = String(t || '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (!n || NON_TEATRI_RE.test(n)) return n;
  return 'TEATRO ' + n.replace(/^TEATRO\s+/, '');
}
function _keyManif(m) {
  // chiave di raggruppamento: maiuscolo, via anni, "EDIZIONE", punteggiatura, spazi
  return String(m || '').toUpperCase().replace(/\b(19|20)\d\d(\s*[\/\-]\s*(19|20)?\d\d)?\b/g, '')
    .replace(/\b(EDIZIONE|ED\.?|STAGIONE)\b/g, '').replace(/[^A-Z0-9]/g, '');
}

function auditAnagrafica() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Spettacoli');
  var rows = sh ? sh.getDataRange().getValues() : [[]];
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iC = hdr.indexOf('CITTA'), iT = hdr.indexOf('TEATRO'), iM = hdr.indexOf('MANIFESTAZIONE');
  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var manif = JSON.parse(props.getProperty('manifCustom') || '[]');

  var usiM = {}, usiT = {};
  for (var r = 1; r < rows.length; r++) {
    var m = String(rows[r][iM] || '').trim(); if (m) usiM[m] = (usiM[m] || 0) + 1;
    var c = String(rows[r][iC] || '').trim().toUpperCase(), t = String(rows[r][iT] || '').trim().toUpperCase();
    if (c) usiT[c + '|' + t] = (usiT[c + '|' + t] || 0) + 1;
  }

  // ── manifestazioni: gruppi per chiave, o per prefisso (APP / APP ASCOLI) ──
  var nomi = Object.keys(usiM);
  manif.forEach(function(m){ if (nomi.indexOf(m) < 0) nomi.push(m); });
  var byKey = {};
  nomi.forEach(function(n){ var k = _keyManif(n); if (!k) return; (byKey[k] = byKey[k] || []).push(n); });
  var keys = Object.keys(byKey);
  // unisci chiavi in relazione di prefisso (una parola almeno di 3 lettere)
  var parent = {}; keys.forEach(function(k){ parent[k] = k; });
  function find(k){ while (parent[k] !== k) k = parent[k]; return k; }
  for (var i = 0; i < keys.length; i++) for (var j = i + 1; j < keys.length; j++) {
    var a = keys[i], b = keys[j];
    if (a.length >= 3 && b.length >= 3 && (a.indexOf(b) === 0 || b.indexOf(a) === 0)) parent[find(a)] = find(b);
  }
  var cluster = {};
  keys.forEach(function(k){ var root = find(k); (cluster[root] = cluster[root] || []).push.apply(cluster[root], byKey[k]); });
  var gruppiManif = [];
  Object.keys(cluster).forEach(function(root) {
    var vs = cluster[root]; if (vs.length < 2) return;
    var varianti = vs.map(function(n){ return { nome: n, usi: usiM[n] || 0, inLista: manif.indexOf(n) >= 0 }; })
                     .sort(function(x, y){ return (y.usi - x.usi) || (x.nome.length - y.nome.length); });
    // canonico proposto: il più usato; a parità il più corto (senza anno)
    var can = varianti[0].nome;
    var senzaAnno = varianti.filter(function(v){ return !/(19|20)\d\d/.test(v.nome); });
    if (senzaAnno.length) can = senzaAnno.sort(function(x, y){ return (y.usi - x.usi) || (x.nome.length - y.nome.length); })[0].nome;
    gruppiManif.push({ canonico: can, varianti: varianti });
  });
  var manifTutte = nomi.map(function(n){ return { nome: n, usi: usiM[n] || 0, inLista: manif.indexOf(n) >= 0 }; })
                       .sort(function(x, y){ return y.usi - x.usi; });

  // ── teatri: per città, gruppi di varianti dello stesso spazio ──
  // "MUSE", "TEATRO MUSE", "TEATRO DELLE MUSE" → un gruppo; canonico = forma "TEATRO …"
  // già esistente più usata, altrimenti la forma canonica della variante più lunga.
  function core(t){ return _canonTeatro(t).replace(/^TEATRO\s+/, '').replace(/\b(DEL|DELLA|DELLE|DEI|DEGLI|DELLO|DI|D')\b/g, ' ').replace(/\s+/g, ' ').trim(); }
  var gruppiTeatri = [];
  Object.keys(tdb).sort().forEach(function(citta) {
    var lista = (tdb[citta] || []).map(function(t){ return { nome: t, can: _canonTeatro(t), core: core(t), usi: usiT[citta.toUpperCase() + '|' + String(t).toUpperCase()] || 0 }; });
    var par = lista.map(function(_, i){ return i; });
    function f(i){ while (par[i] !== i) i = par[i]; return i; }
    for (var a = 0; a < lista.length; a++) for (var b = a + 1; b < lista.length; b++) {
      var A = lista[a], B = lista[b];
      if (NON_TEATRI_RE.test(A.can) !== NON_TEATRI_RE.test(B.can)) continue;
      var same = A.can === B.can || (A.core.length >= 4 && B.core.length >= 4 && (new RegExp('(^| )' + A.core.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '( |$)').test(B.core) || new RegExp('(^| )' + B.core.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '( |$)').test(A.core)));
      if (same) par[f(a)] = f(b);
    }
    var gr = {};
    lista.forEach(function(x, i){ (gr[f(i)] = gr[f(i)] || []).push(x); });
    Object.keys(gr).forEach(function(k) {
      var vs = gr[k];
      var giaCanoniche = vs.filter(function(v){ return v.nome === v.can; });
      var nonConformi = vs.filter(function(v){ return v.nome !== v.can; });
      if (vs.length < 2 && !nonConformi.length) return;   // già a posto
      var can;
      if (giaCanoniche.length) can = giaCanoniche.sort(function(x, y){ return (y.usi - x.usi) || (y.nome.length - x.nome.length); })[0].can;
      else can = vs.slice().sort(function(x, y){ return (y.nome.length - x.nome.length) || (y.usi - x.usi); })[0].can;
      gruppiTeatri.push({ citta: citta, canonico: can, varianti: vs.sort(function(x, y){ return y.usi - x.usi; }).map(function(v){ return { nome: v.nome, usi: v.usi }; }) });
    });
  });
  // città simili (prefisso / senza apostrofi)
  var cittaSospette = [];
  var cs = Object.keys(tdb);
  function nc(s){ return s.toUpperCase().replace(/[^A-Z]/g, ''); }
  for (var p = 0; p < cs.length; p++) for (var q = p + 1; q < cs.length; q++) {
    var A = nc(cs[p]), B = nc(cs[q]);
    if (A === B || A.indexOf(B) === 0 || B.indexOf(A) === 0) {
      var uA = 0, uB = 0; Object.keys(usiT).forEach(function(k){ if (k.split('|')[0] === cs[p].toUpperCase()) uA += usiT[k]; if (k.split('|')[0] === cs[q].toUpperCase()) uB += usiT[k]; });
      cittaSospette.push({ a: cs[p], usiA: uA, b: cs[q], usiB: uB, teatriA: (tdb[cs[p]]||[]).length, teatriB: (tdb[cs[q]]||[]).length });
    }
  }
  // usati negli spettacoli ma assenti in anagrafica
  var orfani = [];
  Object.keys(usiT).forEach(function(k) {
    var pp = k.split('|'); var c = pp[0], t = pp[1];
    var lista = (tdb[c] || []).map(function(x){ return String(x).toUpperCase(); });
    if (!tdb[c] || (t && lista.indexOf(t) < 0)) orfani.push({ citta: c, teatro: t, usi: usiT[k], cittaNota: !!tdb[c] });
  });
  orfani.sort(function(x, y){ return y.usi - x.usi; });

  return {
    manifestazioni: manifTutte, gruppiManif: gruppiManif,
    gruppiTeatri: gruppiTeatri, cittaSospette: cittaSospette, orfani: orfani,
    totSpettacoli: rows.length - 1, anagraficaVer: parseInt(props.getProperty('anagraficaVer') || '0')
  };
}

// remaps: [{tipo:'manif', da, a} | {tipo:'teatro', citta, da, a} | {tipo:'citta', da, a} | {tipo:'aggiungiTeatro', citta, teatro}]
function applicaAnagrafica(remaps) {
  if (!remaps || !remaps.length) return { error: 'niente da applicare' };
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tdb = JSON.parse(props.getProperty('teatriDB') || '{}');
  var manif = JSON.parse(props.getProperty('manifCustom') || '[]');
  var si = JSON.parse(props.getProperty('spazioInfo') || '{}');
  var U = function(s){ return String(s || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var mapM = {}, mapC = {}, mapT = {}, aggiunte = [];
  remaps.forEach(function(x) {
    if (x.tipo === 'manif' && x.da && x.a && x.da !== x.a) mapM[String(x.da).trim()] = String(x.a).trim();
    else if (x.tipo === 'citta' && x.da && x.a && U(x.da) !== U(x.a)) mapC[U(x.da)] = U(x.a);
    else if (x.tipo === 'teatro' && x.citta && x.da && x.a && U(x.da) !== U(x.a)) mapT[U(x.citta) + '|' + U(x.da)] = U(x.a);
    else if (x.tipo === 'aggiungiTeatro' && x.citta) aggiunte.push({ citta: U(x.citta), teatro: U(x.teatro) });
  });
  function fixCitta(c){ var u = U(c); return mapC[u] || u; }
  function fixTeatro(c, t){ var u = U(t); var m = mapT[U(c) + '|' + u]; return m || u; }   // usa la città ORIGINALE come chiave

  var nRighe = 0;
  ['Spettacoli', 'Proposte'].forEach(function(nome) {
    var sh = ss.getSheetByName(nome); if (!sh) return;
    var rng = sh.getDataRange(), rows = rng.getValues();
    var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iC = hdr.indexOf('CITTA'), iT = hdr.indexOf(nome === 'Spettacoli' ? 'TEATRO' : 'SPAZIO'), iM = hdr.indexOf('MANIFESTAZIONE');
    var mod = false;
    for (var r = 1; r < rows.length; r++) {
      var ch = false;
      if (iM >= 0) { var m = String(rows[r][iM] || '').trim(); if (m && mapM[m] && mapM[m] !== m) { rows[r][iM] = mapM[m]; ch = true; } }
      if (iC >= 0 && rows[r][iC]) {
        var oc = rows[r][iC], ot = iT >= 0 ? rows[r][iT] : '';
        var nc2 = fixCitta(oc), nt = iT >= 0 ? fixTeatro(oc, ot) : ot;
        if (nc2 !== String(oc)) { rows[r][iC] = nc2; ch = true; }
        if (iT >= 0 && nt !== String(ot)) { rows[r][iT] = nt; ch = true; }
      }
      if (ch) { nRighe++; mod = true; }
    }
    if (mod) rng.setValues(rows);
  });

  // teatriDB
  var nuovo = {};
  Object.keys(tdb).forEach(function(c) {
    var nc3 = fixCitta(c);
    (tdb[c] || []).forEach(function(t) { var nt = fixTeatro(c, t); if (!nt) return; (nuovo[nc3] = nuovo[nc3] || []); if (nuovo[nc3].indexOf(nt) < 0) nuovo[nc3].push(nt); });
    if (!nuovo[nc3]) nuovo[nc3] = [];
  });
  aggiunte.forEach(function(a){ (nuovo[a.citta] = nuovo[a.citta] || []); if (a.teatro && nuovo[a.citta].indexOf(a.teatro) < 0) nuovo[a.citta].push(a.teatro); });
  Object.keys(nuovo).forEach(function(c){ nuovo[c].sort(); });
  // manifCustom
  var nm = [];
  manif.forEach(function(m){ var v = mapM[m] || m; if (nm.indexOf(v) < 0) nm.push(v); });
  Object.keys(mapM).forEach(function(k){ if (nm.indexOf(mapM[k]) < 0) nm.push(mapM[k]); });
  nm.sort();
  // spazioInfo
  var nsi = {};
  Object.keys(si).forEach(function(k) {
    var pp = k.split('||'); var nk = fixCitta(pp[0]) + '||' + fixTeatro(pp[0], pp[1] || '');
    if (nsi[nk] && nk !== k) { // fondi: tieni i campi compilati
      var a = nsi[nk], b = si[k]; Object.keys(b).forEach(function(f){ if (!a[f] || (Array.isArray(a[f]) && !a[f].length)) a[f] = b[f]; });
    } else nsi[nk] = si[k];
  });

  var ver = parseInt(props.getProperty('anagraficaVer') || '0') + 1;
  var giornale = JSON.parse(props.getProperty('anagraficaRemaps') || '[]');
  remaps.forEach(function(x){ x.ver = ver; giornale.push(x); });
  if (giornale.length > 300) giornale = giornale.slice(-300);
  props.setProperty('teatriDB', JSON.stringify(nuovo));
  props.setProperty('manifCustom', JSON.stringify(nm));
  props.setProperty('spazioInfo', JSON.stringify(nsi));
  props.setProperty('anagraficaRemaps', JSON.stringify(giornale));
  props.setProperty('anagraficaVer', String(ver));
  props.setProperty('lastModified', String(Date.now()));
  return { ok: true, righe: nRighe, ver: ver, teatriDB: nuovo, manifCustom: nm, spazioInfo: nsi };
}

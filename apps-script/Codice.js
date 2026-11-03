// ── SPETTACOLI APP SCRIPT v134 ──

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

      var aiJson = _aiFetch(apiKey, { max_tokens: 16000, messages: messages }, body.action === 'ai_pdf' ? 'fatture/pdf' : 'testo');
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
      var G0 = _mappeGiornale(SpreadsheetApp.getActiveSpreadsheet());   // un'app indietro non può riportare nomi vecchi
      props.setProperty('teatriDB', JSON.stringify(_giornaleTeatriDB(body.teatriDB||{}, G0)));
      if(body.manifCustom) props.setProperty('manifCustom', JSON.stringify(_giornaleManif(body.manifCustom, G0)));
      if(body.dittaDB) props.setProperty('dittaDB', JSON.stringify(body.dittaDB));
      if(body.collaboratori) props.setProperty('collaboratori', JSON.stringify(body.collaboratori));
      if(body.spazioInfo) props.setProperty('spazioInfo', JSON.stringify(_giornaleSpazioInfo(body.spazioInfo, G0)));
      if(body.aziendeDB) props.setProperty('aziendeDB', JSON.stringify(body.aziendeDB));
      if(body.pianiProduzione) props.setProperty('pianiProduzione', JSON.stringify(body.pianiProduzione));
      lock.releaseLock();
      return resp('config ok');
    }

    if (body.action === 'sendEmail') {
      var to      = body.to || '';
      var subject = body.subject || '';
      var message = body.message || '';
      if(!to || !subject || !message){ lock.releaseLock(); return resp('missing fields'); }
      // il testo può contenere bottoni HTML (link "Scarica"): nel corpo testuale restano solo le parole
      var opt = { to: to, subject: subject, body: message.replace(/<a [^>]*href="([^"]+)"[^>]*>([^<]*)<\/a>/g, '$2: $1').replace(/<[^>]+>/g, ''), htmlBody: message.split(String.fromCharCode(10)).join('<br>') };
      if (COPIA_MAIL_A && to.toLowerCase().indexOf(COPIA_MAIL_A.toLowerCase()) < 0) opt.bcc = COPIA_MAIL_A;   // copia di ogni mail inviata dall'app
      var att = [];
      (body.allegati || []).forEach(function(u) {
        var m = String(u || '').match(/[-\w]{25,}/);
        if (!m) return;
        try { var f = DriveApp.getFileById(m[0]); if (f.getSize() < 20 * 1024 * 1024) att.push(f.getBlob()); } catch(e) {}
      });
      // documento HTML (es. piano di produzione) da allegare come PDF
      if (body.pdfHtml) { try { att.push(Utilities.newBlob(body.pdfHtml, 'text/html', 'doc.html').getAs('application/pdf').setName(body.pdfNome || 'documento.pdf')); } catch(e) {} }
      if (att.length) opt.attachments = att;
      MailApp.sendEmail(opt);
      lock.releaseLock();
      return resp('sent');
    }
    if (body.action === 'updateFatturato') {
      var rf = aggiornaFatturato(body.id, !!body.fatturato, body.autore);
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify(rf)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'updateNote') {
      var rn = aggiornaNota(body.id, body.note, body.autore);
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify(rn)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'scansionaOra') {
      lock.releaseLock();   // dura minuti: non bloccare le sync
      var rs2; try { rs2 = scansionaOra(); } catch(errS) { rs2 = { error: errS.message }; }
      return ContentService.createTextOutput(JSON.stringify(rs2)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'applicaAnagrafica') {
      var ra; try { ra = applicaAnagrafica(body.remaps); } catch(errA) { ra = { error: errA.message }; }
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify(ra)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'uploadChunk') {
      lock.releaseLock();
      var rc; try { rc = caricaChunk(body.cartella, body.nome, body.mime, body.b64, +body.idx, +body.total, body.fileId); } catch(errC) { rc = { error: errC.message }; }
      return ContentService.createTextOutput(JSON.stringify(rc)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'uploadFile') {
      lock.releaseLock();
      var ru; try { ru = caricaDocumento(body.cartella, body.nome, body.mime, body.b64); } catch(errU) { ru = { error: errU.message }; }
      return ContentService.createTextOutput(JSON.stringify(ru)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'setModelloAI') {
      var mod = String(body.modello || '').trim();
      if (!mod) props.deleteProperty('modelloAI'); else props.setProperty('modelloAI', mod);
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, modello: _modelloAI() })).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'setMittenti') {
      var lista = (body.mittenti || []).map(function(x){ return String(x || '').trim().toLowerCase(); }).filter(function(x){ return /^[^@\s]+@[^@\s]+$/.test(x); });
      if (!lista.length) { lock.releaseLock(); return ContentService.createTextOutput(JSON.stringify({ error: 'nessun indirizzo valido' })).setMimeType(ContentService.MimeType.JSON); }
      props.setProperty('mittentiCalendari', JSON.stringify(lista));
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, mittenti: lista })).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'agganciaAllegati') {
      lock.releaseLock();
      var rg; try { rg = agganciaAllegatiMail(body.link, body.allegati, body.titolo, body.threadId, body.oggetto); } catch(errG) { rg = { error: errG.message }; }
      return ContentService.createTextOutput(JSON.stringify(rg)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'allegatiMail') {
      lock.releaseLock();
      var ra; try { ra = allegatiMailThread(body.link, body.fonte); } catch(errA) { ra = { error: errA.message }; }
      return ContentService.createTextOutput(JSON.stringify(ra)).setMimeType(ContentService.MimeType.JSON);
    }
    if (body.action === 'inoltraMail') {
      lock.releaseLock();
      var rf2; try { rf2 = inoltraMailThread(body.link, body.fonte, body.to, body.testo, body.oggetto, body.allegati); } catch(errF) { rf2 = { error: errF.message }; }
      return ContentService.createTextOutput(JSON.stringify(rf2)).setMimeType(ContentService.MimeType.JSON);
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
      'N.SART','TAR.SART','SUB.SART','TOTALE','NOTE','FATTURATO','PREMONTAGGIO','COLLAB_JSON','REPLICA_OF','MAILS_JSON','CONTATTI_JSON','ACCORDI','PROD_JSON'];
    // Difesa dal "locale vince": le rinomine/correzioni del giornale si applicano anche ai dati in arrivo
    var G1 = _mappeGiornale(ss);
    if (G1.n) body.data.forEach(function(r){ _giornaleRiga(r, G1); });
    sh.clearContents();
    sh.getRange(1,1,1,header.length).setValues([header]);
    var rows = body.data.map(function(r,i) {
      return [r.id||i+1,r.data,r.citta,r.teatro,r.spettacolo,r.manifestazione,
        r.n_macc,r.tariffa_macc,r.sub_macc,
        r.n_elett,r.tariffa_elett,r.sub_elett,
        r.n_facc,r.tariffa_facc,r.sub_facc,
        r.n_foni||0,r.tariffa_foni||0,r.sub_foni||0,
        r.n_sart||0,r.tariffa_sart||0,r.sub_sart||0,
        r.totale,r.note,r.fatturato,r.premontaggio||'NO',r.collab_json||'',r.replica_of||'',r.mails_json||'',r.contatti_json||'',r.accordi||'',r.prod_json||''];
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
  if(action === 'cercaMail') {
    var rq; try { rq = cercaMailScheda(e.parameter.q, +e.parameter.giorni || 365); } catch(err) { rq = { error: err.message }; }
    return ContentService.createTextOutput(JSON.stringify(rq)).setMimeType(ContentService.MimeType.JSON);
  }
  if(action === 'cartellaSpazio') {
    var rc; try { rc = cartellaSpazio(e.parameter.cartella); } catch(err) { rc = { error: err.message }; }
    return ContentService.createTextOutput(JSON.stringify(rc)).setMimeType(ContentService.MimeType.JSON);
  }
  if(action === 'consumoAI') {
    return ContentService.createTextOutput(JSON.stringify(consumoAI())).setMimeType(ContentService.MimeType.JSON);
  }
  if(action === 'mittenti') {
    return ContentService.createTextOutput(JSON.stringify({ mittenti: _mittenti() })).setMimeType(ContentService.MimeType.JSON);
  }
  if(action === 'auditDate') {
    try { return ContentService.createTextOutput(JSON.stringify(auditDateProposte())).setMimeType(ContentService.MimeType.JSON); }
    catch(err) { return ContentService.createTextOutput(JSON.stringify({error: err.message})).setMimeType(ContentService.MimeType.JSON); }
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
          if(v instanceof Date) v = Utilities.formatDate(v,'Europe/Rome','yyyy-MM-dd');
          o[h] = v;
        });
        if(String(o.stato).toLowerCase()==='pending') out.push(o);
      }
      return ContentService.createTextOutput(JSON.stringify({proposte:out, ultimaScansione: props.getProperty('ultimaScansione') || ''})).setMimeType(ContentService.MimeType.JSON);
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
        pianiProduzione: JSON.parse(props.getProperty('pianiProduzione')||'{}'),
        anagraficaVer: parseInt(props.getProperty('anagraficaVer')||'0'),
        anagraficaRemaps: _leggiGiornale(ss, parseInt(e.parameter.anagVer||'0'))
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
  var threads = GmailApp.search(_fromMittenti() + ' has:attachment', 0, 20);
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

  var threads = GmailApp.search(_fromMittenti() + ' has:attachment', 0, 20);

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

  var j = _aiFetch(KEY, {max_tokens:16000, messages:[{role:'user', content: prompt}]}, 'calendari');
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

function scansionaMail(maxMin) {
  var t0 = Date.now();
  var limiteMin = maxMin || MAX_MIN;
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
    if (d instanceof Date) d = Utilities.formatDate(d, 'Europe/Rome', 'yyyy-MM-dd');
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
      if (dp instanceof Date) dp = Utilities.formatDate(dp, 'Europe/Rome', 'yyyy-MM-dd');
      esistenti[chiave(dp, rP[q][pC], rP[q][pT])] = true;
    }
  }
  Logger.log('Chiavi note: ' + Object.keys(esistenti).length);

  var label = GmailApp.getUserLabelByName('GF-vista') || GmailApp.createLabel('GF-vista');
  var threads = GmailApp.search(_fromMittenti() + ' has:attachment after:' + DA_DATA + ' -label:GF-vista', 0, 300);
  Logger.log('Thread da esaminare: ' + threads.length + '\n');

  var scritte = 0, dup = 0, fatti = 0, senzaSpazio = 0, interrotto = false;

  for (var t = 0; t < threads.length; t++) {
    if ((Date.now() - t0) / 60000 > limiteMin) { interrotto = true; break; }

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
  return 'Calendari: ' + fatti + ' Word letti, ' + scritte + ' proposte' + (interrotto ? ' (fermato per tempo, riprende al prossimo giro)' : '');
}
function scansionaSettimanale() {
  var out = [];
  try { scansionaMail(); out.push('Allegati: scansione eseguita'); }
  catch(e){ out.push('Errore allegati: ' + e.message); }
  try { out.push(analizzaCorpoMail()); }
  catch(e){ out.push('Errore corpo mail: ' + e.message); }
  try { out.push(scansionaSchede()); }
  catch(e){ out.push('Errore schede: ' + e.message); }

  PropertiesService.getScriptProperties().setProperty('ultimaScansione', Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy HH:mm'));
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
      d = Utilities.formatDate(d,'Europe/Rome','yyyy-MM-dd'); }
    else { if (!d || String(d) < Utilities.formatDate(oggi,'Europe/Rome','yyyy-MM-dd')) continue; }
    elenco.push(d + ' | ' + rS[r][iC] + ' | ' + rS[r][iT]);
  }

  var lbl = GmailApp.getUserLabelByName('GF-corpo') || GmailApp.createLabel('GF-corpo');
  var threads = GmailApp.search(_fromMittenti() + ' newer_than:60d -label:GF-corpo', 0, 30);
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

  var jr = _aiFetch(KEY, {max_tokens:8000, messages:[{role:'user', content:prompt}]}, 'corpo mail');
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

function scansionaSchede(maxMin) {
  var ctx = _ctxSchede();
  ctx.maxMin = maxMin || SCHEDE_MAX_MIN;
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
// Trova la conversazione Gmail dal link salvato sullo spettacolo (o dall'oggetto)
function _threadDaLink(link, fonte) {
  var th = null;
  var id = String(link || '').split(/[\/#]/).pop();
  if (id) { try { th = GmailApp.getThreadById(id); } catch(e) {} }
  if (!th && fonte) {
    var ogg = String(fonte).split(' | ')[0].replace(/^\s*((I|R|RE|FW|FWD|TR|Fwd|Re|I):\s*)+/i, '').replace(/"/g, '').trim();
    if (ogg.length > 6) { var found = GmailApp.search('subject:"' + ogg.substring(0, 80) + '" newer_than:365d', 0, 3); if (found.length) th = found[0]; }
  }
  return th;
}
// Inoltra ai referenti (teatro, service…) la mail della compagnia con tutti gli allegati, più un testo introduttivo
// il messaggio del thread con allegati (di solito la scheda tecnica), altrimenti l'ultimo
function _msgConAllegati(th) {
  var msgs = th.getMessages(), msg = msgs[msgs.length - 1];
  for (var i = msgs.length - 1; i >= 0; i--) { if (msg.getAttachments().length) break; if (msgs[i].getAttachments().length) { msg = msgs[i]; break; } }
  return msg;
}
// Ricerca mirata in Gmail per uno spettacolo: mail con allegati PDF/Word che citano le parole del titolo
// (oggetto o testo). Serve quando la scansione automatica non ha agganciato la scheda: si sceglie a mano.
function cercaMailScheda(q, giorni) {
  var parole = String(q || '').replace(/[^\wÀ-ÿ' ]+/g, ' ').split(/\s+/).filter(function(w){ return w.length >= 4; }).slice(0, 5);
  if (!parole.length) return { error: 'scrivi almeno una parola di 4 lettere' };
  var query = 'has:attachment newer_than:' + (giorni || 365) + 'd (filename:pdf OR filename:docx OR filename:doc) ' + parole.map(function(w){ return '"' + w + '"'; }).join(' ');
  var threads = GmailApp.search(query, 0, 12);
  var out = threads.map(function(th){
    var msg = _msgConAllegati(th);
    var atts = msg.getAttachments().map(function(a, i){ return { i: i, nome: a.getName(), kb: Math.round(a.getSize() / 1024) }; }).filter(function(a){ return /\.(pdf|docx?)$/i.test(a.nome); });
    return { link: th.getPermalink(), id: th.getId(), oggetto: msg.getSubject(), da: msg.getFrom(), data: Utilities.formatDate(msg.getDate(), 'Europe/Rome', 'dd/MM/yyyy'), allegati: atts, etichettata: th.getLabels().some(function(l){ return l.getName() === 'GF-scheda'; }) };
  }).filter(function(t){ return t.allegati.length; });
  return { ok: true, query: query, mail: out };
}
// Aggancio manuale (senza AI): salva su Drive gli allegati scelti e restituisce i link; l'app li mette sullo spettacolo
function agganciaAllegatiMail(link, allegati, titolo, threadId, oggetto) {
  var th = null;
  if (threadId) { try { th = GmailApp.getThreadById(threadId); } catch(e) {} }
  if (!th) th = _threadDaLink(link, oggetto || '');
  if (!th) return { error: 'mail non trovata' };
  var msg = _msgConAllegati(th), tutti = msg.getAttachments();
  var out = [];
  (allegati || []).forEach(function(i){ var a = tutti[+i]; if (!a) return; try { out.push({ nome: a.getName(), url: _salvaAllegato(a, titolo) }); } catch(e) { out.push({ nome: a.getName(), error: e.message }); } });
  try { th.addLabel(GmailApp.getUserLabelByName('GF-scheda') || GmailApp.createLabel('GF-scheda')); } catch(e) {}
  return { ok: true, file: out, da: msg.getFrom(), oggetto: msg.getSubject(), contatti: _contattiScheda(msg, []) };
}
// Elenco allegati della mail: l'app li mostra con le spunte, così al teatro va solo la scheda di competenza
function allegatiMailThread(link, fonte) {
  var th = _threadDaLink(link, fonte);
  if (!th) return { error: 'mail non trovata (né link né oggetto corrispondente)' };
  var msg = _msgConAllegati(th);
  var atts = msg.getAttachments().map(function(a, i){ return { i: i, nome: a.getName(), kb: Math.round(a.getSize() / 1024), tipo: a.getContentType() }; });
  return { ok: true, oggetto: msg.getSubject(), da: msg.getFrom(), allegati: atts };
}
// allegati: array di indici (da allegatiMailThread). Se manca → inoltro classico con tutto.
function inoltraMailThread(link, fonte, to, testo, oggetto, allegati) {
  if (!to) return { error: 'destinatario mancante' };
  var th = _threadDaLink(link, fonte);
  if (!th) return { error: 'mail non trovata (né link né oggetto corrispondente)' };
  var msg = _msgConAllegati(th);
  var intro = String(testo || '');
  var corpoOrig = msg.getBody();
  var opt = { htmlBody: intro.replace(/\n/g, '<br>') + '<br><br>---------- Messaggio inoltrato ----------<br>' + corpoOrig };
  if (COPIA_MAIL_A && String(to).toLowerCase().indexOf(COPIA_MAIL_A.toLowerCase()) < 0) opt.bcc = COPIA_MAIL_A;
  var tutti = msg.getAttachments();
  if (Array.isArray(allegati)) {
    // solo gli allegati scelti: mail nuova con il corpo originale in coda (niente immagini inline della mail di partenza)
    var scelti = allegati.map(function(i){ return tutti[+i]; }).filter(Boolean);
    opt.attachments = scelti;
    opt.htmlBody = opt.htmlBody.replace(/<img[^>]*src="cid:[^"]*"[^>]*>/gi, '');
    var subj = oggetto || ('Fwd: ' + msg.getSubject());
    GmailApp.sendEmail(to, subj, intro + '\n\n---------- Messaggio inoltrato ----------\n' + msg.getPlainBody(), opt);
    return { ok: true, oggetto: subj, allegati: scelti.length, esclusi: tutti.length - scelti.length, da: msg.getFrom() };
  }
  if (oggetto) opt.subject = oggetto;
  msg.forward(to, opt);
  return { ok: true, oggetto: oggetto || ('Fwd: ' + msg.getSubject()), allegati: tutti.length, da: msg.getFrom() };
}
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
    if (!forza && ((Date.now() - ctx.t0) / 60000 > (ctx.maxMin || SCHEDE_MAX_MIN) || ctx.esaminati >= SCHEDE_MAX_LETTURE)) { ctx.interrotto = true; completo = false; break; }
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

  var jr = _aiFetch(KEY, { max_tokens: 3000, messages: [{ role: 'user', content: content }] }, 'schede');
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


// ════════════════════ MODELLO AI E CONSUMO ════════════════════
// Tutte le chiamate a Claude passano di qui: un solo punto per modello, errori e conteggio token.
// Il modello si cambia dall'app (Config → Consumo AI) e finisce nella proprietà 'modelloAI'.
// Sonnet legge schede, fatture e mail con la stessa qualità di Opus a una frazione del costo.
var MODELLO_AI_DEFAULT = 'claude-sonnet-4-5';
var MODELLI_AI = { sonnet: 'claude-sonnet-4-5', opus: 'claude-opus-4-8', haiku: 'claude-haiku-4-5' };
function _modelloAI() {
  var m = PropertiesService.getScriptProperties().getProperty('modelloAI') || '';
  return MODELLI_AI[m] || m || MODELLO_AI_DEFAULT;
}
function _aiFetch(KEY, payload, uso) {
  var modello = payload.model || _modelloAI();
  var chiama = function(m) {
    payload.model = m;
    var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'x-api-key': KEY, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    try { return JSON.parse(res.getContentText()); }
    catch(e) { return { error: { type: 'http', message: 'risposta non valida dallo script (' + res.getResponseCode() + ')' } }; }
  };
  var j = chiama(modello);
  // modello non più disponibile (nome vecchio): riprovo col predefinito, così l'app non si blocca
  if (j.error && /not_found/i.test(String(j.error.type)) && modello !== MODELLO_AI_DEFAULT) { modello = MODELLO_AI_DEFAULT; j = chiama(modello); }
  _registraConsumoAI(uso || 'altro', modello, j.usage, !!j.error);
  return j;
}
// Registro per giorno: chiamate, token in/out, per uso e per modello. Tenuto per 90 giorni.
function _registraConsumoAI(uso, modello, usage, errore) {
  try {
    var props = PropertiesService.getScriptProperties();
    var reg = JSON.parse(props.getProperty('aiConsumo') || '{}');
    var g = Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyy-MM-dd');
    var r = reg[g] = reg[g] || { chiamate: 0, errori: 0, input: 0, output: 0, usi: {}, modelli: {} };
    var inp = (usage && usage.input_tokens) || 0, out = (usage && usage.output_tokens) || 0;
    r.chiamate++; if (errore) r.errori++; r.input += inp; r.output += out;
    var u = r.usi[uso] = r.usi[uso] || { chiamate: 0, input: 0, output: 0 }; u.chiamate++; u.input += inp; u.output += out;
    var mm = r.modelli[modello] = r.modelli[modello] || { chiamate: 0, input: 0, output: 0 }; mm.chiamate++; mm.input += inp; mm.output += out;
    var giorni = Object.keys(reg).sort(); while (giorni.length > 90) delete reg[giorni.shift()];
    props.setProperty('aiConsumo', JSON.stringify(reg));
  } catch(e) { Logger.log('consumo AI non registrato: ' + e.message); }
}
function consumoAI() {
  var props = PropertiesService.getScriptProperties();
  return { modello: _modelloAI(), modelloScelto: props.getProperty('modelloAI') || 'sonnet', modelli: MODELLI_AI, consumo: JSON.parse(props.getProperty('aiConsumo') || '{}') };
}

// ════════════════════ DOCUMENTI (upload su Drive) ════════════════════
// Salva un file in "GammaFucker Documenti/<cartella>" con link di sola lettura
// per chi ha il link. Sostituisce il copia-incolla dei link Dropbox.
var DOC_CARTELLA = 'GammaFucker Documenti';
function _cartellaDoc(cartella) {
  var it = DriveApp.getFoldersByName(DOC_CARTELLA);
  var root = it.hasNext() ? it.next() : DriveApp.createFolder(DOC_CARTELLA);
  var sub = String(cartella || 'Varie').replace(/[\\/:*?"<>|]/g, ' ').trim().substring(0, 80) || 'Varie';
  var it2 = root.getFoldersByName(sub);
  return { folder: it2.hasNext() ? it2.next() : root.createFolder(sub), sub: sub };
}
// La cartella dello spazio è condivisa "chi ha il link può vedere": su WhatsApp basta un link solo
// invece dell'elenco dei file, e chi lo apre vede le anteprime di tutti i documenti.
function _cartellaCondivisa(folder) {
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return folder.getUrl();
}
function cartellaSpazio(cartella) {
  var it = DriveApp.getFoldersByName(DOC_CARTELLA);
  if (!it.hasNext()) return { url: '' };
  var sub = String(cartella || '').replace(/[\\/:*?"<>|]/g, ' ').trim().substring(0, 80);
  var it2 = it.next().getFoldersByName(sub);
  if (!it2.hasNext()) return { url: '' };
  var folder = it2.next();
  if (!folder.getFiles().hasNext()) return { url: '' };
  return { url: _cartellaCondivisa(folder) };
}
function caricaDocumento(cartella, nome, mime, b64) {
  if (!nome || !b64) return { error: 'file mancante' };
  var c = _cartellaDoc(cartella);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', nome);
  var f = c.folder.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
  return { url: f.getUrl(), id: f.getId(), nome: f.getName(), cartella: c.sub, cartellaUrl: _cartellaCondivisa(c.folder) };
}
// Upload a pezzi (file grandi: planimetrie PDF ecc.): il primo pezzo crea il file, i successivi lo allungano.
function caricaChunk(cartella, nome, mime, b64, idx, total, fileId) {
  if (!nome || !b64) return { error: 'pezzo mancante' };
  var bytes = Utilities.base64Decode(b64);
  var f, sub = '';
  if (idx === 0 || !fileId) {
    var c = _cartellaDoc(cartella); sub = c.sub;
    f = c.folder.createFile(Utilities.newBlob(bytes, mime || 'application/octet-stream', nome));
  } else {
    f = DriveApp.getFileById(fileId);
    var tutti = f.getBlob().getBytes().concat(bytes);
    Drive.Files.update({}, fileId, Utilities.newBlob(tutti, mime || 'application/octet-stream', nome));
    f = DriveApp.getFileById(fileId);
  }
  var done = idx >= total - 1;
  if (done) { try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} }
  return { url: f.getUrl(), id: f.getId(), nome: f.getName(), cartella: sub, done: done, size: f.getSize(), cartellaUrl: done ? _cartellaCondivisa(f.getParents().next()) : '' };
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
  var mapM = {}, mapC = {}, mapT = {}, aggiunte = [], mapD = [], mapF = {};
  remaps.forEach(function(x) {
    if (x.tipo === 'manif' && x.da && x.a && x.da !== x.a) mapM[String(x.da).trim()] = String(x.a).trim();
    else if (x.tipo === 'citta' && x.da && x.a && U(x.da) !== U(x.a)) mapC[U(x.da)] = U(x.a);
    else if (x.tipo === 'teatro' && x.citta && x.da && x.a && U(x.da) !== U(x.a)) mapT[U(x.citta) + '|' + U(x.da)] = U(x.a);
    else if (x.tipo === 'aggiungiTeatro' && x.citta) aggiunte.push({ citta: U(x.citta), teatro: U(x.teatro) });
    else if (x.tipo === 'data' && x.da && x.a) mapD.push({ id: String(x.id || ''), citta: U(x.citta), spett: U(x.spett), da: String(x.da), a: String(x.a) });
    else if (x.tipo === 'campi' && x.id) mapF[_normId(x.id)] = x;   // {id, note_add, accordi, contatti:[...]}
  });
  function fixCitta(c){ var u = U(c); return mapC[u] || u; }
  function fixTeatro(c, t){ var u = U(t); var m = mapT[U(c) + '|' + u]; return m || u; }   // usa la città ORIGINALE come chiave

  var nRighe = 0;
  ['Spettacoli', 'Proposte'].forEach(function(nome) {
    var sh = ss.getSheetByName(nome); if (!sh) return;
    var rng = sh.getDataRange(), rows = rng.getValues();
    var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
    var iC = hdr.indexOf('CITTA'), iT = hdr.indexOf(nome === 'Spettacoli' ? 'TEATRO' : 'SPAZIO'), iM = hdr.indexOf('MANIFESTAZIONE');
    var iD = hdr.indexOf('DATA'), iId = hdr.indexOf('ID'), iS = hdr.indexOf('SPETTACOLO');
    var iN = hdr.indexOf('NOTE'), iA = hdr.indexOf('ACCORDI'), iK = hdr.indexOf('CONTATTI_JSON'), iFa = hdr.indexOf('FATTURATO');
    if (nome === 'Spettacoli' && Object.keys(mapF).length) {
      if (iA < 0) { sh.getRange(1, hdr.length + 1).setValue('ACCORDI'); hdr.push('ACCORDI'); iA = hdr.length - 1; rows.forEach(function(rw){ rw.push(''); }); rng = sh.getRange(1, 1, rows.length, hdr.length); }
      if (iK < 0) { sh.getRange(1, hdr.length + 1).setValue('CONTATTI_JSON'); hdr.push('CONTATTI_JSON'); iK = hdr.length - 1; rows.forEach(function(rw){ rw.push(''); }); rng = sh.getRange(1, 1, rows.length, hdr.length); }
    }
    var mod = false;
    for (var r = 1; r < rows.length; r++) {
      var ch = false;
      if (nome === 'Spettacoli' && iId >= 0 && mapF[_normId(rows[r][iId])]) {
        var f2 = mapF[_normId(rows[r][iId])];
        if (f2.note_add && iN >= 0) { var nv = String(rows[r][iN] || ''); if (nv.indexOf(f2.note_add) < 0) { rows[r][iN] = nv ? nv + ' | ' + f2.note_add : f2.note_add; ch = true; } }
        if (typeof f2.note_set === 'string' && iN >= 0) { rows[r][iN] = f2.note_set; ch = true; }
        if (typeof f2.fatturato_set === 'boolean' && iFa >= 0) { rows[r][iFa] = f2.fatturato_set ? 'SI' : 'NO'; ch = true; }
        if (f2.accordi && iA >= 0 && !String(rows[r][iA] || '').trim()) { rows[r][iA] = f2.accordi; ch = true; }
        if (f2.contatti && f2.contatti.length && iK >= 0) {
          var cur = []; try { cur = JSON.parse(rows[r][iK] || '[]'); } catch(e) { cur = []; }
          f2.contatti.forEach(function(c){ if (!c) return; var em = String(c.email || '').toLowerCase(); var dup = cur.some(function(o){ return (em && String(o.email||'').toLowerCase() === em) || (!em && String(o.nome||'').toUpperCase() === String(c.nome||'').toUpperCase() && String(o.telefono||'') === String(c.telefono||'')); }); if (!dup) cur.push(c); });
          rows[r][iK] = JSON.stringify(cur); ch = true;
        }
      }
      if (nome === 'Spettacoli' && mapD.length && iD >= 0) {
        var dv = rows[r][iD]; var ds = dv instanceof Date ? Utilities.formatDate(dv, 'Europe/Rome', 'yyyy-MM-dd') : String(dv || '');
        for (var q = 0; q < mapD.length; q++) {
          var f = mapD[q];
          var byId = f.id && _normId(rows[r][iId]) === _normId(f.id);
          var byKey = ds === f.da && U(rows[r][iC]) === f.citta && U(rows[r][iS]) === f.spett;
          if ((byId || byKey) && ds === f.da) { rows[r][iD] = f.a; ch = true; break; }
        }
      }
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
    if (nsi[nk]) { // fondi: tieni i campi compilati
      var a = nsi[nk], b = si[k]; Object.keys(b).forEach(function(f){ if (!a[f] || (Array.isArray(a[f]) && !a[f].length)) a[f] = b[f]; });
    } else nsi[nk] = si[k];
  });

  var ver = parseInt(props.getProperty('anagraficaVer') || '0') + 1;
  _scriviGiornale(ss, ver, remaps);
  props.setProperty('teatriDB', JSON.stringify(nuovo));
  props.setProperty('manifCustom', JSON.stringify(nm));
  props.setProperty('spazioInfo', JSON.stringify(nsi));
  props.deleteProperty('anagraficaRemaps');
  props.setProperty('anagraficaVer', String(ver));
  props.setProperty('lastModified', String(Date.now()));
  return { ok: true, righe: nRighe, ver: ver, teatriDB: nuovo, manifCustom: nm, spazioInfo: nsi };
}

// ════════════════════ SCANSIONE SU RICHIESTA (pulsante Aggiorna nell'app) ════════════════════
// Esegue i tre lettori in sequenza con un budget totale di ~4 minuti (limite web app 6 min).
// Sono tutti incrementali (etichette GF-*): riparte da dove era arrivata l'ultima scansione.
function scansionaOra() {
  var t0 = Date.now(), esiti = [];
  var budget = function(){ return Math.max(0.3, 4.2 - (Date.now() - t0) / 60000); };
  try { esiti.push(scansionaMail(Math.min(1.5, budget()))); } catch(e) { esiti.push('Calendari: errore ' + e.message); }
  try { esiti.push(analizzaCorpoMail()); } catch(e) { esiti.push('Corpo mail: errore ' + e.message); }
  try { esiti.push(scansionaSchede(budget())); } catch(e) { esiti.push('Schede: errore ' + e.message); }
  var ts = Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy HH:mm');
  PropertiesService.getScriptProperties().setProperty('ultimaScansione', ts);
  var out = { ok: true, ultimaScansione: ts, esiti: esiti, secondi: Math.round((Date.now() - t0) / 1000) };
  Logger.log(JSON.stringify(out));
  return out;
}

// ════════════════════ CORREZIONE DATE (bug GMT+1 in ora legale) ════════════════════
// Le proposte approvate tra fine agosto e settembre 2026 avevano la data spostata di -1 giorno
// (formatDate con 'GMT+1' fisso durante l'ora legale). Elenco ricavato dal confronto con
// l'Elenco Ordini Eventi del 16/09/2026. Da eseguire UNA volta: correggiDateDST()
// Ogni mail inviata dall'app (info spazio, richieste ZTL, mail alle compagnie, riepiloghi) va in copia nascosta qui.
var COPIA_MAIL_A = 'aldoamat@gmail.com';

// Colleghi che inviano programmazione e calendari (modificabili dall'app: Config → Mittenti programmazione)
var MITTENTI_DEFAULT = ['r.cantarini@amat.marche.it', 'm.osimani@amat.marche.it', 'd.sepe@amat.marche.it'];
function _mittenti() {
  try { var m = JSON.parse(PropertiesService.getScriptProperties().getProperty('mittentiCalendari') || 'null'); if (m && m.length) return m; } catch(e) {}
  return MITTENTI_DEFAULT;
}
function _fromMittenti() { return 'from:(' + _mittenti().join(' OR ') + ')'; }

var DATE_DA_CORREGGERE = [
 {
  "tipo": "data",
  "id": "1787055513996.0",
  "citta": "SENIGALLIA",
  "spett": "ASSENZA. EURIDICE E ORFEO",
  "da": "2026-09-04",
  "a": "2026-09-05"
 },
 {
  "tipo": "data",
  "id": "1787055514996.0",
  "citta": "SENIGALLIA",
  "spett": "CAPELLI",
  "da": "2026-09-05",
  "a": "2026-09-06"
 },
 {
  "tipo": "data",
  "id": "1787055531487.0",
  "citta": "SIROLO",
  "spett": "METAMORFOSI",
  "da": "2026-09-10",
  "a": "2026-09-11"
 },
 {
  "tipo": "data",
  "id": "1787056546421.0",
  "citta": "PESARO",
  "spett": "ROSENCRANTZ E GUILDENSTERN SONO MORTI",
  "da": "2026-10-07",
  "a": "2026-10-08"
 },
 {
  "tipo": "data",
  "id": "1787054242139.0",
  "citta": "ASCOLI PICENO",
  "spett": "LA MISTERIOSA SCOMPARSA DI W",
  "da": "2026-10-09",
  "a": "2026-10-10"
 },
 {
  "tipo": "data",
  "id": "1788514288430.0",
  "citta": "PESARO",
  "spett": "BAMBU 2026",
  "da": "2026-10-10",
  "a": "2026-10-11"
 },
 {
  "tipo": "data",
  "id": "1787056570370.0",
  "citta": "FANO",
  "spett": "UNO NESSUNO CENTOMILA",
  "da": "2026-10-14",
  "a": "2026-10-15"
 },
 {
  "tipo": "data",
  "id": "1787052516955.0",
  "citta": "ASCOLI PICENO",
  "spett": "GLENN - UNA VARIAZIONE",
  "da": "2026-10-15",
  "a": "2026-10-16"
 },
 {
  "tipo": "data",
  "id": "1787052517955.0",
  "citta": "ASCOLI PICENO",
  "spett": "CANI LUNARI",
  "da": "2026-10-15",
  "a": "2026-10-16"
 },
 {
  "tipo": "data",
  "id": "1787052518955.0",
  "citta": "ASCOLI PICENO",
  "spett": "IT'S NOT FAIR TO DIE",
  "da": "2026-10-15",
  "a": "2026-10-16"
 },
 {
  "tipo": "data",
  "id": "1787052535955.0",
  "citta": "MATELICA",
  "spett": "ME LA CANTO E ME LA SOGNO... ANCORA",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787052519955.0",
  "citta": "ASCOLI PICENO",
  "spett": "TOTEMICA",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787052520955.0",
  "citta": "ASCOLI PICENO",
  "spett": "DODESKADEN",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787052521955.0",
  "citta": "ASCOLI PICENO",
  "spett": "L'ANIMALE",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787052523955.0",
  "citta": "ASCOLI PICENO",
  "spett": "L'ANGELO DEL FOCOLARE",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787052522955.0",
  "citta": "ASCOLI PICENO",
  "spett": "FOLKLORE DYNAMICS",
  "da": "2026-10-16",
  "a": "2026-10-17"
 },
 {
  "tipo": "data",
  "id": "1787051026280.0",
  "citta": "TOLENTINO",
  "spett": "PEOPLE, PLACES & THINGS",
  "da": "2026-10-17",
  "a": "2026-10-18"
 },
 {
  "tipo": "data",
  "id": "1787059222246.0",
  "citta": "SAN SEVERINO MARCHE",
  "spett": "LA MISTERIOSA SCOMPARSA DI W",
  "da": "2026-10-22",
  "a": "2026-10-23"
 },
 {
  "tipo": "data",
  "id": "1787054257139.0",
  "citta": "OSIMO",
  "spett": "LA MISTERIOSA SCOMPARSA DI W",
  "da": "2026-10-23",
  "a": "2026-10-24"
 },
 {
  "tipo": "data",
  "id": "1787056576370.0",
  "citta": "FANO",
  "spett": "GRACES",
  "da": "2027-03-31",
  "a": "2027-04-01"
 },
 {
  "tipo": "data",
  "id": "1787052560955.0",
  "citta": "PORTO SANT'ELPIDIO",
  "spett": "TUTTA COLPA DI ADAMO",
  "da": "2027-04-01",
  "a": "2027-04-02"
 },
 {
  "tipo": "data",
  "id": "1787059228246.0",
  "citta": "SAN SEVERINO MARCHE",
  "spett": "AMADEUS",
  "da": "2027-04-05",
  "a": "2027-04-06"
 },
 {
  "tipo": "data",
  "id": "1788513815691.0",
  "citta": "RECANATI",
  "spett": "CUORE",
  "da": "2027-04-05",
  "a": "2027-04-06"
 },
 {
  "tipo": "data",
  "id": "1787051033280.0",
  "citta": "TOLENTINO",
  "spett": "RUMORI FUORI SCENA",
  "da": "2027-04-07",
  "a": "2027-04-08"
 },
 {
  "tipo": "data",
  "id": "1787054256139.0",
  "citta": "CORINALDO",
  "spett": "PER FUTILI MOTIVI",
  "da": "2027-04-09",
  "a": "2027-04-10"
 },
 {
  "tipo": "data",
  "id": "1787052533955.0",
  "citta": "CIVITANOVA MARCHE",
  "spett": "CARI GENITORI",
  "da": "2027-04-15",
  "a": "2027-04-16"
 },
 {
  "tipo": "data",
  "id": "1789452014290.0",
  "citta": "CHIARAVALLE",
  "spett": "IL PIACERE DELL'ATTESA",
  "da": "2027-04-16",
  "a": "2027-04-17"
 },
 {
  "tipo": "data",
  "id": "1788513804691.0",
  "citta": "GROTTAMMARE",
  "spett": "IL PADRONE",
  "da": "2027-04-20",
  "a": "2027-04-21"
 },
 {
  "tipo": "data",
  "id": "1787052586955.0",
  "citta": "SENIGALLIA",
  "spett": "CONTRAZIONI PERICOLOSE",
  "da": "2027-04-22",
  "a": "2027-04-23"
 },
 {
  "tipo": "data",
  "id": "1789452021290.0",
  "citta": "JESI",
  "spett": "POVERI CRISTI",
  "da": "2027-04-22",
  "a": "2027-04-23"
 },
 {
  "tipo": "data",
  "id": "1787052570955.0",
  "citta": "SAN BENEDETTO",
  "spett": "CENA CON SORPRESA",
  "da": "2027-04-26",
  "a": "2027-04-27"
 },
 {
  "tipo": "data",
  "id": "1787054251139.0",
  "citta": "ASCOLI PICENO",
  "spett": "SMANIE PER LA VILLEGGIATURA",
  "da": "2027-04-28",
  "a": "2027-04-29"
 },
 {
  "tipo": "data",
  "id": "1787052534955.0",
  "citta": "CIVITANOVA MARCHE",
  "spett": "LE NOSTRE DONNE",
  "da": "2027-05-04",
  "a": "2027-05-05"
 }
];
// ════════════════════ VERIFICA DATE dalle proposte ════════════════════
// Le proposte approvate conservano la data letta dal programma; gli spettacoli creati da esse
// durante l'ora legale (bug GMT+1) sono un giorno indietro. Confronto: per ogni proposta approvata
// (città + titolo + data) senza spettacolo in quella data ma con uno il giorno prima → proposta di
// correzione +1 (con la catena di repliche/premontaggi nei giorni precedenti). GET ?action=auditDate
function auditDateProposte() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shP = ss.getSheetByName('Proposte'), shS = ss.getSheetByName('Spettacoli');
  if (!shP || !shS) return { error: 'fogli mancanti' };
  var fmt = function(v){ return v instanceof Date ? Utilities.formatDate(v, 'Europe/Rome', 'yyyy-MM-dd') : String(v || '').substring(0, 10); };
  var U = function(v){ return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var key = function(c, t){ return U(c) + '|' + _normTitoloGS(t); };
  var addDays = function(d, n){ var x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().substring(0, 10); };

  var P = shP.getDataRange().getValues(), hP = P[0].map(function(h){ return U(h); });
  var pS = hP.indexOf('STATO'), pC = hP.indexOf('CITTA'), pD = hP.indexOf('DATA'), pT = hP.indexOf('TITOLO'), pR = hP.indexOf('REPLICHE');
  var prop = {};   // key → { date: {d:true}, repliche }
  for (var i = 1; i < P.length; i++) {
    if (String(P[i][pS]).toLowerCase() !== 'approvata') continue;
    var d = fmt(P[i][pD]); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
    var k = key(P[i][pC], P[i][pT]); if (!prop[k]) prop[k] = { date: {}, repliche: 1 };
    prop[k].date[d] = true; prop[k].repliche = Math.max(prop[k].repliche, +P[i][pR] || 1);
  }
  var S = shS.getDataRange().getValues(), hS = S[0].map(function(h){ return U(h); });
  var sI = hS.indexOf('ID'), sD = hS.indexOf('DATA'), sC = hS.indexOf('CITTA'), sT = hS.indexOf('SPETTACOLO'), sM = hS.indexOf('MANIFESTAZIONE'), sTe = hS.indexOf('TEATRO');
  var rows = {};   // key → [{id,data,citta,spett,teatro,manif}]
  for (var j = 1; j < S.length; j++) {
    var k2 = key(S[j][sC], S[j][sT]); if (!prop[k2]) continue;
    (rows[k2] = rows[k2] || []).push({ id: _normId(S[j][sI]), data: fmt(S[j][sD]), citta: String(S[j][sC] || ''), spett: String(S[j][sT] || ''), teatro: String(S[j][sTe] || ''), manif: String(S[j][sM] || '') });
  }
  var out = [], visti = {};
  Object.keys(prop).forEach(function(k) {
    var R = rows[k] || []; if (!R.length) return;
    var byDate = {}; R.forEach(function(r){ (byDate[r.data] = byDate[r.data] || []).push(r); });
    Object.keys(prop[k].date).forEach(function(d) {
      if (byDate[d]) return;                       // c'è già uno spettacolo in quella data → ok
      var prev = addDays(d, -1); if (!byDate[prev]) return;
      // catena all'indietro: repliche/premontaggi nei giorni consecutivi precedenti (non oltre un'altra data di proposta)
      var x = prev, catena = [];
      while (byDate[x] && !prop[k].date[x] && catena.length < 8) { catena = catena.concat(byDate[x]); x = addDays(x, -1); }
      catena.forEach(function(r) {
        if (visti[r.id]) return; visti[r.id] = true;
        out.push({ tipo: 'data', id: r.id, citta: r.citta, spett: r.spett, teatro: r.teatro, manif: r.manif, da: r.data, a: addDays(r.data, 1), proposta: d });
      });
    });
  });
  out.sort(function(a, b){ return a.da < b.da ? -1 : a.da > b.da ? 1 : 0; });
  return { totProposte: Object.keys(prop).length, totSpettacoli: S.length - 1, correzioni: out };
}
function anteprimaDateProposte() { var r = auditDateProposte(); Logger.log(r.correzioni.length + ' correzioni proposte'); r.correzioni.forEach(function(x){ Logger.log(x.da + ' → ' + x.a + '  ' + x.citta + ' · ' + x.spett); }); return r; }
function correggiDateProposte() { var r = auditDateProposte(); if (!r.correzioni.length) { Logger.log('niente da correggere'); return r; } var a = applicaAnagrafica(r.correzioni.map(function(x){ return { tipo: 'data', id: x.id, citta: x.citta, spett: x.spett, da: x.da, a: x.a }; })); Logger.log('Date corrette: ' + a.righe + ' (ver ' + a.ver + ')'); return a; }

function correggiDateDST() {
  var r = applicaAnagrafica(DATE_DA_CORREGGERE);
  Logger.log('Date corrette nel foglio: ' + r.righe + ' (versione anagrafica ' + r.ver + '); i dispositivi si allineano al prossimo pull');
  return r;
}

// ════════════════════ GIORNALE (foglio "Giornale") ════════════════════
// Una riga per rinomina/correzione: VER, TS, JSON. Il pull restituisce solo le righe
// con VER > versione del dispositivo, che le applica ai propri dati locali.
function _giornale(ss) {
  var sh = ss.getSheetByName('Giornale');
  if (!sh) { sh = ss.insertSheet('Giornale'); sh.getRange(1, 1, 1, 3).setValues([['VER', 'TS', 'JSON']]); }
  return sh;
}
function _scriviGiornale(ss, ver, remaps) {
  var sh = _giornale(ss), now = new Date();
  var righe = remaps.map(function(x){ x.ver = ver; return [ver, now, JSON.stringify(x)]; });
  if (righe.length) sh.getRange(sh.getLastRow() + 1, 1, righe.length, 3).setValues(righe);
}
function _leggiGiornale(ss, daVer) {
  var sh = ss.getSheetByName('Giornale'); if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues(), out = [];
  for (var i = 0; i < rows.length; i++) {
    if (+rows[i][0] > daVer) { try { out.push(JSON.parse(rows[i][2])); } catch(e) {} }
    if (out.length >= 2000) break;
  }
  return out;
}

// ID spettacolo come stringa confrontabile (dal foglio arrivano numeri, dagli export "….0")
function _normId(v) { var t = String(v == null ? '' : v).trim(); if (/^\d+\.0+$/.test(t)) t = t.replace(/\.0+$/, ''); return t; }

// ── Applicazione del giornale ai dati in arrivo dai dispositivi ──
// Copre il caso di un'app vecchia (o non ancora allineata) che risincronizza nomi/date superati.
function _mappeGiornale(ss) {
  var G = { mapM: {}, mapC: {}, mapT: {}, mapD: {}, n: 0 };
  var U = function(v){ return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  _leggiGiornale(ss, 0).forEach(function(x) {
    if (x.tipo === 'manif' && x.da && x.a) { G.mapM[String(x.da).trim()] = String(x.a).trim(); G.n++; }
    else if (x.tipo === 'citta' && x.da && x.a) { G.mapC[U(x.da)] = U(x.a); G.n++; }
    else if (x.tipo === 'teatro' && x.citta && x.da && x.a) { G.mapT[U(x.citta) + '|' + U(x.da)] = U(x.a); G.n++; }
    else if (x.tipo === 'data' && x.id && x.da && x.a) { G.mapD[_normId(x.id)] = { da: String(x.da), a: String(x.a) }; G.n++; }
  });
  G.U = U;
  G.fixC = function(c){ var u = U(c); var seen = 0; while (G.mapC[u] && seen++ < 5) u = G.mapC[u]; return u; };
  G.fixT = function(c, t){ var u = U(t); var m = G.mapT[U(c) + '|' + u]; return m || u; };
  G.fixM = function(m){ var v = String(m || '').trim(); var seen = 0; while (G.mapM[v] && seen++ < 5) v = G.mapM[v]; return v; };
  return G;
}
function _giornaleRiga(r, G) {
  if (!r) return;
  if (r.manifestazione) r.manifestazione = G.fixM(r.manifestazione);
  if (r.citta) { var oc = r.citta; r.teatro = G.fixT(oc, r.teatro || ''); r.citta = G.fixC(oc); }
  var d = G.mapD[_normId(r.id)]; if (d && String(r.data) === d.da) r.data = d.a;
}
function _giornaleTeatriDB(tdb, G) {
  if (!G.n) return tdb;
  var out = {};
  Object.keys(tdb || {}).forEach(function(c) {
    var nc = G.fixC(c); out[nc] = out[nc] || [];
    (tdb[c] || []).forEach(function(t){ var nt = G.fixT(c, t); if (nt && out[nc].indexOf(nt) < 0) out[nc].push(nt); });
  });
  Object.keys(out).forEach(function(c){ out[c].sort(); });
  return out;
}
function _giornaleManif(lista, G) {
  if (!G.n) return lista;
  var out = []; (lista || []).forEach(function(m){ var v = G.fixM(m); if (v && out.indexOf(v) < 0) out.push(v); });
  return out.sort();
}
function _giornaleSpazioInfo(si, G) {
  if (!G.n) return si;
  var out = {};
  Object.keys(si || {}).forEach(function(k) {
    var p = k.split('||'); var nk = G.fixC(p[0]) + '||' + G.fixT(p[0], p[1] || '');
    if (out[nk]) { var a = out[nk], b = si[k]; Object.keys(b).forEach(function(f){ if (!a[f] || (Array.isArray(a[f]) && !a[f].length)) a[f] = b[f]; }); }
    else out[nk] = si[k];
  });
  return out;
}

// ════════════════════ FATTURATO DALL'APP COLLABORATORI ════════════════════
function aggiornaFatturato(id, fatturato, autore) {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Spettacoli'); if (!sh) return { error: 'foglio mancante' };
  var rows = sh.getDataRange().getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iId = hdr.indexOf('ID'), iF = hdr.indexOf('FATTURATO'), iS = hdr.indexOf('SPETTACOLO');
  if (iF < 0) return { error: 'colonna FATTURATO mancante' };
  var target = _normId(id);
  for (var r = 1; r < rows.length; r++) {
    if (_normId(rows[r][iId]) === target) {
      sh.getRange(r + 1, iF + 1).setValue(fatturato ? 'SI' : 'NO');
      var ver = parseInt(props.getProperty('anagraficaVer') || '0') + 1;
      _scriviGiornale(ss, ver, [{ tipo: 'campi', id: target, fatturato_set: !!fatturato, autore: String(autore || '') }]);
      props.setProperty('anagraficaVer', String(ver));
      props.setProperty('lastModified', String(Date.now()));
      props.setProperty('lastAuthor', String(autore || 'Collaboratore'));
      return { ok: true, spettacolo: rows[r][iS], fatturato: !!fatturato, ver: ver };
    }
  }
  return { error: 'spettacolo non trovato: ' + id };
}

// ════════════════════ NOTE DALL'APP COLLABORATORI ════════════════════
// Aggiorna solo la cella NOTE dello spettacolo e registra la modifica nel giornale,
// così anche i dispositivi admin (merge "locale vince") la ricevono.
function aggiornaNota(id, note, autore) {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Spettacoli'); if (!sh) return { error: 'foglio mancante' };
  var rows = sh.getDataRange().getValues();
  var hdr = rows[0].map(function(h){ return String(h).trim().toUpperCase(); });
  var iId = hdr.indexOf('ID'), iN = hdr.indexOf('NOTE'), iS = hdr.indexOf('SPETTACOLO');
  var target = _normId(id);
  for (var r = 1; r < rows.length; r++) {
    if (_normId(rows[r][iId]) === target) {
      sh.getRange(r + 1, iN + 1).setValue(String(note || ''));
      var ver = parseInt(props.getProperty('anagraficaVer') || '0') + 1;
      _scriviGiornale(ss, ver, [{ tipo: 'campi', id: target, note_set: String(note || ''), autore: String(autore || '') }]);
      props.setProperty('anagraficaVer', String(ver));
      props.setProperty('lastModified', String(Date.now()));
      props.setProperty('lastAuthor', String(autore || 'Collaboratore'));
      return { ok: true, spettacolo: rows[r][iS], ver: ver };
    }
  }
  return { error: 'spettacolo non trovato: ' + id };
}

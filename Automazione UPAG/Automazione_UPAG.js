function salvaParolaDiOggi() {
  var idDocumento = "INSERISCI_QUI_IL_TUO_ID_DOCUMENTO"; 
  
  var threads = GmailApp.search('from:unaparolaalgiorno.it newer_than:1d');
  if (threads.length === 0) {
    Logger.log("Nessuna mail da UPAG trovata oggi.");
    return;
  }
  
  var ultimoMessaggio = threads[0].getMessages().pop();
  var corpoMail = ultimoMessaggio.getBody();
  
  var regexLink = /href=["'](https:\/\/unaparolaalgiorno\.it\/significato\/[^"']+)["']/i;
  var linkTrovato = corpoMail.match(regexLink);
  
  if (!linkTrovato) {
    Logger.log("Nessun link alla parola trovato nella mail.");
    return;
  }
  
  var urlSito = linkTrovato[1].replace(/&amp;/g, '&');
  var htmlSito = UrlFetchApp.fetch(urlSito).getContentText();
  
  var parolaEstratta = urlSito.match(/\/significato\/([^\?]+)/);
  var parola = parolaEstratta ? parolaEstratta[1].toUpperCase() : "PAROLA DEL GIORNO";
  
  var inizio = htmlSito.indexOf("Significato");
  var fine = htmlSito.indexOf("Parola pubblicata il");
  
  if (inizio === -1 || fine === -1 || fine <= inizio) {
    Logger.log("Testo non trovato. Controlla il sito.");
    return;
  }
  
  var bloccoHtml = htmlSito.substring(inizio, fine);
  
  // Pulizia testo
  var testoPulito = bloccoHtml.replace(/<(div|p|h[1-6]|li|br)[^>]*>/gi, '\n');
  testoPulito = testoPulito.replace(/<\/(div|p|h[1-6]|li)>/gi, '\n');
  testoPulito = testoPulito.replace(/<[^>]+>/g, '');
  testoPulito = testoPulito.replace(/&nbsp;/g, ' ').replace(/&[a-zA-Z0-9]+;/g, '');
  testoPulito = testoPulito.replace(/[ \t]+/g, ' ');
  testoPulito = testoPulito.replace(/\n\s*\n/g, '\n').trim();
  
  var documento = DocumentApp.openById(idDocumento);
  var foglio = documento.getBody();
  
  if (foglio.getText().trim().length > 0) {
    foglio.appendPageBreak();
  }
  
  // TITOLO PRINCIPALE
  var titolo = foglio.appendParagraph(parola);
  titolo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titolo.setForegroundColor("#1a237e"); 
  titolo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titolo.setSpacingAfter(6);
  
  // DATA
  var dataOggi = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var dataParagrafo = foglio.appendParagraph("Aggiunta il: " + dataOggi);
  dataParagrafo.setItalic(true).setForegroundColor("#757575");
  dataParagrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  dataParagrafo.setSpacingAfter(24);
  
  // ELABORAZIONE CORPO DEL TESTO
  var righe = testoPulito.split('\n');
  var sezioneAttuale = "descrizione"; 
  
  for (var i = 0; i < righe.length; i++) {
    var testo = righe[i].trim();
    if (testo === "") continue;
    
    var testoLower = testo.toLowerCase();
    
    // 1. Riconoscimento Sezione SIGNIFICATO (Cerca la parola all'inizio della riga)
    if (testoLower.match(/^significato\b/)) {
      sezioneAttuale = "significato";
      var pTitolo = foglio.appendParagraph("Significato");
      pTitolo.setBold(true);
      pTitolo.setFontSize(12);
      pTitolo.setBackgroundColor("#e8f5e9");
      pTitolo.setForegroundColor("#1b5e20");
      pTitolo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pTitolo.setSpacingBefore(18);
      pTitolo.setSpacingAfter(0);
      
      // Ritaglia il testo rimasto sulla stessa riga (es. se c'è "Significato: Costanza di...")
      var resto = testo.replace(/^significato\b[:\s-]*/i, '').trim();
      if (resto !== "") {
        var p = foglio.appendParagraph(resto);
        p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        p.setFontSize(11);
        p.setBackgroundColor("#e8f5e9");
        p.setForegroundColor("#1b5e20");
        p.setSpacingBefore(0);
        p.setSpacingAfter(12);
      }
      continue;
    }
    
    // 2. Riconoscimento Sezione ETIMOLOGIA
    if (testoLower.match(/^etimologia\b/)) {
      sezioneAttuale = "etimologia";
      var pTitolo = foglio.appendParagraph("Etimologia");
      pTitolo.setBold(true);
      pTitolo.setFontSize(12);
      pTitolo.setBackgroundColor("#fff3e0");
      pTitolo.setForegroundColor("#e65100");
      pTitolo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pTitolo.setSpacingBefore(18);
      pTitolo.setSpacingAfter(0);
      
      var resto = testo.replace(/^etimologia\b[:\s-]*/i, '').trim();
      if (resto !== "") {
        var p = foglio.appendParagraph(resto);
        p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        p.setFontSize(11);
        p.setBackgroundColor("#fff3e0");
        p.setForegroundColor("#e65100");
        p.setSpacingBefore(0);
        p.setSpacingAfter(12);
      }
      continue;
    }
    
    // 3. Riconoscimento Sezione ESEMPIO
    if (testo.startsWith("«") || testo.startsWith("\"") || testo.startsWith("“")) {
      if (sezioneAttuale !== "esempio") {
        var pTitolo = foglio.appendParagraph("Esempio");
        pTitolo.setBold(true);
        pTitolo.setFontSize(12);
        pTitolo.setBackgroundColor("#e3f2fd");
        pTitolo.setForegroundColor("#0d47a1");
        pTitolo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        pTitolo.setSpacingBefore(18);
        pTitolo.setSpacingAfter(0);
      }
      sezioneAttuale = "esempio";
      
      var p = foglio.appendParagraph(testo);
      p.setItalic(true); 
      p.setFontSize(11);
      p.setBackgroundColor("#e3f2fd");
      p.setForegroundColor("#0d47a1");
      p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
      p.setSpacingBefore(0);
      p.setSpacingAfter(12); 
      continue;
      
    } else if (sezioneAttuale === "esempio") {
      sezioneAttuale = "descrizione";
    }
    
    // 4. Stesura del corpo del testo (Descrizione normale o continuazione blocchi)
    var p = foglio.appendParagraph(testo);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setFontSize(11);
    p.setItalic(false);
    p.setBold(false);
    p.setSpacingBefore(0);
    p.setSpacingAfter(12);
    
    if (sezioneAttuale === "significato") {
      p.setBackgroundColor("#e8f5e9");
      p.setForegroundColor("#1b5e20");
    } else if (sezioneAttuale === "etimologia") {
      p.setBackgroundColor("#fff3e0");
      p.setForegroundColor("#e65100");
    } else {
      p.setBackgroundColor(null);
      p.setForegroundColor("#000000");
    }
  }
  
  Logger.log("Fatto! La parola " + parola + " è stata salvata con i blocchi colorati corretti.");
}
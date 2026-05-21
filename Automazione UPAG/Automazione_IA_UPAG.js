function salvaParolaConAI() {
  // 1. INSERISCI I TUOI DATI QUI
  var idDocumento = "INSERISCI_QUI_IL_TUO_ID_DOCUMENTO"; 
  var apiKeyGemini = "INSERISCI_QUI_LA_TUA_API_KEY";
  
  // 2. Lettura Mail
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
  
  // Estraiamo la parola per usarla come Titolo
  var parolaEstratta = urlSito.match(/\/significato\/([^\?]+)/);
  var parola = parolaEstratta ? parolaEstratta[1].toUpperCase() : "PAROLA DEL GIORNO";
  
  // 3. Scaricamento della pagina web
  var htmlSito = UrlFetchApp.fetch(urlSito).getContentText();
  
  // 4. PULIZIA BASE (Taglio del rumore)
  // Rimuoviamo enormi blocchi inutili per l'AI: script, stili, menu di navigazione, header e footer
  var testoGrezzo = htmlSito.replace(/<(script|style|nav|header|footer|form)[^>]*>[\s\S]*?<\/\1>/gi, ' ');
  // Ora rimuoviamo tutti i tag HTML rimanenti per tenere solo le parole scritte
  testoGrezzo = testoGrezzo.replace(/<[^>]+>/g, '\n');
  // Puliamo gli spazi e compattiamo
  testoGrezzo = testoGrezzo.replace(/&[a-zA-Z0-9]+;/g, ' ').replace(/[ \t]+/g, ' ');
  testoGrezzo = testoGrezzo.replace(/\n\s*\n/g, '\n').trim();
  
  // Tagliamo i primi 10.000 caratteri per stare abbondantemente nei limiti e non dare testo superfluo
  testoGrezzo = testoGrezzo.substring(0, 10000); 

  // 5. CHIAMATA ALL'INTELLIGENZA ARTIFICIALE (LLM)
  var prompt = "Sei un assistente linguistico esperto. Ti fornirò il testo estratto dalla pagina di un dizionario online riguardante la parola '" + parola + "'. Il testo contiene del 'rumore' (bottoni, menu). Ignora il rumore ed estrai l'essenza della spiegazione. Scrivi un riassunto discorsivo, fluido ed elegante che contenga:\n1. Il significato e il contesto d'uso.\n2. La spiegazione dell'etimologia.\n3. Un esempio di utilizzo pratico (mettilo tra virgolette « »).\nNon usare markdown, asterischi o elenchi puntati: scrivi in paragrafi normali e puliti. Ecco il testo:\n\n" + testoGrezzo;
  
  var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKeyGemini;
  
  var payloadDati = {
    "contents": [{
      "parts": [{"text": prompt}]
    }]
  };
  
  var opzioniApi = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payloadDati),
    "muteHttpExceptions": true
  };
  
  var rispostaApi = UrlFetchApp.fetch(apiUrl, opzioniApi);
  var jsonRisposta = JSON.parse(rispostaApi.getContentText());
  
  if (jsonRisposta.error) {
    Logger.log("Errore API Gemini: " + jsonRisposta.error.message);
    return;
  }
  
  // Estraiamo il testo generato dall'AI
  var riassuntoAI = jsonRisposta.candidates[0].content.parts[0].text;
  
  // Ripuliamo eventuali asterischi del markdown se l'AI dovesse metterli per sbaglio
  riassuntoAI = riassuntoAI.replace(/\*\*/g, '').replace(/\*/g, '');

  // 6. SCRITTURA NEL DOCUMENTO GOOGLE
  var documento = DocumentApp.openById(idDocumento);
  var foglio = documento.getBody();
  
  if (foglio.getText().trim().length > 0) {
    foglio.appendPageBreak();
  }
  
  // Titolo (Parola)
  var pTitolo = foglio.appendParagraph(parola);
  pTitolo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  pTitolo.setForegroundColor("#4a148c"); // Viola scuro (stile AI!)
  pTitolo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pTitolo.setSpacingAfter(6);
  
  // Dicitura "Riassunto con AI"
  var dataOggi = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  var pSottotitolo = foglio.appendParagraph("Riassunto dall'AI il: " + dataOggi);
  pSottotitolo.setItalic(true).setForegroundColor("#9e9e9e");
  pSottotitolo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pSottotitolo.setSpacingAfter(24);
  
  // Stesura del riassunto
  var paragrafiAI = riassuntoAI.split('\n');
  for (var i = 0; i < paragrafiAI.length; i++) {
    var pTesto = paragrafiAI[i].trim();
    if (pTesto === "") continue;
    
    var p = foglio.appendParagraph(pTesto);
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setFontSize(11);

    // FORZA IL COLORE NERO SPEZZANDO L'EREDITÀ DEL SOTTOTITOLO GRIGIO
    p.setForegroundColor("#000000");
    
    // Se è un esempio (inizia con virgolette), lo facciamo grigino per staccarlo
    if (pTesto.startsWith("«") || pTesto.startsWith("\"")) {
      p.setItalic(true);
      p.setBackgroundColor("#f5f5f5");
      p.setIndentStart(24);
      p.setIndentEnd(24);
    } else {
      p.setBackgroundColor(null);
      p.setItalic(false);
      p.setIndentStart(0);
      p.setIndentEnd(0);
    }
    p.setSpacingAfter(12);
  }
  
  Logger.log("Fatto! La parola " + parola + " è stata riassunta dall'AI e salvata.");
}


function scopriModelliDisponibili() {
  // Inserisci qui la tua API Key
  var apiKeyGemini = "INSERISCI_QUI_LA_TUA_API_KEY"; 
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKeyGemini;
  var risposta = UrlFetchApp.fetch(url);
  var json = JSON.parse(risposta.getContentText());
  
  Logger.log("Ecco i modelli che puoi usare:");
  for (var i = 0; i < json.models.length; i++) {
    var nome = json.models[i].name;
    // Filtriamo solo i modelli in grado di generare testo
    if (json.models[i].supportedGenerationMethods.indexOf("generateContent") !== -1) {
      Logger.log(nome);
    }
  }
}
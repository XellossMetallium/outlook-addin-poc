/*
 * ===================================================================
 * FILE 3: taskpane.js (Invariato)
 * ===================================================================
 */

// L'evento 'Office.onReady' viene eseguito non appena la libreria Office.js
// è stata caricata e l'ambiente Office è pronto a ricevere comandi.
// È l'equivalente di 'document.addEventListener("DOMContentLoaded")' per gli Add-in.
Office.onReady((info) => {
  // Verifichiamo che l'add-in sia in esecuzione in Outlook (Host) e non in Word, Excel, etc.
  if (info.host === Office.HostType.Outlook) {
    // Rendiamo disponibili i controlli dell'interfaccia
    document.getElementById("extractButton").disabled = false;
    
    // Associamo la funzione 'extractEmailData' all'evento 'click' del nostro pulsante.
    document.getElementById("extractButton").onclick = extractEmailData;
  }
});

/**
 * Funzione principale per estrarre e visualizzare i dati dell'email.
 */
function extractEmailData() {
  const status = document.getElementById("status");
  const output = document.getElementById("output");
  
  // Resettiamo l'interfaccia per una nuova estrazione
  status.textContent = "Elaborazione in corso...";
  output.textContent = "";

  // 'Office.context.mailbox.item' è l'oggetto principale che rappresenta
  // l'elemento correntemente selezionato in Outlook (in questo caso, un'email).
  const item = Office.context.mailbox.item;

  // Molte proprietà sono disponibili direttamente e in modo sincrono.
  const emailData = {
    outlookId: item.itemId,
    subject: item.subject,
    dateTimeCreated: item.dateTimeCreated.toISOString(),
    from: {
      name: item.from.displayName,
      email: item.from.emailAddress,
    },
    // Le liste di destinatari sono array di oggetti. Usiamo '.map' per trasformarle
    // in un formato più pulito e leggibile.
    to: item.to.map(recipient => ({ name: recipient.displayName, email: recipient.emailAddress })),
    cc: item.cc.map(recipient => ({ name: recipient.displayName, email: recipient.emailAddress })),
    // Controlliamo anche la presenza di allegati.
    hasAttachments: item.attachments.length > 0,
    attachmentsInfo: item.attachments.map(att => ({ name: att.name, size: att.size, type: att.attachmentType }))
  };

  // Il corpo dell'email, invece, viene recuperato in modo asincrono per motivi di performance.
  // Questo previene il blocco dell'interfaccia utente mentre Outlook recupera corpi di grandi dimensioni.
  // Chiediamo il corpo in formato testo semplice ('Office.CoercionType.Text').
  item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // Se la chiamata ha successo, aggiungiamo il corpo all'oggetto dei dati.
      emailData.body = asyncResult.value.trim().substring(0, 500) + '...'; // Tronchiamo per leggibilità
      status.textContent = "Dati estratti con successo!";
    } else {
      // In caso di errore, lo segnaliamo.
      emailData.body = "Errore nel recupero del corpo dell'email.";
      status.textContent = "Errore durante l'estrazione.";
      console.error(asyncResult.error.message);
    }
    
    // Una volta completata anche l'operazione asincrona, visualizziamo l'intero
    // oggetto JSON nell'elemento <pre> della nostra pagina HTML.
    // 'JSON.stringify' con i parametri 2 e ' ' formatta l'output in modo che sia ben indentato e leggibile.
    output.textContent = JSON.stringify(emailData, null, 2);
  });
}

# V.I.A. TCG — app gestionale

Interfaccia dell'app gestionale del progetto V.I.A. TCG (carte e prodotti One Piece).

- **App:** https://siralex03.github.io/viatcg-app/ — da Safari: Condividi → *Aggiungi alla schermata Home*.
- **Dati:** restano nel Google Sheet privato del progetto. Questo repository contiene solo l'interfaccia: nessun dato, nessun PIN.
- **Motore:** script Google Apps Script (copia in `apps-script/`), pubblicato come applicazione web; l'app lo chiama con `fetch`.

## Struttura
- `index.html` — l'intera app (HTML, CSS, JavaScript)
- `manifest.webmanifest`, `sw.js`, icone — installazione come app
- `apps-script/Codice.gs` — motore (il PIN iniziale si imposta solo nell'editor di Apps Script)
- `apps-script/Index.html` — pagina che dal vecchio link rimanda alla nuova app

## Aggiornare il motore senza cambiare link
Nell'editor di Apps Script: Esegui il deployment → Gestisci deployment → matita → Versione: **Nuova versione**.

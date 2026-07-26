# Repository Referenze Progettuali — versione HTML statica (senza build)

Stessa applicazione della versione Vite, ma pensata per il deploy più semplice possibile: **tre file
statici**, nessun `npm install`, nessuno step di build. React, Firebase, xlsx e mammoth vengono
caricati da CDN (esm.sh) tramite import map del browser. Il JSX è già stato compilato in anticipo
in `app.js` (JavaScript puro), quindi non c'è compilazione a runtime nel browser.

## File

```
index.html   — pagina + import map (CDN)
main.js      — avvio React (monta App su #root)
app.js       — l'intera applicazione (già compilata da JSX a JS puro)
```

## Come provarlo in locale

Serve un server statico qualsiasi (non si può aprire index.html direttamente da file:// perché i
moduli ES lo impediscono per motivi di sicurezza del browser):

```bash
npx serve .
# oppure
python3 -m http.server 8080
```

Poi apri l'indirizzo mostrato in console.

## Deploy su GitHub Pages (il caso d'uso più semplice)

1. Crea un repository su GitHub e carica questi 3 file (più eventualmente questo README) nella radice
   o in una cartella `docs/`.
2. Nel repository: Settings → Pages → Source → seleziona il branch e la cartella dove si trovano i file.
3. GitHub Pages pubblica automaticamente l'URL statico — nessuna build da configurare.

Va bene allo stesso modo qualunque altro hosting statico (Firebase Hosting, Netlify, Vercel...).

## Configurazione Firebase

`app.js` contiene già la configurazione del progetto Firebase dello Staffing (`staffing-portal-eeef6`),
scritta direttamente nel file. Per un sito statico senza build questa è la prassi normale: la
configurazione client di Firebase **è pubblica per design** — la sicurezza reale è affidata alle
Security Rules del database e dello Storage, non alla segretezza di questi valori. Non contiene
invece nessuna chiave segreta.

## Immagini e loghi: Firebase Storage, non Realtime Database

Le immagini delle referenze (architetture, screenshot) e i loghi cliente sono salvati come file su
**Firebase Storage** — nel Realtime Database viene salvato solo il percorso del file. Caricamento e
scaricamento passano sempre dall'SDK autenticato (`uploadString` / `getBlob`), mai da un link
pubblico condivisibile: l'accesso a ogni file viene verificato dalle Storage Rules a ogni richiesta.

Regole consigliate (Console Firebase → Storage → Rules):

```
rules_version = '2';
service firebase.storage {
  match /b/{bucket}/o {
    match /referenze/{allPaths=**} {
      allow read, write: if request.auth != null;
    }
  }
}
```

Limite noto: queste regole verificano solo che l'utente sia autenticato, non il suo `gruppo` (le
Storage Rules non possono leggere direttamente il Realtime Database). Per un controllo per gruppo
servirebbero custom claims impostati lato server.

## Importazione in blocco con estrazione automatica di loghi e immagini

Nella scheda "Importa in blocco", allegando **un solo file PDF**, l'AI indica anche in quali pagine
si trovano il logo del cliente (per i clienti nuovi) e diagrammi/screenshot rilevanti per ciascuna
referenza. Quelle pagine vengono renderizzate come immagini direttamente nel browser tramite
`pdf.js` (nessun servizio esterno coinvolto) e allegate automaticamente al momento della conferma.
Funziona solo con un singolo PDF allegato, non con testo libero o altri formati.

## ⚠️ Cose da fare prima del go-live

1. **Security Rules**: verifica in Console Firebase che permettano, per utenti autenticati, la lettura
   di `/users/{auth.uid}` e lettura/scrittura di `/referenze/**` (Realtime Database) e del percorso
   `referenze/**` su Storage, senza allargare i permessi sugli altri nodi già in uso dallo Staffing.
2. **Motore AI (Google Gemini via proxy Cloudflare)**: le funzioni "Struttura con AI", l'Assistente AI e
   l'Importazione in blocco chiamano un proxy Cloudflare Worker che inoltra a Google Gemini — il codice
   del proxy è nella cartella `cloudflare-worker/` accanto a questa (vedi il suo README per il deploy,
   che richiede un account Cloudflare gratuito e una API key Gemini da
   [Google AI Studio](https://aistudio.google.com/app/apikey)). Dopo il deploy, apri `app.js`, cerca la
   costante `AI_PROXY_URL` e sostituiscila con l'URL reale del Worker. Finché questo non è fatto, il
   resto dell'app (repository, anagrafica, filtri, export, loghi, Programma/Servizio AMS/Progetto)
   funziona comunque normalmente; solo le tre funzioni AI resteranno inattive.

## Nota su questo ambiente di sviluppo

Il codice è stato verificato per correttezza sintattica, ma il test end-to-end nel browser (caricamento
reale da esm.sh, login Firebase, lettura/scrittura sul Realtime Database) non è stato possibile
dall'ambiente in cui è stato preparato questo pacchetto, che non ha accesso di rete verso esm.sh.
Va quindi validato in un vero browser prima del go-live, sui punti sopra elencati.

## Aggiornare l'app in futuro

Se in futuro serve modificare la logica applicativa, il sorgente JSX leggibile resta quello del
progetto Vite (`repository-referenze-app/src/App.jsx`); va poi ricompilato in `app.js` con Babel
(`@babel/preset-react`, `runtime: "classic"`) per rigenerare questa versione statica.

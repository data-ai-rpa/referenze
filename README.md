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
Security Rules del database, non alla segretezza di questi valori. Non contiene invece nessuna chiave
segreta.

## ⚠️ Stesse due limitazioni della versione Vite

1. **Security Rules**: verifica in Console Firebase che permettano, per utenti autenticati, la lettura
   di `/users/{auth.uid}` e lettura/scrittura di `/referenze/**`, senza allargare i permessi sugli
   altri nodi già in uso dallo Staffing.
2. **Funzioni AI**: "Struttura con AI" e l'Assistente AI chiamano `api.anthropic.com` direttamente dal
   browser — funzionava solo dentro la sandbox degli artifact di Claude.ai, che iniettava
   l'autenticazione automaticamente. Qui serve un backend/proxy proprio con una vera chiave API
   Anthropic lato server (mai nel codice client) perché queste funzioni tornino operative. Il resto
   dell'app (repository, anagrafica, filtri, export, loghi, Programma/Servizio AMS/Progetto) funziona
   normalmente senza alcun proxy.

## Nota su questo ambiente di sviluppo

Il codice è stato verificato per correttezza sintattica, ma il test end-to-end nel browser (caricamento
reale da esm.sh, login Firebase, lettura/scrittura sul Realtime Database) non è stato possibile
dall'ambiente in cui è stato preparato questo pacchetto, che non ha accesso di rete verso esm.sh.
Va quindi validato in un vero browser prima del go-live, sui punti sopra elencati.

## Aggiornare l'app in futuro

Se in futuro serve modificare la logica applicativa, il sorgente JSX leggibile resta quello del
progetto Vite (`repository-referenze-app/src/App.jsx`); va poi ricompilato in `app.js` con Babel
(`@babel/preset-react`, `runtime: "classic"`) per rigenerare questa versione statica.

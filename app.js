import React, { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import { initializeApp } from "firebase/app";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { getDatabase, ref, get as fbGet, set as fbSet, remove as fbRemove } from "firebase/database";

// ─────────────────────────────────────────────────────────────
//  REPOSITORY REFERENZE PROGETTUALI — Practice Data, AI & RPA
//  Alimentazione: file (PDF/DOCX/XLSX/TXT) · testo libero · voce
//  Anagrafica clienti centralizzata + anonimizzazione dinamica NDA
//
//  ATTENZIONE — QUESTO FILE NON GIRA PIÙ NELLA SANDBOX DEGLI ARTIFACT
//  DI CLAUDE.AI: usa Firebase (rete esterna), non disponibile lì.
//  Va eseguito nel tuo ambiente reale (hosting Firebase, Vite/CRA, ecc.),
//  con `npm install firebase` tra le dipendenze.
//
//  Login: Firebase Authentication (email/password), stesso progetto
//  del portale Staffing (staffing-portal-eeef6). Non crea né modifica
//  utenti: si aspetta che esistano già in Firebase Auth e nel nodo
//  /users/{uid} del Realtime Database, con la stessa struttura usata
//  dallo Staffing: { email, gruppo }.
//
//  Persistenza dati: Firebase Realtime Database, sotto il nodo radice
//  "referenze" — separato e non sovrapposto ai nodi già usati dallo
//  Staffing (verificare in Console che "referenze" non sia già in uso).
// ─────────────────────────────────────────────────────────────

// ── configurazione Firebase (stesso progetto del portale Staffing) ──
// Valori letti da variabili d'ambiente (vedi .env.example): mai committare un .env reale.
const firebaseConfig = {
  apiKey: "AIzaSyBIiF5LoFJelxE1AA2ljNFvlEeTV4k0aoE",
  authDomain: "staffing-portal-eeef6.firebaseapp.com",
  databaseURL: "https://staffing-portal-eeef6-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "staffing-portal-eeef6",
  storageBucket: "staffing-portal-eeef6.firebasestorage.app",
  messagingSenderId: "111457601788",
  appId: "1:111457601788:web:0b06d45f52c1acbb9caf79"
};
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getDatabase(firebaseApp);

// radice isolata per non toccare i dati già usati dallo Staffing
const ROOT = "referenze";

// converte le chiavi usate finora (es. "immagini:ref_123") in path Firebase validi
const toPath = key => `${ROOT}/${key.replace(/:/g, "/")}`;

// shim con la stessa interfaccia usata finora (get/set/delete), ma su Firebase Realtime Database
const storage = {
  async get(key) {
    const snap = await fbGet(ref(db, toPath(key)));
    if (!snap.exists()) throw new Error("Chiave non presente: " + key);
    return {
      key,
      value: snap.val()
    };
  },
  async set(key, value) {
    await fbSet(ref(db, toPath(key)), value);
    return {
      key,
      value
    };
  },
  async delete(key) {
    await fbRemove(ref(db, toPath(key)));
    return {
      key,
      deleted: true
    };
  }
};
const REFS_KEY = "referenze-repo-v1";
const CLIENTS_KEY = "referenze-clienti-v1";
const AMBITI = ["Data", "AI", "RPA", "Cloud", "App", "Altro"];
const SETTORI = ["Energia & Utilities", "Pubblica Amministrazione", "Banche & Assicurazioni", "Telco & Media", "Industria", "Trasporti", "Sanità", "Retail", "Altro"];
const PAESI_COMUNI = ["Italia", "Spagna", "Francia", "Germania", "Portogallo", "Regno Unito", "Perù", "Colombia"];
const AMBITO_COLOR = {
  Data: "#3E8E9C",
  AI: "#C2703D",
  RPA: "#7A6BA8",
  Cloud: "#4E7A4E",
  App: "#9C5C6E",
  Altro: "#6B7280"
};

// ── dominio modalità di fornitura ──────────────────────────────
const MODALITA_FORNITURA = [{
  id: "unico",
  label: "Fornitore unico"
}, {
  id: "rti_mandataria",
  label: "RTI — Mandataria / Capofila"
}, {
  id: "rti_mandante",
  label: "RTI — Mandante"
}, {
  id: "subappalto",
  label: "Subappalto"
}];
const richiedeQuota = m => m !== "unico";
const etichettaFornitura = r => {
  const m = MODALITA_FORNITURA.find(x => x.id === r.modalitaFornitura) || MODALITA_FORNITURA[0];
  if (r.modalitaFornitura === "unico") return m.label;
  const quota = r.quotaPercentuale !== "" && r.quotaPercentuale != null ? ` ${r.quotaPercentuale}%` : "";
  const autonoma = r.attivitaAutonoma ? " · attività svolta in autonomia" : "";
  return `${m.label}${quota}${autonoma}`;
};

// ── dominio tipo referenza: Programma e Servizio AMS sono attività di livello superiore (contenitori);
//    Progetto è l'attività evolutiva agganciata a un Programma o a un Servizio AMS; Standalone è indipendente ──
const TIPO_REFERENZA = [{
  id: "standalone",
  label: "Referenza singola (progetto o servizio autonomo)"
}, {
  id: "programma",
  label: "Programma (iniziativa ampia, contenitore di più progetti)"
}, {
  id: "servizio_ams",
  label: "Servizio AMS (assistenza/manutenzione applicativa continuativa)"
}, {
  id: "progetto",
  label: "Progetto (attività evolutiva nell'ambito di un Programma o Servizio AMS)"
}];
const TIPI_CONTENITORE = ["programma", "servizio_ams"];
const puoEssereGenitore = tipo => TIPI_CONTENITORE.includes(tipo);
const etichettaTipoReferenza = tipo => (TIPO_REFERENZA.find(t => t.id === tipo) || TIPO_REFERENZA[0]).label;

// ── date inizio/fine (mese/anno) e durata calcolata dinamicamente ──
const MESI = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
const MESI_BREVI = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"];

// mese/anno di fine effettivi: se il progetto è in corso, usa il mese/anno correnti (oggi)
const dataFineEffettiva = r => {
  if (r.inCorso) {
    const oggi = new Date();
    return {
      mese: oggi.getMonth() + 1,
      anno: oggi.getFullYear()
    };
  }
  return {
    mese: r.fineMese || r.inizioMese,
    anno: r.fineAnno || r.inizioAnno
  };
};

// durata in mesi, calcolata dinamicamente (inclusiva del mese di inizio)
const calcolaDurataMesi = r => {
  if (!r.inizioAnno || !r.inizioMese) return null;
  const fine = dataFineEffettiva(r);
  if (!fine.anno || !fine.mese) return null;
  const mesi = (fine.anno - r.inizioAnno) * 12 + (fine.mese - r.inizioMese) + 1;
  return mesi > 0 ? mesi : null;
};
const fmtPeriodo = r => {
  if (!r.inizioAnno || !r.inizioMese) return "n.d.";
  const ini = `${MESI_BREVI[r.inizioMese - 1]} ${r.inizioAnno}`;
  if (r.inCorso) return `${ini} – in corso`;
  if (!r.fineAnno || !r.fineMese) return ini;
  return `${ini} – ${MESI_BREVI[r.fineMese - 1]} ${r.fineAnno}`;
};

// formato compatto MM/AAAA -> MM/AAAA (o Today se in corso), per la barra laterale della card
const fmtPeriodoSlash = r => {
  if (!r.inizioAnno || !r.inizioMese) return "n.d.";
  const pad = n => String(n).padStart(2, "0");
  const ini = `${pad(r.inizioMese)}/${r.inizioAnno}`;
  if (r.inCorso) return `${ini} → Today`;
  if (!r.fineAnno || !r.fineMese) return ini;
  return `${ini} → ${pad(r.fineMese)}/${r.fineAnno}`;
};

// ── anagrafica clienti: demo iniziale ──────────────────────────
const DEMO_CLIENTS = [{
  id: "cli_1",
  nome: "Enermesh S.p.A.",
  descrizione: "Operatore nazionale della rete di trasmissione elettrica ad alta tensione, responsabile della pianificazione, gestione e sviluppo dell'infrastruttura di trasmissione sul territorio.",
  labelBreve: "Primario operatore energetico nazionale",
  nda: true
}, {
  id: "cli_2",
  nome: "AcquaViva Multiutility S.p.A.",
  descrizione: "Multiutility italiana attiva nei servizi idrici, ambientali e di igiene urbana su base regionale, con oltre un milione di utenze servite.",
  labelBreve: "Multiutility servizi ambientali e idrici",
  nda: true
}, {
  id: "cli_3",
  nome: "Vetrio Trasporti Nazionali S.p.A.",
  descrizione: "Società pubblica di gestione delle infrastrutture di trasporto ferroviario, responsabile di manutenzione, pianificazione e digitalizzazione della rete.",
  labelBreve: "Operatore pubblico trasporti su ferro",
  nda: false
}, {
  id: "cli_bde",
  nome: "Banco de España (BdE)",
  descrizione: "Banca centrale e autorità di vigilanza bancaria spagnola.",
  labelBreve: "Banca centrale e autorità di vigilanza",
  nda: false
}, {
  id: "cli_barclays",
  nome: "Barclays",
  descrizione: "Gruppo bancario internazionale, con operatività anche sul mercato spagnolo.",
  labelBreve: "Gruppo bancario internazionale",
  nda: false
}, {
  id: "cli_bcc",
  nome: "BCC",
  descrizione: "Istituto finanziario spagnolo (denominazione riportata in forma abbreviata nella documentazione originale).",
  labelBreve: "Istituto finanziario",
  nda: false
}, {
  id: "cli_vodafone_es",
  nome: "Vodafone España",
  descrizione: "Operatore di telecomunicazioni attivo sul mercato spagnolo nei segmenti privati, PMI e autonomi.",
  labelBreve: "Operatore di telecomunicazioni",
  nda: false
}, {
  id: "cli_yoigo",
  nome: "Yoigo",
  descrizione: "Operatore di telecomunicazioni mobile spagnolo.",
  labelBreve: "Operatore di telecomunicazioni mobile",
  nda: false
}, {
  id: "cli_sabadell",
  nome: "Banco Sabadell",
  descrizione: "Gruppo bancario spagnolo attivo su privati, imprese e recupero crediti.",
  labelBreve: "Gruppo bancario spagnolo",
  nda: false
}, {
  id: "cli_bbva",
  nome: "BBVA",
  descrizione: "Gruppo bancario internazionale di origine spagnola.",
  labelBreve: "Gruppo bancario internazionale",
  nda: false
}, {
  id: "cli_pontevedra",
  nome: "Diputación de Pontevedra",
  descrizione: "Ente provinciale spagnolo (Galizia), con competenze anche in materia di turismo regionale.",
  labelBreve: "Ente provinciale",
  nda: false
}, {
  id: "cli_micinn",
  nome: "Ministerio de Ciencia e Innovación",
  descrizione: "Ministero spagnolo per la scienza e l'innovazione.",
  labelBreve: "Ministero nazionale",
  nda: false
}, {
  id: "cli_interior_es",
  nome: "Ministerio del Interior (Spagna)",
  descrizione: "Ministero dell'Interno spagnolo.",
  labelBreve: "Ministero nazionale",
  nda: false
}, {
  id: "cli_atc",
  nome: "Agència Tributària de Catalunya (ATC)",
  descrizione: "Agenzia tributaria della Catalogna.",
  labelBreve: "Agenzia tributaria regionale",
  nda: false
}, {
  id: "cli_diue",
  nome: "DIUE - Departament d'Innovació, Universitats i Empresa (Generalitat de Catalunya)",
  descrizione: "Dipartimento della Generalitat de Catalunya competente su innovazione, università e imprese.",
  labelBreve: "Dipartimento regionale",
  nda: false
}, {
  id: "cli_madrid_edu",
  nome: "Consejería de Educación, Juventud y Deporte de la Comunidad de Madrid",
  descrizione: "Assessorato regionale all'Istruzione, Gioventù e Sport della Comunità di Madrid.",
  labelBreve: "Assessorato regionale",
  nda: false
}, {
  id: "cli_gipuzkoa",
  nome: "Agencia Tributaria de Guipúzcoa (Hacienda Foral)",
  descrizione: "Agenzia tributaria forale del territorio di Gipuzkoa (Paesi Baschi).",
  labelBreve: "Agenzia tributaria regionale",
  nda: false
}, {
  id: "cli_icm",
  nome: "Comunidad de Madrid - Consejería de Hacienda (ICM)",
  descrizione: "Assessorato regionale alle Finanze della Comunità di Madrid.",
  labelBreve: "Assessorato regionale",
  nda: false
}, {
  id: "cli_canarias_trib",
  nome: "Agencia Tributaria de Canarias",
  descrizione: "Agenzia tributaria delle Canarie.",
  labelBreve: "Agenzia tributaria regionale",
  nda: false
}, {
  id: "cli_onp_peru",
  nome: "Oficina de Normalización Previsional (ONP) - Perù",
  descrizione: "Ente pubblico peruviano responsabile della gestione previdenziale pensionistica.",
  labelBreve: "Ente previdenziale pubblico",
  nda: false
}, {
  id: "cli_justizia_eus",
  nome: "Administración de Justicia de Euskadi",
  descrizione: "Amministrazione della Giustizia dei Paesi Baschi.",
  labelBreve: "Amministrazione giudiziaria regionale",
  nda: false
}, {
  id: "cli_judicatura_ec",
  nome: "Consejo de la Judicatura - Ecuador",
  descrizione: "Organo di governo della magistratura dell'Ecuador.",
  labelBreve: "Organo di governo della magistratura",
  nda: false
}, {
  id: "cli_justicia_es",
  nome: "Ministerio de Justicia (Fiscalía General del Estado)",
  descrizione: "Ministero della Giustizia spagnolo, con riferimento alla Procura Generale dello Stato.",
  labelBreve: "Ministero nazionale",
  nda: false
}, {
  id: "cli_agricultura",
  nome: "Dipartimento/Ministero dell'Agricoltura",
  descrizione: "Ente pubblico competente in materia agricola (denominazione esatta e paese non specificati con certezza nel documento originale).",
  labelBreve: "Ente pubblico agricolo",
  nda: false
}, {
  id: "cli_aena",
  nome: "AENA",
  descrizione: "Gestore spagnolo degli aeroporti, tra i principali operatori aeroportuali a livello internazionale.",
  labelBreve: "Gestore aeroportuale nazionale",
  nda: false
}, {
  id: "cli_correos",
  nome: "Correos",
  descrizione: "Operatore postale nazionale spagnolo.",
  labelBreve: "Operatore postale nazionale",
  nda: false
}, {
  id: "cli_sanidad_es",
  nome: "Ministerio de Sanidad, Servicios Sociales e Igualdad",
  descrizione: "Ministero spagnolo della Salute, dei Servizi Sociali e delle Pari Opportunità.",
  labelBreve: "Ministero nazionale",
  nda: false
}, {
  id: "cli_quironsalud",
  nome: "Grupo Quirónsalud",
  descrizione: "Uno dei principali gruppi ospedalieri privati spagnoli.",
  labelBreve: "Gruppo ospedaliero privato",
  nda: false
}, {
  id: "cli_sergas",
  nome: "Servizo Galego de Saúde (SERGAS)",
  descrizione: "Servizio sanitario pubblico della Galizia.",
  labelBreve: "Servizio sanitario regionale",
  nda: false
}, {
  id: "cli_sanidad_madrid",
  nome: "Consejería de Sanidad de la Comunidad de Madrid",
  descrizione: "Assessorato regionale alla Sanità della Comunità di Madrid.",
  labelBreve: "Assessorato regionale alla sanità",
  nda: false
}, {
  id: "cli_avs",
  nome: "Agencia Valenciana de Salud",
  descrizione: "Servizio sanitario pubblico della Comunità Valenciana.",
  labelBreve: "Servizio sanitario regionale",
  nda: false
}, {
  id: "cli_sas",
  nome: "Servicio Andaluz de Salud (SAS)",
  descrizione: "Servizio sanitario pubblico dell'Andalusia.",
  labelBreve: "Servizio sanitario regionale",
  nda: false
}, {
  id: "cli_saludresponde",
  nome: "Salud Responde - Consejería de Andalucía",
  descrizione: "Servizio di contact center sanitario della Junta de Andalucía.",
  labelBreve: "Contact center sanitario regionale",
  nda: false
}, {
  id: "cli_sanidad_cyl",
  nome: "Consejería de Sanidad de Castilla y León",
  descrizione: "Assessorato regionale alla Sanità di Castiglia e León.",
  labelBreve: "Assessorato regionale alla sanità",
  nda: false
}, {
  id: "cli_hlf_chile",
  nome: "Hospital de la Florida (Santiago del Cile)",
  descrizione: "Struttura ospedaliera pubblica di Santiago del Cile.",
  labelBreve: "Struttura ospedaliera pubblica",
  nda: false
}];

// ── referenze demo, collegate all'anagrafica sopra ─────────────
const DEMO_REFS = [{
  id: "demo_1",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Piattaforma Dati e Analytics per la rete di trasmissione",
  clienteId: "cli_1",
  cliente: "",
  // fallback libero, usato solo se clienteId non risolve
  settore: "Energia & Utilities",
  paese: "Italia",
  inizioMese: 1,
  inizioAnno: 2023,
  fineMese: 6,
  fineAnno: 2024,
  inCorso: false,
  importoKEuro: 950,
  ambiti: ["Data", "AI"],
  teamSize: 12,
  ruolo: "System integrator, capofila RTI",
  modalitaFornitura: "rti_mandataria",
  quotaPercentuale: 60,
  partnerRTI: "NTT Data (40%)",
  attivitaAutonoma: false,
  tecnologie: ["Azure Synapse", "Databricks", "Power BI", "Python", "MLflow"],
  descrizione: "Enermesh S.p.A. necessitava di consolidare in un'unica piattaforma i dati provenienti da sistemi SCADA, asset management e mercato, al fine di abilitare analisi avanzate a supporto della pianificazione della rete e della manutenzione predittiva.",
  attivita: "Per Enermesh S.p.A. sono state svolte le attività di disegno dell'architettura dati su cloud Azure, migrazione del data warehouse legacy verso un modello lakehouse, sviluppo di pipeline di ingestion in near real-time e realizzazione di modelli di machine learning per la previsione dei guasti sugli asset di rete. Il progetto ha incluso inoltre la definizione del framework di data governance.",
  risultati: "Riduzione del 35% dei tempi di predisposizione della reportistica regolatoria, incremento del 20% della capacità di individuazione anticipata delle anomalie sugli asset critici e consolidamento di oltre 40 fonti dati in un unico punto di accesso governato.",
  nda: true
}, {
  id: "demo_2",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Automazione RPA dei processi amministrativi e di fatturazione",
  clienteId: "cli_2",
  cliente: "",
  settore: "Energia & Utilities",
  paese: "Italia",
  inizioMese: 1,
  inizioAnno: 2023,
  fineMese: 12,
  fineAnno: 2023,
  inCorso: false,
  importoKEuro: 380,
  ambiti: ["RPA", "AI"],
  teamSize: 6,
  ruolo: "Mandante RTI, unico esecutore dello stream RPA",
  modalitaFornitura: "rti_mandante",
  quotaPercentuale: 30,
  partnerRTI: "Capofila: Vantis Consulting",
  attivitaAutonoma: true,
  tecnologie: ["UiPath", "Document Understanding", "SAP", "SQL Server"],
  descrizione: "AcquaViva Multiutility S.p.A. intendeva ridurre l'effort manuale dei processi amministrativi ad alta ripetitività, con particolare riferimento al ciclo passivo, alla gestione dei reclami e alla riconciliazione delle partite di fatturazione.",
  attivita: "È stato condotto per AcquaViva Multiutility S.p.A. un assessment di automabilità su oltre 30 processi candidati, con successiva progettazione e sviluppo di 14 automazioni RPA integrate con SAP e con i sistemi di CRM e billing. Per la lettura dei documenti passivi è stata adottata una componente di intelligent document processing, ed è stato istituito un Centro di Eccellenza interno.",
  risultati: "Automatizzate oltre 120.000 pratiche annue con un tasso di straight-through processing superiore all'85%, recupero stimato di circa 9 FTE riallocati su attività a maggior valore e riduzione dei tempi medi di gestione dei reclami da 12 a 4 giorni lavorativi.",
  nda: true
}, {
  id: "demo_3",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Programma pluriennale di trasformazione Data & AI",
  clienteId: "cli_3",
  cliente: "",
  settore: "Trasporti",
  paese: "Italia",
  inizioMese: 1,
  inizioAnno: 2023,
  fineMese: 12,
  fineAnno: 2026,
  inCorso: true,
  importoKEuro: 4000,
  ambiti: ["Data", "AI", "Cloud"],
  teamSize: "",
  ruolo: "Fornitore aggiudicatario del programma",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "programma",
  parentId: null,
  descrizione: "Vetrio Trasporti Nazionali S.p.A. ha avviato un programma pluriennale di trasformazione basato su data engineering, intelligenza artificiale e digitalizzazione dei processi, articolato in molteplici progetti evolutivi attivati nel corso della durata del programma.",
  attivita: "Il programma definisce il perimetro complessivo di intervento, i profili professionali coinvolti, la governance e le modalità di attivazione dei singoli progetti, ciascuno gestito come iniziativa autonoma con proprio piano di lavoro, team dedicato e obiettivi specifici.",
  risultati: "Attivati ad oggi molteplici progetti nell'ambito del programma, con un valore complessivo delle iniziative in costante crescita nel corso della sua durata.",
  nda: false
}, {
  id: "demo_4",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Progetto — Piattaforma predittiva per la manutenzione della flotta",
  clienteId: "cli_3",
  cliente: "",
  settore: "Trasporti",
  paese: "Italia",
  inizioMese: 3,
  inizioAnno: 2023,
  fineMese: 2,
  fineAnno: 2024,
  inCorso: false,
  importoKEuro: 620,
  ambiti: ["Data", "AI"],
  teamSize: 8,
  ruolo: "Fornitore unico del progetto",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Databricks", "Python", "MLflow", "Power BI"],
  tipoReferenza: "progetto",
  parentId: "demo_3",
  descrizione: "Nell'ambito del Programma con Vetrio Trasporti Nazionali S.p.A., è stato attivato un primo progetto per la realizzazione di una piattaforma predittiva a supporto della manutenzione della flotta ferroviaria, con l'obiettivo di ridurre i fermi macchina non pianificati.",
  attivita: "Sviluppo di modelli di manutenzione predittiva basati su dati di sensoristica e storico guasti, realizzazione di dashboard di monitoraggio per i responsabili di manutenzione e integrazione con i sistemi informativi esistenti del cliente.",
  risultati: "Riduzione del 18% dei fermi macchina non pianificati sulla flotta pilota e adozione della piattaforma come standard per l'estensione ad altre linee nell'ambito del programma.",
  nda: false
}, {
  id: "demo_5",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Servizio AMS — Applicativo di billing e fatturazione",
  clienteId: "cli_2",
  cliente: "",
  settore: "Energia & Utilities",
  paese: "Italia",
  inizioMese: 1,
  inizioAnno: 2022,
  fineMese: "",
  fineAnno: "",
  inCorso: true,
  importoKEuro: 60,
  ambiti: ["App"],
  teamSize: 3,
  ruolo: "Fornitore del servizio AMS",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["SAP", "Java", "Oracle"],
  tipoReferenza: "servizio_ams",
  parentId: null,
  descrizione: "AcquaViva Multiutility S.p.A. ha affidato il servizio di assistenza e manutenzione applicativa (AMS) continuativa della piattaforma di billing e fatturazione, comprensivo di gestione incident, manutenzione correttiva e piccola evolutiva.",
  attivita: "Presidio continuativo dell'applicativo con gestione ticket, manutenzione correttiva su malfunzionamenti, piccola manutenzione evolutiva su richiesta e reportistica periodica sui livelli di servizio (SLA).",
  risultati: "Mantenimento costante degli SLA contrattuali con riduzione del tempo medio di risoluzione incident nel corso del servizio.",
  nda: true
}, {
  id: "demo_6",
  fonte: "manuale",
  creato: new Date().toISOString(),
  titolo: "Progetto evolutivo — Modulo di reportistica avanzata",
  clienteId: "cli_2",
  cliente: "",
  settore: "Energia & Utilities",
  paese: "Italia",
  inizioMese: 6,
  inizioAnno: 2023,
  fineMese: 11,
  fineAnno: 2023,
  inCorso: false,
  importoKEuro: 95,
  ambiti: ["App", "Data"],
  teamSize: 4,
  ruolo: "Fornitore unico del progetto",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Java", "Power BI", "Oracle"],
  tipoReferenza: "progetto",
  parentId: "demo_5",
  descrizione: "Nell'ambito del Servizio AMS attivo su AcquaViva Multiutility S.p.A., è stato realizzato un progetto evolutivo per l'introduzione di un modulo di reportistica avanzata sui consumi e sulla fatturazione, richiesto dal cliente come estensione del servizio in essere.",
  attivita: "Analisi dei requisiti con il cliente, sviluppo del nuovo modulo di reportistica integrato con l'applicativo di billing esistente, test e rilascio in produzione nell'ambito del servizio AMS già attivo.",
  risultati: "Adozione del nuovo modulo da parte degli utenti di back-office, con riduzione dei tempi di predisposizione della reportistica sui consumi.",
  nda: true
}, {
  id: "spa_1",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma di Big Data Analytics per il settore finanziario",
  clienteId: "cli_bde",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Nell'ambito dei servizi di Data Warehousing & Analytics per Banco de España, è stata realizzata un'iniziativa di Big Data Analytics a supporto della funzione finanziaria e di vigilanza.",
  attivita: "Creazione del BigData Lab Financiero, migrazione di datamart aziendali, evoluzione dei sistemi di gestione delle transazioni, miglioramento dell'esperienza del cliente, sviluppo di intelligenza operativa e migrazione di processi e tecniche di sfruttamento dei dati.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_2",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sviluppo e validazione di modelli di rischio di credito (PD, LGD, EAD)",
  clienteId: "cli_bde",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Banco de España per lo sviluppo e la validazione di modelli di rischio di credito.",
  attivita: "Sviluppo di modelli PD (Probability of Default), LGD (Loss Given Default) ed EAD (Exposure at Default), sviluppo del modello di Capitale Economico e supporto nel relativo processo di validazione da parte di Banco de España.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_3",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Costruzione, revisione e monitoraggio di modelli di rischio di credito",
  clienteId: "cli_barclays",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Barclays per la costruzione, revisione e monitoraggio di modelli di rischio di credito.",
  attivita: "Costruzione di modelli PD, LGD ed EAD, monitoraggio e follow-up dei modelli, progettazione e realizzazione di un pannello di controllo per il tracciamento di modelli e parametri, revisione dei modelli di accantonamento ai sensi dello IAS39.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_4",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Progettazione, sviluppo e implementazione di modelli di credito",
  clienteId: "cli_bcc",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di BCC per la costruzione di modelli di ammissione e di rischio.",
  attivita: "Implementazione e sviluppo di modelli di rating generici per aziende di medie dimensioni nei processi di ammissione clienti, impianto del modello limite e stop loss, assegnazione dei limiti di rischio per controparte, calcolo del Mark to Market su derivati OTC di energia elettrica.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Denominazione cliente riportata in forma abbreviata (BCC) come da documento originale, senza ulteriori dettagli identificativi. Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_5",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Implementazione del modello di ammissione e dei sistemi di scoring",
  clienteId: "cli_vodafone_es",
  cliente: "",
  settore: "Telco & Media",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Transact", "Strategy Design Studio"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Vodafone España per l'implementazione del modello di ammissione clienti.",
  attivita: "Implementazione del modello di ammissione e monitoraggio dei sistemi di punteggio per i diversi segmenti Vodafone (privati, PMI e autonomi) attraverso Transact e Strategy Design Studio (Transact SM), sviluppo dei test di integrazione delle soluzioni (SIT) e dei test utente (UAT) per il modello.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_6",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Test e manutenzione del modello di scoring di ammissione",
  clienteId: "cli_yoigo",
  cliente: "",
  settore: "Telco & Media",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Yoigo per il test e la manutenzione del modello di scoring di ammissione.",
  attivita: "Test di follow-up sulla validità della scorecard generica implementata per il modello di punteggio Yoigo, comprensivi di analisi delle decisioni e test di stabilità del modello; manutenzione continuativa del modello.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_7",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Ottimizzazione della piattaforma e dei processi di recupero crediti",
  clienteId: "cli_sabadell",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Banco Sabadell per l'ottimizzazione dei processi di recupero crediti.",
  attivita: "Revisione e ottimizzazione delle azioni automatiche di recupero per i punti vendita, miglioramento dei processi e riprogettazione del flusso di lavoro, supporto nella revisione e definizione dei criteri per il deposito delle richieste, segmentazione del portafoglio in base alla probabilità stimata di recupero, definizione del sistema di valutazione delle prestazioni per gli avvocati e miglioramenti nella gestione delle procedure giudiziarie.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_8",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Coordinamento dell'ufficio di integrazione dei servizi di recupero",
  clienteId: "cli_sabadell",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di Banco Sabadell per il coordinamento dell'integrazione dei servizi di recupero crediti, presumibilmente a seguito di un'operazione di fusione/integrazione aziendale.",
  attivita: "Ufficio di coordinamento dell'integrazione dei servizi di recupero, elaborazione del piano di integrazione dettagliato con strumento di gestione online proprietario, definizione del percorso critico basato su attività chiave e interdipendenze, coordinamento globale del processo di integrazione (monitoraggio e reporting), supporto nella definizione ed esecuzione dei compiti chiave, coordinamento del processo di richiesta e convalida della documentazione, definizione del modello obiettivo del servizio post-integrazione.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa_9",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Gestione e ottimizzazione dei processi di recupero crediti (Spagna, Perù, Colombia)",
  clienteId: "cli_bbva",
  cliente: "",
  settore: "Banche & Assicurazioni",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: [],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Attività nell'ambito della practice Risk Management a supporto di BBVA per la revisione e riprogettazione del modello di recupero crediti, con l'obiettivo di migliorarne l'efficienza e il dinamismo operativo.",
  attivita: "Implementazione di un nuovo strumento di gestione del recupero, segmentazione del portafoglio di recupero e definizione delle strategie, implementazione di una soluzione aziendale multi-paese (Spagna, Perù e Colombia) per la gestione completa dell'Early Recovery, con strategie per segmenti di clientela omogenei e controllo centralizzato delle azioni di recupero.",
  risultati: "",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_10",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma Smart Destination per il turismo regionale",
  clienteId: "cli_pontevedra",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["MySQL", "Talend Open Studio", "Pentaho CE"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione, per la Diputación de Pontevedra (Consejería de Turismo), di un ambiente analitico a supporto dell'attività turistica regionale, capace di integrare più fonti dati e rispondere alle esigenze informative di gestori, operatori del settore e cittadini.",
  attivita: "Progettazione della piattaforma Smart Destination e del modulo di supporto all'Osservatorio del Turismo, con data mart e reportistica corporate per il monitoraggio degli indicatori di attività turistica e il confronto con i risultati medi di settore.",
  risultati: "Messa a disposizione di gestori, operatori turistici e cittadini di indicatori e report utili a valutare l'attività turistica della regione e orientare nuove azioni di sviluppo.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_11",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Unificazione dei sistemi informativi per la gestione degli aiuti ministeriali",
  clienteId: "cli_micinn",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle BI EE", "Oracle Warehouse Builder", "Oracle 11g"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "A seguito di successive fusioni e scorpori ministeriali, il Ministerio de Ciencia e Innovación si trovava con una mappa sistemi molto complessa per la gestione di aiuti e sovvenzioni, priva di una visione unica e consolidata del ciclo di business.",
  attivita: "Realizzazione di un cruscotto direzionale con vista unificata dei diversi aiuti gestiti dal ministero, con monitoraggio degli investimenti secondo assi temporale, geografico, tipologia di aiuto e finanziamento.",
  risultati: "Visione consistente nel tempo dell'informazione, verifica esaustiva della validità del dato e accesso semplificato ai dati per profilo utente.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_12",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema di consultazione e analisi dei risultati elettorali",
  clienteId: "cli_interior_es",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["MicroStrategy", "SPSS", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione, per il Ministerio del Interior, di un sistema per l'archiviazione e la pubblicazione delle consultazioni sui risultati elettorali svolte sul territorio nazionale spagnolo.",
  attivita: "Realizzazione di un data warehouse storico dei dati elettorali, di un ambiente pubblico di consultazione via internet e di un ambiente riservato di analisi con reportistica comparativa tra diverse tornate elettorali e integrazione di dati socio-demografici.",
  risultati: "Accesso ai risultati elettorali con il massimo dettaglio consentito dalla normativa, con strumenti di reportistica avanzata per gli utenti del ministero.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_13",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Caratterizzazione e prevenzione della criminalità sul territorio",
  clienteId: "cli_interior_es",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data", "AI"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "SPSS"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Progetto di modellizzazione dei reati per il Ministerio del Interior, finalizzato alla caratterizzazione criminologica del territorio e al supporto nella pianificazione delle forze di sicurezza.",
  attivita: "Definizione della metodologia di analisi, modellizzazione statistica fattoriale delle variabili sociodemografiche più rilevanti e sviluppo di modelli predittivi (regressione) per la prevenzione dei reati, con creazione di un cruscotto di indicatori di sicurezza interna.",
  risultati: "Disponibilità di un laboratorio dedicato all'analisi dei comportamenti delittuosi e di un cruscotto per il monitoraggio dei tassi di criminalità e delle azioni di contrasto.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_14",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma Big Data & Analytics per la gestione tributaria",
  clienteId: "cli_atc",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data", "AI"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle 12c", "Oracle BI EE 11", "MongoDB", "Hadoop", "Spark", "R"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Agència Tributària de Catalunya è stata realizzata una piattaforma Big Data integrata nel sistema informazionale dell'Agenzia, comprensiva di un modello predittivo per la rilevazione delle frodi nella valutazione di beni immobili, oltre al servizio di supporto e manutenzione del data warehouse.",
  attivita: "Sviluppo del sistema di indicatori per misurare l'efficienza dei circuiti tributari, realizzazione del cruscotto direzionale dell'Agenzia e abilitazione dell'accesso al modello analitico da parte degli utenti finali per soddisfare in autonomia le proprie necessità informative.",
  risultati: "Miglioramento dell'efficienza organizzativa, aumento del gettito grazie alla gestione del contrasto alle frodi, maggiore autonomia di accesso all'informazione e migliore monitoraggio del rispetto delle norme tributarie.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_15",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema informativo per la gestione universitaria",
  clienteId: "cli_diue",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle 10g", "SQL Server", "MicroStrategy 8.0.1", "Microsoft BI"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Departament d'Innovació, Universitats i Empresa è stato realizzato un sistema informativo sulle università, con analisi su titoli di studio, preiscrizioni, gestione accademica, borse di studio e informazioni economiche, oltre al relativo servizio di supporto e manutenzione.",
  attivita: "Gestione della qualità dell'informazione trasmessa dalle università tramite validazioni tecniche e funzionali, con supporto funzionale nella definizione di indicatori di qualità del sistema universitario, di finanziamento e di borse di studio.",
  risultati: "Disponibilità di un sistema affidabile e validato per il monitoraggio del sistema universitario regionale.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_16",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Modernizzazione dei sistemi informativi per la gestione accademica",
  clienteId: "cli_madrid_edu",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle 12c", "SAP Data Services", "SAP Business Objects"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per la Consejería de Educación, Juventud y Deporte della Comunità di Madrid è stato realizzato un modello informazionale per la gestione accademica, capace di integrare e unificare in un unico repository i dati di studenti, centri e corsi, eliminando duplicazioni tra le diverse entità.",
  attivita: "Analisi e definizione del BI corporate per tutte le aree di business dell'assessorato, generazione di un catalogo di report su misura per le diverse aree interne e realizzazione di un cruscotto di indicatori di gestione per il monitoraggio rispetto agli obiettivi fissati.",
  risultati: "Gestione più efficace delle risorse, controllo centralizzato della pianificazione contabile e finanziaria e maggiore autonomia degli utenti nelle interrogazioni ad-hoc.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_17",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema integrale di lotta al fraudolento fiscale",
  clienteId: "cli_gipuzkoa",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data", "AI"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["DB2", "IBM Cognos 8"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Agencia Tributaria de Guipúzcoa è stato progettato e implementato un sistema integrale di lotta alla frode fiscale, con l'obiettivo di individuare i soggetti a rischio e ottimizzare i processi di ispezione e sanzione.",
  attivita: "Ridisegno e semplificazione dei processi di ispezione e sanzione, realizzazione di un data mart per l'ispezione e di un sistema transazionale per l'approvazione del piano ispettivo, implementazione di un sistema di selezione dei contribuenti ad alto rischio di frode fiscale e di un cruscotto per il monitoraggio degli accertamenti.",
  risultati: "Disponibilità di un sistema integrato per la pianificazione, selezione e monitoraggio delle attività di contrasto alla frode fiscale. Durata del progetto: 2 anni.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_18",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma analitica per la gestione tributaria",
  clienteId: "cli_icm",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Business Objects XI - Xcelsius", "Oracle 10g"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione, per la Consejería de Hacienda della Comunità di Madrid, di un ambiente analitico a supporto della gestione tributaria regionale.",
  attivita: "Sviluppo di un ambiente di reporting per l'analisi e il monitoraggio di report statistici gestionali, generazione autonoma di nuovi report da parte degli utenti finali e realizzazione di un ambiente analitico per scenari di simulazione, oltre a un cruscotto di indicatori di gestione.",
  risultati: "Maggiore autonomia degli utenti nella produzione di report e disponibilità di un cruscotto direzionale per il monitoraggio degli obiettivi.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_19",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema integrale di gestione tributaria (M@GIN)",
  clienteId: "cli_canarias_trib",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle 10g", "Oracle Warehouse Builder 10g", "Oracle BI EE"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Agencia Tributaria de Canarias è stato progettato e implementato il sistema M@GIN, per la gestione integrale dei tributi propri e ceduti dell'amministrazione tributaria canaria, nell'ambito di un più ampio piano di modernizzazione.",
  attivita: "Consulenza sul piano di modernizzazione della gestione tributaria, ridisegno e ottimizzazione dei processi di business, analisi dei requisiti di adattamento della soluzione del Governo di Catalogna (G@UDI) e supporto alla gestione del cambiamento e alla formazione.",
  risultati: "Realizzazione di un sistema transazionale integrale per la gestione tributaria e di un data warehouse per reportistica e cruscotto di gestione, con monitoraggio delle pratiche e analisi ad-hoc dell'informazione.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_20",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Reingegnerizzazione dei processi e trasformazione tecnologica previdenziale",
  clienteId: "cli_onp_peru",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Perù",
  inizioMese: 1,
  inizioAnno: 2013,
  fineMese: 12,
  fineAnno: 2015,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "InfoSphere DataStage", "QualityStage", "Cognos"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Oficina de Normalización Previsional (ONP) del Perù è stato realizzato un progetto di reingegnerizzazione dei processi di business e trasformazione tecnologica, con l'obiettivo di ridurre drasticamente i tempi di riconoscimento delle prestazioni previdenziali e il contenzioso giudiziale.",
  attivita: "Progettazione e implementazione di una base dati istituzionale degli iscritti tramite depurazione, validazione e integrazione con basi dati interne ed esterne, riconversione tecnologica dei sistemi e supporto specialistico sui processi in ambito previdenziale, strutturato in una fase di consulenza (reingegnerizzazione) e una fase operativa (implementazione).",
  risultati: "Ottimizzazione dei processi di business e miglioramento della qualità percepita dall'utente finale, con una fonte dati unica e ufficiale per i processi primari dell'ente. Progetto pianificato su 3 anni (2013-2015).",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_21",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema informazionale per il monitoraggio della gestione giudiziaria",
  clienteId: "cli_justizia_eus",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle Business Intelligence 11g", "Oracle Warehouse Builder 10g", "Oracle 10g"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione, per l'Administración de Justicia de Euskadi, di un nuovo sistema informazionale per il monitoraggio della gestione giudiziaria.",
  attivita: "Realizzazione di un cruscotto di indicatori chiave di gestione a supporto delle decisioni, di un ambiente di reporting per la generazione autonoma di report da parte degli utenti finali e automazione dei caricamenti informativi verso il nuovo data warehouse.",
  risultati: "Ampliamento del numero di indicatori chiave disponibili per conoscere rapidamente lo stato di ciascun ufficio giudiziario.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_22",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Portale di statistiche sull'attività giudiziaria",
  clienteId: "cli_judicatura_ec",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Ecuador",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Data warehouse", "Cuadro de Mando", "Modello multidimensionale", "Excel"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Consejo de la Judicatura dell'Ecuador è stato realizzato un ambiente informazionale per la valutazione delle prestazioni degli organi giudiziari, integrando diversi strumenti di informazione e supporto decisionale per i vari attori dell'amministrazione della giustizia.",
  attivita: "Realizzazione di un cruscotto con KPI per il monitoraggio degli obiettivi da parte dei responsabili di governo della giustizia, copertura di diverse aree tematiche dell'attività giudiziaria (domanda, pianificazione risorse, produttività, qualità), report per il monitoraggio operativo di ciascun organo giudiziario e pubblicazione periodica di statistiche regolatorie.",
  risultati: "Soluzione flessibile, adattabile e modulare basata su tecnologia Microsoft, distribuita ai tribunali anche tramite Excel per un basso costo totale di possesso.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_23",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema di informazione, controllo e consultazione per la Procura Generale dello Stato",
  clienteId: "cli_justicia_es",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["IBM Software Development Platform", "Unix/Shell-script", "Oracle 10g", "Oracle Business Intelligence 10g", "Oracle Warehouse Builder 11g"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Ministerio de Justicia (Fiscalía General del Estado) è stato realizzato il Sistema de Información Control y Consulta (SICC), con integrazione dell'informazione processuale delle procure di tutto il territorio nazionale.",
  attivita: "Omogeneizzazione dell'informazione ricevuta tramite il test di compatibilità del Consiglio Generale del Potere Giudiziario, realizzazione di un'interfaccia web per l'accesso online alle informazioni caricate nel sistema e sviluppo di un ambiente analitico per la valutazione delle prestazioni di ciascuna procura.",
  risultati: "Disponibilità di indicatori omogenei per valutare le prestazioni delle procure su tutto il territorio nazionale.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte salvo diversa indicazione; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_24",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Data Warehouse per il monitoraggio statistico ministeriale",
  clienteId: "cli_agricultura",
  cliente: "",
  settore: "Pubblica Amministrazione",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle Data Integrator", "Oracle Database EE 11g", "Oracle Business Intelligence EE 10g"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione di un data warehouse a supporto della Rete Nazionale di Informazione dell'ente, con l'obiettivo di centralizzare i dati e fornirne l'accesso alle agenzie e agli uffici regionali collegati.",
  attivita: "Messa a disposizione degli utenti dell'Ufficio Statistiche Agricole di indicatori su produzione, mercato, colture e pesca, oltre a informazioni su operazioni di campo e pianificazione delle politiche di settore.",
  risultati: "Centralizzazione dei dati statistici agricoli in un'unica infrastruttura informativa condivisa.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Denominazione esatta dell'ente e paese non specificati con certezza nel documento originale (contenuto generico, non necessariamente riferito alla Spagna). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_25",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Cruscotto di Bilanciamento Integrato per la gestione degli obiettivi aeroportuali",
  clienteId: "cli_aena",
  cliente: "",
  settore: "Trasporti",
  paese: "Spagna",
  inizioMese: 2,
  inizioAnno: 2009,
  fineMese: 1,
  fineAnno: 2013,
  inCorso: false,
  importoKEuro: 300,
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "MicroStrategy (SDK)", "MicroStrategy (mobile)"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Manutenzione ed evoluzione, per AENA, dell'applicazione di Cuadro de Mando Integral che consente agli utenti di valutare gli elementi del piano operativo di ciascuna unità (aeroporto o direzione), strutturato per assi, obiettivi, indicatori e programmi.",
  attivita: "Coordinamento dei team MicroStrategy e monitoraggio continuativo con gli utenti AENA, accesso ai cruscotti tramite icone applicative adattate alla struttura organizzativa aziendale, integrazione di dati non presenti nelle fonti corporate e moduli di amministrazione per la gestione e parametrizzazione dei piani operativi.",
  risultati: "Cruscotti di bilanciamento integrato pienamente operativi su tutta l'organizzazione AENA, con accesso ai report di dettaglio direttamente dal cruscotto.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Periodo e importo economico (300 k€) riportati come indicati esplicitamente nel documento originale.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_26",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Supporto ai sistemi informazionali corporate",
  clienteId: "cli_correos",
  cliente: "",
  settore: "Trasporti",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "SAP Business Objects", "JasperReports", "Java"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per Correos è stato erogato un servizio continuativo di supporto ai sistemi informazionali corporate, con l'obiettivo di rendere disponibili le informazioni aziendali ai diversi utenti di business.",
  attivita: "Gestione di data mart tematici in base alle esigenze di analisi degli utenti, realizzazione di cruscotti direzionali e di report consolidati per gli organi di gestione intermedi, generazione di pannelli informativi con JasperReport e sviluppi su misura per minimizzare l'impatto sulle licenze software, oltre a formazione e supporto agli utenti finali.",
  risultati: "Disponibilità continuativa di informazione corporate affidabile per il business, con riduzione dei costi di licenza grazie a sviluppi su misura.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_27",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Repository informativo del Sistema Sanitario Nazionale",
  clienteId: "cli_sanidad_es",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Business Objects", "Microsoft SQL Server", ".NET"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Ministerio de Sanidad, Servicios Sociales e Igualdad è stato realizzato un sistema che integra diversi applicativi analitici con informazioni sanitarie gestite da distinte unità organizzative, reso disponibile anche tramite portale statistico ad accesso libero per il cittadino.",
  attivita: "Consolidamento e integrazione dell'evoluzione temporale dei principali indicatori sanitari per la rilevazione e pianificazione delle tendenze di politica sanitaria, messa a disposizione dei cittadini di strutture multidimensionali per la generazione rapida di report dinamici.",
  risultati: "Accesso pubblico e trasparente agli indicatori del Sistema Sanitario Nazionale spagnolo.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_28",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma di gestione della conoscenza sulla violenza di genere",
  clienteId: "cli_sanidad_es",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Business Objects"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Ministerio de Sanidad, Servicios Sociales e Igualdad è stata realizzata una piattaforma a supporto della Delegazione del Governo per la Violenza di Genere, per il monitoraggio, la valutazione, la comunicazione e l'elaborazione di proposte di intervento in materia.",
  attivita: "Analisi, sfruttamento e diffusione dell'informazione per l'insieme delle istituzioni pubbliche e private impegnate nella prevenzione e nel rilevamento precoce lungo l'intero ciclo del fenomeno, con necessità di conoscenza e azione coordinata.",
  risultati: "Supporto informativo coordinato all'azione istituzionale di contrasto alla violenza di genere.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_29",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Cruscotti direzionali dipartimentali",
  clienteId: "cli_quironsalud",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["SAP Business Objects 4.0", "SQL Server", "Integration Services"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Progettazione, realizzazione e implementazione, per il Grupo Quirónsalud, di quattro cruscotti dipartimentali relativi alle aree di portale del paziente, centrale acquisti, risorse umane e controllo di gestione.",
  attivita: "Sviluppo del progetto end-to-end, dalla raccolta dei requisiti fino all'implementazione in produzione, con validazione da parte degli stakeholder di ciascun cruscotto; il cruscotto Risorse Umane è stato il primo a consolidare le informazioni corporate dei due grandi gruppi fusi nel 2014.",
  risultati: "Cruscotti realizzati sulla piattaforma BI corporate (SAP BO 4.0), con dashboard aggregati, report di dettaglio e possibilità di query libera.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_30",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Cruscotto di Bilanciamento Integrato di gruppo",
  clienteId: "cli_quironsalud",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["SAP Business Objects 4.0", "SQL Server", "Integration Services", "SAP Dashboards", "Cmaps"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Progettazione, realizzazione e implementazione del Cuadro de Mando Integral del Grupo Quirónsalud, con sfruttamento congiunto a livello corporate delle informazioni delle principali aree di business del gruppo.",
  attivita: "Realizzazione di filtri rapidi dell'informazione, benchmarking tra centri e analisi del contributo di ciascuna struttura al risultato complessivo, con geolocalizzazione degli indicatori principali tramite il plug-in Cmaps nei dashboard.",
  risultati: "Nuovo punto di accesso comune con visione integralmente orientata al business, navigazione intuitiva e piena fruibilità da dispositivi mobili.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_31",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Modulo di analisi e stratificazione dei pazienti",
  clienteId: "cli_sergas",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data", "AI"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle 11g", "Talend Open Studio", "Pentaho CE", "R"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Servizo Galego de Saúde è stato realizzato uno strumento di classificazione della popolazione basato su un modello piramidale di Kaiser adattato alle caratteristiche della popolazione galiziana, per migliorare la gestione dei pazienti sotto il profilo clinico e strategico.",
  attivita: "Segmentazione della popolazione secondo il modello di Kaiser, realizzazione di un repository centrale che unisce dati del paziente, risultati della segmentazione e probabilità di evoluzione tra segmenti per l'esecuzione di studi analitici.",
  risultati: "Migliore controllo e previsione dell'evoluzione dei pazienti cronici e supporto alla definizione di politiche per un uso più efficiente delle risorse sanitarie.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_33",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Cruscotto direzionale per la gestione unificata dell'attività sanitaria",
  clienteId: "cli_sanidad_madrid",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Business Objects", "Xcelsius", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Realizzazione, per la Consejería de Sanidad della Comunità di Madrid, di cruscotti per la misurazione dell'attività dei centri di assistenza sanitaria primaria e ospedaliera della regione.",
  attivita: "Misurazione dell'attività di prenotazione in assistenza primaria e specialistica, con particolare attenzione ai processi di invio da primaria a specialistica e alla libera scelta del medico; misurazione dell'attività generata dalla libera scelta dello specialista negli ospedali, con mappe di migrazione dei pazienti e quantificazione dell'impatto economico tra ospedali.",
  risultati: "Visione direzionale unificata dell'attività sanitaria regionale, a supporto della redistribuzione dei costi tra strutture.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_34",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema di gestione dell'attività ambulatoriale regionale",
  clienteId: "cli_avs",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle BI", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Agencia Valenciana de Salud è stato realizzato un repository informativo associato alla gestione ambulatoriale regionale, per il monitoraggio dei dati clinici e di attività dell'assistenza primaria.",
  attivita: "Consolidamento dell'informazione assistenziale con quella del paziente, consultazione online di grandi volumi di dati con controlli di accesso per profilo e certificato digitale, produzione di report esecutivi e analitici su indicatori quali numero di visite, appuntamenti, tassi di invio a specialistica e prevalenza delle patologie.",
  risultati: "Miglioramento continuo delle prestazioni sanitarie in assistenza primaria grazie al monitoraggio puntuale dell'attività clinica.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_35",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Repository informativo sulla spesa farmaceutica e classificazione dei pazienti",
  clienteId: "cli_avs",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data", "AI"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle BI", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Agencia Valenciana de Salud è stato realizzato un sistema di classificazione dei pazienti per l'analisi della spesa farmaceutica e del consumo di risorse sanitarie in relazione al carico di malattia e al rischio clinico dei cittadini valenziani.",
  attivita: "Identificazione e classificazione dei pazienti cronici che generano la maggiore deviazione economica ed effetti avversi, tramite modelli CRG (Clinical Risk Groups), sistema di allerta su gravità e deviazioni economiche e valutazione della spesa farmaceutica per patologia e gruppo terapeutico.",
  risultati: "Supporto alla decisione clinica e alla previsione del consumo di risorse sanitarie e farmaci per paziente o gruppo di pazienti.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_36",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sistema di supporto alla decisione clinica (Diraya)",
  clienteId: "cli_sas",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["MicroStrategy", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Diraya è il sistema informativo assistenziale del Servicio Andaluz de Salud, che integra in un fascicolo sanitario unico tutti gli episodi assistenziali del cittadino registrati da professionisti di assistenza primaria, specialistica e ospedaliera.",
  attivita: "Gestione delle agende di prenotazione per assistenza primaria, specialistica e prove diagnostiche, monitoraggio dei trattamenti indicati ed estensione della ricetta elettronica, con un sistema completo di supporto alla decisione clinica integrato nell'infrastruttura.",
  risultati: "Ottimizzazione della relazione tra cittadini e professionisti del servizio sanitario andaluso.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_37",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Data mart per il monitoraggio dei pronto soccorso",
  clienteId: "cli_sas",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["MicroStrategy", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il Servicio Andaluz de Salud è stato realizzato un sistema specializzato per l'analisi dei fattori interni ed esterni che incidono sul funzionamento dei servizi di pronto soccorso.",
  attivita: "Monitoraggio degli indicatori sui tempi di risposta assistenziale lungo l'intero percorso del paziente, sulla dinamica dei flussi assistenziali e sulla mortalità, a supporto del miglioramento continuo della qualità assistenziale.",
  risultati: "Migliore conoscenza dei fattori che incidono sulla funzionalità dei servizi di urgenza ed emergenza.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_38",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Modello informazionale corporate per la gestione del call center sanitario",
  clienteId: "cli_saludresponde",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "Unix PL/SQL", "MicroStrategy Mobile Suite"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per il servizio Salud Responde della Consejería de Andalucía è stato realizzato un modello informazionale per misurare l'attività dei diversi servizi erogati, sia telefonici sia su altri canali, e i relativi tempi di gestione.",
  attivita: "Pianificazione della domanda per ciascun servizio su base oraria, settimanale, annuale e non periodica, valutazione della qualità dell'attenzione ricevuta e gestione preventiva dell'attività del call center, con fruizione analitica ed esecutiva disponibile anche su iPad.",
  risultati: "Gestione preventiva e monitorata dell'attività del call center sanitario regionale, con reportistica fruibile in mobilità.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_39",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Piattaforma per la pianificazione sanitaria regionale",
  clienteId: "cli_sanidad_cyl",
  cliente: "",
  settore: "Sanità",
  paese: "Spagna",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle BI", "Oracle"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Progettazione, sviluppo e implementazione, per la Consejería de Sanidad di Castiglia e León, di un data mart sanitario a supporto della Direzione Generale di Pianificazione Sanitaria, orientato all'analisi della copertura e distribuzione della popolazione rispetto a risorse e servizi.",
  attivita: "Realizzazione di un cruscotto a supporto della pianificazione e riorganizzazione sanitaria, con analisi dell'informazione sulla popolazione rispetto a professionisti, offerta di servizi e zonizzazione sanitaria, orientato a dirigenti e responsabili di servizio con diversi livelli di accesso analitico.",
  risultati: "Disponibilità di uno strumento di supporto alla pianificazione e riorganizzazione dei servizi sanitari regionali.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}, {
  id: "spa2_40",
  fonte: "file",
  creato: new Date().toISOString(),
  titolo: "Sfruttamento dei sistemi informativi clinici",
  clienteId: "cli_hlf_chile",
  cliente: "",
  settore: "Sanità",
  paese: "Cile",
  inizioMese: 1,
  inizioAnno: 2016,
  fineMese: 12,
  fineAnno: 2018,
  inCorso: false,
  importoKEuro: "",
  ambiti: ["Data"],
  teamSize: "",
  ruolo: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  tecnologie: ["Oracle", "Oracle BI", "WebLogic"],
  tipoReferenza: "standalone",
  parentId: null,
  descrizione: "Per l'Hospital de la Florida di Santiago del Cile è stato realizzato un modello centralizzato di informazione a supporto delle decisioni, organizzando i dati disponibili secondo standard e classificandoli per area di business.",
  attivita: "Monitoraggio dei piani di intervento con confronto tra risultati ottenuti e attesi, individuazione di nuove tendenze e proiezioni future, identificazione di aree di rischio e risultati indesiderati, produzione periodica di statistiche globali e specifiche.",
  risultati: "Miglioramento del monitoraggio della qualità dei processi assistenziali e maggiore capacità di individuazione precoce di criticità.",
  codiciProgetto: [],
  noteAggiuntive: "Referenza tradotta e importata da presentazione aziendale (deck referenze Spagna, gennaio 2019). Il documento originale non riporta importo economico né date esatte; il periodo indicato è stimato in base alla data del documento.",
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  nda: false
}];
const emptyClient = () => ({
  id: "cli_" + Date.now() + "_" + Math.floor(Math.random() * 9999),
  nome: "",
  descrizione: "",
  labelBreve: "",
  nda: false
});
const emptyRef = () => ({
  id: "ref_" + Date.now() + "_" + Math.floor(Math.random() * 9999),
  titolo: "",
  clienteId: null,
  cliente: "",
  settore: "Altro",
  paese: "Italia",
  inizioMese: new Date().getMonth() + 1,
  inizioAnno: new Date().getFullYear(),
  fineMese: "",
  fineAnno: "",
  inCorso: true,
  importoKEuro: "",
  ambiti: [],
  tecnologie: [],
  descrizione: "",
  attivita: "",
  risultati: "",
  ruolo: "",
  teamSize: "",
  modalitaFornitura: "unico",
  quotaPercentuale: "",
  partnerRTI: "",
  attivitaAutonoma: false,
  contestoOrganizzativo: "",
  metodologia: "",
  kpiDettagliati: "",
  elementiUnicita: "",
  referenteContatto: "",
  noteAggiuntive: "",
  codiciProgetto: [],
  tipoReferenza: "standalone",
  parentId: null,
  nda: false,
  fonte: "manuale",
  creato: new Date().toISOString()
});

// ── AI: struttura testo libero in referenza JSON ──────────────
async function aiStructure(rawText, extraContent) {
  const sys = `Sei un assistente per una practice di consulenza IT (Data, AI, RPA).
Ricevi la descrizione di una referenza progettuale (da testo, dettato vocale o documento) e la trasformi in JSON.
Rispondi SOLO con un JSON valido, senza backtick, senza testo prima o dopo. Schema:
{
 "titolo": "titolo sintetico della referenza (max 10 parole)",
 "cliente": "nome del cliente così come citato (o 'Cliente riservato' se assente)",
 "settore": "uno tra: ${SETTORI.join(" | ")}",
 "paese": "Paese in cui si è svolta la referenza (es. Italia, Spagna, Francia...); se non specificato usa 'Italia'",
 "inizioMese": "numero 1-12 del mese di inizio progetto",
 "inizioAnno": 2024,
 "fineMese": "numero 1-12 del mese di fine, null se il progetto è ancora in corso",
 "fineAnno": "anno di fine, null se il progetto è ancora in corso",
 "inCorso": "true se il testo indica che il progetto è ancora attivo/in corso, altrimenti false",
 "importoKEuro": 350,
 "ambiti": ["uno o più tra: ${AMBITI.join(", ")}"],
 "tecnologie": ["lista tecnologie citate"],
 "descrizione": "contesto e obiettivi, 2-4 frasi in italiano formale da offerta di gara, citando il nome del cliente",
 "attivita": "attività svolte, 2-4 frasi in italiano formale, citando il nome del cliente",
 "risultati": "benefici e risultati misurabili, 1-3 frasi",
 "ruolo": "ruolo del fornitore (es. system integrator, capofila RTI...)",
 "teamSize": 8,
 "modalitaFornitura": "uno tra: unico | rti_mandataria | rti_mandante | subappalto (deduci da termini come 'da soli', 'RTI', 'capofila', 'mandante', 'subappalto'; se non specificato usa 'unico')",
 "quotaPercentuale": "quota percentuale di partecipazione se in RTI/subappalto, altrimenti null",
 "partnerRTI": "nome degli altri partner del RTI se menzionati, altrimenti stringa vuota",
 "attivitaAutonoma": "true se il testo indica che, pur essendo in RTI, l'attività descritta è stata svolta interamente in autonomia dal nostro team; altrimenti false",
 "contestoOrganizzativo": "se il testo fornisce dettagli sul contesto organizzativo, le criticità o i vincoli di partenza del cliente oltre a quanto già in 'descrizione', riportali qui in modo esteso; altrimenti stringa vuota",
 "metodologia": "metodologia, framework o standard di lavoro adottati (es. Agile/Scrum, DevOps, certificazioni, normative di riferimento) se menzionati; altrimenti stringa vuota",
 "kpiDettagliati": "eventuali metriche o KPI aggiuntivi oltre a quelli già in 'risultati', se presenti nel testo; altrimenti stringa vuota",
 "elementiUnicita": "eventuali elementi di unicità, innovazione o valore differenziante del progetto, se menzionati; altrimenti stringa vuota",
 "referenteContatto": "nome, ruolo e contatti di un referente del cliente citato nel testo come verificabile per la referenza, se presente; altrimenti stringa vuota",
 "noteAggiuntive": "qualunque altra informazione utile fornita nel testo che non rientri nei campi precedenti; altrimenti stringa vuota",
 "codiciProgetto": "array di codici progetto, commessa, contratto o ordine citati nel testo (es. codici interni, CIG, numero contratto); array vuoto se non presenti"
}
Se un dato numerico non è presente usa null. Scrivi tutto in italiano professionale adatto a una risposta di gara.`;
  const content = extraContent ? [extraContent, {
    type: "text",
    text: "Estrai la referenza progettuale da questo documento." + (rawText ? " Note aggiuntive: " + rawText : "")
  }] : "Testo da strutturare:\n" + rawText;

  // ATTENZIONE: questa chiamata diretta a api.anthropic.com funzionava solo dentro la sandbox
  // degli artifact di Claude.ai, che inietta automaticamente l'autenticazione.
  // Fuori da lì (questo repository), fallirà con 401/CORS finché non si predispone un piccolo
  // backend/proxy proprio (es. Cloud Function) che aggiunga una vera chiave API Anthropic
  // lato server — non va MAI inserita una chiave API nel codice client. Vedi README.
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: sys,
      messages: [{
        role: "user",
        content
      }]
    })
  });
  const data = await resp.json();
  const text = (data.content || []).map(b => b.text || "").join("\n");
  const clean = text.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

// ── AI: assistente per ricerche libere, analisi e confezionamento referenze ──
async function aiAssistente(promptUtente, archivio) {
  const sys = `Sei un assistente esperto nella redazione di risposte a gare d'appalto e procedure di qualifica per una practice di consulenza IT specializzata in Data, AI e RPA.
Ricevi un archivio di referenze progettuali in formato JSON e una richiesta libera dell'utente in linguaggio naturale. La richiesta può chiedere:
- ricerche, valutazioni, analisi o confronti tra le referenze (es. copertura per un requisito di gara, punti di forza/debolezza, referenze più adatte);
- di confezionare una o più referenze secondo un template specifico di gara o qualifica, in formato sintetico o super dettagliato, nella lingua eventualmente richiesta.
Regole obbligatorie:
- Usa esclusivamente i dati presenti nell'archivio fornito. Se mancano informazioni per rispondere con precisione, dichiaralo esplicitamente invece di inventare dati, importi, o dettagli.
- Alcune referenze sono anonimizzate per NDA (campo "nda": true): in questo caso il nome del cliente è già stato sostituito con un'etichetta generica nei dati che ricevi. Non tentare mai di indovinare o dedurre il nome reale del cliente, e mantieni l'etichetta anonima in qualunque output tu produca per quelle referenze.
- Il campo "tipoReferenza" indica la natura della referenza: "programma" (iniziativa ampia e pluriennale, contenitore di più progetti), "servizio_ams" (servizio continuativo di assistenza/manutenzione applicativa, anch'esso contenitore di eventuali progetti evolutivi), "progetto" (un'attività puntuale ed evolutiva attivata nell'ambito di un Programma o di un Servizio AMS, indicato dal campo "programmaOServizioDiRiferimento") oppure "standalone" (referenza singola indipendente, non collegata a un contenitore). Quando la richiesta riguarda un Programma o un Servizio AMS, considera sia la referenza di livello superiore sia i progetti collegati, distinguendo chiaramente tra il perimetro/valore complessivo dell'iniziativa e i risultati concreti dei singoli progetti attivati.
- Rispondi nella lingua in cui è scritta la richiesta dell'utente, a meno che l'utente non chieda esplicitamente una lingua diversa nel testo del prompt.
- Formatta la risposta in testo semplice leggibile: usa intestazioni con "## ", elenchi puntati con "- " e grassetto con "**testo**" dove utile alla chiarezza. Non usare tabelle markdown.
- Sii concreto e concentrato sulla richiesta specifica; evita premesse superflue.`;
  const userMsg = `ARCHIVIO REFERENZE (JSON, ${archivio.length} referenze):\n${JSON.stringify(archivio)}\n\nRICHIESTA DELL'UTENTE:\n${promptUtente}`;

  // ATTENZIONE: vedi nota identica in aiStructure() sopra — stessa limitazione, stessa soluzione.
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system: sys,
      messages: [{
        role: "user",
        content: userMsg
      }]
    })
  });
  const data = await resp.json();
  return (data.content || []).map(b => b.text || "").join("\n").trim();
}

// ── util ──────────────────────────────────────────────────────
const fmtImporto = k => k === "" || k === null || k === undefined ? "n.d." : k >= 1000 ? (k / 1000).toLocaleString("it-IT", {
  maximumFractionDigits: 1
}) + " M€" : Number(k).toLocaleString("it-IT") + " k€";
const readFileB64 = file => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(r.result.split(",")[1]);
  r.onerror = () => rej(new Error("Lettura file non riuscita"));
  r.readAsDataURL(file);
});

// ridimensiona e comprime un'immagine lato client prima di salvarla nello storage
const comprimiImmagine = (file, maxW = 1100, quality = 0.78) => new Promise((res, rej) => {
  const img = new Image();
  const url = URL.createObjectURL(file);
  img.onload = () => {
    const scale = Math.min(1, maxW / img.width);
    const w = Math.round(img.width * scale);
    const h = Math.round(img.height * scale);
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    canvas.getContext("2d").drawImage(img, 0, 0, w, h);
    URL.revokeObjectURL(url);
    res(canvas.toDataURL("image/jpeg", quality));
  };
  img.onerror = () => {
    URL.revokeObjectURL(url);
    rej(new Error("Immagine non leggibile"));
  };
  img.src = url;
});
const imgKey = refId => "immagini:" + refId;
const logoKey = clientId => "logo:" + clientId;

// loghi cliente estratti dal documento originale (sfondo reso trasparente lato server)
const LOGHI_SEED = {
  cli_interior_es: "data:image/png;base64,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",
  cli_onp_peru: "data:image/png;base64,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",
  cli_madrid_edu: "data:image/png;base64,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",
  cli_justizia_eus: "data:image/png;base64,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",
  cli_sanidad_cyl: "data:image/png;base64,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",
  cli_sergas: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHUAAAA0CAYAAAC0AHd1AAA2iUlEQVR42uW8eZhU1Z3//zrnLlXV3dUr3Q0NNNKALM2ObIqCAiKIigsYNZqoETKJu8lEJ2aSyZiMJmrUaNw33MWFqGziggoBRGRr9qWhGxrofamu5S7nfP+41U1jTGZ+v+88M7/nl8tznrpU3Xvu7fM5n+39WYSjNEKA1CAECEBrEBroOBcEP4jgO4VGkL4AiUYjtUApF6U8pGmQSDphKxRJxmLJ8XnZmevx4eGHn9Q7d+7mjp/9M336dxfrvtqq33pzEVdeeSVjRg4Wrg9hAxxXIaWk42kd50JovGQKhEc4FEYrBcIAOq4FIQR//1Df+L/8G9//Z4dMr0THvebfeVbXuc1vvUKmz/7q7XX6/dI/6PR1CgVIjPT7665P+LY1EOLkq7RWiDQZhRBdCAq+8jGkgVYKaRi4rouQgnAkK6k05GRnrq851vDmnx55bO7+fZX0Ku1DXl7ezKXLP9X3/f53nD5hIg888AATx4/TzY31XD5vLoMHDRCOq7AsieP46XfQgML3/ayMjEgMNELKk3ab1jp93X+FuP8dh0Sg0Eg6luzkz+D3kzfPt2+pb31bLf/mHd/8pSvJhK/TO0R/4yLdhcJao0WwPzoWSwMaHewxpTCFge8FXIoQ7Nq/X6eSLo2NTfzx4UcYP2YcP/rRj0p+9/sHag4freFoXS0//dk/c8YZp4klSz7SS95/D1OCbZn84Q8PCiP91p4fEMv3XYTQZNih9P5WKN9HGtZJf2IHUUHw9+n69xf7//Hx1xTt8qn+BoH+i3PCN7j1W+jUdWcopf7u0EqnOcA/afjawdcOSvsorfE8D6U0yZRPQ0ts/LOvvK7ffG+JvviK7+qnX3hZK61xUpoHH/yTnnXBPL1u007drjTtSpPSmrir2X2gSp89/Ty9/KNP9Rdrv9Q1tQ0POUrjak0s6eIoTTyZQGsf5bvpd9H/L4f/f3l/eqj/gfF3n++nx4nrpUD8bVElBEp7aO2ftMM1PiItfHR630hpgAbTluzes3/dwUPVeL5m1uzzufLqq0TKhU1bNmshNWPHjqF/2SkjDAHvvv2evmH+jXr58uW6qKgo//bbb+eRRx7h1ltv5amnnrqlPZ7KqzlauzaeSFwmBZimie/7adF7ssj9Xzk67Q31XxsowAuGUCBA/yfjhGTpOjzQ3on/a9UpfSRag9InsXHnhGi0kOnzrnpEIDjZQNFpSfbhytX6RzfeysBBQ9FCsnffAWpr69ZaFpSV9V0QjWaxbt1feOedd7YIwLIsysrKaG1txXXdyTPPmy5efvnl8uXLl1989NgxHnzwwcYvvlg9obW1bVHKVSQTTqFC4HveyZLqbxD27xNc/rdJ3m9f+K7Pkd/YBPo/Mc7U3x5afUMGnyymJT4oBVqLzqF8UD54SqGFQCHxESgESgXX+77G932UUvj4aBE88tPPv6DPKX05Xl/PpClnX5eZHWXHrh0TfCCvIOep6667Rrz/wWLR2FjP2+98oC++aKaItTZTXl5OVlbWYgDLtHfk5kYX/+u//vJhKU1y8vN45plnWLNmjQ6Hw3WGYWCYJlophNCI9I7UWv+V+gDBt9P1v4+gqtPGkOnPrucnvvv2ZyqE7kLrziHT4xsP0/LEIGA4H9BdpIH85iL4SqG1xtMKpUgvisD3NY6ncD2F62ocxyPpuCgBSiscV+EpuO3W26+8+19/yeYt21j+4crniop7YNpW4ABohe8Hf94111xzz8qVK1EKRo8ejWEYmCZ4HoQjJkrDsbraWxKpFN2LSzBNm7feeptYIjlEpN9JSPlXXPlNwv5PHQr5N8YJfjxBZPnXVo/uMr5180kQsgvXyzRPSnzoHIEqtAxM00DKYGdoLTBNSSrl5jU0Nr+o0n5qyvFwHM+UlokdsrBDIUJ2BI1EaYE0JO0JP4w06sr6l4qevfvQ0NAQLHLasG9PJHscOXb0SDyezIvH43ef0rcUpUDKwIL2fXDSYrWq6qh+7tkXuOyyeQwbNlAcrqlh5syZZEczdsTjyROLmSacSIszwzDS1rKPYRi4bmDAeel5O8TxCYKfEGu+7550TccG6bje9/3Oe5Ty0t8Hy9shTWXnv+A7Cei0NOywVrWSoE20nyaY9ukqTjqu9RUoHWwEv2OIgEkV4OpgMwVYgehED6Tj+AGX+RphBH+I48P2HTsaX3xx4TUVFTu10pARsbFDIS/laNoTiu079+mk4+Mj0NJg+/ad+vXXX0888cQTKy+77Lt606ZNTJlyDkopLCsEwMaNG2u++93vlrzw4ouNnlbYto1hQjKZREqJ7yvCYZPd+w/qe3/3O34w/wYGDDxVXDjncl3WfwCzzp8hkk5wX4dxJ6XsJELHuRAiPZ+PEALDEJ0Axt/TsYZhnOTjChEYkVprpJTpOdRJm8KQaZ7TJ86Vr0EFzOC7YBhgiICgvg8yzaSdkkaKTuRH6YAhlQYpA9sm6Tq4npsmtMZLe0hSggEYSEwMDAQCMIU0Og0hx1GEQpKWtmTvTZs2M2jQEH53/4OcffbZevr06Qd69SzsZwjB9q179Jo1a9i9f5+uqz/G4ZpqKvcfomdJHy6+ZC57Kw8w+8ILGDqoVHz91Xodj8dRQHY0lwsvupiCggKaWpppa48FrphpoARIS7Jj1wH9zHPPcv0N8ynrXyZWfbpGF3UvZtCgQTQ1x+fk52YsNtPEEWh8FXBnB2cqpfA8D9M0SaVSWZmZmbH29oQZDoe9DsK7rotlWX9F1GAOeRLHSHkC2BBdUBnTNAP0TANS0qkJtMSQJzShaQYbQCCC+U0z4Eqt03owbaug8L3gGTYygAfSlkzY6pjPTb+PDIiORPhdDSUJ+JiO5yJNk7AtqDpct3/FihVl6zd8iWWFGDRkKNnZ2TzxxBNs3VpRZoVsHYlksm/fPq677jo++eQjGhqP094eY8ENCwiFs7Ask7v+5We8uWgR5848m3g8jmWGkEBZWb8JS5YsXVdQVET/AQNKho8YUfPp56t1a2srjuPw3Asv6j2793Hld6+ivHywQELKc5l94Ry09onFYk/mZmcs1gK0UhiG6BS3HVwqpSQUCpFKpcjMzIz5viYUCnlSBnZB13tOYCs6zdFWWi8LuqrrDtsCNEpppCQtDTqUoQs67Q34PloE+j7wIz2kITq5OnBnJGgQwsTTLggLIcAyOzaXxjQEWvsIAUp5GNJIq97g/QxhnayP1Qk3VIYiFp7yUcDO3bvKfK0oKiritttu48xJ48Rdd/183RVXfZek49J/wEA++XQVF825mJTvM7B8KI89+ogYP34CE8aPF6eNLhc7d27nkUceJpqdxQfvLdWZkQwsw8D1ICMjY/34iWdQ0K2IUCR0tLikJ1u3VDD7gouW7dq9l1hbnEgkk+HDhwjbgNWrN+iNG7/myJEjrFmzhnBG5MEOQhhGAE12xXyVChCviooK/eqrr+rGxsb5UgoMQ6aJEVjCUsoTohWJFhLXT+O4QiBkoM+EDBAtkbZRNAJpGCAM4slUmuIOKC8YAnwtEGk5rNBIQ6RxWo2X1tlK+fhpr8SQFlorfF+hPDA0GCKQw1KYCGQnQZXyERhIYQR61u/cR4GbaqSlBh0LBNTV1XH06NFg16f1So+S/IkLFswXgwYPplefUl594/VFO/fuobmljUhGJgpoam5FAfsrD+vt27fRr28fMiNhXNelsbERQ1pYJmz4cpPeurWC3r17X7Phq0166bJljB0/jm7d8mdNnjx5S3ZuDpfNm8uefVW6YvchvWzZMq648krm33C1qDlyjM8+++xeQ4Lr+p06yXGcTh3YoRMjkQjFxcVEIpGnOlSolIJkMoXvKxzHw/d9XNfFV37gdUsT1/XT7l3AnamUT3pKfD/YTJ4HjuMRCUdwfQXSwFOBHaK1xNMCH3B8QEiSnkZhdMANuIDrC7SUJF1NQ3PLFZ7nYabFfsoNrKtkysdxNRqThCuIOZqmFH1jWDQrkziQNKHRg3YD4gYkPI3ru8i444EhOFB1WL/x1pvMmHUeU8+dxr33/ZYDB6u1CIiP7/scOHCAHkXRedu2befDjz7heH0D+w7U6CFDhpJKuXlLliwhNzeX0tJenDZ2NIn2GCXF3Snp0aNt49e79EcffcIFF1zEsdq6hcuWL6ekpITnn3+eNWvX6ZIe3UbOmzfPqq6u5vHHH2fv3r3cede/TOrXt7vYun2fjkajfPzxxySSHqYZ6FC0xrbNtDUacG0ymaR///7i/PPPF+FwOG0sBQSxbRvTlGndGwAfUgRU87wAx5ayww9X2LaBVpBM+lRXH9GHDx/bFRh+Jr6v8TxF0vHxlIFhh/GB9oTI2n+wuWH3vqP64JHmr1ri3jQfg4QvwIjQmvALlWmiDPARRLKzXhMGJByHhOti2IKkgpRh0uKIvBhQm2JOux3GjWRUVqeYvbPBPbSrjYcqmnjvkMsv603y4iYkbEHS0Jgq2FDcf//9TJo0ibFjRgkVGAb6wQcf5KabbtIDB5QJywRDaHzAMAR9y3pz7FgNhmVywYXniQcf+KNOxtu5865/FpYpOXDwsG5sbGTO9+bc09zcfPc77y7m+9//Pm2xGEuXL+X737+O3/7m37ENEyeZCoJXpuntO3CAWDxBW6wd2zbXrN9Qod/782J+8tPbuefX/8amTZv0GRPHCmFItFaB2BEKkSaOaZpUV1e3VlZWRkePHl2SlZV19KuvvtZCCEaNGiWWLVuh6+vrcV2XSCRCvwFlDBkyJJqVkRnTwEsvL9bbtm8nOzvKVVddVVVbc6T0xRdfxHEcPK0G5uUW6CuuupIRwwcJhQlYGJbBgerm/W+99V7Z5m0VxJMJMsIhhCHHZEQzV44cPYKL5px/QXYWH1TVNNS++NJrZBf0oLAgh0sumHxNTlbopbBt4QMHGlP3v/bmu3eEwzahsM2M6dMWFhSEv7e7jQ+XfPL59FgsTltbjFTcucWKhPAt/4KevYp+NWrgqQzvWzKpVNprTIHBjTfdok3T5Nrrr5+Zcn1ClsFZZ0wUTjyh//iHP/Bv//ZvD9lCkxmykBq65WXRv6w3JSUl9OldJJ58aqE+fvw4Jd2Lef/99/XUqVPnogIZlhHN/MWflyy5e+zE02lNtPPue4u5/vrrqlasWFY6dPAQehQVMva00ddpBWvXrte19Q2MGTuOCWecztbtO/T7H7zH/Pnz6VdaJK6+6rt64SsvM3Hi2LThYmBIH6GNTmPHNE0OHToUXbZsBaedNu6o5yk2bdrEgQMHqK6u1vv27WPAgAFkZmZSU3OYN954hTkXX9Q2afxkgZR8vmYjry5aTPeehRxraCv9as1nVO7dR35BIZFoLsfrtrD2qz387r5f6PETyoWr4C/rd+l7f/c4X23eRk5+Bp6XIt7STGZWNn4owrufruWjjTvf/9Xddy7KLOh+x4Zdux/YfOgLIhkW5aNOXThx8ICXHCeFY4V46/ONdzz/yTpSrfXMnXUuOTnh7x1t5d/vWfjq9G1VRwgrRWFWBgN69qAhEWNvQzvL9h2i375abv3OnNV53SxhLlr0tp4y+RyKiorIjWYub2qNXXbkSOOifqeUinOnnyPClq0fffiRW/Ly8ohEIniOi+slmXfJRcIFnnjqeW3bYS6bN49Fb7wRGN6OPykcDh+QUpa1tcUetSyLI7XHaGxr4gc33LBu9RefT3CSSW6/+Ueiw/9fu36jXrx4MTfefCtHao6ycOFCUqkk1197LZZlHNhbWa1Hjh51zQsvPrdw87YdesSwISLQOAZae12C6pCVlUVeXh5SSkxT0rNnTw4ePIjneSxYsKA0Gs2s7jAanefjeunSpYwbPRbTyiI3twcFRb0pLunBO+99wKQxQ/nJrbdy9FgdH69ai5XhUtdYz0cff86YseU0NCd/+cLC19i2cy/9B5yK67fw09vvpH9Jb9Zv3MQrSz+mti3OR2u+JOtPT8198BcL5l33Tz9+4Ie/upfcokLeWLmK8sEDyLZDxBzyln+5mbiVwcDyEcw4bxYZJnzx+Rd37913EDu3gDNGjeb8c8ZSHOWxRs3Fb63fV7Lm6+0cizl8vauKqZP6IzOzsoknU3z0yae0xp3e8UTqlj89+TT7D9VoHzhzyiRx/gUX8s7iPxPNzkGaJoY0STg+FrB9524mTpzIyJHDxRVXXAFS8Of337vF8dyyoUOH0tTU9OPDR6pZuXIFR44coWf3nImVlZVMnjw5wCyBtV9t0G+++SY/++lPVttSuBvXr6e5vp6hgwYTjUYXLluxvOyhR//Irv17F44YOToAKwChT0SJuiJAHYaT53l5vq+Jx+MUFhYye/ZsEY1mVmsduA0azYjhoxFI2hPtPxMy0Lv4gqb6GNOmnsdzzz0t5s6bJW66+Xvin278J+rqGsiIRNm0aRNaw64dO361+vNVFOVFsUnx8hP3852Zp4uJw0rFLdfNEb+66zZSiRiF3YtZueozDh5Pto4aPeKB0ydMIO5oPq/Yz/44L7YDiz9c31hd34CdlcWo8iGMOrWbMIARA/tz17XX8/2pMzh/3FgKwqw80qx+XHU8WZJMKHAMvHafpmNNhAB5yUXniu7du7N48WLu+vkvqiKZ0Z9ceulcnnz6KfYeqNYKGDd+tHjy6WcYNGQoiXhqfDLpYFkGimARpGlyrLYuvmvPHn7wgx9MqK+vZ8mSJZw742zhOEmi0SijRgwnkQYbDEMEQDawdsMG/fbbb3PjTT+iuLDgzN49i+xhQwaTnx3le1deIvbt23NNZdUhBgwaSCgcpr6xgYrtO9m7r0obaQOuK1rUATkmEom0RRwgQqFQiHDYpisc7Ps+th1OBya8MiGgtbWVVNIlN6eA78y7nEgEfDyUgD59SiktLUUIQWtrK1pDfV0jvu9j2Ta5edkM6t9TWMJHxVsRQHl52YjJ084hoTwyc7P5etNX0eIoP5kycRy+61HTpnjlw03X1MC09fuqaE62U1CQyTkTTiMEJIFQQeHqhliS5pYEH3/4BU889vb0+//j99zzq3vYW7GHiAyTZWdgCzNAlLSGaVPOEu+8tUjf+7v7ePrJJ9fdfvtNQkqpn3nmOX74w/m6d88e4tRBZUIADU2tN1umxEumCsORUJ3QmkOHDvHJRx9GamtrKeiWt+6ff3qb+I/fP6hfWPiK/v41Vwk7HNIHDx4kHo/j+RC2TaKZWWz8erN+5513uOqqq+hf1keQhtFmTD9LzDj3LCp2HtCffbqKQYMGMX369McMw9hTUFj8sOP5VOzYxan9S1F+4NUrpdD6BHSY9kXbtAbHcQiFQmkftcMlUkgz0MUZGVmYpr1aKbDS0R/fcUkmA4xZaRcpTBAeRUXdqK2tDZAfG7QwaU8qcjExIlHaFSANDDuCoyHmM6U55YFt4+gUbqKNCHD+xDG7V6z4dOC2RsFnFUcp7D1o5ao9+8ksymDk0J4M7Vs40wL2u8z/9SuLJm1rSOFpg1LbZmBJIeeeciaZ3bpzzMjhs3WbScaThKwwLiANDU4y2Xtw/z7ivt/8x8OrP/8M34exo4eKWLydgwcPIo3Ad0ukPAxEU9gwQak8A9C+R35+PkOGDKG2tpa9e/eigNtvv11UVlayZOlHemC/UjF+7LiLtfJABSjLF6s/Y9GiRSy4YT4jhg8Trg/79lXq48eP75cCqg/XHn/11ZeZNm0aV8+7ROTmRm+MxWIPH6s9TmtrK+u+XE/1kboK0zRxXfekyIxpmmRkZKC1jnbE/zs4F8CyDDzPQwrZGRf2fVUmJZimJGwb2GGB0s5JqS/RzPBjQjsoncL3fdpjmEXdi4lkZpDwfQ4cPkp1bbyiHYkTsklK2H+k9eGaujpc3wMvxeljRxECBhXag04fMQzLCNHuSV75YAWtShESioumnkF+RC7XwLote548lnJpi0TI79+fuddfzrXXnvPA1XNOnzJ7UtmVIcukrb0ZbSgwFO2aPKldRYZtVxuAKTmqfJdNG7/Wu3Yf0In2GIZhIADXU2SGTBrq6xYq3yXDMPYYQKK9nUgoxPe+e7l47plnl7U1t/DsMy/oiAHnTp0GysMAmhrq3o23thCxBCHLYsP69dx68827Ty3rK7SCXbt26sXv/5msnOi8w7V1nz7+1J+Kzj1vOmNGjRACaG9u/9nTTzzJpRdfxrzvXFnVHGtj45bN5Ug6xW/AnTId81WdEJ/WOh0EOBGeC4fDOK6Dr1xSqRSmEV7lpADh4XoxtG7HkE4ab7fSG6P9x77fhildhNZoTbi8fPCE6TNm0NDcRGMswe13/6b8zx9v0l9sO66ffXuNvvOXv+bgoUOEDTjrtNGUFuaVCjeB6Soun3k2+SFNe8sxjjcdJTMrzDkjhzO4sGh6ZhogwpC0tMeQIUGr106zq2lQXHE0wcMfrK589b2PFiOjCjtb0Oq1oASmaZmSppbmq2VWzkutLc33lhR35/e//z1ZWVnU1jfQragYj2AH79i5R7/80kt877tXY9s2f1n7pS7pXkxBXv66NX/5ShcXFnLnnT8VDzzwkF63YZO2TYuQZSOAjEhou21a5UrDpZdeukpKWdOtW8FVAli35i965Scfc8MNN9TUNzZ89eSTT3LBhRdx+vixQgENLbGbH3ro4XvPnjyF0aPLhSkgKzuqE6lkZ26W53lpsWvQ3t5OS0tLmtCQSqVobGzs1KMBzmtgWzZSSlra2vF8UWLb0N7WSENDNVnHNcpPnvCFJQipaW1r4NixgzjxTOLtsQe6dc9acOGsc6isPsr+g1Vs2vg1O3ZUYEsLMyOThlSSkGXQt6iYH117DbmhULVOtKF8zEG9s8TkUf31n9duwDc1doZg7pTvUyyNj4TykdJgfHn/W0Zt2/Hw6kOHUX6CF556lGLbLtFNzSVmyKa4uBvtXoyjBw9h985FQsJUyiMajb6kfEWf0j7ikUf/RDhi4AN3/PQXeufefQwY0Ifd+w7pRx79E9ddczUTxo8Ra9es1S8sfInbfvIz6muPTVixbBm52TnMnTtX33HbreKDD5bq/Nw8pNBBAEHpaMg0ScWThSXFhWf7fhBTqKw6ov+yfj3z58/fXVxcPGjt+r/oadOmMXjw4Csrdu7WgwcPFBXbdz7cs3cpk846UwgRwG1NDY1EM7MCKeK6mKaJlBLP8ykoKGDixIlIKZsAhg0bRiKRSAdHdCecqNEUFxe/P3Xq1AtsO/QuwLTpZxLOMsjJidKvrDdSSFxfYxlgmub22bOnlU+cOIr8nFwsU1XZwMxzxoiRg4e8+fzC1+buqz5AXVMjUtiEo5m40uXssyczfcpZNacURnvqlIslQwjL9lwHrp19OkU5GmVKCrOzGJgTvSCazvlyfU1fWzzysznnTemxYuXFh5tbce1cIraNypKcftpY+vQ5hb27dpMsO4WiDJscn5jwPEV7e/uQaHbWjpQfBBsU8PRzb+iXX3uVR/74R0KhEPf//j7+acF8xo0YJL78cqN+9aWXuemmmzilrEx8+ukXeteuXaRch8GDBzN16jniow8/1OFwmJBlMGXyGeJIzfGKzz77rPzKK+YJXwURCwWs/PQzHY1GGTNqlOiwVhXw9DPP6aTrMW78RJKOx/Hjx5k16zzxwXsf6GVLPqBP7978aMEPqkqLi/r4XmAIBW6MwjRlZ7QjgPWMLqG0dBDLdQNLGY00THw3ILhpiy7BD43rOIStEClXY6d/0x1BcQHbt1do5SpGjhguHA8a21Oza47Xvy9CERQ+RgjKehSKAHsCnfTxU8keGdHMo56ApAgs3JSvzWxDeBHA1JByPYRpIIQgJaAZONKaXKtl6EDKc6dlhc2ncsPykTDUZXQENRRENJhKC8JZWTtiSQ/TDtJIHn1ioW5sbaf/qeUkPdi5dyeDh49i+IhBYu3GHfqNRe/w41tup6yst7AknDPlTLFtyyadk5PD2ZOniER74jLf97Esi2PHavB9MA1jR3ustfxIzfGK4uLioQlXYVsSK5JB3ElhGgJfBfnrxxua7j9YWYk2TBKuRygSIal83l+6XH/+xRd85zvfYeJpo+4uzs/5TRD96sBvvXTWg+7UpUGsNYjQCBEQWgiBZVk4joM0glQbjI5EdUXCTWEaNoY0sAwDpX2EMDo3heO4hGyL/fv26qeffJyL51yI8IcQEpIeOaEPCqI9hTag3YUMCxKpZNgww0lTgqcFGdmZR//87hKtUMy+aKaIGBIM6SWSLmbYwvV8EApTGihPYSLINQXZ2eGJQfjUxkdhd6kN0IAlCfxUDPA8PwDGJWzfdVB/sWYtl1x6Gd1LevHW2++yYdMW/rL2Sx5+fKH+zX2/4/IrrmRA/97CcQJjJNHePq2oqBuDBw4kI0OilCqJx+OMHjNyxLZt23jjjTd0UVG3eacOGsgzzzxT3tTaenPICvJ3XM/DV0F+jSmhpbX9jFdfeuWO0aNO48abblk2pHzoxe8vX0637j1Y/vFKrr3+Os4/d4ooyA0Iqnwf0zRxHAfTNLEsi1QqlUaTzLRP3JmTnoYTgw1gWTZSCuJxB601rucAiohlodPxrI7YqWFAKuWhNERsC9dJsXbtGiZNmsRZZ50ltJsCPHwniZHOQsqyQPsuEVMk8Xx8F+ywxHU17fEY9XVHMYVCKg9DaUJC4KQUpmVgWjaOF0STQqZAOj4hBWFPE9GaqPbJwCMThaEcIijMdMqoTCbjYc9zsiAICVZs2Yrvuvzu3vsYPHgw0WgO48ePp7B7Mb6GkaNH0adv33tcghzfwDJ2JtfW1pKTkxMsnAiMk+ysyNa7775bHKis5NXX3tCTzjhTjBw5igcffPDhY3XNz2oglBHB8T0k0BpLDH/ggQdXDxg0kEsvu0hkZede+OLLr7w7fORoWltjNDW10LdvnysTjocpwXPdznIM2zZx3cDVsCzrpHzgVCpFQ0PTjzpcllisvbcQBrFYe+9UyiESsTEMgW3ZOCkXgUwbeIH/m0zGw77rY0qBma438hy3x8iRo5k9+8KokDbYETBs2hOpydIQOMmAsJZhYUgrSKyz05xlC0Ihg8KiApJeCiVNkr5A2CYyJGlqSwypq2+4X6HRZgD028LA8HyEG8dUHoawUFpiAmENhudipFNYzMyIlRQE6JDva4oKCqg6sJ/p519INDPCRRfMZmvFNvAVE8ePo7RncU1xXsYvguzAgO1NU+4Ih2wKigq2i4CgV9jhEO0JJ5wZsZMLFiy499FHH73z7cWL9WWXzBHhrKh+4A8PXXfLbbdFpTSxDZO6xtZ/X/zu23ePmTCOGedNFwkfHnzoIbd85BgGDy3n17/+NaeffjqWaa7rMHRMMwh2p1IpQqEQlmXheaozaP7hhx/q48eP09TUhFKK7Ozsx0455RTOOeccobUmGs2s9hW0tradsX79+tVVVVXE21uwrBDZ0XwmTZpEn769hBkxkyD56KNV2nMTzJgxQ5imfXT40OGiIylw374DeuPGjVw29xKRTDqYwsbQ8NXGr/WWrZuItSXJyckhIxrmsksvEcISNDY3Y1hByE5aUN+UuOyjFR8uampuoPZ4FTl52Xf0O3UwE8dOvLs4N/c3QknMcIiUmwJfYNsm2lcIYZ4oXvM1UvtJUsnWcGtz3fyIJTht5LC7H3vkIX5x123inMkTIvm50YXLl33AOWefxcQxg0TP7nk9a47XV7z62ht69+5dWmtoaWl5NS8vj9zc3NM8BYYld0QyM/C16hFPuhR0y7vrtp/ccdr+/ft56NE/6cHlQ1ZPmjKZr77eODcVTxAOhRCmcaCpLcZ5M2eI5nhqzH0PPKSHjx7FzBlnirAV3njJRXMIGxb4qsiQojNLoWtwvGue0p///Gf95ZdfUlxczMSJE5kyZQrRaJSVK1eyePFirbUmmXRwnSQLX3x+9YYNG8nP78aZZ05m6JBhtDa38uyzz7FzV5B4B7B7904qKrYiJUhh4nvQ3treWwB7DlTyyedfEE8mCzU+EsVLz72k33rtVbrl5TLpjIn07duH2vrjPPfyi7q2pYX8oh6kYk5vC8m27bv0gw//ftH+fXsYM2IY8y67hOHDytm+5Wuef+aJexKxRB6moCUWOwPTwrLMwFgzJG5nYF8jTAMplCBshZMZ4chTWinyc6O/OXPiCJFhQW6Y5OuvvHxNW1MjOdEoHdz58ccfl1cdruLjT1bS1NR0e0ZGxlOu65IVsZOGAEPImtbWVjIi4cqIbaF8iGZlbJw+YwaO6/PQI49O6t69Oz6BIeOkPHxPl53Stx8aePLZF77q2acP0Zxcbv/Jv+q8vIzpmZEM6o7XYkprixTg6yAFx9cBgtRhFAkhqKmp0Vu3bmXWrFnMnDlDjJ8wTowZM0p85zvzxIUXXshXX33VKaaXLl2qW1pauPrqq9suueRCMWrUaHHWlDPFggXzRY8ePdhesRMpJG1t7X2j0SglJSXpZwXJYeFwuDoWi+f16NGDXr16obXOCofDrFixTO/fv5ebb745MWfOHHHaaaPFGZMmiu9973v5RUVFfPX1RizLIhLOrI61Jnp88snHDCsv5yc/vU1MnDhWlJeXi+lnTxXzr7/uuYxwiIWvLGz0PUVWNHeNRiLFiQwNaUowZFCcFqSlRUBbhOzMAOC2BL7rYwP7dh3S2zd9zdlnnsWunXt4b/kXumLHPl1TW0dLW4yBg4aQn5/3oBCiTSmPXTv36CDX1ajLysoGwJZgKB+hIT8/n7y8AjKjWfjogBgySE3Ny8v6RTgURQNtrUnCoSy+WLOGkWNG88wzLzZOOWucqK05ipNIXmcQVBT4KkjC0kiSSacTrDdNs609EaPq8CEam1ouS3l+pysydtw48cMf/pBUKjVeSkFeXgFTp51LaWlJdkc4ri0WL6zYsV3X1tcSCoVwXZ9oNLPS930SKQdfB2Wx0gzsh3Ak0hRvb8P3HDLDmZXtsUTh7r17OGvq2fTsU5qBMPB8DVqSEc5sOnPCpJkjBg2joa4ew4Kqw4dqjh6p5pwpZ70aDpld6n4Fhd2Kr5/3nSverz5SReWhA1oCwg8S7oVMZ5fqoJBY60AfmtIySKUcQmEDwxA4TjKIXHgQb2+juLiQzGgOvfqV8fCjf6R/WR9OnziW0RdeyIC+pUIB3boV/GTcuNPuePnlhcxfcKOTU9CtZyyZejjhKUKGxLIMkkCvXiXizMmTdTgjc52rvfL9Bw5E87vlEUsm2V95RG/atIkZ551JNJqDQtO//6ls3LiB8sGD2b+/Wmfn5bJhw4bHZpx71p+CEkoPUwp81yHIhNC4rktxcXH2FVdcoRcvXszGrzYtKu17CtnZ2ViWxZCBg+jXr58IhyyUgrMmTxa1tbXvvP76m9p1XZqbmzshx5aWJqQM3t91AyQqbGd0JoIFebgKyzCwDROhNI7jUFtbW3vs+HFKevYMEtjS8KRpSnxfk5edszw3mk0kFEYBzW3NmFLw7ttvXZmRkXGl53mYUpJKpTrDiLX1ddTW19G/f39MeULdBEn4J9dmmI6TwjDB95MYphlkoPouhmkxdPhQcX1uvt6y9wCZ3Qr4/g8XsHXzJrqV9GLFJ5+zs6S7vmjWNCGVZtiw4aLq8BG9dPlya/r5F+zCFJiWJOEp3nr9db1j3z4W/NOPVvfv21O4QFVNvc6MZuG4PpGsKBW7tlNcUsTGLbu0rx1yotkMGDKI9Ru+ZMDgQaz/+ksuuOxCnn/ySYYMH9RQ2qOoIOUF3C4A7XkkUsksoSEzMzM2cvgo0bOk96ObN2/+sZSS7du34zgOmzd+Td++ffXl8+aV2rZd/fWmTfrDDz8kKyuLbt26MWjQICKRCAMHDrz3nXfeubMjUtMR0gvGiaTwUCgwzjpEv2VZne6UlJL29va+0WhmpWlK2tsT4XA4nAzuC+G6Lo4buFOWZVFSUkI0GiUej2ObZmdwwrIsRowYQffu3U+qMOia6tr1MD9Z9bGeMWOG0MoDrbBMA1/5uK7EtAx69S0R++pq9fvLlrJm/VqmTz2H6qPHqKmto29pbwxASEEsHu8xa+ZM8dJr7+jKQwfzc/Pz8ICQKYkl4oRDEZYsWTbp+u9fjfIgFAotam5unVtQUIhSHls2byYejzN16lSk0LheipBprQtZ9oSwHcIwJGNGDhf1F5yv7/7FL/KvuPxKPWv6ZBFPJc3MUMhDSizLioUsm6NHj36aSrlTSktLxYxzp93oeopJkyahtaa2tlY//vjjVFRUVJWVlT23fPlyJk+ezBlnnCGsdOJtkGsLQog7g4ULiOh5Hjk5OZ1Ah2EYOE7g4ycSiU6UyrbtdUKICa7rEo1mVnagXJmZkaTvB3VIzc3NnXPl5+fjeR6TJ0+em5uT+1a6gqkzRScWi53a2Ni4u2fPniKRSBAKhTqJ+q2lX15HsYwUeEHZNoZh4nkOiZTH8uUf6XiynfxueeR1y6OoqIg/Pf442ZlZnHfuZOFpqKtteTESyjgqASGCvFQ35YAOfLXs7GwqKytpbmik9ljDFxkmaNcbEo1k4jkuXsrhtptunHvX7bffU9arUHjJOGHTwPecCal4gnhrGzmRTKTnc8nsWeK671/Porff4YXX3tLxlHOdoxSGaRFPOmM0UF/fOOXZp5/h8OHD8SDQIQmHbSKRED179hSB6+MhpaxJJpMUFBQE1qToNKjZs2ePrqioIBQKBcsjJf369aOiooKmpqabQyELw5DYtsnRo8c/XLVqFUopUqkUJSXdJ44dO5aXX36ZurqG/zBNeZKrtWLFCr1169bA5rBsCgoK7s3IyGDVqlWLAoQsyOiXQlJfX3//E088sXvFsuVYRhBS7MhbPlE18A1OPXPy2SW+0qAkpmXiJROYIRvLsjBMk1islUR7G2dOmshll19290fLP7xnUP9+ZGdFsYCtFQf0zm1bGNK/9Joxp40Rbhqqi7e3ExaCZDLJvEsvE/PmzmPNl5v1m6+/MWnu5d9ZGzKs9ZY0y0OWTVNrC/nRjLeU5i0TiITtAHn1XPKiWTjxdlQyiS0NpIYzJ40TQ0aOmv3b3/7m/W4FOU/OnD71qZTvEYlENrquT8+ePa/Jyspa+MpLL0dOO+003bN3L8LhMIeqqti+fTuWZdG/f/9lhYUFvxg2bNjdixYtoqamRufl5SGE4ODBg1RXV9O7d2/27t1LWVmZLi0tFaeeemrjBx98kP/uu+8+PHjw4IcTiQSmabJx40ZKSko4dOgQSqmw1iTPP/98UVVVpf/4xz/eOXny5Dtt2wZg27ZtKKUoLS0NDEU0+fn5d82ZM+fOF154ATeV0qWlpYTDYRKJBBs2bKCpqYlZ583ESyNMpmlimubfrAkyszIzjiZTHhE7fZG0O8vnlA/XzLtExIDNu/bqxx64/57crCg3/fAG/vLJZwCsX7+eSNhm286dnDZuDPFUCiUkmeHMoHjHtBBIEi6MHTdS1NYc1evWrZswddq51ycTiesMIdIBbfBdH8M2OkNmSvvEYq0MHDCAD7fvoKW2aWn3krxZCQUZWdYHsy+5mKpdO1DpwiBP+URMm/z83Jfmz5/P559/vnDjxo1s3rqF3Nxsjh07RmlpKdOmnk1xj6JZiVSS888/f2ZeXt6ytWvXUlhYSGtrK7179+bSSy8lGo3e8tRTTz188OBBSktL6d69e8GNN96oly5dyurVqztKO5gzZw7hcHj78uXLyzs62vi+z4IFC8TKlSv11q1bicViCCEYPHgwEyZM2H7w4MHylpYWfNfFtmyGDBok5v/gB/qLL75g1apVRKNRamtr6d+/PxdccAH9+pYJ3/c7Axd/qx4IAMfRKK1xvWB0lAE6rk9KaVq05vWVn+rLbrpNX/Wzn+vNxxsbNlYd1fc+9rRu15q99fGH7rjnfl3Z0Phsm9bc//yretn6bXrR8lW61dUkPU3S1xxLeHmfbt2l31j+sX51yYd65+H6hrdWfq5XfLVZL133pW71Neu2VuiE1vzst/fp1z/6WG+trW297mf/out8zXsrP9d/+NML+mhzanKtq6lM6lNv+f0f9BsrV+pmzyOuFe2ui+sFNbA6PdpaYr2PHjm2tKqqqrWpqWl2B3zo+h5KB3+7rzRtsXg41p6gvqHpWs/Xnb/F2hM4rt95rrTGcX0am1rOa25pG+75wf2O65NyPDxfk3K89JoqlA7mbm5pG97U3HqGr4J5U45HMuXiaxWUvaSf5ymfhqbGKw5VV7U2tTRPVlrj66CMssO6dxynU8d/W02uacgTHT6EEQDrSkikCY1t8TN+c+99qw/X1fHTO+9iX+UBnFQyXzseoUgAb+UWRG7NLym+JTM/7/qa5sb3DjfU0q25gUwZgOAKcHzF5h07Gtdv3Eh1dRVTJk/Gt2RzazKen2FGUUjaPLd3xe49DBpWjh0Jk3KTdC8szJ45e6Z+6tln9LXXXFukbLv2seefXdW9V0927dtPa1sTo0aNWhVk6AcFndIOrGE8DYYgK5pZnRXNnPXNnjYdbolIG0ZZmZEkQGZG+Pmu12VmhP/q3DIlebnZy08yTkz5V3Ob6WBCeu6tXa+3LeNEA5YuetEQkvzcvNfyc/NeO1H9LzrrH7uK3a6GUtc5pBRBcU1nzQjgCY2j4eE/PrI6kUjw2COPRQaX9RbJZBLTsIhkZiRyC/LxgO17KnVzczNvvfuu/s19914w5ZwpbNuxjU1btqLSMUfTlOR1K+BYfR2l/fthREIkXadMWjIoYzQESc+f0pYIUj89T2EZJgYwddIkUdyrhGdfWVjbr3zwumhBHlbIpkdxIb/++c8vKMjJOduUBsr18H0/S4qOqj7xLVXZ/xiHDKjsp7uuKFRaQx2sqtJbtm7lRzfeDFIkkwosO0RbIokZDi+KZOeyZdc+fcdtt3PoYCXtrW3MOPdc6muPM2HsWJqamliy/FOtgJSGwm5FMydMmEBBcXc0AicRxxSB7nE9DzsUfk0hcBTkZGSh4ynC6Z188cxZondZP95d8v6E4w1BycSsqedQkBH6wBJBh69IKEQkEooF5YPqRBn3f9JU5X/7+Gavh/+Ow0S7adYOauE6cn5aY21kZ2fjOA62na7YcgNFHU8lr3nrnbeJRiJce/VVTJk82S3tUWhrYP/hI7pXr54iN5ytn3jiCXbs26NnnHce/fr1ERHLpvpoLSHDpHtR8R1hy3zgWHMDlmHjOO55WkFYglCCSCgDpSGZSo5xhOo9bcpZ4t//4z7dVN/AzfMXTM8PWx9ZgBYS33NBSCzDQOgToqqzB5/4x+JUs6MpgRDGSTvGTJf2GYaFAJpbkz+vqjrI+vVraWtrY/jgci6cPZMBvXsIpYKWPkLCwF49RUKDHQkTS3ls31/JZN9HCsiwQngJl4gVIdO2Xiru0/vBptY27XsuFsaBrEgGGoi1teF5CiEgOxze2O6rjas//0IbKYc7b7mVnLD1kQm4SZdIyCDluXi+KhSmWWdZFkIE9aFB+z3zbzYQ+/8vUdGdiTsajZXm1sOHjlCQ142KLVs53tCgP/lsFb37llJbU8OPb/gBIwYPEMF+0ISNgBti8VSeskWTr8lb+ekqLrricqyMMGZWFraEiB3BlCbKVeCrQgl1iUQKIQ0yI3KHJYP2ib1O6UMoGgWg8shxvebzz/j000+5/bZbGFRaIpQGT2nssIXnOphCYoSsOiklQukgWceQJ3Vv+UdSr6YWQaMrjezsJKp8cBMOY4aP5uuNG2n5/DPKRwynOD+fbuPHMmxAPxHpKJg30q33NGRGQk2+hOP19SurjtUw+3vXLNqwZdPcVLrHUkFBAYlEioxIJpZl7QCwwxFSrseByhpdfbiKzTsP6MuvuFjEgWfefFcfrTqMdpI88NvfPta9W96NrquQlsQ1BEmCrqO26NpNw+jagweE+gZB5clU/t9mXf0tSvb/nlODSmhDnmgJa0gYMbSc7bt2cuOCBfQ+pVik0glOIv1pAr7rYZgCZJCUpQChNbZprBPSH1OxddPcosJ8GhrqSPUqQpiGW99YZ23bWcEZ44ZiAL5hYFkhNm3fjpmVyZKVH3Lq4B/y1uIP9P6DlVx35dWc0qNQZADSh7ApcbTGEgIfhSlk+mWCiImBn26u0AXz+wc7hOqC8mt0YGh0JGoBrqtRQnXpJSuDXF4d+Ea+VkG3Mx+0YSAkJLTm6+079AcrV3K8qYmamhoG9C1j7559DCwfSo8ePThUuY8Zs2ax88BBhGExbvQoaqqq6V/Wl6a6ep5/7hn+/Ze/rDqlMK+PESTOI3wfKYJmHB09+4L30EiME01a5T8mMU8iatdyQJT+qwZRXcM78hsNKoPGFuD5QZc0F4VpWjhac6y+fldtfd3AaHYunob2RJy9e/ZzvL6Ofv36sWnrVtpdxYgxo3BTDr2698CWgpdeeJ5Lzp/NpIkTRJYlkBqkAnwvIGp6YyGDAHlnu9r/0V6//x8make/32/2j+2AnDqiAX+LqEKmmylridKKpJNCC4NQyO7s2ed3qDeCRhhKQSgEx+pbn928c9d1bck4e3fvoamhHtu0OG/aVMaPGS1sE7SrMWRA2KB1mH9SdxWk+Kt37Erkf0QCn0RUQWdy3EmY4slmhjixoKQbOEkJSgQ9fzmBOQbR/hMxSgAn6aez5i0MExriqSF1DfXbS0t7CknQGcxIBxOMdPK1kc691VoHP3QhnhZ/zZlSyn9oTv0/yWmZykD7GaMAAAAASUVORK5CYII=",
  cli_bde: "data:image/png;base64,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",
  cli_barclays: "data:image/png;base64,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",
  cli_bcc: "data:image/png;base64,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",
  cli_bbva: "data:image/png;base64,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",
  cli_sabadell: "data:image/png;base64,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",
  cli_vodafone_es: "data:image/png;base64,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",
  cli_yoigo: "data:image/png;base64,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"
};

// escape per uso in RegExp
const escapeRegExp = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

// sostituisce tutte le occorrenze (case-insensitive) del nome cliente con l'etichetta anonima
const sostituisciNome = (testo, nomeReale, label) => {
  if (!testo || !nomeReale || !nomeReale.trim()) return testo;
  const re = new RegExp(escapeRegExp(nomeReale.trim()), "gi");
  return testo.replace(re, label);
};

// risolve il cliente collegato a una referenza (con fallback al testo libero)
function risolviCliente(ref, clientsById) {
  const c = ref.clienteId ? clientsById[ref.clienteId] : null;
  const nomeReale = c ? c.nome : ref.cliente || "Cliente riservato";
  const label = c ? c.labelBreve || "Cliente riservato" : "Cliente riservato";
  const anonimo = !!ref.nda;
  return {
    nomeReale,
    nomeVisualizzato: anonimo ? label : nomeReale,
    label,
    anonimo,
    descrizioneCliente: c ? c.descrizione : ""
  };
}

// applica l'anonimizzazione a tutti i campi testuali di una referenza, se richiesta
function testiVisualizzati(ref, clientsById) {
  const {
    nomeReale,
    nomeVisualizzato,
    anonimo
  } = risolviCliente(ref, clientsById);
  if (!anonimo) {
    return {
      cliente: nomeReale,
      titolo: ref.titolo,
      descrizione: ref.descrizione,
      attivita: ref.attivita,
      risultati: ref.risultati,
      contestoOrganizzativo: ref.contestoOrganizzativo,
      kpiDettagliati: ref.kpiDettagliati,
      elementiUnicita: ref.elementiUnicita,
      noteAggiuntive: ref.noteAggiuntive,
      referenteContatto: ref.referenteContatto
    };
  }
  const sub = t => sostituisciNome(t, nomeReale, nomeVisualizzato);
  return {
    cliente: nomeVisualizzato,
    titolo: sub(ref.titolo),
    descrizione: sub(ref.descrizione),
    attivita: sub(ref.attivita),
    risultati: sub(ref.risultati),
    contestoOrganizzativo: sub(ref.contestoOrganizzativo),
    kpiDettagliati: sub(ref.kpiDettagliati),
    elementiUnicita: sub(ref.elementiUnicita),
    noteAggiuntive: sub(ref.noteAggiuntive),
    referenteContatto: ref.nda ? "" : ref.referenteContatto // il referente non va mai esposto se la referenza è NDA
  };
}

// ─────────────────────────────────────────────────────────────
function RepositoryReferenzeInterna({
  gruppo,
  userEmail,
  onLogout
}) {
  const [refs, setRefs] = useState([]);
  const [clients, setClients] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("repo"); // repo | add | clienti | assistente
  const [addMode, setAddMode] = useState("testo"); // testo | voce | file | manuale
  const [q, setQ] = useState("");
  const [fAmbito, setFAmbito] = useState("");
  const [fSettore, setFSettore] = useState("");
  const [fTipo, setFTipo] = useState("");
  const [fPaese, setFPaese] = useState("");
  const [fAnnoMin, setFAnnoMin] = useState("");
  const [fAnnoMax, setFAnnoMax] = useState("");
  const [fImpMin, setFImpMin] = useState("");
  const [fImpMax, setFImpMax] = useState("");
  const [sel, setSel] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [clientSel, setClientSel] = useState(null);
  const [clientEditDraft, setClientEditDraft] = useState(null);
  const [toast, setToast] = useState("");
  const [aiPrompt, setAiPrompt] = useState("");
  const [aiScope, setAiScope] = useState("tutte"); // tutte | filtrate
  const [aiRunning, setAiRunning] = useState(false);
  const [aiResult, setAiResult] = useState("");
  const [aiErr, setAiErr] = useState("");
  const [rawText, setRawText] = useState("");
  const [busy, setBusy] = useState(false);
  const [preview, setPreview] = useState(null);
  const [recording, setRecording] = useState(false);
  const recogRef = useRef(null);
  const fileRef = useRef(null);
  const [fileName, setFileName] = useState("");
  const [fileErr, setFileErr] = useState("");
  const clientsById = useMemo(() => {
    const m = {};
    clients.forEach(c => {
      m[c.id] = c;
    });
    return m;
  }, [clients]);

  // ── immagini per referenza: caricate/salvate su chiave dedicata ──
  const caricaImmagini = async refId => {
    try {
      const r = await storage.get(imgKey(refId));
      return r && r.value ? JSON.parse(r.value) : [];
    } catch (e) {
      return [];
    }
  };
  const salvaImmagini = async (refId, arr) => {
    try {
      await storage.set(imgKey(refId), JSON.stringify(arr));
    } catch (e) {
      showToast("Salvataggio immagini non riuscito");
    }
  };

  // ── logo cliente: caricato/salvato su chiave dedicata (stringa data URL o null) ──
  const caricaLogo = async clientId => {
    try {
      const r = await storage.get(logoKey(clientId));
      return r && r.value ? r.value : null;
    } catch (e) {
      return null;
    }
  };
  const salvaLogo = async (clientId, dataUrl) => {
    try {
      if (dataUrl) await storage.set(logoKey(clientId), dataUrl);else await storage.delete(logoKey(clientId));
    } catch (e) {
      showToast("Salvataggio logo non riuscito");
    }
  };

  // ── caricamento iniziale (referenze + anagrafica clienti, con seed demo) ──
  useEffect(() => {
    (async () => {
      let refsData = [];
      let clientsData = [];
      try {
        const r = await storage.get(REFS_KEY);
        if (r && r.value) refsData = JSON.parse(r.value);
      } catch (e) {}
      try {
        const c = await storage.get(CLIENTS_KEY);
        if (c && c.value) clientsData = JSON.parse(c.value);
      } catch (e) {}
      const clientiMancanti = DEMO_CLIENTS.filter(d => !clientsData.some(x => x.id === d.id));
      if (clientiMancanti.length > 0) {
        clientsData = [...clientiMancanti, ...clientsData];
        try {
          await storage.set(CLIENTS_KEY, JSON.stringify(clientsData));
        } catch (e) {}
      }
      const refMancanti = DEMO_REFS.filter(d => !refsData.some(x => x.id === d.id));
      if (refMancanti.length > 0) {
        refsData = [...refMancanti, ...refsData];
      }
      // migrazione: completa i campi introdotti dopo il primo salvataggio (periodo, modalità fornitura, codici progetto, ecc.)
      let migrato = refMancanti.length > 0;
      refsData = refsData.map(r => {
        const demo = DEMO_REFS.find(d => d.id === r.id);
        let out = r;
        if (demo && r.inizioAnno === undefined) {
          migrato = true;
          out = {
            ...demo,
            ...r,
            inizioMese: demo.inizioMese,
            inizioAnno: demo.inizioAnno,
            fineMese: demo.fineMese,
            fineAnno: demo.fineAnno,
            inCorso: demo.inCorso,
            modalitaFornitura: r.modalitaFornitura ?? demo.modalitaFornitura,
            quotaPercentuale: r.quotaPercentuale ?? demo.quotaPercentuale,
            partnerRTI: r.partnerRTI ?? demo.partnerRTI,
            attivitaAutonoma: r.attivitaAutonoma ?? demo.attivitaAutonoma
          };
        } else if (r.inizioAnno === undefined) {
          migrato = true;
          out = {
            ...emptyRef(),
            ...r,
            inizioMese: r.inizioMese ?? new Date().getMonth() + 1,
            inizioAnno: r.inizioAnno ?? new Date().getFullYear(),
            inCorso: r.inCorso ?? true
          };
        }
        if (!Array.isArray(out.codiciProgetto)) {
          migrato = true;
          out = {
            ...out,
            codiciProgetto: []
          };
        }
        if (out.tipoReferenza === undefined) {
          migrato = true;
          out = {
            ...out,
            tipoReferenza: "standalone",
            parentId: out.parentId ?? out.accordoQuadroId ?? null
          };
        }
        // migrazione dal vecchio dominio (accordo_quadro/stream) al nuovo (programma/servizio_ams/progetto)
        if (out.tipoReferenza === "accordo_quadro") {
          migrato = true;
          out = {
            ...out,
            tipoReferenza: "programma"
          };
        }
        if (out.tipoReferenza === "stream") {
          migrato = true;
          out = {
            ...out,
            tipoReferenza: "progetto",
            parentId: out.parentId ?? out.accordoQuadroId ?? null
          };
        }
        if (out.parentId === undefined) {
          migrato = true;
          out = {
            ...out,
            parentId: out.accordoQuadroId ?? null
          };
        }
        if (out.paese === undefined) {
          migrato = true;
          out = {
            ...out,
            paese: "Italia"
          };
        }
        return out;
      });
      if (migrato) {
        try {
          await storage.set(REFS_KEY, JSON.stringify(refsData));
        } catch (e) {}
      }

      // seed dei loghi cliente estratti dal documento originale (solo se non già presenti)
      for (const [clientId, dataUrl] of Object.entries(LOGHI_SEED)) {
        try {
          const esiste = await storage.get(logoKey(clientId));
          if (!esiste || !esiste.value) await storage.set(logoKey(clientId), dataUrl);
        } catch (e) {
          try {
            await storage.set(logoKey(clientId), dataUrl);
          } catch (e2) {}
        }
      }
      setRefs(refsData);
      setClients(clientsData);
      setLoaded(true);
    })();
  }, []);
  const persistRefs = async next => {
    setRefs(next);
    try {
      await storage.set(REFS_KEY, JSON.stringify(next));
    } catch (e) {
      showToast("Salvataggio referenze non riuscito — dati solo in sessione");
    }
  };
  const persistClients = async next => {
    setClients(next);
    try {
      await storage.set(CLIENTS_KEY, JSON.stringify(next));
    } catch (e) {
      showToast("Salvataggio anagrafica non riuscito — dati solo in sessione");
    }
  };
  const showToast = m => {
    setToast(m);
    setTimeout(() => setToast(""), 2600);
  };

  // ── voce ──
  const toggleRec = () => {
    if (recording) {
      recogRef.current && recogRef.current.stop();
      setRecording(false);
      return;
    }
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) {
      showToast("Dettatura non supportata da questo browser");
      return;
    }
    const rec = new SR();
    rec.lang = "it-IT";
    rec.continuous = true;
    rec.interimResults = false;
    rec.onresult = e => {
      let t = "";
      for (let i = e.resultIndex; i < e.results.length; i++) if (e.results[i].isFinal) t += e.results[i][0].transcript + " ";
      if (t) setRawText(p => (p ? p + " " : "") + t.trim());
    };
    rec.onend = () => setRecording(false);
    rec.onerror = () => {
      setRecording(false);
      showToast("Errore microfono o permesso negato");
    };
    recogRef.current = rec;
    rec.start();
    setRecording(true);
  };

  // ── abbina il nome cliente estratto dall'AI a un cliente esistente in anagrafica ──
  const abbinaCliente = nomeCliente => {
    if (!nomeCliente) return null;
    const n = nomeCliente.trim().toLowerCase();
    const match = clients.find(c => c.nome.trim().toLowerCase() === n);
    return match ? match.id : null;
  };
  const previewDaStruttura = (s, fonte) => {
    const clienteId = abbinaCliente(s.cliente);
    const clienteMatch = clienteId ? clientsById[clienteId] : null;
    return {
      ...emptyRef(),
      ...sanitize(s),
      clienteId,
      nda: clienteMatch ? !!clienteMatch.nda : false,
      fonte
    };
  };

  // ── file ──
  const handleFile = async file => {
    setFileErr("");
    setFileName(file.name);
    setBusy(true);
    setPreview(null);
    try {
      const ext = file.name.split(".").pop().toLowerCase();
      let structured;
      if (ext === "pdf") {
        const b64 = await readFileB64(file);
        structured = await aiStructure(rawText, {
          type: "document",
          source: {
            type: "base64",
            media_type: "application/pdf",
            data: b64
          }
        });
      } else if (["png", "jpg", "jpeg", "webp"].includes(ext)) {
        const b64 = await readFileB64(file);
        const mt = ext === "png" ? "image/png" : ext === "webp" ? "image/webp" : "image/jpeg";
        structured = await aiStructure(rawText, {
          type: "image",
          source: {
            type: "base64",
            media_type: mt,
            data: b64
          }
        });
      } else if (ext === "docx") {
        const buf = await file.arrayBuffer();
        const {
          value
        } = await mammoth.extractRawText({
          arrayBuffer: buf
        });
        structured = await aiStructure(value.slice(0, 12000) + (rawText ? "\nNote: " + rawText : ""));
      } else if (["xlsx", "xls", "csv"].includes(ext)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf);
        const txt = wb.SheetNames.map(n => n + "\n" + XLSX.utils.sheet_to_csv(wb.Sheets[n])).join("\n\n");
        structured = await aiStructure(txt.slice(0, 12000) + (rawText ? "\nNote: " + rawText : ""));
      } else if (["txt", "md"].includes(ext)) {
        const txt = await file.text();
        structured = await aiStructure(txt.slice(0, 12000) + (rawText ? "\nNote: " + rawText : ""));
      } else {
        throw new Error("Formato non supportato: usa PDF, DOCX, XLSX, CSV, TXT o immagini");
      }
      setPreview(previewDaStruttura(structured, "file"));
    } catch (e) {
      setFileErr(e.message || "Elaborazione non riuscita, riprova");
    }
    setBusy(false);
  };
  const sanitize = s => ({
    titolo: s.titolo || "Referenza senza titolo",
    cliente: s.cliente || "Cliente riservato",
    settore: SETTORI.includes(s.settore) ? s.settore : "Altro",
    paese: s.paese && s.paese.trim() ? s.paese.trim() : "Italia",
    inizioMese: Number(s.inizioMese) >= 1 && Number(s.inizioMese) <= 12 ? Number(s.inizioMese) : new Date().getMonth() + 1,
    inizioAnno: s.inizioAnno || new Date().getFullYear(),
    fineMese: s.fineMese ? Number(s.fineMese) : "",
    fineAnno: s.fineAnno ? Number(s.fineAnno) : "",
    inCorso: s.inCorso !== undefined ? !!s.inCorso : !(s.fineMese && s.fineAnno),
    importoKEuro: s.importoKEuro ?? "",
    ambiti: Array.isArray(s.ambiti) ? s.ambiti.filter(a => AMBITI.includes(a)) : [],
    tecnologie: Array.isArray(s.tecnologie) ? s.tecnologie : [],
    descrizione: s.descrizione || "",
    attivita: s.attivita || "",
    risultati: s.risultati || "",
    ruolo: s.ruolo || "",
    teamSize: s.teamSize ?? "",
    modalitaFornitura: MODALITA_FORNITURA.some(m => m.id === s.modalitaFornitura) ? s.modalitaFornitura : "unico",
    quotaPercentuale: s.quotaPercentuale ?? "",
    partnerRTI: s.partnerRTI || "",
    attivitaAutonoma: !!s.attivitaAutonoma,
    contestoOrganizzativo: s.contestoOrganizzativo || "",
    metodologia: s.metodologia || "",
    kpiDettagliati: s.kpiDettagliati || "",
    elementiUnicita: s.elementiUnicita || "",
    referenteContatto: s.referenteContatto || "",
    noteAggiuntive: s.noteAggiuntive || "",
    codiciProgetto: Array.isArray(s.codiciProgetto) ? s.codiciProgetto.filter(Boolean) : []
  });
  const runAiOnText = async () => {
    if (!rawText.trim()) {
      showToast("Scrivi o detta prima una descrizione");
      return;
    }
    setBusy(true);
    setPreview(null);
    try {
      const s = await aiStructure(rawText);
      setPreview(previewDaStruttura(s, addMode === "voce" ? "voce" : "testo"));
    } catch (e) {
      showToast("Strutturazione AI non riuscita, riprova");
    }
    setBusy(false);
  };
  const savePreview = async () => {
    await persistRefs([preview, ...refs]);
    setPreview(null);
    setRawText("");
    setFileName("");
    setView("repo");
    showToast("Referenza salvata nel repository");
  };

  // ── export (rispetta l'anonimizzazione NDA) ──
  const exportXlsx = () => {
    const rows = filtered.map(r => {
      const t = testiVisualizzati(r, clientsById);
      return {
        Titolo: t.titolo,
        Cliente: t.cliente,
        Settore: r.settore,
        Paese: r.paese || "Italia",
        Periodo: fmtPeriodo(r),
        "Durata (mesi)": calcolaDurataMesi(r) ?? "",
        "Importo (k€)": r.importoKEuro,
        Ambiti: r.ambiti.join(", "),
        Tecnologie: r.tecnologie.join(", "),
        Descrizione: t.descrizione,
        Attività: t.attivita,
        Risultati: t.risultati,
        Ruolo: r.ruolo,
        Team: r.teamSize,
        "Modalità di fornitura": etichettaFornitura(r),
        "Partner RTI": r.partnerRTI || "",
        "Contesto organizzativo": t.contestoOrganizzativo || "",
        "Metodologia": r.metodologia || "",
        "KPI dettagliati": t.kpiDettagliati || "",
        "Elementi di unicità": t.elementiUnicita || "",
        "Referente contatto": t.referenteContatto || "",
        "Note aggiuntive": t.noteAggiuntive || "",
        "Codici progetto": (r.codiciProgetto || []).join(", "),
        NDA: r.nda ? "Sì" : "No"
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Referenze");
    XLSX.writeFile(wb, "referenze_progettuali.xlsx");
  };
  const copyForGara = r => {
    const t = testiVisualizzati(r, clientsById);
    const extra = [t.contestoOrganizzativo && `\nContesto organizzativo e criticità di partenza\n${t.contestoOrganizzativo}`, r.metodologia && `\nMetodologia e standard adottati\n${r.metodologia}`, t.kpiDettagliati && `\nKPI e metriche di dettaglio\n${t.kpiDettagliati}`, t.elementiUnicita && `\nElementi di unicità e innovazione\n${t.elementiUnicita}`, t.referenteContatto && `\nReferente cliente contattabile\n${t.referenteContatto}`, t.noteAggiuntive && `\nNote aggiuntive\n${t.noteAggiuntive}`].filter(Boolean).join("\n");
    const txt = `${t.titolo}\nCliente: ${t.cliente} — Settore: ${r.settore} (${r.paese || "Italia"}) — Periodo: ${fmtPeriodo(r)}${calcolaDurataMesi(r) ? " (" + calcolaDurataMesi(r) + " mesi)" : ""} — Importo: ${fmtImporto(r.importoKEuro)}\nModalità di fornitura: ${etichettaFornitura(r)}${r.partnerRTI ? " — Partner RTI: " + r.partnerRTI : ""}\n\nContesto e obiettivi\n${t.descrizione}\n\nAttività svolte\n${t.attivita}\n\nRisultati conseguiti\n${t.risultati}\n\nTecnologie: ${r.tecnologie.join(", ")}${r.ruolo ? "\nRuolo: " + r.ruolo : ""}${r.teamSize ? "\nTeam: " + r.teamSize + " risorse" : ""}${extra}`;
    navigator.clipboard && navigator.clipboard.writeText(txt).then(() => showToast("Testo referenza copiato negli appunti" + (r.nda ? " (versione anonimizzata)" : ""))).catch(() => showToast("Copia non riuscita"));
  };
  const removeRef = async id => {
    const rimosse = refs.filter(r => r.id !== id).map(r => r.parentId === id ? {
      ...r,
      parentId: null,
      tipoReferenza: "standalone"
    } : r);
    await persistRefs(rimosse);
    try {
      await storage.delete(imgKey(id));
    } catch (e) {}
    setSel(null);
    showToast("Referenza eliminata");
  };
  const saveEdit = async () => {
    await persistRefs(refs.map(r => r.id === editDraft.id ? editDraft : r));
    setSel(editDraft);
    setEditDraft(null);
    showToast("Modifiche salvate");
  };

  // ── anagrafica clienti: CRUD ──
  const saveNewClient = async draft => {
    await persistClients([draft, ...clients]);
    showToast("Cliente aggiunto all'anagrafica");
    return draft.id;
  };
  const saveClientEdit = async () => {
    await persistClients(clients.map(c => c.id === clientEditDraft.id ? clientEditDraft : c));
    setClientSel(clientEditDraft);
    setClientEditDraft(null);
    showToast("Anagrafica aggiornata");
  };
  const removeClient = async id => {
    await persistClients(clients.filter(c => c.id !== id));
    // scollega (non elimina) le referenze che puntavano a questo cliente
    const scollegate = refs.map(r => r.clienteId === id ? {
      ...r,
      clienteId: null,
      cliente: clientsById[id]?.nome || r.cliente
    } : r);
    await persistRefs(scollegate);
    setClientSel(null);
    showToast("Cliente eliminato dall'anagrafica");
  };

  // ── filtri repository ──
  const filtered = useMemo(() => {
    const s = q.toLowerCase();
    const annoMin = fAnnoMin !== "" ? Number(fAnnoMin) : null;
    const annoMax = fAnnoMax !== "" ? Number(fAnnoMax) : null;
    const impMin = fImpMin !== "" ? Number(fImpMin) : null;
    const impMax = fImpMax !== "" ? Number(fImpMax) : null;
    return refs.filter(r => {
      const annoInizio = Number(r.inizioAnno) || 0;
      const annoFine = dataFineEffettiva(r).anno || annoInizio;
      const imp = Number(r.importoKEuro) || 0;
      const t = testiVisualizzati(r, clientsById);
      return (!fAmbito || r.ambiti.includes(fAmbito)) && (!fSettore || r.settore === fSettore) && (!fTipo || (r.tipoReferenza || "standalone") === fTipo) && (!fPaese || (r.paese || "Italia") === fPaese) && (annoMin === null || annoFine >= annoMin) && (annoMax === null || annoInizio <= annoMax) && (impMin === null || imp >= impMin) && (impMax === null || imp <= impMax) && (!s || [t.titolo, t.cliente, t.descrizione, t.attivita, r.tecnologie.join(" "), (r.codiciProgetto || []).join(" ")].join(" ").toLowerCase().includes(s));
    });
  }, [refs, clientsById, q, fAmbito, fSettore, fTipo, fPaese, fAnnoMin, fAnnoMax, fImpMin, fImpMax]);
  const resetFiltriRange = () => {
    setFAnnoMin("");
    setFAnnoMax("");
    setFImpMin("");
    setFImpMax("");
  };
  const filtriRangeAttivi = fAnnoMin !== "" || fAnnoMax !== "" || fImpMin !== "" || fImpMax !== "";
  const totKEuro = refs.reduce((a, r) => a + (Number(r.importoKEuro) || 0), 0);
  const numClientiUnici = new Set(refs.map(r => r.clienteId || r.cliente)).size;
  const MAX_REFS_CONTESTO = 40;
  const contestoReferenze = lista => lista.slice(0, MAX_REFS_CONTESTO).map(r => {
    const t = testiVisualizzati(r, clientsById);
    const parent = r.parentId ? refs.find(x => x.id === r.parentId) : null;
    return {
      titolo: t.titolo,
      cliente: t.cliente,
      settore: r.settore,
      paese: r.paese || "Italia",
      periodo: fmtPeriodo(r),
      durataMesi: calcolaDurataMesi(r),
      importoKEuro: r.importoKEuro,
      ambiti: r.ambiti,
      tecnologie: r.tecnologie,
      modalitaFornitura: etichettaFornitura(r),
      descrizione: t.descrizione,
      attivita: t.attivita,
      risultati: t.risultati,
      ruolo: r.ruolo,
      teamSize: r.teamSize,
      contestoOrganizzativo: t.contestoOrganizzativo,
      metodologia: r.metodologia,
      kpiDettagliati: t.kpiDettagliati,
      elementiUnicita: t.elementiUnicita,
      noteAggiuntive: t.noteAggiuntive,
      tipoReferenza: r.tipoReferenza || "standalone",
      programmaOServizioDiRiferimento: parent ? parent.titolo : null,
      nda: r.nda
    };
  });
  const eseguiAssistente = async () => {
    if (!aiPrompt.trim()) {
      showToast("Scrivi prima una richiesta");
      return;
    }
    setAiRunning(true);
    setAiErr("");
    setAiResult("");
    try {
      const lista = aiScope === "filtrate" ? filtered : refs;
      const archivio = contestoReferenze(lista);
      const risposta = await aiAssistente(aiPrompt, archivio);
      setAiResult(risposta);
    } catch (e) {
      setAiErr("Non sono riuscito a elaborare la richiesta. Riprova.");
    }
    setAiRunning(false);
  };
  const copiaRisultatoAI = () => {
    navigator.clipboard && navigator.clipboard.writeText(aiResult).then(() => showToast("Risultato copiato negli appunti")).catch(() => showToast("Copia non riuscita"));
  };

  // ─────────────────────────────────────────── UI ──
  return /*#__PURE__*/React.createElement("div", {
    style: st.page
  }, /*#__PURE__*/React.createElement("style", null, css), /*#__PURE__*/React.createElement("header", {
    style: st.header
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: st.eyebrow
  }, "Practice Data, AI & RPA · Business Development"), /*#__PURE__*/React.createElement("h1", {
    style: st.h1
  }, "Repository Referenze Progettuali")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 20
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: st.stats
  }, /*#__PURE__*/React.createElement(Stat, {
    n: refs.length,
    l: "referenze"
  }), /*#__PURE__*/React.createElement(Stat, {
    n: fmtImporto(totKEuro),
    l: "valore complessivo"
  }), /*#__PURE__*/React.createElement(Stat, {
    n: numClientiUnici,
    l: "clienti"
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      textAlign: "right"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      color: "#F3EBD8"
    }
  }, userEmail), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      color: "#8FA1C4"
    }
  }, gruppo || "—"), /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    style: {
      marginTop: 6,
      padding: "5px 10px",
      fontSize: 12
    },
    onClick: onLogout
  }, "Esci")))), /*#__PURE__*/React.createElement("nav", {
    style: st.nav
  }, /*#__PURE__*/React.createElement("button", {
    className: "tab" + (view === "repo" ? " on" : ""),
    onClick: () => setView("repo")
  }, "Repository"), /*#__PURE__*/React.createElement("button", {
    className: "tab" + (view === "add" ? " on" : ""),
    onClick: () => {
      setView("add");
      setPreview(null);
    }
  }, "+ Nuova referenza"), /*#__PURE__*/React.createElement("button", {
    className: "tab" + (view === "clienti" ? " on" : ""),
    onClick: () => setView("clienti")
  }, "Anagrafica clienti"), /*#__PURE__*/React.createElement("button", {
    className: "tab" + (view === "assistente" ? " on" : ""),
    onClick: () => setView("assistente")
  }, "Assistente AI"), view === "repo" && refs.length > 0 && /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    style: {
      marginLeft: "auto"
    },
    onClick: exportXlsx
  }, "Esporta Excel")), view === "repo" && /*#__PURE__*/React.createElement("section", null, /*#__PURE__*/React.createElement("div", {
    style: st.filters
  }, /*#__PURE__*/React.createElement("input", {
    placeholder: "Cerca per cliente, tecnologia, parola chiave…",
    value: q,
    onChange: e => setQ(e.target.value),
    style: st.search
  }), /*#__PURE__*/React.createElement("select", {
    value: fAmbito,
    onChange: e => setFAmbito(e.target.value),
    style: st.select
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "Tutti gli ambiti"), AMBITI.map(a => /*#__PURE__*/React.createElement("option", {
    key: a
  }, a))), /*#__PURE__*/React.createElement("select", {
    value: fSettore,
    onChange: e => setFSettore(e.target.value),
    style: st.select
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "Tutti i settori"), SETTORI.map(s => /*#__PURE__*/React.createElement("option", {
    key: s
  }, s))), /*#__PURE__*/React.createElement("select", {
    value: fTipo,
    onChange: e => setFTipo(e.target.value),
    style: st.select
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "Tutti i tipi"), TIPO_REFERENZA.map(tp => /*#__PURE__*/React.createElement("option", {
    key: tp.id,
    value: tp.id
  }, tp.label))), /*#__PURE__*/React.createElement("select", {
    value: fPaese,
    onChange: e => setFPaese(e.target.value),
    style: st.select
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "Tutti i paesi"), Array.from(new Set(refs.map(r => r.paese || "Italia"))).sort().map(p => /*#__PURE__*/React.createElement("option", {
    key: p
  }, p)))), /*#__PURE__*/React.createElement("div", {
    style: st.filtersRange
  }, /*#__PURE__*/React.createElement("span", {
    style: st.rangeLabel
  }, "Anno"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    placeholder: "da",
    value: fAnnoMin,
    onChange: e => setFAnnoMin(e.target.value),
    style: st.rangeInput
  }), /*#__PURE__*/React.createElement("span", {
    style: st.rangeDash
  }, "–"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    placeholder: "a",
    value: fAnnoMax,
    onChange: e => setFAnnoMax(e.target.value),
    style: st.rangeInput
  }), /*#__PURE__*/React.createElement("span", {
    style: {
      ...st.rangeLabel,
      marginLeft: 14
    }
  }, "Importo (k€)"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    placeholder: "min",
    value: fImpMin,
    onChange: e => setFImpMin(e.target.value),
    style: st.rangeInput
  }), /*#__PURE__*/React.createElement("span", {
    style: st.rangeDash
  }, "–"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    placeholder: "max",
    value: fImpMax,
    onChange: e => setFImpMax(e.target.value),
    style: st.rangeInput
  }), filtriRangeAttivi && /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    style: st.rangeReset,
    onClick: resetFiltriRange
  }, "Azzera range")), !loaded ? /*#__PURE__*/React.createElement("p", {
    style: st.empty
  }, "Caricamento archivio…") : filtered.length === 0 ? /*#__PURE__*/React.createElement("div", {
    style: st.empty
  }, refs.length === 0 ? /*#__PURE__*/React.createElement(React.Fragment, null, "L'archivio è vuoto. Aggiungi la prima referenza da file, testo o dettatura vocale.") : /*#__PURE__*/React.createElement(React.Fragment, null, "Nessuna referenza corrisponde ai filtri.")) : /*#__PURE__*/React.createElement("div", {
    style: st.grid
  }, filtered.map(r => {
    const t = testiVisualizzati(r, clientsById);
    return /*#__PURE__*/React.createElement("article", {
      key: r.id,
      className: "card",
      onClick: () => setSel(r)
    }, /*#__PURE__*/React.createElement("div", {
      className: "spine"
    }, /*#__PURE__*/React.createElement("span", {
      className: "periodo"
    }, fmtPeriodoSlash(r)), /*#__PURE__*/React.createElement("span", {
      className: "imp"
    }, fmtImporto(r.importoKEuro))), /*#__PURE__*/React.createElement("div", {
      style: {
        padding: "14px 16px",
        flex: 1
      }
    }, /*#__PURE__*/React.createElement(RefThumb, {
      refId: r.id,
      loader: caricaImmagini
    }), /*#__PURE__*/React.createElement("div", {
      style: st.cardClient
    }, t.cliente, r.nda && /*#__PURE__*/React.createElement("span", {
      className: "ndaTag"
    }, "NDA"), r.tipoReferenza === "programma" && /*#__PURE__*/React.createElement("span", {
      className: "aqTag"
    }, "PROGRAMMA"), r.tipoReferenza === "servizio_ams" && /*#__PURE__*/React.createElement("span", {
      className: "amsTag"
    }, "SERVIZIO AMS")), /*#__PURE__*/React.createElement("h3", {
      style: st.cardTitle
    }, t.titolo), r.tipoReferenza === "progetto" && (() => {
      const parent = refs.find(x => x.id === r.parentId);
      return parent ? /*#__PURE__*/React.createElement("div", {
        style: st.streamOf,
        onClick: e => {
          e.stopPropagation();
          setSel(parent);
        }
      }, "↳ Progetto del ", parent.tipoReferenza === "servizio_ams" ? "Servizio AMS" : "Programma", ": ", parent.titolo) : null;
    })(), /*#__PURE__*/React.createElement("p", {
      style: st.cardDesc
    }, t.descrizione), /*#__PURE__*/React.createElement("div", {
      style: st.tagRow
    }, r.ambiti.map(a => /*#__PURE__*/React.createElement("span", {
      key: a,
      className: "pill",
      style: {
        background: AMBITO_COLOR[a]
      }
    }, a)), /*#__PURE__*/React.createElement("span", {
      style: st.settoreTag
    }, r.settore, " · ", r.paese || "Italia")), /*#__PURE__*/React.createElement("div", {
      style: st.settoreTag
    }, fmtPeriodo(r), calcolaDurataMesi(r) ? ` · ${calcolaDurataMesi(r)} mesi` : ""), /*#__PURE__*/React.createElement("div", {
      style: st.fornituraTag
    }, etichettaFornitura(r))));
  }))), view === "add" && /*#__PURE__*/React.createElement("section", {
    style: {
      maxWidth: 760
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: st.modeRow
  }, [["testo", "Testo"], ["voce", "Voce"], ["file", "File"], ["manuale", "Manuale"]].map(([m, l]) => /*#__PURE__*/React.createElement("button", {
    key: m,
    className: "mode" + (addMode === m ? " on" : ""),
    onClick: () => {
      setAddMode(m);
      setPreview(m === "manuale" ? {
        ...emptyRef()
      } : null);
    }
  }, l))), (addMode === "testo" || addMode === "voce") && !preview && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: st.hint
  }, addMode === "voce" ? "Detta a voce la referenza: cliente, periodo, importo, attività, tecnologie e risultati. L'AI la struttura in formato gara." : "Descrivi liberamente il progetto, anche in modo informale. L'AI genera la scheda referenza in italiano formale da offerta."), /*#__PURE__*/React.createElement("textarea", {
    value: rawText,
    onChange: e => setRawText(e.target.value),
    placeholder: "Es. Progetto per Enermesh S.p.A. 2024, piattaforma dati su Azure, 18 mesi, circa 900 mila euro, team di 12, migrazione DWH e modelli predittivi manutenzione…",
    style: st.textarea,
    rows: 8
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 12,
      flexWrap: "wrap",
      padding: "0 22px"
    }
  }, addMode === "voce" && /*#__PURE__*/React.createElement("button", {
    className: "rec" + (recording ? " live" : ""),
    onClick: toggleRec
  }, recording ? "■ Ferma dettatura" : "● Avvia dettatura"), /*#__PURE__*/React.createElement("button", {
    className: "primary",
    onClick: runAiOnText,
    disabled: busy
  }, busy ? "Strutturazione in corso…" : "Struttura con AI"))), addMode === "file" && !preview && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: st.hint
  }, "Carica un documento della referenza (scheda, slide esportata in PDF, estratto d'offerta). Formati: PDF, DOCX, XLSX, CSV, TXT, immagini."), /*#__PURE__*/React.createElement("div", {
    className: "drop",
    onClick: () => fileRef.current && fileRef.current.click()
  }, busy ? "Analisi del documento in corso…" : fileName ? fileName : "Tocca per scegliere un file"), /*#__PURE__*/React.createElement("input", {
    ref: fileRef,
    type: "file",
    accept: ".pdf,.docx,.xlsx,.xls,.csv,.txt,.md,.png,.jpg,.jpeg,.webp",
    style: {
      display: "none"
    },
    onChange: e => e.target.files[0] && handleFile(e.target.files[0])
  }), /*#__PURE__*/React.createElement("textarea", {
    value: rawText,
    onChange: e => setRawText(e.target.value),
    placeholder: "Note aggiuntive facoltative (es. importo effettivo, nome da usare per il cliente)…",
    style: {
      ...st.textarea,
      marginTop: 12
    },
    rows: 2
  }), fileErr && /*#__PURE__*/React.createElement("p", {
    style: st.err
  }, fileErr)), preview && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "0 22px"
    }
  }, /*#__PURE__*/React.createElement(RefEditor, {
    value: preview,
    onChange: setPreview,
    clients: clients,
    onCreateClient: saveNewClient,
    loadImages: caricaImmagini,
    saveImages: salvaImmagini,
    tutteLeReferenze: refs,
    onSave: savePreview,
    onCancel: () => {
      setPreview(null);
      if (addMode === "manuale") setAddMode("testo");
    },
    saveLabel: "Salva nel repository"
  }))), view === "clienti" && /*#__PURE__*/React.createElement("section", null, /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "0 22px 14px"
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "primary",
    onClick: () => setClientEditDraft(emptyClient())
  }, "+ Nuovo cliente")), clients.length === 0 ? /*#__PURE__*/React.createElement("div", {
    style: st.empty
  }, "Nessun cliente in anagrafica. Aggiungine uno per iniziare.") : /*#__PURE__*/React.createElement("div", {
    style: st.grid
  }, clients.map(c => {
    const nRef = refs.filter(r => r.clienteId === c.id).length;
    return /*#__PURE__*/React.createElement("article", {
      key: c.id,
      className: "clientCard",
      onClick: () => setClientSel(c)
    }, /*#__PURE__*/React.createElement(ClientLogo, {
      clientId: c.id,
      loader: caricaLogo,
      style: {
        marginBottom: 8
      }
    }), /*#__PURE__*/React.createElement("div", {
      style: st.cardClient
    }, c.nome, c.nda && /*#__PURE__*/React.createElement("span", {
      className: "ndaTag"
    }, "NDA")), /*#__PURE__*/React.createElement("p", {
      style: st.cardDesc
    }, c.descrizione || "Nessuna descrizione."), /*#__PURE__*/React.createElement("div", {
      style: st.tagRow
    }, /*#__PURE__*/React.createElement("span", {
      style: st.settoreTag
    }, "Etichetta anonima: “", c.labelBreve || "—", "”")), /*#__PURE__*/React.createElement("div", {
      style: {
        fontSize: 12,
        color: "#7C8AA6",
        marginTop: 8
      }
    }, nRef, " referenza/e collegata/e"));
  }))), view === "assistente" && /*#__PURE__*/React.createElement("section", {
    style: {
      maxWidth: 780
    }
  }, /*#__PURE__*/React.createElement("p", {
    style: st.hint
  }, "Scrivi una richiesta libera in linguaggio naturale: puoi chiedere analisi e valutazioni sull'archivio (es. copertura per un requisito di gara, punti di forza), oppure chiedere di confezionare una o più referenze secondo un template specifico, in formato sintetico o super dettagliato, anche in un'altra lingua."), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      flexWrap: "wrap",
      padding: "0 22px 10px"
    }
  }, ["Valuta la copertura delle nostre referenze per una gara PA che richiede esperienza RPA negli ultimi 3 anni", "Confeziona in formato sintetico, in inglese, le referenze più adatte per una gara sul settore energia", "Elenca le referenze in RTI con quota di partecipazione superiore al 40%", "Scrivi una versione super dettagliata della referenza Enermesh per una qualifica ANAC"].map(esempio => /*#__PURE__*/React.createElement("button", {
    key: esempio,
    className: "ghost",
    style: {
      fontSize: 12,
      padding: "6px 10px"
    },
    onClick: () => setAiPrompt(esempio)
  }, esempio))), /*#__PURE__*/React.createElement("textarea", {
    className: "inp",
    style: {
      margin: "0 22px",
      width: "calc(100% - 44px)"
    },
    rows: 5,
    placeholder: "Es. Confeziona in formato sintetico, in spagnolo, le referenze Data & AI in ambito energia per la gara XYZ…",
    value: aiPrompt,
    onChange: e => setAiPrompt(e.target.value)
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 14,
      alignItems: "center",
      flexWrap: "wrap",
      padding: "10px 22px"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      color: "#5A6478"
    }
  }, "Referenze da considerare:"), /*#__PURE__*/React.createElement("label", {
    style: {
      fontSize: 13,
      display: "flex",
      alignItems: "center",
      gap: 5
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "radio",
    name: "aiScope",
    checked: aiScope === "tutte",
    onChange: () => setAiScope("tutte")
  }), "Tutto l'archivio (", refs.length, ")"), /*#__PURE__*/React.createElement("label", {
    style: {
      fontSize: 13,
      display: "flex",
      alignItems: "center",
      gap: 5
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "radio",
    name: "aiScope",
    checked: aiScope === "filtrate",
    onChange: () => setAiScope("filtrate")
  }), "Solo quelle filtrate in Repository (", filtered.length, ")")), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "0 22px"
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "primary",
    onClick: eseguiAssistente,
    disabled: aiRunning
  }, aiRunning ? "Elaborazione in corso…" : "Genera risposta"), refs.length > MAX_REFS_CONTESTO && /*#__PURE__*/React.createElement("span", {
    style: {
      marginLeft: 10,
      fontSize: 12,
      color: "#7C8AA6"
    }
  }, "Nota: vengono considerate al massimo ", MAX_REFS_CONTESTO, " referenze per richiesta.")), aiErr && /*#__PURE__*/React.createElement("p", {
    style: st.err
  }, aiErr), aiResult && /*#__PURE__*/React.createElement("div", {
    style: {
      margin: "16px 22px 0",
      background: "#FBF8F1",
      border: "1px solid #DDD6C6",
      borderRadius: 12,
      padding: "18px 20px"
    }
  }, /*#__PURE__*/React.createElement(MarkdownLite, {
    text: aiResult
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    onClick: copiaRisultatoAI
  }, "Copia risultato")))), sel && (() => {
    const t = testiVisualizzati(sel, clientsById);
    return /*#__PURE__*/React.createElement("div", {
      style: st.overlay,
      onClick: () => {
        setSel(null);
        setEditDraft(null);
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: st.modal,
      onClick: e => e.stopPropagation()
    }, !editDraft ? /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
      style: st.eyebrow
    }, t.cliente, " · ", sel.settore, " · ", sel.paese || "Italia", " · ", fmtPeriodo(sel), sel.nda && /*#__PURE__*/React.createElement("span", {
      className: "ndaTag",
      style: {
        marginLeft: 8
      }
    }, "NDA"), sel.tipoReferenza === "programma" && /*#__PURE__*/React.createElement("span", {
      className: "aqTag",
      style: {
        marginLeft: 8
      }
    }, "PROGRAMMA"), sel.tipoReferenza === "servizio_ams" && /*#__PURE__*/React.createElement("span", {
      className: "amsTag",
      style: {
        marginLeft: 8
      }
    }, "SERVIZIO AMS")), /*#__PURE__*/React.createElement("h2", {
      style: {
        ...st.h1,
        fontSize: 22,
        margin: "6px 0 4px",
        color: "#22304A"
      }
    }, t.titolo), sel.tipoReferenza === "progetto" && (() => {
      const parent = refs.find(x => x.id === sel.parentId);
      return parent ? /*#__PURE__*/React.createElement("div", {
        style: {
          ...st.streamOf,
          marginBottom: 10
        },
        onClick: () => setSel(parent)
      }, "↳ Progetto del ", parent.tipoReferenza === "servizio_ams" ? "Servizio AMS" : "Programma", ": ", parent.titolo) : null;
    })(), /*#__PURE__*/React.createElement("div", {
      style: {
        ...st.tagRow,
        marginBottom: 14
      }
    }, sel.ambiti.map(a => /*#__PURE__*/React.createElement("span", {
      key: a,
      className: "pill",
      style: {
        background: AMBITO_COLOR[a]
      }
    }, a)), /*#__PURE__*/React.createElement("span", {
      style: st.settoreTag
    }, fmtImporto(sel.importoKEuro), calcolaDurataMesi(sel) ? " · " + calcolaDurataMesi(sel) + " mesi" : "", sel.teamSize ? " · team " + sel.teamSize : "", sel.inCorso ? " · in corso" : "")), /*#__PURE__*/React.createElement(Block, {
      t: "Modalità di fornitura",
      v: etichettaFornitura(sel) + (sel.partnerRTI ? ` — Partner RTI: ${sel.partnerRTI}` : "")
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Contesto e obiettivi",
      v: t.descrizione
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Attività svolte",
      v: t.attivita
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Risultati conseguiti",
      v: t.risultati
    }), sel.ruolo ? /*#__PURE__*/React.createElement(Block, {
      t: "Ruolo",
      v: sel.ruolo
    }) : null, sel.tecnologie.length > 0 && /*#__PURE__*/React.createElement(Block, {
      t: "Tecnologie",
      v: sel.tecnologie.join(", ")
    }), sel.codiciProgetto && sel.codiciProgetto.length > 0 && /*#__PURE__*/React.createElement(Block, {
      t: "Codici progetto / commessa / contratto",
      v: sel.codiciProgetto.join(", ")
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Contesto organizzativo e criticità di partenza",
      v: t.contestoOrganizzativo
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Metodologia e standard adottati",
      v: sel.metodologia
    }), /*#__PURE__*/React.createElement(Block, {
      t: "KPI e metriche di dettaglio",
      v: t.kpiDettagliati
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Elementi di unicità e innovazione",
      v: t.elementiUnicita
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Referente cliente contattabile",
      v: t.referenteContatto
    }), /*#__PURE__*/React.createElement(Block, {
      t: "Note aggiuntive",
      v: t.noteAggiuntive
    }), puoEssereGenitore(sel.tipoReferenza) && (() => {
      const figli = refs.filter(x => x.parentId === sel.id);
      const etichettaFigli = sel.tipoReferenza === "servizio_ams" ? "progetti evolutivi" : "progetti";
      if (figli.length === 0) return /*#__PURE__*/React.createElement("p", {
        style: {
          fontSize: 13,
          color: "#7C8AA6",
          marginBottom: 12
        }
      }, "Nessun progetto ancora collegato a questo ", sel.tipoReferenza === "servizio_ams" ? "servizio AMS" : "programma", ".");
      const totFigli = figli.reduce((a, f) => a + (Number(f.importoKEuro) || 0), 0);
      return /*#__PURE__*/React.createElement("div", {
        style: {
          marginBottom: 12
        }
      }, /*#__PURE__*/React.createElement("div", {
        style: {
          fontSize: 11,
          letterSpacing: ".1em",
          textTransform: "uppercase",
          color: "#7C8AA6",
          marginBottom: 6
        }
      }, etichettaFigli.charAt(0).toUpperCase() + etichettaFigli.slice(1), " collegati (", figli.length, ") · valore attivato: ", fmtImporto(totFigli)), /*#__PURE__*/React.createElement("div", {
        style: {
          display: "flex",
          flexDirection: "column",
          gap: 6
        }
      }, figli.map(f => /*#__PURE__*/React.createElement("div", {
        key: f.id,
        onClick: () => setSel(f),
        style: {
          padding: "8px 10px",
          background: "#F0ECDF",
          borderRadius: 8,
          cursor: "pointer",
          fontSize: 13
        }
      }, /*#__PURE__*/React.createElement("strong", null, f.titolo), " — ", fmtPeriodo(f), " · ", fmtImporto(f.importoKEuro)))));
    })(), /*#__PURE__*/React.createElement(RefGallery, {
      refId: sel.id,
      loader: caricaImmagini
    }), /*#__PURE__*/React.createElement("div", {
      style: {
        display: "flex",
        gap: 10,
        marginTop: 18,
        flexWrap: "wrap"
      }
    }, /*#__PURE__*/React.createElement("button", {
      className: "primary",
      onClick: () => copyForGara(sel)
    }, "Copia testo per gara"), /*#__PURE__*/React.createElement("button", {
      className: "ghost",
      onClick: () => setEditDraft({
        ...sel
      })
    }, "Modifica"), /*#__PURE__*/React.createElement("button", {
      className: "danger",
      onClick: () => removeRef(sel.id)
    }, "Elimina"), /*#__PURE__*/React.createElement("button", {
      className: "ghost",
      style: {
        marginLeft: "auto"
      },
      onClick: () => setSel(null)
    }, "Chiudi"))) : /*#__PURE__*/React.createElement(RefEditor, {
      value: editDraft,
      onChange: setEditDraft,
      clients: clients,
      onCreateClient: saveNewClient,
      loadImages: caricaImmagini,
      saveImages: salvaImmagini,
      tutteLeReferenze: refs,
      onSave: saveEdit,
      onCancel: () => setEditDraft(null),
      saveLabel: "Salva modifiche"
    })));
  })(), clientSel && /*#__PURE__*/React.createElement("div", {
    style: st.overlay,
    onClick: () => {
      setClientSel(null);
      setClientEditDraft(null);
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: st.modal,
    onClick: e => e.stopPropagation()
  }, !clientEditDraft ? /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: st.eyebrow
  }, "Anagrafica cliente"), /*#__PURE__*/React.createElement(ClientLogo, {
    clientId: clientSel.id,
    loader: caricaLogo,
    size: 64,
    style: {
      margin: "8px 0"
    }
  }), /*#__PURE__*/React.createElement("h2", {
    style: {
      ...st.h1,
      fontSize: 22,
      margin: "6px 0 4px",
      color: "#22304A"
    }
  }, clientSel.nome, clientSel.nda && /*#__PURE__*/React.createElement("span", {
    className: "ndaTag",
    style: {
      marginLeft: 8
    }
  }, "NDA di default")), /*#__PURE__*/React.createElement(Block, {
    t: "Descrizione",
    v: clientSel.descrizione
  }), /*#__PURE__*/React.createElement(Block, {
    t: "Etichetta anonima (usata al posto del nome quando la referenza è NDA)",
    v: clientSel.labelBreve
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 18,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    onClick: () => setClientEditDraft({
      ...clientSel
    })
  }, "Modifica"), /*#__PURE__*/React.createElement("button", {
    className: "danger",
    onClick: () => removeClient(clientSel.id)
  }, "Elimina dall'anagrafica"), /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    style: {
      marginLeft: "auto"
    },
    onClick: () => setClientSel(null)
  }, "Chiudi"))) : /*#__PURE__*/React.createElement(ClientEditor, {
    value: clientEditDraft,
    onChange: setClientEditDraft,
    loadLogo: caricaLogo,
    saveLogo: salvaLogo,
    onSave: async () => {
      if (clients.some(c => c.id === clientEditDraft.id)) await saveClientEdit();else {
        await saveNewClient(clientEditDraft);
        setClientEditDraft(null);
      }
    },
    onCancel: () => setClientEditDraft(null),
    saveLabel: clients.some(c => c.id === clientEditDraft.id) ? "Salva modifiche" : "Aggiungi cliente"
  }))), toast && /*#__PURE__*/React.createElement("div", {
    style: st.toast
  }, toast));
}

// ── componenti ────────────────────────────────────────────────
function Stat({
  n,
  l
}) {
  return /*#__PURE__*/React.createElement("div", {
    style: {
      textAlign: "right"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "'IBM Plex Mono', monospace",
      fontSize: 20,
      fontWeight: 600,
      color: "#F3EBD8"
    }
  }, n), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#8FA1C4"
    }
  }, l));
}
function Block({
  t,
  v
}) {
  if (!v) return null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".1em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 3
    }
  }, t), /*#__PURE__*/React.createElement("p", {
    style: {
      margin: 0,
      lineHeight: 1.55,
      color: "#22304A",
      whiteSpace: "pre-wrap"
    }
  }, v));
}

// formattazione inline: **grassetto**
const inlineFmt = testo => testo.split(/(\*\*[^*]+\*\*)/g).map((p, i) => p.startsWith("**") && p.endsWith("**") ? /*#__PURE__*/React.createElement("strong", {
  key: i
}, p.slice(2, -2)) : /*#__PURE__*/React.createElement(React.Fragment, {
  key: i
}, p));

// rendering leggero di markdown (##, -, **) per le risposte testuali dell'AI
function MarkdownLite({
  text
}) {
  if (!text) return null;
  const righe = text.split("\n");
  const elementi = [];
  let elenco = [];
  const chiudiElenco = () => {
    if (elenco.length > 0) {
      elementi.push(/*#__PURE__*/React.createElement("ul", {
        key: "ul" + elementi.length,
        style: {
          margin: "4px 0 12px 18px",
          padding: 0
        }
      }, elenco.map((li, i) => /*#__PURE__*/React.createElement("li", {
        key: i,
        style: {
          marginBottom: 4,
          lineHeight: 1.5
        }
      }, inlineFmt(li)))));
      elenco = [];
    }
  };
  righe.forEach((riga, i) => {
    const t = riga.trim();
    if (/^#{1,3}\s+/.test(t)) {
      chiudiElenco();
      const livello = t.match(/^#+/)[0].length;
      const testo = t.replace(/^#{1,3}\s+/, "");
      elementi.push(/*#__PURE__*/React.createElement("div", {
        key: i,
        style: {
          fontWeight: 700,
          fontSize: livello === 1 ? 16 : 14,
          marginTop: i === 0 ? 0 : 14,
          marginBottom: 5,
          color: "#0E1F3D"
        }
      }, inlineFmt(testo)));
    } else if (/^[-*]\s+/.test(t)) {
      elenco.push(t.replace(/^[-*]\s+/, ""));
    } else if (t === "") {
      chiudiElenco();
    } else {
      chiudiElenco();
      elementi.push(/*#__PURE__*/React.createElement("p", {
        key: i,
        style: {
          margin: "0 0 9px",
          lineHeight: 1.55,
          color: "#22304A"
        }
      }, inlineFmt(t)));
    }
  });
  chiudiElenco();
  return /*#__PURE__*/React.createElement("div", null, elementi);
}

// miniatura caricata pigramente per ogni card della repository
function RefThumb({
  refId,
  loader
}) {
  const [img, setImg] = useState(undefined); // undefined = in caricamento, null = nessuna immagine
  useEffect(() => {
    let vivo = true;
    loader(refId).then(arr => {
      if (vivo) setImg(arr && arr.length > 0 ? arr[0] : null);
    });
    return () => {
      vivo = false;
    };
  }, [refId]);
  if (!img) return null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      margin: "-14px -16px 10px",
      overflow: "hidden",
      borderRadius: "0",
      height: 130
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: img.dataUrl,
    alt: img.didascalia || "Immagine referenza",
    style: {
      width: "100%",
      height: "100%",
      objectFit: "cover",
      display: "block"
    }
  }));
}

// logo cliente, caricato pigramente su sfondo a scacchiera per mostrare la trasparenza
function ClientLogo({
  clientId,
  loader,
  size,
  style
}) {
  const [logo, setLogo] = useState(undefined); // undefined = in caricamento, null = assente
  useEffect(() => {
    let vivo = true;
    setLogo(undefined);
    loader(clientId).then(d => {
      if (vivo) setLogo(d);
    });
    return () => {
      vivo = false;
    };
  }, [clientId]);
  if (!logo) return null;
  const h = size || 44;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      height: h,
      display: "flex",
      alignItems: "center",
      ...style
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: logo,
    alt: "Logo cliente",
    style: {
      maxHeight: h,
      maxWidth: h * 3.2,
      objectFit: "contain",
      display: "block"
    }
  }));
}

// galleria completa con lightbox, usata nel dettaglio referenza
function RefGallery({
  refId,
  loader
}) {
  const [imgs, setImgs] = useState(null);
  const [zoom, setZoom] = useState(null);
  useEffect(() => {
    let vivo = true;
    setImgs(null);
    loader(refId).then(arr => {
      if (vivo) setImgs(arr || []);
    });
    return () => {
      vivo = false;
    };
  }, [refId]);
  if (!imgs || imgs.length === 0) return null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".1em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 6
    }
  }, "Immagini e architetture (", imgs.length, ")"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      flexWrap: "wrap"
    }
  }, imgs.map(im => /*#__PURE__*/React.createElement("div", {
    key: im.id,
    onClick: () => setZoom(im),
    style: {
      cursor: "zoom-in"
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: im.dataUrl,
    alt: im.didascalia || "Immagine referenza",
    style: {
      width: 120,
      height: 90,
      objectFit: "cover",
      borderRadius: 8,
      border: "1px solid #DDD6C6",
      display: "block"
    }
  })))), zoom && /*#__PURE__*/React.createElement("div", {
    style: st.lightbox,
    onClick: () => setZoom(null)
  }, /*#__PURE__*/React.createElement("img", {
    src: zoom.dataUrl,
    alt: zoom.didascalia || "",
    style: {
      maxWidth: "92vw",
      maxHeight: "80vh",
      borderRadius: 10
    }
  }), zoom.didascalia && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "#F3EBD8",
      marginTop: 10,
      fontSize: 13
    }
  }, zoom.didascalia)));
}
function Field({
  l,
  children,
  w
}) {
  return /*#__PURE__*/React.createElement("label", {
    style: {
      display: "block",
      flex: w ? "0 0 " + w : "1 1 160px"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, l), children);
}
function ClientEditor({
  value: v,
  onChange,
  onSave,
  onCancel,
  saveLabel,
  loadLogo,
  saveLogo
}) {
  const set = (k, val) => onChange({
    ...v,
    [k]: val
  });
  const [logo, setLogo] = useState(undefined);
  const [logoBusy, setLogoBusy] = useState(false);
  const logoInputRef = useRef(null);
  useEffect(() => {
    let vivo = true;
    loadLogo(v.id).then(d => {
      if (vivo) setLogo(d);
    });
    return () => {
      vivo = false;
    };
  }, [v.id]);

  // rimozione sfondo chiaro lato client via canvas (soglia semplice su pixel quasi bianchi)
  const rimuoviSfondoESalva = file => {
    setLogoBusy(true);
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const maxW = 320;
      const scale = Math.min(1, maxW / img.width);
      const w = Math.round(img.width * scale),
        h = Math.round(img.height * scale);
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);
      const imgData = ctx.getImageData(0, 0, w, h);
      const d = imgData.data;
      const soglia = 240,
        feather = 25;
      for (let i = 0; i < d.length; i += 4) {
        const avg = (d[i] + d[i + 1] + d[i + 2]) / 3;
        if (d[i + 3] === 0) continue;
        if (avg >= soglia) d[i + 3] = 0;else if (avg >= soglia - feather) d[i + 3] = Math.round(d[i + 3] * (soglia - avg) / feather);
      }
      ctx.putImageData(imgData, 0, 0);
      URL.revokeObjectURL(url);
      const dataUrl = canvas.toDataURL("image/png");
      setLogo(dataUrl);
      saveLogo(v.id, dataUrl).finally(() => setLogoBusy(false));
    };
    img.onerror = () => setLogoBusy(false);
    img.src = url;
  };
  const rimuoviLogo = () => {
    setLogo(null);
    saveLogo(v.id, null);
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      color: "#3E8E9C",
      fontWeight: 600,
      marginBottom: 10,
      letterSpacing: ".05em",
      textTransform: "uppercase"
    }
  }, "Scheda cliente"), /*#__PURE__*/React.createElement(Field, {
    l: "Nome cliente (ragione sociale)"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    value: v.nome,
    onChange: e => set("nome", e.target.value)
  })), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Descrizione del cliente"
  }, /*#__PURE__*/React.createElement("textarea", {
    className: "inp",
    rows: 3,
    value: v.descrizione,
    onChange: e => set("descrizione", e.target.value)
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Etichetta anonima (sostituisce il nome nelle referenze NDA)"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    placeholder: "Es. \"Primario operatore energetico nazionale\"",
    value: v.labelBreve,
    onChange: e => set("labelBreve", e.target.value)
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 14
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, "Logo cliente"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 12,
      marginTop: 6
    }
  }, logo ? /*#__PURE__*/React.createElement("div", {
    className: "logoCheck",
    style: {
      height: 56,
      padding: 8,
      borderRadius: 8
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: logo,
    alt: "Logo",
    style: {
      maxHeight: 56,
      maxWidth: 180,
      objectFit: "contain",
      display: "block"
    }
  })) : /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      color: "#7C8AA6"
    }
  }, "Nessun logo caricato"), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "ghost",
    style: {
      padding: "8px 12px",
      fontSize: 12
    },
    onClick: () => logoInputRef.current && logoInputRef.current.click(),
    disabled: logoBusy
  }, logoBusy ? "Elaborazione…" : "Carica / sostituisci"), logo && /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "danger",
    style: {
      padding: "8px 12px",
      fontSize: 12
    },
    onClick: rimuoviLogo
  }, "Rimuovi")), /*#__PURE__*/React.createElement("input", {
    ref: logoInputRef,
    type: "file",
    accept: "image/*",
    style: {
      display: "none"
    },
    onChange: e => e.target.files[0] && rimuoviSfondoESalva(e.target.files[0])
  }), /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 11,
      color: "#7C8AA6",
      marginTop: 4
    }
  }, "Lo sfondo chiaro viene reso automaticamente trasparente al caricamento.")), /*#__PURE__*/React.createElement("label", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginTop: 14,
      fontSize: 13,
      color: "#5A6478"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!v.nda,
    onChange: e => set("nda", e.target.checked)
  }), "Cliente riservato per NDA (nuove referenze di questo cliente saranno anonimizzate di default)"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 16
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "primary",
    onClick: onSave,
    disabled: !v.nome.trim()
  }, saveLabel), /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    onClick: onCancel
  }, "Annulla")));
}
function RefEditor({
  value: v,
  onChange,
  onSave,
  onCancel,
  saveLabel,
  clients,
  onCreateClient,
  loadImages,
  saveImages,
  tutteLeReferenze
}) {
  const set = (k, val) => onChange({
    ...v,
    [k]: val
  });
  const toggleAmbito = a => set("ambiti", v.ambiti.includes(a) ? v.ambiti.filter(x => x !== a) : [...v.ambiti, a]);
  const [nuovoClienteMode, setNuovoClienteMode] = useState(false);
  const [nuovoClienteNome, setNuovoClienteNome] = useState("");
  const [imgs, setImgs] = useState([]);
  const [imgBusy, setImgBusy] = useState(false);
  const [imgErr, setImgErr] = useState("");
  const [savingImgs, setSavingImgs] = useState(false);
  const imgInputRef = useRef(null);
  const [altreInfoAperte, setAltreInfoAperte] = useState(!!(v.contestoOrganizzativo || v.metodologia || v.kpiDettagliati || v.elementiUnicita || v.referenteContatto || v.noteAggiuntive));
  useEffect(() => {
    let vivo = true;
    loadImages(v.id).then(arr => {
      if (vivo) setImgs(arr || []);
    });
    return () => {
      vivo = false;
    };
  }, [v.id]);
  const aggiungiImmagini = async fileList => {
    setImgErr("");
    setImgBusy(true);
    try {
      const nuove = [];
      for (const file of Array.from(fileList)) {
        if (!file.type.startsWith("image/")) continue;
        const dataUrl = await comprimiImmagine(file);
        nuove.push({
          id: "img_" + Date.now() + "_" + Math.floor(Math.random() * 9999),
          dataUrl,
          didascalia: ""
        });
      }
      setImgs(p => [...p, ...nuove]);
    } catch (e) {
      setImgErr("Una o più immagini non sono state elaborate");
    }
    setImgBusy(false);
  };
  const rimuoviImmagine = id => setImgs(p => p.filter(im => im.id !== id));
  const setDidascalia = (id, testo) => setImgs(p => p.map(im => im.id === id ? {
    ...im,
    didascalia: testo
  } : im));
  const handleSave = async () => {
    setSavingImgs(true);
    await saveImages(v.id, imgs);
    setSavingImgs(false);
    onSave();
  };
  const clienteScelto = clients.find(c => c.id === v.clienteId);
  const handleSelectCliente = e => {
    const id = e.target.value;
    if (id === "__nuovo__") {
      setNuovoClienteMode(true);
      return;
    }
    const c = clients.find(x => x.id === id);
    onChange({
      ...v,
      clienteId: id || null,
      nda: c ? !!c.nda : v.nda
    });
  };
  const confermaNuovoCliente = async () => {
    if (!nuovoClienteNome.trim()) return;
    const draft = {
      ...emptyClient(),
      nome: nuovoClienteNome.trim()
    };
    const id = await onCreateClient(draft);
    onChange({
      ...v,
      clienteId: id,
      cliente: ""
    });
    setNuovoClienteMode(false);
    setNuovoClienteNome("");
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      color: "#3E8E9C",
      fontWeight: 600,
      marginBottom: 10,
      letterSpacing: ".05em",
      textTransform: "uppercase"
    }
  }, "Scheda referenza — verifica e completa"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Titolo"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    value: v.titolo,
    onChange: e => set("titolo", e.target.value)
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Cliente (anagrafica)"
  }, !nuovoClienteMode ? /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.clienteId || "",
    onChange: handleSelectCliente
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "— ", v.cliente || "seleziona un cliente", " (testo libero) —"), clients.map(c => /*#__PURE__*/React.createElement("option", {
    key: c.id,
    value: c.id
  }, c.nome)), /*#__PURE__*/React.createElement("option", {
    value: "__nuovo__"
  }, "+ Aggiungi nuovo cliente…")) : /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      marginTop: 4
    }
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    placeholder: "Nome nuovo cliente",
    value: nuovoClienteNome,
    onChange: e => setNuovoClienteNome(e.target.value)
  }), /*#__PURE__*/React.createElement("button", {
    className: "primary",
    style: {
      padding: "9px 14px"
    },
    onClick: confermaNuovoCliente
  }, "Crea"), /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    style: {
      padding: "9px 14px"
    },
    onClick: () => setNuovoClienteMode(false)
  }, "Annulla"))), !v.clienteId && !nuovoClienteMode && /*#__PURE__*/React.createElement("input", {
    className: "inp",
    style: {
      marginTop: 6
    },
    placeholder: "Nome cliente (testo libero, non collegato ad anagrafica)",
    value: v.cliente,
    onChange: e => set("cliente", e.target.value)
  })), /*#__PURE__*/React.createElement("label", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginTop: 12,
      fontSize: 13,
      color: "#5A6478"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!v.nda,
    onChange: e => set("nda", e.target.checked)
  }), "Anonimizza questa referenza per NDA (il nome cliente viene sostituito ovunque con l'etichetta anonima)", clienteScelto && !clienteScelto.labelBreve && v.nda && /*#__PURE__*/React.createElement("span", {
    style: {
      color: "#A83A3A"
    }
  }, " — imposta un'etichetta anonima per questo cliente in anagrafica")), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12,
      padding: "12px",
      background: "#F0ECDF",
      borderRadius: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 8
    }
  }, "Tipo di referenza"), /*#__PURE__*/React.createElement(Field, {
    l: "Tipo"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.tipoReferenza,
    onChange: e => onChange({
      ...v,
      tipoReferenza: e.target.value,
      parentId: e.target.value === "progetto" ? v.parentId : null
    })
  }, TIPO_REFERENZA.map(tp => /*#__PURE__*/React.createElement("option", {
    key: tp.id,
    value: tp.id
  }, tp.label)))), v.tipoReferenza === "progetto" && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Programma o Servizio AMS di riferimento"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.parentId || "",
    onChange: e => set("parentId", e.target.value || null)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "— seleziona il programma o servizio AMS —"), (tutteLeReferenze || []).filter(r => puoEssereGenitore(r.tipoReferenza) && r.id !== v.id).map(r => /*#__PURE__*/React.createElement("option", {
    key: r.id,
    value: r.id
  }, r.titolo, " (", etichettaTipoReferenza(r.tipoReferenza), ")")))), (tutteLeReferenze || []).filter(r => puoEssereGenitore(r.tipoReferenza)).length === 0 && /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 12,
      color: "#A83A3A",
      marginTop: 6
    }
  }, "Nessun Programma o Servizio AMS presente in archivio. Crea prima quella referenza di livello superiore, poi collega qui i suoi progetti.")), puoEssereGenitore(v.tipoReferenza) && /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 12,
      color: "#5A6478",
      marginTop: 8
    }
  }, "Dopo aver salvato, potrai collegare a questa referenza i singoli progetti evolutivi creandoli come referenze di tipo \"Progetto\" e selezionandola qui.")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap",
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Settore",
    w: "220px"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.settore,
    onChange: e => set("settore", e.target.value)
  }, SETTORI.map(s => /*#__PURE__*/React.createElement("option", {
    key: s
  }, s)))), /*#__PURE__*/React.createElement(Field, {
    l: "Paese",
    w: "160px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    list: "paesi-noti",
    value: v.paese,
    onChange: e => set("paese", e.target.value)
  }), /*#__PURE__*/React.createElement("datalist", {
    id: "paesi-noti"
  }, PAESI_COMUNI.map(p => /*#__PURE__*/React.createElement("option", {
    key: p,
    value: p
  }))))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12,
      padding: "12px",
      background: "#F0ECDF",
      borderRadius: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 8
    }
  }, "Periodo del progetto"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Mese inizio",
    w: "150px"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.inizioMese,
    onChange: e => set("inizioMese", Number(e.target.value))
  }, MESI.map((m, i) => /*#__PURE__*/React.createElement("option", {
    key: m,
    value: i + 1
  }, m)))), /*#__PURE__*/React.createElement(Field, {
    l: "Anno inizio",
    w: "100px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    type: "number",
    value: v.inizioAnno,
    onChange: e => set("inizioAnno", Number(e.target.value))
  }))), /*#__PURE__*/React.createElement("label", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginTop: 10,
      fontSize: 13,
      color: "#5A6478"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!v.inCorso,
    onChange: e => onChange({
      ...v,
      inCorso: e.target.checked,
      fineMese: e.target.checked ? "" : v.fineMese,
      fineAnno: e.target.checked ? "" : v.fineAnno
    })
  }), "Progetto ancora in corso (la durata si aggiorna automaticamente ad oggi)"), !v.inCorso && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap",
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Mese fine",
    w: "150px"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.fineMese || "",
    onChange: e => set("fineMese", Number(e.target.value))
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "—"), MESI.map((m, i) => /*#__PURE__*/React.createElement("option", {
    key: m,
    value: i + 1
  }, m)))), /*#__PURE__*/React.createElement(Field, {
    l: "Anno fine",
    w: "100px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    type: "number",
    value: v.fineAnno,
    onChange: e => set("fineAnno", Number(e.target.value))
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10,
      fontSize: 13,
      color: "#3E8E9C",
      fontWeight: 600
    }
  }, fmtPeriodo(v), calcolaDurataMesi(v) ? ` · durata calcolata: ${calcolaDurataMesi(v)} mesi` : "")), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap",
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Importo (k€)",
    w: "110px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    type: "number",
    value: v.importoKEuro,
    onChange: e => set("importoKEuro", e.target.value)
  })), /*#__PURE__*/React.createElement(Field, {
    l: "Team",
    w: "80px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    type: "number",
    value: v.teamSize,
    onChange: e => set("teamSize", e.target.value)
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 14,
      padding: "12px",
      background: "#F0ECDF",
      borderRadius: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 8
    }
  }, "Modalità di fornitura"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 12,
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Fornitura",
    w: "230px"
  }, /*#__PURE__*/React.createElement("select", {
    className: "inp",
    value: v.modalitaFornitura,
    onChange: e => set("modalitaFornitura", e.target.value)
  }, MODALITA_FORNITURA.map(m => /*#__PURE__*/React.createElement("option", {
    key: m.id,
    value: m.id
  }, m.label)))), richiedeQuota(v.modalitaFornitura) && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Field, {
    l: "Quota %",
    w: "90px"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    type: "number",
    min: "0",
    max: "100",
    value: v.quotaPercentuale,
    onChange: e => set("quotaPercentuale", e.target.value)
  })), /*#__PURE__*/React.createElement(Field, {
    l: "Partner RTI"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    placeholder: "Es. NTT Data (40%)",
    value: v.partnerRTI,
    onChange: e => set("partnerRTI", e.target.value)
  })))), richiedeQuota(v.modalitaFornitura) && /*#__PURE__*/React.createElement("label", {
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8,
      marginTop: 10,
      fontSize: 13,
      color: "#5A6478"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!v.attivitaAutonoma,
    onChange: e => set("attivitaAutonoma", e.target.checked)
  }), "Pur essendo in RTI, l'attività descritta in questa referenza è stata svolta interamente in autonomia dal nostro team")), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, "Ambiti"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 6,
      flexWrap: "wrap",
      marginTop: 5
    }
  }, AMBITI.map(a => /*#__PURE__*/React.createElement("button", {
    key: a,
    className: "pillBtn",
    style: {
      background: v.ambiti.includes(a) ? AMBITO_COLOR[a] : "#E4E0D4",
      color: v.ambiti.includes(a) ? "#fff" : "#5A6478"
    },
    onClick: () => toggleAmbito(a)
  }, a)))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Tecnologie (separate da virgola)"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    value: v.tecnologie.join(", "),
    onChange: e => set("tecnologie", e.target.value.split(",").map(t => t.trim()).filter(Boolean))
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: "Codici progetto / commessa / contratto (separati da virgola, per il recupero dati di fatturazione)"
  }, /*#__PURE__*/React.createElement("input", {
    className: "inp",
    placeholder: "Es. PRJ-2024-0187, CIG 9876543210",
    value: (v.codiciProgetto || []).join(", "),
    onChange: e => set("codiciProgetto", e.target.value.split(",").map(c => c.trim()).filter(Boolean))
  }))), [["descrizione", "Contesto e obiettivi"], ["attivita", "Attività svolte"], ["risultati", "Risultati conseguiti"], ["ruolo", "Ruolo del fornitore"]].map(([k, l]) => /*#__PURE__*/React.createElement("div", {
    key: k,
    style: {
      marginTop: 12
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: l
  }, /*#__PURE__*/React.createElement("textarea", {
    className: "inp",
    rows: k === "ruolo" ? 1 : 3,
    value: v[k],
    onChange: e => set(k, e.target.value)
  })))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 14
    }
  }, /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "ghost",
    style: {
      fontSize: 13,
      padding: "8px 14px"
    },
    onClick: () => setAltreInfoAperte(p => !p)
  }, altreInfoAperte ? "− Nascondi altre informazioni" : "+ Altre informazioni (contesto, metodologia, KPI, referente…)"), altreInfoAperte && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10
    }
  }, [["contestoOrganizzativo", "Contesto organizzativo e criticità di partenza", 3], ["metodologia", "Metodologia e standard adottati (Agile/Waterfall, framework, certificazioni, compliance)", 2], ["kpiDettagliati", "KPI e metriche di dettaglio", 3], ["elementiUnicita", "Elementi di unicità e innovazione", 2], ["referenteContatto", "Referente cliente contattabile (nome, ruolo, contatti)", 2], ["noteAggiuntive", "Note libere / altre informazioni utili", 3]].map(([k, l, rows]) => /*#__PURE__*/React.createElement("div", {
    key: k,
    style: {
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Field, {
    l: l
  }, /*#__PURE__*/React.createElement("textarea", {
    className: "inp",
    rows: rows,
    value: v[k],
    onChange: e => set(k, e.target.value)
  })))))), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 14
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, "Immagini e architetture"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 8,
      flexWrap: "wrap",
      marginTop: 6
    }
  }, imgs.map(im => /*#__PURE__*/React.createElement("div", {
    key: im.id,
    style: {
      width: 130
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative"
    }
  }, /*#__PURE__*/React.createElement("img", {
    src: im.dataUrl,
    alt: "",
    style: {
      width: 130,
      height: 90,
      objectFit: "cover",
      borderRadius: 8,
      border: "1px solid #C9C2B2",
      display: "block"
    }
  }), /*#__PURE__*/React.createElement("button", {
    type: "button",
    onClick: () => rimuoviImmagine(im.id),
    style: {
      position: "absolute",
      top: 4,
      right: 4,
      background: "rgba(14,31,61,.8)",
      color: "#fff",
      border: "none",
      borderRadius: "50%",
      width: 22,
      height: 22,
      fontSize: 12,
      lineHeight: "22px",
      padding: 0
    }
  }, "×")), /*#__PURE__*/React.createElement("input", {
    className: "inp",
    style: {
      marginTop: 4,
      fontSize: 12,
      padding: "5px 7px"
    },
    placeholder: "Didascalia",
    value: im.didascalia,
    onChange: e => setDidascalia(im.id, e.target.value)
  }))), /*#__PURE__*/React.createElement("div", {
    className: "drop",
    style: {
      margin: 0,
      width: 130,
      height: 90,
      padding: 0,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 12
    },
    onClick: () => imgInputRef.current && imgInputRef.current.click()
  }, imgBusy ? "Caricamento…" : "+ Aggiungi")), /*#__PURE__*/React.createElement("input", {
    ref: imgInputRef,
    type: "file",
    accept: "image/*",
    multiple: true,
    style: {
      display: "none"
    },
    onChange: e => e.target.files.length && aggiungiImmagini(e.target.files)
  }), imgErr && /*#__PURE__*/React.createElement("p", {
    style: st.err
  }, imgErr)), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      gap: 10,
      marginTop: 16
    }
  }, /*#__PURE__*/React.createElement("button", {
    className: "primary",
    onClick: handleSave,
    disabled: savingImgs
  }, savingImgs ? "Salvataggio…" : saveLabel), /*#__PURE__*/React.createElement("button", {
    className: "ghost",
    onClick: onCancel
  }, "Annulla")));
}

// ── stile ─────────────────────────────────────────────────────
const st = {
  page: {
    minHeight: "100vh",
    background: "#EFEAE0",
    fontFamily: "'Libre Franklin', system-ui, sans-serif",
    color: "#22304A",
    padding: "0 0 60px"
  },
  header: {
    background: "#0E1F3D",
    padding: "26px 22px 22px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-end",
    gap: 16,
    flexWrap: "wrap"
  },
  eyebrow: {
    fontSize: 11,
    letterSpacing: ".14em",
    textTransform: "uppercase",
    color: "#8FA1C4"
  },
  h1: {
    fontFamily: "'Archivo', sans-serif",
    fontWeight: 700,
    fontSize: 26,
    color: "#F3EBD8",
    margin: "6px 0 0",
    lineHeight: 1.15
  },
  stats: {
    display: "flex",
    gap: 22
  },
  nav: {
    display: "flex",
    gap: 8,
    padding: "14px 22px",
    alignItems: "center",
    flexWrap: "wrap"
  },
  filters: {
    display: "flex",
    gap: 8,
    padding: "0 22px 16px",
    flexWrap: "wrap"
  },
  search: {
    flex: "1 1 220px",
    padding: "10px 12px",
    border: "1px solid #C9C2B2",
    borderRadius: 8,
    background: "#FBF8F1",
    fontSize: 14
  },
  select: {
    padding: "10px 12px",
    border: "1px solid #C9C2B2",
    borderRadius: 8,
    background: "#FBF8F1",
    fontSize: 14
  },
  filtersRange: {
    display: "flex",
    gap: 8,
    alignItems: "center",
    padding: "0 22px 16px",
    flexWrap: "wrap"
  },
  rangeLabel: {
    fontSize: 11,
    letterSpacing: ".08em",
    textTransform: "uppercase",
    color: "#7C8AA6"
  },
  rangeInput: {
    width: 76,
    padding: "9px 8px",
    border: "1px solid #C9C2B2",
    borderRadius: 8,
    background: "#FBF8F1",
    fontSize: 14
  },
  rangeDash: {
    color: "#B7AF9C"
  },
  rangeReset: {
    padding: "6px 12px",
    fontSize: 12,
    marginLeft: 6
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill,minmax(290px,1fr))",
    gap: 14,
    padding: "0 22px"
  },
  cardClient: {
    fontSize: 11,
    letterSpacing: ".1em",
    textTransform: "uppercase",
    color: "#3E8E9C",
    fontWeight: 600,
    display: "flex",
    alignItems: "center",
    gap: 6
  },
  cardTitle: {
    fontFamily: "'Archivo', sans-serif",
    fontSize: 16,
    margin: "4px 0 6px",
    lineHeight: 1.25
  },
  cardDesc: {
    fontSize: 13,
    color: "#5A6478",
    margin: "0 0 10px",
    display: "-webkit-box",
    WebkitLineClamp: 3,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
    lineHeight: 1.45
  },
  tagRow: {
    display: "flex",
    gap: 6,
    flexWrap: "wrap",
    alignItems: "center"
  },
  settoreTag: {
    fontSize: 11,
    color: "#7C8AA6"
  },
  fornituraTag: {
    fontSize: 11,
    color: "#7A6BA8",
    fontWeight: 600,
    marginTop: 6
  },
  streamOf: {
    fontSize: 11.5,
    color: "#3E8E9C",
    fontWeight: 600,
    margin: "2px 0 4px",
    cursor: "pointer",
    textDecoration: "underline"
  },
  lightbox: {
    position: "fixed",
    inset: 0,
    background: "rgba(14,31,61,.9)",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 80,
    padding: 20,
    cursor: "zoom-out"
  },
  empty: {
    padding: "50px 22px",
    textAlign: "center",
    color: "#7C8AA6",
    fontSize: 15,
    lineHeight: 1.5
  },
  modeRow: {
    display: "flex",
    gap: 8,
    padding: "0 22px 16px",
    flexWrap: "wrap"
  },
  hint: {
    padding: "0 22px",
    color: "#5A6478",
    fontSize: 14,
    lineHeight: 1.5
  },
  textarea: {
    display: "block",
    width: "calc(100% - 44px)",
    margin: "0 22px",
    padding: "12px",
    border: "1px solid #C9C2B2",
    borderRadius: 10,
    background: "#FBF8F1",
    fontSize: 14,
    fontFamily: "inherit",
    lineHeight: 1.5
  },
  err: {
    color: "#A83A3A",
    padding: "8px 22px",
    fontSize: 13
  },
  overlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(14,31,61,.55)",
    display: "flex",
    alignItems: "flex-end",
    justifyContent: "center",
    zIndex: 40
  },
  modal: {
    background: "#FBF8F1",
    width: "min(680px, 100%)",
    maxHeight: "88vh",
    overflowY: "auto",
    borderRadius: "16px 16px 0 0",
    padding: "22px 20px 30px",
    boxShadow: "0 -8px 40px rgba(0,0,0,.25)"
  },
  toast: {
    position: "fixed",
    bottom: 18,
    left: "50%",
    transform: "translateX(-50%)",
    background: "#0E1F3D",
    color: "#F3EBD8",
    padding: "10px 18px",
    borderRadius: 10,
    fontSize: 13,
    zIndex: 60,
    boxShadow: "0 4px 18px rgba(0,0,0,.3)"
  }
};
const css = `
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@600;700&family=Libre+Franklin:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');
* { box-sizing: border-box; }
button { font-family: 'Libre Franklin', system-ui, sans-serif; cursor: pointer; border: none; }
.tab { padding: 9px 16px; border-radius: 8px; background: transparent; color: #22304A; font-size: 14px; font-weight: 600; border: 1px solid transparent; }
.tab.on { background: #0E1F3D; color: #F3EBD8; }
.mode { padding: 8px 14px; border-radius: 20px; background: #E4E0D4; color: #5A6478; font-size: 13px; font-weight: 600; }
.mode.on { background: #22304A; color: #F3EBD8; }
.primary { background: #0E1F3D; color: #F3EBD8; padding: 11px 18px; border-radius: 9px; font-size: 14px; font-weight: 600; }
.primary:disabled { opacity: .55; cursor: not-allowed; }
.ghost { background: transparent; border: 1px solid #C9C2B2; color: #22304A; padding: 10px 16px; border-radius: 9px; font-size: 14px; font-weight: 600; }
.danger { background: transparent; border: 1px solid #C99; color: #A83A3A; padding: 10px 16px; border-radius: 9px; font-size: 14px; font-weight: 600; }
.rec { background: #FBF8F1; border: 1.5px solid #A83A3A; color: #A83A3A; padding: 11px 18px; border-radius: 9px; font-size: 14px; font-weight: 600; }
.rec.live { background: #A83A3A; color: #fff; animation: pulse 1.2s infinite; }
@keyframes pulse { 50% { opacity: .75; } }
@media (prefers-reduced-motion: reduce) { .rec.live { animation: none; } }
.card, .clientCard { background: #FBF8F1; border: 1px solid #DDD6C6; border-radius: 12px; cursor: pointer; transition: box-shadow .15s, transform .15s; }
.card { display: flex; overflow: hidden; }
.clientCard { padding: 14px 16px; }
.card:hover, .clientCard:hover { box-shadow: 0 6px 20px rgba(14,31,61,.12); transform: translateY(-2px); }
.spine { background: #0E1F3D; color: #F3EBD8; writing-mode: vertical-rl; transform: rotate(180deg); display: flex; justify-content: space-between; align-items: center; padding: 12px 6px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: .06em; }
.spine .imp { color: #C2A35A; font-weight: 600; }
.spine .periodo { font-size: 10.5px; letter-spacing: .02em; white-space: nowrap; }
.pill { color: #fff; font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; letter-spacing: .03em; }
.pillBtn { font-size: 12px; font-weight: 600; padding: 6px 12px; border-radius: 20px; }
.ndaTag { background: #A83A3A; color: #fff; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 20px; letter-spacing: .04em; }
.logoCheck { display: inline-flex; align-items: center; background-image: linear-gradient(45deg, #ccc 25%, transparent 25%), linear-gradient(-45deg, #ccc 25%, transparent 25%), linear-gradient(45deg, transparent 75%, #ccc 75%), linear-gradient(-45deg, transparent 75%, #ccc 75%); background-size: 16px 16px; background-position: 0 0, 0 8px, 8px -8px, -8px 0px; }
.aqTag { background: #C2A35A; color: #0E1F3D; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 20px; letter-spacing: .04em; }
.amsTag { background: #4E7A4E; color: #fff; font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 20px; letter-spacing: .04em; }
.inp { display: block; width: 100%; margin-top: 4px; padding: 9px 11px; border: 1px solid #C9C2B2; border-radius: 8px; background: #fff; font-size: 14px; font-family: inherit; }
.drop { margin: 0 22px; padding: 34px 16px; border: 2px dashed #C9C2B2; border-radius: 12px; text-align: center; color: #5A6478; font-size: 14px; background: #F5F1E7; cursor: pointer; }
.drop:hover { border-color: #0E1F3D; color: #0E1F3D; }
button:focus-visible, input:focus-visible, select:focus-visible, textarea:focus-visible { outline: 2px solid #3E8E9C; outline-offset: 2px; }
@media (max-width: 640px) { .spine { font-size: 10px; } }
`;

// ── schermata di login (Firebase Authentication, stesso progetto dello Staffing) ──
function LoginScreen({
  onLogin
}) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [errore, setErrore] = useState("");
  const handleSubmit = async e => {
    e.preventDefault();
    if (!email.trim() || !password) return;
    setBusy(true);
    setErrore("");
    try {
      await signInWithEmailAndPassword(auth, email.trim(), password);
      // onAuthStateChanged nel wrapper gestisce il resto
    } catch (err) {
      const msg = err.code === "auth/invalid-credential" || err.code === "auth/wrong-password" || err.code === "auth/user-not-found" ? "Email o password non corrette." : err.code === "auth/too-many-requests" ? "Troppi tentativi falliti. Riprova tra qualche minuto." : "Accesso non riuscito. Riprova.";
      setErrore(msg);
      setBusy(false);
    }
  };
  return /*#__PURE__*/React.createElement("div", {
    style: {
      minHeight: "100vh",
      background: "#0E1F3D",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "'Libre Franklin', system-ui, sans-serif",
      padding: 20
    }
  }, /*#__PURE__*/React.createElement("style", null, `@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@700&family=Libre+Franklin:wght@400;500;600&display=swap');`), /*#__PURE__*/React.createElement("form", {
    onSubmit: handleSubmit,
    style: {
      background: "#FBF8F1",
      borderRadius: 16,
      padding: "36px 32px",
      width: "min(380px, 100%)",
      boxShadow: "0 20px 60px rgba(0,0,0,.35)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 11,
      letterSpacing: ".14em",
      textTransform: "uppercase",
      color: "#7C8AA6",
      marginBottom: 6
    }
  }, "Practice Data, AI & RPA"), /*#__PURE__*/React.createElement("h1", {
    style: {
      fontFamily: "'Archivo', sans-serif",
      fontSize: 22,
      fontWeight: 700,
      color: "#22304A",
      margin: "0 0 24px"
    }
  }, "Repository Referenze"), /*#__PURE__*/React.createElement("label", {
    style: {
      display: "block",
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, "Email aziendale"), /*#__PURE__*/React.createElement("input", {
    type: "email",
    required: true,
    autoFocus: true,
    value: email,
    onChange: e => setEmail(e.target.value),
    style: {
      display: "block",
      width: "100%",
      marginTop: 5,
      padding: "10px 12px",
      border: "1px solid #C9C2B2",
      borderRadius: 8,
      fontSize: 14,
      boxSizing: "border-box"
    }
  })), /*#__PURE__*/React.createElement("label", {
    style: {
      display: "block",
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 11,
      letterSpacing: ".08em",
      textTransform: "uppercase",
      color: "#7C8AA6"
    }
  }, "Password"), /*#__PURE__*/React.createElement("input", {
    type: "password",
    required: true,
    value: password,
    onChange: e => setPassword(e.target.value),
    style: {
      display: "block",
      width: "100%",
      marginTop: 5,
      padding: "10px 12px",
      border: "1px solid #C9C2B2",
      borderRadius: 8,
      fontSize: 14,
      boxSizing: "border-box"
    }
  })), errore && /*#__PURE__*/React.createElement("p", {
    style: {
      color: "#A83A3A",
      fontSize: 13,
      margin: "6px 0 0"
    }
  }, errore), /*#__PURE__*/React.createElement("button", {
    type: "submit",
    disabled: busy,
    style: {
      width: "100%",
      marginTop: 18,
      padding: "11px 0",
      background: "#0E1F3D",
      color: "#F3EBD8",
      border: "none",
      borderRadius: 9,
      fontSize: 14,
      fontWeight: 600,
      cursor: busy ? "wait" : "pointer",
      opacity: busy ? .6 : 1
    }
  }, busy ? "Accesso in corso…" : "Accedi"), /*#__PURE__*/React.createElement("p", {
    style: {
      fontSize: 12,
      color: "#7C8AA6",
      marginTop: 16,
      marginBottom: 0,
      lineHeight: 1.5
    }
  }, "Stessa utenza e stesse credenziali del portale Staffing. Per problemi di accesso o per richiedere un'utenza, contatta chi gestisce oggi gli accessi allo Staffing.")));
}

// ── wrapper con autenticazione: gate su Firebase Auth + lettura grant da /users/{uid} ──
export default function RepositoryReferenzeConAuth() {
  const [stato, setStato] = useState("verifica"); // verifica | loggedOut | loggedIn | errorePermessi
  const [utente, setUtente] = useState(null);
  const [gruppo, setGruppo] = useState(null);
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async user => {
      if (!user) {
        setUtente(null);
        setGruppo(null);
        setStato("loggedOut");
        return;
      }
      setUtente(user);
      try {
        // legge il grant di gruppo dallo stesso nodo /users/{uid} usato dal portale Staffing
        const snap = await fbGet(ref(db, `users/${user.uid}`));
        if (!snap.exists()) {
          setStato("errorePermessi");
          return;
        }
        const dati = snap.val();
        setGruppo(dati.gruppo || null);
        setStato("loggedIn");
      } catch (e) {
        setStato("errorePermessi");
      }
    });
    return () => unsub();
  }, []);
  const handleLogout = () => {
    signOut(auth);
  };
  if (stato === "verifica") {
    return /*#__PURE__*/React.createElement("div", {
      style: {
        minHeight: "100vh",
        background: "#0E1F3D",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "#8FA1C4",
        fontFamily: "system-ui, sans-serif"
      }
    }, "Verifica accesso…");
  }
  if (stato === "loggedOut") {
    return /*#__PURE__*/React.createElement(LoginScreen, null);
  }
  if (stato === "errorePermessi") {
    return /*#__PURE__*/React.createElement("div", {
      style: {
        minHeight: "100vh",
        background: "#0E1F3D",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 20
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        background: "#FBF8F1",
        borderRadius: 16,
        padding: "32px 28px",
        maxWidth: 420,
        textAlign: "center"
      }
    }, /*#__PURE__*/React.createElement("h2", {
      style: {
        fontFamily: "'Archivo', sans-serif",
        fontSize: 18,
        color: "#A83A3A",
        marginTop: 0
      }
    }, "Utenza non abilitata"), /*#__PURE__*/React.createElement("p", {
      style: {
        fontSize: 14,
        color: "#22304A",
        lineHeight: 1.5
      }
    }, "L'accesso è riuscito, ma non risulta un profilo associato a questa utenza nel sistema condiviso con lo Staffing. Contatta chi gestisce oggi gli accessi allo Staffing per abilitare l'utenza."), /*#__PURE__*/React.createElement("button", {
      className: "ghost",
      onClick: handleLogout,
      style: {
        marginTop: 10,
        padding: "9px 16px"
      }
    }, "Esci")));
  }
  return /*#__PURE__*/React.createElement(RepositoryReferenzeInterna, {
    gruppo: gruppo,
    userEmail: utente.email,
    onLogout: handleLogout
  });
}
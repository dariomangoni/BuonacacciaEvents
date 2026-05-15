# BuonaCaccia Events  
Riformattazione e visualizzazione avanzata degli eventi scout AGESCI pubblicati su BuonaCaccia.

---

## A chi è rivolto

**BuonaCaccia Events** è uno strumento pensato per rendere più semplice ed efficace la consultazione degli eventi scout AGESCI pubblicati su BuonaCaccia.  
Campetti dall'altra parte d'Italia, eventi con lista d'attesa ormai satura, altri con iscrizioni già chiuse, altri in periodi in cui abbiamo già altri impegni... ora non sono più un problema.

## Cosa puoi fare
- filtrare gli eventi per **Regione**, **Posti rimanenti**, **Periodo** e molto altro
- usare filtri avanzati sulle date (e.g. >20/05/2026)
- visualizzare gli eventi in un formato più leggibile e immediato  
- accedere rapidamente alla pagina ufficiale dell’evento su BuonaCaccia  
- scaricare la lista degli eventi in formato Excel (già predisposta per filtrare efficacemente i risultati)  

## Limitazioni
I dati vengono periodicamente raccolti dal sito ufficiale [BuonaCaccia](https://buonacaccia.net/), ma non sono aggiornati in tempo reale: per questo è sempre necessario verificare le informazioni sul sito ufficiale (semplicemente cliccando sul link dell'evento).

Vista la mia residenza sono pre-impostati filtri per le regioni di mio interesse.

---

## Per sviluppatori e contributori

BuonaCaccia Events è una **web app statica** ospitata su **GitHub Pages**.  

I dati sono recuperati da uno script Python che effettua lo scraping di BuonaCaccia recuperando la lista degli eventi più gettonati (PiccoleOrme, Specialità, Competenza) non solo dalla pagina riassuntiva, ma navigando all'interno di ciascun evento al fine di recuperare anche le date di apertura e chiusura iscrizioni. Lo script genera infine un file JSON e un file Excel con i dati recuperati (una coppia di file per ciascun tipo di eventi).

Il file JSON funge da database e viene letto dalla pagina *index_template.html* per generare la table con gli eventi. Avrei potuto generare direttamente la pagina HTML, ma avere un database vero e proprio consente di evidenziare eventuali campetti inseriti su BuonaCaccia tra la run attuale e quella precedente.

Il file Excel è anch'esso fornito di filtri preimpostati per permetterne un'immediata consultazione.

La pagina *index.html* funge solamente da entry point. Sono invece le singole pagine *index_template.html* a contenere la lista di eventi di quella specifica tipologia (PiccoleOrme, Specialità, Competenza). Tale pagina è la medesima per tutte le tipologie: essa viene infatti copiata automaticamente nelle sottocartelle in cui è stato generato il file JSON specifico per tale tipo di eventi.


### Stack
- script Python per lo scraping di Buonacaccia con generazione file JSON ed Excel per ciascun tipo di eventi
- generazione **HTML / CSS / JavaScript** (vanilla) con parsing del file JSON
- deploy su **GitHub Pages** tramite **GitHub Actions** ad intervalli prestabiliti (più frequenti nei mesi "caldi" per le iscrizioni)

### Licenza
MIT License — libero utilizzo e modifica, con attribuzione.

---

## Contatti
Per domande o suggerimenti:  
**Dario Mangoni** – GitHub: https://github.com/dariomangoni

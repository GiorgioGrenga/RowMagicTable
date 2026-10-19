# Tabella → Excel

Web app responsive per estrarre tabelle da immagini e PDF con Gemini, rivedere
e correggere i dati nel browser e scaricare un file Excel. È possibile caricare
anche un `.xlsx`: l'app mostra i fogli, abbina le intestazioni compatibili e
permette di aggiungere, rimappare o ignorare le altre colonne.

## Avvio in locale

Serve Python 3.10 o più recente e una chiave Gemini.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:GEMINI_API_KEY = "la-tua-chiave"
uvicorn server:app --reload
```

Apri <http://127.0.0.1:8000>. La chiave si crea in Google AI Studio e va
configurata solo come variabile d'ambiente del server; non inserirla nel
frontend o nei file del repository.

## Pubblicazione su Render

1. Pubblica questi file in un repository GitHub.
2. In Render crea un nuovo **Blueprint** collegando il repository: Render
   leggerà `render.yaml` e creerà il servizio web.
3. Nelle impostazioni del servizio aggiungi `GEMINI_API_KEY` con la tua chiave
   Google AI Studio.
4. Apri l'URL HTTPS assegnato da Render.

Il piano gratuito può sospendere il servizio quando non viene usato; la prima
apertura successiva può quindi richiedere qualche secondo. L'estrazione AI
dipende dalle quote e dalla disponibilità del piano gratuito Gemini. Una
chiave API valida rimane necessaria.

## Uso

1. Carica un'immagine PNG/JPG o un PDF (massimo 15 MB e 40 pagine).
2. Facoltativamente carica un file Excel `.xlsx` (massimo 15 MB).
3. Controlla e, se necessario, modifica celle e intestazioni estratte.
4. Con un Excel esistente scegli il foglio e mappa le colonne: le intestazioni
   uguali vengono abbinate, le altre si possono aggiungere o ignorare.
5. Scarica il nuovo Excel oppure la copia aggiornata del file caricato.

I file vengono elaborati in memoria e non salvati dall'app. Il documento
immagine/PDF è trasmesso a Google Gemini per l'estrazione; il file Excel viene
letto e modificato sul server ma non inviato al provider AI. Sul piano gratuito,
Google può utilizzare i contenuti inviati per migliorare i propri prodotti:
non caricare documenti riservati o dati personali sensibili.

## Note

- Sono supportati `.png`, `.jpg`, `.jpeg`, `.pdf` e workbook `.xlsx`.
- Quando si aggiunge al workbook, le altre schede sono mantenute e le righe
  vengono accodate al foglio scelto.
- Il server non implementa account o login: il link pubblico può essere usato
  da chiunque lo riceva. Proteggi il deploy e la quota API se la condivisione
  pubblica non è desiderata.

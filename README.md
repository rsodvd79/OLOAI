# Outlook AI Assistant Add-in

Estensione VSTO per Microsoft Outlook Classic 2024 LTSC che integra i modelli OpenAI per assisterti nella scrittura, riscrittura e correzione delle email direttamente dal client desktop.

## Requisiti
- Windows con Microsoft Outlook Classic 2024 LTSC installato
- Visual Studio 2022 (o superiore) con carico di lavoro **Office/SharePoint**
- .NET Framework 4.8
- Chiave API OpenAI valida (modelli compatibili con l'endpoint `chat/completions`)

## Configurazione della chiave API
1. Imposta la variabile di ambiente `OPENAI_API_KEY` con la tua chiave:
   ```powershell
   setx OPENAI_API_KEY "sk-..."
   ```
   > In alternativa, puoi compilare `app.config` con i valori `OpenAI__ApiKey`, `OpenAI__Model`, `OpenAI__BaseUrl`, `OpenAI__MaxTokens` e facoltativamente `OpenAI__Temperature`.

2. (Opzionale) Specifica un modello diverso da `gpt-4o-mini` aggiornando `app.config` oppure impostando la variabile `OPENAI_MODEL` prima di lanciare Outlook.

## Struttura del progetto
- `OutlookAiAddIn.sln` – soluzione Visual Studio
- `OutlookAiAddIn/OutlookAiAddIn.csproj` – progetto VSTO targeting .NET Framework 4.8
- `Services/OpenAIService.cs` – wrapper per le chiamate all'API OpenAI
- `Services/OutlookContextService.cs` – helper per interagire con Outlook (selezioni, bozze, reply)
- `UI/AiTaskPane*.cs` – pannello laterale per generare e applicare i suggerimenti
- `Ribbon/AiRibbon.*` – Ribbon XML e code-behind con i comandi Outlook
- `ThisAddIn.*` – bootstrap dell'add-in e wiring del ribbon/pannello

## Build & debug
1. Apri la soluzione `OutlookAiAddIn.sln` in Visual Studio.
2. Imposta il progetto `OutlookAiAddIn` come avviabile.
3. Verifica che la chiave API sia disponibile nell'ambiente di debug (`Project Properties` → `Debug` → `Environment variables`).
4. Premi **F5** per compilare ed avviare Outlook in modalità debug. Il Ribbon "Assistente AI" apparirà nella scheda Posta.

## Uso dell'add-in
- Dal ribbon "Assistente AI" puoi aprire il pannello e scegliere:
  - **Suggerisci risposta**: analizza l'email selezionata e genera una bozza di risposta.
  - **Migliora bozza**: riscrive un messaggio che stai componendo.
  - **Correggi errori**: corregge ortografia e grammatica del testo selezionato.
- Nel pannello inserisci eventuali istruzioni aggiuntive (tono, punti chiave) e premi **Genera suggerimento**.
- Usa **Inserisci in Outlook** per creare/revisionare automaticamente la bozza oppure **Copia testo** per incollare manualmente.

## Distribuzione
1. Compila in modalità `Release`.
2. Firma l'assembly (Project Properties → Signing) e genera un installer ClickOnce o MSI secondo la tua policy aziendale.
3. Distribuisci includendo i prerequisiti VSTO e assicurando che gli aggiornamenti di sicurezza di Office siano applicati.

## Note e limiti
- L'add-in utilizza l'endpoint `chat/completions`; adatta il modello se OpenAI cambia politiche o naming.
- La conversione HTML è basilare; per scenari complessi valuta librerie dedicate.
- Gestisci la chiave API in modo sicuro (Azure Key Vault, Windows Credential Manager, ecc.).
- L'add-in effettua chiamate HTTP; assicurati che la rete aziendale consenta raggiungibilità verso `api.openai.com`.

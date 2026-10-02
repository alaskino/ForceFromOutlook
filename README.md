# Force From for Outlook

A lightweight **Microsoft Outlook Classic VSTO add-in** that forces outgoing messages to use the **default Outlook profile account** when enabled.

Developed by **Mathieu Licata (Alaska)**.

> Windows + Outlook Classic only.

---

## English

### What it does

**Force From** adds an ON/OFF toggle to the Outlook ribbon.

When **ON**:

- new messages use the default Outlook profile account;
- replies, Reply All and forwarded messages use the default account;
- inline replies are handled;
- the sender account is checked again immediately before sending;
- if the default account cannot be determined, the message is blocked instead of risking an incorrect sender.

When **OFF**:

- Outlook behaves normally;
- the sender can be changed freely using the **From** field.

The ON/OFF state is saved for the current Windows user and restored after Outlook restarts.

### Interface

The ribbon icon changes dynamically:

- **ON**: blue envelope with a green status indicator;
- **OFF**: gray envelope with a gray status indicator.

The UI automatically uses **Italian or English** according to the current UI culture.

### Requirements

- Windows
- Microsoft Outlook Classic desktop
- .NET Framework 4.8
- Visual Studio Tools for Office Runtime (VSTO Runtime)

For development:

- Microsoft Visual Studio 2022
- Office/SharePoint development workload

### Build from source

1. Open `ForceFromOutlook.sln` in Visual Studio 2022.
2. Select **Release**.
3. Build the solution with `Ctrl+Shift+B`.
4. Publish the VSTO project using **ClickOnce** if you want an installer.

### Installation

Download the installer ZIP from the **Releases** section.

1. Extract the ZIP.
2. Close Outlook.
3. Run `setup.exe`.
4. Complete the installation.
5. Start Outlook Classic.

The **Force From** button should appear in the Outlook ribbon.

### Privacy

Force From:

- does not send data to external servers;
- does not store email contents;
- does not require an online service;
- works locally inside Outlook;
- stores only the ON/OFF preference for the current Windows user.

### Technical notes

The add-in uses the Outlook Object Model and `MailItem.SendUsingAccount`.

The default sending account is resolved by matching the Outlook session default store with the delivery store of the configured Outlook accounts.

A final sender check is performed in Outlook's `ItemSend` event.

---

## Italiano

### Cosa fa

**Force From** aggiunge un interruttore ON/OFF alla barra multifunzione di Outlook.

Quando è **ON**:

- i nuovi messaggi usano l'account predefinito del profilo Outlook;
- Rispondi, Rispondi a tutti e Inoltra usano l'account predefinito;
- vengono gestite anche le risposte inline;
- il mittente viene ricontrollato immediatamente prima dell'invio;
- se non è possibile determinare l'account predefinito, il messaggio viene bloccato invece di rischiare l'invio dalla casella sbagliata.

Quando è **OFF**:

- Outlook funziona normalmente;
- il mittente può essere modificato liberamente tramite il campo **Da**.

Lo stato ON/OFF viene salvato per l'utente Windows corrente e ripristinato dopo il riavvio di Outlook.

### Interfaccia

L'icona nella Ribbon cambia dinamicamente:

- **ON**: busta blu con indicatore verde;
- **OFF**: busta grigia con indicatore grigio.

L'interfaccia usa automaticamente **Italiano o Inglese** in base alla lingua dell'interfaccia utente.

### Requisiti

- Windows
- Microsoft Outlook classico desktop
- .NET Framework 4.8
- Visual Studio Tools for Office Runtime (VSTO Runtime)

Per lo sviluppo:

- Microsoft Visual Studio 2022
- workload Sviluppo per Office/SharePoint

### Compilazione dai sorgenti

1. Apri `ForceFromOutlook.sln` in Visual Studio 2022.
2. Seleziona **Release**.
3. Compila la soluzione con `Ctrl+Shift+B`.
4. Pubblica il progetto VSTO tramite **ClickOnce** se vuoi generare l'installer.

### Installazione

Scarica lo ZIP dell'installer dalla sezione **Releases**.

1. Estrai lo ZIP.
2. Chiudi Outlook.
3. Avvia `setup.exe`.
4. Completa l'installazione.
5. Avvia Outlook classico.

Il pulsante **Forza DA** dovrebbe comparire nella barra multifunzione di Outlook.

### Privacy

Force From:

- non invia dati a server esterni;
- non salva il contenuto delle email;
- non richiede servizi online;
- funziona localmente all'interno di Outlook;
- salva soltanto la preferenza ON/OFF dell'utente Windows corrente.

### Note tecniche

L'add-in utilizza l'Outlook Object Model e `MailItem.SendUsingAccount`.

L'account di invio predefinito viene individuato confrontando lo store predefinito della sessione Outlook con il delivery store degli account configurati.

Prima dell'invio viene eseguito un controllo finale tramite l'evento `ItemSend`.

---

## Author

**Mathieu Licata**  
Nickname: **Alaska**

## License

No license selected yet.

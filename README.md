# ETAP SLD Batch Printer

Questo programma automatizza l'esportazione dei Single Line Diagrams (SLD) da ETAP in formato PDF.
Permette di eseguire scenari multipli (incluso il Motor Starting) e salvare automaticamente i risultati in una cartella specifica.

## Requisiti

Il programma richiede l'installazione di alcune librerie Python aggiuntive nell'ambiente Python di ETAP.
Le librerie necessarie sono:

*   **customtkinter**: Per l'interfaccia grafica moderna.
*   **pyautogui**: Per l'automazione di mouse e tastiera.

## Istruzioni per l'Installazione

Poiché non hai accesso diretto alla PowerShell di sistema, devi installare i pacchetti utilizzando la **Console Python di ETAP**.

1.  Apri ETAP.
2.  Apri la **Console Python** (di solito si trova in basso o nel menu "Tools" -> "Python Console").
3.  Copia e incolla i seguenti blocchi di codice (uno alla volta) nella console e premi **Invio**.

### 1. Installazione di CustomTkinter

Copia questo codice nella console Python di ETAP:

```python
import sys
import subprocess

subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'customtkinter'])
```

### 2. Installazione di PyAutoGUI

Copia questo codice nella console Python di ETAP:

```python
import sys
import subprocess
  
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyautogui'])

```

## Come Eseguire il Programma

1.  Assicurati che il file `.py` (o il nome del tuo script principale) sia nella cartella corretta.
2.  Nella Console Python di ETAP, esegui il file usando il comando:

```python
with open(r"c:\ETAP 2400\ThirdParty\Python\Distro\Scripts\_temp\main_final.py") as f:
    exec(f.read())
```
*(Assicurati che il percorso del file sia corretto)*

## Funzionalità Principali

*   **Connessione**: Si connette all'istanza ETAP aperta.
*   **Selezione Scenari**: Visualizza tutti gli scenari disponibili raggruppati per modalità.
*   **Automazione**: Esegue gli scenari selezionati, gestisce la stampa PDF e, nel caso di Motor Starting, avanza automaticamente i time step.
*   **Impostazioni**: Permette di configurare i ritardi temporali e la cartella di output tramite il pulsante ingranaggio (⚙️).

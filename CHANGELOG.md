# CHANGELOG - AD-ExpiryManager

Dalla più recente. Lo stato "provato" indica solo ciò che l'utente ha confermato in esecuzione reale.

## 2026-10-22 (2) - v4.32
- Caso A: se non è Family Web né Due Diligence, menu dei profili di AD-UserProvisioning (`esterno_gruppi_default`, `commesse`, `legal_factory`) o `utente_esistente` (copia OU+gruppi da un riferimento). Status `Profiled_Completed` / `Error_Profilo`.
- Nuove funzioni `Invoke-ProfilazioneUtente`, `Read-UtenteRiferimento`; dati in `$profiliProvisioning`.
- Stato: **non provato** (sintassi verificata, nessuna esecuzione su AD).

## 2026-10-22 - v4.31
- Salvataggio facoltativo e cifrato delle credenziali admin (`Get-CredenzialiSalvate`, `-DimenticaCredenziali` inoltrato al RunAs). Origine: `AD-UserProvisioningNEW\strumenti\Salvataggio-Credenziali_istruzioni.md`.
- Stato: **non provato**.

## 2026-10-08 - v4.30
- Caso E: utente attivo con gruppi in OU Disabilitati → segnalazione, scelta OU 1-6, sposta + proroga (`Moved_Updated`). Prima veniva scartato senza messaggio né log.
- Stato: **non provato**.

## 2026-10-05 - v4.29
- Verifica reale su AD a fine batch (`VERIFY_FAILED` se non applicato); testo ticket negli appunti solo se almeno un utente è verificato.
- Log scritto nella cartella dello script (fallback Desktop dell'utente originale), non più sul Desktop dell'account admin del RunAs.
- Stato: log confermato funzionante dall'utente; esito del controllo in AD non comunicato.

## 2026-09-21 - v4.28
- Caso A: dopo "Family Web? NO" chiede "Due Diligence?"; se sì gruppi (7), OU `\ConsulentiEsterni\DueDiligence`, abilita, proroga. Status `DueDiligence_Completed` / `Error_DueDiligence`.
- Stato: l'utente ha dichiarato lo script testato e funzionante.

## v4.27 e precedenti (2026-08-25 e prima)
Separatore selezione multipla, ricerca Nome Cognome con menu, riepilogo visivo, Family Web, riattivazione, log Excel. Dettaglio nell'header del `.ps1`.

## Rollback
La cartella non è un repository git e non ci sono copie delle versioni precedenti: **non esiste un rollback automatico**. Prima di nuove modifiche salvare una copia del `.ps1` (es. `PROROGA MASSIVA UTENZE AD_v4.32.ps1`) o inizializzare git.

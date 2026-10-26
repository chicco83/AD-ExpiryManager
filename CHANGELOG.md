# CHANGELOG - AD-ExpiryManager

Dalla più recente. Lo stato "provato" indica solo ciò che l'utente ha confermato in esecuzione reale. Il 2026-10-26 l'utente ha confermato provate e funzionanti le versioni 4.29-4.32.

## 2026-10-22 (2) - v4.32
- Caso A: se non è Family Web né Due Diligence, menu dei profili di AD-UserProvisioning (`esterno_gruppi_default`, `commesse`, `legal_factory`) o `utente_esistente` (copia OU+gruppi da un riferimento). Status `Profiled_Completed` / `Error_Profilo`.
- Nuove funzioni `Invoke-ProfilazioneUtente`, `Read-UtenteRiferimento`; dati in `$profiliProvisioning`.
- Stato: provato e funzionante (conferma utente, 2026-10-26).

## 2026-10-22 - v4.31
- Salvataggio facoltativo e cifrato delle credenziali admin (`Get-CredenzialiSalvate`, `-DimenticaCredenziali` inoltrato al RunAs). Origine: `AD-UserProvisioningNEW\strumenti\Salvataggio-Credenziali_istruzioni.md`.
- Stato: provato e funzionante (conferma utente, 2026-10-26).

## 2026-10-08 - v4.30
- Caso E: utente attivo con gruppi in OU Disabilitati → segnalazione, scelta OU 1-6, sposta + proroga (`Moved_Updated`). Prima veniva scartato senza messaggio né log.
- Stato: provato e funzionante (conferma utente, 2026-10-26).

## 2026-10-05 - v4.29
- Verifica reale su AD a fine batch (`VERIFY_FAILED` se non applicato); testo ticket negli appunti solo se almeno un utente è verificato.
- Log scritto nella cartella dello script (fallback Desktop dell'utente originale), non più sul Desktop dell'account admin del RunAs.
- Stato: provato e funzionante (conferma utente, 2026-10-26); log confermato il 2026-10-05.

## 2026-09-21 - v4.28
- Caso A: dopo "Family Web? NO" chiede "Due Diligence?"; se sì gruppi (7), OU `\ConsulentiEsterni\DueDiligence`, abilita, proroga. Status `DueDiligence_Completed` / `Error_DueDiligence`.
- Stato: l'utente ha dichiarato lo script testato e funzionante.

## v4.27 e precedenti (2026-08-25 e prima)
Separatore selezione multipla, ricerca Nome Cognome con menu, riepilogo visivo, Family Web, riattivazione, log Excel. Dettaglio nell'header del `.ps1`.

## Rollback
Per le versioni 4.28-4.31 non esistono copie separate. Dal commit `912178c` (v4.32) il progetto è su git (https://github.com/chicco83/AD-ExpiryManager): per tornare a una versione `git checkout <commit> -- "PROROGA MASSIVA UTENZE AD.ps1"`.

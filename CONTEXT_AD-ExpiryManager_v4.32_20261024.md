# CONTEXT AD-ExpiryManager - v4.32 - 2026-10-24

Fonte di verità tra sessioni. Legenda: ✅ fatto e verificato · ⏳ da fare · ⏸️ rinviato · ⛔ voluto così.

## Stato

| Voce | Stato | Note |
|---|---|---|
| Proroga diretta (Caso C), riattivazione (B, D), Family Web | ✅ | in uso da prima della v4.28 |
| Due Diligence (v4.28) | ✅ | utente: "script testato e funzionante" |
| Log nella cartella dello script (v4.29) | ✅ | l'utente ha trovato il log |
| Verifica su AD a fine batch (v4.29) | ⏳ | l'utente stava controllando in AD; esito non comunicato |
| Caso E attivo in OU disabilitati (v4.30) | ⏳ | mai eseguito |
| Credenziali salvate (v4.31) | ⏳ | mai eseguito; la funzione stessa è "non provata su VDI reale" anche nel progetto d'origine |
| Menu profili Caso A (v4.32) | ⏳ | mai eseguito |
| **Segnalazione 2026-10-05: il 21/09 lo script dava esito OK ma in AD non risultava nulla** | ⏳ | causa **non trovata**: nessuna modifica in AD il 21/09; log del 21/09 mai recuperato |

## Problema aperto (21/09)
Lo script segnalava successo, ma gli utenti risultavano ancora disabilitati, in OU Disabilitati, senza gruppi; il metadata di replica non mostrava modifiche quel giorno. Dal codice non risulta un percorso che dia esito positivo senza scrivere (tutte le scritture usano `-ErrorAction Stop`). Ipotesi non verificate: sessione/copia diversa da quella attesa, utente non effettivamente processato (es. `Skipped_*`/`NotFound`), log non letto. Per chiudere serve il file `AD_Expiry_Log_*` di quella sessione (cercarlo anche nel Desktop dell'account admin del RunAs). Dalla v4.29 il `VERIFY_FAILED` rende il problema visibile.

## Decisioni (con il perché)
- **Profili agente esclusi dal menu**: scadenza "Mai" e attributi propri (extensionAttribute9, Description, Note) non si adattano a una proroga a 6 mesi.
- **`consulente_societa_recupero` escluso**: nel template non ha gruppi configurati.
- **`dipendente_*` esclusi**: modalità dipendenti disattivata in AD-UserProvisioning.
- **Profili copiati nello script, non letti dal template Excel**: lo script gira dove il template non è garantito (SharePoint, RunAs); costo: allineamento manuale con `AD-UserProvisioningNEW\v*\Template_Creazione_Utenti_AD`.
- **Family Web e Due Diligence restano domande S/N separate** (richieste e provate prima del menu), il menu compare solo se entrambe sono NO.
- **Proroga sempre scritta anche se già uguale**: `NoChange` è solo un'etichetta del log, nessuna operazione viene saltata.
- **Non copiata nel documento la lista gruppi differenze**: Family Web nello script non include `ext.users.automationscript`, presente invece nel template v2.52; Due Diligence nello script ha in più `app.commesse20.duediligencevendita.op`. Non uniformati finché non richiesto.

## Da fare, in ordine
1. Provare v4.32 su utenti di test: Caso E, credenziali salvate, ogni voce del menu profili, `utente_esistente`.
2. Recuperare il log del 21/09 e chiudere il problema aperto.
3. Decidere se uniformare le liste gruppi Family Web / Due Diligence al template.
4. Salvare una copia del `.ps1` per ogni versione o inizializzare git (oggi nessun rollback).

## Rollback
Nessuno disponibile (non è un repository git, nessuna copia delle versioni). Vedi punto 4.

## Dipendenze
Template profili: `..\AD-UserProvisioningNEW\v2.52\Template_Creazione_Utenti_AD.xlsx` (fogli `Config_OU_Profilo`, `Config_Gruppi_Profilo`). Istruzioni credenziali: `..\AD-UserProvisioningNEW\strumenti\Salvataggio-Credenziali_istruzioni.md`.

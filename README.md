# AD-ExpiryManager

Script PowerShell per la **proroga massiva della scadenza degli account AD** (utenti esterni/consulenti) di Banca Ifis, con riattivazione e profilazione quando serve.

File: `PROROGA MASSIVA UTENZE AD.ps1` (versione corrente: vedi `CHANGELOG.md`).

## Cosa fa

1. Si auto-eleva (UAC) e chiede le credenziali AD admin (salvabili in modo cifrato, vedi sotto).
2. L'operatore incolla una lista di utenti (SAM, UPN o `Nome Cognome`; tre INVIO per terminare).
3. Per ogni utente imposta `accountExpirationDate` a **+6 mesi, fine giornata** (calcolata una volta all'avvio).
4. A fine batch **rilegge AD** e verifica ogni utente dato per riuscito (scadenza, abilitato, OU): se non corrisponde lo status diventa `VERIFY_FAILED`.
5. Copia negli appunti "utenza prorogata al gg/mm/aaaa" solo se almeno un utente è verificato.
6. Salva il log Excel (o CSV) **nella cartella dello script** (fallback: Desktop dell'utente originale), un file per sessione, scritto solo quando si risponde N a "Vuoi prorogare altri utenti?".

## Casi gestiti per utente

| Caso | Condizione | Azione |
|---|---|---|
| A | solo Domain Users (nessun gruppo extra) | chiede Family Web → Due Diligence → altri profili di AD-UserProvisioning; con profilo: gruppi + sposta OU + abilita + proroga; senza: `Blocked_NoGroups` |
| B | disabilitato, in OU Disabilitati, con gruppi | conferma, scelta OU (1-6), sposta + abilita + proroga |
| C | attivo, OU normale | proroga diretta |
| D | disabilitato, OU normale | menu: riattiva+sposta+proroga / solo proroga / salta |
| E | attivo, con gruppi, in OU Disabilitati | segnala anomalia, propone spostamento in OU (1-6) + proroga |

Ricerca per Nome Cognome: 3 tentativi (GivenName/Surname senza accenti, LDAP `cn=`, LDAP ordine inverso); con più risultati menu numerato (anche selezione multipla).

## Profili (Caso A)

| Profilo | OU | Origine |
|---|---|---|
| Family Web | `\IFISNplServicing\Users\External` | lista nello script |
| Due Diligence | `\ConsulentiEsterni\DueDiligence` | lista nello script |
| `esterno_gruppi_default` | `\IFIS\Users\External` | copia del template AD-UserProvisioning v2.52 |
| `commesse` | `\ConsulentiEsterni\Legal Factory` | idem |
| `legal_factory` | `\ConsulentiEsterni\Legal Factory` | idem |
| `utente_esistente` | OU del riferimento | copia OU+gruppi da un utente attivo indicato dall'operatore |

I profili sono scritti nello script (`$profiliProvisioning`), non letti dal template: se cambiano in `AD-UserProvisioningNEW\v*\Template_Creazione_Utenti_AD` vanno aggiornati anche qui.

## Credenziali salvate

Dopo `Get-Credential` lo script chiede se salvarle (`Export-Clixml`, cifrate DPAPI: leggibili solo dallo stesso utente Windows sullo stesso PC, cioè l'account della sessione elevata). File: `%LOCALAPPDATA%\AD-UserProvisioning\credenziali_expirymanager.xml`. Eliminare con il parametro `-DimenticaCredenziali`. Mai su OneDrive/repository.

## Come si usa

Avvio dal file `.ps1` (anche da SharePoint). Il banner iniziale mostra la versione in esecuzione: controllare che sia quella attesa, altrimenti si sta usando una copia vecchia.

Requisiti: PowerShell 5.1 (Constrained Language Mode), modulo ActiveDirectory (RSAT), modulo ImportExcel in `OneDrive - Banca Ifis S.p.A\PSModules` (se manca: log CSV), account AD con permessi di scrittura.

## Regole del file

Solo caratteri ASCII nelle istruzioni (accenti solo nei commenti), fine riga CRLF, nessun BOM. Versionare ogni modifica nell'header e nel banner.

## Documenti

`CHANGELOG.md` cronologia · `CONTEXT_AD-ExpiryManager_*.md` stato attuale e decisioni.

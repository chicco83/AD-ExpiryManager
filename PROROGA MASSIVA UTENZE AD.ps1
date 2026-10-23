# ============================================================
# AD_Expiry_Manager_v4.ps1
# Versione: 4.32
# Data revisione: 2026-10-22
#
# Fix v4.32 (2026-10-22):
#   [NEW] Caso A: se l'utente (solo Domain Users) non e' ne' Family Web ne' Due
#         Diligence, propone i profili di AD-UserProvisioning (esterno_gruppi_default,
#         commesse, legal_factory) oppure utente_esistente (copia OU+gruppi da un
#         utente di riferimento). Assegna gruppi, sposta, abilita, proroga.
#         Status log: Profiled_Completed / Error_Profilo. Con 0 resta Blocked_NoGroups.
#
# Fix v4.31 (2026-10-22):
#   [NEW] Salvataggio facoltativo e cifrato (DPAPI, Export-Clixml) delle credenziali
#         admin in %LOCALAPPDATA%\AD-UserProvisioning\credenziali_expirymanager.xml.
#         Parametro -DimenticaCredenziali per eliminarlo (inoltrato anche al RunAs).
#
# Fix v4.30 (2026-10-08):
#   [NEW] Caso E: utente ATTIVO con gruppi in OU disabilitati -> segnalazione
#         anomalia, proposta di spostamento in una delle 6 OU di ripristino e
#         proroga. Status log: Moved_Updated (prima veniva scartato in silenzio).
#
# Fix v4.29 (2026-10-05):
#   [FIX] Log salvato nella cartella dello script ($PSScriptRoot) se scrivibile,
#         altrimenti sul Desktop dell'utente originale ($OriginalUserProfile)
#         e non su quello dell'account admin usato da RunAs.
#   [FIX] Verifica reale su AD a fine batch (rilettura scadenza/abilitato/OU):
#         se non applicato -> Status VERIFY_FAILED. Il testo ticket negli
#         appunti viene copiato solo se almeno un utente e' verificato
#         (prima veniva copiato sempre, anche con 0 utenti prorogati).
#
# Fix v4.28 (2026-09-21):
#   [NEW] Caso A: se l'utente (solo Domain Users) non e' Family Web,
#         viene chiesto se e' un utente DUE DILIGENCE. Se confermato:
#         assegna gruppi (ext.users.automationscript, GPO.NoLocalAndRDPLogon.AppUser,
#         app.o365.ssrp, cbox.npl.ro.prd, nac.Consulenti, cbox.npl.prd,
#         app.commesse20.duediligencevendita.op), sposta in OU
#         \ConsulentiEsterni\DueDiligence, abilita e proroga.
#         Status log: DueDiligence_Completed / Error_DueDiligence
#
# Fix v4.27 (2026-08-25):
#   [NEW] Separatore visivo tra la fase di selezione degli utenti
#         (menu ambiguo) e la fase di proroga degli utenti accodati
#         dalla selezione multipla.
#           Citrix.MFA.O365, nac.Consulenti,
#           GPO.NoLocalAndRDPLogon.AppUser
#         Status log: FamilyWeb_Completed / Error_FamilyWeb
#
# [v4.4 - 2026-07-13]
#   [FIX] param() in cima / percorso PSModules Banca Ifis / no download
# [v4.3 - 2026-07-13]
#   [FIX] ImportExcel: Save-Module nel profilo utente originale
# [v4.2 - 2026-07-13]
#   [NEW] Riepilogo visivo compatto per ogni utente trovato
# [v4.1 - 2026-07-13]
#   [FIX] Ricerca Nome Cognome multipla: menu numerato
# [v4 - 2026-07-13]
#   [FIX] Contatore Updated a 0 / Loop ripartenza a fine batch
# [v3.1 - 2026-07-13]
#   [FIX] $beforeDT init / ObjectGUID dopo Move / -ErrorAction Stop
# [v3 - 2026-07-13]
#   Caso C proroga diretta / Caso D menu 1-2-3
# [v2 - 2025-07-13]
#   Controlli disabilitato, gruppi, OU / output Excel
# ============================================================

# IMPORTANTE: param() deve essere la prima istruzione eseguibile dello script.
# Riceve il profilo dell'utente originale (non-admin) passato dal blocco RunAs.
# Default: $env:USERPROFILE (funziona se lo script gira gia' come utente corretto)
param(
    [string]$OriginalUserProfile = $env:USERPROFILE,
    [switch]$DimenticaCredenziali
)

# -------------------------------------------------------
# Forza codepage UTF-8 (65001) sulla console per gestire
# correttamente caratteri accentati nei nomi (es. Arnò).
# Si usa chcp invece di [Console]::OutputEncoding perche'
# l'ambiente e' in Constrained Language Mode (CLM) che
# blocca l'accesso diretto alle classi .NET.
# L'output di chcp viene soppresso (> $null).
# -------------------------------------------------------
chcp 65001 > $null

# -------------------------------------------------------
# Funzione: verifica privilegi amministrativi con 'net session' (ASCII only)
# -------------------------------------------------------
function Test-IsAdmin {
    net session > $null 2>&1
    if ($LASTEXITCODE -eq 0) { return $true } else { return $false }
}

# -------------------------------------------------------
# Auto-elevazione se non Admin, mantenendo la finestra aperta.
# Passa il profilo dell'utente corrente (non-admin) al processo elevato
# in modo che sotto RunAs sappia dove trovare il modulo ImportExcel.
# -------------------------------------------------------
if (-not (Test-IsAdmin)) {
    Write-Host "Lo script non e' in esecuzione come amministratore. Riavvio con privilegi elevati..." -ForegroundColor Yellow

    if (-not $PSCommandPath) {
        Write-Host "Errore: impossibile determinare il percorso dello script. Salva il codice in un file .ps1 e rilancialo." -ForegroundColor Red
        Read-Host -Prompt "Premi INVIO per chiudere"
        exit 1
    }

    # $env:USERPROFILE qui e' ancora quello dell'utente non-admin
    $elevArgs = "-NoProfile -ExecutionPolicy Bypass -NoExit -File `"$PSCommandPath`" -OriginalUserProfile `"$($env:USERPROFILE)`""
    if ($DimenticaCredenziali) { $elevArgs += " -DimenticaCredenziali" }
    Start-Process powershell -ArgumentList $elevArgs -Verb RunAs
    exit
}

# ========================
# 0. PREPARAZIONE
# ========================

# Banner versione: primo output visibile all'operatore
# per confermare che si sta eseguendo la versione corretta
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  AD Expiry Manager  -  Versione 4.32" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# --- Percorso PSModules dell'utente originale (non-admin) ---
# Il modulo ImportExcel deve essere pre-copiato in:
#   C:\Users\<utente>\OneDrive - Banca Ifis S.p.A\PSModules\ImportExcel
# Sotto RunAs $env:USERPROFILE punta al profilo admin, quindi usiamo
# $OriginalUserProfile ricevuto come parametro.

$psModulesPath = Join-Path $OriginalUserProfile "OneDrive - Banca Ifis S.p.A\PSModules"

Write-Host ("Percorso PSModules utente: {0}" -f $psModulesPath) -ForegroundColor DarkGray

# Aggiunge il percorso a PSModulePath se esiste e non e' gia' presente
if (Test-Path $psModulesPath) {
    if ($env:PSModulePath -notlike "*$psModulesPath*") {
        $env:PSModulePath = "$psModulesPath;$env:PSModulePath"
        Write-Host "PSModulePath aggiornato con il percorso utente." -ForegroundColor DarkGray
    }
} else {
    Write-Host ("ATTENZIONE: cartella PSModules non trovata in: {0}" -f $psModulesPath) -ForegroundColor Yellow
    Write-Host "Crea la cartella e copia la cartella ImportExcel al suo interno." -ForegroundColor Yellow
}

# Import modulo AD
# -WarningAction SilentlyContinue sopprime il warning "Errore di inizializzazione
# dell'unita' predefinita" che il provider AD:\ genera quando la sessione elevata
# (RunAs con account admin locale) non ha accesso diretto al dominio.
# E' un avviso innocuo: le query AD funzionano correttamente perche' usano
# -Credential esplicito su ogni cmdlet.
try {
    Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
}
catch {
    Write-Host "Modulo ActiveDirectory non disponibile. Installa RSAT / AD PowerShell." -ForegroundColor Red
    exit 1
}

# Import modulo ImportExcel dal percorso PSModules utente
# Se non disponibile si usa CSV come fallback (nessun tentativo di download:
# la macchina potrebbe non avere accesso a PSGallery)
$script:UseCSVFallback = $false

try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "Modulo ImportExcel caricato." -ForegroundColor DarkGray
}
catch {
    Write-Host ("ImportExcel non disponibile ({0})." -f $_.Exception.Message) -ForegroundColor Yellow
    Write-Host ("Verifica che la cartella ImportExcel sia presente in: {0}" -f $psModulesPath) -ForegroundColor Yellow
    Write-Host "Proseguo con output CSV." -ForegroundColor Yellow
    $script:UseCSVFallback = $true
}

# -----------------------------------------------------------------------------
# CREDENZIALI SALVATE IN MODO CIFRATO - FACOLTATIVO
# Dopo l'inserimento delle credenziali lo script chiede se salvarle. Il file
# (Export-Clixml) e cifrato con DPAPI di Windows: si decifra SOLO con lo stesso
# utente Windows sulla stessa macchina (qui: l'account con cui gira la sessione
# elevata). Se il file non e leggibile (altra macchina/utente, VDI ricreata) o
# le credenziali non risultano valide (password cambiata) vengono richieste
# di nuovo e si puo risalvarle. Il file sta in %LOCALAPPDATA%, mai su OneDrive
# o nel repository. Eliminazione: parametro -DimenticaCredenziali.
# -----------------------------------------------------------------------------
function Get-CredenzialiSalvate {
    param(
        [string]$NomeFile,                 # es. 'credenziali_ad.xml'
        [string]$Messaggio,                # testo della finestra Get-Credential
        [scriptblock]$Verifica = $null,    # facoltativo: riceve la credenziale, errore = non valida
        [switch]$Dimentica
    )

    $cartella = Join-Path $env:LOCALAPPDATA 'AD-UserProvisioning'
    $file     = Join-Path $cartella $NomeFile

    if ($Dimentica -and (Test-Path -LiteralPath $file)) {
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        Write-Host "[INFO] Credenziali salvate eliminate ($NomeFile)." -ForegroundColor Green
    }

    if (Test-Path -LiteralPath $file) {
        $salvata = $null
        try { $salvata = Import-Clixml -LiteralPath $file -ErrorAction Stop }
        catch {
            Write-Host "[AVVISO] Credenziali salvate non leggibili (file creato da un altro utente o su un altro PC): vengono richieste di nuovo." -ForegroundColor Yellow
        }
        if ($salvata -is [System.Management.Automation.PSCredential]) {
            Write-Host "[INFO] Credenziali salvate trovate per: $($salvata.UserName)" -ForegroundColor Cyan
            $valida = $true
            if ($Verifica) {
                try { & $Verifica $salvata } catch {
                    $valida = $false
                    Write-Host "[AVVISO] Le credenziali salvate non risultano valide ($($_.Exception.Message))." -ForegroundColor Yellow
                    Write-Host "         Password cambiata? Vengono richieste di nuovo." -ForegroundColor Yellow
                }
            }
            if ($valida) {
                Write-Host "[OK] Uso le credenziali salvate: nessun inserimento richiesto." -ForegroundColor Green
                Write-Host "     (Per eliminarle rilanciare lo script con -DimenticaCredenziali)" -ForegroundColor DarkGray
                return $salvata
            }
        }
    }

    $cred = Get-Credential -Message $Messaggio
    if (-not $cred) { return $null }

    $risp = ''
    do {
        $risp = (Read-Host "Salvare queste credenziali, cifrate e valide solo per questo utente Windows su questo PC, per non doverle reinserire? [S/N]").Trim().ToUpper()
        if ($risp -ne 'S' -and $risp -ne 'N') { Write-Host "Risposta non valida. Digitare S o N." -ForegroundColor Red }
    } while ($risp -ne 'S' -and $risp -ne 'N')
    if ($risp -eq 'S') {
        try {
            if (-not (Test-Path -LiteralPath $cartella)) { New-Item -ItemType Directory -Path $cartella -Force -ErrorAction Stop | Out-Null }
            $cred | Export-Clixml -LiteralPath $file -Force -ErrorAction Stop
            Write-Host "[OK] Credenziali salvate in: $file" -ForegroundColor Green
        } catch {
            Write-Host "[AVVISO] Salvataggio non riuscito: $($_.Exception.Message). Si prosegue senza salvare." -ForegroundColor Yellow
        }
    }
    return $cred
}

# Credenziali AD (obbligatorie)
Write-Host "Inserisci le credenziali AD con permessi per modificare gli account." -ForegroundColor Yellow
$script:Cred = $null
# 2026-10-22: sostituita la richiesta diretta con credenziali salvate facoltative. Originale:
# try { $script:Cred = Get-Credential -Message "Credenziali AD (dominio\utente o utente@dominio)" } catch {}
try {
    $script:Cred = Get-CredenzialiSalvate -NomeFile 'credenziali_expirymanager.xml' `
        -Messaggio "Credenziali AD (dominio\utente o utente@dominio)" -Dimentica:$DimenticaCredenziali `
        -Verifica { param($c) Get-ADDomain -Credential $c -ErrorAction Stop | Out-Null }
} catch {}
if (-not $script:Cred) {
    Write-Host "Nessuna credenziale fornita. Interrotto." -ForegroundColor Red
    exit 1
}

# -------------------------------------------------------
# Funzione: escape apici per filtri LDAP/AD
# -------------------------------------------------------
function Escape-ADFilterValue([string]$v) {
    if ($null -eq $v) { return $v }
    return $v.Replace("'", "''")
}

# -------------------------------------------------------
# Funzione: rimuove i diacritici da una stringa
# (es. "Arnò" -> "Arno", "Müller" -> "Muller")
# Usata come fallback quando la ricerca esatta fallisce
# a causa di problemi di encoding della console (CLM/chcp)
# che corrompono i caratteri accentati letti da Read-Host.
# Non usa classi .NET (incompatibili con CLM): sostituisce
# manualmente i caratteri accentati piu' comuni.
# -------------------------------------------------------
function Remove-Diacritics([string]$s) {
    # Tabella di sostituzione scritta con codici [char] numerici
    # per evitare caratteri non-ASCII nel sorgente del file
    # (che causano errori di parsing quando il file viene salvato
    # con encoding diverso da UTF-8 o letto in ambienti CLM).
    # Formato: coppie (codice-char-accentato, char-base)
    $pairs = @(
        97,  @(224,225,226,227,228,229)   # a: a-grave,acute,circ,tilde,uml,ring
        101, @(232,233,234,235)            # e: e-grave,acute,circ,uml
        105, @(236,237,238,239)            # i: i-grave,acute,circ,uml
        111, @(240,242,243,244,245,246)    # o: eth,o-grave,acute,circ,tilde,uml
        117, @(249,250,251,252)            # u: u-grave,acute,circ,uml
        121, @(253,255)                    # y: y-acute,uml
        110, @(241)                        # n: n-tilde
        99,  @(231)                        # c: c-cedilla
        65,  @(192,193,194,195,196,197)   # A: A-grave,acute,circ,tilde,uml,ring
        69,  @(200,201,202,203)            # E: E-grave,acute,circ,uml
        73,  @(204,205,206,207)            # I: I-grave,acute,circ,uml
        79,  @(210,211,212,213,214)        # O: O-grave,acute,circ,tilde,uml
        85,  @(217,218,219,220)            # U: U-grave,acute,circ,uml
        89,  @(221)                        # Y: Y-acute
        78,  @(209)                        # N: N-tilde
        67,  @(199)                        # C: C-cedilla
    )
    for ($i = 0; $i -lt $pairs.Count; $i += 2) {
        $base    = [char]$pairs[$i]
        $accents = $pairs[$i+1]
        foreach ($code in $accents) {
            $s = $s.Replace([char]$code, $base)
        }
    }
    return $s
}

# -------------------------------------------------------
# Funzione: verifica se il DN dell'utente ricade in una OU
# disabilitati (diretta o sotto-OU)
# -------------------------------------------------------
function Test-IsInDisabledOU([string]$dn) {
    $disabledOUPatterns = @('OU=Disabilitati', 'OU=Disabilitati_Ripristinabili')
    foreach ($pattern in $disabledOUPatterns) {
        if ($dn -match [regex]::Escape($pattern)) { return $true }
    }
    return $false
}

# -------------------------------------------------------
# Funzione: esegue Get-ADUser gestendo separatamente
# AuthenticationException (credenziali errate/scadute).
# Restituisce l'array dei risultati oppure $null se le
# credenziali sono errate e l'operatore sceglie di uscire.
# Imposta $script:CredRejected = $true per segnalare al
# chiamante di interrompere l'elaborazione del batch.
# -------------------------------------------------------
function Invoke-ADSearch {
    param(
        [string]$Filter,
        [string[]]$Properties
        # NON riceve -Credential come parametro: legge sempre $script:Cred
        # direttamente, cosi' le credenziali aggiornate dopo un reinserimento
        # sono usate automaticamente senza doverle ripassare per valore
    )
    try {
        $result = @( Get-ADUser -Credential $script:Cred -Filter $Filter `
            -Properties $Properties -ErrorAction Stop )
        return $result
    }
    catch {
        # Rileva credenziali rifiutate separatamente da altri errori
        if ($_.Exception -is [System.Security.Authentication.AuthenticationException] -or
            $_.FullyQualifiedErrorId -like '*AuthenticationException*' -or
            $_.Exception.Message -like '*rifiutato*' -or
            $_.Exception.Message -like '*rejected*' -or
            $_.Exception.Message -like '*credentials*') {

            Write-Host ""
            Write-Host "  [ERRORE] Il server AD ha rifiutato le credenziali." -ForegroundColor Red
            Write-Host "  Le credenziali potrebbero essere errate o scadute." -ForegroundColor Red
            Write-Host ""
            $retry = Read-Host "  Vuoi inserire nuove credenziali e riprovare? (S/N)"
            if ($retry -match '^[Ss]$') {
                $newCred = $null
                try { $newCred = Get-Credential -Message "Nuove credenziali AD (dominio\utente o utente@dominio)" } catch {}
                if ($newCred) {
                    # Aggiorna $script:Cred: le chiamate successive useranno
                    # automaticamente le nuove credenziali
                    $script:Cred = $newCred
                    Write-Host "  Nuove credenziali acquisite. Riprovo..." -ForegroundColor Yellow
                    try {
                        $result = @( Get-ADUser -Credential $script:Cred -Filter $Filter `
                            -Properties $Properties -ErrorAction Stop )
                        return $result
                    }
                    catch {
                        Write-Host ("  Ancora errore con le nuove credenziali: {0}" -f $_.Exception.Message) -ForegroundColor Red
                        $script:CredRejected = $true
                        return $null
                    }
                } else {
                    Write-Host "  Nessuna credenziale fornita. Interrompo il batch." -ForegroundColor Red
                    $script:CredRejected = $true
                    return $null
                }
            } else {
                Write-Host "  Elaborazione interrotta." -ForegroundColor Red
                $script:CredRejected = $true
                return $null
            }
        }
        # Qualsiasi altro errore (es. utente non trovato): restituisce array vuoto
        return @()
    }
}


# -------------------------------------------------------
# Funzione: come Invoke-ADSearch ma usa -LDAPFilter invece
# di -Filter. Il filtro LDAP supporta wildcard native
# (es. cn=Federico Arn*) che funzionano in modo affidabile
# su tutti i DC, inclusi casi con caratteri accentati.
# Il parametro $LDAPFilter deve essere gia' in formato
# LDAP valido, es: "(cn=Federico Arn*)"
# -------------------------------------------------------
function Invoke-ADSearchLDAP {
    param(
        [string]$LDAPFilter,
        [string[]]$Properties
    )
    try {
        $result = @( Get-ADUser -Credential $script:Cred -LDAPFilter $LDAPFilter `
            -Properties $Properties -ErrorAction Stop )
        return $result
    }
    catch {
        if ($_.Exception -is [System.Security.Authentication.AuthenticationException] -or
            $_.FullyQualifiedErrorId -like '*AuthenticationException*' -or
            $_.Exception.Message -like '*rifiutato*' -or
            $_.Exception.Message -like '*rejected*' -or
            $_.Exception.Message -like '*credentials*') {
            Write-Host "  [ERRORE] Credenziali rifiutate durante la ricerca LDAP." -ForegroundColor Red
            $script:CredRejected = $true
            return $null
        }
        # Altri errori (filtro non valido, ecc.): restituisce array vuoto
        return @()
    }
}

# -------------------------------------------------------
# Funzione: assegna gruppi, sposta in OU, abilita e proroga un utente
# (stesso flusso di Family Web / Due Diligence, parametrizzato).
# Lancia eccezione se Move/Enable/Set falliscono; i gruppi falliti
# non bloccano e vengono restituiti in .Failed.
# -------------------------------------------------------
function Invoke-ProfilazioneUtente {
    param(
        $User,
        [string[]]$GruppiNomi = @(),
        [string[]]$GruppiDN   = @(),
        [string]$OUDN,
        [string]$OUDisplay
    )
    $assigned = @()
    $failed   = @()
    $targets  = @()

    foreach ($n in $GruppiNomi) {
        try {
            $g = Get-ADGroup -Credential $script:Cred -Filter ("Name -eq '{0}'" -f (Escape-ADFilterValue $n)) -ErrorAction Stop | Select-Object -First 1
            if ($g) { $targets += [pscustomobject]@{ Label = $n; DN = $g.DistinguishedName } }
            else {
                $failed += "$n (non trovato)"
                Write-Host ("    Gruppo NON trovato: {0}" -f $n) -ForegroundColor Yellow
            }
        }
        catch {
            $failed += "$n ($($_.Exception.Message))"
            Write-Host ("    Errore cercando {0}: {1}" -f $n, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
    foreach ($dnG in $GruppiDN) {
        $lbl = if ($dnG -match '^CN=([^,]+)') { $Matches[1] } else { $dnG }
        $targets += [pscustomobject]@{ Label = $lbl; DN = $dnG }
    }

    foreach ($t in $targets) {
        try {
            Add-ADGroupMember -Credential $script:Cred -Identity $t.DN -Members $User.ObjectGUID -ErrorAction Stop
            $assigned += $t.Label
            Write-Host ("    Gruppo assegnato: {0}" -f $t.Label) -ForegroundColor Green
        }
        catch {
            $failed += "$($t.Label) ($($_.Exception.Message))"
            Write-Host ("    Errore assegnando {0}: {1}" -f $t.Label, $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    Move-ADObject -Credential $script:Cred -Identity $User.ObjectGUID -TargetPath $OUDN -ErrorAction Stop
    Write-Host ("    Utente spostato in: {0}" -f $OUDisplay) -ForegroundColor Green

    Enable-ADAccount -Credential $script:Cred -Identity $User.ObjectGUID -ErrorAction Stop
    Write-Host ("    Account abilitato: {0}" -f $User.SamAccountName) -ForegroundColor Green

    $before    = $User.accountExpirationDate
    $beforeDT  = if ($before) { [datetime]$before } else { $null }
    $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
    $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')
    Set-ADUser -Credential $script:Cred -Identity $User.ObjectGUID -AccountExpirationDate $ExpiryDate -ErrorAction Stop
    Write-Host ("    Scadenza impostata: {0} -> {1}" -f $beforeTxt, $afterTxt) -ForegroundColor Green

    return [pscustomobject]@{
        Assigned  = $assigned
        Failed    = $failed
        BeforeTxt = $beforeTxt
        AfterTxt  = $afterTxt
    }
}

# -------------------------------------------------------
# Funzione: chiede il SAM di un utente di riferimento e ne valida
# l'idoneita (abilitato, non in OU disabilitati, con gruppi) come
# Get-DatiUtenteEsistente di AD-UserProvisioning.
# Restituisce @{ Gruppi; OU; Sam } oppure $null se l'operatore annulla (INVIO vuoto).
# -------------------------------------------------------
function Read-UtenteRiferimento {
    param([string]$SamEscluso)
    while ($true) {
        $sam = (Read-Host "  SAM utente di riferimento (INVIO = annulla)").Trim()
        if ([string]::IsNullOrWhiteSpace($sam)) { return $null }
        if ($sam -ieq $SamEscluso) {
            Write-Host "  L'utente di riferimento non puo' essere l'utente stesso." -ForegroundColor Red
            continue
        }
        try {
            $ref = Get-ADUser -Credential $script:Cred -Identity $sam `
                -Properties MemberOf, DistinguishedName, Enabled -ErrorAction Stop
            if (-not $ref.Enabled) { throw "L'account '$sam' e' DISABILITATO: usare un utente attivo come riferimento." }
            $refOU = $ref.DistinguishedName -replace '^CN=[^,]+,', ''
            if ($refOU -imatch 'OU=Disabilitati') { throw "L'account '$sam' e' in una OU di disabilitati ('$refOU')." }
            if (-not $ref.MemberOf -or $ref.MemberOf.Count -eq 0) { throw "L'account '$sam' ha solo Domain Users: non e' un riferimento valido." }
            Write-Host ("  Riferimento valido: {0} - OU {1} - {2} gruppi" -f $ref.SamAccountName, $refOU, $ref.MemberOf.Count) -ForegroundColor Green
            return @{ Gruppi = @($ref.MemberOf); OU = $refOU; Sam = $ref.SamAccountName }
        }
        catch {
            Write-Host ("  {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}

function Get-UserGroupNames {
    param($UserObj)
    $groups = @()
    try {
        foreach ($groupDN in $UserObj.MemberOf) {
            if ($groupDN -match '^CN=([^,]+)') { $groups += $Matches[1] }
        }
    }
    catch {}
    return $groups
}

# ========================
# 1. DATA DI SCADENZA AUTOMATICA
# ========================

# Calcolata una volta sola all'avvio (6 mesi da oggi, fine giornata 23:59:59)
$ExpiryDate = (Get-Date).AddMonths(6).Date.AddDays(1).AddSeconds(-1)
Write-Host ("Data impostata automaticamente a +6 mesi: {0}" -f $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')) -ForegroundColor Green

# ========================
# 2. PERCORSI LOG (un solo file per tutta la sessione, anche se il loop riparte)
# ========================

$timestamp         = (Get-Date).ToString('yyyyMMdd_HHmmss')
# Sotto RunAs GetFolderPath('Desktop') restituisce il Desktop dell'account admin:
# si usa il profilo dell'utente originale; fallback sul Desktop corrente.
$Desktop = Join-Path $OriginalUserProfile 'Desktop'
if (-not (Test-Path $Desktop)) { $Desktop = [Environment]::GetFolderPath('Desktop') }

# Cartella log: quella dello script se scrivibile, altrimenti Desktop dell'utente originale
$logDir = $Desktop
if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
    $probe = Join-Path $PSScriptRoot ("_probe_{0}.tmp" -f $timestamp)
    try {
        Set-Content -Path $probe -Value "" -ErrorAction Stop
        Remove-Item $probe -ErrorAction SilentlyContinue
        $logDir = $PSScriptRoot
    } catch {
        Write-Host "Cartella dello script non scrivibile: log sul Desktop." -ForegroundColor Yellow
    }
}
$Desktop = $logDir
Write-Host ("Log di sessione: {0}" -f $Desktop) -ForegroundColor DarkGray
$xlsxPath          = Join-Path $Desktop ("AD_Expiry_Log_{0}.xlsx" -f $timestamp)
$csvPath           = Join-Path $Desktop ("AD_Expiry_Log_{0}.csv"  -f $timestamp)
$script:logs       = @()   # accumula su tutti i giri del loop

# -------------------------------------------------------
# Funzione: aggiunge una riga al log
# -------------------------------------------------------
function Add-Log {
    param(
        [string]$InputText,
        [string]$Sam        = "",
        [string]$UPN        = "",
        [string]$DN         = "",
        [string]$PrevExpiry = "",
        [string]$NewExpiry  = "",
        [string]$Status,
        [string]$Message    = "",
        [string]$IsEnabled  = "",
        [string]$Groups     = "",
        [string]$OU_Flag    = ""
    )
    $script:logs += [pscustomobject]@{
        Timestamp              = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Input                  = $InputText
        SamAccountName         = $Sam
        UserPrincipalName      = $UPN
        DistinguishedName      = $DN
        IsEnabled              = $IsEnabled
        Groups                 = $Groups
        OU_Flag                = $OU_Flag
        PreviousExpirationDate = $PrevExpiry
        NewExpirationDate      = $NewExpiry
        Status                 = $Status
        Message                = $Message
    }
}

# OU di destinazione per riattivazione utenti disabilitati
$targetOUs = @(
    # DN corretti in base alla struttura reale del dominio:
    # dominio: DC=intranet,DC=ifis,DC=local
    # le sigle (IFIS, IFISNplServicing, ecc.) sono OU, non domini separati
    [pscustomobject]@{ Label = "1"; Display = "\IFIS\Users\External";                    DN = "OU=External,OU=Users,OU=IFIS,DC=intranet,DC=ifis,DC=local" },
    [pscustomobject]@{ Label = "2"; Display = "\IFISNplServicing\Users\External  (FAMILY WEB)"; DN = "OU=External,OU=Users,OU=IFISNplServicing,DC=intranet,DC=ifis,DC=local" },
    [pscustomobject]@{ Label = "3"; Display = "\IfisNplInvesting\Users\External";        DN = "OU=External,OU=Users,OU=IfisNplInvesting,DC=intranet,DC=ifis,DC=local" },
    [pscustomobject]@{ Label = "4"; Display = "\Illimity\Users\External";                DN = "OU=External,OU=Users,OU=Illimity,DC=intranet,DC=ifis,DC=local" },
    [pscustomobject]@{ Label = "5"; Display = "\ConsulentiEsterni\DueDiligence";         DN = "OU=DueDiligence,OU=ConsulentiEsterni,DC=intranet,DC=ifis,DC=local" },
    [pscustomobject]@{ Label = "6"; Display = "\ConsulentiEsterni\Legal Factory";        DN = "OU=Legal Factory,OU=ConsulentiEsterni,DC=intranet,DC=ifis,DC=local" }
)

# Profili di AD-UserProvisioning proposti agli utenti senza gruppi che non sono
# ne' Family Web ne' Due Diligence. Copia dei fogli Config_OU_Profilo e
# Config_Gruppi_Profilo di Template_Creazione_Utenti_AD (v2.52 del 2026-10-21):
# se i profili cambiano nel template vanno aggiornati anche qui.
# Esclusi di proposito: agenti (scadenza "Mai" + attributi propri),
# consulente_societa_recupero (nessun gruppo configurato nel template),
# dipendente_* (modalita' dipendente disattivata nel progetto).
$profiliProvisioning = @(
    [pscustomobject]@{
        Label = "1"; Key = "esterno_gruppi_default"
        OUDN = "OU=External,OU=Users,OU=IFIS,DC=intranet,DC=ifis,DC=local"
        OUDisplay = "\IFIS\Users\External"
        Gruppi = @("app.o365.ssrp","O365.AccessPolicy.NoDownload","o365.mfa.ifisuser.oneday","prisma.ext","ext.users.automationscript")
    },
    [pscustomobject]@{
        Label = "2"; Key = "commesse"
        OUDN = "OU=Legal Factory,OU=ConsulentiEsterni,DC=intranet,DC=ifis,DC=local"
        OUDisplay = "\ConsulentiEsterni\Legal Factory"
        Gruppi = @("cbox.npl.ro.prd","app.commesse20.extview","app.o365.ssrp","cbox.npl.prd","ext.users.automationscript","GPO.NoLocalAndRDPLogon.AppUser")
    },
    [pscustomobject]@{
        Label = "3"; Key = "legal_factory"
        OUDN = "OU=Legal Factory,OU=ConsulentiEsterni,DC=intranet,DC=ifis,DC=local"
        OUDisplay = "\ConsulentiEsterni\Legal Factory"
        Gruppi = @("app.legalfactory20.ifis4you","app.legalfactory20.legaleesterno","app.o365.ssrp","app.smartransaction.frontoffice","cbox.npl.prd","ext.users.automationscript","GPO.NoLocalAndRDPLogon.AppUser","nac.LegaliEsterni")
    }
)

# ========================
# LOOP PRINCIPALE
# Riparte finche' l'operatore non sceglie di uscire
# ========================
$continueLoop = $true

while ($continueLoop) {

    # --------------------------------------------------
    # INPUT LISTA UTENTI (tre INVIO consecutivi per terminare)
    # Tre INVIO perche' le liste incollate possono contenere
    # fino a due righe vuote consecutive come separatori interni
    # --------------------------------------------------
    Write-Host ""
    Write-Host "Incolla la lista degli utenti (SAM, UPN o Nome Cognome)." -ForegroundColor Yellow
    Write-Host "Puoi includere righe vuote come separatori." -ForegroundColor Yellow
    Write-Host "Per TERMINARE la lista premi INVIO TRE VOLTE su riga vuota." -ForegroundColor Yellow
    Write-Host ""

    $inputLines  = @()
    $blankStreak = 0

    while ($true) {
        $line = Read-Host
        if ([string]::IsNullOrWhiteSpace($line)) {
            $blankStreak++
            if ($blankStreak -ge 3) { break }
            $inputLines += ""
            continue
        }
        $blankStreak = 0
        $inputLines += $line.Trim()
    }

    # Rimuove righe vuote finali
    while ($inputLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($inputLines[-1])) {
        $inputLines = $inputLines[0..($inputLines.Count - 2)]
    }

    $hasInput = ($inputLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -gt 0

    if (-not $hasInput) {
        Write-Host "Nessun input fornito per questo giro." -ForegroundColor Yellow
    }
    else {

        Write-Host ""
        $batchLogStart = $script:logs.Count
        Write-Host "--- INIZIO ELABORAZIONE ---" -ForegroundColor Cyan
        Write-Host ""

        # ========================
        # 4. ELABORAZIONE
        # ========================

        # $script:extraQueue non serve piu': usiamo un ArrayList mutabile
        # con indice numerico cosi' i SAM extra da selezione multipla
        # vengono aggiunti e processati nello stesso ciclo senza loop separati
        # Deduplicazione: rimuove righe identiche (case-insensitive) mantenendo l'ordine
        # Evita che lo stesso Nome Cognome inserito due volte generi elaborazioni doppie
        $seenLines = @{}
        $dedupedLines = @()
        foreach ($line in $inputLines) {
            $key = $line.Trim().ToLower()
            if (-not [string]::IsNullOrWhiteSpace($key) -and -not $seenLines.ContainsKey($key)) {
                $seenLines[$key] = $true
                $dedupedLines += $line
            } elseif ([string]::IsNullOrWhiteSpace($line)) {
                $dedupedLines += $line  # conserva righe vuote come separatori
            }
        }
        $inputLines = $dedupedLines

        $workList = New-Object System.Collections.ArrayList
        foreach ($line in $inputLines) { [void]$workList.Add($line) }

        # Memorizza la dimensione iniziale: quando $wi supera questa soglia
        # si entra nella fase di elaborazione degli utenti accodati (multi-select)
        $originalWorkCount = $workList.Count
        $separatorShown    = $false

        $wi = 0
        while ($wi -lt $workList.Count) {
            $raw = $workList[$wi]
            $wi++
            $entry = $raw.Trim()
            if ([string]::IsNullOrWhiteSpace($entry)) { continue }

            # Separatore visivo tra utenti della lista originale e utenti accodati
            if ($wi -gt ($originalWorkCount + 1) -and -not $separatorShown) {
                $separatorShown = $true
                Write-Host ""
                Write-Host "=============================================" -ForegroundColor DarkCyan
                Write-Host "  PROROGA UTENTI DA SELEZIONE MULTIPLA" -ForegroundColor DarkCyan
                Write-Host "=============================================" -ForegroundColor DarkCyan
                Write-Host ""
            }

            Write-Host ("Ricerca: {0}" -f $entry) -ForegroundColor DarkCyan

            $foundUser = $null
            # Flag credenziali rifiutate: viene impostato da Invoke-ADSearch
            $script:CredRejected = $false

            # --- Ricerca: Caso 1 UPN ---
            if ($entry -match '@') {
                $val = Escape-ADFilterValue $entry
                $adMatches = Invoke-ADSearch `
                    -Filter "UserPrincipalName -eq '$val'" `
                    -Properties @('accountExpirationDate','Enabled','MemberOf','DistinguishedName')
                if ($script:CredRejected) { break }
                if ($null -eq $adMatches) { $adMatches = @() }
                if (($adMatches | Measure-Object).Count -eq 1) {
                    $foundUser = $adMatches | Select-Object -First 1
                }
            }

            # --- Ricerca: Caso 2 SAM (nessuno spazio) ---
            if (-not $foundUser -and -not $script:CredRejected -and $entry -notmatch '\s') {
                $val = Escape-ADFilterValue $entry
                $adMatches = Invoke-ADSearch `
                    -Filter "SamAccountName -eq '$val'" `
                    -Properties @('accountExpirationDate','Enabled','MemberOf','DistinguishedName')
                if ($script:CredRejected) { break }
                if ($null -eq $adMatches) { $adMatches = @() }
                if (($adMatches | Measure-Object).Count -eq 1) {
                    $foundUser = $adMatches | Select-Object -First 1
                }
            }

            # --- Ricerca: Caso 3 Nome Cognome ---
            if (-not $foundUser -and -not $script:CredRejected -and $entry -match '\s+') {
                $parts   = $entry -split '\s+', 2
                $nome    = $parts[0]
                $cognome = $parts[1]

                # Tentativo 1: GivenName + Surname senza diacritici
                # Il test manuale conferma che AD tratta 'Arno' e 'Arno' (accentato)
                # come equivalenti nel filtro PowerShell (accent-insensitive)
                $nomeNoAcc    = Remove-Diacritics $nome
                $cognomeNoAcc = Remove-Diacritics $cognome
                $nomeEsc      = Escape-ADFilterValue $nomeNoAcc
                $cognomeEsc   = Escape-ADFilterValue $cognomeNoAcc
                $adMatches = Invoke-ADSearch `
                    -Filter "GivenName -eq '$nomeEsc' -and Surname -eq '$cognomeEsc'" `
                    -Properties @('accountExpirationDate','Enabled','MemberOf','DistinguishedName')
                if ($script:CredRejected) { break }
                if ($null -eq $adMatches) { $adMatches = @() }
                $adCount = ($adMatches | Measure-Object).Count
                Write-Host ("  T1 GivenName/Surname '{0} {1}': {2} risultati" -f $nomeNoAcc, $cognomeNoAcc, $adCount) -ForegroundColor DarkGray

                # Tentativo 2: LDAP cn=* come fallback se T1 non trova nulla
                if ($adCount -eq 0 -and -not $script:CredRejected) {
                    $cnPattern = "$nomeNoAcc $cognomeNoAcc*"
                    Write-Host ("  T2 LDAP cn='{0}*'" -f "$nomeNoAcc $cognomeNoAcc") -ForegroundColor DarkGray
                    $adMatches = Invoke-ADSearchLDAP `
                        -LDAPFilter "(cn=$cnPattern)" `
                        -Properties @('accountExpirationDate','Enabled','MemberOf','DistinguishedName')
                    if ($script:CredRejected) { break }
                    if ($null -eq $adMatches) { $adMatches = @() }
                    $adCount = ($adMatches | Measure-Object).Count
                    Write-Host ("  T2: {0} risultati" -f $adCount) -ForegroundColor DarkGray
                }

                # Tentativo 3: LDAP ordine inverso Cognome*Nome*
                if ($adCount -eq 0 -and -not $script:CredRejected) {
                    $cnPattern = "$cognomeNoAcc*$nomeNoAcc*"
                    Write-Host ("  T3 LDAP ordine inverso: cn='{0}*'" -f "$cognomeNoAcc*$nomeNoAcc") -ForegroundColor DarkGray
                    $adMatches = Invoke-ADSearchLDAP `
                        -LDAPFilter "(cn=$cnPattern)" `
                        -Properties @('accountExpirationDate','Enabled','MemberOf','DistinguishedName')
                    if ($script:CredRejected) { break }
                    if ($null -eq $adMatches) { $adMatches = @() }
                    $adCount = ($adMatches | Measure-Object).Count
                    Write-Host ("  T3: {0} risultati" -f $adCount) -ForegroundColor DarkGray
                }

                if ($adCount -eq 1) {
                    $foundUser = $adMatches | Select-Object -First 1
                }
                elseif ($adCount -gt 1) {
                    # Piu' utenti trovati: mostra elenco numerato
                    # L'operatore puo' selezionarne uno o piu' separati da virgola
                    Write-Host ""
                    Write-Host ("  Trovati {0} utenti per '{1}':" -f $adCount, $entry) -ForegroundColor Yellow

                    $adList = @($adMatches)
                    for ($i = 0; $i -lt $adCount; $i++) {
                        $u = $adList[$i]
                        $uEnabled = if ($u.Enabled) { "Abilitato" } else { "DISABILITATO" }
                        Write-Host ("  [{0}] SAM: {1,-20} | UPN: {2,-35} | Nome: {3} | {4}" -f `
                            ($i + 1), $u.SamAccountName, $u.UserPrincipalName, $u.Name, $uEnabled) -ForegroundColor Yellow
                    }
                    Write-Host ("  [0] Salta - non prorogare nessuno") -ForegroundColor Gray
                    Write-Host ""

                    # Accetta una o piu' scelte separate da virgola, con spazi opzionali
                    # es. "1", "1,2", "1, 2", "1 ,2 , 3"
                    $selectedUsers = $null
                    do {
                        $choiceInput = Read-Host ("  Scegli (es. 1 oppure 1,2) o 0 per saltare")

                        # Divide su virgola con spazi opzionali usando -split e regex
                        # Piu' robusto di -replace in CLM
                        $tokens = @($choiceInput.Trim() -split '\s*,\s*' | Where-Object { $_ -ne '' })

                        # Caso: scelta singola "0" = salta tutti
                        if ($tokens.Count -eq 1 -and $tokens[0].Trim() -eq '0') {
                            $selectedUsers = @()
                            break
                        }

                        # Valida ogni token: deve essere un numero tra 1 e $adCount
                        $valid = $true
                        $chosenIndices = @()
                        foreach ($token in $tokens) {
                            $t = $token.Trim()
                            if ($t -match '^\d+$') {
                                $n = [int]$t
                                if ($n -ge 1 -and $n -le $adCount) {
                                    $chosenIndices += $n
                                } else {
                                    $valid = $false
                                    break
                                }
                            } else {
                                $valid = $false
                                break
                            }
                        }

                        if ($valid -and $chosenIndices.Count -gt 0) {
                            # Rimuove duplicati mantenendo l'ordine
                            $seen = @{}
                            $unique = @()
                            foreach ($idx in $chosenIndices) {
                                if (-not $seen.ContainsKey($idx)) {
                                    $seen[$idx] = $true
                                    $unique += $idx
                                }
                            }
                            $selectedUsers = $unique
                        } else {
                            Write-Host ("  Valore non valido. Inserisci numeri tra 1 e {0} separati da virgola, o 0 per saltare." -f $adCount) -ForegroundColor Red
                            $selectedUsers = $null
                        }
                    } while ($null -eq $selectedUsers)

                    if ($selectedUsers.Count -eq 0) {
                        Write-Host ("  '{0}' saltato per scelta operatore." -f $entry) -ForegroundColor Gray
                        Add-Log -InputText $entry -Status "Skipped_ByOperator" `
                            -Message "Ambiguous (Nome Cognome): operatore ha scelto di non selezionare nessun candidato"
                        continue
                    }

                    if ($selectedUsers.Count -eq 1) {
                        # Selezione singola: assegna $foundUser e lascia proseguire il flusso normale
                        $foundUser = $adList[$selectedUsers[0] - 1]
                        Write-Host ("  Selezionato: {0} ({1})" -f $foundUser.Name, $foundUser.SamAccountName) -ForegroundColor Green
                    } else {
                        # Selezione multipla: aggiunge i SAM alla workList
                        # che e' un ArrayList mutabile: il while ($wi -lt $workList.Count)
                        # li elaborera' nei giri successivi con tutti i controlli completi
                        Write-Host ("  Selezionati {0} utenti. Li accodo per l'elaborazione:" -f $selectedUsers.Count) -ForegroundColor Green
                        foreach ($idx in $selectedUsers) {
                            $u = $adList[$idx - 1]
                            Write-Host ("    - {0} ({1})" -f $u.Name, $u.SamAccountName) -ForegroundColor Green
                            # Aggiunge solo se il SAM non e' gia' in workList (deduplicazione)
                            $samLower = $u.SamAccountName.ToLower()
                            $alreadyInList = $false
                            foreach ($existing in $workList) {
                                if ($existing.Trim().ToLower() -eq $samLower) { $alreadyInList = $true; break }
                            }
                            if (-not $alreadyInList) {
                                [void]$workList.Add($u.SamAccountName)
                            }
                        }
                        Add-Log -InputText $entry -Status "Skipped_MultiSelect" `
                            -Message ("Ambiguous: selezionati {0} SAM accodati per elaborazione" -f $selectedUsers.Count)
                        continue
                    }
                }
            }

            # --- Utente non trovato ---
            if (-not $foundUser) {
                Write-Host ("Utente non trovato: {0}" -f $entry) -ForegroundColor Red
                Add-Log -InputText $entry -Status "NotFound" -Message "User not found"
                continue
            }

            # --------------------------------------------------
            # CONTROLLI: stato account, gruppi, OU
            # --------------------------------------------------

            # Stato abilitazione
            $isEnabled  = $foundUser.Enabled
            $enabledTxt = if ($isEnabled) { "Yes" } else { "No" }

            # Gruppi extra (MemberOf non include Domain Users = gruppo primario implicito)
            $userGroups         = Get-UserGroupNames -UserObj $foundUser
            $groupsTxt          = if ($userGroups.Count -gt 0) { $userGroups -join "; " } else { "(nessun gruppo extra)" }
            $hasOnlyDomainUsers = ($userGroups.Count -eq 0)

            # OU
            $dn           = $foundUser.DistinguishedName
            $inDisabledOU = Test-IsInDisabledOU -dn $dn

            $ouFlag = "OK"
            if ($dn -match [regex]::Escape("OU=Disabilitati_Ripristinabili")) {
                $ouFlag = "OU_Disabilitati_Ripristinabili"
            } elseif ($dn -match [regex]::Escape("OU=Disabilitati")) {
                $ouFlag = "OU_Disabilitati"
            }

            # --------------------------------------------------
            # RIEPILOGO VISIVO per l'operatore
            # Mostra sempre stato, OU e gruppi per ogni utente trovato
            # --------------------------------------------------
            Write-Host ""
            Write-Host ("  +-- {0} ({1})" -f $foundUser.Name, $foundUser.SamAccountName) -ForegroundColor White

            # Riga stato account
            if ($isEnabled) {
                Write-Host "  |  Stato    : Abilitato" -ForegroundColor Green
            } else {
                Write-Host "  |  Stato    : DISABILITATO" -ForegroundColor Red
            }

            # Riga OU: mostra la parte significativa del DN (senza il CN dell'utente)
            $ouPath = $dn -replace '^CN=[^,]+,', ''
            if ($inDisabledOU) {
                Write-Host ("  |  OU       : {0}  [OU disabilitati]" -f $ouFlag) -ForegroundColor Yellow
            } else {
                Write-Host ("  |  OU       : {0}" -f $ouPath) -ForegroundColor Green
            }

            # Riga gruppi
            if ($hasOnlyDomainUsers) {
                Write-Host "  |  Gruppi   : solo Domain Users  [nessun gruppo extra - proroga bloccata]" -ForegroundColor Yellow
            } else {
                Write-Host ("  |  Gruppi   : {0}" -f $groupsTxt) -ForegroundColor Cyan
            }

            Write-Host "  +------------------------------------------------" -ForegroundColor DarkGray
            Write-Host ""

            # ==================================================
            # CASO A: solo Domain Users
            #   - Chiede se e' un utente Family Web
            #   - Se SI: assegna i 6 gruppi standard, sposta in OU
            #     IFISNplServicing\Users\External, abilita, proroga
            #   - Se NO: blocca la proroga e logga
            # ==================================================
            if ($hasOnlyDomainUsers) {

                Write-Host ""
                $isFamilyWeb = Read-Host "  Il ticket indica che si tratta di un utente FAMILY WEB? (S/N)"

                # --- Caso A-NO: non e' Family Web -> chiede se e' Due Diligence ---
                if ($isFamilyWeb -notmatch '^[Ss]$') {

                    Write-Host ""
                    $isDueDiligence = Read-Host "  Il ticket indica che si tratta di un utente DUE DILIGENCE? (S/N)"

                    # --- Caso A-NO-NO: non e' ne' Family Web ne' Due Diligence ->
                    #     propone gli altri profili di AD-UserProvisioning; 0 = blocco proroga ---
                    if ($isDueDiligence -notmatch '^[Ss]$') {

                        Write-Host ""
                        Write-Host "  Vuoi profilarlo con un altro profilo di AD-UserProvisioning?" -ForegroundColor Cyan
                        foreach ($pf in $profiliProvisioning) {
                            Write-Host ("  [{0}] {1,-24} -> {2}" -f $pf.Label, $pf.Key, $pf.OUDisplay) -ForegroundColor Cyan
                        }
                        $labelEsistente = [string]($profiliProvisioning.Count + 1)
                        Write-Host ("  [{0}] {1,-24} -> copia OU e gruppi da un utente di riferimento" -f $labelEsistente, "utente_esistente") -ForegroundColor Cyan
                        Write-Host "  [0] Nessuno - blocca la proroga" -ForegroundColor Gray
                        Write-Host ""

                        $profScelta = $null
                        $profSceltaTxt = ''
                        do {
                            $profScelta = $null
                            $profSceltaTxt = (Read-Host ("  Scelta (0-{0})" -f $labelEsistente)).Trim()
                            if ($profSceltaTxt -eq '0' -or $profSceltaTxt -eq $labelEsistente) { break }
                            $profScelta = $profiliProvisioning | Where-Object { $_.Label -eq $profSceltaTxt }
                            if (-not $profScelta) { Write-Host "  Scelta non valida. Riprova." -ForegroundColor Red }
                        } while (-not $profScelta)

                        # Profilo predefinito scelto, oppure utente_esistente con riferimento valido
                        $profGruppiNomi = @()
                        $profGruppiDN   = @()
                        $profOUDN       = ''
                        $profOUDisplay  = ''
                        $profKey        = ''
                        if ($profScelta) {
                            $profGruppiNomi = @($profScelta.Gruppi)
                            $profOUDN       = $profScelta.OUDN
                            $profOUDisplay  = $profScelta.OUDisplay
                            $profKey        = $profScelta.Key
                        }
                        elseif ($profSceltaTxt -eq $labelEsistente) {
                            $refDati = Read-UtenteRiferimento -SamEscluso $foundUser.SamAccountName
                            if ($refDati) {
                                $profGruppiDN  = @($refDati.Gruppi)
                                $profOUDN      = $refDati.OU
                                $profOUDisplay = $refDati.OU
                                $profKey       = "utente_esistente (rif. {0})" -f $refDati.Sam
                            }
                        }

                        # --- Profilo scelto: assegna gruppi, sposta, abilita, proroga ---
                        if ($profKey) {
                            Write-Host ""
                            Write-Host ("  Profilo '{0}' confermato. Procedo con assegnazione gruppi, spostamento OU, abilitazione e proroga." -f $profKey) -ForegroundColor Cyan
                            try {
                                $esito = Invoke-ProfilazioneUtente -User $foundUser `
                                    -GruppiNomi $profGruppiNomi -GruppiDN $profGruppiDN `
                                    -OUDN $profOUDN -OUDisplay $profOUDisplay

                                $groupsAssignedTxt = if ($esito.Assigned.Count -gt 0) { $esito.Assigned -join "; " } else { "(nessuno)" }
                                $msgLog = "Profilo {0}: gruppi assegnati=[{1}]" -f $profKey, $groupsAssignedTxt
                                if ($esito.Failed.Count -gt 0) { $msgLog += " | ATTENZIONE gruppi falliti=[" + ($esito.Failed -join "; ") + "]" }

                                Write-Host ("  {0}: profilazione completata. {1}" -f $foundUser.SamAccountName, $msgLog) -ForegroundColor Green

                                Add-Log `
                                    -InputText  $entry `
                                    -Sam        $foundUser.SamAccountName `
                                    -UPN        $foundUser.UserPrincipalName `
                                    -DN         $profOUDN `
                                    -IsEnabled  "Yes (riattivato)" `
                                    -Groups     $groupsAssignedTxt `
                                    -OU_Flag    ("Moved_to: {0}" -f $profOUDisplay) `
                                    -PrevExpiry $esito.BeforeTxt `
                                    -NewExpiry  $esito.AfterTxt `
                                    -Status     "Profiled_Completed" `
                                    -Message    $msgLog
                            }
                            catch {
                                $errMsg = $_ | Out-String
                                Write-Host ("  Errore nella profilazione '{0}' per '{1}': {2}" -f $profKey, $entry, $errMsg) -ForegroundColor Red
                                Add-Log `
                                    -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                                    -DN $dn -IsEnabled $enabledTxt -Groups "(solo Domain Users)" -OU_Flag $ouFlag `
                                    -PrevExpiry "" -NewExpiry "" -Status "Error_Profilo" -Message $errMsg
                            }
                            continue
                        }

                        # Nessun profilo scelto (0 o riferimento annullato) -> blocco proroga
                        Write-Host ("  [BLOCCATO] {0}: solo Domain Users, non e' Family Web ne' Due Diligence. Proroga non consentita." -f $foundUser.SamAccountName) -ForegroundColor Red
                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $dn -IsEnabled $enabledTxt -Groups "(solo Domain Users)" -OU_Flag $ouFlag `
                            -PrevExpiry "" -NewExpiry "" -Status "Blocked_NoGroups" `
                            -Message "Solo Domain Users, non Family Web ne' Due Diligence: proroga non consentita"
                        continue
                    }

                    # --- Caso A-NO-SI: e' Due Diligence -> assegna gruppi, sposta, abilita, proroga ---
                    # Gruppi da assegnare agli utenti Due Diligence privi di gruppi
                    $dueDiligenceGroups = @(
                        "ext.users.automationscript",
                        "GPO.NoLocalAndRDPLogon.AppUser",
                        "app.o365.ssrp",
                        "cbox.npl.ro.prd",
                        "nac.Consulenti",
                        "cbox.npl.prd",
                        "app.commesse20.duediligencevendita.op"
                    )
                    # OU di destinazione fissa per Due Diligence
                    $dueDiligenceOU        = "OU=DueDiligence,OU=ConsulentiEsterni,DC=intranet,DC=ifis,DC=local"
                    $dueDiligenceOUDisplay = "\ConsulentiEsterni\DueDiligence"

                    Write-Host ""
                    Write-Host ("  Utente Due Diligence confermato. Procedo con assegnazione gruppi, spostamento OU, abilitazione e proroga.") -ForegroundColor Cyan

                    try {
                        # 1. Assegna i gruppi Due Diligence
                        # Usa $foundUser.ObjectGUID come identity del membro:
                        # il GUID e' immutabile e non soffre di cache stale come il DN
                        $groupsAssigned  = @()
                        $groupsFailed    = @()
                        foreach ($grpName in $dueDiligenceGroups) {
                            try {
                                $grpObj = Get-ADGroup -Credential $script:Cred -Filter "Name -eq '$grpName'" -ErrorAction Stop
                                if ($grpObj) {
                                    Add-ADGroupMember -Credential $script:Cred -Identity $grpObj.DistinguishedName `
                                        -Members $foundUser.ObjectGUID -ErrorAction Stop
                                    $groupsAssigned += $grpName
                                    Write-Host ("    Gruppo assegnato: {0}" -f $grpName) -ForegroundColor Green
                                } else {
                                    $groupsFailed += "$grpName (non trovato)"
                                    Write-Host ("    Gruppo NON trovato: {0}" -f $grpName) -ForegroundColor Yellow
                                }
                            }
                            catch {
                                $groupsFailed += "$grpName ($($_.Exception.Message))"
                                Write-Host ("    Errore assegnando {0}: {1}" -f $grpName, $_.Exception.Message) -ForegroundColor Yellow
                            }
                        }

                        # 2. Sposta nella OU ConsulentiEsterni\DueDiligence
                        # Usa ObjectGUID come -Identity: il DN potrebbe essere stale
                        # dopo le operazioni di gruppo (problema di cache AD)
                        Move-ADObject -Credential $script:Cred -Identity $foundUser.ObjectGUID -TargetPath $dueDiligenceOU -ErrorAction Stop
                        Write-Host ("    Utente spostato in: {0}" -f $dueDiligenceOUDisplay) -ForegroundColor Green

                        # 3. Abilita l'account (usa ObjectGUID, stabile dopo Move)
                        Enable-ADAccount -Credential $script:Cred -Identity $foundUser.ObjectGUID -ErrorAction Stop
                        Write-Host ("    Account abilitato: {0}" -f $foundUser.SamAccountName) -ForegroundColor Green

                        # 4. Proroga di 6 mesi
                        $before    = $foundUser.accountExpirationDate
                        $beforeDT  = if ($before) { [datetime]$before } else { $null }
                        $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                        $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                        Set-ADUser -Credential $script:Cred -Identity $foundUser.ObjectGUID `
                            -AccountExpirationDate $ExpiryDate -ErrorAction Stop
                        Write-Host ("    Scadenza impostata: {0} -> {1}" -f $beforeTxt, $afterTxt) -ForegroundColor Green

                        # Riepilogo finale per l'operatore
                        $groupsAssignedTxt = if ($groupsAssigned.Count -gt 0) { $groupsAssigned -join "; " } else { "(nessuno)" }
                        $groupsFailedTxt   = if ($groupsFailed.Count -gt 0)   { $groupsFailed -join "; " }   else { "" }
                        $msgLog = "Due Diligence: gruppi assegnati=[$groupsAssignedTxt]"
                        if ($groupsFailedTxt) { $msgLog += " | ATTENZIONE gruppi falliti=[$groupsFailedTxt]" }

                        Write-Host ("  {0}: Due Diligence completato. {1}" -f $foundUser.SamAccountName, $msgLog) -ForegroundColor Green

                        Add-Log `
                            -InputText  $entry `
                            -Sam        $foundUser.SamAccountName `
                            -UPN        $foundUser.UserPrincipalName `
                            -DN         $dueDiligenceOU `
                            -IsEnabled  "Yes (riattivato)" `
                            -Groups     $groupsAssignedTxt `
                            -OU_Flag    ("Moved_to: {0}" -f $dueDiligenceOUDisplay) `
                            -PrevExpiry $beforeTxt `
                            -NewExpiry  $afterTxt `
                            -Status     "DueDiligence_Completed" `
                            -Message    $msgLog
                    }
                    catch {
                        $errMsg = $_ | Out-String
                        Write-Host ("  Errore nel workflow Due Diligence per '{0}': {1}" -f $entry, $errMsg) -ForegroundColor Red
                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $dn -IsEnabled $enabledTxt -Groups "(solo Domain Users)" -OU_Flag $ouFlag `
                            -PrevExpiry "" -NewExpiry "" -Status "Error_DueDiligence" -Message $errMsg
                    }
                    continue
                }

                # --- Caso A-SI: e' Family Web -> assegna gruppi, sposta, abilita, proroga ---
                # Gruppi da assegnare agli utenti Family Web privi di gruppi
                $familyWebGroups = @(
                    "app.o365.ssrp",
                    "cbox.npl.prd",
                    "cbox.npls.prd",
                    "Citrix.MFA.O365",
                    "nac.Consulenti",
                    "GPO.NoLocalAndRDPLogon.AppUser"
                )
                # OU di destinazione fissa per Family Web
                # DN corretto: IFISNplServicing e' una OU nel dominio intranet.ifis.local
                $familyWebOU        = "OU=External,OU=Users,OU=IFISNplServicing,DC=intranet,DC=ifis,DC=local"
                $familyWebOUDisplay = "\IFISNplServicing\Users\External"

                Write-Host ""
                Write-Host ("  Utente Family Web confermato. Procedo con assegnazione gruppi, spostamento OU, abilitazione e proroga.") -ForegroundColor Cyan

                try {
                    # 1. Assegna i gruppi Family Web
                    # Usa $foundUser.ObjectGUID come identity del membro:
                    # il GUID e' immutabile e non soffre di cache stale come il DN
                    $groupsAssigned  = @()
                    $groupsFailed    = @()
                    foreach ($grpName in $familyWebGroups) {
                        try {
                            $grpObj = Get-ADGroup -Credential $script:Cred -Filter "Name -eq '$grpName'" -ErrorAction Stop
                            if ($grpObj) {
                                Add-ADGroupMember -Credential $script:Cred -Identity $grpObj.DistinguishedName `
                                    -Members $foundUser.ObjectGUID -ErrorAction Stop
                                $groupsAssigned += $grpName
                                Write-Host ("    Gruppo assegnato: {0}" -f $grpName) -ForegroundColor Green
                            } else {
                                $groupsFailed += "$grpName (non trovato)"
                                Write-Host ("    Gruppo NON trovato: {0}" -f $grpName) -ForegroundColor Yellow
                            }
                        }
                        catch {
                            $groupsFailed += "$grpName ($($_.Exception.Message))"
                            Write-Host ("    Errore assegnando {0}: {1}" -f $grpName, $_.Exception.Message) -ForegroundColor Yellow
                        }
                    }

                    # 2. Sposta nella OU IFISNplServicing\Users\External
                    # Usa ObjectGUID come -Identity: il DN potrebbe essere stale
                    # dopo le operazioni di gruppo (problema di cache AD)
                    Move-ADObject -Credential $script:Cred -Identity $foundUser.ObjectGUID -TargetPath $familyWebOU -ErrorAction Stop
                    Write-Host ("    Utente spostato in: {0}" -f $familyWebOUDisplay) -ForegroundColor Green

                    # 3. Abilita l'account (usa ObjectGUID, stabile dopo Move)
                    Enable-ADAccount -Credential $script:Cred -Identity $foundUser.ObjectGUID -ErrorAction Stop
                    Write-Host ("    Account abilitato: {0}" -f $foundUser.SamAccountName) -ForegroundColor Green

                    # 4. Proroga di 6 mesi
                    $before    = $foundUser.accountExpirationDate
                    $beforeDT  = if ($before) { [datetime]$before } else { $null }
                    $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                    $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                    Set-ADUser -Credential $script:Cred -Identity $foundUser.ObjectGUID `
                        -AccountExpirationDate $ExpiryDate -ErrorAction Stop
                    Write-Host ("    Scadenza impostata: {0} -> {1}" -f $beforeTxt, $afterTxt) -ForegroundColor Green

                    # Riepilogo finale per l'operatore
                    $groupsAssignedTxt = if ($groupsAssigned.Count -gt 0) { $groupsAssigned -join "; " } else { "(nessuno)" }
                    $groupsFailedTxt   = if ($groupsFailed.Count -gt 0)   { $groupsFailed -join "; " }   else { "" }
                    $msgLog = "Family Web: gruppi assegnati=[$groupsAssignedTxt]"
                    if ($groupsFailedTxt) { $msgLog += " | ATTENZIONE gruppi falliti=[$groupsFailedTxt]" }

                    Write-Host ("  {0}: Family Web completato. {1}" -f $foundUser.SamAccountName, $msgLog) -ForegroundColor Green

                    Add-Log `
                        -InputText  $entry `
                        -Sam        $foundUser.SamAccountName `
                        -UPN        $foundUser.UserPrincipalName `
                        -DN         $familyWebOU `
                        -IsEnabled  "Yes (riattivato)" `
                        -Groups     $groupsAssignedTxt `
                        -OU_Flag    ("Moved_to: {0}" -f $familyWebOUDisplay) `
                        -PrevExpiry $beforeTxt `
                        -NewExpiry  $afterTxt `
                        -Status     "FamilyWeb_Completed" `
                        -Message    $msgLog
                }
                catch {
                    $errMsg = $_ | Out-String
                    Write-Host ("  Errore nel workflow Family Web per '{0}': {1}" -f $entry, $errMsg) -ForegroundColor Red
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $dn -IsEnabled $enabledTxt -Groups "(solo Domain Users)" -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Error_FamilyWeb" -Message $errMsg
                }
                continue
            }

            # ==================================================
            # CASO B: disabilitato + OU disabilitati + ha gruppi
            #         -> workflow riattivazione
            # ==================================================
            if (-not $isEnabled -and $inDisabledOU) {
                Write-Host ""
                $confirm = Read-Host "  Vuoi RIATTIVARE questo utente? (S/N)"

                if ($confirm -notmatch '^[Ss]$') {
                    Write-Host ("  Riattivazione annullata per {0}." -f $foundUser.SamAccountName) -ForegroundColor Gray
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $dn -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Skipped_ReactivationRefused" `
                        -Message "Operatore ha rifiutato la riattivazione"
                    continue
                }

                # Selezione OU destinazione
                Write-Host ""
                Write-Host "  Seleziona la OU di destinazione:" -ForegroundColor Cyan
                foreach ($ou in $targetOUs) { Write-Host ("  [{0}] {1}" -f $ou.Label, $ou.Display) -ForegroundColor Cyan }
                Write-Host ""

                $ouChoice = $null
                do {
                    $ouInput  = Read-Host "  Inserisci il numero (1-6)"
                    $ouChoice = $targetOUs | Where-Object { $_.Label -eq $ouInput }
                    if (-not $ouChoice) { Write-Host "  Scelta non valida. Riprova." -ForegroundColor Red }
                } while (-not $ouChoice)

                Write-Host ("  OU selezionata: {0}" -f $ouChoice.Display) -ForegroundColor Green

                try {
                    Move-ADObject  -Credential $script:Cred -Identity $foundUser.ObjectGUID -TargetPath $ouChoice.DN -ErrorAction Stop
                    Write-Host ("  Utente spostato in: {0}" -f $ouChoice.Display) -ForegroundColor Green

                    # Usa ObjectGUID (stabile dopo Move-ADObject, univoco in multi-dominio)
                    Enable-ADAccount -Credential $script:Cred -Identity $foundUser.ObjectGUID -ErrorAction Stop
                    Write-Host ("  Account abilitato: {0}" -f $foundUser.SamAccountName) -ForegroundColor Green

                    $before    = $foundUser.accountExpirationDate
                    $beforeDT  = if ($before) { [datetime]$before } else { $null }
                    $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                    $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                    Set-ADUser -Credential $script:Cred -Identity $foundUser.ObjectGUID -AccountExpirationDate $ExpiryDate -ErrorAction Stop
                    Write-Host ("{0}: riattivato + spostato + prorogato. {1} -> {2}" -f $foundUser.SamAccountName, $beforeTxt, $afterTxt) -ForegroundColor Green

                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $ouChoice.DN -IsEnabled "Yes (riattivato)" -Groups $groupsTxt `
                        -OU_Flag ("Moved_to: {0}" -f $ouChoice.Display) `
                        -PrevExpiry $beforeTxt -NewExpiry $afterTxt `
                        -Status "Reactivated_Moved_Updated" `
                        -Message ("Riattivato, spostato in {0}, prorogato" -f $ouChoice.Display)
                }
                catch {
                    $errMsg = $_ | Out-String
                    Write-Host ("  Errore riattivazione '{0}': {1}" -f $entry, $errMsg) -ForegroundColor Red
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $dn -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Error_Reactivation" -Message $errMsg
                }
                continue
            }

            # ==================================================
            # CASO C: attivo + OU normale -> proroga diretta
            # ==================================================
            if ($isEnabled -and (-not $inDisabledOU)) {
                # FIX v4: $statusResult e $beforeDT inizializzati prima del try
                # per evitare che il catch o il contatore leggano valori residui
                $statusResult = $null
                $beforeDT     = $null
                try {
                    $before    = $foundUser.accountExpirationDate
                    $beforeDT  = if ($before) { [datetime]$before } else { $null }
                    $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                    $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                    Set-ADUser -Credential $script:Cred -Identity $foundUser.DistinguishedName -AccountExpirationDate $ExpiryDate -ErrorAction Stop

                    $statusResult  = if ($beforeDT -and $beforeDT -eq $ExpiryDate) { "NoChange" } else { "Updated" }
                    $messageResult = if ($statusResult -eq "NoChange") { "Expiration already set to requested value" } else { "OK" }

                    Write-Host ("{0}: {1} -> {2}" -f $foundUser.SamAccountName, $beforeTxt, $afterTxt) -ForegroundColor Green

                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry $beforeTxt -NewExpiry $afterTxt -Status $statusResult -Message $messageResult
                }
                catch {
                    $msg           = $_ | Out-String
                    $beforeTxtSafe = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                    Write-Host ("Errore aggiornando '{0}': {1}" -f $entry, $msg) -ForegroundColor Yellow
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry $beforeTxtSafe -NewExpiry $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss') `
                        -Status "Error" -Message $msg
                }
                continue
            }

            # ==================================================
            # CASO E: attivo + ha gruppi + OU disabilitati
            #         -> segnala anomalia e propone spostamento in OU
            # ==================================================
            if ($isEnabled -and $inDisabledOU) {
                Write-Host "  [ANOMALIA] Utente ATTIVO ma si trova in una OU di utenti disabilitati." -ForegroundColor Yellow
                Write-Host ("  OU attuale: {0}" -f $ouPath) -ForegroundColor Yellow
                Write-Host ""
                $confirmMove = Read-Host "  Vuoi SPOSTARLO in una OU attiva e prorogare? (S/N)"

                if ($confirmMove -notmatch '^[Ss]$') {
                    Write-Host ("  Spostamento annullato per {0}. Nessuna modifica." -f $foundUser.SamAccountName) -ForegroundColor Gray
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $dn -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Skipped_ByOperator" `
                        -Message "Attivo in OU disabilitati: operatore ha rifiutato lo spostamento"
                    continue
                }

                Write-Host ""
                Write-Host "  Seleziona la OU di destinazione:" -ForegroundColor Cyan
                foreach ($ou in $targetOUs) { Write-Host ("  [{0}] {1}" -f $ou.Label, $ou.Display) -ForegroundColor Cyan }
                Write-Host ""

                $ouChoice = $null
                do {
                    $ouInput  = Read-Host "  Inserisci il numero (1-6)"
                    $ouChoice = $targetOUs | Where-Object { $_.Label -eq $ouInput }
                    if (-not $ouChoice) { Write-Host "  Scelta non valida. Riprova." -ForegroundColor Red }
                } while (-not $ouChoice)

                Write-Host ("  OU selezionata: {0}" -f $ouChoice.Display) -ForegroundColor Green

                try {
                    Move-ADObject -Credential $script:Cred -Identity $foundUser.ObjectGUID -TargetPath $ouChoice.DN -ErrorAction Stop
                    Write-Host ("  Utente spostato in: {0}" -f $ouChoice.Display) -ForegroundColor Green

                    $before    = $foundUser.accountExpirationDate
                    $beforeDT  = if ($before) { [datetime]$before } else { $null }
                    $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                    $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                    Set-ADUser -Credential $script:Cred -Identity $foundUser.ObjectGUID -AccountExpirationDate $ExpiryDate -ErrorAction Stop
                    Write-Host ("{0}: spostato + prorogato. {1} -> {2}" -f $foundUser.SamAccountName, $beforeTxt, $afterTxt) -ForegroundColor Green

                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $ouChoice.DN -IsEnabled $enabledTxt -Groups $groupsTxt `
                        -OU_Flag ("Moved_to: {0}" -f $ouChoice.Display) `
                        -PrevExpiry $beforeTxt -NewExpiry $afterTxt `
                        -Status "Moved_Updated" `
                        -Message ("Attivo in OU disabilitati, spostato in {0}, prorogato" -f $ouChoice.Display)
                }
                catch {
                    $errMsg = $_ | Out-String
                    Write-Host ("  Errore spostamento '{0}': {1}" -f $entry, $errMsg) -ForegroundColor Red
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $dn -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Error_Reactivation" -Message $errMsg
                }
                continue
            }

            # ==================================================
            # CASO D: disabilitato + OU normale -> menu 1/2/3
            # ==================================================
            if (-not $isEnabled -and (-not $inDisabledOU)) {
                Write-Host ""
                Write-Host "  Cosa vuoi fare?" -ForegroundColor Cyan
                Write-Host "  [1] Riattiva, sposta in OU scelta e proroga di 6 mesi" -ForegroundColor Cyan
                Write-Host "  [2] Solo proroga di 6 mesi (rimane disabilitato)" -ForegroundColor Cyan
                Write-Host "  [3] Salta (non fare nulla)" -ForegroundColor Cyan
                Write-Host ""

                $actionChoice = $null
                do {
                    $actionInput = Read-Host "  Scelta (1/2/3)"
                    if ($actionInput -in @('1','2','3')) { $actionChoice = $actionInput }
                    else { Write-Host "  Scelta non valida. Inserisci 1, 2 o 3." -ForegroundColor Red }
                } while (-not $actionChoice)

                # --- Scelta 3: salta ---
                if ($actionChoice -eq '3') {
                    Write-Host ("  Operazione saltata per {0}." -f $foundUser.SamAccountName) -ForegroundColor Gray
                    Add-Log `
                        -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                        -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                        -PrevExpiry "" -NewExpiry "" -Status "Skipped_ByOperator" `
                        -Message "Operatore ha scelto di non fare nulla (disabilitato, OU normale)"
                    continue
                }

                # --- Scelta 1: riattiva + sposta + proroga ---
                if ($actionChoice -eq '1') {
                    Write-Host ""
                    Write-Host "  Seleziona la OU di destinazione:" -ForegroundColor Cyan
                    foreach ($ou in $targetOUs) { Write-Host ("  [{0}] {1}" -f $ou.Label, $ou.Display) -ForegroundColor Cyan }
                    Write-Host ""

                    $ouChoice = $null
                    do {
                        $ouInput  = Read-Host "  Inserisci il numero (1-6)"
                        $ouChoice = $targetOUs | Where-Object { $_.Label -eq $ouInput }
                        if (-not $ouChoice) { Write-Host "  Scelta non valida. Riprova." -ForegroundColor Red }
                    } while (-not $ouChoice)

                    Write-Host ("  OU selezionata: {0}" -f $ouChoice.Display) -ForegroundColor Green

                    try {
                        Move-ADObject -Credential $script:Cred -Identity $foundUser.ObjectGUID -TargetPath $ouChoice.DN -ErrorAction Stop
                        Write-Host ("  Utente spostato in: {0}" -f $ouChoice.Display) -ForegroundColor Green

                        # Usa ObjectGUID (stabile dopo Move-ADObject)
                        Enable-ADAccount -Credential $script:Cred -Identity $foundUser.ObjectGUID -ErrorAction Stop
                        Write-Host ("  Account abilitato: {0}" -f $foundUser.SamAccountName) -ForegroundColor Green

                        $before    = $foundUser.accountExpirationDate
                        $beforeDT  = if ($before) { [datetime]$before } else { $null }
                        $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                        $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                        Set-ADUser -Credential $script:Cred -Identity $foundUser.ObjectGUID -AccountExpirationDate $ExpiryDate -ErrorAction Stop
                        Write-Host ("{0}: riattivato + spostato + prorogato. {1} -> {2}" -f $foundUser.SamAccountName, $beforeTxt, $afterTxt) -ForegroundColor Green

                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $ouChoice.DN -IsEnabled "Yes (riattivato)" -Groups $groupsTxt `
                            -OU_Flag ("Moved_to: {0}" -f $ouChoice.Display) `
                            -PrevExpiry $beforeTxt -NewExpiry $afterTxt `
                            -Status "Reactivated_Moved_Updated" `
                            -Message ("Riattivato da OU normale, spostato in {0}, prorogato" -f $ouChoice.Display)
                    }
                    catch {
                        $errMsg = $_ | Out-String
                        Write-Host ("  Errore riattivazione '{0}': {1}" -f $entry, $errMsg) -ForegroundColor Red
                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                            -PrevExpiry "" -NewExpiry "" -Status "Error_Reactivation" -Message $errMsg
                    }
                    continue
                }

                # --- Scelta 2: solo proroga (rimane disabilitato) ---
                if ($actionChoice -eq '2') {
                    # FIX v4: $statusResult e $beforeDT inizializzati prima del try
                    $statusResult = $null
                    $beforeDT     = $null
                    try {
                        $before    = $foundUser.accountExpirationDate
                        $beforeDT  = if ($before) { [datetime]$before } else { $null }
                        $beforeTxt = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                        $afterTxt  = $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss')

                        Set-ADUser -Credential $script:Cred -Identity $foundUser.DistinguishedName -AccountExpirationDate $ExpiryDate -ErrorAction Stop

                        $statusResult  = if ($beforeDT -and $beforeDT -eq $ExpiryDate) { "NoChange" } else { "Updated" }
                        $messageResult = if ($statusResult -eq "NoChange") {
                            "Expiration already set to requested value (account rimane disabilitato)"
                        } else {
                            "Prorogato (account lasciato disabilitato per scelta operatore)"
                        }

                        Write-Host ("{0}: {1} -> {2} (account rimane DISABILITATO)" -f $foundUser.SamAccountName, $beforeTxt, $afterTxt) -ForegroundColor Green

                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                            -PrevExpiry $beforeTxt -NewExpiry $afterTxt -Status $statusResult -Message $messageResult
                    }
                    catch {
                        $msg           = $_ | Out-String
                        $beforeTxtSafe = if ($beforeDT) { $beforeDT.ToString('dd/MM/yyyy HH:mm:ss') } else { 'Nessuna (senza scadenza)' }
                        Write-Host ("Errore aggiornando '{0}': {1}" -f $entry, $msg) -ForegroundColor Yellow
                        Add-Log `
                            -InputText $entry -Sam $foundUser.SamAccountName -UPN $foundUser.UserPrincipalName `
                            -DN $foundUser.DistinguishedName -IsEnabled $enabledTxt -Groups $groupsTxt -OU_Flag $ouFlag `
                            -PrevExpiry $beforeTxtSafe -NewExpiry $ExpiryDate.ToString('dd/MM/yyyy HH:mm:ss') `
                            -Status "Error" -Message $msg
                    }
                    continue
                }
            }

        } # fine while ($wi -lt $workList.Count)

        Write-Host ""
        Write-Host "--- FINE ELABORAZIONE BATCH ---" -ForegroundColor Cyan

        # --- VERIFICA REALE su AD: rilegge ogni utente dichiarato OK in questo batch ---
        # Controlla scadenza, stato abilitato e OU (dove il flusso li ha modificati)
        $verifiedOk     = 0
        $verifyFailed   = 0
        $successStatuses = @('Updated','NoChange','Reactivated_Moved_Updated','Moved_Updated','FamilyWeb_Completed','DueDiligence_Completed','Profiled_Completed')
        Write-Host ""
        Write-Host "--- VERIFICA SU ACTIVE DIRECTORY ---" -ForegroundColor Cyan
        for ($li = $batchLogStart; $li -lt $script:logs.Count; $li++) {
            $row = $script:logs[$li]
            if ($row.Status -notin $successStatuses) { continue }
            $problems = @()
            try {
                $chk = Get-ADUser -Credential $script:Cred -Filter ("SamAccountName -eq '{0}'" -f (Escape-ADFilterValue $row.SamAccountName)) `
                    -Properties accountExpirationDate, Enabled -ErrorAction Stop
                if (-not $chk) {
                    $problems += "utente non trovato in rilettura"
                } else {
                    $chk = @($chk)[0]
                    $expOk = $false
                    if ($chk.accountExpirationDate) {
                        $expOk = ([math]::Abs((([datetime]$chk.accountExpirationDate) - $ExpiryDate).TotalHours) -le 24)
                    }
                    if (-not $expOk) { $problems += ("scadenza in AD = {0}" -f $(if ($chk.accountExpirationDate) { ([datetime]$chk.accountExpirationDate).ToString('dd/MM/yyyy HH:mm:ss') } else { 'nessuna' })) }
                    if ($row.Status -ne 'Updated' -and $row.Status -ne 'NoChange') {
                        if (-not $chk.Enabled) { $problems += "account ancora DISABILITATO" }
                        if ($row.DistinguishedName -and $chk.DistinguishedName -notlike ("*," + $row.DistinguishedName)) {
                            $problems += ("OU in AD = {0}" -f ($chk.DistinguishedName -replace '^CN=[^,]+,', ''))
                        }
                    }
                }
            }
            catch {
                $problems += ("errore rilettura: {0}" -f $_.Exception.Message)
            }

            if ($problems.Count -eq 0) {
                $verifiedOk++
                Write-Host ("  [VERIFICATO]   {0}" -f $row.SamAccountName) -ForegroundColor Green
            } else {
                $verifyFailed++
                $script:logs[$li].Message = ("{0} | VERIFICA FALLITA: {1}" -f $row.Message, ($problems -join '; '))
                $script:logs[$li].Status  = "VERIFY_FAILED"
                Write-Host ("  [NON APPLICATO] {0}: {1}" -f $row.SamAccountName, ($problems -join '; ')) -ForegroundColor Red
            }
        }
        if (($verifiedOk + $verifyFailed) -eq 0) {
            Write-Host "  Nessun utente modificato in questo batch." -ForegroundColor Yellow
        }

        # --- Copia negli appunti il testo per il ticket SOLO se almeno un utente e' stato verificato ---
        if ($verifyFailed -gt 0) {
            Write-Host ""
            Write-Host ("ATTENZIONE: {0} utenti NON risultano modificati in AD. Testo ticket NON copiato." -f $verifyFailed) -ForegroundColor Red
        }
        elseif ($verifiedOk -gt 0) {
            $clipText = "utenza prorogata al {0}" -f $ExpiryDate.ToString('dd/MM/yyyy')
            try {
                $clipText | Set-Clipboard
                Write-Host ("Copiato negli appunti: ""{0}""" -f $clipText) -ForegroundColor Cyan
                Write-Host "(incolla direttamente nel ticket con CTRL+V)" -ForegroundColor DarkGray
            }
            catch {
                Write-Host ("Testo per il ticket: ""{0}""" -f $clipText) -ForegroundColor Cyan
                Write-Host "(Set-Clipboard non disponibile, copia manualmente il testo sopra)" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Nessuna proroga eseguita: testo ticket NON copiato." -ForegroundColor Yellow
        }

    } # fine else (hasInput)

    # --------------------------------------------------
    # PROMPT DI FINE CICLO: proroga altri utenti o esci?
    # --------------------------------------------------
    Write-Host ""
    # INVIO senza digitare nulla equivale a "S" (proroga altri utenti)
    $again = Read-Host "Vuoi prorogare altri utenti? [INVIO = Si, N = No]"
    if ($again -match '^[Nn]$') {
        $continueLoop = $false
    }

} # fine while ($continueLoop)

# ========================
# 5. RIEPILOGO E EXPORT LOG
# ========================

$updated     = ($script:logs | Where-Object { $_.Status -eq 'Updated' }).Count
$nochange    = ($script:logs | Where-Object { $_.Status -eq 'NoChange' }).Count
$errors      = ($script:logs | Where-Object { $_.Status -in @('Error','Error_Reactivation','Error_FamilyWeb','Error_DueDiligence','Error_Profilo','VERIFY_FAILED') }).Count
$notFound    = ($script:logs | Where-Object { $_.Status -eq 'NotFound' }).Count
$ambig       = ($script:logs | Where-Object { $_.Status -eq 'Ambiguous' }).Count
$blocked     = ($script:logs | Where-Object { $_.Status -eq 'Blocked_NoGroups' }).Count
$reactivated = ($script:logs | Where-Object { $_.Status -in @('Reactivated_Moved_Updated','Moved_Updated') }).Count
$familyWeb   = ($script:logs | Where-Object { $_.Status -eq 'FamilyWeb_Completed' }).Count
$dueDiligenceCompleted = ($script:logs | Where-Object { $_.Status -eq 'DueDiligence_Completed' }).Count
$profiled    = ($script:logs | Where-Object { $_.Status -eq 'Profiled_Completed' }).Count
$skipped     = ($script:logs | Where-Object { $_.Status -in @('Skipped_ReactivationRefused','Skipped_ByOperator') }).Count

# Riga di riepilogo nel log
$script:logs += [pscustomobject]@{
    Timestamp              = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Input                  = '--- SUMMARY ---'
    SamAccountName         = ''
    UserPrincipalName      = ''
    DistinguishedName      = ''
    IsEnabled              = ''
    Groups                 = ''
    OU_Flag                = ''
    PreviousExpirationDate = ''
    NewExpirationDate      = ''
    Status                 = 'Summary'
    Message                = "Updated=$updated; NoChange=$nochange; Errors=$errors; NotFound=$notFound; Ambiguous=$ambig; Blocked_NoGroups=$blocked; Reactivated=$reactivated; FamilyWeb=$familyWeb; DueDiligence=$dueDiligenceCompleted; Profiled=$profiled; Skipped=$skipped"
}

# --- Export Excel o CSV ---
if (-not $script:UseCSVFallback) {
    try {
        $conditionalFormats = @(
            New-ConditionalText -Text "Error"            -BackgroundColor LightSalmon -ConditionalTextColor DarkRed
            New-ConditionalText -Text "NotFound"         -BackgroundColor LightSalmon -ConditionalTextColor DarkRed
            New-ConditionalText -Text "Blocked_NoGroups" -BackgroundColor LightYellow -ConditionalTextColor DarkOrange
            New-ConditionalText -Text "Updated"          -BackgroundColor LightGreen  -ConditionalTextColor DarkGreen
            New-ConditionalText -Text "Reactivated"      -BackgroundColor PaleGreen   -ConditionalTextColor DarkGreen
            New-ConditionalText -Text "FamilyWeb"        -BackgroundColor PaleGreen   -ConditionalTextColor DarkGreen
            New-ConditionalText -Text "Profiled"         -BackgroundColor PaleGreen   -ConditionalTextColor DarkGreen
            New-ConditionalText -Text "Ambiguous"        -BackgroundColor LightYellow -ConditionalTextColor DarkOrange
        )

        $script:logs | Export-Excel `
            -Path $xlsxPath `
            -WorksheetName "AD_Expiry_Log" `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
            -ConditionalText $conditionalFormats `
            -ErrorAction Stop

        Write-Host ""
        Write-Host ("Log Excel salvato:{0}" -f $xlsxPath) -ForegroundColor Cyan
    }
    catch {
        Write-Host ("Errore export Excel: {0}. Uso CSV di fallback." -f ($_ | Out-String)) -ForegroundColor Yellow
        $script:UseCSVFallback = $true
    }
}

if ($script:UseCSVFallback) {
    try {
        $script:logs | Export-Csv -Path $csvPath -NoTypeInformation -UseCulture -Encoding UTF8
        Write-Host ""
        Write-Host ("Log CSV salvato:{0}" -f $csvPath) -ForegroundColor Cyan
    }
    catch {
        Write-Host ("Impossibile scrivere il log. Errore: {0}" -f ($_ | Out-String)) -ForegroundColor Yellow
    }
}

# ========================
# 6. RIEPILOGO FINALE + CLIPBOARD
# ========================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "--- OPERAZIONE COMPLETATA ---" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host ("  Updated      : {0}" -f $updated)     -ForegroundColor Green
Write-Host ("  NoChange     : {0}" -f $nochange)    -ForegroundColor DarkGray
Write-Host ("  FamilyWeb    : {0}" -f $familyWeb)   -ForegroundColor Green
Write-Host ("  DueDiligence : {0}" -f $dueDiligenceCompleted) -ForegroundColor Green
Write-Host ("  Profilati    : {0}" -f $profiled)    -ForegroundColor Green
Write-Host ("  Reactivated  : {0}" -f $reactivated) -ForegroundColor Green
Write-Host ("  Blocked      : {0}" -f $blocked)     -ForegroundColor Yellow
Write-Host ("  NotFound     : {0}" -f $notFound)    -ForegroundColor Red
Write-Host ("  Errors       : {0}" -f $errors)      -ForegroundColor Red
Write-Host ("  Ambiguous    : {0}" -f $ambig)       -ForegroundColor Yellow
Write-Host ("  Skipped      : {0}" -f $skipped)     -ForegroundColor DarkGray
Write-Host ""

Write-Host ""
# Read-Host finale: tiene la finestra aperta finche' l'operatore non preme INVIO
# (necessario perche' lo script gira con -NoExit ma i contatori scorrerebbero
# fuori schermo non appena PowerShell stampa il prompt)
Read-Host "Premi INVIO per chiudere"
<#
.SYNOPSIS
    Déploie un Web Add-in Outlook sur Exchange 2019 On-Premise pour un groupe AD.
.NOTES
    Exécuter depuis l'Exchange Management Shell.
    Nécessite le module ActiveDirectory (RSAT) pour la résolution du groupe.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ADGroupName,

    [switch]$Force
)

$ErrorActionPreference = "Stop"

# ============================================================================
# MANIFEST XML — Remplace le contenu ci-dessous par ton manifest complet
# ============================================================================
$ManifestXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="MailApp">
  <Id>xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Mon Entreprise</ProviderName>
  <DefaultLocale>fr-FR</DefaultLocale>
  <DisplayName DefaultValue="Mon Plugin Outlook"/>
  <Description DefaultValue="Description du plugin"/>
  <Hosts><Host Name="Mailbox"/></Hosts>
  <Requirements><Sets><Set Name="Mailbox" MinVersion="1.1"/></Sets></Requirements>
  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://monserveur.local/addin/index.html"/>
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
  </Rule>
</OfficeApp>
"@

# ============================================================================
# FONCTIONS
# ============================================================================

function Get-ManifestId([string]$Xml) {
    ([xml]$Xml).OfficeApp.Id.Trim()
}

function Get-ManifestVersion([string]$Xml) {
    ([xml]$Xml).OfficeApp.Version.Trim()
}

function Resolve-ADGroupMailboxes([string]$GroupName) {
    Import-Module ActiveDirectory -ErrorAction Stop

    $members = Get-ADGroupMember -Identity $GroupName -Recursive |
               Where-Object { $_.objectClass -eq "user" } |
               Select-Object -ExpandProperty SamAccountName

    if (-not $members) {
        throw "Le groupe AD '$GroupName' est vide ou introuvable."
    }

    # Ne garder que les membres qui ont une boîte mail Exchange
    $mailboxes = @()
    foreach ($sam in $members) {
        $mbx = Get-Mailbox -Identity $sam -ErrorAction SilentlyContinue
        if ($mbx) { $mailboxes += $mbx.PrimarySmtpAddress.ToString() }
    }

    if (-not $mailboxes) {
        throw "Aucun membre du groupe '$GroupName' ne possède de boîte mail Exchange."
    }

    Write-Host "[OK] $($mailboxes.Count) boîte(s) mail trouvée(s) dans '$GroupName'." -ForegroundColor Green
    return $mailboxes
}

# ============================================================================
# EXÉCUTION
# ============================================================================

try {
    $appId   = Get-ManifestId $ManifestXml
    $version = Get-ManifestVersion $ManifestXml
    $bytes   = [System.Text.Encoding]::UTF8.GetBytes($ManifestXml)

    Write-Host "[*] Add-in: $appId v$version" -ForegroundColor Cyan
    Write-Host "[*] Résolution du groupe AD '$ADGroupName'..." -ForegroundColor Cyan
    $userList = Resolve-ADGroupMailboxes -GroupName $ADGroupName

    # Vérifie si l'add-in existe déjà
    $existing = Get-App -OrganizationApp -ErrorAction SilentlyContinue |
                Where-Object { $_.AppId -eq $appId }

    if ($existing) {
        if (($existing.AppVersion -eq $version) -and (-not $Force)) {
            Write-Host "[OK] Déjà installé en v$version. Rien à faire (-Force pour forcer)." -ForegroundColor Green
        }
        else {
            Write-Host "[*] Mise à jour vers v$version..." -ForegroundColor Cyan
            Set-App -Identity $existing.Identity `
                    -OrganizationApp `
                    -FileData $bytes `
                    -ProvidedTo SpecificUsers `
                    -UserList $userList `
                    -DefaultStateForUser Enabled `
                    -Enabled:$true `
                    -Confirm:$false
            Write-Host "[OK] Add-in mis à jour." -ForegroundColor Green
        }
    }
    else {
        Write-Host "[*] Installation..." -ForegroundColor Cyan
        New-App -OrganizationApp `
                -FileData $bytes `
                -ProvidedTo SpecificUsers `
                -UserList $userList `
                -DefaultStateForUser Enabled `
                -Enabled:$true `
                -Confirm:$false
        Write-Host "[OK] Add-in installé pour $($userList.Count) utilisateur(s)." -ForegroundColor Green
    }
}
catch {
    Write-Host "[ERREUR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

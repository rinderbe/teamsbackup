#Requires -Version 7.0
<#
    TeamsConfigBackup.psm1

    Backup och återläsning av Teams Phone / Direct Routing-konfiguration.

    Kräver:
      - PowerShell 7+
      - Modulen MicrosoftTeams (6.x eller senare rekommenderas)
      - Connect-MicrosoftTeams med ett konto som är Teams Administrator

    Exporterade funktioner:
      Backup-TeamsConfig   – läser ut konfigurationen till en tidsstämplad mapp (JSON + CSV)
      Restore-TeamsConfig  – läser tillbaka hela eller delar av en backup (stöder -WhatIf)
      Get-TeamsConfigBackup – visar innehållet i en backup
      Compare-TeamsConfigBackup – visar skillnaderna mellan två backuper (kräver ingen anslutning)
      Get-TeamsCallForwarding – listar användare med aktiv vidarekoppling
#>

$script:ScriptVersion   = '1.1.0'
$script:JsonDepth       = 15
$script:NoiseProperties = 'Element', 'XsAnyElements', 'ScopeClass', 'Anchor'

# Parametrar som aldrig tas från backup-objektet vid återläsning
$script:ExcludedParams = @(
    'Identity', 'Name', 'Force', 'InMemory', 'Instance', 'Tenant', 'PassThru', 'MsftInternalProcessingMode',
    'Break', 'HttpPipelineAppend', 'HttpPipelinePrepend', 'Proxy', 'ProxyCredential', 'ProxyUseDefaultCredentials',
    'WhatIf', 'Confirm'
) + [System.Management.Automation.PSCmdlet]::CommonParameters + [System.Management.Automation.PSCmdlet]::OptionalCommonParameters

# Ordningen styr återläsningen (beroenden: usages -> gateways -> routes -> policies -> användare).
# Borttagning (-RemoveExtra) sker i omvänd ordning.
$script:RestoreOrder = @(
    'PstnUsages', 'TranslationRules', 'Gateways', 'VoiceRoutes', 'VoiceRoutingPolicies', 'DialPlans',
    'Policies', 'NetworkConfig', 'NumberBlocking', 'GroupPolicyAssignments', 'PhoneNumbers', 'Users', 'UserPolicies'
)
$script:UserCategories = 'PhoneNumbers', 'Users', 'UserPolicies'

# Objekt som backas upp och återläses generiskt via Get-/New-/Set-/Remove-Cs<Noun>
$script:ObjectSets = [ordered]@{
    TranslationRules     = @(@{ Noun = 'TeamsTranslationRule';     File = 'DirectRouting/TranslationRules.json' })
    Gateways             = @(@{ Noun = 'OnlinePSTNGateway';        File = 'DirectRouting/Gateways.json' })
    VoiceRoutes          = @(@{ Noun = 'OnlineVoiceRoute';         File = 'DirectRouting/VoiceRoutes.json' })
    VoiceRoutingPolicies = @(@{ Noun = 'OnlineVoiceRoutingPolicy'; File = 'DirectRouting/VoiceRoutingPolicies.json' })
    DialPlans            = @(@{ Noun = 'TenantDialPlan';           File = 'DirectRouting/DialPlans.json' })
    NetworkConfig        = @(
        @{ Noun = 'TenantNetworkRegion';    File = 'Network/NetworkRegions.json' }
        @{ Noun = 'TenantNetworkSite';      File = 'Network/NetworkSites.json' }
        @{ Noun = 'TenantNetworkSubnet';    File = 'Network/NetworkSubnets.json' }
        @{ Noun = 'TenantTrustedIPAddress'; File = 'Network/TrustedIPAddresses.json' }
    )
    NumberBlocking       = @(
        @{ Noun = 'InboundBlockedNumberPattern'; File = 'NumberBlocking/InboundBlockedNumberPatterns.json' }
        @{ Noun = 'InboundExemptNumberPattern';  File = 'NumberBlocking/InboundExemptNumberPatterns.json' }
    )
}

# Policytyper som alltid backas upp (kompletteras med de policyer som syns på användarobjekten)
$script:KnownPolicyTypes = @(
    'CallingLineIdentity', 'OnlineVoicemailPolicy', 'TeamsCallingPolicy', 'TeamsEmergencyCallingPolicy',
    'TeamsEmergencyCallRoutingPolicy', 'TeamsCallParkPolicy', 'TeamsCallHoldPolicy', 'TeamsIPPhonePolicy',
    'TeamsVoiceApplicationsPolicy', 'TeamsSharedCallingRoutingPolicy', 'TeamsNetworkRoamingPolicy',
    'TeamsSurvivableBranchAppliancePolicy', 'OnlineAudioConferencingRoutingPolicy', 'OnlineDialOutPolicy',
    'TeamsMeetingPolicy', 'TeamsMessagingPolicy', 'TeamsAppSetupPolicy', 'TeamsAppPermissionPolicy',
    'TeamsChannelsPolicy', 'TeamsUpdateManagementPolicy', 'TeamsFeedbackPolicy', 'TeamsMobilityPolicy',
    'TeamsVdiPolicy', 'TeamsEventsPolicy', 'TeamsFilesPolicy', 'TeamsComplianceRecordingPolicy',
    'ExternalAccessPolicy', 'TeamsUpgradePolicy'
)
# Dessa har egna kategorier under DirectRouting
$script:DirectRoutingPolicyTypes = 'OnlineVoiceRoutingPolicy', 'TenantDialPlan'

# Egenskapsnamn på Get-objektet som skiljer sig från parameternamnet på New/Set
$script:PropertyAliases = @{
    EmergencyCallRoutingPolicy = 'EmergencyCallRoutingPolicyName'
    EmergencyCallingPolicy     = 'EmergencyCallingPolicyName'
    NetworkRoamingPolicy       = 'NetworkRoamingPolicyName'
}

# Konfiguration som backas upp men inte återläses automatiskt (för komplex / ID:n ändras)
$script:BackupOnlyItems = [ordered]@{
    'BackupOnly/AutoAttendants.json'                    = { Get-AllPaged 'Get-CsAutoAttendant' }
    'BackupOnly/CallQueues.json'                        = { Get-AllPaged 'Get-CsCallQueue' }
    'BackupOnly/ResourceAccounts.json'                  = { Get-CsOnlineApplicationInstance -ErrorAction Stop }
    'BackupOnly/OnlineSchedules.json'                   = { Get-CsOnlineSchedule -ErrorAction Stop }
    'BackupOnly/SurvivableBranchAppliances.json'        = { Get-CsTeamsSurvivableBranchAppliance -ErrorAction Stop }
    'BackupOnly/Emergency/LisCivicAddresses.json'       = { Get-CsOnlineLisCivicAddress -ErrorAction Stop }
    'BackupOnly/Emergency/LisLocations.json'            = { Get-CsOnlineLisLocation -ErrorAction Stop }
    'BackupOnly/Emergency/LisSubnets.json'              = { Get-CsOnlineLisSubnet -ErrorAction Stop }
    'BackupOnly/Emergency/LisSwitches.json'             = { Get-CsOnlineLisSwitch -ErrorAction Stop }
    'BackupOnly/Emergency/LisPorts.json'                = { Get-CsOnlineLisPort -ErrorAction Stop }
    'BackupOnly/Emergency/LisWirelessAccessPoints.json' = { Get-CsOnlineLisWirelessAccessPoint -ErrorAction Stop }
    'BackupOnly/Tenant/Tenant.json'                     = { Get-CsTenant -ErrorAction Stop }
    'BackupOnly/Tenant/FederationConfiguration.json'    = { Get-CsTenantFederationConfiguration -ErrorAction Stop }
    'BackupOnly/Tenant/TeamsUpgradeConfiguration.json'  = { Get-CsTeamsUpgradeConfiguration -ErrorAction Stop }
    'BackupOnly/Tenant/TeamsClientConfiguration.json'   = { Get-CsTeamsClientConfiguration -ErrorAction Stop }
    'BackupOnly/Tenant/TeamsMeetingConfiguration.json'  = { Get-CsTeamsMeetingConfiguration -ErrorAction Stop }
    'BackupOnly/Tenant/BlockedCallingNumbers.json'      = { Get-CsTenantBlockedCallingNumbers -ErrorAction Stop }
    'BackupOnly/Tenant/DialInConferencingSettings.json' = { Get-CsOnlineDialinConferencingTenantSettings -ErrorAction Stop }
}

# Jämförelse av backuper: nyckel per fil där Identity/Name inte räcker för att para ihop objekten
$script:CompareKeys = @{
    'BackupOnly/ResourceAccounts.json'                  = 'ObjectId'
    'BackupOnly/EffectivePolicies.json'                 = 'ObjectId'
    'BackupOnly/UserCallingSettings.json'               = 'ObjectId'
    'BackupOnly/OnlineSchedules.json'                   = 'Id'
    'BackupOnly/Emergency/LisCivicAddresses.json'       = 'CivicAddressId'
    'BackupOnly/Emergency/LisLocations.json'            = 'LocationId'
    'BackupOnly/Emergency/LisSubnets.json'              = 'Subnet'
    'BackupOnly/Emergency/LisSwitches.json'             = 'ChassisId'
    'BackupOnly/Emergency/LisPorts.json'                = 'ChassisId', 'PortId'
    'BackupOnly/Emergency/LisWirelessAccessPoints.json' = 'BSSID'
}
# Egenskaper som ändras av sig själva och inte säger något om konfigurationen
$script:CompareIgnoreProperties = 'WhenChanged', 'WhenCreated', 'LastSyncTimeStamp'

#region Hjälpfunktioner ---------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'CHANGE', 'WHATIF')][string]$Level = 'INFO',
        [switch]$NoHost
    )
    $line = '{0} [{1,-6}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if (-not $NoHost) {
        $color = switch ($Level) {
            'OK'     { 'Green' }
            'WARN'   { 'Yellow' }
            'ERROR'  { 'Red' }
            'CHANGE' { 'Cyan' }
            'WHATIF' { 'Magenta' }
            default  { 'Gray' }
        }
        Write-Host $line -ForegroundColor $color
    }
    if ($script:LogFile) { Add-Content -Path $script:LogFile -Value $line -Encoding utf8 }
}

function Assert-TeamsConnection {
    if (-not (Get-Module MicrosoftTeams)) {
        Import-Module MicrosoftTeams -ErrorAction Stop
    }
    $version = (Get-Module MicrosoftTeams).Version
    if ($version.Major -lt 6) {
        Write-Log "MicrosoftTeams $version är installerad – version 6 eller senare rekommenderas." -Level WARN
    }
    try {
        return Get-CsTenant -ErrorAction Stop
    }
    catch {
        throw "Ingen anslutning till Teams. Kör Connect-MicrosoftTeams först. ($($_.Exception.Message))"
    }
}

function Get-PolicyName($Value) {
    if ($null -eq $Value) { return $null }
    $name = if ($Value -is [string]) { $Value }
            elseif ($Value.PSObject.Properties['Name']) { [string]$Value.Name }
            else { [string]$Value }
    $name = $name -replace '^Tag:', ''
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    return $name
}

function Get-ObjectKey($Object) {
    $id = if ($Object.PSObject.Properties['Identity'] -and $Object.Identity) { [string]$Object.Identity }
          elseif ($Object.PSObject.Properties['Name']) { [string]$Object.Name }
    if ($id) { return $id -replace '^Tag:', '' }
}

function Get-AllPaged([string]$Command, [int]$PageSize = 100) {
    $skip = 0
    do {
        $page = @(& $Command -First $PageSize -Skip $skip -ErrorAction Stop)
        $page
        $skip += $page.Count
    } while ($page.Count -eq $PageSize)
}

function Get-AllPhoneNumberAssignments {
    $pageSize = 500
    $skip = 0
    do {
        $page = @(Get-CsPhoneNumberAssignment -Top $pageSize -Skip $skip -ErrorAction Stop)
        $page
        $skip += $page.Count
    } while ($page.Count -eq $pageSize)
}

# Normaliserar ett värde så att backup (JSON) och aktuell konfiguration (typade objekt) kan jämföras
function ConvertTo-Comparable($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [enum] -or $Value -is [ValueType]) { return $Value.ToString() }
    $json = ConvertTo-Json -InputObject $Value -Depth $script:JsonDepth -Compress -EnumsAsStrings -WarningAction SilentlyContinue
    if ($json -in '[]', '""', 'null', '{}') { return '' }
    return $json
}

function Test-ComplexValue($Value) {
    foreach ($item in @($Value)) {
        if ($item -is [System.Management.Automation.PSCustomObject]) { return $true }
    }
    return $false
}

# Värde att skicka när backupen innehöll $null men nuvarande värde är satt
function Get-EmptyValue([type]$Type) {
    if ($Type -eq [string]) { return '' }
    if ($Type.IsArray -or $Type.Name -like 'PSListModifier*' -or
        ($Type.IsGenericType -and [System.Collections.IEnumerable].IsAssignableFrom($Type))) {
        return , @()
    }
    return $null
}

function Get-RuleSignature($Rules) {
    @($Rules | Where-Object { $_ } | ForEach-Object {
        '{0}|{1}|{2}|{3}' -f $_.Name, $_.Pattern, $_.Translation, [bool]$_.IsInternalExtension
    }) -join ';'
}

function New-NormalizationRuleFromBackup($Rule) {
    $params = @{
        Parent      = 'Global'
        Name        = $Rule.Name
        Pattern     = $Rule.Pattern
        Translation = $Rule.Translation
        InMemory    = $true
    }
    if ($Rule.Description) { $params.Description = $Rule.Description }
    if ($null -ne $Rule.IsInternalExtension) { $params.IsInternalExtension = [bool]$Rule.IsInternalExtension }
    New-CsVoiceNormalizationRule @params
}

# Väljer den parameteruppsättning på New-/Set-cmdleten som använder -Identity (alt. -Name)
$script:ParamInfoCache = @{}
function Get-RestoreParamInfo([string]$CommandName) {
    if ($script:ParamInfoCache.ContainsKey($CommandName)) { return $script:ParamInfoCache[$CommandName] }
    $info = $null
    $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
    if ($cmd) {
        foreach ($idParam in 'Identity', 'Name') {
            $set = $cmd.ParameterSets | Where-Object { $_.Parameters.Name -contains $idParam } | Select-Object -First 1
            if (-not $set) { continue }
            $parameters = @{}
            foreach ($p in $set.Parameters) {
                if ($p.Name -notin $script:ExcludedParams) { $parameters[$p.Name] = $p.ParameterType }
            }
            $info = [pscustomobject]@{ IdParam = $idParam; Parameters = $parameters }
            break
        }
    }
    $script:ParamInfoCache[$CommandName] = $info
    return $info
}

function Format-Value($Value) {
    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { return '$' + $Value.ToString().ToLower() }
    if ($Value -is [string]) { return "'$Value'" }
    if ($Value -is [hashtable]) {
        return '@{' + (($Value.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, (Format-Value $_.Value) }) -join '; ') + '}'
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value)
        if ($items.Count -eq 0) { return '@()' }
        return (@($items | ForEach-Object { Format-Value $_ }) -join ',')
    }
    if ($Value -is [ValueType] -and $Value -isnot [enum] -and $Value -isnot [guid] -and $Value -isnot [datetime]) { return [string]$Value }
    if ($Value -isnot [ValueType] -and $Value.PSObject.Properties['Name']) { return "<$($Value.Name)>" }
    return "'$Value'"
}

function Format-Splat([hashtable]$Parameters) {
    $ordered = $Parameters.GetEnumerator() | Sort-Object @{ Expression = { $_.Key -notin 'Identity', 'GroupId' } }, Key
    ($ordered | ForEach-Object { '-{0} {1}' -f $_.Key, (Format-Value $_.Value) }) -join ' '
}

# Utför (eller simulerar vid -WhatIf) en ändring och loggar exakt kommando
function Invoke-Change {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Command,
        [Parameter(Mandatory)][hashtable]$Parameters,
        [Parameter(Mandatory)][string]$Reason
    )
    $commandText = '{0} {1}' -f $Command, (Format-Splat $Parameters)
    $record = [pscustomobject]@{
        Category = $script:CurrentCategory
        Change   = $Reason
        Command  = $commandText
        Status   = $null
        Error    = $null
    }
    if ($null -ne $script:Changes) { $script:Changes.Add($record) }

    $description = '{0}{1}    {2}' -f $Reason, [Environment]::NewLine, $commandText
    if ($PSCmdlet.ShouldProcess($description, "Köra: $commandText ?", 'Bekräfta ändring')) {
        $splat = $Parameters.Clone()
        $splat.ErrorAction = 'Stop'
        $cmdInfo = Get-Command $Command -ErrorAction SilentlyContinue
        if ($cmdInfo -and $cmdInfo.Parameters.ContainsKey('Confirm')) { $splat.Confirm = $false }
        try {
            $null = & $Command @splat
            $script:Stats.Changed++
            $record.Status = 'Utförd'
            Write-Log ('{0}: {1}' -f $Reason, $commandText) -Level CHANGE
        }
        catch {
            $script:Stats.Failed++
            $record.Status = 'Misslyckad'
            $record.Error  = $_.Exception.Message
            Write-Log ('MISSLYCKADES – {0}: {1}{2}    {3}' -f $Reason, $commandText, [Environment]::NewLine, $_.Exception.Message) -Level ERROR
        }
    }
    elseif ($WhatIfPreference) {
        $script:Stats.Planned++
        $record.Status = 'Planerad'
        Write-Log ('{0}: {1}' -f $Reason, $commandText) -Level WHATIF -NoHost
    }
    else {
        $record.Status = 'Hoppades över'
    }
}

function Save-Json([string]$RelativePath, [object[]]$Data) {
    $file = Join-Path $script:BackupRoot $RelativePath
    $null = New-Item -ItemType Directory -Force -Path (Split-Path $file -Parent)
    $clean = @($Data | Where-Object { $null -ne $_ } | Select-Object -Property * -ExcludeProperty $script:NoiseProperties)
    ConvertTo-Json -InputObject $clean -Depth $script:JsonDepth -EnumsAsStrings -WarningAction SilentlyContinue |
        Set-Content -Path $file -Encoding utf8
    $script:Manifest.Items[$RelativePath] = $clean.Count
    Write-Log ('{0,-55} {1,6} objekt' -f $RelativePath, $clean.Count) -Level OK
}

function Export-Item([string]$RelativePath, [scriptblock]$ScriptBlock) {
    try {
        Save-Json $RelativePath @(& $ScriptBlock)
    }
    catch {
        Write-Log "Kunde inte läsa ut $RelativePath – $($_.Exception.Message)" -Level WARN
        $script:Manifest.Errors += [pscustomobject]@{ Item = $RelativePath; Error = $_.Exception.Message }
    }
}

function Read-BackupFile([string]$Root, [string]$RelativePath, [switch]$Quiet) {
    $file = Join-Path $Root $RelativePath
    if (-not (Test-Path $file)) {
        if (-not $Quiet) { Write-Log "Filen $RelativePath saknas i backupen – hoppar över" -Level WARN }
        return $null
    }
    $data = @(Get-Content -Path $file -Raw -Encoding utf8 | ConvertFrom-Json)
    return , $data
}

function ConvertTo-UserRecord($User, [string[]]$PolicyProperties) {
    $policies = [ordered]@{}
    foreach ($p in $PolicyProperties) { $policies[$p] = Get-PolicyName $User.$p }
    [pscustomobject][ordered]@{
        ObjectId                  = [string]$User.Identity
        UserPrincipalName         = $User.UserPrincipalName
        DisplayName               = $User.DisplayName
        SipAddress                = $User.SipAddress
        LineUri                   = $User.LineUri
        EnterpriseVoiceEnabled    = $User.EnterpriseVoiceEnabled
        HostedVoiceMail           = $User.HostedVoiceMail
        AccountEnabled            = $User.AccountEnabled
        UsageLocation             = $User.UsageLocation
        InterpretedUserType       = $User.InterpretedUserType
        TeamsUpgradeEffectiveMode = [string]$User.TeamsUpgradeEffectiveMode
        FeatureTypes              = @($User.FeatureTypes)
        Policies                  = [pscustomobject]$policies
    }
}

function Get-TargetName([string]$ObjectId) {
    if ($script:TargetNames -and $script:TargetNames.ContainsKey($ObjectId)) { return $script:TargetNames[$ObjectId] }
    return $ObjectId
}

function Get-BackupIdentities([string]$Root) {
    foreach ($u in (Read-BackupFile $Root 'Users/Users.json' -Quiet)) { [pscustomobject]@{ ObjectId = [string]$u.ObjectId; Upn = $u.UserPrincipalName } }
    foreach ($r in (Read-BackupFile $Root 'BackupOnly/ResourceAccounts.json' -Quiet)) { [pscustomobject]@{ ObjectId = [string]$r.ObjectId; Upn = $r.UserPrincipalName } }
}

# Översätter -User (UPN eller ObjectId) till ett uppslag på ObjectId
function Resolve-TargetFilter([object[]]$Identities, [string[]]$User) {
    $filter = @{}
    foreach ($entry in $User) {
        $match = @($Identities | Where-Object { $_.Upn -eq $entry -or $_.ObjectId -eq $entry })
        if ($match) { foreach ($m in $match) { $filter[$m.ObjectId] = $true } }
        elseif ($entry -as [guid]) { $filter[$entry] = $true }
        else { Write-Log "Hittar inte '$entry' i backupen" -Level WARN }
    }
    if ($filter.Count -eq 0) { throw 'Ingen av de angivna användarna finns i backupen.' }
    return $filter
}

function Format-PolicySource($Sources) {
    @($Sources | ForEach-Object {
        if ($_.AssignmentType -eq 'Group') { 'Group:{0} ({1}, prio {2})' -f $_.Reference, $_.PolicyName, $_.Priority }
        else { '{0} ({1})' -f $_.AssignmentType, $_.PolicyName }
    }) -join '; '
}

function Select-FilteredUsers([object[]]$Users) {
    if ($script:TargetFilter) {
        return @($Users | Where-Object { $script:TargetFilter.ContainsKey([string]$_.ObjectId) })
    }
    return @($Users)
}

function Get-CurrentUserMap([string[]]$ObjectIds) {
    if ($null -eq $script:CurrentUserMap) {
        $script:CurrentUserMap = @{}
        if ($ObjectIds.Count -le 50) {
            foreach ($id in $ObjectIds) {
                try {
                    $u = Get-CsOnlineUser -Identity $id -ErrorAction Stop
                    $script:CurrentUserMap[[string]$u.Identity] = $u
                }
                catch { }
            }
        }
        else {
            Write-Log 'Läser aktuella användare från Teams...'
            foreach ($u in Get-CsOnlineUser -ErrorAction Stop) { $script:CurrentUserMap[[string]$u.Identity] = $u }
        }
    }
    return $script:CurrentUserMap
}

#endregion

#region Synkronisering per kategori -------------------------------------------------------------------

function Sync-CsObjectSet {
    param(
        [Parameter(Mandatory)][string]$Noun,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Desired,
        [Parameter(Mandatory)][ValidateSet('Upsert', 'Remove')][string]$Phase
    )
    $getCmd = "Get-Cs$Noun"
    if (-not (Get-Command $getCmd -ErrorAction SilentlyContinue)) {
        Write-Log "$getCmd finns inte i installerad MicrosoftTeams-modul – hoppar över $Noun" -Level WARN
        return
    }
    $current = @{}
    foreach ($obj in @(& $getCmd -ErrorAction Stop)) { $current[(Get-ObjectKey $obj)] = $obj }
    $desiredKeys = @{}
    foreach ($obj in $Desired) { $desiredKeys[(Get-ObjectKey $obj)] = $true }

    if ($Phase -eq 'Remove') {
        foreach ($key in @($current.Keys)) {
            if ($key -eq 'Global' -or $desiredKeys.ContainsKey($key)) { continue }
            Invoke-Change -Command "Remove-Cs$Noun" -Parameters @{ Identity = $key } -Reason "Ta bort $Noun (finns inte i backup)"
        }
        return
    }

    foreach ($obj in $Desired) {
        $key = Get-ObjectKey $obj
        if (-not $key) { continue }
        $exists = $current.ContainsKey($key)
        if (-not $exists -and $key -eq 'Global') { continue }

        $command = if ($exists) { "Set-Cs$Noun" } else { "New-Cs$Noun" }
        $info = Get-RestoreParamInfo $command
        if (-not $info) {
            Write-Log "Hittar ingen användbar parameteruppsättning för $command – '$key' hoppas över" -Level WARN
            continue
        }

        $params  = @{}
        $changed = [System.Collections.Generic.List[string]]::new()
        foreach ($paramName in $info.Parameters.Keys) {
            $propName = $paramName
            if (-not $obj.PSObject.Properties[$propName]) {
                $propName = $script:PropertyAliases[$paramName]
                if (-not $propName -or -not $obj.PSObject.Properties[$propName]) { continue }
            }
            $value        = $obj.$propName
            $currentValue = if ($exists) { $current[$key].$propName } else { $null }

            if ($Noun -eq 'TenantDialPlan' -and $paramName -eq 'NormalizationRules') {
                if ((Get-RuleSignature $value) -eq (Get-RuleSignature $currentValue)) { continue }
                $rules = @(foreach ($rule in @($value)) { if ($rule) { New-NormalizationRuleFromBackup $rule } })
                if ($exists) { $params[$paramName] = $rules } else { $params[$paramName] = @{ Add = $rules } }
                $changed.Add($paramName)
                continue
            }

            if ((ConvertTo-Comparable $value) -eq (ConvertTo-Comparable $currentValue)) { continue }
            if (Test-ComplexValue $value) {
                Write-Log "$Noun '$key': $paramName har ett sammansatt värde som inte återläses automatiskt – jämför manuellt mot backupfilen" -Level WARN
                continue
            }
            if ($null -eq $value) {
                $value = Get-EmptyValue $info.Parameters[$paramName]
                if ($null -eq $value) { continue }
            }
            $params[$paramName] = $value
            $changed.Add($paramName)
        }

        if ($exists -and $changed.Count -eq 0) {
            $script:Stats.Unchanged++
            Write-Verbose "$Noun '$key' är oförändrad"
            continue
        }
        $params[$info.IdParam] = $key
        $reason = if ($exists) { "Uppdatera $Noun ($($changed -join ', '))" } else { "Skapa $Noun" }
        Invoke-Change -Command $command -Parameters $params -Reason $reason
    }
}

# MicrosoftTeams 8.x heter cmdleten Get-/Set-CsOnlinePstnUsage, äldre versioner har pluralformen
function Get-PstnUsageCommand([ValidateSet('Get', 'Set')][string]$Verb) {
    foreach ($name in "$Verb-CsOnlinePstnUsage", "$Verb-CsOnlinePstnUsages") {
        if (Get-Command $name -ErrorAction SilentlyContinue) { return $name }
    }
    throw "Varken $Verb-CsOnlinePstnUsage eller $Verb-CsOnlinePstnUsages finns i MicrosoftTeams-modulen"
}

function Sync-PstnUsages([object[]]$Backup, [string]$Phase) {
    $setCommand = Get-PstnUsageCommand Set
    $desired = @($Backup | ForEach-Object { $_.Usage } | Where-Object { $_ })
    $current = @((& (Get-PstnUsageCommand Get) -ErrorAction Stop).Usage | Where-Object { $_ })
    if ($Phase -eq 'Upsert') {
        foreach ($usage in ($desired | Where-Object { $_ -notin $current })) {
            Invoke-Change -Command $setCommand -Parameters @{ Identity = 'Global'; Usage = @{ Add = $usage } } -Reason 'Lägg till PSTN usage'
        }
    }
    else {
        foreach ($usage in ($current | Where-Object { $_ -notin $desired })) {
            Invoke-Change -Command $setCommand -Parameters @{ Identity = 'Global'; Usage = @{ Remove = $usage } } -Reason 'Ta bort PSTN usage (finns inte i backup)'
        }
    }
}

function Sync-GroupPolicyAssignments([object[]]$Backup, [string]$Phase, [string[]]$PolicyType) {
    $inScope = { -not $PolicyType -or $_.PolicyType -in $PolicyType }
    $desired = @($Backup | Where-Object $inScope)
    $current = @{}
    foreach ($a in @(Get-CsGroupPolicyAssignment -ErrorAction Stop | Where-Object $inScope)) {
        $current['{0}|{1}' -f $a.GroupId, $a.PolicyType] = $a
    }

    if ($Phase -eq 'Remove') {
        $desiredKeys = @{}
        foreach ($d in $desired) { $desiredKeys['{0}|{1}' -f $d.GroupId, $d.PolicyType] = $true }
        foreach ($key in @($current.Keys)) {
            if ($desiredKeys.ContainsKey($key)) { continue }
            $a = $current[$key]
            Invoke-Change -Command 'Remove-CsGroupPolicyAssignment' -Parameters @{ GroupId = [string]$a.GroupId; PolicyType = $a.PolicyType } `
                -Reason 'Ta bort grupptilldelning (finns inte i backup)'
        }
        return
    }

    foreach ($d in $desired) {
        $a = $current['{0}|{1}' -f $d.GroupId, $d.PolicyType]
        $params = @{ GroupId = [string]$d.GroupId; PolicyType = $d.PolicyType; PolicyName = (Get-PolicyName $d.PolicyName); Rank = [int]$d.Priority }
        if (-not $a) {
            Invoke-Change -Command 'New-CsGroupPolicyAssignment' -Parameters $params -Reason 'Skapa grupptilldelning'
        }
        elseif ((Get-PolicyName $a.PolicyName) -ne $params.PolicyName -or [int]$a.Priority -ne $params.Rank) {
            Invoke-Change -Command 'Set-CsGroupPolicyAssignment' -Parameters $params -Reason 'Uppdatera grupptilldelning'
        }
    }
}

function Remove-NumberAssignment($Assignment, [string]$Reason) {
    $params = @{
        Identity        = [string]$Assignment.AssignedPstnTargetId
        PhoneNumber     = $Assignment.TelephoneNumber
        PhoneNumberType = [string]$Assignment.NumberType
    }
    Invoke-Change -Command 'Remove-CsPhoneNumberAssignment' -Parameters $params -Reason ('{0} – {1}' -f $Reason, (Get-TargetName $params.Identity))
}

function Sync-PhoneNumbers([object[]]$Backup, [string]$Phase) {
    $assigned = @($Backup | Where-Object { $_.AssignedPstnTargetId -and $_.PstnAssignmentStatus -ne 'Unassigned' })
    $current  = @(Get-AllPhoneNumberAssignments | Where-Object { $_.AssignedPstnTargetId })
    $inScope  = { -not $script:TargetFilter -or $script:TargetFilter.ContainsKey([string]$_.AssignedPstnTargetId) }

    if ($Phase -eq 'Remove') {
        $backupNumbers = @{}
        foreach ($n in $assigned) { $backupNumbers[$n.TelephoneNumber] = $true }
        foreach ($c in @($current | Where-Object $inScope)) {
            if (-not $backupNumbers.ContainsKey($c.TelephoneNumber)) {
                Remove-NumberAssignment $c 'Ta bort nummer (ej tilldelat i backup)'
            }
        }
        return
    }

    $desired = @($assigned | Where-Object $inScope)
    $currentByNumber = @{}
    foreach ($c in $current) { $currentByNumber[$c.TelephoneNumber] = $c }

    $toAssign = @($desired | Where-Object {
        $c = $currentByNumber[$_.TelephoneNumber]
        -not ($c -and $c.AssignedPstnTargetId -eq $_.AssignedPstnTargetId -and "$($c.LocationId)" -eq "$($_.LocationId)")
    })
    if ($toAssign.Count -eq 0) {
        Write-Log 'Telefonnummer: inga avvikelser mot backup'
        return
    }

    $released = @{}
    # 1) Frigör nummer som i dag sitter på ett annat mål
    foreach ($d in $toAssign) {
        $c = $currentByNumber[$d.TelephoneNumber]
        if ($c -and $c.AssignedPstnTargetId -ne $d.AssignedPstnTargetId -and -not $released.ContainsKey($c.TelephoneNumber)) {
            Remove-NumberAssignment $c ('Frigör {0} (ska tillbaka till {1})' -f $c.TelephoneNumber, (Get-TargetName $d.AssignedPstnTargetId))
            $released[$c.TelephoneNumber] = $true
        }
    }
    # 2) Ta bort nummer som berörda mål fått efter backupen
    $wantedByTarget = $desired | Group-Object { [string]$_.AssignedPstnTargetId } -AsHashTable -AsString
    foreach ($target in @($toAssign | ForEach-Object { [string]$_.AssignedPstnTargetId } | Select-Object -Unique)) {
        $wanted = @($wantedByTarget[$target] | ForEach-Object { $_.TelephoneNumber })
        foreach ($c in @($current | Where-Object { [string]$_.AssignedPstnTargetId -eq $target })) {
            if ($c.TelephoneNumber -notin $wanted -and -not $released.ContainsKey($c.TelephoneNumber)) {
                Remove-NumberAssignment $c ('Ta bort {0} (inte tilldelat målet i backup)' -f $c.TelephoneNumber)
                $released[$c.TelephoneNumber] = $true
            }
        }
    }
    # 3) Tilldela enligt backup
    foreach ($d in $toAssign) {
        $params = @{
            Identity        = [string]$d.AssignedPstnTargetId
            PhoneNumber     = $d.TelephoneNumber
            PhoneNumberType = [string]$d.NumberType
        }
        if ($d.LocationId -and "$($d.LocationId)" -ne [guid]::Empty.ToString()) { $params.LocationId = [string]$d.LocationId }
        if ($d.AssignmentCategory -eq 'Private') { $params.AssignmentCategory = 'Private' }
        Invoke-Change -Command 'Set-CsPhoneNumberAssignment' -Parameters $params -Reason ('Tilldela nummer till {0}' -f (Get-TargetName $params.Identity))
    }
}

function Sync-Users([object[]]$Backup) {
    $users = Select-FilteredUsers $Backup
    $map = Get-CurrentUserMap -ObjectIds @($users | ForEach-Object { [string]$_.ObjectId })
    foreach ($u in $users) {
        $cur = $map[[string]$u.ObjectId]
        if (-not $cur) {
            Write-Log "Användaren $($u.UserPrincipalName) finns inte längre i tenanten – hoppar över" -Level WARN
            continue
        }
        $params = @{}
        foreach ($prop in 'EnterpriseVoiceEnabled', 'HostedVoiceMail') {
            if ($null -ne $u.$prop -and [bool]$u.$prop -ne [bool]$cur.$prop) { $params[$prop] = [bool]$u.$prop }
        }
        if ($params.Count -eq 0) { $script:Stats.Unchanged++; continue }
        $params.Identity = [string]$u.ObjectId
        Invoke-Change -Command 'Set-CsUser' -Parameters $params -Reason "Uppdatera röstinställningar för $($u.UserPrincipalName)"
    }
}

function Sync-UserPolicies([object[]]$Backup, [string[]]$PolicyType) {
    $users = Select-FilteredUsers $Backup
    $map = Get-CurrentUserMap -ObjectIds @($users | ForEach-Object { [string]$_.ObjectId })
    foreach ($u in $users) {
        $cur = $map[[string]$u.ObjectId]
        if (-not $cur) {
            Write-Log "Användaren $($u.UserPrincipalName) finns inte längre i tenanten – hoppar över" -Level WARN
            continue
        }
        foreach ($prop in $u.Policies.PSObject.Properties) {
            if ($PolicyType -and $prop.Name -notin $PolicyType) { continue }
            $wanted = Get-PolicyName $prop.Value
            $actual = Get-PolicyName $cur.($prop.Name)
            if ($wanted -eq $actual) { continue }
            Invoke-Change -Command "Grant-Cs$($prop.Name)" -Parameters @{ Identity = [string]$u.ObjectId; PolicyName = $wanted } `
                -Reason ('Tilldela {0} till {1} (nu: {2})' -f $prop.Name, $u.UserPrincipalName, ($actual ?? '<ingen>'))
        }
    }
}

function Invoke-RestoreCategory([string]$Category, [string]$Phase, [string]$Path, [string[]]$PolicyType) {
    switch ($Category) {
        'PstnUsages' {
            $data = Read-BackupFile $Path 'DirectRouting/PstnUsages.json'
            if ($null -ne $data) { Sync-PstnUsages -Backup $data -Phase $Phase }
        }
        { $_ -in $script:ObjectSets.Keys } {
            $sets = @($script:ObjectSets[$Category])
            if ($Phase -eq 'Remove') { [array]::Reverse($sets) }
            foreach ($set in $sets) {
                $data = Read-BackupFile $Path $set.File
                if ($null -ne $data) { Sync-CsObjectSet -Noun $set.Noun -Desired $data -Phase $Phase }
            }
        }
        'Policies' {
            $files = @(Get-ChildItem -Path (Join-Path $Path 'Policies') -Filter '*.json' -ErrorAction SilentlyContinue)
            foreach ($file in $files) {
                if ($PolicyType -and $file.BaseName -notin $PolicyType) { continue }
                $data = Read-BackupFile $Path "Policies/$($file.Name)"
                Sync-CsObjectSet -Noun $file.BaseName -Desired $data -Phase $Phase
            }
        }
        'GroupPolicyAssignments' {
            $data = Read-BackupFile $Path 'Assignments/GroupPolicyAssignments.json'
            if ($null -ne $data) { Sync-GroupPolicyAssignments -Backup $data -Phase $Phase -PolicyType $PolicyType }
        }
        'PhoneNumbers' {
            $data = Read-BackupFile $Path 'Assignments/PhoneNumberAssignments.json'
            if ($null -ne $data) { Sync-PhoneNumbers -Backup $data -Phase $Phase }
        }
        'Users' {
            if ($Phase -eq 'Upsert') {
                $data = Read-BackupFile $Path 'Users/Users.json'
                if ($null -ne $data) { Sync-Users -Backup $data }
            }
        }
        'UserPolicies' {
            if ($Phase -eq 'Upsert') {
                $data = Read-BackupFile $Path 'Users/Users.json'
                if ($null -ne $data) { Sync-UserPolicies -Backup $data -PolicyType $PolicyType }
            }
        }
    }
}

#endregion

#region Jämförelse av backuper ---------------------------------------------------------------------

function Format-CompareValue($Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [array] -and -not (Test-ComplexValue $Value)) { return (@($Value | Where-Object { $null -ne $_ }) -join ', ') }
    return ConvertTo-Comparable $Value
}

# Plattar till nästlade objekt, t.ex. Policies.TeamsCallingPolicy eller CallingSettings.IsForwardingEnabled
function Get-FlatProperties($Object, [string]$Prefix = '') {
    $result = [ordered]@{}
    if ($null -eq $Object) { return $result }
    foreach ($p in $Object.PSObject.Properties) {
        $name = $Prefix + $p.Name
        if ($p.Value -is [System.Management.Automation.PSCustomObject]) {
            $nested = Get-FlatProperties $p.Value "$name."
            foreach ($k in $nested.Keys) { $result[$k] = $nested[$k] }
        }
        else { $result[$name] = $p.Value }
    }
    return $result
}

function Test-CompareProperty([string]$Name, [string[]]$Include, [string[]]$Exclude) {
    $leaf = ($Name -split '\.')[-1]
    if ($leaf -in $script:CompareIgnoreProperties) { return $false }
    if ($Include -and -not ($Include | Where-Object { $Name -like $_ })) { return $false }
    if ($Exclude | Where-Object { $Name -like $_ }) { return $false }
    if ($script:CompareProperty -and -not ($script:CompareProperty | Where-Object { $Name -like $_ -or $leaf -like $_ })) { return $false }
    return $true
}

function Get-CompareKey($Object, [string[]]$KeyProperty) {
    if (-not $KeyProperty) { return Get-ObjectKey $Object }
    $key = @(foreach ($p in $KeyProperty) { [string]$Object.$p }) -join '|'
    if ($key.Trim('|')) { return $key }
}

function Get-CompareLabel($Object, [string]$Key) {
    if ($Object.PSObject.Properties['UserPrincipalName'] -and $Object.UserPrincipalName) { return [string]$Object.UserPrincipalName }
    if (($Key -as [guid]) -and $Object.PSObject.Properties['Name'] -and $Object.Name) { return [string]$Object.Name }
    return $Key
}

function New-DiffRecord($Category, $Type, $Item, $Property, $Change, $Before, $After) {
    [pscustomobject][ordered]@{
        Category = $Category
        Type     = $Type
        Item     = $Item
        Property = $Property
        Change   = $Change
        Before   = $Before
        After    = $After
    }
}

# Parar ihop objekten i två backuper och returnerar en rad per egenskap som skiljer sig.
# Objekt som bara finns i ena backupen ger en rad (Tillagd/Borttagen), eller en rad per ifylld
# egenskap när urvalet är begränsat till vissa egenskaper.
function Compare-RecordSet {
    param(
        [string]$Category,
        [string]$Type,
        [AllowEmptyCollection()][object[]]$Reference,
        [AllowEmptyCollection()][object[]]$Difference,
        [string[]]$KeyProperty,
        [scriptblock]$Label,
        [string[]]$Include,
        [string[]]$Exclude
    )
    $refMap, $difMap = foreach ($side in $Reference, $Difference) {
        $map = @{}
        $items = @($side | Where-Object { $null -ne $_ })
        for ($i = 0; $i -lt $items.Count; $i++) {
            $key = Get-CompareKey $items[$i] $KeyProperty
            if (-not $key) { $key = if ($items.Count -eq 1) { $Type } else { "#$i" } }
            $map[$key] = $items[$i]
        }
        $map
    }
    $keys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($k in @($refMap.Keys) + @($difMap.Keys)) { $null = $keys.Add($k) }

    $entries = foreach ($key in $keys) {
        $obj = $difMap[$key] ?? $refMap[$key]
        [pscustomobject]@{ Key = $key; Label = ($Label ? (& $Label $obj) : (Get-CompareLabel $obj $key)) }
    }
    $propertyLimited = $Include -or $script:CompareProperty

    foreach ($entry in ($entries | Sort-Object Label)) {
        $ref = $refMap[$entry.Key]
        $dif = $difMap[$entry.Key]
        $refProps = Get-FlatProperties $ref
        $difProps = Get-FlatProperties $dif
        $names = @(@($refProps.Keys) + @($difProps.Keys) | Select-Object -Unique |
            Where-Object { Test-CompareProperty $_ $Include $Exclude })

        if ($null -eq $ref -or $null -eq $dif) {
            $change = if ($null -eq $ref) { 'Tillagd' } else { 'Borttagen' }
            if (-not $propertyLimited) {
                New-DiffRecord $Category $Type $entry.Label '' $change '' ''
                continue
            }
            $props = if ($null -eq $ref) { $difProps } else { $refProps }
            foreach ($name in $names) {
                $value = Format-CompareValue $props[$name]
                if ($value -eq '') { continue }
                if ($null -eq $ref) { New-DiffRecord $Category $Type $entry.Label $name $change '' $value }
                else { New-DiffRecord $Category $Type $entry.Label $name $change $value '' }
            }
            continue
        }

        foreach ($name in $names) {
            if ((ConvertTo-Comparable $refProps[$name]) -eq (ConvertTo-Comparable $difProps[$name])) { continue }
            $before = Format-CompareValue $refProps[$name]
            $after  = Format-CompareValue $difProps[$name]
            $change = if ($before -eq '') { 'Tillagd' } elseif ($after -eq '') { 'Borttagen' } else { 'Ändrad' }
            New-DiffRecord $Category $Type $entry.Label $name $change $before $after
        }
    }
}

# Läser samma fil ur båda backuperna. Saknas filen i den ena hoppas den över, annars skulle
# alla objekt i den felaktigt visas som tillagda eller borttagna.
function Read-BackupPair([string]$RelativePath) {
    $ref = Read-BackupFile $script:CompareRoots[0] $RelativePath -Quiet
    $dif = Read-BackupFile $script:CompareRoots[1] $RelativePath -Quiet
    if ($null -eq $ref -and $null -eq $dif) { return $null }
    if ($null -eq $ref -or $null -eq $dif) {
        Write-Log ('{0} finns bara i den {1} backupen – hoppar över' -f $RelativePath, ($null -eq $ref ? 'andra' : 'första')) -Level WARN
        return $null
    }
    return [pscustomobject]@{ Reference = $ref; Difference = $dif }
}

function ConvertTo-EffectivePolicyRecord($Entry) {
    $policies = [ordered]@{}
    foreach ($p in @($Entry.Policies)) {
        if (-not $p -or -not $p.PolicyType) { continue }
        $policies[$p.PolicyType] = '{0} via {1}' -f ($p.PolicyName ?? '<ingen>'), (Format-PolicySource $p.PolicySource)
    }
    [pscustomobject]@{ ObjectId = $Entry.ObjectId; UserPrincipalName = $Entry.UserPrincipalName; Policies = [pscustomobject]$policies }
}

function Invoke-CompareCategory([string]$Category, [string[]]$PolicyType) {
    switch ($Category) {
        'PstnUsages' {
            # Listan har inga egenskaper att filtrera på
            if ($script:CompareProperty) { return }
            $pair = Read-BackupPair 'DirectRouting/PstnUsages.json'
            if (-not $pair) { return }
            $before = @($pair.Reference | ForEach-Object { $_.Usage } | Where-Object { $_ })
            $after  = @($pair.Difference | ForEach-Object { $_.Usage } | Where-Object { $_ })
            foreach ($u in ($before | Where-Object { $_ -notin $after })) { New-DiffRecord $Category 'PstnUsage' $u '' 'Borttagen' '' '' }
            foreach ($u in ($after | Where-Object { $_ -notin $before })) { New-DiffRecord $Category 'PstnUsage' $u '' 'Tillagd' '' '' }
        }
        { $_ -in $script:ObjectSets.Keys } {
            foreach ($set in $script:ObjectSets[$Category]) {
                $pair = Read-BackupPair $set.File
                if ($pair) { Compare-RecordSet -Category $Category -Type $set.Noun -Reference $pair.Reference -Difference $pair.Difference }
            }
        }
        'Policies' {
            $types = @(foreach ($root in $script:CompareRoots) {
                Get-ChildItem -Path (Join-Path $root 'Policies') -Filter '*.json' -ErrorAction SilentlyContinue | ForEach-Object BaseName
            }) | Sort-Object -Unique
            foreach ($type in $types) {
                if ($PolicyType -and $type -notin $PolicyType) { continue }
                $pair = Read-BackupPair "Policies/$type.json"
                if ($pair) { Compare-RecordSet -Category $Category -Type $type -Reference $pair.Reference -Difference $pair.Difference }
            }
        }
        'GroupPolicyAssignments' {
            $pair = Read-BackupPair 'Assignments/GroupPolicyAssignments.json'
            if (-not $pair) { return }
            $inScope = { -not $PolicyType -or $_.PolicyType -in $PolicyType }
            Compare-RecordSet -Category $Category -Type 'GroupPolicyAssignment' -KeyProperty GroupId, PolicyType `
                -Reference @($pair.Reference | Where-Object $inScope) -Difference @($pair.Difference | Where-Object $inScope) `
                -Label { param($a) '{0} ({1})' -f $a.GroupId, $a.PolicyType }
        }
        'PhoneNumbers' {
            $pair = Read-BackupPair 'Assignments/PhoneNumberAssignments.json'
            if (-not $pair) { return }
            # Med -User tas nummer med som satt på någon av användarna i någon av backuperna
            $numbers = $null
            if ($script:TargetFilter) {
                $numbers = @{}
                foreach ($n in @($pair.Reference) + @($pair.Difference)) {
                    if ($n -and $script:TargetFilter.ContainsKey([string]$n.AssignedPstnTargetId)) { $numbers[$n.TelephoneNumber] = $true }
                }
            }
            $ref, $dif = foreach ($side in $pair.Reference, $pair.Difference) {
                , @($side | Where-Object { $_ -and (-not $numbers -or $numbers.ContainsKey($_.TelephoneNumber)) } |
                    Select-Object @{ Name = 'AssignedTo'; Expression = { if ($_.AssignedPstnTargetId) { Get-TargetName ([string]$_.AssignedPstnTargetId) } } }, *)
            }
            Compare-RecordSet -Category $Category -Type 'PhoneNumber' -KeyProperty TelephoneNumber -Exclude AssignedPstnTargetId `
                -Reference $ref -Difference $dif -Label { param($n) $n.TelephoneNumber }
        }
        'Users' {
            $pair = Read-BackupPair 'Users/Users.json'
            if ($pair) {
                Compare-RecordSet -Category $Category -Type 'User' -KeyProperty ObjectId -Exclude 'Policies.*' `
                    -Reference (Select-FilteredUsers $pair.Reference) -Difference (Select-FilteredUsers $pair.Difference)
            }
        }
        'UserPolicies' {
            $pair = Read-BackupPair 'Users/Users.json'
            if ($pair) {
                $include = if ($PolicyType) { @($PolicyType | ForEach-Object { "Policies.$_" }) } else { 'Policies.*' }
                Compare-RecordSet -Category $Category -Type 'User' -KeyProperty ObjectId -Include $include `
                    -Reference (Select-FilteredUsers $pair.Reference) -Difference (Select-FilteredUsers $pair.Difference)
            }
        }
        'BackupOnly' {
            $files = @(foreach ($root in $script:CompareRoots) {
                $dir = Join-Path $root 'BackupOnly'
                Get-ChildItem -Path $dir -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue |
                    ForEach-Object { 'BackupOnly/' + [System.IO.Path]::GetRelativePath($dir, $_.FullName).Replace('\', '/') }
            }) | Sort-Object -Unique
            foreach ($file in $files) {
                $keyProperty = $script:CompareKeys[$file]
                # Med -User jämförs bara filer som har en rad per användare eller resurskonto
                if ($script:TargetFilter -and $keyProperty -notcontains 'ObjectId') { continue }
                $pair = Read-BackupPair $file
                if (-not $pair) { continue }
                $ref = @(Select-FilteredUsers $pair.Reference)
                $dif = @(Select-FilteredUsers $pair.Difference)
                if ($file -eq 'BackupOnly/EffectivePolicies.json') {
                    $ref = @($ref | ForEach-Object { ConvertTo-EffectivePolicyRecord $_ })
                    $dif = @($dif | ForEach-Object { ConvertTo-EffectivePolicyRecord $_ })
                }
                Compare-RecordSet -Category $Category -Type ([System.IO.Path]::GetFileNameWithoutExtension($file)) `
                    -KeyProperty $keyProperty -Reference $ref -Difference $dif
            }
        }
    }
}

#endregion

#region Vidarekoppling ----------------------------------------------------------------------------

function Format-ForwardingTarget([string]$TargetType, [string]$Target, $Settings) {
    switch ($TargetType) {
        'Voicemail'   { 'Röstbrevlåda' }
        'MyDelegates' { 'Delegater: ' + (@($Settings.Delegates | ForEach-Object { $_.Id }) -join ', ') }
        'Group'       { 'Samtalsgrupp: ' + (@($Settings.CallGroupTargets) -join ', ') }
        default       { $Target }
    }
}

# En rad per aktiv vidarekoppling. Ej svar till röstbrevlåda är standard och tas inte med.
function ConvertTo-ForwardingRecords([string]$Upn, $Settings) {
    if (-not $Settings) { return }
    if ($Settings.IsForwardingEnabled) {
        [pscustomobject]@{
            UserPrincipalName = $Upn
            Type              = [string]$Settings.ForwardingType -eq 'Simultaneous' ? 'Samtidig ringning' : 'Direkt'
            TargetType        = [string]$Settings.ForwardingTargetType
            Target            = Format-ForwardingTarget ([string]$Settings.ForwardingTargetType) $Settings.ForwardingTarget $Settings
            Delay             = $null
        }
    }
    if ($Settings.IsUnansweredEnabled -and [string]$Settings.UnansweredTargetType -ne 'Voicemail') {
        [pscustomobject]@{
            UserPrincipalName = $Upn
            Type              = 'Vid ej svar'
            TargetType        = [string]$Settings.UnansweredTargetType
            Target            = Format-ForwardingTarget ([string]$Settings.UnansweredTargetType) $Settings.UnansweredTarget $Settings
            Delay             = [string]$Settings.UnansweredDelay
        }
    }
}

#endregion

#region Publika funktioner -----------------------------------------------------------------------

function Backup-TeamsConfig {
    <#
    .SYNOPSIS
        Läser ut Teams Phone / Direct Routing-konfiguration till en tidsstämplad backupmapp.
    .PARAMETER Path
        Katalog där backupmappen skapas. Standard: aktuell katalog.
    .PARAMETER VoiceUsersOnly
        Läs bara ut användare med EnterpriseVoiceEnabled = $true (snabbare i stora tenants).
    .PARAMETER IncludeUserCallingSettings
        Läs även ut vidarekoppling, delegering och röstbrevlådeinställningar per röstanvändare.
        Tar ca 1–2 s per användare. Backas upp men återläses inte automatiskt.
    .PARAMETER IncludeEffectivePolicies
        Läs ut vilken policy som faktiskt gäller per användare och varifrån den kommer (direkt/grupp).
        Tar ca 1 s per användare. Används som facit för att kontrollera en återläsning.
    .PARAMETER SkipUsers
        Hoppa över användare (endast tenantkonfiguration).
    .EXAMPLE
        Backup-TeamsConfig -Path C:\Backup -IncludeUserCallingSettings
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Get-Location).Path,
        [switch]$VoiceUsersOnly,
        [switch]$IncludeUserCallingSettings,
        [switch]$IncludeEffectivePolicies,
        [switch]$SkipUsers
    )

    $tenant = Assert-TeamsConnection
    $tenantName = ([string]$tenant.DisplayName) -replace '[^\w\-]', '_'
    $root = Join-Path (Resolve-Path $Path).Path ('TeamsBackup_{0}_{1:yyyyMMdd-HHmmss}' -f $tenantName, (Get-Date))
    $null = New-Item -ItemType Directory -Path $root -Force

    $script:BackupRoot = $root
    $script:LogFile    = Join-Path $root 'backup.log'
    $script:Manifest   = [ordered]@{
        BackupDate          = (Get-Date).ToString('o')
        TenantId            = [string]$tenant.TenantId
        TenantName          = [string]$tenant.DisplayName
        Operator            = [Environment]::UserName
        Computer            = [Environment]::MachineName
        TeamsModuleVersion  = [string](Get-Module MicrosoftTeams).Version
        ScriptVersion       = $script:ScriptVersion
        VoiceUsersOnly      = [bool]$VoiceUsersOnly
        Items               = [ordered]@{}
        Errors              = @()
    }
    Write-Log "Backup av tenant '$($tenant.DisplayName)' ($($tenant.TenantId)) -> $root"

    # --- Direct Routing ---
    Export-Item 'DirectRouting/PstnUsages.json' { & (Get-PstnUsageCommand Get) -ErrorAction Stop }
    foreach ($sets in $script:ObjectSets.Values) {
        foreach ($set in $sets) {
            Export-Item $set.File { & "Get-Cs$($set.Noun)" -ErrorAction Stop }
        }
    }

    # --- Användare ---
    $policyProperties = @()
    if (-not $SkipUsers) {
        Write-Log 'Läser ut användare (kan ta en stund i stora tenants)...'
        $userParams = @{ ErrorAction = 'Stop' }
        if ($VoiceUsersOnly) { $userParams.Filter = 'EnterpriseVoiceEnabled -eq $true' }
        try {
            $users = @(Get-CsOnlineUser @userParams)
            if ($users.Count -gt 0) {
                # Alla egenskaper som har en motsvarande Grant-Cs<namn> räknas som policytilldelning
                $policyProperties = @($users[0].PSObject.Properties.Name |
                    Where-Object { Get-Command "Grant-Cs$_" -ErrorAction SilentlyContinue })
            }
            $records = @(foreach ($u in $users) { ConvertTo-UserRecord $u $policyProperties })
            Save-Json 'Users/Users.json' $records

            $csv = foreach ($r in $records) {
                $row = [ordered]@{}
                foreach ($p in $r.PSObject.Properties) {
                    if ($p.Name -eq 'Policies') { continue }
                    $row[$p.Name] = if ($p.Value -is [array]) { $p.Value -join ', ' } else { $p.Value }
                }
                foreach ($p in $r.Policies.PSObject.Properties) { $row[$p.Name] = $p.Value }
                [pscustomobject]$row
            }
            $csv | Export-Csv -Path (Join-Path $root 'Users/Users.csv') -NoTypeInformation -Encoding utf8 -Delimiter ';'

            if ($IncludeEffectivePolicies) {
                $i = 0
                $effective = foreach ($u in $users) {
                    $i++
                    Write-Progress -Activity 'Läser gällande policyer' -Status $u.UserPrincipalName -PercentComplete ($i * 100 / $users.Count)
                    try {
                        $assignments = @(Get-CsUserPolicyAssignment -Identity ([string]$u.Identity) -ErrorAction Stop)
                        [pscustomobject]@{
                            ObjectId          = [string]$u.Identity
                            UserPrincipalName = $u.UserPrincipalName
                            Policies          = @(foreach ($a in $assignments) {
                                [pscustomobject]@{
                                    PolicyType   = $a.PolicyType
                                    PolicyName   = Get-PolicyName $a.PolicyName
                                    PolicySource = @($a.PolicySource | ForEach-Object {
                                        [pscustomobject]@{
                                            AssignmentType = [string]$_.AssignmentType
                                            PolicyName     = Get-PolicyName $_.PolicyName
                                            Reference      = [string]$_.Reference
                                            Priority       = $_.Priority
                                        }
                                    })
                                }
                            })
                        }
                    }
                    catch { Write-Log "Gällande policyer för $($u.UserPrincipalName): $($_.Exception.Message)" -Level WARN }
                }
                Write-Progress -Activity 'Läser gällande policyer' -Completed
                Save-Json 'BackupOnly/EffectivePolicies.json' $effective

                $effectiveCsv = foreach ($e in $effective) {
                    foreach ($p in $e.Policies) {
                        [pscustomobject]@{
                            UserPrincipalName = $e.UserPrincipalName
                            PolicyType        = $p.PolicyType
                            PolicyName        = $p.PolicyName
                            Source            = Format-PolicySource $p.PolicySource
                        }
                    }
                }
                $effectiveCsv | Export-Csv -Path (Join-Path $root 'BackupOnly/EffectivePolicies.csv') -NoTypeInformation -Encoding utf8 -Delimiter ';'
            }

            if ($IncludeUserCallingSettings) {
                $voiceUsers = @($users | Where-Object EnterpriseVoiceEnabled)
                $i = 0
                $settings = foreach ($u in $voiceUsers) {
                    $i++
                    Write-Progress -Activity 'Läser samtalsinställningar' -Status $u.UserPrincipalName -PercentComplete ($i * 100 / $voiceUsers.Count)
                    $calling = $null; $voicemail = $null
                    try { $calling = Get-CsUserCallingSettings -Identity $u.UserPrincipalName -ErrorAction Stop }
                    catch { Write-Log "Samtalsinställningar för $($u.UserPrincipalName): $($_.Exception.Message)" -Level WARN }
                    try { $voicemail = Get-CsOnlineVoicemailUserSettings -Identity $u.UserPrincipalName -ErrorAction Stop }
                    catch { Write-Log "Röstbrevlåda för $($u.UserPrincipalName): $($_.Exception.Message)" -Level WARN -NoHost }
                    [pscustomobject]@{
                        ObjectId          = [string]$u.Identity
                        UserPrincipalName = $u.UserPrincipalName
                        CallingSettings   = $calling
                        VoicemailSettings = $voicemail
                    }
                }
                Write-Progress -Activity 'Läser samtalsinställningar' -Completed
                Save-Json 'BackupOnly/UserCallingSettings.json' $settings
            }
        }
        catch {
            Write-Log "Kunde inte läsa ut användare – $($_.Exception.Message)" -Level ERROR
            $script:Manifest.Errors += [pscustomobject]@{ Item = 'Users/Users.json'; Error = $_.Exception.Message }
        }
    }

    # --- Tilldelningar ---
    Export-Item 'Assignments/PhoneNumberAssignments.json' { Get-AllPhoneNumberAssignments }
    Export-Item 'Assignments/GroupPolicyAssignments.json' { Get-CsGroupPolicyAssignment -ErrorAction Stop }

    # --- Policydefinitioner ---
    $policyTypes = @($script:KnownPolicyTypes + $policyProperties) | Select-Object -Unique |
        Where-Object { $_ -notin $script:DirectRoutingPolicyTypes } | Sort-Object
    foreach ($type in $policyTypes) {
        if (-not (Get-Command "Get-Cs$type" -ErrorAction SilentlyContinue)) { continue }
        Export-Item "Policies/$type.json" { & "Get-Cs$type" -ErrorAction Stop }
    }

    # --- Endast backup (återläses manuellt) ---
    foreach ($item in $script:BackupOnlyItems.GetEnumerator()) {
        Export-Item $item.Key $item.Value
    }

    $script:Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $root 'manifest.json') -Encoding utf8

    $errors = @($script:Manifest.Errors).Count
    Write-Log ('Backup klar: {0} filer, {1} fel/varningar. Mapp: {2}' -f $script:Manifest.Items.Count, $errors, $root) -Level ($errors ? 'WARN' : 'OK')
    $script:LogFile = $null
    return Get-Item $root
}

function Restore-TeamsConfig {
    <#
    .SYNOPSIS
        Läser tillbaka hela eller delar av en backup skapad med Backup-TeamsConfig.
    .DESCRIPTION
        Jämför backupen med aktuell konfiguration och ändrar bara det som skiljer sig.
        Kör alltid först med -WhatIf för att se exakt vilka kommandon som kommer att köras.
    .PARAMETER Path
        Backupmappen (den som innehåller manifest.json).
    .PARAMETER Include
        Kategorier att återläsa. Standard 'All'. Ordningen hanteras automatiskt.
    .PARAMETER Exclude
        Kategorier att hoppa över.
    .PARAMETER User
        Begränsa till en eller flera användare/resurskonton (UPN eller ObjectId).
        Utan -Include återläses då bara PhoneNumbers, Users och UserPolicies.
    .PARAMETER PolicyType
        Begränsa Policies, GroupPolicyAssignments och UserPolicies till vissa policytyper,
        t.ex. OnlineVoiceRoutingPolicy, TenantDialPlan, TeamsCallingPolicy.
    .PARAMETER RemoveExtra
        Ta även bort objekt som skapats efter backupen (gateways, routes, policyer, nummer m.m.).
    .PARAMETER AllowDifferentTenant
        Tillåt återläsning till en annan tenant än den backupen togs från.
    .PARAMETER Force
        Fråga inte om bekräftelse innan start.
    .EXAMPLE
        Restore-TeamsConfig -Path .\TeamsBackup_Kund_20261018-090000 -WhatIf
    .EXAMPLE
        Restore-TeamsConfig -Path .\TeamsBackup_Kund_20261018-090000 -Include Gateways,VoiceRoutes,VoiceRoutingPolicies
    .EXAMPLE
        Restore-TeamsConfig -Path .\TeamsBackup_Kund_20261018-090000 -User anna@kund.se,bertil@kund.se
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('All', 'PstnUsages', 'TranslationRules', 'Gateways', 'VoiceRoutes', 'VoiceRoutingPolicies', 'DialPlans',
            'Policies', 'NetworkConfig', 'NumberBlocking', 'GroupPolicyAssignments', 'PhoneNumbers', 'Users', 'UserPolicies')]
        [string[]]$Include = @('All'),
        [ValidateSet('PstnUsages', 'TranslationRules', 'Gateways', 'VoiceRoutes', 'VoiceRoutingPolicies', 'DialPlans',
            'Policies', 'NetworkConfig', 'NumberBlocking', 'GroupPolicyAssignments', 'PhoneNumbers', 'Users', 'UserPolicies')]
        [string[]]$Exclude,
        [string[]]$User,
        [string[]]$PolicyType,
        [switch]$RemoveExtra,
        [switch]$AllowDifferentTenant,
        [switch]$Force
    )

    $Path = (Resolve-Path $Path).Path
    $manifestFile = Join-Path $Path 'manifest.json'
    if (-not (Test-Path $manifestFile)) { throw "Hittar ingen manifest.json i '$Path' – är det en backupmapp?" }
    $manifest = Get-Content $manifestFile -Raw | ConvertFrom-Json

    $tenant = Assert-TeamsConnection
    if ($manifest.TenantId -ne [string]$tenant.TenantId -and -not $AllowDifferentTenant) {
        throw "Backupen kommer från tenant $($manifest.TenantId) ($($manifest.TenantName)) men du är ansluten till $($tenant.TenantId) ($($tenant.DisplayName)). Använd -AllowDifferentTenant om det är avsiktligt."
    }

    $categories = if ($Include -contains 'All') {
        if ($User) { $script:UserCategories } else { $script:RestoreOrder }
    }
    else { $script:RestoreOrder | Where-Object { $_ -in $Include } }
    $categories = @($categories | Where-Object { $_ -notin $Exclude })

    $script:LogFile = Join-Path $Path ('restore-{0:yyyyMMdd-HHmmss}{1}.log' -f (Get-Date), ($WhatIfPreference ? '-whatif' : ''))
    $script:Stats = @{ Changed = 0; Failed = 0; Planned = 0; Unchanged = 0 }
    $script:Changes = [System.Collections.Generic.List[object]]::new()
    $reportFile = [System.IO.Path]::ChangeExtension($script:LogFile, '.csv')
    $script:CurrentUserMap = $null
    $script:TargetFilter = $null

    # Namnuppslag för loggning och -User-filter
    $script:TargetNames = @{}
    $identities = @(Get-BackupIdentities $Path)
    foreach ($id in $identities) { if ($id.ObjectId) { $script:TargetNames[$id.ObjectId] = $id.Upn } }

    if ($User) { $script:TargetFilter = Resolve-TargetFilter $identities $User }

    Write-Log ('Återläsning från backup {0} ({1}) tagen {2}' -f $manifest.TenantName, $manifest.TenantId, $manifest.BackupDate)
    Write-Log ('Kategorier: {0}' -f ($categories -join ', '))
    if ($User)        { Write-Log ('Användarfilter: {0}' -f ($User -join ', ')) }
    if ($PolicyType)  { Write-Log ('Policyfilter: {0}' -f ($PolicyType -join ', ')) }
    if ($RemoveExtra) { Write-Log 'RemoveExtra: objekt som inte finns i backupen kommer att TAS BORT' -Level WARN }
    if ($WhatIfPreference) { Write-Log 'WhatIf-läge: inga ändringar görs' -Level WHATIF }

    if (-not $WhatIfPreference -and -not $Force) {
        if (-not $PSCmdlet.ShouldContinue("Ändringar kommer att göras i tenant '$($tenant.DisplayName)'. Har du kört med -WhatIf först?", 'Bekräfta återläsning')) {
            Write-Log 'Avbrutet av användaren' -Level WARN
            $script:LogFile = $null
            return
        }
    }

    foreach ($category in $categories) {
        $script:CurrentCategory = $category
        Write-Log "--- $category ---"
        try { Invoke-RestoreCategory -Category $category -Phase Upsert -Path $Path -PolicyType $PolicyType }
        catch { $script:Stats.Failed++; Write-Log "Kategori $category avbröts – $($_.Exception.Message)" -Level ERROR }
    }

    if ($RemoveExtra) {
        $reversed = @($categories); [array]::Reverse($reversed)
        foreach ($category in $reversed) {
            if ($category -in 'Users', 'UserPolicies') { continue }
            $script:CurrentCategory = "$category (borttag)"
            Write-Log "--- $category (borttag) ---"
            try { Invoke-RestoreCategory -Category $category -Phase Remove -Path $Path -PolicyType $PolicyType }
            catch { $script:Stats.Failed++; Write-Log "Borttag i $category avbröts – $($_.Exception.Message)" -Level ERROR }
        }
    }

    if ($script:Changes.Count -gt 0) {
        $script:Changes | Export-Csv -Path $reportFile -NoTypeInformation -Encoding utf8 -Delimiter ';'
        Write-Host ''
        Write-Host ($WhatIfPreference ? 'Planerade ändringar per kategori:' : 'Ändringar per kategori:')
        $script:Changes | Group-Object Category | ForEach-Object {
            Write-Host ('  {0,-35} {1,5}' -f $_.Name, $_.Count)
        }
        Write-Host ''
    }

    $summary = if ($WhatIfPreference) {
        'Torrsimning klar – inget har ändrats: {0} planerade ändringar, {1} oförändrade objekt.' -f $script:Stats.Planned, $script:Stats.Unchanged
    }
    else {
        'Återläsning klar: {0} ändringar, {1} fel, {2} oförändrade objekt.' -f $script:Stats.Changed, $script:Stats.Failed, $script:Stats.Unchanged
    }
    Write-Log $summary -Level ($script:Stats.Failed ? 'WARN' : 'OK')
    Write-Log "Logg: $($script:LogFile)"
    if ($script:Changes.Count -gt 0) { Write-Log "Ändringslista: $reportFile" }
    $script:LogFile = $null
    $script:Changes = $null
    $script:CurrentCategory = $null
}

function Get-TeamsConfigBackup {
    <#
    .SYNOPSIS
        Visar information om en backup och antal objekt per fil.
    .EXAMPLE
        Get-TeamsConfigBackup -Path .\TeamsBackup_Kund_20261018-090000 | Format-Table
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $manifest = Get-Content (Join-Path $Path 'manifest.json') -Raw | ConvertFrom-Json
    Write-Host ('Tenant:  {0} ({1})' -f $manifest.TenantName, $manifest.TenantId)
    Write-Host ('Datum:   {0}' -f $manifest.BackupDate)
    Write-Host ('Av:      {0} på {1}, MicrosoftTeams {2}' -f $manifest.Operator, $manifest.Computer, $manifest.TeamsModuleVersion)
    foreach ($e in @($manifest.Errors)) { Write-Host ('Fel:     {0} – {1}' -f $e.Item, $e.Error) -ForegroundColor Yellow }
    foreach ($item in $manifest.Items.PSObject.Properties) {
        [pscustomobject]@{
            File       = $item.Name
            Objects    = $item.Value
            Restorable = -not $item.Name.StartsWith('BackupOnly/')
        }
    }
}

function Compare-TeamsConfigBackup {
    <#
    .SYNOPSIS
        Visar skillnaderna mellan två backuper. Kräver ingen anslutning till Teams.
    .DESCRIPTION
        Returnerar en rad per egenskap som skiljer sig, med kolumnerna Category, Type, Item, Property,
        Change (Tillagd, Borttagen, Ändrad), Before (värdet i ReferencePath) och After (värdet i DifferencePath).
    .PARAMETER ReferencePath
        Backupen att jämföra från, normalt den äldre.
    .PARAMETER DifferencePath
        Backupen att jämföra med, normalt den nyare.
    .PARAMETER Include
        Kategorier att jämföra. Samma kategorier som Restore-TeamsConfig, plus BackupOnly. Standard 'All'.
    .PARAMETER Exclude
        Kategorier att hoppa över.
    .PARAMETER User
        Begränsa till en eller flera användare/resurskonton (UPN eller ObjectId).
        Utan -Include jämförs då PhoneNumbers, Users, UserPolicies och användarfilerna under BackupOnly.
    .PARAMETER PolicyType
        Begränsa Policies, GroupPolicyAssignments och UserPolicies till vissa policytyper.
    .PARAMETER Property
        Visa bara dessa egenskaper. Jokertecken tillåts, t.ex. LineUri, *VoiceRouting*, Policies.*.
    .PARAMETER CsvPath
        Spara resultatet som semikolonseparerad CSV-fil.
    .EXAMPLE
        Compare-TeamsConfigBackup .\TeamsBackup_Kund_20261001-090000 .\TeamsBackup_Kund_20261018-090000 | Format-Table
    .EXAMPLE
        Compare-TeamsConfigBackup $gammal $ny -Include Users -Property LineUri | Where-Object Change -eq 'Borttagen'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][string]$ReferencePath,
        [Parameter(Mandatory, Position = 1)][string]$DifferencePath,
        [ValidateSet('All', 'PstnUsages', 'TranslationRules', 'Gateways', 'VoiceRoutes', 'VoiceRoutingPolicies', 'DialPlans',
            'Policies', 'NetworkConfig', 'NumberBlocking', 'GroupPolicyAssignments', 'PhoneNumbers', 'Users', 'UserPolicies', 'BackupOnly')]
        [string[]]$Include = @('All'),
        [ValidateSet('PstnUsages', 'TranslationRules', 'Gateways', 'VoiceRoutes', 'VoiceRoutingPolicies', 'DialPlans',
            'Policies', 'NetworkConfig', 'NumberBlocking', 'GroupPolicyAssignments', 'PhoneNumbers', 'Users', 'UserPolicies', 'BackupOnly')]
        [string[]]$Exclude,
        [string[]]$User,
        [string[]]$PolicyType,
        [string[]]$Property,
        [string]$CsvPath
    )

    $roots = @($ReferencePath, $DifferencePath | ForEach-Object { (Resolve-Path $_).Path })
    $refManifest, $difManifest = foreach ($root in $roots) {
        $file = Join-Path $root 'manifest.json'
        if (-not (Test-Path $file)) { throw "Hittar ingen manifest.json i '$root' – är det en backupmapp?" }
        Get-Content $file -Raw | ConvertFrom-Json
    }

    Write-Host ('Före:  {0} {1}  ({2})' -f $refManifest.TenantName, $refManifest.BackupDate, $roots[0])
    Write-Host ('Efter: {0} {1}  ({2})' -f $difManifest.TenantName, $difManifest.BackupDate, $roots[1])
    if ($refManifest.TenantId -ne $difManifest.TenantId) {
        Write-Log "Backuperna kommer från olika tenants ($($refManifest.TenantId) och $($difManifest.TenantId))" -Level WARN
    }
    if ([datetime]$refManifest.BackupDate -gt [datetime]$difManifest.BackupDate) {
        Write-Log 'Den första backupen är nyare än den andra – Before och After är omvända i tid' -Level WARN
    }
    if ([bool]$refManifest.VoiceUsersOnly -ne [bool]$difManifest.VoiceUsersOnly) {
        Write-Log 'Bara den ena backupen togs med -VoiceUsersOnly – användare utan Enterprise Voice visas som tillagda eller borttagna' -Level WARN
    }

    $allCategories = $script:RestoreOrder + 'BackupOnly'
    $categories = if ($Include -contains 'All') {
        if ($User) { $script:UserCategories + 'BackupOnly' } else { $allCategories }
    }
    else { $allCategories | Where-Object { $_ -in $Include } }
    $categories = @($categories | Where-Object { $_ -notin $Exclude })

    $script:CompareRoots    = $roots
    $script:CompareProperty = $Property
    $script:TargetFilter    = $null
    $script:TargetNames     = @{}
    $identities = @(foreach ($root in $roots) { Get-BackupIdentities $root })
    foreach ($id in $identities) { if ($id.ObjectId) { $script:TargetNames[$id.ObjectId] = $id.Upn } }

    try {
        if ($User) { $script:TargetFilter = Resolve-TargetFilter $identities $User }

        $results = [System.Collections.Generic.List[object]]::new()
        foreach ($category in $categories) {
            try { foreach ($r in @(Invoke-CompareCategory -Category $category -PolicyType $PolicyType)) { $results.Add($r) } }
            catch { Write-Log "Jämförelsen av $category avbröts – $($_.Exception.Message)" -Level ERROR }
        }
    }
    finally {
        $script:CompareRoots = $null
        $script:CompareProperty = $null
        $script:TargetFilter = $null
        $script:TargetNames = $null
    }

    if ($results.Count -eq 0) {
        Write-Host 'Inga skillnader.'
        return
    }
    Write-Host ('{0} skillnader: {1}' -f $results.Count, (($results | Group-Object Category | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join ', '))
    if ($CsvPath) {
        $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8 -Delimiter ';'
        Write-Host "Sparad: $CsvPath"
    }
    $results
}

function Get-TeamsCallForwarding {
    <#
    .SYNOPSIS
        Listar användare som har vidarekoppling aktiverad. Ändrar ingenting.
    .DESCRIPTION
        Returnerar en rad per aktiv vidarekoppling med kolumnerna UserPrincipalName, Type (Direkt,
        Samtidig ringning, Vid ej svar), TargetType, Target och Delay. Vidarekoppling vid ej svar till
        röstbrevlådan är standard och visas inte.
        Utan -Path läses inställningarna från tenanten (ca 1 s per röstanvändare). Med -Path läses de
        från en backup tagen med -IncludeUserCallingSettings, och ingen anslutning krävs.
    .PARAMETER Path
        Backupmapp att läsa från i stället för tenanten.
    .PARAMETER User
        Begränsa till en eller flera användare (UPN).
    .PARAMETER CsvPath
        Spara resultatet som semikolonseparerad CSV-fil.
    .EXAMPLE
        Get-TeamsCallForwarding -CsvPath C:\Kund\vidarekoppling.csv | Format-Table
    .EXAMPLE
        Get-TeamsCallForwarding -Path .\TeamsBackup_Kund_20261018-090000 | Format-Table
    #>
    [CmdletBinding()]
    param(
        [string]$Path,
        [string[]]$User,
        [string]$CsvPath
    )

    if ($Path) {
        $entries = Read-BackupFile (Resolve-Path $Path).Path 'BackupOnly/UserCallingSettings.json' -Quiet
        if ($null -eq $entries) { throw "Backupen saknar BackupOnly/UserCallingSettings.json – den togs utan -IncludeUserCallingSettings." }
        if ($User) { $entries = @($entries | Where-Object UserPrincipalName -in $User) }
        $results = @(foreach ($e in $entries) { ConvertTo-ForwardingRecords $e.UserPrincipalName $e.CallingSettings })
    }
    else {
        $null = Assert-TeamsConnection
        $upns = if ($User) { $User }
                else { @(Get-CsOnlineUser -Filter 'EnterpriseVoiceEnabled -eq $true' -ErrorAction Stop | ForEach-Object UserPrincipalName) }
        $i = 0
        $results = @(foreach ($upn in $upns) {
            $i++
            Write-Progress -Activity 'Läser vidarekopplingar' -Status $upn -PercentComplete ($i * 100 / $upns.Count)
            try { ConvertTo-ForwardingRecords $upn (Get-CsUserCallingSettings -Identity $upn -ErrorAction Stop) }
            catch { Write-Log "Samtalsinställningar för ${upn}: $($_.Exception.Message)" -Level WARN }
        })
        Write-Progress -Activity 'Läser vidarekopplingar' -Completed
    }

    $userCount = @($results | Select-Object -ExpandProperty UserPrincipalName -Unique).Count
    Write-Host ('{0} användare med aktiv vidarekoppling' -f $userCount)
    if ($results.Count -eq 0) { return }
    if ($CsvPath) {
        $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8 -Delimiter ';'
        Write-Host "Sparad: $CsvPath"
    }
    $results
}

#endregion

Export-ModuleMember -Function Backup-TeamsConfig, Restore-TeamsConfig, Get-TeamsConfigBackup, Compare-TeamsConfigBackup, Get-TeamsCallForwarding

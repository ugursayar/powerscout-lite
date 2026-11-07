<#
.SYNOPSIS
    PowerScout - Power Platform waste audit. Finds Dataverse capacity add-ons you no longer
    need, unused premium licenses, orphaned flows/apps and empty environments, and quantifies
    how much money you can recover.

.DESCRIPTION
    Runs entirely inside YOUR tenant (read-only). No data ever leaves your environment -
    PowerScout has no servers and collects nothing.

    Produces:
      * CSV exports for each finding (for remediation / renewal negotiation)
      * A console summary

    Findings:
      * Redundant capacity add-ons     (Dataverse DB/File/Log you buy but no longer need
                                        after Microsoft's Dec-2025 / Apr-2026 entitlement increases)
      * Unassigned premium seats       (Power Apps / Power Automate premium bought, assigned to no one)
      * Disabled-but-licensed users    (blocked accounts still holding premium licenses)
      * Orphaned flows & apps          (flows: no valid owner or co-owner; apps: owner deleted or disabled) [risk]
      * Ownership not verified         (solution-aware flows, owner held in Dataverse; listed, never counted as orphaned)
      * Empty environments             (no apps, no flows, still provisioned) [advisory]

    Data sources: Microsoft Graph (licenses/users) + Power Platform admin PowerShell
    (environments/apps/flows, optional) + your Power Platform admin center capacity numbers
    (CSV or parameters - there is no clean public capacity API; see README).

.PARAMETER PriceTablePath
    CSV of prices: capacity $/GB/month per type + premium license $/user/month.
    Default: .\config\PPPriceTable.csv

.PARAMETER EntitlementTablePath
    CSV of Dataverse capacity entitlements (one-time default when the tenant holds a qualifying
    license + per-license accruals). Entitlements are DATA, not code - Microsoft changed them twice within six
    months (Dec 2025, Apr 2026). Default: .\config\EntitlementTable.csv

.PARAMETER CapacityCsvPath
    CSV with your per-environment usage from Power Platform admin center -> Resources ->
    Capacity: columns Environment,DbUsedGB,FileUsedGB,LogUsedGB. Totals are summed.

.PARAMETER DbUsedGB / FileUsedGB / LogUsedGB
    Tenant-total usage overrides if you'd rather not build the CSV.

.PARAMETER AddOnDbGB / AddOnFileGB / AddOnLogGB
    Capacity add-ons you currently PAY for (from your subscriptions page). Default 0.

.PARAMETER Lite
    Free mode: capacity position + orphaned CSV + summary only (no $ totals, no HTML).

.PARAMETER MockDataPath
    Internal/testing: load tenant data from a JSON file instead of connecting.

.EXAMPLE
    .\PowerScout-Audit.ps1 -CapacityCsvPath .\my-capacity.csv -AddOnDbGB 10 -AddOnFileGB 50

.NOTES
    PowerScout  |  powerscout.dev  |  read-only, runs in your tenant.
    Independent tool - not affiliated with Microsoft. "Power Platform", "Power Apps",
    "Power Automate" and "Dataverse" are Microsoft trademarks.
#>

[CmdletBinding()]
param(
    [string] $PriceTablePath       = (Join-Path (Join-Path $PSScriptRoot 'config') 'PPPriceTable.csv'),
    [string] $EntitlementTablePath = (Join-Path (Join-Path $PSScriptRoot 'config') 'EntitlementTable.csv'),
    [string] $CapacityCsvPath,
    [double] $DbUsedGB   = -1,
    [double] $FileUsedGB = -1,
    [double] $LogUsedGB  = -1,
    [double] $AddOnDbGB   = 0,
    [double] $AddOnFileGB = 0,
    [double] $AddOnLogGB  = 0,
    [string] $OutputPath,
    [switch] $Lite,
    [string] $MockDataPath
)

$ErrorActionPreference = 'Stop'
$script:Version = '1.0.2'
$script:Edition = 'Lite'   # build edition (set per package): Lite | Solo | Pro
$nowUtc = (Get-Date).ToUniversalTime()

function Write-Step { param($Msg) Write-Host "[PowerScout] $Msg" -ForegroundColor Cyan }
function Write-Warn { param($Msg) Write-Host "[PowerScout] $Msg" -ForegroundColor Yellow }

# ----------------------------------------------------------------------------------------------
# Config tables (data, not code)
# ----------------------------------------------------------------------------------------------
function Get-PriceTable {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Price table not found: $Path" }
    $rows = Import-Csv -Path $Path | Where-Object { $_.Key -and -not $_.Key.StartsWith('#') }
    $map = @{}
    foreach ($r in $rows) {
        $val = $null
        if ($r.Value -and ($r.Value -as [double]) -ne $null) { $val = [double]$r.Value }
        $map[$r.Key] = [pscustomobject]@{
            Key = $r.Key; Value = $val
            FriendlyName = $r.FriendlyName
            IsPPLicense  = ($r.PSObject.Properties['IsPPLicense'] -and $r.IsPPLicense -eq '1')
        }
    }
    return $map
}

function Get-EntitlementTable {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Entitlement table not found: $Path" }
    Import-Csv -Path $Path | Where-Object { $_.Rule -and -not $_.Rule.StartsWith('#') } | ForEach-Object {
        [pscustomobject]@{
            Rule          = $_.Rule                     # display name of the entitlement rule
            Kind          = $_.Kind                     # tenantbase | base | perlicense
            SkuPartNumber = $_.SkuPartNumber            # blank for tenantbase
            DbGB          = [double]($_.DbGB   | ForEach-Object { if ($_ -eq '') { 0 } else { $_ } })
            FileGB        = [double]($_.FileGB | ForEach-Object { if ($_ -eq '') { 0 } else { $_ } })
            LogGB         = [double]($_.LogGB  | ForEach-Object { if ($_ -eq '') { 0 } else { $_ } })
        }
    }
}

function ConvertFrom-FlowPermissions {
    # Owners and co-owners (roleName Owner or CanEdit) from a raw /flows/{name}/permissions response,
    # as @{type; id}, every principal type kept. Pure (no API call) so test/check_cases.ps1 can run it
    # on a fixture. Throws when the response carries no permission list.
    param($Response)
    if ($null -eq $Response -or $null -eq $Response.value) { throw 'no permissions list in the response' }
    return ,@($Response.value | Where-Object { [string]$_.properties.roleName -in 'Owner','CanEdit' } | ForEach-Object {
        [pscustomobject]@{ type = [string]$_.properties.principal.type; id = [string]$_.properties.principal.id } })
}

# ----------------------------------------------------------------------------------------------
# Data acquisition (Graph + PP admin module + capacity inputs, OR mock)
# ----------------------------------------------------------------------------------------------
function Get-TenantData {
    param([string]$MockPath)

    if ($MockPath) {
        Write-Step "Loading MOCK data from $MockPath"
        return (Get-Content -Path $MockPath -Raw | ConvertFrom-Json)
    }

    Write-Step "Connecting to Microsoft Graph (read-only)..."
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
    } catch {
        throw "Microsoft Graph PowerShell SDK not installed. Run:  Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Connect-MgGraph -Scopes @('Organization.Read.All','User.Read.All','Directory.Read.All') -NoWelcome

    $org  = Get-MgOrganization
    $skus = Get-MgSubscribedSku -All
    Write-Step "Reading users (license assignments)..."
    $users = Get-MgUser -All -Property 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses'

    # Power Platform admin module - optional; findings degrade gracefully without it.
    $apps = @(); $flows = @(); $environments = @(); $ppOK = $true
    try {
        Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
        Write-Step "Connecting to Power Platform admin (you may be prompted)..."
        Add-PowerAppsAccount | Out-Null

        Write-Step "Reading environments / apps / flows (read-only)..."
        $environments = @(Get-AdminPowerAppEnvironment | ForEach-Object {
            [pscustomobject]@{ name = $_.EnvironmentName; displayName = $_.DisplayName; isDefault = [bool]$_.IsDefault }
        })
        $apps = @(Get-AdminPowerApp | ForEach-Object {
            [pscustomobject]@{
                name = $_.DisplayName; environment = $_.EnvironmentName
                ownerId = $_.Owner.id; ownerUpn = $_.Owner.email
            }
        })
        $flows = @(Get-AdminFlow | ForEach-Object {
            [pscustomobject]@{
                name = $_.DisplayName; environment = $_.EnvironmentName; flowId = $_.FlowName
                ownerId = $_.CreatedBy.userId; ownerObjectId = $_.CreatedBy.objectId; ownerUpn = $_.CreatedBy.userPrincipalName
                creatorType = [string]$_.CreatedBy.userType
                state = $_.Enabled
            }
        })

        # Orphaned = no valid owner or co-owner (Microsoft). Owner roles are read only for flows whose
        # creator is not an enabled user. Same read-only GET as Get-AdminFlowOwnerRole, which drops
        # ServicePrincipal and Tenant principals. ownerRead = raw | cmdlet | failed; no owners property = unread.
        $enabledUser = @{}
        foreach ($u in $users) { if ($u.AccountEnabled) { $enabledUser[[string]$u.Id] = $true; if ($u.UserPrincipalName) { $enabledUser[([string]$u.UserPrincipalName).ToLower()] = $true } } }
        foreach ($f in $flows) {
            if ($enabledUser[[string]$f.ownerId] -or $enabledUser[[string]$f.ownerObjectId] -or ($f.ownerUpn -and $enabledUser[([string]$f.ownerUpn).ToLower()])) { continue }
            $read = 'raw'
            try {
                $perm = InvokeApi -Method GET -Route "https://{flowEndpoint}/providers/Microsoft.ProcessSimple/scopes/admin/environments/$($f.environment)/flows/$($f.flowId)/permissions?api-version={apiVersion}" -ThrowOnFailure
                $owners = ConvertFrom-FlowPermissions $perm
            } catch {
                # Get-AdminFlowOwnerRole returns nothing on error, so only trust it when it names a valid owner
                $owners = @()
                try {
                    $owners = @(Get-AdminFlowOwnerRole -EnvironmentName $f.environment -FlowName $f.flowId | Where-Object { [string]$_.RoleType -in 'Owner','CanEdit' } |
                        ForEach-Object { [pscustomobject]@{ type = [string]$_.PrincipalType; id = [string]$_.PrincipalObjectId } } |
                        Where-Object { $_.type -eq 'Group' -or $enabledUser[$_.id] })
                } catch { }
                $read = if ($owners.Count) { 'cmdlet' } else { 'failed' }
            }
            $f | Add-Member -NotePropertyName ownerRead -NotePropertyValue $read
            if ($read -ne 'failed') { $f | Add-Member -NotePropertyName owners -NotePropertyValue $owners }
        }
    } catch {
        Write-Warn "Power Platform admin module unavailable ($($_.Exception.Message.Split("`n")[0]))."
        Write-Warn "Install-Module Microsoft.PowerApps.Administration.PowerShell - continuing without app/flow findings."
        $ppOK = $false
    }

    # Capacity: CSV or parameters (PPAC has no clean public API - numbers come from you).
    $cap = $null
    if ($CapacityCsvPath) {
        Write-Step "Reading capacity usage from $CapacityCsvPath"
        $rows = Import-Csv -Path $CapacityCsvPath
        $cap = [pscustomobject]@{
            dbUsedGB   = [math]::Round((($rows | Measure-Object -Property DbUsedGB -Sum).Sum), 2)
            fileUsedGB = [math]::Round((($rows | Measure-Object -Property FileUsedGB -Sum).Sum), 2)
            logUsedGB  = [math]::Round((($rows | Measure-Object -Property LogUsedGB -Sum).Sum), 2)
            perEnvironment = @($rows | ForEach-Object {
                [pscustomobject]@{ environment=$_.Environment; dbUsedGB=[double]$_.DbUsedGB; fileUsedGB=[double]$_.FileUsedGB; logUsedGB=[double]$_.LogUsedGB }
            })
        }
    } elseif ($DbUsedGB -ge 0 -or $FileUsedGB -ge 0 -or $LogUsedGB -ge 0) {
        $cap = [pscustomobject]@{
            dbUsedGB = [math]::Max(0,$DbUsedGB); fileUsedGB = [math]::Max(0,$FileUsedGB); logUsedGB = [math]::Max(0,$LogUsedGB)
            perEnvironment = @()
        }
    }
    if ($cap) {
        $cap | Add-Member -NotePropertyName addOnDbGB   -NotePropertyValue $AddOnDbGB
        $cap | Add-Member -NotePropertyName addOnFileGB -NotePropertyValue $AddOnFileGB
        $cap | Add-Member -NotePropertyName addOnLogGB  -NotePropertyValue $AddOnLogGB
    }

    return [pscustomobject]@{
        organization    = [pscustomobject]@{ displayName = $org.DisplayName; id = $org.Id }
        ppDataAvailable = $ppOK
        capacity        = $cap
        subscribedSkus  = $skus | ForEach-Object {
            [pscustomobject]@{
                skuId = $_.SkuId; skuPartNumber = $_.SkuPartNumber
                prepaidUnits = [pscustomobject]@{ enabled = $_.PrepaidUnits.Enabled }
                consumedUnits = $_.ConsumedUnits
            }
        }
        users = $users | ForEach-Object {
            [pscustomobject]@{
                id = $_.Id; displayName = $_.DisplayName; userPrincipalName = $_.UserPrincipalName
                accountEnabled = $_.AccountEnabled
                assignedLicenses = @($_.AssignedLicenses | ForEach-Object { [pscustomobject]@{ skuId = $_.SkuId } })
            }
        }
        apps = $apps; flows = $flows; environments = $environments
    }
}

# ----------------------------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------------------------
function Invoke-Analysis {
    param($Data, $PriceMap, $EntitlementRules)

    $ppOK = if ($null -ne $Data.ppDataAvailable) { [bool]$Data.ppDataAvailable } else { $true }

    # ---- premium license position (Graph) ----
    $ppSkuParts = @($PriceMap.Values | Where-Object { $_.IsPPLicense } | ForEach-Object { $_.Key })
    $skuMeta = @(); $skuById = @{}; $notInTable = @()
    # Free / trial / developer part numbers per Microsoft's licensing CSV - never worth a warning
    $nonPaidParts = @('FLOW_FREE','POWERAPPS_INDIVIDUAL_USER','POWERAPPS_VIRAL','POWERAPPS_PER_APP_IW','POWERAPPS_DEV')
    foreach ($s in $Data.subscribedSkus) {
        if ($ppSkuParts -notcontains $s.skuPartNumber) {
            # Power Platform-looking SKU with no price-table row: name it instead of skipping it silently
            if ($s.skuPartNumber -match 'POWER_?APPS|POWER_?AUTOMATE|FLOW' -and $nonPaidParts -notcontains $s.skuPartNumber) { $notInTable += $s.skuPartNumber }
            continue
        }
        $meta = $PriceMap[$s.skuPartNumber]
        $purchased = [int]$s.prepaidUnits.enabled
        $consumed  = [int]$s.consumedUnits
        $unassigned = [math]::Max(0, $purchased - $consumed)
        $info = [ordered]@{
            SkuId=$s.skuId; PartNumber=$s.skuPartNumber
            Friendly = if ($meta.FriendlyName) { $meta.FriendlyName } else { $s.skuPartNumber }
            Price = $meta.Value
            Purchased=$purchased; Consumed=$consumed; Unassigned=$unassigned
        }
        $info = [pscustomobject]$info
        $skuById[$s.skuId] = $info; $skuMeta += $info
    }
    if ($notInTable.Count) { Write-Warn "Not audited - Power Platform SKU(s) with no row in the price table: $($notInTable -join ', '). Add an IsPPLicense=1 row in config\PPPriceTable.csv for each per-user premium licence." }

    # license counts per part number (for entitlement accrual), counted from assignments
    $licCount = @{}
    foreach ($u in $Data.users) {
        foreach ($l in @($u.assignedLicenses)) {
            $inf = $skuById[$l.skuId]; if ($inf) { $licCount[$inf.PartNumber] = 1 + [int]$licCount[$inf.PartNumber] }
        }
    }

    # ---- entitlement computation (transparent: keep the applied rules) ----
    # purchased units per part number, any SKU: a base (default) rule applies once the tenant has bought a listed license
    $purchasedByPart = @{}
    foreach ($s in $Data.subscribedSkus) { $purchasedByPart[[string]$s.skuPartNumber] = [int]$purchasedByPart[[string]$s.skuPartNumber] + [int]$s.prepaidUnits.enabled }
    $entDb = 0.0; $entFile = 0.0; $entLog = 0.0; $defaultApplied = $false
    foreach ($r in $EntitlementRules) {
        $mult = 0
        # SkuPartNumber may list several part numbers separated by | - any of them counts
        $held = 0; $bought = 0
        foreach ($p in ([string]$r.SkuPartNumber -split '\|')) { if ($p) { $held += [int]$licCount[$p]; $bought += [int]$purchasedByPart[$p] } }
        switch ($r.Kind) {
            'tenantbase' { $mult = 1 }
            'base'       { if ($bought -gt 0) { $mult = 1 } }
            'perlicense' { $mult = $held }      # assigned licenses - see BACKLOG-accrual-purchased-licences.md
        }
        if ($mult -le 0) { continue }
        if ($r.Kind -ne 'perlicense') { $defaultApplied = $true }
        $db = $r.DbGB * $mult; $file = $r.FileGB * $mult; $log = $r.LogGB * $mult
        $entDb += $db; $entFile += $file; $entLog += $log
    }
    if (-not $defaultApplied) { Write-Warn "No default Dataverse capacity applied: this tenant holds none of the licenses in config\EntitlementTable.csv. Dynamics 365 and other subscriptions also grant a default; if Power Platform admin center (Licensing > Dataverse) shows an entitlement, add a row for that license." }

    # ---- capacity add-on redundancy ----
    $cap = $Data.capacity
    $capacityAvailable = $null -ne $cap
    $capFindings = @()
    if ($capacityAvailable) {
        $types = @(
            @{ t='Database'; used=[double]$cap.dbUsedGB;   addon=[double]$cap.addOnDbGB;   ent=$entDb;   priceKey='DbPerGBMonthly' },
            @{ t='File';     used=[double]$cap.fileUsedGB; addon=[double]$cap.addOnFileGB; ent=$entFile; priceKey='FilePerGBMonthly' },
            @{ t='Log';      used=[double]$cap.logUsedGB;  addon=[double]$cap.addOnLogGB;  ent=$entLog;  priceKey='LogPerGBMonthly' }
        )
        foreach ($x in $types) {
            $surplus = [math]::Round(($x.ent + $x.addon) - $x.used, 2)
            $finding = [ordered]@{
                Type=$x.t; EntitledGB=[math]::Round($x.ent,2); AddOnGB=$x.addon; UsedGB=$x.used
                SurplusGB=$surplus
            }
            $capFindings += [pscustomobject]$finding
        }
    }

    # ---- user buckets (licensed with PP premium) ----
    $userById = @{}; $userByUpn = @{}
    foreach ($u in $Data.users) {
        $userById[[string]$u.id] = $u
        if ($u.userPrincipalName) { $userByUpn[([string]$u.userPrincipalName).ToLower()] = $u }
    }


    $disabled=@()
    foreach ($u in $Data.users) {
        $lic = @($u.assignedLicenses | ForEach-Object { $skuById[$_.skuId] } | Where-Object { $_ })
        if ($lic.Count -eq 0) { continue }
        $monthly = ($lic | Where-Object { $_.Price } | Measure-Object -Property Price -Sum).Sum
        if (-not $monthly) { $monthly = 0 }
        $row = [pscustomobject]@{
            DisplayName=$u.displayName; UserPrincipalName=$u.userPrincipalName
            Licenses=(($lic | ForEach-Object { $_.Friendly }) -join '; ')
            MonthlyCost=[math]::Round($monthly,2)
        }
        if (-not $u.accountEnabled) { $disabled += $row; continue }
    }

    # ---- orphaned apps (owner deleted or disabled) / flows (no valid owner or co-owner) ----
    $orphApps=@(); $orphFlows=@(); $unverifiedFlows=@()
    if ($ppOK) {
        $ownerState = {
            param($ids, $upn)
            $u = $null
            foreach ($id in @($ids)) { if (-not $u -and $id -and $userById.ContainsKey([string]$id)) { $u = $userById[[string]$id] } }
            if (-not $u -and $upn -and $userByUpn.ContainsKey(([string]$upn).ToLower())) { $u = $userByUpn[([string]$upn).ToLower()] }
            if (-not $u) { return 'deleted' }
            if (-not $u.accountEnabled) { return 'disabled' }
            return 'active'
        }
        foreach ($a in @($Data.apps)) {
            $st = & $ownerState $a.ownerId $a.ownerUpn
            if ($st -ne 'active') {
                $orphApps += [pscustomobject]@{ App=$a.name; Environment=$a.environment; Owner=$a.ownerUpn; OwnerStatus=$st }
            }
        }
        # counts only (no names or ids) for the one-line flow ownership summary
        $fc = @{ userId=0; objectId=0; UPN=0; none=0; raw=0; cmdlet=0; failed=0; active=0; coOwner=0; orphaned=0; notVerified=0 }
        $creatorTypes = @{}
        foreach ($f in @($Data.flows)) {
            $ct = if ($f.creatorType) { [string]$f.creatorType } else { '(not set)' }
            $creatorTypes[$ct] = 1 + [int]$creatorTypes[$ct]
            $u = $null; $by = 'none'
            if ($f.ownerId -and $userById.ContainsKey([string]$f.ownerId)) { $u = $userById[[string]$f.ownerId]; $by = 'userId' }
            elseif ($f.ownerObjectId -and $userById.ContainsKey([string]$f.ownerObjectId)) { $u = $userById[[string]$f.ownerObjectId]; $by = 'objectId' }
            elseif ($f.ownerUpn -and $userByUpn.ContainsKey(([string]$f.ownerUpn).ToLower())) { $u = $userByUpn[([string]$f.ownerUpn).ToLower()]; $by = 'UPN' }
            $fc[$by]++
            $st = if (-not $u) { 'deleted' } elseif (-not $u.accountEnabled) { 'disabled' } else { 'active' }
            if ($st -eq 'active') { $fc.active++; continue }
            $hasOwners = [bool]$f.PSObject.Properties['owners']
            $ownerRead = if ($f.ownerRead) { [string]$f.ownerRead } elseif ($hasOwners) { 'raw' } else { 'failed' }   # mocks omit ownerRead
            $fc[$ownerRead]++
            if ($hasOwners) {
                # a group, service principal, the tenant or an enabled user as owner/co-owner keeps the flow owned
                $valid = @($f.owners | Where-Object { $_.type -in 'Group','ServicePrincipal','Tenant' -or ($_.type -eq 'User' -and (& $ownerState $_.id $null) -eq 'active') })
                if ($valid.Count) { $fc.coOwner++; continue }
            }
            $row = [pscustomobject]@{ Flow=$f.name; Environment=$f.environment; Owner=$f.ownerUpn; OwnerStatus=$st; Enabled=$f.state }
            if ($f.creatorType -eq 'DataverseSystemUser' -and $by -eq 'none') {
                # solution-aware flow: the owner is held in Dataverse, which this audit does not read - never counted as orphaned
                $row.OwnerStatus = 'ownership not verified (solution-aware flow, owner held in Dataverse)'
                $unverifiedFlows += $row; $fc.notVerified++; continue
            }
            if (-not $hasOwners) { $row.OwnerStatus = "creator $st, co-owners not checked" }
            $orphFlows += $row; $fc.orphaned++
        }
        $typeText = if ($creatorTypes.Count) { ($creatorTypes.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) $($_.Value)" }) -join ', ' } else { '-' }
        Write-Step ("Flow ownership counts: flows {0} | creator matched userId {1} / objectId {2} / UPN {3} / none {4} | owner list raw {5} / cmdlet {6} / failed {7} | active {8} / co-owner {9} / orphaned {10} / not verified {11} | creator types {12}" -f `
            @($Data.flows).Count, $fc.userId, $fc.objectId, $fc.UPN, $fc.none, $fc.raw, $fc.cmdlet, $fc.failed, $fc.active, $fc.coOwner, $fc.orphaned, $fc.notVerified, $typeText)
    }

    # ---- empty environments (advisory) ----
    $emptyEnvs=@()
    if ($ppOK) {
        $envUse = @{}
        foreach ($a in @($Data.apps))  { $envUse[[string]$a.environment] = $true }
        foreach ($f in @($Data.flows)) { $envUse[[string]$f.environment] = $true }
        foreach ($e in @($Data.environments)) {
            if (-not $envUse[[string]$e.name]) {
                $emptyEnvs += [pscustomobject]@{ Environment=$e.displayName; Id=$e.name; IsDefault=[bool]$e.isDefault }
            }
        }
    }


    $result = [ordered]@{
        Org=$Data.organization
        Skus=$skuMeta
        EntitledDbGB=[math]::Round($entDb,2); EntitledFileGB=[math]::Round($entFile,2); EntitledLogGB=[math]::Round($entLog,2)
        CapacityAvailable=$capacityAvailable
        CapacityFindings=$capFindings
        PerEnvironment = if ($capacityAvailable) { @($cap.perEnvironment) } else { @() }
        Disabled=$disabled
        OrphanedApps=$orphApps; OrphanedFlows=$orphFlows; UnverifiedFlows=$unverifiedFlows; EmptyEnvironments=$emptyEnvs
        UnassignedSeats=[int](($skuMeta | Measure-Object -Property Unassigned -Sum).Sum)   # Sum is $null when no audited SKU
        PpDataAvailable=$ppOK
    }
    return [pscustomobject]$result
}

# ----------------------------------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "  PowerScout v$script:Version - Power Platform waste audit" -ForegroundColor Green
Write-Host "  Read-only. Runs in your tenant. Nothing leaves your environment." -ForegroundColor DarkGray
Write-Host ""

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("PowerScout-Report-" + $nowUtc.ToString('yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$priceMap = Get-PriceTable -Path $PriceTablePath
$entRules = Get-EntitlementTable -Path $EntitlementTablePath
$data     = Get-TenantData -MockPath $MockDataPath
$R        = Invoke-Analysis -Data $data -PriceMap $priceMap -EntitlementRules $entRules

# CSV exports (always)
$R.OrphanedFlows | Export-Csv (Join-Path $OutputPath 'orphaned-flows.csv') -NoTypeInformation -Encoding utf8
$R.OrphanedApps  | Export-Csv (Join-Path $OutputPath 'orphaned-apps.csv') -NoTypeInformation -Encoding utf8
$R.Disabled      | Export-Csv (Join-Path $OutputPath 'disabled-but-licensed.csv') -NoTypeInformation -Encoding utf8
if ($R.UnverifiedFlows.Count) { $R.UnverifiedFlows | Export-Csv (Join-Path $OutputPath 'ownership-not-verified-flows.csv') -NoTypeInformation -Encoding utf8 }

if ($Lite -or $script:Edition -eq 'Lite') {
    Write-Host ""
    Write-Step "LITE mode - free edition"
    if ($R.CapacityAvailable) {
        foreach ($c in $R.CapacityFindings) {
            Write-Host ([string]::Format([cultureinfo]::InvariantCulture, "  {0,-8} entitled {1,7:N1} GB | addon {2,6:N1} GB | used {3,7:N1} GB | surplus {4,7:N1} GB", $c.Type, $c.EntitledGB, $c.AddOnGB, $c.UsedGB, $c.SurplusGB))
        }
    } else {
        Write-Host "  (no capacity inputs - pass -CapacityCsvPath or -DbUsedGB etc.)"
    }
    Write-Host ("  Unassigned premium seats : {0}" -f $R.UnassignedSeats)
    Write-Host ("  Disabled but licensed    : {0}" -f $R.Disabled.Count)
    Write-Host ("  Orphaned flows / apps    : {0} / {1}" -f $R.OrphanedFlows.Count, $R.OrphanedApps.Count)
    if ($R.UnverifiedFlows.Count) { Write-Host ("  Ownership not verified   : {0} (solution-aware flows, owner held in Dataverse - not orphaned)" -f $R.UnverifiedFlows.Count) }
    Write-Host ("  Empty environments       : {0}" -f $R.EmptyEnvironments.Count)
    Write-Host ""
    Write-Host "  Solo/Pro add the HTML report — powerscout.dev" -ForegroundColor Yellow
    Write-Host ""
    Write-Step "CSV saved to: $OutputPath"
    return
}


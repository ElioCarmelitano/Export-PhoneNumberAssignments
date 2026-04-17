<#
.SYNOPSIS
    Exports Microsoft Teams phone number assignments with user and emergency location details.

.DESCRIPTION
    Retrieves all phone number assignments from Microsoft Teams, joins them with
    voice-enabled user details and LIS (Location Information Service) emergency
    location information, then exports the result to CSV.

    The script is designed to be run interactively by a Teams administrator who
    has already connected to Microsoft Teams PowerShell via Connect-MicrosoftTeams.

.PARAMETER PageSize
    Number of phone assignments to retrieve per API call. Default is 500.
    The Get-CsPhoneNumberAssignment cmdlet is paged, so larger tenants
    require multiple calls to retrieve all records.

.PARAMETER ExportPath
    Full path to the output CSV file. The parent directory will be created
    automatically if it does not already exist.

.PARAMETER ShowGrid
    If specified, displays the results in Out-GridView after the CSV export
    has completed. Export happens first so the grid window cannot block it.

.EXAMPLE
    .\Export-PhoneAssignments.ps1 -ExportPath "C:\Reports\PhoneAssignments.csv" -ShowGrid

.NOTES
    Requires the MicrosoftTeams PowerShell module and an active connection
    established via Connect-MicrosoftTeams.
#>
[CmdletBinding()]
param(
    [int]$PageSize = 500,
    [string]$ExportPath = "C:\Temp\PhoneAssignments.csv",
    [switch]$ShowGrid
)

# --- Preflight checks ---------------------------------------------------------
# Verify an active Teams PowerShell session exists. Get-CsTenant is a cheap
# call that will throw if the session is missing or expired, giving us a clear
# error message rather than a confusing failure later on.
try {
    $null = Get-CsTenant -ErrorAction Stop
} catch {
    throw "Not connected to Microsoft Teams. Run Connect-MicrosoftTeams first."
}

# Ensure the export directory exists so Export-Csv doesn't fail at the final step.
$exportDir = Split-Path $ExportPath -Parent
if ($exportDir -and -not (Test-Path $exportDir)) {
    New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
}

try {
    # --- 1. Cache LIS Locations -------------------------------------------------
    # LIS locations hold emergency address information (company name, street,
    # city, postcode). We pull them all once into a hashtable keyed by LocationId
    # so we can look them up in O(1) time when joining with phone assignments.
    Write-Host "Retrieving LIS locations..." -ForegroundColor Cyan
    $locationMap = @{}
    Get-CsOnlineLisLocation | ForEach-Object {
        if ($_.LocationId) { $locationMap["$($_.LocationId)"] = $_ }
    }
    Write-Host ("Cached {0} locations." -f $locationMap.Count)

    # --- 2. Cache Voice-Enabled Users -------------------------------------------
    # Only users with EnterpriseVoiceEnabled = $true can have phone numbers
    # assigned, so we filter server-side to keep the cache small and the call
    # fast. ResultSize is set to MaxValue because the default cap is 1000,
    # which silently truncates results in larger tenants.
    Write-Host "Caching voice-enabled users..." -ForegroundColor Cyan
    $userLookup = @{}
    Get-CsOnlineUser -Filter {EnterpriseVoiceEnabled -eq $true} `
        -ResultSize ([int]::MaxValue) `
        -Properties DisplayName, UserPrincipalName, Identity, OnlineDialOutPolicy |
        ForEach-Object {
            # Identity is a GUID; we normalise to lowercase strings so lookups
            # are case-insensitive regardless of how the source cmdlet formats them.
            if ($_.Identity) {
                $userLookup["$($_.Identity)".ToLower()] = $_
            }
        }
    Write-Host ("Cached {0} voice users." -f $userLookup.Count)

    # --- 3. Get Phone Assignments (paged) ---------------------------------------
    # Get-CsPhoneNumberAssignment returns results in pages of up to $PageSize.
    # We use a generic List to collect pages because PowerShell's array += operator
    # copies the entire array on each addition (O(n^2) behaviour) and would be
    # very slow for tenants with thousands of numbers.
    Write-Host "Retrieving phone assignments..." -ForegroundColor Cyan
    $allNumbers = [System.Collections.Generic.List[object]]::new()
    $skip = 0

    do {
        # Wrapping the result in @(...) ensures .Count works correctly even when
        # the cmdlet returns a single object rather than an array.
        $page = @(Get-CsPhoneNumberAssignment -Top $PageSize -Skip $skip)
        if ($page.Count -gt 0) {
            $allNumbers.AddRange([object[]]$page)
            $skip += $page.Count
            Write-Host ("...fetched {0} assignments" -f $allNumbers.Count)
        }
        # We stop when a page returns fewer records than requested, which
        # indicates we have reached the end of the result set.
    } while ($page.Count -eq $PageSize)

    Write-Host ("Retrieved {0} total phone assignments." -f $allNumbers.Count)

    # --- 4. Process and Join ----------------------------------------------------
    # For each phone assignment, look up the assigned user (if any) and the
    # emergency location (if any) from the caches we built earlier, then emit
    # a flat PSCustomObject with the combined fields.
    Write-Host "Joining data..." -ForegroundColor Cyan
    $result = foreach ($n in $allNumbers) {
        $userData   = $null
        $locDetails = $null

        # Match emergency location by LocationId.
        if ($n.LocationId) {
            $locDetails = $locationMap["$($n.LocationId)"]
        }

        # Match user via Identity GUID. AssignedPstnTargetId is the identifier
        # of whatever holds the number (usually a user, but can also be a
        # resource account). We only resolve user matches here.
        if ($n.AssignedPstnTargetId) {
            $targetGUID = "$($n.AssignedPstnTargetId)".ToLower()
            if ($userLookup.ContainsKey($targetGUID)) {
                $userData = $userLookup[$targetGUID]
            }
        }

        # [ordered] guarantees column order in the CSV output regardless of
        # the PowerShell version. Empty strings are used for missing values
        # so the CSV has consistent columns on every row.
        [PSCustomObject][ordered]@{
            TelephoneNumber      = $n.TelephoneNumber
            DisplayName          = if ($userData)   { $userData.DisplayName }         else { "" }
            UserPrincipalName    = if ($userData)   { $userData.UserPrincipalName }   else { "" }
            OnlineDialOutPolicy  = if ($userData)   { $userData.OnlineDialOutPolicy } else { "" }
            LocationId           = $n.LocationId
            CompanyName          = if ($locDetails) { $locDetails.CompanyName } else { "" }
            HouseNumber          = if ($locDetails) { $locDetails.HouseNumber } else { "" }
            StreetName           = if ($locDetails) { $locDetails.StreetName } else { "" }
            City                 = if ($locDetails) { $locDetails.City }        else { "" }
            PostCode             = if ($locDetails) { $locDetails.PostalCode }  else { "" }
            PstnAssignmentStatus = $n.PstnAssignmentStatus
        }
    }

    # --- 5. Export first, then display ------------------------------------------
    # CSV export happens before Out-GridView so that closing or dismissing the
    # grid window (or it failing to render in a non-interactive session) cannot
    # prevent the file from being written.
    $result | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Success! {0} rows exported to {1}" -f @($result).Count, $ExportPath) -ForegroundColor Green

    if ($ShowGrid) {
        $result | Out-GridView -Title "Phone Assignments"
    }
}
catch {
    Write-Error "Script failed: $_"
    throw
}

# Export-PhoneAssignments

A PowerShell script for Microsoft Teams administrators that exports every phone number assignment in a tenant to CSV, enriched with the assigned user's details and the emergency location address.

## What it does

Microsoft Teams stores phone assignment data, user data, and emergency location (LIS) data in three separate places. Answering a simple question like *"who has which phone number, and what address is registered against it?"* requires joining all three. This script automates that join and produces a single flat CSV.

For each assigned phone number in the tenant, the output includes:

- The telephone number and its PSTN assignment status
- The assigned user's display name, UPN, and dial-out policy (where applicable)
- The emergency location's company name, street address, city, and postcode (where applicable)

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- The [MicrosoftTeams](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install) PowerShell module
- An account with sufficient Teams administrative permissions to run `Get-CsOnlineUser`, `Get-CsPhoneNumberAssignment`, and `Get-CsOnlineLisLocation`
- An active Teams PowerShell session established via `Connect-MicrosoftTeams` before running the script

## Usage

Connect to Teams first, then run the script:

```powershell
Connect-MicrosoftTeams
.\Export-PhoneAssignments.ps1 -ExportPath "C:\Reports\PhoneAssignments.csv" -ShowGrid
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ExportPath` | string | `C:\Temp\PhoneAssignments.csv` | Full path for the output CSV. The parent directory is created automatically if missing. |
| `PageSize` | int | `500` | Number of phone assignments to fetch per API call. |
| `ShowGrid` | switch | off | If set, opens the result set in `Out-GridView` after the CSV has been written. |

### Examples

Default export to `C:\Temp`:

```powershell
.\Export-PhoneAssignments.ps1
```

Custom path with interactive grid view:

```powershell
.\Export-PhoneAssignments.ps1 -ExportPath "D:\TeamsReports\numbers.csv" -ShowGrid
```

Unattended run with verbose progress output:

```powershell
.\Export-PhoneAssignments.ps1 -ExportPath "D:\TeamsReports\numbers.csv" -Verbose
```

## How it works

The script runs in five stages.

**1. Preflight checks.** `Get-CsTenant` is called to verify there is a live Teams session. The export directory is created if it doesn't exist.

**2. Cache LIS locations.** All emergency locations are pulled once via `Get-CsOnlineLisLocation` and stored in a hashtable keyed by `LocationId`. This gives O(1) lookups during the join stage rather than repeatedly calling the API.

**3. Cache voice-enabled users.** `Get-CsOnlineUser` is filtered server-side to `EnterpriseVoiceEnabled -eq $true` with `-ResultSize ([int]::MaxValue)` to avoid the default 1000-record cap that silently truncates results in larger tenants. Users are cached in a hashtable keyed by their lowercase `Identity` GUID.

**4. Retrieve phone assignments.** `Get-CsPhoneNumberAssignment` is paged using `-Top` and `-Skip`. Pages are collected into a `List[object]` rather than a standard array to avoid the O(n²) performance penalty of PowerShell's `+=` operator. Paging continues until a page returns fewer records than requested.

**5. Join and export.** Each phone assignment is enriched by looking up the assigned user (via `AssignedPstnTargetId`) and the emergency location (via `LocationId`) in the pre-built hashtables. The combined records are written to CSV first, then optionally displayed in `Out-GridView`. Doing the export before the grid view ensures the file is always written even if the grid is closed abruptly.

## Output columns

| Column | Source |
|--------|--------|
| TelephoneNumber | Phone assignment |
| DisplayName | User |
| UserPrincipalName | User |
| OnlineDialOutPolicy | User |
| LocationId | Phone assignment |
| CompanyName | LIS location |
| HouseNumber | LIS location |
| StreetName | LIS location |
| City | LIS location |
| PostCode | LIS location |
| PstnAssignmentStatus | Phone assignment |

## Notes and caveats

- **Users only.** The script resolves user assignments only. Numbers assigned to resource accounts (auto attendants, call queues) will appear in the output with blank user fields. Extending the script to resolve resource accounts would require caching them separately via `Get-CsOnlineApplicationInstance`.
- **Identity matching.** The join assumes that `Identity` from `Get-CsOnlineUser` matches `AssignedPstnTargetId` from `Get-CsPhoneNumberAssignment`. If user fields come back blank for numbers you know are assigned, the matching key may need to be `ObjectId` instead.
- **Unassigned numbers.** Numbers without an assigned target or location will still appear in the output, with the relevant columns empty.

## Troubleshooting

**"Not connected to Microsoft Teams."** Run `Connect-MicrosoftTeams` and sign in before executing the script.

**Blank user fields for known-assigned numbers.** Check whether your tenant returns `Identity` as a GUID or a distinguished name from `Get-CsOnlineUser`. If it's a DN, change the cache key to `ObjectId`.

**Script runs but CSV is empty.** Verify the account has permission to read phone assignments. Try running `Get-CsPhoneNumberAssignment -Top 10` manually to confirm.

**Slow performance on large tenants.** Increase `PageSize` (up to the cmdlet's maximum, typically 999) to reduce the number of round trips.

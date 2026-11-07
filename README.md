# PowerScout — Power Platform Savings Audit

Find the Power Platform spend you no longer need — redundant Dataverse capacity add-ons, unused premium licenses, orphaned flows and apps — then recover it.
PowerScout runs **read-only inside your own tenant**. It has no server and collects nothing; the report is written to your machine.

---

## This repository: the free Lite build

This is **PowerScout Lite**, free and open source under the MIT License. A run prints a console summary with the capacity position and writes CSVs of orphaned flows and apps and of disabled accounts that still hold licences; it shows no dollar totals and builds no HTML report. Anything below that mentions the HTML report, dollar totals, Solo or Pro belongs to the paid editions at [powerscout.dev](https://powerscout.dev); that code is not in this script. The `-Lite` switch changes nothing in this build: it always runs as Lite.

## What you get

| File | Purpose |
|------|---------|
| `PowerScout-Audit.ps1` | The audit engine (PowerShell + Microsoft Graph + PP admin module) |
| `config/PPPriceTable.csv` | Editable prices — capacity $/GB/mo + premium license $/user/mo |
| `config/EntitlementTable.csv` | Editable, date-stamped Dataverse entitlement rules |
| `test/mock-tenant.json` | Demo data to try the tool with no tenant access |

Licensed MIT — Solo and Pro are under the EULA.

Outputs per run: a polished **HTML savings report** + **CSV exports** for each finding.

---

## Requirements

- **PowerShell 7+** (Windows recommended) — or Windows PowerShell 5.1
- **Microsoft Graph PowerShell SDK** and the **Power Platform admin module**:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser
  ```
  Install both modules before the first run. The Power Platform admin module works in PowerShell 7; importing it prints "unapproved verbs" warnings, which are expected and harmless.
- A **Power Platform Administrator** (or Global admin) account — read-only sign-in at run time
- Your capacity usage numbers from **Power Platform admin center → Resources → Capacity**
  (there is no clean public capacity API; PowerScout takes them as a small CSV or parameters)

Without the PP admin module, PowerScout still runs — it audits licenses and capacity and tells you which findings were skipped.

---

## Quick start (try it with demo data — no tenant needed)

```powershell
.\PowerScout-Audit.ps1 -MockDataPath .\test\mock-tenant.json
```
Open the generated `PowerScout-Report-<timestamp>\PowerScout-Report.html`.

## Run against your tenant

```powershell
# 1) (first time) install the modules — see Requirements

# 2) note your usage from PPAC -> Resources -> Capacity, then either:
.\PowerScout-Audit.ps1 -DbUsedGB 24 -FileUsedGB 60 -LogUsedGB 1.5 -AddOnDbGB 10 -AddOnFileGB 50

# or put per-environment rows in a CSV (Environment,DbUsedGB,FileUsedGB,LogUsedGB):
.\PowerScout-Audit.ps1 -CapacityCsvPath .\my-capacity.csv -AddOnDbGB 10 -AddOnFileGB 50
```

`-AddOn*GB` = the add-on capacity you currently **pay** for (subscriptions page). Omit them and the add-on finding simply reports 0 cancellable.

---

## Options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-CapacityCsvPath` | — | CSV of per-environment usage (columns above); totals are summed |
| `-DbUsedGB` / `-FileUsedGB` / `-LogUsedGB` | — | Tenant-total usage, if you skip the CSV |
| `-AddOnDbGB` / `-AddOnFileGB` / `-AddOnLogGB` | `0` | Add-on capacity you currently pay for |
| `-HeadroomGB` | `2` | Safety buffer excluded from the cancellable figure |
| `-PriceTablePath` | `.\config\PPPriceTable.csv` | Your prices |
| `-EntitlementTablePath` | `.\config\EntitlementTable.csv` | Your entitlement rules |
| `-CompanyName` | — | **(Pro)** White-label the report header |
| `-CurrencySymbol` / `-CurrencyLabel` | `$` / `USD` | Display currency |
| `-Lite` | off | Free mode: position + orphan CSVs + summary only |
| `-MockDataPath` | — | Use demo JSON instead of connecting |

---

## Set your tables (important)

- **Prices** (`PPPriceTable.csv`): database ≈ $40/GB/mo, file ≈ $2/GB/mo ship as defaults — set the log price and your real license prices. The shipped license rows cover `POWERAPPS_PER_USER` (Power Apps Premium), `POWERAUTOMATE_ATTENDED_RPA` (Power Automate Premium), the legacy `FLOW_PER_USER` (Power Automate per user plan) and `POWERAPPS_PER_APP`; part numbers are historical, and Microsoft's [product names and service plan identifiers](https://learn.microsoft.com/entra/identity/users/licensing-service-plan-reference) table maps them. Add extra `IsPPLicense=1` rows for per-app/process plans (`Get-MgSubscribedSku` lists your part numbers).
- **Entitlements** (`EntitlementTable.csv`): Microsoft changed defaults twice within six months (database Dec 2025, file Apr 2026). The shipped table models the post-increase state: the default capacity (20 GB database, 20 GB file, 2 GB log) is granted once, and only when the tenant has bought a Power Apps or Power Automate license listed in the table; a tenant without one gets 0 and the run says so. The report prints **every applied rule** so you can check the math. The authoritative entitled number is on PPAC's Capacity page — if it differs, edit the table to match and re-run.

---

## What it checks

1. **Redundant capacity add-ons** — entitled + purchased vs used, per type; cancellable GB × your $/GB
2. **Unassigned premium seats** — Power Apps / Power Automate premium bought, assigned to no one
3. **Disabled but licensed** — blocked accounts still holding premium licenses
4. **Orphaned flows & apps** — flows with no valid owner or co-owner (creator deleted or disabled, and no enabled user, group or service principal owns the flow); apps whose owner was deleted or disabled *(risk list)*. Solution-aware flows created by a Dataverse system user, whose owner list is empty or unreadable, are listed separately as *ownership not verified* (the owner is held in Dataverse, which PowerScout does not read) and never counted as orphaned.
5. **No-maker-footprint licenses** *(advisory)* — premium users owning no app and no flow; verify usage before reclaiming
6. **Empty environments** *(advisory)* — no apps, no flows, still provisioned

Headline savings = items 1–3 (conservative). Advisory items are listed but not added to the total.

---

## Security & privacy

- **Read-only.** PowerScout never writes or changes anything in your tenant. Apart from sign-in (`Connect-MgGraph`, `Add-PowerAppsAccount`), every call is a `Get-*` cmdlet except one: for a flow whose creator is no longer an enabled user, it sends a GET to `/flows/{name}/permissions` through the admin module's `InvokeApi`. That is the same request `Get-AdminFlowOwnerRole` makes, but the cmdlet drops ServicePrincipal and Tenant owners, so a flow owned by a service principal would look orphaned. If the GET fails, PowerScout falls back to `Get-AdminFlowOwnerRole`.
- **No exfiltration.** No network calls except to Microsoft's own endpoints. The report is generated locally.
- Review the script before running — it's plain PowerShell, intentionally readable.

---

## Troubleshooting

- *"Power Platform admin module unavailable"* → `Install-Module Microsoft.PowerApps.Administration.PowerShell`; without it, orphan/footprint/environment findings are skipped and the report says so.
- *"No capacity inputs"* → pass `-CapacityCsvPath` or the `-DbUsedGB`-style parameters; the add-on finding needs your PPAC numbers.
- *Entitled GB doesn't match PPAC* → PPAC is authoritative; edit `EntitlementTable.csv` to match your agreement and re-run.
- *"No default Dataverse capacity applied"* → the tenant holds none of the licenses in `EntitlementTable.csv`. Dynamics 365 and other subscriptions also grant a default; if PPAC (Licensing > Dataverse) shows an entitlement, add a row for that license.
- *An orphaned flow shows "co-owners not checked"* → its creator is gone and the flow's owner list could not be read, so it may still have a valid co-owner; check the flow's owners in the Power Platform admin center.
- *"Ownership not verified"* → a solution-aware flow created by a Dataverse system user: its creator is not an Entra account and its owner list is empty or unreadable, because the owner is held in Dataverse. It is listed in `ownership-not-verified-flows.csv` and not counted as orphaned; check the owner in the flow's solution.
- *"Flow ownership counts: …"* → one counts-only line per run (no names or ids): how creators were matched, how owner lists were read and what each flow came out as. Paste it when reporting a mismatch.
- *A SKU shows "not set"* → add its price in `PPPriceTable.csv`.
- *"Not audited - Power Platform SKU(s) with no row in the price table"* → those part numbers are in your tenant but not in `PPPriceTable.csv`, so their seats are not counted; add an `IsPPLicense=1` row for each per-user premium licence and re-run.

---

PowerScout · powerscout.dev · Part of the Scout suite (SeatScout · StorageScout · PilotScout) · Independent tool, not affiliated with or endorsed by Microsoft. "Power Platform", "Power Apps", "Power Automate" and "Dataverse" are trademarks of Microsoft Corporation.

## Author

Built and maintained by Uğur Sayar, a SharePoint and Microsoft 365 architect in Istanbul ([LinkedIn](https://www.linkedin.com/in/ugursayar)). Bug reports and pull requests are welcome through GitHub issues.

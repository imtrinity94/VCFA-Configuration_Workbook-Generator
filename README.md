# VCF Automation Configuration Workbook Generator

Read-only PowerShell scripts that connect to **VCF Automation 9.x** and capture its configuration into Excel workbooks, plus a raw JSON dump for quick reference. They replace filling in the configuration workbooks by hand from the UI.

There is one generator per workbook type, each in its own folder with its own README:

| Folder | Script | Captures | Output folder |
|---|---|---|---|
| [`VM-Apps/`](VM-Apps/README.md) | `Get-VCFAVMAppsConfig.ps1` | One **VM Apps** organization: infrastructure, projects, network profiles, tags, onboarding, catalog and policies, Orchestrator endpoints | `VM Apps - <org>` |
| [`All-Apps/`](All-Apps/README.md) | `Get-VCFAAllAppsConfig.ps1` | One **All Apps** organization: catalog, blueprints, VM images, content libraries, event subscriptions, projects, policies, regions, namespaces and classes, VPCs, transit gateways, IP management, access control, integrations, identity providers, secrets, configuration | `All Apps - <org>` |
| `Provider/` | `Get-VCFAProviderConfig.ps1` | The **provider** (system) configuration | `Provider - <name>` |

## Shared behaviour

All generators work the same way:

- **Read-only:** only `GET` calls, plus the login `POST`.
- **No dependencies:** Windows PowerShell **5.1** or PowerShell **7+**. The `.xlsx` is written natively, so neither Excel nor ImportExcel is needed on the jump host.
- **Interactive by default:** each asks for the FQDN, organization, TLS choice and credentials. Each also has a `$Preset` block at the top to hard-code them, and accepts an encrypted credential file.
- **Excel + JSON:** every run writes both. `-FromJson` rebuilds the workbook from a saved JSON with no connection.
- **Never overwrites:** a second run on the same day gets `_v2`, `_v3`, …
- **Robust on large environments:** full paging with PARTIAL flags, retries with backoff, automatic re-login. One failing section never stops the rest, and every call is listed on the **Collection Log** sheet.
- **Secrets stay out:** passwords are never returned by the APIs, and any password / secret / token field is redacted before the JSON is written.

## Quick start

```powershell
cd .\All-Apps            # or VM-Apps / Provider
.\Get-VCFAAllAppsConfig.ps1
```

## Execution policy

If Windows refuses to run a script:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Unblock-File .\Get-VCFA*Config.ps1
```

If it is blocked by **Software Restriction Policies** or an enforced Group Policy, do not work around it on a customer machine. Run it from another machine that can reach VCF Automation, or ask the customer's IT for an exception or to sign the script.

## Keeping customer data out of the repository

The generated workbooks and JSON contain customer infrastructure details. The root `.gitignore` excludes the `VM Apps - *`, `All Apps - *` and `Provider - *` output folders, any `VCFA_*_Config_*` file, and credential files. The only exceptions are the lab samples under `*/samples/`.

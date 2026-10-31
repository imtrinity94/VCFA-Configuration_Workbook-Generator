# Provider Configuration Workbook Generator

`Get-VCFAProviderConfig.ps1` is a read-only PowerShell script that connects to **VCF Automation 9.x as a provider** (System organization) and captures the provider configuration into an Excel workbook, plus a raw JSON dump for quick reference.

It replaces filling in the Provider configuration workbook by hand from the UI. The sheets follow the same names as that workbook.

## What it captures

| Sheet | Contents |
|---|---|
| **Index** | Run details (FQDN, time, user, login method) and a linked table of contents with item counts |
| **Version History** | Version history row for the run |
| **Organizations** | Every organization (name, display name, URL, enabled, metadata, first user and role, content library settings). Also region quotas (region, supervisor, per-zone CPU / memory limits and reservations, VM classes added / not added, storage class limits), reserved VM classes, and per-org networking (log name, regional networking, provider gateway, edge cluster, ingress / egress, IP space quotas and usage) |
| **Networking** | Provider gateways (T0 gateways, associated IP spaces), IP spaces (blocks, region, default quotas, utilization), edge clusters, transit gateways |
| **Regions** | Regions (status, NSX Local Manager, supervisors and vCenter, zones, capacity), zones with allocation and reservation, VM classes (CPU / memory reservation, size, reserved), storage classes, reserved VM classes |
| **Supervisors** | Supervisor, region, vCenter, zones and zone capacity |
| **Content Libraries** | Provider and tenant libraries with storage classes, and library items |
| **Roles and Credentials** | The built-in and design accounts to record (passwords are never collected), plus the provider roles for reference |
| **VCF Data Services** | Data Services endpoints, if the API publishes them |
| **General Settings** | Activity log, timeouts, organization and operation limits, plus every general setting |
| **Connections** | vCenter (state, connected, user, license, VCF instance) and NSX Managers |
| **VCF Instances** | Name, host name, user, version |
| **Access Control** | Provider users, groups, roles, system tokens, service accounts, global roles and rights bundles (with the organizations they are published to) |
| **Identity Providers** | VCF SSO / OIDC, LDAP (connection, user and group attributes), SAML |
| **OIDC Services** | Relying parties, keys and the OpenID provider settings |
| **Certificate Management** | Trusted certificates and the certificates library: common name, issuer and expiry are read from the certificate itself. Private keys are never exported |
| **Branding** | Theme, platform name, colors, custom links, themes |
| **Feature Flags** | Name, usage, description, state |
| **Collection Log** | One line per API call group: status, item count, error message, duration |

Layout: objects with a handful of instances are shown side by side (#1, #2, …). Above 10 objects (configurable), a section switches to one row per object.

### Where the data comes from

- `/cloudapi/1.0.0` for organizations, users, groups, roles, global roles, rights bundles, service accounts, tokens, vCenters, feature flags, certificates and the OpenID provider.
- `/cloudapi/v1` (9.1) or `/cloudapi/vcf` (9.0) for regions, zones, supervisors, VM and storage classes, region quotas, regional networking, provider gateways, IP spaces, edge clusters, content libraries, NSX Managers and VCF instances. Where a resource has had different names across releases, each name is tried in turn.
- `/api/admin/extension/settings/*` for general and LDAP settings, and `/api/admin/org/{System}/settings/*` for OIDC and SAML.

Organization users and settings are read in each organization's context. Anything the environment does not expose shows as FAILED (or `<not configured>` for a 404) in the Collection Log, and the rest of the run continues.

## Requirements

- Windows PowerShell **5.1** or PowerShell **7+** (Windows, macOS or Linux)
- HTTPS access to the VCF Automation FQDN
- No modules needed: the `.xlsx` is written natively (no Excel or ImportExcel required)
- A provider account in the **System** organization with the **System Administrator** role

## Usage

```powershell
.\Get-VCFAProviderConfig.ps1
```

It asks for the FQDN, whether to skip TLS validation, and the provider credentials (`@System` is appended to the user name automatically).

Output goes to a folder named after the first label of the FQDN (for example `vcfa` for `vcfa.corp.local`):

```
Provider - <name>\VCFA_<name>_Provider_Config_<yyyyMMdd>.xlsx
Provider - <name>\VCFA_<name>_Provider_Config_<yyyyMMdd>.json
```

Existing files are never overwritten; a second run on the same day gets `_v2`, `_v3`, …

### Parameters

| Parameter | Purpose |
|---|---|
| `-Server`, `-Credential` | Skip the prompts |
| `-SkipCertificateCheck` | Accept self-signed / untrusted certificates |
| `-UseApiToken` | Log in with a provider API token instead of a user name and password |
| `-OutputFolder` | Parent folder for the `Provider - <name>` folder (default: current folder) |
| `-FromJson <file>` | Rebuild the workbook from a saved JSON, no connection needed |
| `-MaxColumns <n>` | Side-by-side limit before a section switches to rows (default 10) |

### Hard-coding inputs

The `$Preset` block at the top of the script can hold the FQDN, TLS choice and user name, so the script only asks for what is left empty. Prefer an encrypted credential file over a plain-text password:

```powershell
Get-Credential | Export-Clixml -Path C:\path\vcfa-cred.xml   # then set CredentialFile in $Preset
```

Only the same Windows user on the same machine can decrypt that file.

## Security

- Read-only: only `GET` calls, plus the login `POST`.
- Passwords and private keys are never written: any field that looks like a password, secret, token or private key is redacted before the JSON is saved.
- The generated workbook and JSON contain infrastructure details (FQDNs, service account names); `.gitignore` keeps them out of this repository.

See the [repository README](../README.md) for execution policy advice.

# All Apps Configuration Workbook Generator

`Get-VCFAAllAppsConfig.ps1` is a read-only PowerShell script that connects to a **VCF Automation 9.x All Apps organization** and captures its configuration into an Excel workbook, plus a raw JSON dump for quick reference.

It replaces filling in the All Apps configuration workbook by hand from the UI. The sheets follow the same names and row labels as that workbook.

## What it captures

| Sheet | Contents |
|---|---|
| **Index** | Run details (org, FQDN, time, user, login method) and a linked table of contents with item counts |
| **Version History** | Version history row for the run |
| **Catalog** | Catalog items: name, project, catalog type, icon, scope |
| **Blueprint Design** | Blueprints (source control, read-only, project, sharing, published, full YAML), input and constant property groups with their properties and the blueprints that use them, custom resources, resource actions |
| **VM Images** | Content library items: version, status, image identifier, library, type |
| **Content Libraries** | Name, status, type (tenant / provider), subscribed, version, creation date |
| **Event Subscriptions** | Topic, condition, workflow, blocking, timeout, priority, recover workflow, project scope |
| **Projects** | Projects, their users and groups, and their namespaces |
| **Policies** | Approval, Day 2 actions, lease and IaaS resource policies with their template fields; every other policy type with its full definition |
| **Regions** | Regions with zones, namespaces and CPU / memory / storage figures; zones |
| **Namespaces** | Namespace, project, status, class, region, VPC, zones; storage class usage |
| **Namespace Class** | Classes with CPU / memory limits and reservations, VM classes, storage classes, content libraries; organization VM classes and storage classes |
| **Virtual Private Clouds** | VPCs, connectivity profiles, general details and health status |
| **Transit Gateways** | Transit gateways and connected VPCs |
| **IP Management** | IP address blocks (CIDR, scope, VPCs, quotas, connectivity profiles) and IP quotas |
| **Access Control** | Users, groups, roles, service accounts |
| **External Orchestrators** | Orchestrator integrations: VIP, organization, version. Node, DNS, NTP and placement rows are marked for manual entry, since the VCF Automation API does not expose them |
| **Integrations** | Orchestrator, GitHub, GitLab, Bitbucket and any other integration type (tokens are never exported) |
| **Identity Providers** | OIDC, LDAP (source of users, organizational unit, custom LDAP details), SAML |
| **Secrets** | Names, scope and description only; values are never read |
| **Configuration** | Settings, every configuration property, organization settings, notification scenarios, email server |
| **Other CCI Resources** | Every other resource type found through the Cloud Consumption Interface, with counts and names |
| **Collection Log** | One line per API call group: status, item count, error message, duration |

Layout: objects with a handful of instances are shown side by side (#1, #2, …). Above 10 objects (configurable), a section switches to one row per object.

### Where the data comes from

- **Organization REST APIs:** catalog, blueprints, property groups, custom resources, event broker, Orchestrator gateway, project service, policies, `/cloudapi` (regions, zones, namespaces, VM and storage classes, content libraries, users, groups, roles, service accounts, organization settings), integrations, secrets, notifications. Paths are tried under `/cloudapi/v1` (9.1) first, then `/cloudapi/vcf` (9.0).
- **Cloud Consumption Interface (`/cci/kubernetes`):** the script discovers every VMware API group the organization publishes and lists every resource in it. Supervisor namespaces, namespace classes, VPCs, connectivity profiles, transit gateways, IP blocks and IP quotas are found by resource kind, so the script keeps working when API versions change. Namespaced resources are listed cluster-wide where allowed, otherwise once per project.

  Credential-like resources (secrets, tokens, kubeconfigs) are never read.

Fields whose layout differs between API versions are matched by name. Each networking and namespace class section also has a **Spec (all fields)** row with the full object, so nothing is hidden if a field moves. Every object is kept unchanged in the JSON.

## Requirements

- Windows PowerShell **5.1** or PowerShell **7+** (Windows, macOS or Linux)
- HTTPS access to the VCF Automation FQDN
- No modules needed: the `.xlsx` is written natively (no Excel or ImportExcel required)
- An **Organization Administrator** account in the All Apps organization. With fewer rights, sections the account cannot read show as FAILED in the Collection Log.

## Usage

```powershell
.\Get-VCFAAllAppsConfig.ps1
```

It asks for the FQDN, organization, whether to skip TLS validation, and credentials (`@<org>` is appended to the user name automatically).

Output goes to a per-org folder:

```
All Apps - <org>\VCFA_<org>_AllApps_Config_<yyyyMMdd>.xlsx
All Apps - <org>\VCFA_<org>_AllApps_Config_<yyyyMMdd>.json
```

Existing files are never overwritten; a second run on the same day gets `_v2`, `_v3`, …

### Parameters

| Parameter | Purpose |
|---|---|
| `-Server`, `-Organization`, `-Credential` | Skip the prompts |
| `-SkipCertificateCheck` | Accept self-signed / untrusted certificates |
| `-UseApiToken` | Log in with an organization API token instead of a user name and password |
| `-OutputFolder` | Parent folder for the `All Apps - <org>` folder (default: current folder) |
| `-FromJson <file>` | Rebuild the workbook from a saved JSON, no connection needed |
| `-MaxColumns <n>` | Side-by-side limit before a section switches to rows (default 10) |
| `-SkipKubernetes` | Skip the Cloud Consumption Interface discovery (namespace, VPC and IP sections then come from the REST APIs only) |

### Hard-coding inputs

The `$Preset` block at the top of the script can hold the FQDN, organization, TLS choice and user name, so the script only asks for what is left empty. Prefer an encrypted credential file over a plain-text password:

```powershell
Get-Credential | Export-Clixml -Path C:\path\vcfa-cred.xml   # then set CredentialFile in $Preset
```

Only the same Windows user on the same machine can decrypt that file.

## Behaviour on large environments

- Pages through every list API (including Kubernetes `continue` tokens), following the server's own totals. A short result is flagged **PARTIAL** in the Collection Log instead of being silently truncated.
- Retries throttling and 5xx errors with backoff, and logs in again automatically if the session expires mid-run.
- One failing section never stops the rest; it is reported in the Collection Log, and the sheet shows `<not collected - see Collection Log sheet>`.

## Security

- Read-only: only `GET` calls, plus the login `POST`.
- Passwords are never returned by the APIs. Any field that looks like a password, secret, token or private key is redacted before the JSON is written.
- Blueprint YAML is captured as written. If a blueprint contains a password in plain text, it will appear in the workbook.
- The generated workbook and JSON can contain customer infrastructure details; `.gitignore` keeps them out of this repository.

See the [repository README](../README.md) for execution policy advice.

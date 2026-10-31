#Requires -Version 5.1
<#
.SYNOPSIS
    Collects the configuration of one VCF Automation 9.x "All Apps" organization and writes
    a configuration workbook (.xlsx) plus a raw JSON dump for quick reference.

.DESCRIPTION
    Read-only. Uses GET calls only (plus the login POST). Prompts for anything not passed:
    VCF Automation FQDN, organization name, TLS check, and credentials.

    Covers the sheets of the All Apps configuration workbook: Catalog, Blueprint Design, VM Images,
    Content Libraries, Event Subscriptions, Projects, Policies, Regions, Namespaces, Namespace Class,
    Virtual Private Clouds, Transit Gateways, IP Management, Access Control, External Orchestrators,
    Integrations, Identity Providers, Secrets and Configuration.

    Sources:
      - the organization REST APIs (/cloudapi, /catalog, /blueprint, /policy, /event-broker,
        /project-service, /iaas, /platform, /notification, /form-service, /properties, /vro)
      - the Cloud Consumption Interface (/cci/kubernetes). Every VMware API group it publishes is
        discovered at run time and every listable resource is captured, so namespaces, namespace
        classes, VPCs, transit gateways, connectivity profiles, IP blocks and IP quotas are found
        whatever API version the environment runs.

    Output (in a sub-folder "All Apps - <Org>" of -OutputFolder, default = current folder):
        All Apps - <Org>\VCFA_<Org>_AllApps_Config_<yyyyMMdd>.xlsx
        All Apps - <Org>\VCFA_<Org>_AllApps_Config_<yyyyMMdd>.json
    If either file already exists, a suffix _v2, _v3 ... is added so nothing is overwritten.

    The workbook is written natively (Open XML), so no ImportExcel / Excel install is needed
    on the customer jump host.

    Passwords are never returned by the APIs; any field that looks like a password, secret
    or token is redacted before the JSON is saved.

.PARAMETER Server
    VCF Automation FQDN, e.g. vcfa.corp.local

.PARAMETER Organization
    All Apps organization name (as used at login).

.PARAMETER Credential
    Org user. "@<Organization>" is appended to the user name if it is not already there.

.PARAMETER UseApiToken
    Log in with an organization API token instead of user name / password.

.PARAMETER SkipCertificateCheck
    Skip TLS certificate validation (self-signed / untrusted CA). Asked interactively if omitted.

.PARAMETER SkipKubernetes
    Skip the Cloud Consumption Interface (/cci/kubernetes) discovery. Namespace, VPC, transit gateway
    and IP management sections are then filled only from the REST APIs.

.PARAMETER MaxColumns
    Objects shown side by side (#1, #2 ...) up to this count; above it the section switches to a row-per-object table. Default 10.

.PARAMETER FromJson
    Re-build the workbook from a JSON file produced by an earlier run (no connection needed).

.EXAMPLE
    ./Get-VCFAAllAppsConfig.ps1
    Fully interactive.

.EXAMPLE
    ./Get-VCFAAllAppsConfig.ps1 -Server vcfa.corp.local -Organization my-allapps-org -SkipCertificateCheck

.EXAMPLE
    ./Get-VCFAAllAppsConfig.ps1 -FromJson .\VCFA_my-allapps-org_AllApps_Config_20261102.json
#>
[CmdletBinding(DefaultParameterSetName = 'Collect')]
param(
    [Parameter(ParameterSetName = 'Collect')][string]$Server,
    [Parameter(ParameterSetName = 'Collect')][string]$Organization,
    [Parameter(ParameterSetName = 'Collect')][pscredential]$Credential,
    [Parameter(ParameterSetName = 'Collect')][switch]$UseApiToken,
    [Parameter(ParameterSetName = 'Collect')][switch]$SkipCertificateCheck,
    [Parameter(ParameterSetName = 'Collect')][string]$BaseUri,
    [Parameter(ParameterSetName = 'Render', Mandatory)][string]$FromJson,
    [Parameter(ParameterSetName = 'Collect')][switch]$SkipKubernetes,
    [int]$MaxColumns = 10,
    [string]$OutputFolder = (Get-Location).Path
)

# =====================================================================================
#  OPTIONAL HARD-CODED INPUTS
#  Fill in any of these to skip the matching prompt. Leave a value empty ('' or $null) to be asked.
#  Command-line parameters always win over these values.
# =====================================================================================
$Preset = @{
    Server               = ''        # VCF Automation FQDN, e.g. 'vcfa.corp.local'
    Organization         = ''        # All Apps organization name, e.g. 'my-allapps-org'
    SkipCertificateCheck = $null     # $true / $false, or $null to be asked
    Username             = ''        # e.g. 'admin'  (@<Organization> is added automatically)
    Password             = ''        # Plain text - avoid on customer machines; prefer CredentialFile below
    CredentialFile       = ''        # Path to an encrypted credential file made with:
                                     #   Get-Credential | Export-Clixml -Path C:\path\vcfa-cred.xml
                                     # Only the same Windows user on the same machine can decrypt it.
}
# =====================================================================================

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
Add-Type -AssemblyName System.IO.Compression
if ($PSVersionTable.PSVersion.Major -lt 6) { Add-Type -AssemblyName System.IO.Compression.FileSystem }

# =====================================================================================
#  Generic helpers
# =====================================================================================

function Items($v) {
    if ($null -eq $v) { return , @() }
    return , @($v | Where-Object { $null -ne $_ })
}

function Get-P {
    # First non-empty value among dotted property paths, searched on one or more objects.
    param($Obj, [string[]]$Path)
    foreach ($c in @($Obj)) {
        if ($null -eq $c) { continue }
        foreach ($p in $Path) {
            $cur = $c
            foreach ($seg in $p.Split('.')) {
                if ($null -eq $cur) { break }
                if ($cur -is [System.Collections.IDictionary]) { $cur = if (([System.Collections.IDictionary]$cur).Contains($seg)) { $cur[$seg] } else { $null } }
                else {
                    $pp = $cur.PSObject.Properties[$seg]
                    $cur = if ($pp) { $pp.Value } else { $null }
                }
            }
            if ($null -eq $cur) { continue }
            if ($cur -is [string] -and $cur -eq '') { continue }
            if ($cur -is [array] -and $cur.Count -eq 0) { continue }
            return $cur
        }
    }
    return $null
}

function Get-Members($o) {
    if ($null -eq $o) { return , @() }
    if ($o -is [System.Collections.IDictionary]) {
        return , @($o.GetEnumerator() | ForEach-Object { [pscustomobject]@{ Name = "$($_.Key)"; Value = $_.Value } })
    }
    if ($o -is [System.Management.Automation.PSCustomObject]) {
        return , @($o.PSObject.Properties | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.Value } })
    }
    return , @()
}

function Get-PropsExcept($o, [string[]]$Except) {
    $d = [ordered]@{}
    foreach ($m in (Get-Members $o)) {
        if ($Except -contains $m.Name) { continue }
        if ($m.Name -match '(?i)password|privateKey$|secret|token') { continue }
        $d[$m.Name] = $m.Value
    }
    if ($d.Count -eq 0) { return $null }
    return $d
}

function LinkIds($o, [string]$Rel) {
    $l = Get-P $o "_links.$Rel"
    if ($null -eq $l) { return , @() }
    $h = @()
    $one = Get-P $l 'href'; if ($one) { $h += $one }
    $many = Get-P $l 'hrefs'; if ($many) { $h += @($many) }
    return , @($h | ForEach-Object { (("$_" -split '\?')[0].TrimEnd('/') -split '/')[-1] } | Where-Object { $_ })
}

function Format-Value {
    param($v, [int]$Depth = 0)
    if ($null -eq $v) { return '' }
    if ($v -is [bool]) { if ($v) { return 'Yes' } else { return 'No' } }
    if ($v -is [string]) {
        if ($v -match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?$') {
            $dto = [DateTimeOffset]::MinValue
            if ([DateTimeOffset]::TryParse($v, [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) { return $dto.UtcDateTime.ToString('yyyy-MM-dd HH:mm') }
        }
        return $v
    }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm') }
    if ($v -is [System.DateTimeOffset]) { return $v.ToString('yyyy-MM-dd HH:mm') }
    if ($v -is [ValueType]) { return $v }
    if ($v -is [System.Collections.IDictionary] -or $v -is [System.Management.Automation.PSCustomObject]) {
        $m = Get-Members $v
        $names = @($m | ForEach-Object { $_.Name })
        if ($names -contains 'key' -and $m.Count -le 2) {
            $k = Get-P $v 'key'; $val = Get-P $v 'value'
            if ($null -ne $val -and "$val" -ne '') { return "${k}:$val" } else { return "$k" }
        }
        if ($names -contains 'expression') {
            $hard = if (Get-P $v 'mandatory') { 'hard' } else { 'soft' }
            return "$(Get-P $v 'expression') ($hard)"
        }
        if ($Depth -ge 1) { return ($v | ConvertTo-Json -Depth 10 -Compress) }
        return (@($m | ForEach-Object { '{0}: {1}' -f $_.Name, (Format-Value $_.Value ($Depth + 1)) }) -join "`n")
    }
    if ($v -is [System.Collections.IEnumerable]) {
        return (@($v | ForEach-Object { Format-Value $_ ($Depth + 1) }) -join "`n")
    }
    return "$v"
}

function Add-RestOf($Dict, $Obj, [string[]]$Used) {
    # Appends all fields of $Obj not already shown, under their raw API names, so nothing is dropped.
    foreach ($m in (Get-Members $Obj)) {
        if ($Used -contains $m.Name) { continue }
        if ($m.Name -match '(?i)password|secret|token|privateKey$') { continue }
        if ($m.Name -match '^document(Version|Epoch|Kind|UpdateTimeMicros|UpdateAction|ExpirationTimeMicros|Owner|AuthPrincipalLink|SourceLink|SelfLink)$') { continue }
        if ($Dict.Contains($m.Name)) { continue }
        $Dict[$m.Name] = $m.Value
    }
    return $Dict
}

function Join-Lines($list) { (@($list | Where-Object { $null -ne $_ -and "$_" -ne '' }) -join "`n") }

# =====================================================================================
#  Minimal native .xlsx writer
# =====================================================================================
# Styles: 1 sheet title | 2 section banner | 3 table header | 4 data | 5 row label | 6 note | 7 link

function New-Book {
    @{ Sheets = [System.Collections.Generic.List[object]]::new(); Index = [System.Collections.Generic.List[object]]::new() }
}

function Add-Sheet($Book, [string]$Name) {
    $s = @{ Name = $Name; Rows = [System.Collections.Generic.List[object]]::new(); Links = [System.Collections.Generic.List[object]]::new() }
    $Book.Sheets.Add($s)
    return $s
}

function C($v, [int]$s = 4) { @{ v = $v; s = $s } }

function Add-Row($Sheet, [object[]]$Cells) {
    $Sheet.Rows.Add(@($Cells))
    return $Sheet.Rows.Count
}

function Add-Blank($Sheet) { [void](Add-Row $Sheet @()) }

function Add-Title($Sheet, [string]$Text, [string]$Sub) {
    [void](Add-Row $Sheet @(C $Text 1))
    if ($Sub) { [void](Add-Row $Sheet @(C $Sub 6)) }
    Add-Blank $Sheet
}

function Get-EmptyText([string[]]$Key) {
    foreach ($k in $Key) { if ($k -and $script:FailedKeys -contains $k) { return '<not collected - see Collection Log sheet>' } }
    return '<not configured>'
}

function Add-Section($Book, $Sheet, [string]$Title, [int]$Width, $Count) {
    $cells = @(C $Title 2)
    for ($i = 1; $i -lt [math]::Max(2, $Width); $i++) { $cells += (C '' 2) }
    $r = Add-Row $Sheet $cells
    $Book.Index.Add(@{ Sheet = $Sheet.Name; Row = $r; Title = $Title; Count = $Count })
}

function Get-Columns($Rows) {
    $cols = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $Rows) { foreach ($k in $r.Keys) { if (-not $cols.Contains($k)) { $cols.Add($k) } } }
    return , $cols.ToArray()
}

function Add-Table {
    # Rows-as-records layout (one object per row).
    param($Book, $Sheet, [string]$Title, $Rows, [string[]]$Key, [string[]]$Columns)
    $list = @($Rows | Where-Object { $null -ne $_ })
    if (-not $Columns) { $Columns = Get-Columns $list }
    Add-Section $Book $Sheet $Title ([math]::Max(1, $Columns.Count)) $list.Count
    if ($list.Count -eq 0) { [void](Add-Row $Sheet @(C (Get-EmptyText $Key) 6)); Add-Blank $Sheet; return }
    [void](Add-Row $Sheet @($Columns | ForEach-Object { C $_ 3 }))
    foreach ($r in $list) {
        [void](Add-Row $Sheet @($Columns | ForEach-Object { C (Format-Value $r[$_]) 4 }))
    }
    Add-Blank $Sheet
}

function Add-Vertical {
    # Properties-as-rows layout (one object per column: #1, #2 ...).
    param($Book, $Sheet, [string]$Title, $Objects, [string[]]$Key, [string]$NameKey = 'Name')
    $list = @($Objects | Where-Object { $null -ne $_ })
    $labels = Get-Columns $list
    if ($list.Count -gt $script:MaxCols) {
        # Too many objects to read side by side: fall back to one row per object
        Add-Table $Book $Sheet $Title $list $Key $labels
        return
    }
    Add-Section $Book $Sheet $Title ($list.Count + 1) $list.Count
    if ($list.Count -eq 0) { [void](Add-Row $Sheet @(C (Get-EmptyText $Key) 6)); Add-Blank $Sheet; return }
    $hdr = @(C 'Property' 3)
    for ($i = 0; $i -lt $list.Count; $i++) {
        $nm = Get-P $list[$i] $NameKey, 'Name', 'name', 'displayName', 'id'
        $hdr += (C ('#{0}  {1}' -f ($i + 1), $nm) 3)
    }
    [void](Add-Row $Sheet $hdr)
    foreach ($l in $labels) {
        $cells = @(C $l 5)
        foreach ($o in $list) { $cells += (C (Format-Value $o[$l]) 4) }
        [void](Add-Row $Sheet $cells)
    }
    Add-Blank $Sheet
}

function Add-KeyValue($Sheet, $Pairs) {
    foreach ($k in $Pairs.Keys) { [void](Add-Row $Sheet @((C $k 5), (C (Format-Value $Pairs[$k]) 4))) }
    Add-Blank $Sheet
}

function ColName([int]$n) {
    $s = ''
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }
    return $s
}

function EscX([string]$s) {
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + ' ...[truncated]' }
    return [System.Security.SecurityElement]::Escape($s)
}

function Test-Numeric($v) {
    return ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal] -or $v -is [int16] -or $v -is [single] -or $v -is [byte])
}

function Get-SheetXml($S) {
    $inv = [cultureinfo]::InvariantCulture
    $maxCol = 1; $w = @{}
    foreach ($row in $S.Rows) {
        for ($i = 0; $i -lt $row.Count; $i++) {
            $c = $row[$i]
            if ($i + 1 -gt $maxCol) { $maxCol = $i + 1 }
            if ($c.s -in 1, 2, 6) { continue }
            $len = 0
            foreach ($ln in ("$($c.v)" -split "`n")) { if ($ln.Length -gt $len) { $len = $ln.Length } }
            if (-not $w.ContainsKey($i) -or $w[$i] -lt $len) { $w[$i] = $len }
        }
    }
    $widths = @{}
    for ($i = 0; $i -lt $maxCol; $i++) {
        $l = if ($w.ContainsKey($i)) { $w[$i] } else { 10 }
        $min = if ($i -eq 0) { 30 } else { 14 }
        $max = if ($i -eq 0) { 45 } else { 60 }
        $widths[$i] = [math]::Min($max, [math]::Max($min, $l + 2))
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="13.5"/><cols>')
    for ($i = 0; $i -lt $maxCol; $i++) { [void]$sb.AppendFormat('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), $widths[$i]) }
    [void]$sb.Append('</cols><sheetData>')
    $r = 0
    foreach ($row in $S.Rows) {
        $r++
        if ($row.Count -eq 0) { continue }
        $lines = 1
        for ($i = 0; $i -lt $row.Count; $i++) {
            $c = $row[$i]
            if ($c.s -notin 3, 4, 5) { continue }
            $n = 0
            foreach ($ln in ("$($c.v)" -split "`n")) { $n += [math]::Max(1, [math]::Ceiling($ln.Length / [math]::Max(1, $widths[$i] - 1))) }
            if ($n -gt $lines) { $lines = $n }
        }
        $ht = ''
        if ($lines -gt 1) { $ht = ' ht="{0}" customHeight="1"' -f ([math]::Min(409, 13.5 * $lines)).ToString($inv) }
        elseif ($row[0].s -eq 1) { $ht = ' ht="24" customHeight="1"' }
        elseif ($row[0].s -eq 2) { $ht = ' ht="18" customHeight="1"' }
        [void]$sb.AppendFormat('<row r="{0}"{1}>', $r, $ht)
        for ($i = 0; $i -lt $row.Count; $i++) {
            $c = $row[$i]
            $ref = (ColName ($i + 1)) + $r
            if (Test-Numeric $c.v) {
                [void]$sb.AppendFormat('<c r="{0}" s="{1}"><v>{2}</v></c>', $ref, $c.s, [System.Convert]::ToString($c.v, $inv))
            }
            elseif ($null -eq $c.v -or "$($c.v)" -eq '') {
                [void]$sb.AppendFormat('<c r="{0}" s="{1}"/>', $ref, $c.s)
            }
            else {
                [void]$sb.AppendFormat('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>', $ref, $c.s, (EscX "$($c.v)"))
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($S.Links.Count -gt 0) {
        [void]$sb.Append('<hyperlinks>')
        foreach ($l in $S.Links) { [void]$sb.AppendFormat('<hyperlink ref="{0}" location="{1}" display="{2}"/>', $l.Ref, (EscX $l.Location), (EscX $l.Display)) }
        [void]$sb.Append('</hyperlinks>')
    }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

$StylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
<font><sz val="10"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="10"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><i/><sz val="10"/><color rgb="FF808080"/><name val="Calibri"/><family val="2"/></font>
<font><u/><sz val="10"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@

function Save-Book($Book, [string]$Path) {
    $enc = [System.Text.UTF8Encoding]::new($false)
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $put = {
            param([string]$Name, [string]$Text)
            $e = $zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
            $sw = [System.IO.StreamWriter]::new($e.Open(), $enc)
            $sw.Write($Text); $sw.Dispose()
        }
        $n = $Book.Sheets.Count
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        for ($i = 1; $i -le $n; $i++) { $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }
        $ct += '</Types>'
        & $put '[Content_Types].xml' $ct
        & $put '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets>'
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        for ($i = 1; $i -le $n; $i++) {
            $wb += '<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (EscX $Book.Sheets[$i - 1].Name), $i
            $rels += '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $i
        }
        $wb += '</sheets></workbook>'
        $rels += '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        & $put 'xl/workbook.xml' $wb
        & $put 'xl/_rels/workbook.xml.rels' $rels
        & $put 'xl/styles.xml' $StylesXml
        for ($i = 1; $i -le $n; $i++) { & $put "xl/worksheets/sheet$i.xml" (Get-SheetXml $Book.Sheets[$i - 1]) }
    }
    finally { $zip.Dispose(); $fs.Dispose() }
}

# =====================================================================================
#  REST layer
# =====================================================================================

# Works on Windows PowerShell 5.1 and PowerShell 7+
$WebOpts = @{ UseBasicParsing = $true }

function Set-TlsOptions([bool]$Skip) {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        if ($Skip) { $script:WebOpts['SkipCertificateCheck'] = $true }
        return
    }
    $proto = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    try { $proto = $proto -bor [Net.SecurityProtocolType]'Tls13' } catch { }
    [Net.ServicePointManager]::SecurityProtocol = $proto
    if ($Skip) {
        if (-not ('VcfaTrustAllCerts' -as [type])) {
            Add-Type -TypeDefinition @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class VcfaTrustAllCerts : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
"@
        }
        [Net.ServicePointManager]::CertificatePolicy = New-Object VcfaTrustAllCerts
    }
}

function Read-Json([string]$Text) {
    # Parses into dictionaries / arrays. Avoids the 2 MB limit of ConvertFrom-Json in Windows PowerShell 5.1.
    if ($PSVersionTable.PSVersion.Major -ge 6) { return (ConvertFrom-Json -InputObject $Text -AsHashtable) }
    Add-Type -AssemblyName System.Web.Extensions
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = [int]::MaxValue
    $ser.RecursionLimit = 1000
    return , $ser.DeserializeObject($Text)
}

$script:Conn = @{ Base = ''; Headers = @{}; Skip = $false; IaasApiVersion = $null; CloudApiVersion = '40.0' }
$script:Reconnect = $false
$script:AuthInfo = ''
$script:ApiTokenPlain = $null
$script:Warnings = [System.Collections.Generic.List[string]]::new()
$script:TimeoutSec = 300
$script:MaxCols = 10
$script:Raw = [ordered]@{}
$script:Log = [System.Collections.Generic.List[object]]::new()
$script:FailedKeys = [System.Collections.Generic.List[string]]::new()

function Get-StatusCode($err) {
    try { return [int]$err.Exception.Response.StatusCode } catch { return 0 }
}

function Get-ErrMsg($err) {
    $code = Get-StatusCode $err
    $msg = $err.Exception.Message
    if ($err.ErrorDetails -and $err.ErrorDetails.Message) { $msg = $err.ErrorDetails.Message }
    $msg = ($msg -replace '\s+', ' ')
    if ($msg.Length -gt 250) { $msg = $msg.Substring(0, 250) + '...' }
    if ($code) { return "HTTP $code - $msg" } else { return $msg }
}

function Invoke-Vcfa([string]$Path) {
    # GET with retry on throttling / transient errors and one silent re-login if the token expires.
    $uri = $script:Conn.Base + $Path
    if ($Path.StartsWith('/iaas/') -and $Path -notmatch 'apiVersion=' -and $script:Conn.IaasApiVersion) {
        $uri += $(if ($uri.Contains('?')) { '&' } else { '?' }) + "apiVersion=$($script:Conn.IaasApiVersion)"
    }
    $attempt = 0; $reauth = $false
    while ($true) {
        $attempt++
        $hdr = @{ Authorization = $script:Conn.Headers.Authorization; Accept = 'application/json' }
        if ($Path.StartsWith('/cloudapi/')) { $hdr.Accept = "application/json;version=$($script:Conn.CloudApiVersion)" }
        elseif ($Path.StartsWith('/api/')) { $hdr.Accept = "application/*+json;version=$($script:Conn.CloudApiVersion)" }
        try {
            return (Invoke-RestMethod -Uri $uri -Method Get -Headers $hdr -TimeoutSec $script:TimeoutSec @WebOpts -ErrorAction Stop)
        }
        catch {
            $code = Get-StatusCode $_
            if ($code -eq 401 -and $script:Reconnect -and -not $reauth) {
                $reauth = $true
                Write-Host ' [session expired - logging in again] ' -NoNewline -ForegroundColor DarkYellow
                $script:Conn.Headers = @{ Authorization = "Bearer $(Connect-Vcfa)"; Accept = 'application/json' }
                continue
            }
            if (($code -eq 0 -or $code -eq 429 -or $code -ge 500) -and $code -ne 501 -and $attempt -lt 4) {
                Start-Sleep -Seconds ([int][math]::Pow(2, $attempt))
                continue
            }
            throw
        }
    }
}

function Get-Items($r) {
    if ($null -eq $r) { return , @() }
    if ($r -is [array]) { return , $r }
    foreach ($name in 'content', 'values') {
        $p = $r.PSObject.Properties[$name]
        if ($p) { if ($null -eq $p.Value) { return , @() } else { return , @($p.Value) } }
    }
    return , @($r)
}

function Get-All([string]$Path, [int]$PageSize = 200) {
    # Pages through any VCFA list API. Stops on the server's own paging metadata (not on a short page),
    # so a server that caps the page size below $PageSize does not silently truncate the result.
    $sep = if ($Path.Contains('?')) { '&' } else { '?' }
    $all = [System.Collections.Generic.List[object]]::new()
    $mode = if ($Path.StartsWith('/iaas/')) { 'iaas' } elseif ($Path.StartsWith('/cloudapi/')) { 'cloud' } else { 'page' }
    if ($mode -eq 'cloud') { $PageSize = [math]::Min($PageSize, 128) }
    $expected = $null
    for ($i = 0; $i -lt 5000; $i++) {
        $q = switch ($mode) {
            'iaas'  { "`$top=$PageSize&`$skip=$($all.Count)" }
            'cloud' { "page=$($i + 1)&pageSize=$PageSize" }
            default { "page=$i&size=$PageSize" }
        }
        try { $r = Invoke-Vcfa "$Path$sep$q" }
        catch {
            if ($i -eq 0 -and (Get-StatusCode $_) -eq 400) {
                foreach ($x in (Get-Items (Invoke-Vcfa $Path))) { $all.Add($x) }
                break
            }
            throw
        }
        $batch = Get-Items $r
        foreach ($x in $batch) { $all.Add($x) }
        if ($r -is [array] -or $batch.Count -eq 0) { break }
        if (($i % 10) -eq 9) { Write-Host '.' -NoNewline }
        $meta = $r.PSObject.Properties
        if ($meta['totalElements'] -and $null -ne $r.totalElements) { $expected = [int]$r.totalElements }
        elseif ($meta['resultTotal'] -and $null -ne $r.resultTotal) { $expected = [int]$r.resultTotal }
        if ($null -ne $expected) { if ($all.Count -ge $expected) { break } else { continue } }
        if ($meta['last'] -and $null -ne $r.last) { if ($r.last -eq $true) { break } else { continue } }
        if ($meta['totalPages'] -and $null -ne $r.totalPages) { if (($i + 1) -ge [int]$r.totalPages) { break } else { continue } }
        if ($meta['pageCount'] -and $null -ne $r.pageCount) { if (($i + 1) -ge [int]$r.pageCount) { break } else { continue } }
        if ($batch.Count -lt $PageSize) { break }
    }
    if ($null -ne $expected -and $all.Count -lt $expected) {
        $script:Warnings.Add("$Path returned $($all.Count) of $expected reported items")
    }
    return , $all.ToArray()
}

function Add-Log($Key, $Endpoint, $Status, $Count, $Message, $Ms) {
    $script:Log.Add([ordered]@{ Section = $Key; Endpoint = $Endpoint; Status = $Status; Count = $Count; Message = $Message; 'Time (ms)' = $Ms })
}

function Collect {
    # -EmptyOn404: the API answers 404 when the object simply is not configured (e.g. email server).
    param([string]$Key, [string]$Endpoint, [scriptblock]$Script, [switch]$Single, [switch]$EmptyOn404)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $script:Warnings.Clear()
    Write-Host ('  {0,-28}' -f $Key) -NoNewline
    try {
        if ($Single) { $res = Invoke-Vcfa $Endpoint }
        elseif ($Script) { $res = @(& $Script | ForEach-Object { $_ }) }
        else { $res = Get-All $Endpoint }
        $script:Raw[$Key] = $res
        $cnt = if ($Single) { if ($null -ne $res) { 1 } else { 0 } } else { @($res).Count }
        if ($script:Warnings.Count) {
            $w = $script:Warnings -join '; '
            Write-Host "PARTIAL ($cnt)  $w" -ForegroundColor Yellow
            Add-Log $Key $Endpoint 'PARTIAL' $cnt $w $sw.ElapsedMilliseconds
        }
        else {
            Write-Host "OK ($cnt)" -ForegroundColor Green
            Add-Log $Key $Endpoint 'OK' $cnt '' $sw.ElapsedMilliseconds
        }
    }
    catch {
        $msg = Get-ErrMsg $_
        if ($EmptyOn404 -and (Get-StatusCode $_) -eq 404) {
            Write-Host 'OK (not configured)' -ForegroundColor Green
            $script:Raw[$Key] = $null
            Add-Log $Key $Endpoint 'OK' 0 "Not configured: $msg" $sw.ElapsedMilliseconds
            return
        }
        Write-Host "FAILED  $msg" -ForegroundColor Yellow
        $script:Raw[$Key] = $null
        $script:FailedKeys.Add($Key)
        Add-Log $Key $Endpoint 'FAILED' 0 $msg $sw.ElapsedMilliseconds
    }
}

function Get-Header($resp, [string]$Name) {
    foreach ($k in $resp.Headers.Keys) {
        if ($k -ieq $Name) { return (@($resp.Headers[$k]) | Select-Object -First 1) }
    }
    return $null
}

function Connect-Vcfa {
    $base = $script:Conn.Base; $skip = $script:Conn.Skip
    if ($UseApiToken) {
        if (-not $script:ApiTokenPlain) {
            $sec = Read-Host -AsSecureString "API token for organization '$Organization'"
            $script:ApiTokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
        }
        $r = Invoke-RestMethod @WebOpts -Method Post -Uri "$base/oauth/tenant/$Organization/token" -Body @{ grant_type = 'refresh_token'; refresh_token = $script:ApiTokenPlain } -ContentType 'application/x-www-form-urlencoded' -Headers @{ Accept = 'application/json' }
        $script:AuthInfo = 'API token (/oauth/tenant)'
        return $r.access_token
    }
    $user = $Credential.UserName
    if ($user -notmatch "@$([regex]::Escape($Organization))$") { $user = "$user@$Organization" }
    $pass = $Credential.GetNetworkCredential().Password
    $basic = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${user}:$pass"))
    $lastErr = $null
    foreach ($ver in '40.0', '41.0', '39.0') {
        try {
            $resp = Invoke-WebRequest @WebOpts -Method Post -Uri "$base/cloudapi/1.0.0/sessions" -Headers @{ Accept = "application/json;version=$ver"; Authorization = "Basic $basic" }
            $tok = Get-Header $resp 'x-vmware-vcloud-access-token'
            if ($tok) { $script:AuthInfo = "cloudapi session (API version $ver) as $user"; $script:Conn.CloudApiVersion = $ver; return $tok }
        }
        catch {
            $lastErr = $_
            if ((Get-StatusCode $_) -eq 401) { break }
        }
    }
    # Fallback: 8.x-style login (kept for older / upgraded environments)
    try {
        $u = $Credential.UserName; $dom = $null
        if ($u -match '^(.+)@([^@]+)$' -and $Matches[2] -ne $Organization) { $u = $Matches[1]; $dom = $Matches[2] }
        $body = @{ username = $u; password = $pass }; if ($dom) { $body.domain = $dom }
        $r = Invoke-RestMethod @WebOpts -Method Post -Uri "$base/csp/gateway/am/api/login?access_token" -Body ($body | ConvertTo-Json) -ContentType 'application/json'
        $r2 = Invoke-RestMethod @WebOpts -Method Post -Uri "$base/iaas/api/login" -Body (@{ refreshToken = $r.refresh_token } | ConvertTo-Json) -ContentType 'application/json'
        if ($r2.token) { $script:AuthInfo = 'IaaS login (8.x style)'; return $r2.token }
    }
    catch { }
    if ($lastErr) { throw "Login failed: $(Get-ErrMsg $lastErr)" }
    throw 'Login failed: no access token returned.'
}


# =====================================================================================
#  All Apps specific helpers
# =====================================================================================

function Get-CloudList([string]$Name) {
    # VCF Automation 9.1 publishes the tenant APIs under /cloudapi/v1, 9.0 under /cloudapi/vcf.
    try { return , (Get-All "/cloudapi/v1/$Name") }
    catch {
        if ((Get-StatusCode $_) -eq 404) { return , (Get-All "/cloudapi/vcf/$Name") }
        throw
    }
}

function Get-CloudOne([string]$Name) {
    try { return (Invoke-Vcfa "/cloudapi/v1/$Name") }
    catch {
        if ((Get-StatusCode $_) -eq 404) { return (Invoke-Vcfa "/cloudapi/vcf/$Name") }
        throw
    }
}

function Get-K8sList([string]$Path) {
    # Kubernetes-style list (items + metadata.continue)
    $all = [System.Collections.Generic.List[object]]::new(); $cont = $null
    for ($i = 0; $i -lt 1000; $i++) {
        $q = 'limit=500'
        if ($cont) { $q += '&continue=' + [uri]::EscapeDataString("$cont") }
        $r = Invoke-Vcfa ('{0}?{1}' -f $Path, $q)
        foreach ($x in (Items (Get-P $r 'items'))) { $all.Add($x) }
        $cont = Get-P $r 'metadata.continue'
        if (-not $cont) { break }
    }
    return , $all.ToArray()
}

function Add-Flat($Out, $V, [string]$P) {
    if ($null -eq $V) { return }
    if ($V -is [string] -or $V -is [ValueType]) { if ($P) { $Out[$P] = $V }; return }
    if ($V -is [System.Collections.IDictionary] -or $V -is [System.Management.Automation.PSCustomObject]) {
        foreach ($m in (Get-Members $V)) { Add-Flat $Out $m.Value $(if ($P) { "$P.$($m.Name)" } else { $m.Name }) }
        return
    }
    if ($V -is [System.Collections.IEnumerable]) {
        $scalar = $true
        foreach ($x in $V) { if (-not ($null -eq $x -or $x -is [string] -or $x -is [ValueType])) { $scalar = $false; break } }
        if ($scalar) { if ($P) { $Out[$P] = (@($V) -join ', ') }; return }
        $i = 0
        foreach ($x in $V) { Add-Flat $Out $x "$P[$i]"; $i++ }
    }
}

function Get-Flat($Obj) {
    # Object -> ordered { 'a.b[0].c' = value } under the raw API names
    $out = [ordered]@{}
    Add-Flat $out $Obj ''
    return $out
}

function Find-Deep {
    # First non-empty value whose property name matches $KeyRegex, searched breadth-first
    # through one or more objects. Used where API shapes differ between versions.
    param($Obj, [string]$KeyRegex, [int]$MaxDepth = 7)
    $queue = [System.Collections.Generic.Queue[object]]::new()
    foreach ($o in @($Obj)) { if ($null -ne $o) { $queue.Enqueue(@{ o = $o; d = 0 }) } }
    while ($queue.Count -gt 0) {
        $e = $queue.Dequeue(); $o = $e.o
        if ($null -eq $o -or $e.d -gt $MaxDepth -or $o -is [string] -or $o -is [ValueType]) { continue }
        if ($o -is [System.Collections.IDictionary] -or $o -is [System.Management.Automation.PSCustomObject]) {
            $mem = Get-Members $o
            foreach ($m in $mem) {
                if ($m.Name -match $KeyRegex -and $null -ne $m.Value -and "$($m.Value)" -ne '') { return $m.Value }
            }
            foreach ($m in $mem) { $queue.Enqueue(@{ o = $m.Value; d = $e.d + 1 }) }
        }
        elseif ($o -is [System.Collections.IEnumerable]) { foreach ($x in $o) { $queue.Enqueue(@{ o = $x; d = $e.d + 1 }) } }
    }
    return $null
}

function Get-FlatMatch($Obj, [string]$Regex) {
    # "key: value" lines for every flattened field whose path matches $Regex
    $f = Get-Flat $Obj
    return (Join-Lines @($f.Keys | Where-Object { $_ -match $Regex } | ForEach-Object { '{0}: {1}' -f $_, (Format-Value $f[$_]) }))
}

function Get-NameList($v) {
    # list of objects or strings -> one name per line
    return (Join-Lines @((Items $v) | ForEach-Object {
                if ($_ -is [string] -or $_ -is [ValueType]) { "$_" } else { Get-P $_ 'name', 'displayName', 'metadata.name', 'id' }
            }))
}

function K8sName($o) { return (Get-P $o 'metadata.name', 'name') }

function K8sStatus($o) {
    $txt = @()
    $ph = Get-P $o 'status.phase', 'status.state', 'status.status'
    if ($ph -and ($ph -is [string] -or $ph -is [ValueType])) { $txt += "$ph" }
    foreach ($c in (Items (Get-P $o 'status.conditions'))) {
        if ("$(Get-P $c 'type')" -match '^(Ready|Realized|Available|Succeeded)$') { $txt += ('{0}: {1}' -f (Get-P $c 'type'), (Get-P $c 'status')) }
    }
    return (Join-Lines $txt)
}

function K8sConditions($o) {
    return (Join-Lines ((Items (Get-P $o 'status.conditions')) | ForEach-Object {
                'Type: {0} | Status: {1} | Reason: {2} | Message: {3}' -f (Get-P $_ 'type'), (Get-P $_ 'status'), (Get-P $_ 'reason'), (Get-P $_ 'message')
            }))
}

function K8sRegion($o) {
    $r = Get-P $o 'spec.regionName', 'spec.region', 'status.regionName', 'status.region', 'metadata.labels.region'
    if (-not $r) { $r = Find-Deep @((Get-P $o 'spec'), (Get-P $o 'status')) '(?i)^region(Name|Ref)?$' }
    if ($r -and -not ($r -is [string] -or $r -is [ValueType])) { $r = Get-P $r 'name', 'id' }
    return $r
}

function Get-Refs([string]$Name, $Objs) {
    # Names of the objects (spec / status) that mention $Name as a value
    if (-not $Name) { return @() }
    $hits = [System.Collections.Generic.List[string]]::new()
    foreach ($o in (Items $Objs)) {
        $n = "$(K8sName $o)"
        if ($n -eq $Name -or $hits.Contains($n)) { continue }
        $flat = Get-Flat ([ordered]@{ spec = (Get-P $o 'spec'); status = (Get-P $o 'status') })
        foreach ($v in $flat.Values) {
            if (@("$v" -split ',\s*') -contains $Name) { $hits.Add($n); break }
        }
    }
    return $hits.ToArray()
}

# Kubernetes resources never read (credentials, auth reviews, events)
$K8sSkip = '(?i)secret|token|credential|kubeconfig|^events$|reviews$|^bindings$'

# =====================================================================================
#  Collection
# =====================================================================================

function Invoke-Collection {
    Write-Host "`nCollecting from $($script:Conn.Base) ..." -ForegroundColor Cyan

    Collect 'about' '/iaas/api/about' -Single
    $script:Conn.IaasApiVersion = Get-P $script:Raw.about 'latestApiVersion'
    if (-not $script:Conn.IaasApiVersion) { $script:Conn.IaasApiVersion = '2021-07-15' }
    Collect 'orgs'                     '/cloudapi/1.0.0/orgs'

    # --- Build & Deploy
    Collect 'catalogSources'           '/catalog/api/admin/sources'
    Collect 'catalogItems'             '/catalog/api/admin/items'
    Collect 'blueprints'               '/blueprint/api/blueprints'
    Collect 'blueprintDetails'         '/blueprint/api/blueprints/{id}' -Script {
        $all = Items $script:Raw.blueprints; $n = 0
        foreach ($b in $all) {
            $n++
            if ($all.Count -gt 20) { Write-Progress -Activity 'Blueprint content' -Status "$n of $($all.Count)" -PercentComplete (100 * $n / $all.Count) }
            try { Invoke-Vcfa "/blueprint/api/blueprints/$($b.id)" }
            catch { $script:Warnings.Add("blueprint $($b.name): $(Get-ErrMsg $_)") }
        }
        Write-Progress -Activity 'Blueprint content' -Completed
    }
    Collect 'propertyGroups'           '/properties/api/property-groups'
    Collect 'customResourceTypes'      '/form-service/api/custom/resource-types'
    Collect 'customResourceActions'    '/form-service/api/custom/resource-actions'
    Collect 'contentLibraries'         '/cloudapi/v1/contentLibraries' -Script { Get-CloudList 'contentLibraries' }
    Collect 'contentLibraryItems'      '/cloudapi/v1/contentLibraryItems' -Script { Get-CloudList 'contentLibraryItems' }
    Collect 'eventTopics'              '/event-broker/api/topics'
    Collect 'eventSubscriptions'       '/event-broker/api/subscriptions'
    Collect 'vroWorkflows'             '/vro/workflows'

    # --- Manage & Govern
    Collect 'projectService'           '/project-service/api/projects'
    Collect 'policies'                 '/policy/api/policies?expandDefinition=true'
    Collect 'regions'                  '/cloudapi/v1/regions' -Script { Get-CloudList 'regions' }
    Collect 'regionZones'              '/cloudapi/v1/regions/{urn}/namespaceCompatibleZones' -Script {
        foreach ($r in (Items $script:Raw.regions)) {
            $z = $null
            try { $z = Get-All "/cloudapi/v1/regions/$($r.id)/namespaceCompatibleZones" }
            catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("region $($r.name): $(Get-ErrMsg $_)") } }
            [pscustomobject]@{ regionId = "$($r.id)"; zones = @($z) }
        }
    }
    Collect 'zones'                    '/cloudapi/v1/zones' -Script { Get-CloudList 'zones' }
    Collect 'namespaceSummaries'       '/cloudapi/v1/namespaceSummaries' -Script { Get-CloudList 'namespaceSummaries' }
    Collect 'namespaceStorageMetrics'  '/cloudapi/v1/namespaceStorageClassMetrics' -Script { Get-CloudList 'namespaceStorageClassMetrics' }
    Collect 'vmClasses'                '/cloudapi/v1/virtualMachineClasses' -Script { Get-CloudList 'virtualMachineClasses' }
    Collect 'storageClasses'           '/cloudapi/v1/storageClasses' -Script { Get-CloudList 'storageClasses' }
    if (-not $SkipKubernetes) {
        Collect 'k8s'                  '/cci/kubernetes/apis/{group}/{version}/{resource}' -Script {
            $groups = Invoke-Vcfa '/cci/kubernetes/apis'
            $projNames = @((Items $script:Raw.projectService) | ForEach-Object { "$($_.name)"; "$($_.name)".ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
            foreach ($g in (Items (Get-P $groups 'groups'))) {
                $gname = "$(Get-P $g 'name')"
                if ($gname -notmatch '(?i)vmware') { continue }
                $gv = Get-P $g 'preferredVersion.groupVersion'
                if (-not $gv) { $gv = Get-P ((Items (Get-P $g 'versions')) | Select-Object -First 1) 'groupVersion' }
                if (-not $gv) { continue }
                try { $rl = Invoke-Vcfa "/cci/kubernetes/apis/$gv" }
                catch { $script:Warnings.Add("${gv}: $(Get-ErrMsg $_)"); continue }
                foreach ($res in (Items (Get-P $rl 'resources'))) {
                    $rn = "$(Get-P $res 'name')"
                    if (-not $rn -or $rn.Contains('/') -or $rn -match $K8sSkip) { continue }
                    if (@(Get-P $res 'verbs') -notcontains 'list') { continue }
                    $kind = "$(Get-P $res 'kind')"
                    $nsd = [bool](Get-P $res 'namespaced')
                    $items = [System.Collections.Generic.List[object]]::new()
                    $errs = [System.Collections.Generic.List[string]]::new()
                    $done = $false
                    # cluster-wide list first; namespaced resources fall back to one call per project namespace
                    try { foreach ($x in (Get-K8sList "/cci/kubernetes/apis/$gv/$rn")) { $items.Add($x) }; $done = $true }
                    catch { if (-not $nsd) { $errs.Add((Get-ErrMsg $_)) } }
                    if (-not $done -and $nsd) {
                        foreach ($p in $projNames) {
                            try { foreach ($x in (Get-K8sList "/cci/kubernetes/apis/$gv/namespaces/$p/$rn")) { $items.Add($x) } }
                            catch { if ((Get-StatusCode $_) -ne 404) { $errs.Add("namespace ${p}: $(Get-ErrMsg $_)") } }
                        }
                    }
                    if ($errs.Count) { $script:Warnings.Add("$gv/${rn}: $($errs -join '; ')") }
                    if ($items.Count) { Write-Host ("`n      {0,-45} {1}" -f "$kind ($gv)", $items.Count) -NoNewline -ForegroundColor DarkGray }
                    [pscustomobject]@{ groupVersion = $gv; resource = $rn; kind = $kind; namespaced = $nsd; count = $items.Count; errors = @($errs); items = @($items) }
                }
            }
            Write-Host ''
            Write-Host ('  {0,-28}' -f '') -NoNewline
        }
    }

    # --- Administer
    Collect 'orgUsers'                 '/cloudapi/1.0.0/users'
    Collect 'orgGroups'                '/cloudapi/1.0.0/groups'
    Collect 'orgRoles'                 '/cloudapi/1.0.0/roles'
    Collect 'serviceAccounts'          '/cloudapi/1.0.0/serviceAccounts'
    Collect 'integrations'             '/iaas/api/integrations'
    Collect 'orgSettings'              '/cloudapi/v1/orgSettings' -Script { Get-CloudOne 'orgSettings' } -EmptyOn404
    Collect 'identityProviders'        '/api/admin/org/{id}/settings/(oauth|ldap|federation) + /cloudapi/v1/orgSettings/ldap' -Script {
        $orgs = Items $script:Raw.orgs
        if ($orgs.Count -eq 0) { $orgs = Get-All '/cloudapi/1.0.0/orgs' }
        $o = $orgs | Where-Object { "$($_.name)" -eq $Organization } | Select-Object -First 1
        if (-not $o) { $o = $orgs | Select-Object -First 1 }
        $uuid = if ($o) { ("$($o.id)" -split ':')[-1] } else { $null }
        foreach ($kind in 'oauth', 'ldap', 'federation') {
            $res = $null; $err = $null
            if ($uuid) {
                try { $res = Invoke-Vcfa "/api/admin/org/$uuid/settings/$kind" }
                catch { if ((Get-StatusCode $_) -ne 404) { $err = Get-ErrMsg $_ } }
            }
            else { $err = 'Organization not visible through /cloudapi/1.0.0/orgs' }
            [pscustomobject]@{ kind = $kind; settings = $res; error = $err }
        }
        $res = $null; $err = $null
        try { $res = Get-CloudOne 'orgSettings/ldap' }
        catch { if ((Get-StatusCode $_) -ne 404) { $err = Get-ErrMsg $_ } }
        [pscustomobject]@{ kind = 'ldapV1'; settings = $res; error = $err }
    }
    Collect 'secrets'                  '/platform/api/secrets'
    Collect 'configProperties'         '/iaas/api/configuration-properties'
    Collect 'emailConfig'              '/notification/api/email-config' -Single -EmptyOn404
    Collect 'notificationScenarios'    '/notification/api/scenario-configs'
}

# =====================================================================================
#  Workbook rendering (works only from the saved JSON, so -FromJson gives identical output)
# =====================================================================================

$PolicyTypes = [ordered]@{
    'com.vmware.policy.approval'            = 'Approval Policy'
    'com.vmware.policy.deployment.action'   = 'Day 2 Actions Policy'
    'com.vmware.policy.deployment.lease'    = 'Lease Policy'
    'com.vmware.policy.iaas.resource'       = 'IaaS Resource Policy'
    'com.vmware.policy.catalog.entitlement' = 'Content Sharing Policy'
    'com.vmware.policy.resource.quota'      = 'Resource Quota Policy'
    'com.vmware.policy.deployment.limit'    = 'Deployment Limit Policy'
}

$NotFromApi = 'Not available from the VCF Automation API - fill in manually'

function Build-Workbook {
    param($Data, [string]$Path)
    $Src = $Data.raw
    $script:FailedKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($k in (Items $Data.failed)) { $script:FailedKeys.Add("$k") }
    $org = $Data.meta.organization

    $Bk = New-Book
    $shIdx = Add-Sheet $Bk 'Index'
    $shVer = Add-Sheet $Bk 'Version History'
    $shCat = Add-Sheet $Bk 'Catalog'
    $shBp  = Add-Sheet $Bk 'Blueprint Design'
    $shImg = Add-Sheet $Bk 'VM Images'
    $shCl  = Add-Sheet $Bk 'Content Libraries'
    $shEvt = Add-Sheet $Bk 'Event Subscriptions'
    $shPrj = Add-Sheet $Bk 'Projects'
    $shPol = Add-Sheet $Bk 'Policies'
    $shReg = Add-Sheet $Bk 'Regions'
    $shNs  = Add-Sheet $Bk 'Namespaces'
    $shNsc = Add-Sheet $Bk 'Namespace Class'
    $shVpc = Add-Sheet $Bk 'Virtual Private Clouds'
    $shTgw = Add-Sheet $Bk 'Transit Gateways'
    $shIp  = Add-Sheet $Bk 'IP Management'
    $shAcc = Add-Sheet $Bk 'Access Control'
    $shExt = Add-Sheet $Bk 'External Orchestrators'
    $shInt = Add-Sheet $Bk 'Integrations'
    $shIdp = Add-Sheet $Bk 'Identity Providers'
    $shSec = Add-Sheet $Bk 'Secrets'
    $shCfg = Add-Sheet $Bk 'Configuration'
    $shK8s = Add-Sheet $Bk 'Other CCI Resources'
    $shLog = Add-Sheet $Bk 'Collection Log'

    # ---------------- lookups ----------------
    $k8sRes = Items $Src.k8s
    $usedKinds = [System.Collections.Generic.List[string]]::new()
    function KindItems([string]$Regex, [string]$NotRegex) {
        $out = [System.Collections.Generic.List[object]]::new()
        foreach ($r in $k8sRes) {
            $k = "$(Get-P $r 'kind')"
            if ($k -notmatch $Regex) { continue }
            if ($NotRegex -and $k -match $NotRegex) { continue }
            if (-not $usedKinds.Contains($k)) { $usedKinds.Add($k) }
            foreach ($x in (Items (Get-P $r 'items'))) { $out.Add($x) }
        }
        return , $out.ToArray()
    }

    $projects = Items $Src.projectService
    $projById = @{}; foreach ($p in $projects) { $projById["$($p.id)"] = $p }
    function PName($id) { if ($id -and $projById.ContainsKey("$id")) { return $projById["$id"].name } elseif ($id) { return "$id" } else { return '' } }
    function Scope($o) { $prj = Get-P $o 'projectId'; if ($prj) { return "Project: $(PName $prj)" } else { return 'Organization' } }

    $regions = Items $Src.regions
    $regionById = @{}; foreach ($r in $regions) { $regionById["$($r.id)"] = $r }
    function RegName($v) {
        if ($null -eq $v) { return '' }
        if ($v -is [string]) { if ($regionById.ContainsKey($v)) { return $regionById[$v].name } else { return $v } }
        $n = Get-P $v 'name'; if ($n) { return $n }
        $id = Get-P $v 'id'; if ($id -and $regionById.ContainsKey("$id")) { return $regionById["$id"].name }
        return "$id"
    }

    # namespaces: REST summaries merged with CCI SupervisorNamespace objects (by name)
    $nsK8s = KindItems '(?i)^SupervisorNamespace$'
    $nsRecs = [ordered]@{}
    foreach ($n in (Items $Src.namespaceSummaries)) { $nm = "$(Get-P $n 'name')"; if ($nm) { $nsRecs[$nm] = @{ S = $n; K = $null } } }
    foreach ($k in $nsK8s) { $nm = "$(K8sName $k)"; if (-not $nm) { continue }; if (-not $nsRecs.Contains($nm)) { $nsRecs[$nm] = @{ S = $null; K = $k } } else { $nsRecs[$nm].K = $k } }
    $nsRows = foreach ($nm in $nsRecs.Keys) {
        $s = $nsRecs[$nm].S; $k = $nsRecs[$nm].K
        $proj = Get-P $s 'project.name', 'projectName', 'projectRef.name'; if (-not $proj) { $proj = Get-P $k 'metadata.namespace' }
        $status = Get-P $s 'status', 'phase', 'state'; if ($status -and -not ($status -is [string])) { $status = Format-Value $status }
        if (-not $status) { $status = K8sStatus $k }
        $cls = Get-P $s 'className', 'namespaceClassName', 'supervisorNamespaceClass.name', 'namespaceClass.name', 'classRef.name'
        if (-not $cls) { $cls = Get-P $k 'spec.className', 'spec.classRef.name' }
        if (-not $cls) { $cls = Find-Deep @($s, (Get-P $k 'spec')) '(?i)class(Name|Ref)?$' }
        $reg = Get-P $s 'region.name', 'regionName', 'regionRef.name'; if (-not $reg) { $reg = K8sRegion $k }
        $vpc = Get-P $s 'vpc.name', 'vpcName', 'vpcRef.name'; if (-not $vpc) { $vpc = Get-P $k 'spec.vpcName', 'spec.vpc.name', 'spec.vpcRef.name' }
        if (-not $vpc) { $vpc = Find-Deep @($s, (Get-P $k 'spec')) '(?i)^vpc(Name|Ref)?$' }
        $zon = Get-P $s 'zones', 'zoneNames'; if (-not $zon) { $zon = Get-P $k 'spec.zones', 'status.zones', 'spec.zoneNames' }
        [ordered]@{
            'Name'            = $nm
            'Project'         = $proj
            'Description'     = Get-P @($s, (Get-P $k 'spec')) 'description'
            'Status'          = $status
            'Namespace Class' = if ($cls -and -not ($cls -is [string])) { Get-P $cls 'name' } else { $cls }
            'Region'          = RegName $reg
            'VPC'             = if ($vpc -and -not ($vpc -is [string])) { Get-P $vpc 'name' } else { $vpc }
            'Zones'           = Get-NameList $zon
        }
    }
    $nsRows = @($nsRows)
    function NsWhere([string]$Field, [string]$Value) { return (Join-Lines @($nsRows | Where-Object { "$($_[$Field])" -eq $Value } | ForEach-Object { $_['Name'] })) }

    # =========================================================================
    #  Catalog
    # =========================================================================
    Add-Title $shCat 'Catalog' "Organization: $org"
    $rows = foreach ($i in (Items $Src.catalogItems)) {
        $pids = @(Get-P $i 'projectIds')
        $icon = Get-P $i 'iconId'
        [ordered]@{
            'Name'                           = $i.name
            'Project'                        = Join-Lines ($pids | Where-Object { $_ } | ForEach-Object { PName $_ })
            'Catalog Type'                   = Get-P $i 'type.name', 'type.id'
            'Icon'                           = if ($icon) { "custom ($icon)" } else { 'default' }
            'scope (organization\provider)'  = Get-P $i 'scope', 'sharingScope', 'visibility'
        }
    }
    Add-Vertical $Bk $shCat 'Catalog Items' $rows 'catalogItems'

    # =========================================================================
    #  Blueprint Design
    # =========================================================================
    Add-Title $shBp 'Blueprint Design' "Organization: $org"
    $bpDet = @{}; foreach ($b in (Items $Src.blueprintDetails)) { $bpDet["$($b.id)"] = $b }
    $bpContent = @{}
    $rows = foreach ($b in (Items $Src.blueprints)) {
        $d = $bpDet["$($b.id)"]; $both = @($d, $b)
        $content = Get-P $both 'content'; $bpContent["$($b.name)"] = "$content"
        $csId = Get-P $both 'contentSourceId'
        $ro = Get-P $both 'readOnly'
        $rel = Get-P $both 'totalReleasedVersions'
        [ordered]@{
            'Name'                       = $b.name
            'Source Control'             = if ($csId) { Join-Lines @('Yes', (Get-P $both 'contentSourceType'), (Get-P $both 'contentSourcePath')) } else { 'No' }
            # templates synced from source control are read-only in the designer
            'Read-only'                  = if ($null -ne $ro) { $ro } elseif ($csId) { 'Yes' } else { 'No' }
            'Project'                    = Get-P $both 'projectName'
            'Shared with other projects' = Get-P $both 'requestScopeOrg'
            'Published'                  = if ($null -ne $rel) { [bool]([int]$rel -gt 0) } else { Get-P $both 'status' }
            'YAML Code'                  = $content
        }
    }
    Add-Vertical $Bk $shBp 'Blueprints' $rows 'blueprints', 'blueprintDetails'

    function UsedIn([string]$Name) {
        if (-not $Name) { return '' }
        $pat = 'propgroup\.' + [regex]::Escape($Name) + '\b'
        return (Join-Lines @($bpContent.Keys | Where-Object { $bpContent[$_] -match $pat }))
    }
    foreach ($pgType in @(@{ T = 'INPUT'; Title = 'Property Groups - Input Values'; Val = 'Default Value'; Key = 'default' }, @{ T = 'CONSTANT'; Title = 'Property Groups - Constant Values'; Val = 'Default Value'; Key = 'const' })) {
        $groups = @((Items $Src.propertyGroups) | Where-Object { "$(Get-P $_ 'type')" -eq $pgType.T })
        $rows = foreach ($g in $groups) {
            [ordered]@{
                'Name' = $g.name; 'Display Name' = Get-P $g 'displayName'; 'Description' = Get-P $g 'description'
                'Scope' = if (Get-P $g 'projectId') { 'Project' } else { 'Organization' }; 'Project' = PName (Get-P $g 'projectId')
                'Used in Blueprints' = UsedIn $g.name
            }
        }
        Add-Vertical $Bk $shBp $pgType.Title $rows 'propertyGroups'
        $rows = foreach ($g in $groups) {
            $i = 0
            foreach ($m in (Get-Members (Get-P $g 'properties'))) {
                $i++
                [ordered]@{
                    'Property Group' = $g.name; 'Property' = "Property #$i"; 'Name' = $m.Name
                    'Display Name' = Get-P $m.Value 'title'; 'Type' = Get-P $m.Value 'type'
                    'Default Value' = Get-P $m.Value $pgType.Key, 'default', 'const'
                }
            }
        }
        Add-Table $Bk $shBp "$($pgType.Title) - Properties" $rows 'propertyGroups'
    }

    $rows = foreach ($c in (Items $Src.customResourceTypes)) {
        [ordered]@{
            'Name' = Get-P $c 'displayName', 'name'; 'Description' = Get-P $c 'description'; 'Resource Type' = Get-P $c 'resourceType'
            'Available in Blueprints' = if ("$(Get-P $c 'status')" -eq 'RELEASED') { 'Yes' } elseif (Get-P $c 'status') { "No ($(Get-P $c 'status'))" } else { '' }
        }
    }
    Add-Vertical $Bk $shBp 'Custom Resources' $rows 'customResourceTypes'

    $rows = foreach ($a in (Items $Src.customResourceActions)) {
        [ordered]@{
            'Name' = Get-P $a 'displayName', 'name'; 'Description' = Get-P $a 'description'
            'Acts on' = Get-P $a 'resourceTypeDisplayName', 'resourceType'
            'Active' = if ("$(Get-P $a 'status')" -eq 'RELEASED') { 'Yes' } elseif (Get-P $a 'status') { "No ($(Get-P $a 'status'))" } else { '' }
        }
    }
    Add-Vertical $Bk $shBp 'Resource Actions' $rows 'customResourceActions'

    # =========================================================================
    #  VM Images / Content Libraries
    # =========================================================================
    Add-Title $shImg 'VM Images' "Organization: $org  -  content library items"
    $rows = foreach ($i in (Items $Src.contentLibraryItems)) {
        [ordered]@{
            'Name' = $i.name; 'Version' = Get-P $i 'versionNumber', 'version'; 'Status' = Get-P $i 'status'
            'Image Identifier' = Get-P $i 'imageIdentifier', 'imageId'; 'Content Library' = Get-P $i 'contentLibrary.name', 'library.name'
            'Type' = Get-P $i 'itemType', 'type'
        }
    }
    Add-Vertical $Bk $shImg 'VM Images' $rows 'contentLibraryItems'

    Add-Title $shCl 'Content Libraries' "Organization: $org"
    $rows = foreach ($l in (Items $Src.contentLibraries)) {
        [ordered]@{
            'Name' = $l.name; 'Status' = Get-P $l 'status'; 'Type' = Get-P $l 'libraryType'; 'Subscribed' = Get-P $l 'isSubscribed'
            'Version' = Get-P $l 'versionNumber'; 'Creation Date' = Get-P $l 'creationDate'
        }
    }
    Add-Vertical $Bk $shCl 'Content Libraries' $rows 'contentLibraries'

    # =========================================================================
    #  Event Subscriptions
    # =========================================================================
    Add-Title $shEvt 'Event Subscriptions' "Organization: $org"
    $topicName = @{}; foreach ($t in (Items $Src.eventTopics)) { $topicName["$(Get-P $t 'id')"] = Get-P $t 'name' }
    $wfName = @{}; foreach ($w in (Items $Src.vroWorkflows)) { $wfName["$(Get-P $w 'id')"] = Get-P $w 'name' }
    function Runnable($type, $id) {
        if (-not $id) { return '' }
        $nm = if ($wfName.ContainsKey("$id")) { $wfName["$id"] } else { "$id" }
        if ($type) { return "$nm ($type)" } else { return $nm }
    }
    $rows = foreach ($s in (Items $Src.eventSubscriptions)) {
        $tid = Get-P $s 'eventTopicId'
        $pids = @(Get-P $s 'constraints.projectId', 'constraints.projectIds', 'projectIds')
        [ordered]@{
            'Name' = $s.name; 'Description' = Get-P $s 'description'
            'Status' = if (Get-P $s 'disabled') { 'Disabled' } else { 'Enabled' }
            'Event Topic' = if ($topicName.ContainsKey("$tid")) { "$($topicName["$tid"]) ($tid)" } else { $tid }
            'Condition' = Get-P $s 'constraints.condition', 'criteria'
            'Workflow' = Runnable (Get-P $s 'runnableType') (Get-P $s 'runnableId')
            'Blocking' = Get-P $s 'blocking'; 'Timeout' = Get-P $s 'timeout'; 'Priority' = Get-P $s 'priority'
            'Recover Workflow' = Runnable (Get-P $s 'recoverRunnableType') (Get-P $s 'recoverRunnableId')
            'Project Scope' = if (@($pids | Where-Object { $_ }).Count) { Join-Lines ($pids | Where-Object { $_ } | ForEach-Object { PName $_ }) } else { 'Any project' }
        }
    }
    Add-Vertical $Bk $shEvt 'Subscriptions' $rows 'eventSubscriptions'

    # =========================================================================
    #  Projects
    # =========================================================================
    Add-Title $shPrj 'Projects' "Organization: $org"
    $rows = foreach ($p in $projects) {
        $n = 0
        foreach ($role in 'administrators', 'members', 'viewers', 'supervisors', 'principals') { $n += (Items (Get-P $p $role)).Count }
        [ordered]@{
            'Name' = $p.name
            'Users & Groups' = $n
            'Namespaces' = @($nsRows | Where-Object { "$($_['Project'])" -eq "$($p.name)" }).Count
            'Description' = Get-P $p 'description'
            'Data Service Admin Password' = 'not stored in the sheet'
        }
    }
    Add-Vertical $Bk $shPrj 'Projects' $rows 'projectService'

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $projects) {
        foreach ($role in 'administrators', 'members', 'viewers', 'supervisors') {
            foreach ($u in (Items (Get-P $p $role))) {
                $rows.Add([ordered]@{ 'Project' = $p.name; 'Role' = $role.TrimEnd('s'); 'Type' = Get-P $u 'type'; 'Name' = Get-P $u 'email', 'name', 'id' })
            }
        }
        foreach ($u in (Items (Get-P $p 'principals'))) {
            $rows.Add([ordered]@{ 'Project' = $p.name; 'Role' = Get-P $u 'role', 'roles'; 'Type' = Get-P $u 'type'; 'Name' = Get-P $u 'email', 'name', 'id' })
        }
    }
    Add-Table $Bk $shPrj 'Projects - Users & Groups' $rows 'projectService'
    Add-Vertical $Bk $shPrj 'Namespaces' $nsRows 'namespaceSummaries', 'k8s'

    # =========================================================================
    #  Policies
    # =========================================================================
    Add-Title $shPol 'Policies' "Organization: $org"
    $pols = Items $Src.policies
    function Criteria($p) { $c = Get-P $p 'criteria'; if ($c) { return ($c | ConvertTo-Json -Depth 20 -Compress) } else { return '' } }
    $typesSeen = @($PolicyTypes.Keys) + @($pols | ForEach-Object { Get-P $_ 'typeId' } | Where-Object { $_ -and $PolicyTypes.Keys -notcontains $_ } | Select-Object -Unique)
    foreach ($tid in $typesSeen) {
        $title = if ($PolicyTypes.Contains("$tid")) { $PolicyTypes["$tid"] } else { "Policy - $tid" }
        $rows = foreach ($p in ($pols | Where-Object { (Get-P $_ 'typeId') -eq $tid })) {
            $def = Get-P $p 'definition'
            $r = [ordered]@{ 'Name' = $p.name; 'Description' = Get-P $p 'description'; 'Scope' = Scope $p; 'Criteria' = Criteria $p }
            switch ("$tid") {
                'com.vmware.policy.approval' {
                    $r['Approval level'] = Get-P $def 'level'
                    $r['Approval type'] = Get-P $def 'approvalType'
                    $r['Approvers'] = Get-P $def 'approvers'
                    $r['Approver Mode'] = Get-P $def 'approvalMode', 'approverMode'
                    $r['Auto expiry decision'] = Get-P $def 'autoApprovalDecision'
                    $r['Auto expiry trigger'] = Get-P $def 'autoApprovalExpiry'
                    $r['Actions'] = Get-P $def 'actions'
                }
                'com.vmware.policy.deployment.action' {
                    $allowed = Items (Get-P $def 'allowedActions')
                    $r['Enforcement type'] = Get-P $p 'enforcementType'
                    $r['Entitlement type'] = Get-P $def 'entitlementType'
                    $r['Role'] = Join-Lines @($allowed | ForEach-Object { @(Get-P $_ 'authorities') } | Where-Object { $_ } | Select-Object -Unique)
                    $r['Actions'] = Join-Lines @($allowed | ForEach-Object { @(Get-P $_ 'actions') } | Where-Object { $_ } | Select-Object -Unique)
                }
                'com.vmware.policy.deployment.lease' {
                    $r['Enforcement type'] = Get-P $p 'enforcementType'
                    $r['Maximum lease (days)'] = Get-P $def 'leaseTermMax'
                    $r['Maximum total lease (days)'] = Get-P $def 'leaseTotalTermMax'
                    $r['Grace period (days)'] = Get-P $def 'leaseGrace'
                }
                'com.vmware.policy.iaas.resource' {
                    $r['Enforcement type'] = Get-P $p 'enforcementType'
                    $tpl = Get-P $def 'template', 'policyTemplate', 'resource'
                    $r['Policy Template (YAML code)'] = if ($tpl) { $tpl } else { Format-Value $def }
                }
                default {
                    $r['Enforcement type'] = Get-P $p 'enforcementType'
                    $r['Definition'] = Format-Value $def
                }
            }
            $r
        }
        Add-Vertical $Bk $shPol $title $rows 'policies'
    }

    # =========================================================================
    #  Regions
    # =========================================================================
    Add-Title $shReg 'Regions' "Organization: $org"
    $zones = Items $Src.zones
    $rzByRegion = @{}; foreach ($x in (Items $Src.regionZones)) { $rzByRegion["$(Get-P $x 'regionId')"] = Items (Get-P $x 'zones') }
    function ZonesOf($r) {
        $z = @($zones | Where-Object { "$(Get-P $_ 'region.id', 'regionId', 'regionRef.id')" -eq "$($r.id)" -or "$(Get-P $_ 'region.name', 'regionName')" -eq "$($r.name)" } | ForEach-Object { $_.name })
        if ($rzByRegion.ContainsKey("$($r.id)")) { $z += @($rzByRegion["$($r.id)"] | ForEach-Object { Get-P $_ 'name', 'id' }) }
        return (Join-Lines @($z | Where-Object { $_ } | Select-Object -Unique))
    }
    $regObjs = @($regions)
    if ($regObjs.Count -eq 0) { $regObjs = @((KindItems '(?i)^Region$') | ForEach-Object { [pscustomobject]@{ id = K8sName $_; name = K8sName $_; spec = Get-P $_ 'spec'; status = Get-P $_ 'status' } }) }
    $rows = foreach ($r in $regObjs) {
        [ordered]@{
            'Name' = $r.name; 'Description' = Get-P $r 'description'
            'Zones' = ZonesOf $r
            'Namespaces' = NsWhere 'Region' "$($r.name)"
            'Utilization - CPU' = Get-FlatMatch $r '(?i)cpu'
            'Utilization - Memory' = Get-FlatMatch $r '(?i)mem'
            'Utilization - Storage' = Get-FlatMatch $r '(?i)storage'
        }
    }
    Add-Vertical $Bk $shReg 'Regions' $rows 'regions'

    $rows = foreach ($z in $zones) {
        [ordered]@{ 'Name' = $z.name; 'Region' = RegName (Get-P $z 'region', 'regionRef', 'regionId'); 'Description' = Get-P $z 'description'; 'CPU' = Get-FlatMatch $z '(?i)cpu'; 'Memory' = Get-FlatMatch $z '(?i)mem' }
    }
    Add-Table $Bk $shReg 'Zones' $rows 'zones'

    # =========================================================================
    #  Namespaces
    # =========================================================================
    Add-Title $shNs 'Namespaces' "Organization: $org"
    Add-Vertical $Bk $shNs 'Namespaces' $nsRows 'namespaceSummaries', 'k8s'
    $rows = foreach ($m in (Items $Src.namespaceStorageMetrics)) { Get-Flat $m }
    Add-Table $Bk $shNs 'Namespace Storage Class Usage' $rows 'namespaceStorageMetrics'

    # =========================================================================
    #  Namespace Class
    # =========================================================================
    Add-Title $shNsc 'Namespace Class' "Organization: $org"
    $classes = KindItems '(?i)NamespaceClass$'
    $classCfg = @{}; foreach ($c in (KindItems '(?i)NamespaceClassConfig$')) { $classCfg["$(K8sName $c)"] = $c }
    $rows = foreach ($c in $classes) {
        $nm = "$(K8sName $c)"; $both = @((Get-P $c 'spec'), (Get-P $classCfg[$nm] 'spec'), (Get-P $c 'status'))
        [ordered]@{
            'Name'                = $nm
            'Description'         = Get-P $both 'description'
            'CPU Limit'           = Find-Deep $both '(?i)^cpu.?limit'
            'CPU Reservation'     = Find-Deep $both '(?i)^cpu.?reserv'
            'Memory limit'        = Find-Deep $both '(?i)^mem(ory)?.?limit'
            'Memory reservation'  = Find-Deep $both '(?i)^mem(ory)?.?reserv'
            'VM Classes'          = Get-NameList (Find-Deep $both '(?i)^(vm|virtualMachine)Class(es|Names|Refs)?$')
            'Storage Classes'     = Format-Value (Find-Deep $both '(?i)^storageClass(es|Names|Refs)?$')
            'Content Libraries'   = Get-NameList (Find-Deep $both '(?i)^contentLibrar(y|ies)(Names|Refs)?$')
            'Spec (all fields)'   = Get-P $c 'spec'
        }
    }
    Add-Vertical $Bk $shNsc 'Namespace Classes' $rows 'k8s'

    $rows = foreach ($v in (Items $Src.vmClasses)) {
        [ordered]@{ 'Name' = $v.name; 'Regions' = Get-NameList (Get-P $v 'region', 'regions', 'regionRef'); 'CPU' = Get-FlatMatch $v '(?i)cpu'; 'Memory' = Get-FlatMatch $v '(?i)mem' }
    }
    Add-Table $Bk $shNsc 'VM Classes' $rows 'vmClasses'
    $rows = foreach ($s in (Items $Src.storageClasses)) {
        [ordered]@{ 'Name' = $s.name; 'Regions' = Get-NameList (Get-P $s 'region', 'regions', 'regionRef'); 'Limits' = Get-FlatMatch $s '(?i)limit|capacity|MiB|GiB' }
    }
    Add-Table $Bk $shNsc 'Storage Classes' $rows 'storageClasses'
    $rows = foreach ($l in (Items $Src.contentLibraries)) {
        [ordered]@{ 'Name' = $l.name; 'Regions' = Join-Lines @((Items (Get-P $l 'storageClasses')) | ForEach-Object { Get-P $_ 'region.name', 'regionName' } | Where-Object { $_ } | Select-Object -Unique) }
    }
    Add-Table $Bk $shNsc 'Content Libraries' $rows 'contentLibraries'

    # =========================================================================
    #  Networking (CCI: VPCs, connectivity profiles, transit gateways, IP blocks, IP quotas)
    # =========================================================================
    $vpcs     = KindItems '(?i)^(vpc|virtualprivatecloud)$'
    $profiles = KindItems '(?i)connectivityprofile'
    $tgws     = KindItems '(?i)transitgateway$'
    $blocks   = KindItems '(?i)^ip(address)?blocks?$'
    $quotas   = KindItems '(?i)^ip.*quota'
    function Pick($o, [string]$Regex) {
        $v = Find-Deep @((Get-P $o 'spec'), (Get-P $o 'status')) $Regex
        if ($v -is [System.Collections.IDictionary] -or $v -is [System.Management.Automation.PSCustomObject]) { $n = Get-P $v 'name'; if ($n) { return $n } }
        return $v
    }
    $profByName = @{}; foreach ($p in $profiles) { $profByName["$(K8sName $p)"] = $p }
    function VpcTgw($v) {
        $t = Pick $v '(?i)^transitGateway(Name|Ref)?$'
        if (-not $t) { $cp = "$(Pick $v '(?i)connectivityProfile(Name|Ref)?$')"; if ($cp -and $profByName.ContainsKey($cp)) { $t = Pick $profByName[$cp] '(?i)^transitGateway(Name|Ref)?$' } }
        return $t
    }
    function VpcsVia([string]$Name) {
        # VPCs naming the object directly, or through a connectivity profile that names it
        $out = @(Get-Refs $Name $vpcs)
        foreach ($pn in @(Get-Refs $Name $profiles)) { $out += @(Get-Refs $pn $vpcs) }
        return (Join-Lines @($out | Where-Object { $_ } | Select-Object -Unique))
    }

    Add-Title $shVpc 'Virtual Private Clouds' "Organization: $org"
    $rows = foreach ($v in $vpcs) {
        $nm = "$(K8sName $v)"
        [ordered]@{
            'Name' = $nm; 'Region' = K8sRegion $v
            'Transit Gateway' = VpcTgw $v
            'Private VPC IPv4 CIDRs' = Pick $v '(?i)^private(IPs|Ipv4Cidrs|IPv4CIDRs|Cidrs|IpBlocks?)$'
            'Default SNAT IP' = Pick $v '(?i)snat'
            'Namespaces' = NsWhere 'VPC' $nm
            'IP quota status' = Pick $v '(?i)quota'
            'Status' = K8sStatus $v
        }
    }
    Add-Vertical $Bk $shVpc 'VPCs' $rows 'k8s'

    $rows = foreach ($p in $profiles) {
        $nm = "$(K8sName $p)"
        [ordered]@{
            'Name' = $nm; 'Region' = K8sRegion $p
            'Transit Gateway' = Pick $p '(?i)^transitGateway(Name|Ref)?$'
            'External IPv4 blocks' = Pick $p '(?i)^external.*(block|cidr)'
            'Private VPC IPv4 CIDRs' = Pick $p '(?i)^private.*(block|cidr)'
            'VPCs' = Join-Lines (Get-Refs $nm $vpcs)
            'Spec (all fields)' = Get-P $p 'spec'
        }
    }
    Add-Vertical $Bk $shVpc 'Connectivity Profiles' $rows 'k8s'

    $rows = foreach ($v in $vpcs) {
        $nm = "$(K8sName $v)"
        [ordered]@{
            'Name' = $nm; 'Region' = K8sRegion $v
            'Transit gateway' = VpcTgw $v
            'Connectivity profile' = Pick $v '(?i)connectivityProfile(Name|Ref)?$'
            'Private - VPC IPv4 CIDRs' = Pick $v '(?i)^private(IPs|Ipv4Cidrs|IPv4CIDRs|Cidrs|IpBlocks?)$'
            'Default SNAT IP' = Pick $v '(?i)snat'
            'Load balancing' = Pick $v '(?i)loadBalanc'
            'Namespace' = NsWhere 'VPC' $nm
            'IP quota status' = Pick $v '(?i)quota'
            'Status' = K8sStatus $v
            'Description' = Get-P $v 'spec.description', 'metadata.annotations.description'
            'Health Status' = K8sConditions $v
            'Spec (all fields)' = Get-P $v 'spec'
            'Status (all fields, incl. IP block and quota usage)' = Get-P $v 'status'
        }
    }
    Add-Vertical $Bk $shVpc 'VPCs - General Details' $rows 'k8s'

    Add-Title $shTgw 'Transit Gateways' "Organization: $org"
    $rows = foreach ($t in $tgws) {
        $nm = "$(K8sName $t)"
        [ordered]@{
            'Name' = $nm; 'Region' = K8sRegion $t
            'Connected VPCs' = VpcsVia $nm
            'Status' = K8sStatus $t
            'Spec (all fields)' = Get-P $t 'spec'
        }
    }
    Add-Vertical $Bk $shTgw 'Transit Gateways' $rows 'k8s'

    Add-Title $shIp 'IP Management' "Organization: $org"
    $rows = foreach ($b in $blocks) {
        $nm = "$(K8sName $b)"
        [ordered]@{
            'Name' = $nm; 'CIDR' = Pick $b '(?i)^cidrs?$'; 'Region' = K8sRegion $b
            'Scope' = Pick $b '(?i)^(visibility|scope|type)$'
            'VPCs' = VpcsVia $nm
            'Applicable quotas' = Join-Lines (Get-Refs $nm $quotas)
            'Connectivity Profiles' = Join-Lines (Get-Refs $nm $profiles)
            'Spec (all fields)' = Get-P $b 'spec'
        }
    }
    Add-Vertical $Bk $shIp 'IP Address Blocks' $rows 'k8s'
    $rows = foreach ($q in $quotas) {
        $nm = "$(K8sName $q)"
        [ordered]@{
            'Name' = $nm; 'Region' = K8sRegion $q
            'Scope' = Pick $q '(?i)^(visibility|scope|type)$'
            'IP blocks' = Pick $q '(?i)ipBlock'
            'Number of single IPs' = Pick $q '(?i)single|ipCount|numberOfIps'
            'Number of CIDRs' = Pick $q '(?i)cidrCount|numberOfCidrs|cidrs$'
            'Max size of subnet' = Pick $q '(?i)maxSubnet|subnetSize|prefix'
            'Associated with' = Join-Lines @(@(Get-Refs $nm $vpcs) + @(Get-Refs $nm $profiles))
            'Spec (all fields)' = Get-P $q 'spec'
        }
    }
    Add-Vertical $Bk $shIp 'IP Quotas' $rows 'k8s'

    # =========================================================================
    #  Access Control
    # =========================================================================
    Add-Title $shAcc 'Access Control' "Organization: $org"
    $rows = foreach ($u in (Items $Src.orgUsers)) {
        $en = Get-P $u 'enabled'; $lk = Get-P $u 'isLocked', 'locked'
        [ordered]@{
            'User Name' = Get-P $u 'username', 'name'; 'Full Name' = Get-P $u 'fullName'; 'Email' = Get-P $u 'email'
            'State' = if ($null -eq $en) { '' } elseif ($en) { 'Enabled' } else { 'Disabled' }
            'Locked' = if ($null -eq $lk) { '' } elseif ($lk) { 'Locked' } else { 'Unlocked' }
            'Effective Roles' = Join-Lines (@(Get-P $u 'roleEntityRefs') | ForEach-Object { Get-P $_ 'name' })
            'Provider Type' = Get-P $u 'providerType'; 'Stranded' = Get-P $u 'stranded'
        }
    }
    Add-Vertical $Bk $shAcc 'Users' $rows 'orgUsers' 'User Name'
    $rows = foreach ($g in (Items $Src.orgGroups)) {
        [ordered]@{
            'Name' = Get-P $g 'name'; 'Description' = Get-P $g 'description'
            'Role' = Join-Lines (@(Get-P $g 'roleEntityRefs', 'roleEntityRef') | ForEach-Object { Get-P $_ 'name' })
            'Type' = Get-P $g 'providerType'
        }
    }
    Add-Vertical $Bk $shAcc 'Groups' $rows 'orgGroups'
    $rows = foreach ($r in (Items $Src.orgRoles)) {
        [ordered]@{ 'Role' = $r.name; 'Description' = Get-P $r 'description'; 'Predefined' = Get-P $r 'readOnly'; 'Bundle Key' = Get-P $r 'bundleKey' }
    }
    Add-Table $Bk $shAcc 'Roles' $rows 'orgRoles'
    $rows = foreach ($s in (Items $Src.serviceAccounts)) {
        [ordered]@{
            'Name' = Get-P $s 'name'; 'Role' = Get-P $s 'role.name', 'role'; 'Status' = Get-P $s 'status'
            'Software ID' = Get-P $s 'softwareId'; 'Client URI' = Get-P $s 'uri', 'clientUri'; 'Software Version' = Get-P $s 'softwareVersion'
        }
    }
    Add-Vertical $Bk $shAcc 'Service Accounts' $rows 'serviceAccounts'

    # =========================================================================
    #  External Orchestrators / Integrations
    # =========================================================================
    $ints = @((Items $Src.integrations) | Where-Object { "$(Get-P $_ 'integrationType')" -notmatch '^abx' })
    $addrKeys = 'hostName', 'hostname', 'url', 'endpoint', 'address', 'serverUrl', 'vroHostname', 'apiUrl', 'apiEndpoint'
    function IntAddr($i) { return (Get-P (Get-P $i 'integrationProperties') $addrKeys) }
    function IntUser($i) { return (Get-P @($i, (Get-P $i 'integrationProperties')) 'privateKeyId', 'username', 'userName', 'user') }

    Add-Title $shExt 'External Orchestrators' "Organization: $org  -  the VIP and organization come from the Orchestrator integrations; node, DNS, NTP and placement details are not exposed by the VCF Automation API"
    $rows = foreach ($i in ($ints | Where-Object { "$(Get-P $_ 'integrationType')" -match '(?i)^vro' })) {
        [ordered]@{
            'Orchestrator' = $i.name; 'VIP' = IntAddr $i
            'Node #1 FQDN' = $NotFromApi; 'Node #2 FQDN' = $NotFromApi; 'Node #3 FQDN' = $NotFromApi
            'Node #1 IP' = $NotFromApi; 'Node #2 IP' = $NotFromApi; 'Node #3 IP' = $NotFromApi
            'DNS #1' = $NotFromApi; 'DNS #2' = $NotFromApi; 'NTP' = $NotFromApi
            'Connected to which Organization' = $org
            'Kubernetes Cluster IP Range #1' = $NotFromApi; 'Kubernetes Cluster IP Range #2' = $NotFromApi
            'Version' = Find-Deep (Get-P $i 'integrationProperties') '(?i)version'
            'Deployed on which site' = $NotFromApi; 'Cluster' = $NotFromApi; 'Network' = $NotFromApi; 'Storage' = $NotFromApi
        }
    }
    Add-Vertical $Bk $shExt 'Orchestrators' $rows 'integrations' 'Orchestrator'

    Add-Title $shInt 'Integrations' "Organization: $org"
    $intTypes = [ordered]@{ 'vro' = 'Orchestrator'; 'github' = 'GitHub'; 'gitlab' = 'GitLab'; 'bitbucket' = 'BitBucket' }
    $seenTypes = @($intTypes.Keys) + @($ints | ForEach-Object { "$(Get-P $_ 'integrationType')" } | Where-Object { $_ -and $intTypes.Keys -notcontains $_ } | Select-Object -Unique)
    foreach ($t in $seenTypes) {
        $title = if ($intTypes.Contains($t)) { $intTypes[$t] } else { $t }
        $rows = foreach ($i in ($ints | Where-Object { "$(Get-P $_ 'integrationType')" -eq $t })) {
            switch ($t) {
                'vro'       { [ordered]@{ 'Name' = $i.name; 'Description' = Get-P $i 'description'; 'Credentials' = IntUser $i } }
                'bitbucket' { [ordered]@{ 'Name' = $i.name; 'Description' = Get-P $i 'description'; 'Server URL' = IntAddr $i; 'Username' = IntUser $i; 'Token' = '(not exported)' } }
                { $_ -in 'github', 'gitlab' } { [ordered]@{ 'Name' = $i.name; 'Description' = Get-P $i 'description'; 'Server URL' = IntAddr $i; 'Token' = '(not exported)' } }
                default     { [ordered]@{ 'Name' = $i.name; 'Description' = Get-P $i 'description'; 'Server URL' = IntAddr $i; 'Credentials' = IntUser $i } }
            }
        }
        Add-Vertical $Bk $shInt $title $rows 'integrations'
    }

    # =========================================================================
    #  Identity Providers
    # =========================================================================
    Add-Title $shIdp 'Identity Providers' "Organization: $org"
    $idp = @{}; foreach ($x in (Items $Src.identityProviders)) { $idp["$(Get-P $x 'kind')"] = $x }
    function IdpNote($k) { $x = $idp[$k]; if ($x -and (Get-P $x 'error')) { "<not collected: $(Get-P $x 'error')>" } else { $null } }
    $o = Get-P $idp['oauth'] 'settings'
    $rows = if ($o -and ((Get-P $o 'enabled') -or (Get-P $o 'issuerId', 'clientId'))) {
        Add-RestOf ([ordered]@{
                'Name' = 'OIDC'; 'Enabled' = Get-P $o 'enabled'; 'Issuer' = Get-P $o 'issuerId'; 'Client ID' = Get-P $o 'clientId'
                'Authorization endpoint' = Get-P $o 'userAuthorizationEndpoint'; 'Token endpoint' = Get-P $o 'accessTokenEndpoint'
                'User info endpoint' = Get-P $o 'userInfoEndpoint'; 'Scope' = Get-P $o 'scope'
                'Claim mapping' = Get-P $o 'oAuthClaimMapping', 'claimMapping'; 'Max clock skew (s)' = Get-P $o 'maxClockSkew'
            }) $o @('enabled', 'issuerId', 'clientId', 'userAuthorizationEndpoint', 'accessTokenEndpoint', 'userInfoEndpoint', 'scope', 'oAuthClaimMapping', 'claimMapping', 'maxClockSkew')
    }
    Add-Vertical $Bk $shIdp 'OIDC' $rows 'identityProviders'
    if (IdpNote 'oauth') { [void](Add-Row $shIdp @(C (IdpNote 'oauth') 6)); Add-Blank $shIdp }

    $l = Get-P $idp['ldap'] 'settings'
    if (-not $l) { $l = Get-P $idp['ldapV1'] 'settings' }
    $mode = "$(Get-P $l 'orgLdapMode', 'ldapMode', 'mode')"
    $cl = Get-P $l 'customOrgLdapSettings', 'customLdapSettings'
    $d = [ordered]@{
        'Name'                                                = 'LDAP'
        'What is the source of users for this organization?' = $mode
        'Do not use LDAP'                                     = if ($mode) { [bool]($mode -eq 'NONE') } else { '' }
        'VCF Automation system LDAP service'                  = if ($mode) { [bool]($mode -eq 'SYSTEM') } else { '' }
        'Distinguished name for the organizational unit'      = Get-P $l 'customUsersOu'
        'Custom LDAP service'                                 = if ($mode) { [bool]($mode -eq 'CUSTOM') } else { '' }
    }
    if ($cl) {
        foreach ($kv in ([ordered]@{ 'Host' = 'hostName'; 'Port' = 'port'; 'SSL' = 'isSsl'; 'Search base' = 'searchBase'; 'Group search base' = 'groupSearchBase'; 'Bind user' = 'userName'; 'Authentication' = 'authenticationMechanism'; 'Connector type' = 'connectorType' }).GetEnumerator()) {
            $d[$kv.Key] = Get-P $cl $kv.Value
        }
        $d = Add-RestOf $d $cl @('hostName', 'port', 'isSsl', 'searchBase', 'groupSearchBase', 'userName', 'authenticationMechanism', 'connectorType')
    }
    $rows = if ($l) { Add-RestOf $d $l @('orgLdapMode', 'ldapMode', 'mode', 'customOrgLdapSettings', 'customLdapSettings', 'customUsersOu') }
    Add-Vertical $Bk $shIdp 'LDAP' $rows 'identityProviders'
    if (-not $l -and (IdpNote 'ldap')) { [void](Add-Row $shIdp @(C (IdpNote 'ldap') 6)); Add-Blank $shIdp }

    $f = Get-P $idp['federation'] 'settings'
    $rows = if ($f -and ((Get-P $f 'enabled') -or (Get-P $f 'samlMetadata'))) {
        Add-RestOf ([ordered]@{
                'Name' = 'SAML'; 'Enabled' = Get-P $f 'enabled'; 'Entity ID' = Get-P $f 'samlSPEntityId'
                'Certificate expiration' = Get-P $f 'certificateExpiration'; 'Attribute mapping' = Get-P $f 'samlAttributeMapping'
            }) $f @('enabled', 'samlSPEntityId', 'certificateExpiration', 'samlAttributeMapping')
    }
    Add-Vertical $Bk $shIdp 'SAML' $rows 'identityProviders'
    if (IdpNote 'federation') { [void](Add-Row $shIdp @(C (IdpNote 'federation') 6)); Add-Blank $shIdp }

    # =========================================================================
    #  Secrets
    # =========================================================================
    Add-Title $shSec 'Secrets' "Organization: $org"
    $rows = foreach ($s in (Items $Src.secrets)) {
        [ordered]@{ 'Name' = $s.name; 'Scope' = if (Get-P $s 'orgScoped') { 'Organization' } else { PName (Get-P $s 'projectId') }; 'Value' = '(hidden)'; 'Description' = Get-P $s 'description' }
    }
    Add-Vertical $Bk $shSec 'Secrets' $rows 'secrets'

    # =========================================================================
    #  Configuration
    # =========================================================================
    Add-Title $shCfg 'Configuration' "Organization: $org"
    $cfg = [ordered]@{}
    foreach ($c in (Items $Src.configProperties)) { $cfg["$(Get-P $c 'key', 'name')"] = Get-P $c 'value' }
    $os = Get-Flat (@($Src.orgSettings) | Select-Object -First 1)
    function CfgFind([string]$Regex) {
        foreach ($k in $cfg.Keys) { if ($k -match $Regex) { return ('{0} = {1}' -f $k, (Format-Value $cfg[$k])) } }
        foreach ($k in $os.Keys) { if ($k -match $Regex) { return ('{0} = {1}' -f $k, (Format-Value $os[$k])) } }
        return ''
    }
    Add-Section $Bk $shCfg 'Settings' 2 3
    Add-KeyValue $shCfg ([ordered]@{
            'Create new resource'  = CfgFind '(?i)new.?resource|create.?resource'
            'Show names of users'  = CfgFind '(?i)show.?(user|name)|user.?names?'
            'IP Release timeout'   = CfgFind '(?i)ip.?release'
        })
    $rows = foreach ($k in $cfg.Keys) { [ordered]@{ 'Name' = $k; 'Value' = $cfg[$k] } }
    Add-Table $Bk $shCfg 'Settings - All Configuration Properties' $rows 'configProperties'
    $rows = foreach ($k in $os.Keys) { [ordered]@{ 'Name' = $k; 'Value' = $os[$k] } }
    Add-Table $Bk $shCfg 'Organization Settings' $rows 'orgSettings'

    $rows = foreach ($s in (Items $Src.notificationScenarios)) {
        $cust = Get-P $s 'customized', 'isCustomized', 'customTemplate'
        [ordered]@{
            'Scenario' = Get-P $s 'scenarioName', 'scenarioId'; 'Category' = Get-P $s 'scenarioCategory'
            'Enable' = [bool](Get-P $s 'enabled')
            'Customized' = if ($null -ne $cust) { if ($cust -is [bool]) { $cust } else { $true } } else { $false }
        }
    }
    Add-Table $Bk $shCfg 'Scenarios' $rows 'notificationScenarios'

    $e = $Src.emailConfig
    $rows = if ($e -and (Get-P $e 'host', 'serverName')) {
        [ordered]@{
            'Name' = Get-P $e 'name'; 'Description' = Get-P $e 'description'; 'Server Name' = Get-P $e 'host', 'serverName'
            'Sender Name' = Get-P $e 'senderName'; 'Sender Address' = Get-P $e 'sender', 'senderAddress'
            'Authentication Required' = Get-P $e 'authentication', 'authenticationRequired'
            'Connection Security' = Get-P $e 'connectionSecurity'; 'Server Port' = Get-P $e 'port'
            'Enable Insecure Connection Upgrade with STARTTLS' = Get-P $e 'insecureConnectionUpgrade', 'allowInsecureUpgrade'
            'Trust Certificates Presented by the Host' = Get-P $e 'trustHost', 'trustCert'
        }
    }
    Add-Vertical $Bk $shCfg 'Email Server' $rows 'emailConfig'

    # =========================================================================
    #  Other CCI resources (everything discovered that no sheet above uses)
    # =========================================================================
    Add-Title $shK8s 'Other CCI Resources' 'Every other resource type discovered through the Cloud Consumption Interface. Full objects are in the JSON.'
    $rows = foreach ($r in $k8sRes) {
        $k = "$(Get-P $r 'kind')"
        if ($usedKinds.Contains($k) -or $k -match '(?i)^Region$') { continue }
        [ordered]@{
            'Kind' = $k; 'API' = Get-P $r 'groupVersion'; 'Resource' = Get-P $r 'resource'; 'Namespaced' = Get-P $r 'namespaced'
            'Count' = (Items (Get-P $r 'items')).Count
            'Names' = Join-Lines @((Items (Get-P $r 'items')) | ForEach-Object { $ns = Get-P $_ 'metadata.namespace'; if ($ns) { "$ns/$(K8sName $_)" } else { K8sName $_ } })
            'Errors' = Join-Lines (Get-P $r 'errors')
        }
    }
    Add-Table $Bk $shK8s 'Resource Types' $rows 'k8s'

    # =========================================================================
    #  Version History, Collection Log, Index
    # =========================================================================
    Add-Title $shVer 'Version History'
    [void](Add-Row $shVer @((C 'Date' 3), (C 'Version' 3), (C 'Author' 3), (C 'Description' 3), (C 'Reviewers' 3)))
    [void](Add-Row $shVer @((C (Format-Value $Data.meta.collectedAt)), (C '0.1'), (C $Data.meta.collectedBy), (C "Auto-collected from $($Data.meta.server) (org $org) by script v$($Data.meta.scriptVersion)"), (C '')))

    Add-Title $shLog 'Collection Log' 'One line per API call group. FAILED usually means the endpoint does not exist in this version or the account lacks rights.'
    Add-Table $Bk $shLog 'API calls' (Items $Data.log)

    Add-Title $shIdx 'VCF Automation - All Apps Organization Configuration'
    Add-KeyValue $shIdx ([ordered]@{
            'Organization'     = $org
            'VCF Automation'   = $Data.meta.server
            'Collected at'     = $Data.meta.collectedAt
            'Collected by'     = $Data.meta.collectedBy
            'Login method'     = $Data.meta.auth
            'IaaS API version' = $Data.meta.iaasApiVersion
            'Script version'   = $Data.meta.scriptVersion
            'Failed sections'  = if ($script:FailedKeys.Count) { $script:FailedKeys -join ', ' } else { 'none' }
        })
    [void](Add-Row $shIdx @((C 'Contents' 2), (C '' 2), (C '' 2)))
    [void](Add-Row $shIdx @((C 'Sheet' 3), (C 'Section' 3), (C 'Items' 3)))
    foreach ($ix in $Bk.Index) {
        if ($ix.Sheet -eq 'Collection Log') { continue }
        $r = Add-Row $shIdx @((C $ix.Sheet), (C $ix.Title 7), (C $ix.Count))
        $shIdx.Links.Add(@{ Ref = "B$r"; Location = "'$($ix.Sheet)'!A$($ix.Row)"; Display = $ix.Title })
    }

    Save-Book $Bk $Path
}
# =====================================================================================
#  Main
# =====================================================================================

function Get-FreePath([string]$Folder, [string]$Stem, [string[]]$Ext) {
    $name = $Stem; $i = 1
    while ($true) {
        $taken = $false
        foreach ($x in $Ext) { if (Test-Path -LiteralPath (Join-Path $Folder "$name$x")) { $taken = $true } }
        if (-not $taken) { return (Join-Path $Folder $name) }
        $i++; $name = '{0}_v{1}' -f $Stem, $i
    }
}

$script:MaxCols = [math]::Max(1, $MaxColumns)
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path

if ($PSCmdlet.ParameterSetName -eq 'Render') {
    $jsonPath = (Resolve-Path -LiteralPath $FromJson).Path
    $data = Read-Json ([System.IO.File]::ReadAllText($jsonPath))
    $stem = [IO.Path]::GetFileNameWithoutExtension($jsonPath)
    $out = (Get-FreePath (Split-Path $jsonPath -Parent) $stem @('.xlsx')) + '.xlsx'
    Build-Workbook $data $out
    Write-Host "Workbook written: $out" -ForegroundColor Green
    return
}

# ---- hard-coded presets (only fill what was not passed on the command line)
if (-not $Server -and $Preset.Server) { $Server = $Preset.Server }
if (-not $Organization -and $Preset.Organization) { $Organization = $Preset.Organization }
if (-not $PSBoundParameters.ContainsKey('SkipCertificateCheck') -and $null -ne $Preset.SkipCertificateCheck) {
    $SkipCertificateCheck = [bool]$Preset.SkipCertificateCheck
    $PSBoundParameters['SkipCertificateCheck'] = $SkipCertificateCheck
}
if (-not $UseApiToken -and -not $Credential) {
    if ($Preset.CredentialFile -and (Test-Path -LiteralPath $Preset.CredentialFile)) {
        $Credential = Import-Clixml -LiteralPath $Preset.CredentialFile
    }
    elseif ($Preset.Username -and $Preset.Password) {
        $Credential = New-Object System.Management.Automation.PSCredential ($Preset.Username, (ConvertTo-SecureString $Preset.Password -AsPlainText -Force))
    }
    elseif ($Preset.Username) {
        $sec = Read-Host -AsSecureString "Password for $($Preset.Username)"
        $Credential = New-Object System.Management.Automation.PSCredential ($Preset.Username, $sec)
    }
}

# ---- interactive prompts (only for what is still missing)
if (-not $BaseUri) {
    while (-not $Server) { $Server = Read-Host 'VCF Automation FQDN (e.g. vcfa.corp.local)' }
    $Server = ($Server -replace '^https?://', '').TrimEnd('/')
    $BaseUri = "https://$Server"
}
else { $BaseUri = $BaseUri.TrimEnd('/'); if (-not $Server) { $Server = ([uri]$BaseUri).Authority } }
while (-not $Organization) { $Organization = Read-Host 'All Apps organization name' }
if (-not $PSBoundParameters.ContainsKey('SkipCertificateCheck')) {
    $ans = Read-Host 'Skip TLS certificate validation (self-signed / untrusted CA)? [y/N]'
    $SkipCertificateCheck = [bool]($ans -match '^(y|yes)$')
}
if (-not $UseApiToken -and -not $Credential) {
    $Credential = Get-Credential -Message "VCF Automation user for organization '$Organization' (@$Organization is added automatically)"
    if (-not $Credential) { throw 'No credentials entered.' }
}

$script:Conn.Base = $BaseUri
$script:Conn.Skip = [bool]$SkipCertificateCheck
Set-TlsOptions ([bool]$SkipCertificateCheck)

Write-Host "Logging in to $BaseUri as organization '$Organization' ..." -ForegroundColor Cyan
$token = Connect-Vcfa
$script:Conn.Headers = @{ Authorization = "Bearer $token"; Accept = 'application/json' }
$script:Reconnect = $true
Write-Host "  $($script:AuthInfo)" -ForegroundColor DarkGray
$authInfo = $script:AuthInfo

$started = Get-Date
Invoke-Collection

$data = [ordered]@{
    meta   = [ordered]@{
        organization   = $Organization
        server         = $Server
        collectedAt    = $started.ToString('yyyy-MM-dd HH:mm')
        collectedBy    = if ($Credential) { $Credential.UserName } else { 'API token' }
        auth           = $authInfo
        iaasApiVersion = $script:Conn.IaasApiVersion
        scriptVersion  = $ScriptVersion
    }
    failed = @($script:FailedKeys)
    log    = @($script:Log)
    raw    = $script:Raw
}

$safeOrg = $Organization -replace '[^\w\-]', '_'
# One folder per org type + org, e.g. "All Apps - my-allapps-org"
$orgFolder = Join-Path $OutputFolder ('All Apps - ' + ($Organization -replace '[\\/:*?"<>|]', '_').Trim())
if (-not (Test-Path -LiteralPath $orgFolder)) { New-Item -ItemType Directory -Path $orgFolder | Out-Null }
$base = Get-FreePath $orgFolder ("VCFA_{0}_AllApps_Config_{1}" -f $safeOrg, (Get-Date -Format 'yyyyMMdd')) @('.xlsx', '.json')

# Redact anything that looks like a credential before it touches disk
$json = $data | ConvertTo-Json -Depth 64
$json = [regex]::Replace($json, '"([A-Za-z_]*(?:[Pp]assword|[Ss]ecret|[Tt]oken)|privateKey)"\s*:\s*"(?:[^"\\]|\\.)*"', '"$1": "<redacted>"')
[System.IO.File]::WriteAllText("$base.json", $json, [System.Text.UTF8Encoding]::new($false))

Build-Workbook (Read-Json $json) "$base.xlsx"

$ok = @($script:Log | Where-Object { $_.Status -eq 'OK' }).Count
Write-Host ''
Write-Host "Done. $ok of $($script:Log.Count) API sections collected." -ForegroundColor Green
if ($script:FailedKeys.Count) { Write-Host "Failed: $($script:FailedKeys -join ', ')  (details on the 'Collection Log' sheet)" -ForegroundColor Yellow }
Write-Host "Workbook : $base.xlsx"
Write-Host "JSON     : $base.json"

#Requires -Version 5.1
<#
.SYNOPSIS
    Collects the configuration of one VCF Automation 9.x "VM Apps" organization and writes
    a configuration workbook (.xlsx) plus a raw JSON dump for quick reference.

.DESCRIPTION
    Read-only. Uses GET calls only (plus the login POST). Prompts for anything not passed:
    VCF Automation FQDN, organization name, TLS check, and credentials.

    Output (in a sub-folder "VM Apps - <Org>" of -OutputFolder, default = current folder):
        VM Apps - <Org>\VCFA_<Org>_VMApps_Config_<yyyyMMdd>.xlsx
        VM Apps - <Org>\VCFA_<Org>_VMApps_Config_<yyyyMMdd>.json
    If either file already exists, a suffix _v2, _v3 ... is added so nothing is overwritten.

    The workbook is written natively (Open XML), so no ImportExcel / Excel install is needed
    on the customer jump host.

    Passwords are never returned by the APIs; any field that looks like a password, secret
    or token is redacted before the JSON is saved.

.PARAMETER Server
    VCF Automation FQDN, e.g. vcfa.corp.local

.PARAMETER Organization
    VM Apps organization name (as used at login).

.PARAMETER Credential
    Org user. "@<Organization>" is appended to the user name if it is not already there.

.PARAMETER UseApiToken
    Log in with an organization API token instead of user name / password.

.PARAMETER SkipCertificateCheck
    Skip TLS certificate validation (self-signed / untrusted CA). Asked interactively if omitted.

.PARAMETER SkipFormCheck
    Skip the per-catalog-item custom request form lookup (one API call per item). Useful for very large catalogs.

.PARAMETER MaxColumns
    Objects shown side by side (#1, #2 ...) up to this count; above it the section switches to a row-per-object table. Default 10.

.PARAMETER OnlyUsedNetworks
    In the Networks table, list only networks that are in a network profile, have a CIDR, or have an IP range.
    (Everything is still kept in the JSON.)

.PARAMETER FromJson
    Re-build the workbook from a JSON file produced by an earlier run (no connection needed).

.EXAMPLE
    ./Get-VCFAVMAppsConfig.ps1
    Fully interactive.

.EXAMPLE
    ./Get-VCFAVMAppsConfig.ps1 -Server vcfa.corp.local -Organization acme -SkipCertificateCheck

.EXAMPLE
    ./Get-VCFAVMAppsConfig.ps1 -FromJson .\VCFA_acme_VMApps_Config_20261027.json
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
    [Parameter(ParameterSetName = 'Collect')][switch]$SkipFormCheck,
    [int]$MaxColumns = 10,
    [switch]$OnlyUsedNetworks,
    [string]$OutputFolder = (Get-Location).Path
)

# =====================================================================================
#  OPTIONAL HARD-CODED INPUTS
#  Fill in any of these to skip the matching prompt. Leave a value empty ('' or $null) to be asked.
#  Command-line parameters always win over these values.
# =====================================================================================
$Preset = @{
    Server               = ''        # VCF Automation FQDN, e.g. 'vcfa.corp.local'
    Organization         = ''        # VM Apps organization name, e.g. 'my-vmapps-org'
    SkipCertificateCheck = $null     # $true / $false, or $null to be asked
    Username             = ''        # e.g. 'admin'  (@<Organization> is added automatically)
    Password             = ''        # Plain text - avoid on customer machines; prefer CredentialFile below
    CredentialFile       = ''        # Path to an encrypted credential file made with:
                                     #   Get-Credential | Export-Clixml -Path C:\path\vcfa-cred.xml
                                     # Only the same Windows user on the same machine can decrypt it.
}
# =====================================================================================

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.3.4'
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

function Get-ProfileNetLinks($np) {
    # All networks a network profile points to: id fields plus every _links entry that names networks
    # (fabric-networks, and any other *network* link the API version uses), except domains / self / region.
    $out = [ordered]@{}
    foreach ($id in @(Get-P $np 'fabricNetworkIds') + @(Get-P $np 'fabricNetworksIds')) { if ($id -and -not $out.Contains("$id")) { $out["$id"] = $null } }
    foreach ($m in (Get-Members (Get-P $np '_links'))) {
        if ($m.Name -notmatch '(?i)network' -or $m.Name -match '(?i)domain|profile|self|region|security') { continue }
        $h = @()
        $one = Get-P $m.Value 'href'; if ($one) { $h += $one }
        $many = Get-P $m.Value 'hrefs'; if ($many) { $h += @($many) }
        foreach ($x in $h) {
            $id = (("$x" -split '\?')[0].TrimEnd('/') -split '/')[-1]
            if ($id -and -not $out.Contains($id)) { $out[$id] = "$x" }
        }
    }
    return $out
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

function Convert-VroLink($l) {
    # One vRO inventory object -> ordered dictionary of ALL its attributes, under their raw names
    # (name, @fullType, vraHost, sessionMode ... as in the workbook template). Only credential fields are left out.
    $d = [ordered]@{}
    foreach ($a in (Items (Get-P $l 'attributes'))) {
        $n = "$($a.name)"
        if (-not $n -or $n -match '(?i)password|secret|token') { continue }
        if (-not $d.Contains($n)) { $d[$n] = $a.value }
    }
    return $d
}

function Get-Xenon([string]$Path) {
    # Older-style APIs (e.g. onboarding) return { documentLinks, documents{}, nextPageLink }.
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Path + $(if ($Path.Contains('?')) { '&' } else { '?' }) + 'expand=true'
    for ($i = 0; $i -lt 1000 -and $next; $i++) {
        $r = Invoke-Vcfa $next
        $docs = Get-P $r 'documents'
        if ($null -ne $docs -and -not ($docs -is [array])) {
            foreach ($m in (Get-Members $docs)) { $all.Add($m.Value) }
        }
        else {
            foreach ($x in (Get-Items $r)) { if ($null -ne $x) { $all.Add($x) } }
            break
        }
        $next = Get-P $r 'nextPageLink'
        if ($next -and -not "$next".StartsWith('/')) { $next = "/$next" }
    }
    return , $all.ToArray()
}

function Get-VroCatalog([string]$Ns, [string]$Type) {
    $all = [System.Collections.Generic.List[object]]::new()
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    for ($i = 0; $i -lt 1000; $i++) {
        $r = Invoke-Vcfa ("/vco/api/catalog/{0}/{1}?maxResult=500&startIndex={2}" -f $Ns, $Type, $all.Count)
        $new = 0
        foreach ($l in (Items (Get-P $r 'link'))) {
            if ($seen.Add("$(Get-P $l 'href')")) { $new++; $all.Add((Convert-VroLink $l)) }
        }
        $total = Get-P $r 'total'
        if ($new -eq 0 -or $null -eq $total -or $all.Count -ge [int]$total) { break }
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
#  Collection
# =====================================================================================

# Plug-ins whose root inventory holds no configured endpoints (skipped when discovering endpoints)
$VroSkipModules = @()   # nothing skipped: every plug-in's root inventory is captured
$VroTypeTitles = @{
    'AD:AdHost'                 = 'Active Directory'
    'AD:AD_Host'                = 'Active Directory'
    'VRA:Host'                  = 'VCF Automation (VM Apps) hosts'
    'VCFA:Host'                 = 'VCF Automation (All Apps) hosts'
    'PowerShell:PowerShellHost' = 'PowerShell hosts'
    'REST:RESTHost'             = 'REST hosts'
    'VC:SdkConnection'          = 'vCenter connections'
    'SSH:Host'                  = 'SSH hosts'
    'SOAP:SOAPHost'             = 'SOAP hosts'
    'SQL:Database'              = 'SQL databases'
    'AMQP:Broker'               = 'AMQP brokers'
    'VAPI:VAPIEndpoint'         = 'vAPI endpoints'
    'Mail:MailServer'           = 'Mail servers'
    'SNMP:SnmpDevice'           = 'SNMP devices'
    'DynamicTypes:DynamicNamespaceDefinition' = 'Dynamic Types namespaces'
    'Configurator:Keystores'    = 'Keystores'
}

function Invoke-Collection {
    Write-Host "`nCollecting from $($script:Conn.Base) ..." -ForegroundColor Cyan

    Collect 'about' '/iaas/api/about' -Single
    $script:Conn.IaasApiVersion = Get-P $script:Raw.about 'latestApiVersion'
    if (-not $script:Conn.IaasApiVersion) { $script:Conn.IaasApiVersion = '2021-07-15' }

    # --- Infrastructure
    Collect 'integrations'           '/iaas/api/integrations'
    Collect 'cloudAccounts'          '/iaas/api/cloud-accounts'
    Collect 'cloudAccountsVsphere'   '/iaas/api/cloud-accounts-vsphere'
    Collect 'cloudAccountsNsxt'      '/iaas/api/cloud-accounts-nsx-t'
    Collect 'regions'                '/iaas/api/regions'
    Collect 'zones'                  '/iaas/api/zones'
    Collect 'zoneComputes'           '/iaas/api/zones/{id}/computes' -Script {
        foreach ($z in (Items $script:Raw.zones)) {
            [pscustomobject]@{ zoneId = $z.id; computes = @(Get-All "/iaas/api/zones/$($z.id)/computes") }
        }
    }
    Collect 'fabricComputes'         '/iaas/api/fabric-computes'
    Collect 'flavorProfiles'         '/iaas/api/flavor-profiles'
    Collect 'imageProfiles'          '/iaas/api/image-profiles'
    Collect 'storageProfilesVsphere' '/iaas/api/storage-profiles-vsphere'
    Collect 'datastores'             '/iaas/api/fabric-vsphere-datastores'
    Collect 'storagePolicies'        '/iaas/api/fabric-vsphere-storage-policies'
    Collect 'fabricNetworks'         '/iaas/api/fabric-networks'
    Collect 'fabricNetworksVsphere'  '/iaas/api/fabric-networks-vsphere'
    Collect 'ipRanges'               '/iaas/api/network-ip-ranges'
    Collect 'externalIpRanges'       '/iaas/api/external-network-ip-ranges'
    Collect 'networkDomains'         '/iaas/api/network-domains'
    Collect 'loadBalancers'          '/iaas/api/load-balancers'
    Collect 'securityGroups'         '/iaas/api/security-groups'
    Collect 'networkProfiles'        '/iaas/api/network-profiles'
    Collect 'networkProfileNetworks' '/iaas/api/network-profiles/{id}' -Script {
        $known = @{}
        foreach ($n in (Items $script:Raw.fabricNetworks) + (Items $script:Raw.fabricNetworksVsphere)) { $known["$($n.id)"] = $true }
        foreach ($np in (Items $script:Raw.networkProfiles)) {
            $links = Get-ProfileNetLinks $np
            # the list response can omit the network links; the single-profile call carries them
            try { $det = Invoke-Vcfa "/iaas/api/network-profiles/$($np.id)"; foreach ($e in (Get-ProfileNetLinks $det).GetEnumerator()) { if (-not $links.Contains($e.Key)) { $links[$e.Key] = $e.Value } } }
            catch { $script:Warnings.Add("profile $($np.name): $(Get-ErrMsg $_)") }
            $fetched = [System.Collections.Generic.List[object]]::new()
            foreach ($e in $links.GetEnumerator()) {
                if ($known.ContainsKey("$($e.Key)")) { continue }
                $path = if ($e.Value) { '/' + (("$($e.Value)" -replace '^https?://[^/]+', '').TrimStart('/')) } else { "/iaas/api/fabric-networks/$($e.Key)" }
                try { $fetched.Add((Invoke-Vcfa $path)); $known["$($e.Key)"] = $true }
                catch { $script:Warnings.Add("network $($e.Key) in profile $($np.name): $(Get-ErrMsg $_)") }
            }
            [pscustomobject]@{ profileId = "$($np.id)"; networkIds = @($links.Keys); fetched = @($fetched) }
        }
    }
    Collect 'pricingCards'           '/price/api/private/pricing-cards'

    # --- Administration
    Collect 'projects'               '/iaas/api/projects'
    Collect 'projectService'         '/project-service/api/projects'
    Collect 'orgRoles'               '/cloudapi/1.0.0/roles'
    Collect 'orgUsers'               '/cloudapi/1.0.0/users'
    Collect 'orgGroups'              '/cloudapi/1.0.0/groups'
    Collect 'naming'                 '/iaas/api/naming'
    Collect 'secrets'                '/platform/api/secrets'
    Collect 'identityProviders'      '/api/admin/org/{id}/settings/(oauth|ldap|federation)' -Script {
        $orgs = Get-All '/cloudapi/1.0.0/orgs'
        $o = $orgs | Where-Object { "$($_.name)" -eq $Organization } | Select-Object -First 1
        if (-not $o) { $o = $orgs | Select-Object -First 1 }
        if (-not $o) { throw 'Organization not visible through /cloudapi/1.0.0/orgs' }
        $uuid = ("$($o.id)" -split ':')[-1]
        foreach ($kind in 'oauth', 'ldap', 'federation') {
            $res = $null; $err = $null
            try { $res = Invoke-Vcfa "/api/admin/org/$uuid/settings/$kind" }
            catch { $code = Get-StatusCode $_; if ($code -ne 404) { $err = Get-ErrMsg $_ } }
            [pscustomobject]@{ kind = $kind; settings = $res; error = $err }
        }
    }
    Collect 'configProperties'       '/iaas/api/configuration-properties'
    Collect 'tags'                   '/iaas/api/tags'
    Collect 'onboardingPlans'        '/relocation/onboarding/plan' -Script { Get-Xenon '/relocation/onboarding/plan' }
    Collect 'onboardingDeployments'  '/relocation/onboarding/deployment' -Script { Get-Xenon '/relocation/onboarding/deployment' }
    Collect 'onboardingResources'    '/relocation/onboarding/resource' -Script { Get-Xenon '/relocation/onboarding/resource' }

    # --- Content & Policies
    Collect 'contentSources'         '/catalog/api/admin/sources'
    Collect 'catalogItems'           '/catalog/api/admin/items'
    if (-not $SkipFormCheck) { Collect 'catalogForms' '/form-service/api/forms/fetchBySourceAndType' -Script {
        $all = Items $script:Raw.catalogItems; $n = 0
        foreach ($it in $all) {
            $n++
            if ($all.Count -gt 20) { Write-Progress -Activity 'Custom request forms' -Status "$n of $($all.Count)" -PercentComplete (100 * $n / $all.Count) }
            $st = Get-P $it 'type.id'
            $status = 'Not customised'
            try {
                $f = Invoke-Vcfa ("/form-service/api/forms/fetchBySourceAndType?formType=requestForm&sourceType={0}&sourceId={1}" -f [uri]::EscapeDataString("$st"), $it.id)
                if ($f) { $s = Get-P $f 'status'; $status = if ($s -eq 'ON') { 'Enabled' } elseif ($s) { "Customised ($s)" } else { 'Customised' } }
            }
            catch {
                $code = Get-StatusCode $_
                if ($code -ne 404 -and $code -ne 204) { $status = "Unknown (HTTP $code)" }
            }
            [pscustomobject]@{ itemId = $it.id; status = $status }
        }
        Write-Progress -Activity 'Custom request forms' -Completed
    } }
    Collect 'blueprints'             '/blueprint/api/blueprints'
    Collect 'policies'               '/policy/api/policies?expandDefinition=true'
    Collect 'emailConfig'            '/notification/api/email-config' -Single -EmptyOn404
    Collect 'notificationScenarios'  '/notification/api/scenario-configs'

    # --- Orchestrator
    Collect 'vroAbout'               '/vco/api/about' -Single
    Collect 'vroPlugins'             '/vco/api/plugins' -Script { @(Get-P (Invoke-Vcfa '/vco/api/plugins') 'plugins', 'plugin') | Where-Object { $_ } }
    Collect 'vroEndpoints'           '/vco/api/catalog/{plug-in}/{type}' -Script {
        # Discover which endpoint types each plug-in has at its inventory root, then page through each type.
        foreach ($p in (Items $script:Raw.vroPlugins)) {
            $mod = Get-P $p 'moduleName', 'name'
            if (-not $mod -or $VroSkipModules -contains $mod) { continue }
            try { $root = Invoke-Vcfa "/vco/api/catalog/$mod" }
            catch {
                $code = Get-StatusCode $_
                if ($code -ne 404) { $script:Warnings.Add("$mod root: $(Get-ErrMsg $_)") }
                continue
            }
            # group the root's objects by type (kept as a fallback if the per-type listing is not supported)
            $byType = [ordered]@{}
            foreach ($l in (Items (Get-P $root 'relations.link'))) {
                $ft = $null
                foreach ($a in (Items (Get-P $l 'attributes'))) { if ($a.name -eq '@fullType') { $ft = "$($a.value)" } }
                if (-not $ft) { continue }
                if (-not $byType.Contains($ft)) { $byType[$ft] = [System.Collections.Generic.List[object]]::new() }
                $byType[$ft].Add((Convert-VroLink $l))
            }
            foreach ($ft in @($byType.Keys)) {
                $ty = ($ft -split ':', 2)[1]
                # each type on its own: one unsupported type must never stop the others
                try { $objs = Get-VroCatalog $mod $ty }
                catch {
                    $objs = $byType[$ft].ToArray()
                    if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("${ft}: $(Get-ErrMsg $_) (root listing used)") }
                }
                if (@($objs).Count -eq 0) { $objs = $byType[$ft].ToArray() }
                # keep only real configured endpoints: containers / folders (e.g. VAPI:Endpoints, folders, keystores) have no 'name'
                $objs = @($objs | Where-Object { Get-P $_ 'name' })
                if ($objs.Count -eq 0) { continue }
                Write-Host ("`n      {0,-45} {1}" -f $ft, @($objs).Count) -NoNewline -ForegroundColor DarkGray
                [pscustomobject]@{ module = $mod; fullType = $ft; objects = @($objs) }
            }
        }
        Write-Host ''
        Write-Host ('  {0,-28}' -f '') -NoNewline
    }
}

# =====================================================================================
#  Workbook rendering (works only from the saved JSON, so -FromJson gives identical output)
# =====================================================================================

$ContentSourceTypes = @{
    'com.vmw.blueprint'          = 'VCF Automation Templates'
    'com.vmw.vro.workflow'       = 'Automation Orchestrator Workflows'
    'com.vmw.abx.actions'        = 'Extensibility Actions'
    'com.vmw.codestream'         = 'Pipelines'
    'com.vmw.mp.content.gitlab'  = 'Marketplace / GitLab'
    'com.vmw.mp.content.github'  = 'Marketplace / GitHub'
    'com.vmw.blueprint.gitlab'   = 'GitLab'
    'com.vmw.blueprint.github'   = 'GitHub'
}

$PolicyTypes = [ordered]@{
    'com.vmware.policy.catalog.entitlement' = 'Content Sharing Policies'
    'com.vmware.policy.approval'            = 'Approval Policies'
    'com.vmware.policy.deployment.action'   = 'Day 2 Actions Policies'
    'com.vmware.policy.deployment.lease'    = 'Lease Policies'
    'com.vmware.policy.resource.quota'      = 'Resource Quota Policies'
    'com.vmware.policy.deployment.limit'    = 'Deployment Limit Policies'
    'com.vmware.policy.iaas.resource'       = 'IaaS Resource Policies'
}

function Build-Workbook {
    param($Data, [string]$Path)
    $Src = $Data.raw
    $script:FailedKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($k in (Items $Data.failed)) { $script:FailedKeys.Add("$k") }

    $Bk = New-Book
    $shIdx = Add-Sheet $Bk 'Index'
    $shVer = Add-Sheet $Bk 'Version'
    $shInf = Add-Sheet $Bk 'Infrastructure'
    $shAdm = Add-Sheet $Bk 'Administration'
    $shNp  = Add-Sheet $Bk 'Infra - Network Profiles'
    $shTag = Add-Sheet $Bk 'Tags'
    $shOnb = Add-Sheet $Bk 'Onboarding'
    $shCp  = Add-Sheet $Bk 'Content & Policies'
    $shVro = Add-Sheet $Bk 'Orchestrator'
    $shLog = Add-Sheet $Bk 'Collection Log'
    $org = $Data.meta.organization

    # ---------------- lookups ----------------
    $acctById = @{}; $acctDet = @{}; $acctType = @{}
    foreach ($a in (Items $Src.cloudAccounts)) { $acctById["$($a.id)"] = $a; $acctType["$($a.id)"] = Get-P $a 'cloudAccountType' }
    foreach ($a in (Items $Src.cloudAccountsVsphere)) { $acctDet["$($a.id)"] = $a; if (-not $acctById.ContainsKey("$($a.id)")) { $acctById["$($a.id)"] = $a }; if (-not $acctType["$($a.id)"]) { $acctType["$($a.id)"] = 'vsphere' } }
    foreach ($a in (Items $Src.cloudAccountsNsxt))    { $acctDet["$($a.id)"] = $a; if (-not $acctById.ContainsKey("$($a.id)")) { $acctById["$($a.id)"] = $a }; if (-not $acctType["$($a.id)"]) { $acctType["$($a.id)"] = 'nsxt' } }
    $acctOrder = @($acctById.Keys | Sort-Object { if ($acctType[$_] -match "nsx") { 1 } else { 0 } }, { $acctById[$_].name })
    $AcctName = { param($id) if ($id -and $acctById.ContainsKey("$id")) { $acctById["$id"].name } else { "$id" } }

    $regionById = @{}; $regionList = [System.Collections.Generic.List[object]]::new()
    foreach ($rg in (Items $Src.regions)) {
        $caId = Get-P $rg 'cloudAccountId'
        if (-not $caId) { $caId = (LinkIds $rg 'cloud-account') | Select-Object -First 1 }
        $an = if ($caId -and $acctById.ContainsKey("$caId")) { $acctById["$caId"].name } else { $null }
        $lbl = if ($an) { "$an / $($rg.name)" } else { "$($rg.name)" }
        $o = [pscustomobject]@{ Id = "$($rg.id)"; Ext = "$(Get-P $rg 'externalRegionId')"; Ca = "$caId"; Label = $lbl }
        $regionById[$o.Id] = $o; $regionList.Add($o)
    }
    function RegionOf($o) {
        $rid = Get-P $o 'regionId'
        if (-not $rid) { $rid = (LinkIds $o 'region') | Select-Object -First 1 }
        if ($rid -and $regionById.ContainsKey("$rid")) { return $regionById["$rid"].Label }
        $ext = Get-P $o 'externalRegionId'
        if ($ext) {
            $cas = @(@(Get-P $o 'cloudAccountIds') + @(Get-P $o 'cloudAccountId') + (LinkIds $o 'cloud-accounts') + (LinkIds $o 'cloud-account') | Where-Object { $_ })
            $m = $regionList | Where-Object { $_.Ext -eq $ext -and ($cas.Count -eq 0 -or $cas -contains $_.Ca) } | Select-Object -First 1
            if (-not $m) { $m = $regionList | Where-Object { $_.Ext -eq $ext } | Select-Object -First 1 }
            if ($m) { return $m.Label }
            return "$ext"
        }
        return ''
    }

    $projects = Items $Src.projects
    $projById = @{}; foreach ($p in $projects) { $projById["$($p.id)"] = $p }
    foreach ($p in (Items $Src.projectService)) { if (-not $projById.ContainsKey("$($p.id)")) { $projById["$($p.id)"] = $p } }
    if ($projects.Count -eq 0) { $projects = Items $Src.projectService }
    $psById = @{}; foreach ($p in (Items $Src.projectService)) { $psById["$($p.id)"] = $p }
    function PName($id) { if ($id -and $projById.ContainsKey("$id")) { return $projById["$id"].name } elseif ($id) { return "$id" } else { return '' } }
    function ZoneAssign($p) { return , (Items (Get-P $p 'zones', 'zoneAssignmentConfigurations')) }

    $zones = Items $Src.zones
    $zoneById = @{}; foreach ($z in $zones) { $zoneById["$($z.id)"] = $z }
    $zoneProjects = @{}
    foreach ($p in $projects) {
        foreach ($za in (ZoneAssign $p)) {
            $zid = "$(Get-P $za 'zoneId')"
            if (-not $zoneProjects.ContainsKey($zid)) { $zoneProjects[$zid] = [System.Collections.Generic.List[string]]::new() }
            $zoneProjects[$zid].Add("$($p.name)")
        }
    }
    $zoneComp = @{}; $compZones = @{}
    foreach ($zc in (Items $Src.zoneComputes)) {
        $zoneComp["$($zc.zoneId)"] = Items $zc.computes
        $zn = if ($zoneById.ContainsKey("$($zc.zoneId)")) { $zoneById["$($zc.zoneId)"].name } else { "$($zc.zoneId)" }
        foreach ($c in (Items $zc.computes)) {
            if (-not $compZones.ContainsKey("$($c.id)")) { $compZones["$($c.id)"] = [System.Collections.Generic.List[string]]::new() }
            $compZones["$($c.id)"].Add($zn)
        }
    }

    $dsById = @{}; foreach ($d in (Items $Src.datastores)) { $dsById["$($d.id)"] = $d }
    $spolById = @{}; foreach ($d in (Items $Src.storagePolicies)) { $spolById["$($d.id)"] = $d }

    $netVs = @{}; foreach ($n in (Items $Src.fabricNetworksVsphere)) { $netVs["$($n.id)"] = $n }
    $nets = Items $Src.fabricNetworks
    if ($nets.Count -eq 0) { $nets = Items $Src.fabricNetworksVsphere }
    $netById = @{}; foreach ($n in $nets) { $netById["$($n.id)"] = $n }
    foreach ($n in (Items $Src.fabricNetworksVsphere)) { if (-not $netById.ContainsKey("$($n.id)")) { $netById["$($n.id)"] = $n } }
    $ndById = @{}; foreach ($d in (Items $Src.networkDomains)) { $ndById["$($d.id)"] = $d }
    function NetP($n, [string[]]$Paths) { return (Get-P @($netVs["$($n.id)"], $n) $Paths) }
    function NetDomain($n) {
        $id = Get-P $n 'networkDomainId'
        if (-not $id) { $id = (LinkIds $n 'network-domain') | Select-Object -First 1 }
        if ($id -and $ndById.ContainsKey("$id")) { return $ndById["$id"].name }
        return (Get-P $n 'networkDomainName')
    }
    function VpcInfo($n) {
        $cp = Get-P @($netVs["$($n.id)"], $n) 'customProperties'
        $out = @((Get-Members $cp) | Where-Object { $_.Name -match '(?i)vpc|nsxProject|projectPath' } | ForEach-Object { "$($_.Name)=$($_.Value)" })
        return (Join-Lines $out)
    }
    $zonesByRegion = @{}
    foreach ($z in $zones) {
        $zr = RegionOf $z
        if (-not $zonesByRegion.ContainsKey($zr)) { $zonesByRegion[$zr] = [System.Collections.Generic.List[string]]::new() }
        $zonesByRegion[$zr].Add("$($z.name)")
    }
    $netProfiles = Items $Src.networkProfiles
    $profNetIds = @{}
    foreach ($x in (Items $Src.networkProfileNetworks)) {
        $profNetIds["$(Get-P $x 'profileId')"] = [object[]](Items (Get-P $x 'networkIds'))
        foreach ($n in (Items (Get-P $x 'fetched'))) { if ($n -and -not $netById.ContainsKey("$($n.id)")) { $netById["$($n.id)"] = $n } }
    }
    function ProfNets($np) {
        $ids = @(); if ($profNetIds.ContainsKey("$($np.id)")) { $ids += $profNetIds["$($np.id)"] }
        foreach ($k in (Get-ProfileNetLinks $np).Keys) { $ids += "$k" }
        return , @($ids | Where-Object { $_ } | Select-Object -Unique)
    }
    $netInProfiles = @{}
    foreach ($np in $netProfiles) {
        foreach ($nid in (ProfNets $np)) {
            if (-not $netInProfiles.ContainsKey("$nid")) { $netInProfiles["$nid"] = [System.Collections.Generic.List[string]]::new() }
            $netInProfiles["$nid"].Add("$($np.name)")
        }
    }

    $srcById = @{}; foreach ($s in (Items $Src.contentSources)) { $srcById["$($s.id)"] = $s }
    $itemById = @{}; foreach ($i in (Items $Src.catalogItems)) { $itemById["$($i.id)"] = $i }
    $formByItem = @{}; foreach ($f in (Items $Src.catalogForms)) { $formByItem["$($f.itemId)"] = $f.status }

    # =========================================================================
    #  Infrastructure
    # =========================================================================
    Add-Title $shInf 'Infrastructure' "Organization: $org"

    $rows = foreach ($i in ((Items $Src.integrations) | Where-Object { "$(Get-P $_ 'integrationType')" -notmatch '^abx' })) {
        $props = Get-P $i 'integrationProperties'
        $addrKeys = 'hostName', 'hostname', 'url', 'endpoint', 'address', 'serverUrl', 'vroHostname', 'apiUrl', 'apiEndpoint'
        [ordered]@{
            'Type'                    = Get-P $i 'integrationType'
            'Name'                    = $i.name
            'Description'             = Get-P $i 'description'
            'Address/FQDN/Connection' = Get-P $props $addrKeys
            'Username'                = Get-P @($i, $props) 'privateKeyId', 'username', 'userName', 'user'
            'Password'                = '(not exported)'
            'Capability Tags'         = Get-P $i 'tags'
        }
    }
    Add-Vertical $Bk $shInf 'Integrations' $rows 'integrations'

    $rows = foreach ($id in $acctOrder) {
        $a = $acctById[$id]; $d = $acctDet[$id]; $both = @($d, $a)
        $cap = Get-P $both 'cloudAccountProperties'
        $assoc = @(@(Get-P $both 'associatedCloudAccountIds') + (LinkIds $a 'associated-cloud-accounts') | Where-Object { $_ } | Select-Object -Unique)
        $nsx = @($assoc | Where-Object { $acctType["$_"] -match 'nsx' } | ForEach-Object { & $AcctName $_ })
        $regs = Get-P $both 'enabledRegions', 'regions'
        $regTxt = if ($regs) { Join-Lines (@($regs) | ForEach-Object { Get-P $_ 'name', 'externalRegionId' }) } else { Join-Lines (@(Get-P $both 'enabledRegionIds') | ForEach-Object { if ($regionById.ContainsKey("$_")) { $regionById["$_"].Label } else { $_ } }) }
        [ordered]@{
            'Type'                                    = $acctType[$id]
            'Name'                                    = $a.name
            'Description'                             = Get-P $both 'description'
            'IP Address / FQDN'                       = Get-P @($d, $cap, $a) 'hostName', 'hostname'
            'Username'                                = Get-P @($d, $cap, $a) 'username', 'privateKeyId'
            'Password'                                = '(not exported)'
            'Allow provisioning in these datacenters' = $regTxt
            'NSX Manager'                             = Join-Lines $nsx
            'NSX Manager Mode'                        = Get-P @($d, $cap) 'managerMode', 'isGlobalManager'
            'Accept self-signed certificate'          = Get-P @($cap, $d) 'acceptSelfSignedCertificate'
            'Capability Tags'                         = Get-P $both 'tags'
            'Health'                                  = Get-P $both 'healthy', 'status'
            'Created'                                 = Get-P $both 'createdAt'
        }
    }
    Add-Vertical $Bk $shInf 'Cloud Accounts' $rows 'cloudAccounts', 'cloudAccountsVsphere'

    $rows = foreach ($id in $acctOrder) {
        $a = $acctById[$id]
        foreach ($x in @(@(Get-P @($acctDet[$id], $a) 'associatedCloudAccountIds') + (LinkIds $a 'associated-cloud-accounts') | Where-Object { $_ } | Select-Object -Unique)) {
            [ordered]@{ 'Cloud Account' = $a.name; 'Associated Account' = & $AcctName $x; 'Associated Type' = $acctType["$x"]; 'Identifier' = $x }
        }
    }
    Add-Table $Bk $shInf 'Cloud Accounts - Site Associations' $rows 'cloudAccounts'

    $rows = foreach ($z in $zones) {
        $ttm = Get-P $z 'tagsToMatch'
        [ordered]@{
            'Account/Region'    = RegionOf $z
            'Name'              = $z.name
            'Description'       = Get-P $z 'description'
            'Placement Policy'  = Get-P $z 'placementPolicy'
            'Folder'            = Get-P $z 'folder'
            'Capability Tags'   = Get-P $z 'tags'
            'Compute selection' = if ($ttm) { "Dynamic - tags: $(Format-Value $ttm)" } else { 'Manual / all' }
            'Computes'          = Join-Lines (@($zoneComp["$($z.id)"]) | ForEach-Object { $_.name })
            'Projects'          = Join-Lines $zoneProjects["$($z.id)"]
            'Custom Properties' = Get-P $z 'customProperties'
        }
    }
    Add-Vertical $Bk $shInf 'Cloud Zones' $rows 'zones'

    $rows = foreach ($z in $zones) {
        foreach ($c in (Items $zoneComp["$($z.id)"])) {
            [ordered]@{ 'Cloud Zone' = $z.name; 'Compute' = $c.name; 'Account/Region' = RegionOf $c; 'Type' = Get-P $c 'type'; 'Tags' = Get-P $c 'tags' }
        }
    }
    Add-Table $Bk $shInf 'Cloud Zones - Compute' $rows 'zoneComputes'

    # Flavor mapping matrix
    $cols = [System.Collections.Generic.List[string]]::new(); $mx = [ordered]@{}
    foreach ($p in (Items $Src.flavorProfiles)) {
        $rl = RegionOf $p; if (-not $cols.Contains($rl)) { $cols.Add($rl) }
        $map = Get-P $p 'flavorMappings'; $inner = Get-P $map 'mapping'; if ($inner) { $map = $inner }
        foreach ($m in (Get-Members $map)) {
            $f = $m.Value
            $mem = Get-P $f 'memoryInMB'; $cps = Get-P $f 'coresPerSocket'
            $txt = '{0} vCPU / {1} GB' -f (Get-P $f 'cpuCount'), $(if ($mem) { [math]::Round([double]$mem / 1024, 2) } else { '?' })
            if ($cps) { $txt += " / $cps cores per socket" }
            if (-not $mx.Contains($m.Name)) { $mx[$m.Name] = @{} }
            $mx[$m.Name][$rl] = $txt
        }
    }
    $rows = foreach ($n in $mx.Keys) { $r = [ordered]@{ 'Flavor Name' = $n }; foreach ($c in $cols) { $r[$c] = $mx[$n][$c] }; $r }
    Add-Table $Bk $shInf 'Flavor Mapping (flavor x account/region)' $rows 'flavorProfiles' (@('Flavor Name') + $cols)

    # Image mapping matrix + details
    $cols = [System.Collections.Generic.List[string]]::new(); $mx = [ordered]@{}; $detail = [System.Collections.Generic.List[object]]::new()
    foreach ($p in (Items $Src.imageProfiles)) {
        $rl = RegionOf $p; if (-not $cols.Contains($rl)) { $cols.Add($rl) }
        $map = Get-P $p 'imageMappings'; $inner = Get-P $map 'mapping'; if ($inner) { $map = $inner }
        foreach ($m in (Get-Members $map)) {
            $im = $m.Value
            $img = Get-P $im 'name', 'externalId', 'image', 'id'
            if (-not $mx.Contains($m.Name)) { $mx[$m.Name] = @{} }
            $mx[$m.Name][$rl] = $img
            $detail.Add([ordered]@{
                'Image Name' = $m.Name; 'Account/Region' = $rl; 'Image' = $img; 'OS Family' = Get-P $im 'osFamily'
                'Constraints' = Get-P $im 'constraints'; 'Cloud Configuration' = Get-P $im 'cloudConfig'
            })
        }
    }
    $rows = foreach ($n in $mx.Keys) { $r = [ordered]@{ 'Image Name' = $n }; foreach ($c in $cols) { $r[$c] = $mx[$n][$c] }; $r }
    Add-Table $Bk $shInf 'Image Mapping (image x account/region)' $rows 'imageProfiles' (@('Image Name') + $cols)
    Add-Table $Bk $shInf 'Image Mapping - Details' $detail 'imageProfiles'

    $rows = foreach ($s in (Items $Src.storageProfilesVsphere)) {
        $dsId = Get-P $s 'datastoreId', 'diskTargetProperties.datastoreId'; if (-not $dsId) { $dsId = (LinkIds $s 'datastore') | Select-Object -First 1 }
        $spId = Get-P $s 'storagePolicyId', 'diskTargetProperties.storagePolicyId'; if (-not $spId) { $spId = (LinkIds $s 'storage-policy') | Select-Object -First 1 }
        $lvl = Get-P $s 'sharesLevel', 'diskProperties.sharesLevel'; $sh = Get-P $s 'shares', 'diskProperties.shares'
        [ordered]@{
            'Account/Region'                    = RegionOf $s
            'Name'                              = $s.name
            'Description'                       = Get-P $s 'description'
            'Disk type'                         = Get-P $s 'diskType', 'diskProperties.diskType'
            'Storage policy'                    = if ($spId -and $spolById.ContainsKey("$spId")) { $spolById["$spId"].name } elseif ($spId) { $spId } else { 'Datastore default' }
            'Datastore/cluster'                 = if ($dsId -and $dsById.ContainsKey("$dsId")) { $dsById["$dsId"].name } else { $dsId }
            'Provisioning type'                 = Get-P $s 'provisioningType', 'diskProperties.provisioningType'
            'Shares'                            = (Join-Lines @($lvl, $sh)) -replace "`n", ' / '
            'Limit IOPS'                        = Get-P $s 'limitIops', 'diskProperties.limitIops'
            'Disk mode'                         = Get-P $s 'diskMode', 'diskProperties.diskMode'
            'Supports encryption'               = Get-P $s 'supportsEncryption'
            'Preferred storage for this region' = Get-P $s 'defaultItem'
            'Capability Tags'                   = Get-P $s 'tags'
        }
    }
    Add-Vertical $Bk $shInf 'Storage Profiles' $rows 'storageProfilesVsphere'

    $rows = foreach ($c in (Items $Src.fabricComputes)) {
        [ordered]@{ 'Name' = $c.name; 'Account/Region' = RegionOf $c; 'Type' = Get-P $c 'type'; 'Cloud Zones' = Join-Lines $compZones["$($c.id)"]; 'Power/Lifecycle' = Get-P $c 'powerState', 'lifecycleState'; 'Tags' = Get-P $c 'tags' }
    }
    Add-Table $Bk $shInf 'Compute' $rows 'fabricComputes'

    $rows = foreach ($n in $netById.Values) {
        [ordered]@{
            'Name' = $n.name; 'Account/Region' = RegionOf $n; 'Network Domain' = NetDomain $n
            'IPv4 CIDR' = NetP $n 'cidr'; 'IPv4 Default Gateway' = NetP $n 'gateway', 'defaultGateway'
            'IPv6 CIDR' = NetP $n 'ipv6Cidr'; 'IPv6 Default Gateway' = NetP $n 'defaultIpv6Gateway'
            'Domain' = NetP $n 'domain'; 'DNS Servers' = NetP $n 'dnsServerAddresses'; 'DNS Search Domains' = NetP $n 'dnsSearchDomains'
            'Support Public IP' = NetP $n 'isPublic'; 'Default for Zone' = NetP $n 'isDefault'
            'Network Profiles' = Join-Lines $netInProfiles["$($n.id)"]; 'Tags' = NetP $n 'tags'
        }
    }
    $rows = @($rows | Sort-Object { $_['Account/Region'] }, { $_['Name'] })
    $netTitle = 'Networks'
    if ($OnlyUsedNetworks) {
        $inRange = @{}
        foreach ($ir in (Items $Src.ipRanges) + (Items $Src.externalIpRanges)) {
            foreach ($nid in @(@(Get-P $ir 'fabricNetworkIds') + (LinkIds $ir 'fabric-networks') | Where-Object { $_ })) { $inRange["$nid"] = $true }
        }
        $usedNames = @{}
        foreach ($n in $netById.Values) { if ($netInProfiles.ContainsKey("$($n.id)") -or (NetP $n 'cidr') -or $inRange.ContainsKey("$($n.id)")) { $usedNames["$($n.name)"] = $true } }
        $total = $rows.Count
        $rows = @($rows | Where-Object { $usedNames.ContainsKey("$($_['Name'])") })
        $netTitle = "Networks (used only: $($rows.Count) of $total)"
    }
    Add-Table $Bk $shInf $netTitle $rows 'fabricNetworks', 'fabricNetworksVsphere'

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($ir in (Items $Src.ipRanges) + (Items $Src.externalIpRanges)) {
        $isExt = -not (Get-P $ir 'startIPAddress')
        $nids = @(@(Get-P $ir 'fabricNetworkIds', 'fabricNetworkId') + (LinkIds $ir 'fabric-networks') + (LinkIds $ir 'fabric-network') | Where-Object { $_ } | Select-Object -Unique)
        if ($nids.Count -eq 0) { $nids = @($null) }
        foreach ($nid in $nids) {
            $n = if ($nid -and $netById.ContainsKey("$nid")) { $netById["$nid"] } else { $null }
            $rows.Add([ordered]@{
                'Name' = $ir.name; 'Description' = Get-P $ir 'description'
                'Network' = if ($n) { $n.name } else { $nid }; 'Account/Region' = if ($n) { RegionOf $n } else { RegionOf $ir }
                'Network Domain' = if ($n) { NetDomain $n } else { '' }; 'Domain' = if ($n) { NetP $n 'domain' } else { '' }
                'IPv4 CIDR' = if ($n) { NetP $n 'cidr' } else { Get-P $ir 'cidr' }; 'IPv4 Default Gateway' = if ($n) { NetP $n 'gateway' } else { '' }
                'IPv6 CIDR' = if ($n) { NetP $n 'ipv6Cidr' } else { '' }; 'IPv6 Default Gateway' = if ($n) { NetP $n 'defaultIpv6Gateway' } else { '' }
                'DNS Servers' = if ($n) { NetP $n 'dnsServerAddresses' } else { '' }; 'DNS Search domains' = if ($n) { NetP $n 'dnsSearchDomains' } else { '' }
                'Support Public IP' = if ($n) { NetP $n 'isPublic' } else { '' }; 'Default for Zone' = if ($n) { NetP $n 'isDefault' } else { '' }
                'Origin' = if ($n) { NetP $n 'origin' } else { '' }; 'Network Tags' = if ($n) { NetP $n 'tags' } else { '' }
                'Provider' = if ($isExt) { 'External IPAM' } else { 'Internal' }
                'Start IP Address' = Get-P $ir 'startIPAddress'; 'End IP Address' = Get-P $ir 'endIPAddress'
                'Range Tags' = Get-P $ir 'tags'
            })
        }
    }
    Add-Table $Bk $shInf 'Networks - IP Ranges' $rows 'ipRanges', 'externalIpRanges'

    $rows = foreach ($l in (Items $Src.loadBalancers)) { [ordered]@{ 'Name' = $l.name; 'Account/Region' = RegionOf $l; 'Address' = Get-P $l 'address'; 'Project' = PName (Get-P $l 'projectId'); 'Tags' = Get-P $l 'tags' } }
    Add-Table $Bk $shInf 'Networks - Load Balancers' $rows 'loadBalancers'

    $rows = foreach ($d in (Items $Src.networkDomains)) { [ordered]@{ 'Name' = $d.name; 'Account/Region' = RegionOf $d; 'CIDR' = Get-P $d 'cidr'; 'Tags' = Get-P $d 'tags' } }
    Add-Table $Bk $shInf 'Networks - Network Domains' $rows 'networkDomains'

    $rows = foreach ($s in (Items $Src.securityGroups)) { [ordered]@{ 'Name' = $s.name; 'Description' = Get-P $s 'description'; 'Account/Region' = RegionOf $s; 'Rules' = (Items (Get-P $s 'rules')).Count; 'Tags' = Get-P $s 'tags' } }
    Add-Table $Bk $shInf 'Security Groups (discovered)' $rows 'securityGroups'

    $rows = foreach ($d in (Items $Src.datastores)) { [ordered]@{ 'Name' = $d.name; 'Account/Region' = RegionOf $d; 'Type' = Get-P $d 'type'; 'Free (GB)' = Get-P $d 'freeSizeGB'; 'Tags' = Get-P $d 'tags' } }
    Add-Table $Bk $shInf 'Storage Resources - Datastores (discovered)' $rows 'datastores'

    $rows = foreach ($c in (Items $Src.pricingCards)) { [ordered]@{ 'Name' = $c.name; 'Description' = Get-P $c 'description'; 'Charge Model' = Get-P $c 'chargeModel'; 'Created By' = Get-P $c 'createdBy' } }
    Add-Table $Bk $shInf 'Pricing Cards' $rows 'pricingCards'


    # =========================================================================
    #  Administration
    # =========================================================================
    Add-Title $shAdm 'Administration' "Organization: $org"

    $rows = foreach ($p in $projects) {
        $ps = $psById["$($p.id)"]; $both = @($ps, $p)
        $t = Get-P $p 'operationTimeout'
        [ordered]@{
            'Name'                                   = $p.name
            'Description'                            = Get-P $both 'description'
            'Deployment sharing'                     = Get-P $both 'sharedResources'
            'Placement Policy'                       = Get-P $p 'placementPolicy'
            'Request Timeout (s)'                    = if ($t) { $t } else { 'blank (default)' }
            'Allow Terraform cloud zone mapping'     = if ((Get-P $p 'properties.__allowTerraformCloudzoneMapping') -eq 'true') { 'Enabled' } else { 'Disabled' }
            'Cloud Zones'                            = Join-Lines ((ZoneAssign $p) | ForEach-Object { $zid = "$(Get-P $_ 'zoneId')"; if ($zoneById.ContainsKey($zid)) { $zoneById[$zid].name } else { $zid } })
            'Administrators'                         = Join-Lines (@(Get-P $both 'administrators') | ForEach-Object { Get-P $_ 'email', 'name' })
            'Members'                                = Join-Lines (@(Get-P $both 'members') | ForEach-Object { Get-P $_ 'email', 'name' })
            'Viewers'                                = Join-Lines (@(Get-P $both 'viewers') | ForEach-Object { Get-P $_ 'email', 'name' })
            'Supervisors'                            = Join-Lines (@(Get-P $both 'supervisors') | ForEach-Object { Get-P $_ 'email', 'name' })
            'Custom Properties (count)'              = (Get-Members (Get-P $p 'customProperties')).Count
        }
    }
    Add-Vertical $Bk $shAdm 'Projects' $rows 'projects'

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $projects) {
        $both = @($psById["$($p.id)"], $p); $users = [ordered]@{}
        foreach ($role in 'administrators', 'members', 'viewers', 'supervisors') {
            foreach ($u in (Items (Get-P $both $role))) {
                $nm = Get-P $u 'email', 'name'
                if (-not $users.Contains("$nm")) { $users["$nm"] = [ordered]@{ 'Project' = $p.name; 'Type' = Get-P $u 'type'; 'Name' = $nm; 'Administrator' = ''; 'Member' = ''; 'Viewer' = ''; 'Supervisor' = '' } }
                $users["$nm"][(Get-Culture).TextInfo.ToTitleCase($role.TrimEnd('s'))] = 'Yes'
            }
        }
        foreach ($v in $users.Values) { $rows.Add($v) }
    }
    Add-Table $Bk $shAdm 'Projects - Users & Groups' $rows 'projects', 'projectService'

    $rows = foreach ($p in $projects) {
        foreach ($za in (ZoneAssign $p)) {
            $zid = "$(Get-P $za 'zoneId')"; $z = $zoneById[$zid]
            [ordered]@{
                'Project' = $p.name; 'Cloud Zone' = if ($z) { $z.name } else { $zid }; 'Account/Region' = if ($z) { RegionOf $z } else { '' }
                'Priority' = Get-P $za 'priority'; 'Instances (0 = unlimited)' = Get-P $za 'maxNumberInstances'
                'Memory Limit (MB)' = Get-P $za 'memoryLimitMB'; 'CPU Limit' = Get-P $za 'cpuLimit'; 'Storage Limit (GB)' = Get-P $za 'storageLimitGB'
                'Capability Tags' = if ($z) { Get-P $z 'tags' } else { '' }
            }
        }
    }
    Add-Table $Bk $shAdm 'Projects - Provisioning - Cloud Zones' $rows 'projects'

    $rows = foreach ($p in $projects) {
        foreach ($kind in 'network', 'storage', 'extensibility') {
            foreach ($c in (Items (Get-P $p "constraints.$kind"))) {
                [ordered]@{ 'Project' = $p.name; 'Constraint Type' = $kind; 'Expression' = Get-P $c 'expression'; 'Enforcement' = if (Get-P $c 'mandatory') { 'Hard' } else { 'Soft' } }
            }
        }
    }
    Add-Table $Bk $shAdm 'Projects - Provisioning - Constraints' $rows 'projects'

    $rows = foreach ($p in $projects) {
        foreach ($m in (Get-Members (Get-P $p 'customProperties'))) { [ordered]@{ 'Project' = $p.name; 'Name' = $m.Name; 'Value' = $m.Value } }
    }
    Add-Table $Bk $shAdm 'Projects - Provisioning - Custom Properties' $rows 'projects'

    $rows = foreach ($r in (Items $Src.orgRoles)) {
        [ordered]@{ 'Name' = $r.name; 'Description' = Get-P $r 'description'; 'Predefined' = Get-P $r 'readOnly'; 'Bundle Key' = Get-P $r 'bundleKey' }
    }
    Add-Table $Bk $shAdm 'Roles (organization)' $rows 'orgRoles'

    $rows = foreach ($u in (Items $Src.orgUsers)) {
        [ordered]@{
            'User' = Get-P $u 'username', 'name'; 'Full Name' = Get-P $u 'fullName'; 'Provider' = Get-P $u 'providerType'
            'Roles' = Join-Lines (@(Get-P $u 'roleEntityRefs') | ForEach-Object { Get-P $_ 'name' })
            'Enabled' = Get-P $u 'enabled'; 'Email' = Get-P $u 'email'
        }
    }
    Add-Table $Bk $shAdm 'Users (organization)' $rows 'orgUsers'

    $rows = foreach ($g in (Items $Src.orgGroups)) {
        [ordered]@{
            'Group' = Get-P $g 'name'; 'Description' = Get-P $g 'description'; 'Provider' = Get-P $g 'providerType'
            'Roles' = Join-Lines (@(Get-P $g 'roleEntityRefs', 'roleEntityRef') | ForEach-Object { Get-P $_ 'name' })
        }
    }
    Add-Table $Bk $shAdm 'Groups (organization)' $rows 'orgGroups'

    $rows = foreach ($n in (Items $Src.naming)) {
        $scope = Join-Lines (@(Get-P $n 'projects') | ForEach-Object {
                if (Get-P $_ 'orgDefault', 'defaultOrg') { 'Organization default' } else { Get-P $_ 'projectName', 'projectId' } })
        foreach ($t in (Items (Get-P $n 'templates'))) {
            [ordered]@{
                'Custom Name' = $n.name; 'Description' = Get-P $n 'description'; 'Scope' = $scope
                'Resource Type' = Get-P $t 'resourceTypeName', 'resourceType'; 'Template Name' = Get-P $t 'name'
                'Template Format' = Get-P $t 'pattern'; 'Starting Counter Value' = Get-P $t 'startCounter'; 'Increment Step' = Get-P $t 'incrementStep'
                'Current Counter' = Join-Lines (@(Get-P $t 'counters') | ForEach-Object { Get-P $_ 'currentCounter' })
                'Unique Name' = Get-P $t 'uniqueName'; 'Static Pattern' = Get-P $t 'staticPattern'
            }
        }
    }
    Add-Table $Bk $shAdm 'Custom Names' $rows 'naming'

    $rows = foreach ($s in (Items $Src.secrets)) {
        [ordered]@{ 'Name' = $s.name; 'Scope' = if (Get-P $s 'orgScoped') { 'Organization' } else { PName (Get-P $s 'projectId') }; 'Value' = '(hidden)'; 'Description' = Get-P $s 'description'; 'Updated' = Get-P $s 'updatedAt', 'createdAt' }
    }
    Add-Table $Bk $shAdm 'Secrets' $rows 'secrets'

    $rows = foreach ($c in (Items $Src.configProperties)) { [ordered]@{ 'Name' = Get-P $c 'key', 'name'; 'Value' = Get-P $c 'value'; 'Description' = Get-P $c 'description' } }
    Add-Table $Bk $shAdm 'Settings' $rows 'configProperties'

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
    Add-Vertical $Bk $shAdm 'Identity Providers - OIDC' $rows 'identityProviders'
    if (IdpNote 'oauth') { [void](Add-Row $shAdm @(C (IdpNote 'oauth') 6)); Add-Blank $shAdm }

    $l = Get-P $idp['ldap'] 'settings'
    $mode = Get-P $l 'orgLdapMode'
    $rows = if ($l -and $mode -and $mode -ne 'NONE') {
        $cl = Get-P $l 'customOrgLdapSettings'
        $d = [ordered]@{
            'Name' = 'LDAP'; 'Mode' = $mode; 'Host' = Get-P $cl 'hostName'; 'Port' = Get-P $cl 'port'; 'SSL' = Get-P $cl 'isSsl'
            'Search base' = Get-P $cl 'searchBase'; 'Group search base' = Get-P $cl 'groupSearchBase'; 'Bind user' = Get-P $cl 'userName'
            'Authentication' = Get-P $cl 'authenticationMechanism'; 'Connector type' = Get-P $cl 'connectorType'
            'Custom users OU' = Get-P $l 'customUsersOu'
        }
        $d = Add-RestOf $d $cl @('hostName', 'port', 'isSsl', 'searchBase', 'groupSearchBase', 'userName', 'authenticationMechanism', 'connectorType')
        Add-RestOf $d $l @('orgLdapMode', 'customOrgLdapSettings', 'customUsersOu')
    }
    Add-Vertical $Bk $shAdm 'Identity Providers - LDAP' $rows 'identityProviders'
    if (IdpNote 'ldap') { [void](Add-Row $shAdm @(C (IdpNote 'ldap') 6)); Add-Blank $shAdm }

    $f = Get-P $idp['federation'] 'settings'
    $rows = if ($f -and ((Get-P $f 'enabled') -or (Get-P $f 'samlMetadata'))) {
        Add-RestOf ([ordered]@{
            'Name' = 'SAML'; 'Enabled' = Get-P $f 'enabled'; 'Entity ID' = Get-P $f 'samlSPEntityId'
            'Certificate expiration' = Get-P $f 'certificateExpiration'; 'Attribute mapping' = Get-P $f 'samlAttributeMapping'
        }) $f @('enabled', 'samlSPEntityId', 'certificateExpiration', 'samlAttributeMapping')
    }
    Add-Vertical $Bk $shAdm 'Identity Providers - SAML' $rows 'identityProviders'
    if (IdpNote 'federation') { [void](Add-Row $shAdm @(C (IdpNote 'federation') 6)); Add-Blank $shAdm }

    # =========================================================================
    #  Network Profiles
    # =========================================================================
    Add-Title $shNp 'Network Profiles' "Organization: $org"
    $rows = foreach ($np in $netProfiles) {
        $nids = ProfNets $np
        $sgs = @(@(Get-P $np 'securityGroupIds') + (LinkIds $np 'security-groups') | Where-Object { $_ } | Select-Object -Unique)
        [ordered]@{
            'Account / region'  = RegionOf $np
            'Name'              = $np.name
            'Description'       = Get-P $np 'description'
            'Capability tags'   = Get-P $np 'tags'
            'Networks'          = Join-Lines ($nids | ForEach-Object { if ($netById.ContainsKey("$_")) { $netById["$_"].name } else { $_ } })
            'Isolation policy'  = Get-P $np 'isolationType'
            'Isolation network domain / CIDR' = Join-Lines @((Get-P $np 'isolationNetworkDomainCIDR'), (Get-P $np 'isolatedNetworkCIDRPrefix'))
            'Security groups'   = Join-Lines ($sgs | ForEach-Object { $x = "$_"; $m = (Items $Src.securityGroups) | Where-Object { "$($_.id)" -eq $x } | Select-Object -First 1; if ($m) { $m.name } else { $x } })
            'Custom Properties' = Get-P $np 'customProperties'
        }
    }
    Add-Vertical $Bk $shNp 'Network Profiles' $rows 'networkProfiles'

    $rows = foreach ($np in $netProfiles) {
        $nids = ProfNets $np
        foreach ($nid in $nids) {
            $n = $netById["$nid"]
            if (-not $n) { [ordered]@{ 'Network Profile' = $np.name; 'Name' = $nid }; continue }
            $rl = RegionOf $n
            [ordered]@{
                'Network Profile' = $np.name; 'Name' = $n.name; 'Account / Region' = $rl
                'Zone' = Join-Lines $zonesByRegion[$rl]
                'Network Domain' = NetDomain $n; 'CIDR' = NetP $n 'cidr'; 'NSX Project / VPC' = VpcInfo $n
                'Support Public IP' = NetP $n 'isPublic'; 'Default for Zone' = NetP $n 'isDefault'; 'Tags' = NetP $n 'tags'
            }
        }
    }
    Add-Table $Bk $shNp 'Network Profiles - Networks' $rows 'networkProfiles'

    # =========================================================================
    #  Tags
    # =========================================================================
    Add-Title $shTag 'Tags' "Organization: $org  -  capability tags (on zones / profiles / accounts) and resource tags (discovered or user-added on compute, networks, storage)"
    $tagUse = [ordered]@{}
    function Add-TagUse($tags, [string]$Kind, [string]$Where) {
        foreach ($t in (Items $tags)) {
            $k = "$(Get-P $t 'key')"; if (-not $k) { continue }
            $v = "$(Get-P $t 'value')"; $id = "${k}:$v"
            if (-not $tagUse.Contains($id)) { $tagUse[$id] = @{ Key = $k; Value = $v; Kinds = [System.Collections.Generic.List[string]]::new(); Where = [System.Collections.Generic.List[string]]::new(); Refs = [System.Collections.Generic.List[string]]::new() } }
            if (-not $tagUse[$id].Kinds.Contains($Kind)) { $tagUse[$id].Kinds.Add($Kind) }
            $tagUse[$id].Where.Add($Where)
        }
    }
    foreach ($t in (Items $Src.tags)) { Add-TagUse @($t) 'Defined' $null }
    foreach ($a in $acctById.Values) { Add-TagUse (Get-P $a 'tags') 'Capability' "Cloud account: $($a.name)" }
    foreach ($z in $zones) { Add-TagUse (Get-P $z 'tags') 'Capability' "Cloud zone: $($z.name)" }
    foreach ($x in (Items $Src.storageProfilesVsphere)) { Add-TagUse (Get-P $x 'tags') 'Capability' "Storage profile: $($x.name)" }
    foreach ($x in $netProfiles) { Add-TagUse (Get-P $x 'tags') 'Capability' "Network profile: $($x.name)" }
    foreach ($x in (Items $Src.integrations)) { Add-TagUse (Get-P $x 'tags') 'Capability' "Integration: $($x.name)" }
    foreach ($x in (Items $Src.fabricComputes)) { Add-TagUse (Get-P $x 'tags') 'Resource' "Compute: $($x.name)" }
    foreach ($x in $netById.Values) { Add-TagUse (NetP $x 'tags') 'Resource' "Network: $($x.name)" }
    foreach ($x in (Items $Src.datastores)) { Add-TagUse (Get-P $x 'tags') 'Resource' "Datastore: $($x.name)" }
    foreach ($x in (Items $Src.securityGroups)) { Add-TagUse (Get-P $x 'tags') 'Resource' "Security group: $($x.name)" }
    foreach ($x in (Items $Src.ipRanges)) { Add-TagUse (Get-P $x 'tags') 'Resource' "IP range: $($x.name)" }
    # where tags are consumed: project constraints, zone compute selection, image constraints
    $refs = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $projects) { foreach ($kind in 'network', 'storage', 'extensibility') { foreach ($c in (Items (Get-P $p "constraints.$kind"))) { $refs.Add(@{ Expr = "$(Get-P $c 'expression')"; Where = "Project $($p.name) ($kind constraint)" }) } } }
    foreach ($z in $zones) { foreach ($t in (Items (Get-P $z 'tagsToMatch'))) { $refs.Add(@{ Expr = "$(Get-P $t 'key'):$(Get-P $t 'value')"; Where = "Cloud zone $($z.name) (compute selection)" }) } }
    foreach ($ip in (Items $Src.imageProfiles)) {
        $map = Get-P $ip 'imageMappings'; $inner = Get-P $map 'mapping'; if ($inner) { $map = $inner }
        foreach ($m in (Get-Members $map)) { foreach ($c in (Items (Get-P $m.Value 'constraints'))) { $refs.Add(@{ Expr = "$(Get-P $c 'expression')"; Where = "Image $($m.Name) (constraint)" }) } }
    }
    foreach ($r in $refs) {
        $e = ($r.Expr -replace '^!', '').Trim(); if (-not $e) { continue }
        $parts = $e -split ':', 2; $k = $parts[0]; $v = if ($parts.Count -gt 1) { $parts[1] } else { '' }
        $matched = $false
        foreach ($id in @($tagUse.Keys)) { $u = $tagUse[$id]; if ($u.Key -eq $k -and ($v -eq '' -or $u.Value -eq $v)) { $u.Refs.Add($r.Where); $matched = $true } }
        if (-not $matched) {
            $id = "${k}:$v"
            $tagUse[$id] = @{ Key = $k; Value = $v; Kinds = [System.Collections.Generic.List[string]]::new(); Where = [System.Collections.Generic.List[string]]::new(); Refs = [System.Collections.Generic.List[string]]::new() }
            $tagUse[$id].Refs.Add($r.Where)
        }
    }
    function Cap($list) { return (Join-Lines @($list | Where-Object { $_ })) }
    $rows = foreach ($id in $tagUse.Keys) {
        $u = $tagUse[$id]
        $kinds = @($u.Kinds | Where-Object { $_ -ne 'Defined' })
        $type = if ($kinds -contains 'Capability' -and $kinds -contains 'Resource') { 'Capability + Resource' } elseif ($kinds.Count) { $kinds[0] } elseif ($u.Refs.Count) { 'Constraint only (no matching tag found)' } else { 'Defined, not in use' }
        [ordered]@{
            'Key' = $u.Key; 'Value' = $u.Value; 'Tag Type' = $type
            'Used On (count)' = @($u.Where | Where-Object { $_ }).Count; 'Used On' = Cap $u.Where
            'Referenced By' = Cap $u.Refs
        }
    }
    $rows = @($rows | Sort-Object { $_['Key'] }, { $_['Value'] })
    Add-Table $Bk $shTag 'All Tags' $rows 'tags'

    # =========================================================================
    #  Onboarding
    # =========================================================================
    Add-Title $shOnb 'Onboarding' "Organization: $org"
    $plans = Items $Src.onboardingPlans
    $odeps = Items $Src.onboardingDeployments
    $ores = Items $Src.onboardingResources
    $planName = @{}; foreach ($p in $plans) { $planName["$(Get-P $p 'documentSelfLink')"] = "$(Get-P $p 'name')" }
    $depName = @{}; foreach ($d in $odeps) { $depName["$(Get-P $d 'documentSelfLink')"] = "$(Get-P $d 'name')" }
    $resByPlan = @{}; $resByDep = @{}; $depByPlan = @{}
    foreach ($r in $ores) { $k = "$(Get-P $r 'planLink')"; $resByPlan[$k] = 1 + [int]$resByPlan[$k]; $k = "$(Get-P $r 'deploymentLink')"; $resByDep[$k] = 1 + [int]$resByDep[$k] }
    foreach ($d in $odeps) { $k = "$(Get-P $d 'planLink')"; $depByPlan[$k] = 1 + [int]$depByPlan[$k] }
    function LinkTail($x) { if ($x) { return (("$x" -split '/')[-1]) } return $null }
    $rows = foreach ($p in $plans) {
        $self = "$(Get-P $p 'documentSelfLink')"
        $eps = @(@(Get-P $p 'endpointIds') + @(Get-P $p 'endpointId') + @(Get-P $p 'endpointLink') | Where-Object { $_ } | ForEach-Object { $tid = LinkTail $_; if ($acctById.ContainsKey("$tid")) { $acctById["$tid"].name } else { $_ } })
        [ordered]@{
            'Name' = Get-P $p 'name'; 'Description' = Get-P $p 'description'; 'Cloud Account' = Join-Lines $eps
            'Project' = PName (Get-P $p 'projectId'); 'Deployment Tag Type' = Get-P $p 'deploymentTagType'
            'Status' = Get-P $p 'status', 'state', 'enabled'; 'Created By' = Get-P $p 'createdBy', 'ownerId'
            'Deployments' = [int]$depByPlan[$self]
            'Resources' = [int]$resByPlan[$self]
        } | ForEach-Object { Add-RestOf $_ $p @('name', 'description', 'endpointIds', 'endpointId', 'endpointLink', 'projectId', 'deploymentTagType', 'status', 'state', 'enabled', 'createdBy', 'ownerId') }
    }
    Add-Table $Bk $shOnb 'Onboarding Plans' $rows 'onboardingPlans'
    $rows = foreach ($d in $odeps) {
        $self = "$(Get-P $d 'documentSelfLink')"
        [ordered]@{
            'Plan' = $planName["$(Get-P $d 'planLink')"]; 'Deployment' = Get-P $d 'name'; 'Description' = Get-P $d 'description'
            'Owner' = Get-P $d 'owner', 'ownerId'; 'Status' = Get-P $d 'status', 'state'
            'Resources' = [int]$resByDep[$self]
        } | ForEach-Object { Add-RestOf $_ $d @('planLink', 'name', 'description', 'owner', 'ownerId', 'status', 'state') }
    }
    Add-Table $Bk $shOnb 'Onboarding Deployments' $rows 'onboardingDeployments'
    $rows = foreach ($r in $ores) {
        [ordered]@{
            'Plan' = $planName["$(Get-P $r 'planLink')"]; 'Deployment' = $depName["$(Get-P $r 'deploymentLink')"]
            'Resource' = Get-P $r 'resourceName', 'name'; 'Resource Link' = Get-P $r 'resourceLink'
        } | ForEach-Object { Add-RestOf $_ $r @('planLink', 'deploymentLink', 'resourceName', 'name', 'resourceLink') }
    }
    Add-Table $Bk $shOnb 'Onboarding Resources' $rows 'onboardingResources'

    # =========================================================================
    #  Content & Policies
    # =========================================================================
    Add-Title $shCp 'Content & Policies' "Organization: $org"

    $rows = foreach ($s in (Items $Src.contentSources)) {
        $tid = Get-P $s 'typeId'
        [ordered]@{
            'Type' = if ($ContentSourceTypes.ContainsKey("$tid")) { $ContentSourceTypes["$tid"] } else { $tid }
            'Name' = $s.name; 'Get templates from / Source project' = PName (Get-P $s 'config.sourceProjectId', 'projectId')
            'Items Found' = Get-P $s 'itemsFound'; 'Items Imported' = Get-P $s 'itemsImported'
            'Last Import' = Get-P $s 'lastImportCompletedAt'; 'Created By' = Get-P $s 'createdBy'
        }
    }
    Add-Table $Bk $shCp 'Content Sources' $rows 'contentSources'

    $rows = foreach ($i in (Items $Src.catalogItems)) {
        [ordered]@{
            'Name' = $i.name; 'Type' = Get-P $i 'type.name', 'type.id'; 'Content Source' = Get-P $i 'sourceName'
            'Source' = Get-P $i 'type.id'; 'Projects' = Join-Lines (@(Get-P $i 'projectIds') | ForEach-Object { PName $_ })
            'Custom Request Form' = if ($formByItem.ContainsKey("$($i.id)")) { $formByItem["$($i.id)"] } elseif (Get-P $i 'formId') { 'Customised' } else { '' }
            'Description' = Get-P $i 'description'; 'Last Updated' = Get-P $i 'lastUpdatedAt'
        }
    }
    Add-Table $Bk $shCp 'Content (Catalog Items)' $rows 'catalogItems'

    $rows = foreach ($b in (Items $Src.blueprints)) {
        [ordered]@{
            'Name' = $b.name; 'Project' = Get-P $b 'projectName'; 'Status' = Get-P $b 'status'
            'Shared with all projects in org' = Get-P $b 'requestScopeOrg'; 'Versions' = Get-P $b 'totalVersions'; 'Released Versions' = Get-P $b 'totalReleasedVersions'
            'Updated By' = Get-P $b 'updatedBy', 'createdBy'; 'Updated' = Get-P $b 'updatedAt', 'createdAt'
        }
    }
    Add-Table $Bk $shCp 'Cloud Templates' $rows 'blueprints'

    $pols = Items $Src.policies
    $typesSeen = @($PolicyTypes.Keys) + @($pols | ForEach-Object { Get-P $_ 'typeId' } | Where-Object { $PolicyTypes.Keys -notcontains $_ } | Select-Object -Unique)
    foreach ($tid in $typesSeen) {
        $title = if ($PolicyTypes.Contains("$tid")) { $PolicyTypes["$tid"] } else { "Policies - $tid" }
        $rows = foreach ($p in ($pols | Where-Object { (Get-P $_ 'typeId') -eq $tid })) {
            $def = Get-P $p 'definition'
            $summary = if ($tid -eq 'com.vmware.policy.catalog.entitlement') {
                Join-Lines (@(Get-P $def 'entitledUsers') | ForEach-Object {
                        $who = Join-Lines (@(Get-P $_ 'principals') | ForEach-Object { $ref = Get-P $_ 'referenceId'; if ($ref) { '{0}:{1}' -f (Get-P $_ 'type'), $ref } else { '{0} (all)' -f (Get-P $_ 'type') } })
                        $what = Join-Lines (@(Get-P $_ 'items') | ForEach-Object {
                                $iid = "$(Get-P $_ 'id')"
                                $nm = if ($srcById.ContainsKey($iid)) { "source: $($srcById[$iid].name)" } elseif ($itemById.ContainsKey($iid)) { "item: $($itemById[$iid].name)" } else { "$(Get-P $_ 'type') $iid" }
                                $nm })
                        "Users: $($who -replace "`n", ', ')`nContent: $($what -replace "`n", ', ')" })
            } else { Format-Value $def }
            [ordered]@{
                'Name' = $p.name; 'Description' = Get-P $p 'description'; 'Enforcement' = Get-P $p 'enforcementType'
                'Creator' = Get-P $p 'createdBy'; 'Creation Time' = Get-P $p 'createdAt'
                'Scope' = if (Get-P $p 'projectId') { "Project: $(PName (Get-P $p 'projectId'))" } else { 'Organization' }
                'Criteria (fields to check)' = if (Get-P $p 'criteria') { (Get-P $p 'criteria') | ConvertTo-Json -Depth 20 -Compress } else { '' }
                'Definition' = $summary
            }
        }
        Add-Table $Bk $shCp $title $rows 'policies'
    }

    $e = $Src.emailConfig
    $rows = if ($e -and (Get-P $e 'host', 'serverName')) {
        [ordered]@{
            'Name' = Get-P $e 'name'; 'Description' = Get-P $e 'description'; 'Server name' = Get-P $e 'host', 'serverName'
            'Sender name' = Get-P $e 'senderName'; 'Sender address' = Get-P $e 'sender', 'senderAddress'
            'Authentication' = Get-P $e 'authentication', 'authenticationRequired'; 'Username' = Get-P $e 'userName', 'username'
            'Connection security' = Get-P $e 'connectionSecurity'; 'Server port' = Get-P $e 'port'
            'Enable insecure connection upgrade with STARTTLS' = Get-P $e 'insecureConnectionUpgrade', 'allowInsecureUpgrade'
            'Trust certificates presented by the host' = Get-P $e 'trustHost', 'trustCert'
        }
    }
    Add-Vertical $Bk $shCp 'Notifications - Email Server' $rows 'emailConfig'

    $rows = foreach ($s in (Items $Src.notificationScenarios)) {
        [ordered]@{ 'Category' = Get-P $s 'scenarioCategory'; 'Scenario' = Get-P $s 'scenarioName', 'scenarioId'; 'Status' = if (Get-P $s 'enabled') { 'Enabled' } else { 'Disabled' } }
    }
    Add-Table $Bk $shCp 'Notifications - Scenarios' $rows 'notificationScenarios'

    # =========================================================================
    #  Orchestrator
    # =========================================================================
    Add-Title $shVro 'Orchestrator' "Organization: $org"
    $va = $Src.vroAbout
    $rows = if ($va) { [ordered]@{ 'Name' = 'Orchestrator'; 'Version' = Get-P $va 'version'; 'Build' = Get-P $va 'build-number', 'buildNumber'; 'API Version' = Get-P $va 'api-version', 'apiVersion' } }
    Add-Vertical $Bk $shVro 'About' $rows 'vroAbout'

    $groups = Items $Src.vroEndpoints
    if (-not (@($groups | ForEach-Object { "$(Get-P $_ 'fullType')" }) -match '^AD:')) {
        Add-Vertical $Bk $shVro 'Integrations - Active Directory' @() 'vroEndpoints'
    }
    foreach ($g in $groups) {
        $ft = "$(Get-P $g 'fullType')"
        $title = if ($VroTypeTitles.ContainsKey($ft)) { $VroTypeTitles[$ft] } else { $ft }
        Add-Vertical $Bk $shVro "Integrations - $title ($ft)" (Items (Get-P $g 'objects')) 'vroEndpoints' 'name'
    }

    # Plug-in (.dar) list is kept in the JSON only (not needed in the workbook)
    if ($false) { $rows = foreach ($p in (Items $Src.vroPlugins)) {
        Add-RestOf ([ordered]@{ 'moduleName' = Get-P $p 'moduleName', 'name'; 'description' = Get-P $p 'description', 'displayName'; 'version' = Get-P $p 'version'; 'enabled' = Get-P $p 'enabled' }) $p @()
    }
    Add-Table $Bk $shVro 'Plug-ins' $rows 'vroPlugins' }

    # =========================================================================
    #  Version, Collection Log, Index
    # =========================================================================
    Add-Title $shVer 'Version History'
    [void](Add-Row $shVer @((C 'Date' 3), (C 'Version' 3), (C 'Author' 3), (C 'Description' 3), (C 'Reviewers' 3)))
    [void](Add-Row $shVer @((C (Format-Value $Data.meta.collectedAt)), (C '0.1'), (C $Data.meta.collectedBy), (C "Auto-collected from $($Data.meta.server) (org $org) by script v$($Data.meta.scriptVersion)"), (C '')))

    Add-Title $shLog 'Collection Log' 'One line per API call group. FAILED usually means the endpoint does not exist in this version or the account lacks rights.'
    Add-Table $Bk $shLog 'API calls' (Items $Data.log)

    Add-Title $shIdx 'VCF Automation - VM Apps Organization Configuration'
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
while (-not $Organization) { $Organization = Read-Host 'VM Apps organization name' }
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
# One folder per org type + org, e.g. "VM Apps - my-vmapps-org" (All Apps / Provider collectors will use their own prefix)
$orgFolder = Join-Path $OutputFolder ('VM Apps - ' + ($Organization -replace '[\\/:*?"<>|]', '_').Trim())
if (-not (Test-Path -LiteralPath $orgFolder)) { New-Item -ItemType Directory -Path $orgFolder | Out-Null }
$base = Get-FreePath $orgFolder ("VCFA_{0}_VMApps_Config_{1}" -f $safeOrg, (Get-Date -Format 'yyyyMMdd')) @('.xlsx', '.json')

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

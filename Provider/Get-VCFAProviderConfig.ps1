#Requires -Version 5.1
<#
.SYNOPSIS
    Collects the provider (System organization) configuration of VCF Automation 9.x and writes
    a configuration workbook (.xlsx) plus a raw JSON dump for quick reference.

.DESCRIPTION
    Read-only. Uses GET calls only (plus the login POST). Prompts for anything not passed:
    VCF Automation FQDN, TLS check, and provider credentials.

    Covers the sheets of the Provider configuration workbook: Organizations (region quotas, zones,
    VM classes, storage classes, networking, IP usage, first user, settings), Networking (provider
    gateways, IP spaces, edge clusters), Regions, Supervisors, Content Libraries, Roles and
    Credentials, VCF Data Services, General Settings, Connections (vCenter, NSX), VCF Instances,
    Access Control, Identity Providers, OIDC Services, Certificate Management, Branding, Feature Flags.

    API names that changed between releases are tried in turn (/cloudapi/v1 for 9.1, /cloudapi/vcf
    for 9.0, and alternative resource names). Whatever is not found is listed on the Collection Log.

    Output (in a sub-folder "Provider - <name>" of -OutputFolder, default = current folder, where
    <name> is the first label of the VCF Automation FQDN):
        Provider - <name>\VCFA_<name>_Provider_Config_<yyyyMMdd>.xlsx
        Provider - <name>\VCFA_<name>_Provider_Config_<yyyyMMdd>.json
    If either file already exists, a suffix _v2, _v3 ... is added so nothing is overwritten.

    The workbook is written natively (Open XML), so no ImportExcel / Excel install is needed
    on the jump host.

    Passwords are never returned by the APIs; any field that looks like a password, secret,
    token or private key is redacted before the JSON is saved.

.PARAMETER Server
    VCF Automation FQDN, e.g. vcfa.corp.local

.PARAMETER Credential
    Provider user (System organization). "@System" is appended to the user name if it is not already there.

.PARAMETER UseApiToken
    Log in with a provider API token instead of user name / password.

.PARAMETER SkipCertificateCheck
    Skip TLS certificate validation (self-signed / untrusted CA). Asked interactively if omitted.

.PARAMETER MaxColumns
    Objects shown side by side (#1, #2 ...) up to this count; above it the section switches to a row-per-object table. Default 10.

.PARAMETER FromJson
    Re-build the workbook from a JSON file produced by an earlier run (no connection needed).

.EXAMPLE
    ./Get-VCFAProviderConfig.ps1
    Fully interactive.

.EXAMPLE
    ./Get-VCFAProviderConfig.ps1 -Server vcfa.corp.local -SkipCertificateCheck

.EXAMPLE
    ./Get-VCFAProviderConfig.ps1 -FromJson .\VCFA_vcfa_Provider_Config_20261102.json
#>
[CmdletBinding(DefaultParameterSetName = 'Collect')]
param(
    [Parameter(ParameterSetName = 'Collect')][string]$Server,
    [Parameter(ParameterSetName = 'Collect')][pscredential]$Credential,
    [Parameter(ParameterSetName = 'Collect')][switch]$UseApiToken,
    [Parameter(ParameterSetName = 'Collect')][switch]$SkipCertificateCheck,
    [Parameter(ParameterSetName = 'Collect')][string]$BaseUri,
    [Parameter(ParameterSetName = 'Render', Mandatory)][string]$FromJson,
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
    SkipCertificateCheck = $null     # $true / $false, or $null to be asked
    Username             = ''        # e.g. 'admin'  (@System is added automatically)
    Password             = ''        # Plain text - avoid on customer machines; prefer CredentialFile below
    CredentialFile       = ''        # Path to an encrypted credential file made with:
                                     #   Get-Credential | Export-Clixml -Path C:\path\vcfa-cred.xml
                                     # Only the same Windows user on the same machine can decrypt it.
}
# =====================================================================================

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
$Organization = 'System'   # provider login always uses the System organization
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
        if ($script:TenantCtx) { $hdr['X-VMWARE-VCLOUD-TENANT-CONTEXT'] = $script:TenantCtx }
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
            $sec = Read-Host -AsSecureString 'Provider API token'
            $script:ApiTokenPlain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
        }
        $r = Invoke-RestMethod @WebOpts -Method Post -Uri "$base/oauth/provider/token" -Body @{ grant_type = 'refresh_token'; refresh_token = $script:ApiTokenPlain } -ContentType 'application/x-www-form-urlencoded' -Headers @{ Accept = 'application/json' }
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
            $resp = Invoke-WebRequest @WebOpts -Method Post -Uri "$base/cloudapi/1.0.0/sessions/provider" -Headers @{ Accept = "application/json;version=$ver"; Authorization = "Basic $basic" }
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
#  Provider specific helpers
# =====================================================================================

function Get-CloudList([string]$Name) {
    # VCF Automation 9.1 publishes these APIs under /cloudapi/v1, 9.0 under /cloudapi/vcf.
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

function Get-CloudFirst([string[]]$Names) {
    # First of several candidate names that exists (API names differ between releases)
    $last = $null
    foreach ($n in $Names) {
        try { $r = Get-CloudList $n; $script:Resolved[$n] = $true; return , $r }
        catch { $last = $_; if ((Get-StatusCode $_) -ne 404) { throw } }
    }
    throw $last
}

function Get-FirstPath([string[]]$Paths, [switch]$Single) {
    $last = $null
    foreach ($p in $Paths) {
        try { if ($Single) { return (Invoke-Vcfa $p) } else { return , (Get-All $p) } }
        catch { $last = $_; if ((Get-StatusCode $_) -ne 404) { throw } }
    }
    throw $last
}

function Invoke-AsTenant([string]$OrgId, [scriptblock]$Script) {
    # Runs a block in an organization's context (X-VMWARE-VCLOUD-TENANT-CONTEXT)
    $script:TenantCtx = ("$OrgId" -split ':')[-1]
    try { return (& $Script) } finally { $script:TenantCtx = $null }
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
    $out = [ordered]@{}
    Add-Flat $out $Obj ''
    return $out
}

function Find-Deep {
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
    $f = Get-Flat $Obj
    return (Join-Lines @($f.Keys | Where-Object { $_ -match $Regex } | ForEach-Object { '{0}: {1}' -f $_, (Format-Value $f[$_]) }))
}

function Get-NameList($v) {
    return (Join-Lines @((Items $v) | ForEach-Object {
                if ($_ -is [string] -or $_ -is [ValueType]) { "$_" } else { Get-P $_ 'name', 'displayName', 'id' }
            }))
}

function RefName($v) {
    # entity reference {name, id} or plain value -> name
    if ($null -eq $v) { return '' }
    if ($v -is [string] -or $v -is [ValueType]) { return "$v" }
    if ($v -is [System.Collections.IDictionary] -or $v -is [System.Management.Automation.PSCustomObject]) { return "$(Get-P $v 'name', 'displayName', 'id')" }
    return (Get-NameList $v)
}

function Get-CertInfo([string]$Pem) {
    # Subject / issuer / expiry of the first certificate in a PEM string
    if (-not $Pem) { return $null }
    $m = [regex]::Match($Pem, '-----BEGIN CERTIFICATE-----(.+?)-----END CERTIFICATE-----', 'Singleline')
    $b64 = if ($m.Success) { $m.Groups[1].Value } else { $Pem }
    try {
        $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 (, [Convert]::FromBase64String(($b64 -replace '\s', '')))
        return @{
            Cn      = $c.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
            Issuer  = $c.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $true)
            Expires = $c.NotAfter.ToUniversalTime().ToString('yyyy-MM-dd HH:mm')
            Chain   = ([regex]::Matches($Pem, 'BEGIN CERTIFICATE')).Count
        }
    }
    catch { return $null }
}

# =====================================================================================
#  Collection
# =====================================================================================

$script:Resolved = @{}
$script:TenantCtx = $null

function Invoke-Collection {
    Write-Host "`nCollecting from $($script:Conn.Base) ..." -ForegroundColor Cyan

    # --- Organizations
    Collect 'orgs'                  '/cloudapi/1.0.0/orgs'
    $orgList = @((Items $script:Raw.orgs) | Where-Object { "$($_.name)" -ne 'System' })
    Collect 'orgUsers'              '/cloudapi/1.0.0/users (per organization)' -Script {
        foreach ($o in $orgList) {
            try { $u = Invoke-AsTenant $o.id { Get-All '/cloudapi/1.0.0/users' }; [pscustomobject]@{ orgId = "$($o.id)"; users = @($u) } }
            catch { $script:Warnings.Add("org $($o.name): $(Get-ErrMsg $_)") }
        }
    }
    Collect 'orgSettings'           '/cloudapi/v1/orgSettings (per organization)' -Script {
        foreach ($o in $orgList) {
            try { $s = Invoke-AsTenant $o.id { Get-CloudOne 'orgSettings' }; [pscustomobject]@{ orgId = "$($o.id)"; settings = $s } }
            catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("org $($o.name): $(Get-ErrMsg $_)") } }
        }
    }
    Collect 'orgMetadata'           '/cloudapi/1.0.0/orgs/{id}/metadata' -Script {
        foreach ($o in $orgList) {
            try { $m = Get-FirstPath @("/cloudapi/v1/orgs/$($o.id)/metadata", "/cloudapi/1.0.0/orgs/$($o.id)/metadata"); [pscustomobject]@{ orgId = "$($o.id)"; metadata = @($m) } }
            catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("org $($o.name): $(Get-ErrMsg $_)") } }
        }
    }
    Collect 'regionQuotas'          '/cloudapi/v1/regionQuotas | virtualDatacenters' -Script { Get-CloudFirst @('regionQuotas', 'virtualDatacenters') }
    Collect 'regionQuotaDetails'    '/cloudapi/v1/<regionQuotas>/{id}[/vmClasses|storagePolicies|zoneResourceAllocations]' -Script {
        $base = if ($script:Resolved['regionQuotas']) { 'regionQuotas' } else { 'virtualDatacenters' }
        foreach ($q in (Items $script:Raw.regionQuotas)) {
            $det = $null
            try { $det = Get-CloudOne "$base/$($q.id)" } catch { $script:Warnings.Add("quota $($q.name): $(Get-ErrMsg $_)") }
            $subs = [ordered]@{}
            foreach ($sub in 'virtualMachineClasses', 'vmClasses', 'storagePolicies', 'zoneResourceAllocations') {
                try { $subs[$sub] = @(Get-CloudList "$base/$($q.id)/$sub") }
                catch { if ((Get-StatusCode $_) -notin 404, 400) { $script:Warnings.Add("quota $($q.name) ${sub}: $(Get-ErrMsg $_)") } }
            }
            [pscustomobject]@{ id = "$($q.id)"; detail = $det; sub = $subs }
        }
    }
    Collect 'vdcStoragePolicies'    '/cloudapi/v1/virtualDatacenterStoragePolicies' -Script { Get-CloudFirst @('virtualDatacenterStoragePolicies', 'regionQuotaStoragePolicies') }
    Collect 'orgNetworking'         '/cloudapi/(v1|1.0.0)/orgs/{id}/networkingSettings' -Script {
        foreach ($o in $orgList) {
            try { $n = Get-FirstPath @("/cloudapi/v1/orgs/$($o.id)/networkingSettings", "/cloudapi/vcf/orgs/$($o.id)/networkingSettings", "/cloudapi/1.0.0/orgs/$($o.id)/networkingSettings") -Single; [pscustomobject]@{ orgId = "$($o.id)"; settings = $n } }
            catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("org $($o.name): $(Get-ErrMsg $_)") } }
        }
    }
    Collect 'regionalNetworking'    '/cloudapi/v1/regionalNetworkingSettings' -Script { Get-CloudList 'regionalNetworkingSettings' }

    # --- Networking
    Collect 'providerGateways'      '/cloudapi/v1/providerGateways' -Script { Get-CloudList 'providerGateways' }
    Collect 'ipSpaces'              '/cloudapi/v1/ipSpaces' -Script { Get-CloudList 'ipSpaces' }
    Collect 'ipSpaceUtilization'    '/cloudapi/v1/ipSpaces/{id}/utilizationSummary' -Script {
        foreach ($s in (Items $script:Raw.ipSpaces)) {
            try { [pscustomobject]@{ ipSpaceId = "$($s.id)"; summary = (Get-CloudOne "ipSpaces/$($s.id)/utilizationSummary") } }
            catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("ip space $($s.name): $(Get-ErrMsg $_)") } }
        }
    }
    Collect 'ipSpaceAssociations'   '/cloudapi/v1/ipSpaceAssociations' -Script { Get-CloudList 'ipSpaceAssociations' }
    Collect 'ipSpaceOrgQuotas'      '/cloudapi/v1/ipSpaces/orgQuotas' -Script { Get-CloudList 'ipSpaces/orgQuotas' }
    Collect 'ipOrgUtilizations'     '/cloudapi/v1/ipSpaces/orgQuotas/ipUtilizations' -Script { Get-CloudList 'ipSpaces/orgQuotas/ipUtilizations' }
    Collect 'edgeClusters'          '/cloudapi/v1/edgeClusters' -Script { Get-CloudList 'edgeClusters' }
    Collect 'transitGateways'       '/cloudapi/v1/transitGateways' -Script { Get-CloudList 'transitGateways' }

    # --- Regions, supervisors, classes
    Collect 'regions'               '/cloudapi/v1/regions' -Script { Get-CloudList 'regions' }
    Collect 'zones'                 '/cloudapi/v1/zones' -Script { Get-CloudList 'zones' }
    Collect 'supervisors'           '/cloudapi/v1/supervisors' -Script { Get-CloudFirst @('supervisors', 'supervisorClusters') }
    Collect 'vmClasses'             '/cloudapi/v1/virtualMachineClasses' -Script { Get-CloudFirst @('virtualMachineClasses', 'regionVmClasses') }
    Collect 'reservedVmClasses'     '/cloudapi/v1/reservedVMClasses' -Script { Get-CloudList 'reservedVMClasses' }
    Collect 'regionStoragePolicies' '/cloudapi/v1/regionStoragePolicies' -Script { Get-CloudList 'regionStoragePolicies' }
    Collect 'storageClasses'        '/cloudapi/v1/storageClasses' -Script { Get-CloudList 'storageClasses' }
    Collect 'contentLibraries'      '/cloudapi/v1/contentLibraries' -Script { Get-CloudList 'contentLibraries' }
    Collect 'contentLibraryItems'   '/cloudapi/v1/contentLibraryItems' -Script { Get-CloudList 'contentLibraryItems' }
    Collect 'dataServices'          '/cloudapi/v1/vcfServiceCapabilities' -Script { Get-CloudList 'vcfServiceCapabilities' } -EmptyOn404

    # --- Connections
    Collect 'vCenters'              '/cloudapi/1.0.0/virtualCenters'
    Collect 'nsxManagers'           '/cloudapi/v1/nsxManagers | /api/admin/extension/nsxtManagers' -Script {
        try { Get-CloudFirst @('nsxManagers', 'nsxTManagers') }
        catch { if ((Get-StatusCode $_) -ne 404) { throw }; @(Get-P (Invoke-Vcfa '/api/admin/extension/nsxtManagers') 'nsxTManager', 'nsxtManager') | Where-Object { $_ } }
    }
    Collect 'vcfInstances'          '/cloudapi/v1/vcfInfraEndpoints' -Script { Get-CloudFirst @('vcfInfraEndpoints', 'vcfInstances', 'sddcManagers') }

    # --- Settings
    Collect 'generalSettings'       '/api/admin/extension/settings/general' -Single
    Collect 'featureFlags'          '/cloudapi/1.0.0/featureFlags'
    Collect 'branding'              '/cloudapi/branding' -Single -EmptyOn404
    Collect 'brandingThemes'        '/cloudapi/branding/themes' -Script { @(Get-Items (Invoke-Vcfa '/cloudapi/branding/themes')) }

    # --- Access control
    Collect 'users'                 '/cloudapi/1.0.0/users'
    Collect 'groups'                '/cloudapi/1.0.0/groups'
    Collect 'roles'                 '/cloudapi/1.0.0/roles'
    Collect 'globalRoles'           '/cloudapi/1.0.0/globalRoles'
    Collect 'rightsBundles'         '/cloudapi/1.0.0/rightsBundles'
    Collect 'publishedTo'           '/cloudapi/1.0.0/(globalRoles|rightsBundles)/{id}/tenants' -Script {
        foreach ($kind in 'globalRoles', 'rightsBundles') {
            foreach ($x in (Items $script:Raw[$kind])) {
                if (Get-P $x 'publishAll') { continue }
                try { [pscustomobject]@{ kind = $kind; id = "$($x.id)"; tenants = @(Get-All "/cloudapi/1.0.0/$kind/$($x.id)/tenants") } }
                catch { if ((Get-StatusCode $_) -ne 404) { $script:Warnings.Add("$kind $($x.name): $(Get-ErrMsg $_)") } }
            }
        }
    }
    Collect 'serviceAccounts'       '/cloudapi/1.0.0/serviceAccounts'
    Collect 'tokens'                '/cloudapi/1.0.0/tokens'

    # --- Identity
    Collect 'identityProviders'     '/api/admin/org/{System}/settings/(oauth|federation) + /api/admin/extension/settings/ldapSettings' -Script {
        $sys = (Items $script:Raw.orgs) | Where-Object { "$($_.name)" -eq 'System' } | Select-Object -First 1
        $uuid = if ($sys) { ("$($sys.id)" -split ':')[-1] } else { $null }
        foreach ($kind in 'oauth', 'federation') {
            $res = $null; $err = $null
            if ($uuid) {
                try { $res = Invoke-Vcfa "/api/admin/org/$uuid/settings/$kind" }
                catch { if ((Get-StatusCode $_) -ne 404) { $err = Get-ErrMsg $_ } }
            }
            else { $err = 'System organization not visible through /cloudapi/1.0.0/orgs' }
            [pscustomobject]@{ kind = $kind; settings = $res; error = $err }
        }
        $res = $null; $err = $null
        try { $res = Invoke-Vcfa '/api/admin/extension/settings/ldapSettings' }
        catch { if ((Get-StatusCode $_) -ne 404) { $err = Get-ErrMsg $_ } }
        [pscustomobject]@{ kind = 'ldap'; settings = $res; error = $err }
    }
    Collect 'oidcRelyingParties'    '/cloudapi/1.0.0/openIdProvider/relyingParties'
    Collect 'oidcKeys'              '/cloudapi/1.0.0/openIdProvider/keys'
    Collect 'oidcConfig'            '/cloudapi/1.0.0/openIdProvider' -Single -EmptyOn404

    # --- Certificates
    Collect 'trustedCerts'          '/cloudapi/1.0.0/ssl/trustedCertificates'
    Collect 'certLibrary'           '/cloudapi/1.0.0/ssl/certificateLibrary'
}

# =====================================================================================
#  Workbook rendering (works only from the saved JSON, so -FromJson gives identical output)
# =====================================================================================

$NotFromApi = 'Not available from the VCF Automation API - fill in manually'

function Build-Workbook {
    param($Data, [string]$Path)
    $Src = $Data.raw
    $script:FailedKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($k in (Items $Data.failed)) { $script:FailedKeys.Add("$k") }
    $srv = $Data.meta.server

    $Bk = New-Book
    $shIdx = Add-Sheet $Bk 'Index'
    $shVer = Add-Sheet $Bk 'Version History'
    $shOrg = Add-Sheet $Bk 'Organizations'
    $shNet = Add-Sheet $Bk 'Networking'
    $shReg = Add-Sheet $Bk 'Regions'
    $shSup = Add-Sheet $Bk 'Supervisors'
    $shCl  = Add-Sheet $Bk 'Content Libraries'
    $shRc  = Add-Sheet $Bk 'Roles and Credentials'
    $shDs  = Add-Sheet $Bk 'VCF Data Services'
    $shGen = Add-Sheet $Bk 'General Settings'
    $shCon = Add-Sheet $Bk 'Connections'
    $shVcf = Add-Sheet $Bk 'VCF Instances'
    $shAcc = Add-Sheet $Bk 'Access Control'
    $shIdp = Add-Sheet $Bk 'Identity Providers'
    $shOid = Add-Sheet $Bk 'OIDC Services'
    $shCrt = Add-Sheet $Bk 'Certificate Management'
    $shBr  = Add-Sheet $Bk 'Branding'
    $shFf  = Add-Sheet $Bk 'Feature Flags'
    $shLog = Add-Sheet $Bk 'Collection Log'

    # ---------------- lookups ----------------
    $orgs = @((Items $Src.orgs) | Where-Object { "$($_.name)" -ne 'System' })
    $orgById = @{}; foreach ($o in (Items $Src.orgs)) { $orgById["$($o.id)"] = $o }
    function OrgName($v) { $id = if ($v -is [string]) { $v } else { Get-P $v 'id' }; if ($id -and $orgById.ContainsKey("$id")) { return $orgById["$id"].name }; return (RefName $v) }
    function ByOrg($list, [string]$Field) { $h = @{}; foreach ($x in (Items $list)) { $h["$(Get-P $x 'orgId')"] = Get-P $x $Field }; return $h }
    $orgUsers = ByOrg $Src.orgUsers 'users'
    $orgSet   = ByOrg $Src.orgSettings 'settings'
    $orgMeta  = ByOrg $Src.orgMetadata 'metadata'
    $orgNet   = ByOrg $Src.orgNetworking 'settings'
    function OfOrg($list, $org) { return , @((Items $list) | Where-Object { "$(Get-P $_ 'org.id', 'orgRef.id', 'orgId')" -eq "$($org.id)" -or "$(Get-P $_ 'org.name', 'orgRef.name')" -eq "$($org.name)" }) }

    $regions = Items $Src.regions
    $regionById = @{}; foreach ($r in $regions) { $regionById["$($r.id)"] = $r }
    function RegOf($o) {
        $v = Get-P $o 'region', 'regionRef', 'regionId'
        if ($v -is [string] -and $regionById.ContainsKey($v)) { return $regionById[$v].name }
        return (RefName $v)
    }
    $zones = Items $Src.zones
    $zoneById = @{}; foreach ($z in $zones) { $zoneById["$($z.id)"] = $z }
    $vmcById = @{}; foreach ($v in (Items $Src.vmClasses)) { $vmcById["$($v.id)"] = $v }
    $rspById = @{}; foreach ($s in (Items $Src.regionStoragePolicies)) { $rspById["$($s.id)"] = $s }
    $supById = @{}; foreach ($s in (Items $Src.supervisors)) { $supById["$($s.id)"] = $s }
    $vcById = @{}; foreach ($v in (Items $Src.vCenters)) { $vcById["$(Get-P $v 'vcId', 'id')"] = $v }
    function NameOf($h, $v) { $id = if ($v -is [string]) { $v } else { Get-P $v 'id' }; if ($id -and $h.ContainsKey("$id")) { return (Get-P $h["$id"] 'name') }; return (RefName $v) }
    $qDet = @{}; foreach ($d in (Items $Src.regionQuotaDetails)) { $qDet["$(Get-P $d 'id')"] = $d }

    # =========================================================================
    #  Organizations
    # =========================================================================
    Add-Title $shOrg 'Organizations' "VCF Automation: $srv"
    $rows = foreach ($o in $orgs) {
        $set = $orgSet["$($o.id)"]; $flat = Get-Flat ([ordered]@{ org = $o; settings = $set })
        $first = (Items $orgUsers["$($o.id)"]) | Select-Object -First 1
        [ordered]@{
            'Name' = $o.name; 'Display Name' = Get-P $o 'displayName'
            'Organization URL' = "https://$srv/tenant/$($o.name)"
            'Enabled' = Get-P $o 'isEnabled', 'enabled'
            'Description' = Get-P $o 'description'
            'Metadata' = Join-Lines @((Items $orgMeta["$($o.id)"]) | ForEach-Object { '{0} = {1}' -f (Get-P $_ 'keyValue.key', 'key'), (Format-Value (Get-P $_ 'keyValue.value.value', 'keyValue.value', 'value')) })
            'First User - User name' = Get-P $first 'username', 'name'
            'First User - Role' = Join-Lines (@(Get-P $first 'roleEntityRefs') | ForEach-Object { Get-P $_ 'name' })
            'Allow this org to create subscribed content libraries' = Join-Lines @($flat.Keys | Where-Object { $_ -match '(?i)subscri' } | ForEach-Object { '{0} = {1}' -f $_, (Format-Value $flat[$_]) })
            'Enable blocking of content library items upload' = Join-Lines @($flat.Keys | Where-Object { $_ -match '(?i)upload|block' } | ForEach-Object { '{0} = {1}' -f $_, (Format-Value $flat[$_]) })
        }
    }
    Add-Vertical $Bk $shOrg 'Initial Configuration' $rows 'orgs'

    $rows = [System.Collections.Generic.List[object]]::new()
    $zrows = [System.Collections.Generic.List[object]]::new()
    $vrows = [System.Collections.Generic.List[object]]::new()
    $srows = [System.Collections.Generic.List[object]]::new()
    foreach ($q in (Items $Src.regionQuotas)) {
        $d = $qDet["$($q.id)"]; $det = Get-P $d 'detail'; $both = @($det, $q); $sub = Get-P $d 'sub'
        $orgN = OrgName (Get-P $both 'org', 'orgRef', 'orgId')
        $rows.Add([ordered]@{
                'Organization' = $orgN; 'Region Quota' = $q.name
                'Region' = RegOf $q
                'Supervisor' = Join-Lines @(@(Get-P $both 'supervisors', 'supervisorIds', 'supervisorRefs') | ForEach-Object { NameOf $supById $_ })
                'Status' = Get-P $both 'status'
            })
        $za = @(Get-P $both 'zoneResourceAllocation', 'zoneResourceAllocations') + @(Get-P $sub 'zoneResourceAllocations')
        foreach ($z in ($za | Where-Object { $_ })) {
            $zr = Get-P $z 'zone', 'regionZone', 'zoneRef', 'regionZoneId', 'zoneId'
            $zrows.Add([ordered]@{
                    'Organization' = $orgN; 'Region Quota' = $q.name; 'Zone Name' = NameOf $zoneById $zr
                    'CPU Limit (MHz)' = Get-P $z 'resourceAllocation.cpuLimitMHz', 'cpuLimitMHz', 'cpuLimitMhz'
                    'CPU Reservation (MHz)' = Get-P $z 'resourceAllocation.cpuReservationMHz', 'cpuReservationMHz', 'cpuReservationMhz'
                    'Memory Limit (MiB)' = Get-P $z 'resourceAllocation.memoryLimitMiB', 'memoryLimitMiB'
                    'Memory Reservation (MiB)' = Get-P $z 'resourceAllocation.memoryReservationMiB', 'memoryReservationMiB'
                })
        }
        $vmc = @(Get-P $sub 'virtualMachineClasses', 'vmClasses') + @(Get-P $both 'vmClasses', 'vmClassIds', 'regionVmClassIds', 'virtualMachineClasses')
        $added = @($vmc | Where-Object { $_ } | ForEach-Object { NameOf $vmcById $_ } | Select-Object -Unique)
        $regName = RegOf $q
        $allInRegion = @((Items $Src.vmClasses) | Where-Object { -not $regName -or (RegOf $_) -eq $regName } | ForEach-Object { $_.name } | Select-Object -Unique)
        foreach ($n in @(@($allInRegion) + @($added) | Select-Object -Unique)) {
            $vrows.Add([ordered]@{ 'Organization' = $orgN; 'Region Quota' = $q.name; 'VM Class' = $n; 'Status' = if ($added -contains $n) { 'added' } else { 'not added' } })
        }
        $sps = @(Get-P $sub 'storagePolicies') + @(Get-P $both 'storagePolicies', 'regionStoragePolicies', 'regionStoragePolicy') + @((Items $Src.vdcStoragePolicies) | Where-Object { "$(Get-P $_ 'virtualDatacenter.id', 'vdc.id', 'regionQuota.id')" -eq "$($q.id)" })
        foreach ($s in ($sps | Where-Object { $_ })) {
            $srows.Add([ordered]@{
                    'Organization' = $orgN; 'Region Quota' = $q.name
                    'Name' = NameOf $rspById (Get-P $s 'regionStoragePolicy', 'regionStoragePolicyId', 'name')
                    'Limit (MiB)' = Get-P $s 'storageLimitMiB', 'storageLimitMB', 'limit'
                })
        }
    }
    Add-Table $Bk $shOrg 'Region Quota' $rows 'regionQuotas'
    Add-Table $Bk $shOrg 'Region Quota - Zones' $zrows 'regionQuotas', 'regionQuotaDetails'
    Add-Table $Bk $shOrg 'Region Quota - VM Classes' $vrows 'regionQuotas', 'regionQuotaDetails'
    Add-Table $Bk $shOrg 'Region Quota - Storage Classes' $srows 'regionQuotas', 'regionQuotaDetails', 'vdcStoragePolicies'
    $rows = foreach ($o in $orgs) {
        foreach ($r in (OfOrg $Src.reservedVmClasses $o)) { [ordered]@{ 'Organization' = $o.name; 'Name' = Get-P $r 'name', 'vmClass.name'; 'Zone' = RefName (Get-P $r 'zone', 'zoneRef'); 'Used\Total' = Get-FlatMatch $r '(?i)used|total|count' } }
    }
    Add-Table $Bk $shOrg 'Reserved VM Classes' $rows 'reservedVmClasses'

    # Networking per org
    $rows = foreach ($o in $orgs) {
        $ns = $orgNet["$($o.id)"]
        $rn = OfOrg $Src.regionalNetworking $o
        $iq = OfOrg $Src.ipSpaceOrgQuotas $o
        $iu = OfOrg $Src.ipOrgUtilizations $o
        [ordered]@{
            'Organization' = $o.name
            'Log Name' = Get-P $ns 'orgNameForLogs', 'logName', 'networkingTenancyName'
            'Regional Networking - Name' = Join-Lines @($rn | ForEach-Object { $_.name })
            'Regional Networking - Region' = Join-Lines @($rn | ForEach-Object { RegOf $_ })
            'Regional Networking - Provider Gateway' = Join-Lines @($rn | ForEach-Object { RefName (Get-P $_ 'providerGatewayRef', 'providerGateway') })
            'Regional Networking - Edge Cluster' = Join-Lines @($rn | ForEach-Object { RefName (Get-P $_ 'edgeClusterRef', 'edgeCluster', 'serviceEdgeClusterRef') })
            'Regional Networking - Status' = Join-Lines @($rn | ForEach-Object { Get-P $_ 'status', 'realizedState' })
            'VPC Services - Edge Cluster' = Join-Lines @($rn | ForEach-Object { RefName (Find-Deep $_ '(?i)vpc.*edgeCluster|serviceEdgeCluster') })
            'Ingress Traffic' = Join-Lines @($rn | ForEach-Object { Get-FlatMatch $_ '(?i)ingress' })
            'Egress Traffic' = Join-Lines @($rn | ForEach-Object { Get-FlatMatch $_ '(?i)egress' })
            'IP Usage - IP Space' = Join-Lines @($iq | ForEach-Object { RefName (Get-P $_ 'ipSpaceRef', 'ipSpace') })
            'IP Usage - Quotas' = Join-Lines @($iq | ForEach-Object { Get-FlatMatch $_ '(?i)quota|cidr|subnet|ip' })
            'IP Usage - Utilization' = Join-Lines @($iu | ForEach-Object { Get-FlatMatch $_ '(?i)used|count|cidr|subnet' })
        }
    }
    Add-Vertical $Bk $shOrg 'Networking' $rows @('regionalNetworking', 'orgNetworking', 'ipSpaceOrgQuotas') 'Organization'

    # =========================================================================
    #  Networking (provider)
    # =========================================================================
    Add-Title $shNet 'Networking' "VCF Automation: $srv"
    $assoc = Items $Src.ipSpaceAssociations
    $rows = foreach ($g in (Items $Src.providerGateways)) {
        $a = @($assoc | Where-Object { "$(Get-P $_ 'providerGatewayRef.id', 'providerGateway.id')" -eq "$($g.id)" })
        [ordered]@{
            'Name' = $g.name; 'Description' = Get-P $g 'description'; 'Region' = RegOf $g; 'Status' = Get-P $g 'status', 'realizedState'
            'T0 Gateways' = Join-Lines @(@(Get-P $g 'backingRef', 'tier0GatewayRefs', 'backingRefs', 'tier0Gateways') | ForEach-Object { RefName $_ })
            'Associated IP Spaces' = Join-Lines @($a | ForEach-Object { '{0} ({1})' -f (RefName (Get-P $_ 'ipSpaceRef', 'ipSpace')), (Get-P $_ 'status', 'realizedState') })
        }
    }
    Add-Vertical $Bk $shNet 'Provider Gateways' $rows 'providerGateways', 'ipSpaceAssociations'
    $util = @{}; foreach ($u in (Items $Src.ipSpaceUtilization)) { $util["$(Get-P $u 'ipSpaceId')"] = Get-P $u 'summary' }
    $rows = foreach ($s in (Items $Src.ipSpaces)) {
        $pg = @($assoc | Where-Object { "$(Get-P $_ 'ipSpaceRef.id', 'ipSpace.id')" -eq "$($s.id)" } | ForEach-Object { RefName (Get-P $_ 'providerGatewayRef', 'providerGateway') })
        [ordered]@{
            'Name' = $s.name; 'Status' = Get-P $s 'status', 'realizedState'
            'IP Blocks' = Join-Lines @(@(Get-P $s 'internalScopeCidrBlocks', 'ipBlocks', 'cidrBlocks', 'externalScopeCidr') | ForEach-Object { if ($_ -is [string]) { $_ } else { Get-P $_ 'cidr', 'name' } })
            'Region' = RegOf $s
            'Number of Single IPs' = Get-P $s 'defaultQuota.maxIPCount', 'defaultGatewayServiceConfig.maxIpCount', 'quota.maxIpCount', 'maxIpCount'
            'Number of CIDRs' = Get-P $s 'defaultQuota.maxCidrCount', 'quota.maxCidrCount', 'maxCidrCount'
            'Maximum Subnet Size' = Get-P $s 'defaultQuota.maxSubnetSize', 'quota.maxSubnetSize', 'maxSubnetSize'
            'Provider Gateways' = Join-Lines $pg
            'Utilization' = Get-FlatMatch $util["$($s.id)"] '.'
        }
    }
    Add-Vertical $Bk $shNet 'IP Spaces' $rows 'ipSpaces', 'ipSpaceUtilization'
    $rows = foreach ($e in (Items $Src.edgeClusters)) {
        [ordered]@{
            'Name' = $e.name; 'Health' = Get-P $e 'healthStatus', 'health', 'status'
            'Nodes' = Get-P $e 'nodeCount', 'nodes'; 'Organizations' = Get-P $e 'orgCount', 'organizationCount'
            'VPCs' = Get-P $e 'vpcCount'; 'Type' = Get-P $e 'deploymentType', 'type'; 'Region' = RegOf $e
        }
    }
    Add-Vertical $Bk $shNet 'Edge Clusters' $rows 'edgeClusters'
    $rows = foreach ($t in (Items $Src.transitGateways)) { [ordered]@{ 'Name' = $t.name; 'Region' = RegOf $t; 'Organization' = OrgName (Get-P $t 'orgRef', 'org'); 'Status' = Get-P $t 'status', 'realizedState'; 'Description' = Get-P $t 'description' } }
    Add-Table $Bk $shNet 'Transit Gateways' $rows 'transitGateways'

    # =========================================================================
    #  Regions
    # =========================================================================
    Add-Title $shReg 'Regions' "VCF Automation: $srv"
    $rows = foreach ($r in $regions) {
        $sups = @((Items $Src.supervisors) | Where-Object { (RegOf $_) -eq $r.name })
        [ordered]@{
            'Name' = $r.name; 'Status' = Get-P $r 'status'; 'Description' = Get-P $r 'description'
            'NSX Local Manager' = RefName (Get-P $r 'nsxManager', 'nsxManagerRef')
            'Supervisors' = Join-Lines @($sups | ForEach-Object { '{0} (vCenter: {1})' -f $_.name, (NameOf $vcById (Get-P $_ 'virtualCenter', 'vcenter', 'vCenterRef', 'vcId')) })
            'Zones' = Join-Lines @($zones | Where-Object { (RegOf $_) -eq $r.name } | ForEach-Object { $_.name })
            'CPU / Memory' = Get-FlatMatch $r '(?i)cpu|mem'
            'Storage' = Get-FlatMatch $r '(?i)storage'
        }
    }
    Add-Vertical $Bk $shReg 'Regions' $rows 'regions'
    $rows = foreach ($z in $zones) {
        [ordered]@{
            'Region' = RegOf $z; 'Name' = $z.name
            'CPU Allocation' = Get-FlatMatch $z '(?i)cpu.*(alloc|used|limit|capacity)'
            'Memory Allocation' = Get-FlatMatch $z '(?i)mem.*(alloc|used|limit|capacity)'
            'CPU Reservation' = Get-FlatMatch $z '(?i)cpu.*reserv'
            'Memory Reservation' = Get-FlatMatch $z '(?i)mem.*reserv'
        }
    }
    Add-Table $Bk $shReg 'Zones' $rows 'zones'
    $rows = foreach ($v in (Items $Src.vmClasses)) {
        $mem = Get-P $v 'memoryMiB', 'memoryMB', 'memory'
        [ordered]@{
            'Region' = RegOf $v; 'Name' = $v.name
            'CPU Reservation' = Get-P $v 'cpuReservationRequestedMHz', 'cpuReservation', 'cpuReservationMHz'
            'CPUs' = Get-P $v 'cpuCount', 'cpus', 'numCpus'
            'Memory Reservation' = Get-P $v 'memoryReservationRequestedMiB', 'memoryReservation', 'memoryReservationMiB'
            'Memory (GB)' = if ($mem -is [ValueType]) { [math]::Round([double]$mem / 1024, 2) } else { $mem }
            'Reserved VM Class' = Get-P $v 'reserved', 'isReserved'
        }
    }
    Add-Table $Bk $shReg 'VM Classes' $rows 'vmClasses'
    $rows = foreach ($s in (Items $Src.regionStoragePolicies)) {
        [ordered]@{ 'Name' = $s.name; 'Region' = RegOf $s; 'Storage Consumption' = Get-FlatMatch $s '(?i)storage|used|capacity|limit' }
    }
    Add-Table $Bk $shReg 'Storage Classes (region storage policies)' $rows 'regionStoragePolicies'
    $rows = foreach ($s in (Items $Src.storageClasses)) { [ordered]@{ 'Name' = $s.name; 'Region' = RegOf $s; 'Details' = Get-FlatMatch $s '(?i)limit|capacity|used|MiB|GiB' } }
    Add-Table $Bk $shReg 'Storage Classes' $rows 'storageClasses'
    $rows = foreach ($r in (Items $Src.reservedVmClasses)) { [ordered]@{ 'Name' = Get-P $r 'name', 'vmClass.name'; 'Zone' = RefName (Get-P $r 'zone', 'zoneRef'); 'Organization' = OrgName (Get-P $r 'org', 'orgRef'); 'Used\Total' = Get-FlatMatch $r '(?i)used|total|count' } }
    Add-Table $Bk $shReg 'Reserved VM Classes' $rows 'reservedVmClasses'

    # =========================================================================
    #  Supervisors
    # =========================================================================
    Add-Title $shSup 'Supervisors' "VCF Automation: $srv"
    $rows = foreach ($s in (Items $Src.supervisors)) {
        $sz = @($zones | Where-Object { "$(Get-P $_ 'supervisor.id', 'supervisorRef.id')" -eq "$($s.id)" })
        if ($sz.Count -eq 0) { $sz = @(@(Get-P $s 'zones', 'zoneRefs') | Where-Object { $_ }) }
        [ordered]@{
            'Name' = $s.name; 'Region' = RegOf $s
            'vCenter' = NameOf $vcById (Get-P $s 'virtualCenter', 'vcenter', 'vCenterRef', 'vcId')
            'Zones' = Join-Lines @($sz | ForEach-Object { RefName $_ })
            'Zone capacity' = Join-Lines @($sz | ForEach-Object { if ($_ -is [string]) { '' } else { '{0}: {1}' -f (RefName $_), ((Get-FlatMatch $_ '(?i)cpu|mem') -replace "`n", '; ') } })
            'Status' = Get-P $s 'status'
        }
    }
    Add-Vertical $Bk $shSup 'Supervisors' $rows 'supervisors'

    # =========================================================================
    #  Content Libraries
    # =========================================================================
    Add-Title $shCl 'Content Libraries' "VCF Automation: $srv"
    $rows = foreach ($l in (Items $Src.contentLibraries)) {
        [ordered]@{
            'Name' = $l.name; 'Description' = Get-P $l 'description'; 'Status' = Get-P $l 'status'; 'Type' = Get-P $l 'libraryType'
            'Organization' = OrgName (Get-P $l 'org'); 'Subscribed' = Get-P $l 'isSubscribed'; 'Version' = Get-P $l 'versionNumber'
            'Creation Date' = Get-P $l 'creationDate'
            'Storage Classes' = Join-Lines @((Items (Get-P $l 'storageClasses')) | ForEach-Object { '{0} ({1})' -f (Get-P $_ 'name'), (Get-P $_ 'region.name', 'regionName') })
        }
    }
    Add-Vertical $Bk $shCl 'Content Libraries' $rows 'contentLibraries'
    $rows = foreach ($i in (Items $Src.contentLibraryItems)) {
        [ordered]@{
            'Name' = $i.name; 'Version' = Get-P $i 'versionNumber', 'version'; 'Status' = Get-P $i 'status'
            'Image Identifier' = Get-P $i 'imageIdentifier', 'imageId'; 'Content Library' = Get-P $i 'contentLibrary.name', 'library.name'
            'Type' = Get-P $i 'itemType', 'type'
        }
    }
    Add-Table $Bk $shCl 'Content Library Items' $rows 'contentLibraryItems'

    # =========================================================================
    #  Roles and Credentials (design information - never read from the system)
    # =========================================================================
    Add-Title $shRc 'Roles and Credentials' 'Built-in and design accounts. Passwords are never collected: fill in from your password vault reference.'
    $acct = [ordered]@{
        'Provider Portal'          = @('Enterprise Admin / Provider Admin', 'Custom provider roles')
        'LOB/Organization Portal'  = @('Organization Administrator', 'Organization Auditor', 'Organization User', 'Custom tenant roles')
        'Projects Roles'           = @('Project Admin', 'Project Auditor', 'Project User', 'Project AdvanceUser')
        'Integrations'             = @('Active Directory User')
    }
    $rows = foreach ($g in $acct.Keys) { foreach ($a in $acct[$g]) { [ordered]@{ 'Portal' = $g; 'Account' = $a; 'Username' = 'fill in manually'; 'Password' = 'not stored in the sheet'; 'Description' = "Username and password for the $a" } } }
    Add-Table $Bk $shRc 'Appliance Usernames and Passwords (Built-In Accounts)' $rows ''
    $rows = foreach ($r in (Items $Src.roles)) { [ordered]@{ 'Provider Role' = $r.name; 'Description' = Get-P $r 'description'; 'Read Only' = Get-P $r 'readOnly' } }
    Add-Table $Bk $shRc 'Provider Roles (for reference)' $rows 'roles'

    # =========================================================================
    #  VCF Data Services
    # =========================================================================
    Add-Title $shDs 'VCF Data Services' "VCF Automation: $srv"
    $ds = Items $Src.dataServices
    $rows = foreach ($x in $ds) {
        [ordered]@{
            'Name' = Get-P $x 'name', 'displayName', 'id'
            'Region' = RegOf $x
            'Data Services FQDN' = Find-Deep $x '(?i)fqdn|hostname|url|endpoint'
            'Data Services Certificate Status' = Find-Deep $x '(?i)cert.*status|trust'
            'Data Service Admin Username' = Find-Deep $x '(?i)admin.*user|username'
            'Data Service Admin Password' = 'not stored in the sheet'
            'All fields' = Get-FlatMatch $x '.'
        }
    }
    Add-Vertical $Bk $shDs 'Data Services' $rows 'dataServices'

    # =========================================================================
    #  General Settings
    # =========================================================================
    Add-Title $shGen 'General Settings' "VCF Automation: $srv"
    $gs = Get-Flat $Src.generalSettings
    function GsFind([string]$Regex) { return (Join-Lines @($gs.Keys | Where-Object { $_ -match $Regex } | ForEach-Object { '{0} = {1}' -f $_, (Format-Value $gs[$_]) })) }
    Add-Section $Bk $shGen 'Settings' 2 11
    Add-KeyValue $shGen ([ordered]@{
            'Activity Log - Log history to keep'                       = GsFind '(?i)activityLogKeep'
            'Activity Log - Log history shown'                         = GsFind '(?i)activityLogDisplay'
            'Activity Log - Display debug information'                 = GsFind '(?i)stackTrace|debug'
            'Timeouts - API token lifetime'                            = GsFind '(?i)token.*(life|expir|timeout)'
            'Timeouts - Host refresh frequency'                        = GsFind '(?i)hostCheckDelay|hostRefresh'
            'Timeouts - Host hung timeout'                             = GsFind '(?i)hostCheckTimeout|hostHung'
            'Timeouts - Transfer session timeout'                      = GsFind '(?i)transferSession'
            'Organization Limits - Maximum number of Region Quotas'    = GsFind '(?i)(regionQuota|vdc).*(max|limit)|max.*(regionQuota|vdc)'
            'Operation Limits'                                         = GsFind '(?i)operation|resourceIntensive|queued'
        })
    $rows = foreach ($k in $gs.Keys) { [ordered]@{ 'Name' = $k; 'Value' = $gs[$k] } }
    Add-Table $Bk $shGen 'All General Settings' $rows 'generalSettings'

    # =========================================================================
    #  Connections / VCF Instances
    # =========================================================================
    Add-Title $shCon 'Connections' "VCF Automation: $srv"
    $rows = foreach ($v in (Items $Src.vCenters)) {
        [ordered]@{
            'Name' = Get-P $v 'name'; 'URL' = Get-P $v 'url'
            'State' = if ($null -ne (Get-P $v 'isEnabled')) { if (Get-P $v 'isEnabled') { 'Enabled' } else { 'Disabled' } } else { '' }
            'Connected' = Get-P $v 'isConnected', 'listenerState', 'connectionState'
            'Username' = Get-P $v 'username', 'userName'
            'License Status' = Get-P $v 'licenseStatus', 'licensed'
            'VCF Instance' = RefName (Get-P $v 'vcfInstance', 'vcfInfraEndpoint', 'sddcManager')
            'Status' = Get-P $v 'vcStatus', 'status', 'clusterHealthStatus'
            'Version' = Get-P $v 'vcVersion', 'version'
        }
    }
    Add-Vertical $Bk $shCon 'vCenter' $rows 'vCenters'
    $rows = foreach ($n in (Items $Src.nsxManagers)) {
        [ordered]@{ 'Name' = Get-P $n 'name'; 'URL' = Get-P $n 'url'; 'Description' = Get-P $n 'description'; 'Status' = Get-P $n 'status', 'connectionState'; 'Username' = Get-P $n 'username', 'userName'; 'Version' = Get-P $n 'version', 'productVersion' }
    }
    Add-Vertical $Bk $shCon 'NSX Manager' $rows 'nsxManagers'

    Add-Title $shVcf 'VCF Instances' "VCF Automation: $srv"
    $rows = foreach ($i in (Items $Src.vcfInstances)) {
        [ordered]@{ 'Name' = Get-P $i 'name'; 'Host Name' = Get-P $i 'hostName', 'hostname', 'fqdn', 'url'; 'Username' = Get-P $i 'username', 'userName'; 'Version' = Get-P $i 'version', 'productVersion'; 'Status' = Get-P $i 'status' }
    }
    Add-Vertical $Bk $shVcf 'VCF Instances' $rows 'vcfInstances'

    # =========================================================================
    #  Access Control
    # =========================================================================
    Add-Title $shAcc 'Access Control' "VCF Automation: $srv  -  provider (System) organization"
    $rows = foreach ($u in (Items $Src.users)) {
        $en = Get-P $u 'enabled'; $lk = Get-P $u 'isLocked', 'locked'
        [ordered]@{
            'User Name' = Get-P $u 'username', 'name'; 'Full Name' = Get-P $u 'fullName'; 'Email' = Get-P $u 'email'
            'State' = if ($null -eq $en) { '' } elseif ($en) { 'Enabled' } else { 'Disabled' }
            'Locked' = if ($null -eq $lk) { '' } elseif ($lk) { 'Locked' } else { 'Unlocked' }
            'Effective Roles' = Join-Lines (@(Get-P $u 'roleEntityRefs') | ForEach-Object { Get-P $_ 'name' })
            'Provider Type' = Get-P $u 'providerType'; 'Stranded' = Get-P $u 'stranded'
        }
    }
    Add-Vertical $Bk $shAcc 'Users' $rows 'users' 'User Name'
    $rows = foreach ($g in (Items $Src.groups)) {
        [ordered]@{ 'Name' = Get-P $g 'name'; 'Description' = Get-P $g 'description'; 'Role' = Join-Lines (@(Get-P $g 'roleEntityRefs', 'roleEntityRef') | ForEach-Object { Get-P $_ 'name' }); 'Type' = Get-P $g 'providerType' }
    }
    Add-Vertical $Bk $shAcc 'Groups' $rows 'groups'
    $rows = foreach ($r in (Items $Src.roles)) { [ordered]@{ 'Name' = $r.name; 'Description' = Get-P $r 'description'; 'Read Only' = Get-P $r 'readOnly' } }
    Add-Vertical $Bk $shAcc 'Roles' $rows 'roles'
    $rows = foreach ($t in (Items $Src.tokens)) { [ordered]@{ 'Name' = Get-P $t 'name'; 'Type' = Get-P $t 'type'; 'Owner' = RefName (Get-P $t 'owner'); 'Expires On' = Get-P $t 'expirationTimeUtc', 'expiresOn' } }
    Add-Table $Bk $shAcc 'System Tokens' $rows 'tokens'
    $rows = foreach ($s in (Items $Src.serviceAccounts)) {
        [ordered]@{ 'Name' = Get-P $s 'name'; 'Role' = Get-P $s 'role.name', 'role'; 'Status' = Get-P $s 'status'; 'Software ID' = Get-P $s 'softwareId'; 'Client URI' = Get-P $s 'uri', 'clientUri'; 'Software Version' = Get-P $s 'softwareVersion' }
    }
    Add-Table $Bk $shAcc 'Service Accounts' $rows 'serviceAccounts'
    $pub = @{}; foreach ($p in (Items $Src.publishedTo)) { $pub["$(Get-P $p 'kind')|$(Get-P $p 'id')"] = Items (Get-P $p 'tenants') }
    function PublishedTo($kind, $x) { if (Get-P $x 'publishAll') { return 'All' }; $t = $pub["$kind|$($x.id)"]; if ($t) { return ('Specific Tenants: ' + ((@($t) | ForEach-Object { RefName $_ }) -join ', ')) }; return 'Specific Tenants' }
    $rows = foreach ($r in (Items $Src.globalRoles)) { [ordered]@{ 'Name' = $r.name; 'Description' = Get-P $r 'description'; 'Published To' = PublishedTo 'globalRoles' $r } }
    Add-Vertical $Bk $shAcc 'Global Roles' $rows 'globalRoles'
    $rows = foreach ($r in (Items $Src.rightsBundles)) { [ordered]@{ 'Name' = $r.name; 'Description' = Get-P $r 'description'; 'Published To' = PublishedTo 'rightsBundles' $r } }
    Add-Vertical $Bk $shAcc 'Right Bundles' $rows 'rightsBundles'

    # =========================================================================
    #  Identity Providers
    # =========================================================================
    Add-Title $shIdp 'Identity Providers' "VCF Automation: $srv  -  provider (System) organization"
    $idp = @{}; foreach ($x in (Items $Src.identityProviders)) { $idp["$(Get-P $x 'kind')"] = $x }
    function IdpNote($k) { $x = $idp[$k]; if ($x -and (Get-P $x 'error')) { "<not collected: $(Get-P $x 'error')>" } else { $null } }
    $o = Get-P $idp['oauth'] 'settings'
    $rows = if ($o -and ((Get-P $o 'enabled') -or (Get-P $o 'issuerId', 'clientId'))) {
        Add-RestOf ([ordered]@{
                'Name' = 'VCF SSO / OIDC'; 'Identity Provider Name' = Get-P $o 'providerName', 'name'; 'Identity Broker Issuer' = Get-P $o 'issuerId'
                'Enabled' = Get-P $o 'enabled'; 'Client ID' = Get-P $o 'clientId'
                'Authorization endpoint' = Get-P $o 'userAuthorizationEndpoint'; 'Token endpoint' = Get-P $o 'accessTokenEndpoint'
                'User info endpoint' = Get-P $o 'userInfoEndpoint'; 'Scope' = Get-P $o 'scope'; 'Claim mapping' = Get-P $o 'oAuthClaimMapping', 'claimMapping'
            }) $o @('providerName', 'name', 'issuerId', 'enabled', 'clientId', 'userAuthorizationEndpoint', 'accessTokenEndpoint', 'userInfoEndpoint', 'scope', 'oAuthClaimMapping', 'claimMapping')
    }
    Add-Vertical $Bk $shIdp 'VCF SSO' $rows 'identityProviders'
    if (IdpNote 'oauth') { [void](Add-Row $shIdp @(C (IdpNote 'oauth') 6)); Add-Blank $shIdp }

    $l = Get-P $idp['ldap'] 'settings'
    $rows = if ($l -and (Get-P $l 'hostName')) {
        $ua = Get-P $l 'userAttributes'; $ga = Get-P $l 'groupAttributes'
        $d = [ordered]@{
            'Name' = 'LDAP'; 'Server' = Get-P $l 'hostName'; 'Port' = Get-P $l 'port'; 'Base Distinguished Name' = Get-P $l 'searchBase'
            'Use SSL' = Get-P $l 'isSsl'; 'Connector Type' = Get-P $l 'connectorType'; 'Authentication Method' = Get-P $l 'authenticationMechanism'
            'User Name' = Get-P $l 'userName'; 'Password' = '(not exported)'
            'User - Object class' = Get-P $ua 'objectClass'; 'User - Unique identifier' = Get-P $ua 'objectIdentifier'; 'User - User name' = Get-P $ua 'userName'
            'User - Display name' = Get-P $ua 'fullName'; 'User - Given name' = Get-P $ua 'givenName'; 'User - Surname' = Get-P $ua 'surname'
            'User - Email' = Get-P $ua 'email'; 'User - Telephone' = Get-P $ua 'telephone'
            'User - Group membership identifier' = Get-P $ua 'groupMembershipIdentifier'; 'User - Group back link' = Get-P $ua 'groupBackLinkIdentifier'
            'Group - Object class' = Get-P $ga 'objectClass'; 'Group - Unique identifier' = Get-P $ga 'objectIdentifier'; 'Group - Name' = Get-P $ga 'groupName'
            'Group - Membership' = Get-P $ga 'membership'; 'Group - Group membership identifier' = Get-P $ga 'membershipIdentifier'; 'Group - Group back link identifier' = Get-P $ga 'backLinkIdentifier'
        }
        $d = Add-RestOf $d $ua @('objectClass', 'objectIdentifier', 'userName', 'fullName', 'givenName', 'surname', 'email', 'telephone', 'groupMembershipIdentifier', 'groupBackLinkIdentifier')
        $d = Add-RestOf $d $ga @('objectClass', 'objectIdentifier', 'groupName', 'membership', 'membershipIdentifier', 'backLinkIdentifier')
        Add-RestOf $d $l @('hostName', 'port', 'searchBase', 'isSsl', 'connectorType', 'authenticationMechanism', 'userName', 'userAttributes', 'groupAttributes')
    }
    Add-Vertical $Bk $shIdp 'LDAP' $rows 'identityProviders'
    if (IdpNote 'ldap') { [void](Add-Row $shIdp @(C (IdpNote 'ldap') 6)); Add-Blank $shIdp }

    $f = Get-P $idp['federation'] 'settings'
    $rows = if ($f -and ((Get-P $f 'enabled') -or (Get-P $f 'samlMetadata'))) {
        Add-RestOf ([ordered]@{
                'Name' = 'SAML'; 'Use SAML Identity Provider' = Get-P $f 'enabled'; 'Entity ID' = Get-P $f 'samlSPEntityId'
                'Certificate expiration' = Get-P $f 'certificateExpiration'; 'Attribute Mapping' = Get-P $f 'samlAttributeMapping'
            }) $f @('enabled', 'samlSPEntityId', 'certificateExpiration', 'samlAttributeMapping')
    }
    Add-Vertical $Bk $shIdp 'SAML' $rows 'identityProviders'
    if (IdpNote 'federation') { [void](Add-Row $shIdp @(C (IdpNote 'federation') 6)); Add-Blank $shIdp }

    # =========================================================================
    #  OIDC Services (VCF Automation as OpenID provider)
    # =========================================================================
    Add-Title $shOid 'OIDC Services' "VCF Automation: $srv"
    $rows = foreach ($r in (Items $Src.oidcRelyingParties)) {
        [ordered]@{
            'Relying Party Name' = Get-P $r 'clientName', 'name'; 'Relying Party ID' = Get-P $r 'clientId', 'id'
            'Is Public' = Get-P $r 'isPublic', 'public'; 'Requires PKCE' = Get-P $r 'requirePkce', 'requiresPkce'
            'Redirect URIs' = Get-P $r 'redirectUris'
        }
    }
    Add-Vertical $Bk $shOid 'Relying Parties' $rows 'oidcRelyingParties' 'Relying Party Name'
    $rows = foreach ($k in (Items $Src.oidcKeys)) {
        [ordered]@{ 'Description' = Get-P $k 'description', 'name'; 'Creation Time' = Get-P $k 'creationDate', 'createdAt'; 'Expiration' = Get-P $k 'expirationDate'; 'Algorithm' = Get-P $k 'algorithm'; 'Public Key' = Get-P $k 'publicKey' }
    }
    Add-Vertical $Bk $shOid 'Keys' $rows 'oidcKeys' 'Description'
    $c = $Src.oidcConfig
    $rows = if ($c) {
        Add-RestOf ([ordered]@{
                'Name' = 'OpenID provider'
                'Well-known Configuration URL' = Get-P $c 'wellKnownEndpoint', 'wellKnownConfigurationUrl'
                'Access Token Lifetime' = Get-P $c 'accessTokenLifespan', 'accessTokenLifetime'
                'Refresh Token Lifetime' = Get-P $c 'refreshTokenLifespan', 'refreshTokenLifetime'
                'ID Token Lifetime' = Get-P $c 'idTokenLifespan', 'idTokenLifetime'
                'Authorization Code Lifetime' = Get-P $c 'authorizationCodeLifespan', 'authorizationCodeLifetime'
                'Redirect URL Policy' = Get-P $c 'redirectUriPolicy', 'redirectUrlPolicy'
            }) $c @('wellKnownEndpoint', 'wellKnownConfigurationUrl', 'accessTokenLifespan', 'accessTokenLifetime', 'refreshTokenLifespan', 'refreshTokenLifetime', 'idTokenLifespan', 'idTokenLifetime', 'authorizationCodeLifespan', 'authorizationCodeLifetime', 'redirectUriPolicy', 'redirectUrlPolicy')
    }
    Add-Vertical $Bk $shOid 'General Settings' $rows 'oidcConfig'

    # =========================================================================
    #  Certificate Management
    # =========================================================================
    Add-Title $shCrt 'Certificate Management' "VCF Automation: $srv"
    $rows = foreach ($t in (Items $Src.trustedCerts)) {
        $ci = Get-CertInfo (Get-P $t 'certificate')
        [ordered]@{ 'Friendly Name' = Get-P $t 'alias', 'name'; 'Common Name' = $ci.Cn; 'Issuer' = $ci.Issuer; 'Expires On' = $ci.Expires }
    }
    Add-Table $Bk $shCrt 'Trusted Certificates' $rows 'trustedCerts'
    $rows = foreach ($t in (Items $Src.certLibrary)) {
        $ci = Get-CertInfo (Get-P $t 'certificate')
        [ordered]@{
            'Friendly Name' = Get-P $t 'alias', 'name'; 'Description' = Get-P $t 'description'; 'Common Name' = $ci.Cn; 'Issuer' = $ci.Issuer
            'Expires On' = $ci.Expires; 'Certificates in Chain' = $ci.Chain; 'Consumers' = Get-P $t 'consumerCount', 'consumers'
        }
    }
    Add-Table $Bk $shCrt 'Certificates Library' $rows 'certLibrary'

    # =========================================================================
    #  Branding
    # =========================================================================
    Add-Title $shBr 'Branding' "VCF Automation: $srv"
    $b = $Src.branding
    $rows = if ($b) {
        Add-RestOf ([ordered]@{
                'Name' = 'Branding'
                'Enable Login & Logout Page Branding' = Get-P $b 'loginPageBranding', 'brandLoginPage', 'customizeLoginPage'
                'Theme Name' = Get-P $b 'selectedTheme.name', 'selectedTheme.themeType'
                'Platform Name' = Get-P $b 'portalName'
                'Portal Color' = Get-P $b 'portalColor'
                'Custom Links' = Get-P $b 'customLinks'
            }) $b @('loginPageBranding', 'brandLoginPage', 'customizeLoginPage', 'selectedTheme', 'portalName', 'portalColor', 'customLinks')
    }
    Add-Vertical $Bk $shBr 'Branding' $rows 'branding'
    [void](Add-Row $shBr @(C 'Brand logo, browser icon and login background images are not exported; check them in the UI.' 6)); Add-Blank $shBr
    $rows = foreach ($t in (Items $Src.brandingThemes)) { [ordered]@{ 'Name' = Get-P $t 'name'; 'Type' = Get-P $t 'themeType'; 'Colors' = Get-FlatMatch $t '(?i)color' } }
    Add-Table $Bk $shBr 'Themes' $rows 'brandingThemes'

    # =========================================================================
    #  Feature Flags
    # =========================================================================
    Add-Title $shFf 'Feature Flags' "VCF Automation: $srv"
    $rows = foreach ($x in (Items $Src.featureFlags)) {
        [ordered]@{
            'Name' = Get-P $x 'displayName', 'name'; 'Usage' = Get-P $x 'usage'; 'Description' = Get-P $x 'description'
            'State (Disabled\Enabled)' = if (Get-P $x 'enabled') { 'Enabled' } else { 'Disabled' }
        }
    }
    Add-Table $Bk $shFf 'Feature Flags' $rows 'featureFlags'

    # =========================================================================
    #  Version History, Collection Log, Index
    # =========================================================================
    Add-Title $shVer 'Version History'
    [void](Add-Row $shVer @((C 'Date' 3), (C 'Version' 3), (C 'Author' 3), (C 'Description' 3), (C 'Reviewers' 3)))
    [void](Add-Row $shVer @((C (Format-Value $Data.meta.collectedAt)), (C '0.1'), (C $Data.meta.collectedBy), (C "Auto-collected from $srv (provider) by script v$($Data.meta.scriptVersion)"), (C '')))

    Add-Title $shLog 'Collection Log' 'One line per API call group. FAILED usually means the endpoint does not exist in this version or the account lacks rights.'
    Add-Table $Bk $shLog 'API calls' (Items $Data.log)

    Add-Title $shIdx 'VCF Automation - Provider Configuration'
    Add-KeyValue $shIdx ([ordered]@{
            'VCF Automation'  = $srv
            'Collected at'    = $Data.meta.collectedAt
            'Collected by'    = $Data.meta.collectedBy
            'Login method'    = $Data.meta.auth
            'Script version'  = $Data.meta.scriptVersion
            'Failed sections' = if ($script:FailedKeys.Count) { $script:FailedKeys -join ', ' } else { 'none' }
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
if (-not $PSBoundParameters.ContainsKey('SkipCertificateCheck')) {
    $ans = Read-Host 'Skip TLS certificate validation (self-signed / untrusted CA)? [y/N]'
    $SkipCertificateCheck = [bool]($ans -match '^(y|yes)$')
}
if (-not $UseApiToken -and -not $Credential) {
    $Credential = Get-Credential -Message 'VCF Automation provider user (@System is added automatically)'
    if (-not $Credential) { throw 'No credentials entered.' }
}

$script:Conn.Base = $BaseUri
$script:Conn.Skip = [bool]$SkipCertificateCheck
Set-TlsOptions ([bool]$SkipCertificateCheck)

Write-Host "Logging in to $BaseUri as provider ..." -ForegroundColor Cyan
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
        scriptVersion  = $ScriptVersion
    }
    failed = @($script:FailedKeys)
    log    = @($script:Log)
    raw    = $script:Raw
}

# One folder per VCF Automation instance, named after the first label of its FQDN, e.g. "Provider - vcfa"
$short = (($Server -split ':')[0] -split '\.')[0]
$safeOrg = $short -replace '[^\w\-]', '_'
$orgFolder = Join-Path $OutputFolder ('Provider - ' + $safeOrg)
if (-not (Test-Path -LiteralPath $orgFolder)) { New-Item -ItemType Directory -Path $orgFolder | Out-Null }
$base = Get-FreePath $orgFolder ("VCFA_{0}_Provider_Config_{1}" -f $safeOrg, (Get-Date -Format 'yyyyMMdd')) @('.xlsx', '.json')

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

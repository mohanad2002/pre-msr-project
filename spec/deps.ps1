param($d)
Add-Type -AssemblyName System.Web
[xml]$wb = Get-Content "$d\xl\workbook.xml" -Raw -Encoding UTF8
[xml]$rel = Get-Content "$d\xl\_rels\workbook.xml.rels" -Raw
$map = @{}; $rel.Relationships.Relationship | % { $map[$_.Id] = $_.Target }
$sheets = @()
foreach ($s in $wb.workbook.sheets.sheet) {
  $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $sheets += [pscustomobject]@{ Name = $s.name; File = $map[$rid] }
}
$names = $sheets.Name
$out = @()
$fnStats = @{}
foreach ($sh in $sheets) {
  $raw = Get-Content "$d\xl\$($sh.File)" -Raw -Encoding UTF8
  $fs = [regex]::Matches($raw, '<f[^>]*>(.*?)</f>', 'Singleline') | % { [System.Web.HttpUtility]::HtmlDecode($_.Groups[1].Value) }
  $refs = @{}
  $ext = 0
  $fn = @{}
  foreach ($f in $fs) {
    foreach ($m in [regex]::Matches($f, "'((?:[^']|'')+)'!|([A-Za-z_][\w\.]*)!")) {
      $n = if ($m.Groups[1].Success) { $m.Groups[1].Value -replace "''", "'" } else { $m.Groups[2].Value }
      if ($n -match '^\[\d+\]') { $ext++; continue }
      if ($names -contains $n -and $n -ne $sh.Name) { $refs[$n] = 1 + [int]$refs[$n] }
    }
    if ($f -match '\[\d+\]') { $ext++ }
    foreach ($m in [regex]::Matches($f, '([A-Z][A-Z0-9\.]+)\(')) { $fn[$m.Groups[1].Value] = 1 + [int]$fn[$m.Groups[1].Value] }
  }
  $top = ($fn.GetEnumerator() | sort Value -desc | select -first 5 | % { "$($_.Key)($($_.Value))" }) -join ' '
  "### $($sh.Name) | ext=$ext | fns: $top"
  $refs.GetEnumerator() | sort Value -desc | % { "   <- $($_.Key) ($($_.Value))" }
}

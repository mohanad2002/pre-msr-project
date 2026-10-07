param($d, [int[]]$ids, [int]$maxRow = 14, [int]$maxLen = 1800)
Add-Type -AssemblyName System.Web
$ssRaw = Get-Content "$d\xl\sharedStrings.xml" -Raw -Encoding UTF8
$ss = @()
foreach ($si in [regex]::Matches($ssRaw, '<si>(.*?)</si>', 'Singleline')) {
  $t = ([regex]::Matches($si.Groups[1].Value, '<t[^>]*>(.*?)</t>', 'Singleline') | % { $_.Groups[1].Value }) -join ''
  $ss += [System.Web.HttpUtility]::HtmlDecode($t)
}
[xml]$wb = Get-Content "$d\xl\workbook.xml" -Raw -Encoding UTF8
$names = @($wb.workbook.sheets.sheet | % { $_.name })
foreach ($i in $ids) {
  $raw = Get-Content "$d\xl\worksheets\sheet$i.xml" -Raw -Encoding UTF8
  $sb = "=== [$i] $($names[$i-1])`n"
  foreach ($row in [regex]::Matches($raw, '<row r="(\d+)"[^>]*>(.*?)</row>', 'Singleline')) {
    if ([int]$row.Groups[1].Value -gt $maxRow) { break }
    $cells = @()
    foreach ($c in [regex]::Matches($row.Groups[2].Value, '<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', 'Singleline')) {
      $v = [regex]::Match($c.Groups[3].Value, '<v>(.*?)</v>').Groups[1].Value
      if ($v -eq '') { continue }
      if ($c.Groups[2].Value -match 't="s"') { $v = $ss[[int]$v] }
      elseif ($c.Groups[2].Value -match 't="str"') { $v = [System.Web.HttpUtility]::HtmlDecode($v) }
      if ($c.Groups[3].Value -match '<f') { $v = "~" + $v }
      if ($v.Length -gt 40) { $v = $v.Substring(0, 40) }
      $cells += "$($c.Groups[1].Value)=$v"
    }
    if ($cells) { $sb += "r$($row.Groups[1].Value): " + ($cells -join ' | ') + "`n" }
  }
  if ($sb.Length -gt $maxLen) { $sb = $sb.Substring(0, $maxLen) + "...`n" }
  $sb
}

param($src, $d)
New-Item -ItemType Directory -Force $d | Out-Null
$zip = "$d.zip"
Copy-Item $src $zip -Force
Expand-Archive $zip $d -Force
[xml]$wb = Get-Content "$d\xl\workbook.xml" -Raw -Encoding UTF8
[xml]$rel = Get-Content "$d\xl\_rels\workbook.xml.rels" -Raw
$map = @{}; $rel.Relationships.Relationship | % { $map[$_.Id] = $_.Target }
$i = 0
foreach ($sh in $wb.workbook.sheets.sheet) {
  $i++
  $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $f = $map[$rid]
  $raw = Get-Content "$d\xl\$f" -Raw -Encoding UTF8
  $tab = [regex]::Match($raw, '<tabColor[^>]*/>').Value
  $nf = ([regex]::Matches($raw, '<f[ >]')).Count
  $ext = ([regex]::Matches($raw, "<f[^>]*>[^<]*'?\[\d+\][^<]*!")).Count
  $err = ([regex]::Matches($raw, 't="e"')).Count
  $af = [regex]::Match($raw, '<autoFilter ref="([^"]+)"').Groups[1].Value
  "{0}|{1}|{2}|{3}|{4}|f={5}|ext={6}|errCells={7}|af={8}" -f $i, $sh.name, $sh.state, $f, $tab, $nf, $ext, $err, $af
}
"--- names"
$wb.workbook.definedNames.definedName | % { "$($_.name) [$($_.localSheetId)] = $($_.'#text')" }
"--- ext links"
Get-ChildItem "$d\xl\externalLinks\_rels" -ErrorAction SilentlyContinue | % { $_.Name; [regex]::Matches((Get-Content $_.FullName -Raw), 'Target="([^"]+)"') | % { "  " + $_.Groups[1].Value } }
"--- caches"
Get-ChildItem "$d\xl\pivotCache\pivotCacheDefinition*.xml" | % { $r = Get-Content $_.FullName -Raw; "$($_.Name): " + [regex]::Match($r, '<worksheetSource[^>]*/>').Value }
Get-ChildItem "$d\xl\pivotCache\_rels" | % { $t = [regex]::Matches((Get-Content $_.FullName -Raw), 'externalLinkPath" Target="([^"]+)"'); if ($t.Count) { "$($_.Name) EXT: $($t[0].Groups[1].Value)" } }
"--- pivot->cache"
Get-ChildItem "$d\xl\pivotTables\_rels" | % { "$($_.Name): " + [regex]::Match((Get-Content $_.FullName -Raw), 'Target="([^"]+)"').Groups[1].Value }
"--- sheet rels"
Get-ChildItem "$d\xl\worksheets\_rels" | % { $p = [regex]::Matches((Get-Content $_.FullName -Raw), 'Target="([^"]*pivotTable[^"]*)"') | % { $_.Groups[1].Value -replace '.*/', '' }; if ($p) { "$($_.Name): $($p -join ', ')" } }
"--- slicers"
Get-ChildItem "$d\xl\slicerCaches\*.xml" -ErrorAction SilentlyContinue | % { $r = Get-Content $_.FullName -Raw; "$($_.Name): " + [regex]::Match($r, 'sourceName="([^"]+)"').Groups[1].Value + " pivots=" + (([regex]::Matches($r, 'tabId="(\d+)" name="([^"]+)"') | % { "$($_.Groups[1].Value):$($_.Groups[2].Value)" }) -join ',') }

param(
  [Parameter(Mandatory=$true)][ValidateSet('inspect','apply')][string]$Mode,
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$ReportPath,
  [Parameter(Mandatory=$true)][string]$IdentityPath,
  [string]$PlanPath='',
  [string]$OutputDocx=''
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word_pid_contract.ps1')

function Get-AxMathShapes($d){
  $shapes=@()
  for($i=1;$i -le [int]$d.InlineShapes.Count;$i++){
    $shape=$d.InlineShapes.Item($i)
    try{if([string]$shape.OLEFormat.ProgID -eq 'Equation.AxMath'){$shapes += $shape}}catch{}
  }
  return $shapes
}
function Get-SharedSha256([string]$Path){
  $stream=$null
  $sha=$null
  try{
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    $sha=[Security.Cryptography.SHA256]::Create()
    return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','').ToLowerInvariant()
  }finally{
    if($sha){$sha.Dispose()}
    if($stream){$stream.Dispose()}
  }
}

$inputFull=[IO.Path]::GetFullPath($InputDocx)
$reportFull=[IO.Path]::GetFullPath($ReportPath)
$identityFull=[IO.Path]::GetFullPath($IdentityPath)
if([string]::Equals($reportFull,$identityFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Report and identity paths must differ.'}
if(-not (Test-Path -LiteralPath $inputFull -PathType Leaf)){throw 'Input DOCX does not exist.'}
if(Test-Path -LiteralPath $reportFull){throw 'Report path already exists.'}
if(Test-Path -LiteralPath $identityFull){throw 'Identity path already exists.'}
$inputHash=Get-SharedSha256 $inputFull
$plan=$null
$partial=''
$outputFull=''
if($Mode -eq 'apply'){
  if(-not $PlanPath -or -not $OutputDocx){throw 'Apply requires PlanPath and OutputDocx.'}
  $plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
  if($plan.schema -ne 'axmath-body-math-plan/v1'){throw 'Unsupported body-math plan schema.'}
  if(([string]$plan.input_sha256).ToLowerInvariant() -ne $inputHash){throw 'Plan input hash is stale.'}
  if(@($plan.targets).Count -eq 0){throw 'Plan has no targets.'}
  $outputFull=[IO.Path]::GetFullPath($OutputDocx)
  if([string]::Equals($outputFull,$inputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Output must differ from input.'}
  if([string]::Equals($outputFull,$reportFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($outputFull,$identityFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Output, report, and identity paths must differ.'}
  if(Test-Path -LiteralPath $outputFull){throw 'Output path already exists.'}
  $partial=$outputFull+'.'+[guid]::NewGuid().ToString('N')+'.partial.docx'
  [IO.File]::Copy($inputFull,$partial,$false)
  $item=Get-Item -LiteralPath $partial
  if($item.IsReadOnly){$item.IsReadOnly=$false}
}
New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($reportFull)) | Out-Null
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$word=$null
$identity=$null
$doc=$null
$tmp=$null
$donorDocs=New-Object 'System.Collections.Generic.List[object]'
$success=$false
$failure=''
$report=[ordered]@{
  schema='axmath-body-math-operation/v1'; mode=$Mode; input=$inputFull;
  input_sha256=$inputHash; output=$outputFull; partial=$partial; success=$false
}
try{
  $started=Get-Date
  $word=New-Object -ComObject Word.Application
  Start-Sleep -Milliseconds 300
  $identity=Resolve-OwnedWordIdentity -Word $word -BeforePids $beforeWord -LaunchStartedAt $started
  $identity | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $identityFull -Encoding UTF8
  $report.word_identity=$identity
  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  if($Mode -eq 'inspect'){
    $doc=$word.Documents.OpenNoRepairDialog($inputFull,$false,$true,$false)
    $paragraphs=New-Object 'System.Collections.Generic.List[object]'
    $totalCandidates=0
    for($i=1;$i -le [int]$doc.Paragraphs.Count;$i++){
      $p=$doc.Paragraphs.Item($i).Range
      $pText=[string]$p.Text
      $candidates=New-Object 'System.Collections.Generic.List[object]'
      $offset=0
      while($offset -lt $pText.Length){
        $size=1
        if([char]::IsHighSurrogate($pText[$offset]) -and $offset+1 -lt $pText.Length -and [char]::IsLowSurrogate($pText[$offset+1])){$size=2}
        $category=[Globalization.CharUnicodeInfo]::GetUnicodeCategory($pText,$offset).ToString()
        if(@('UppercaseLetter','LowercaseLetter','TitlecaseLetter','ModifierLetter','DecimalDigitNumber','LetterNumber','OtherNumber') -contains $category){
          $leftStart=[Math]::Max(0,$offset-8)
          $rightStart=$offset+$size
          $rightLength=[Math]::Min(8,$pText.Length-$rightStart)
          $candidates.Add([pscustomobject]@{
            candidate_id=('P'+$i+'-T'+$offset); text_offset_utf16=$offset;
            text=$pText.Substring($offset,$size);
            before=$pText.Substring($leftStart,$offset-$leftStart);
            after=$pText.Substring($rightStart,$rightLength)
          })
          $totalCandidates++
        }
        $offset+=$size
      }
      $paragraphs.Add([pscustomobject]@{paragraph_index=$i; text=$pText; fields=[int]$p.Fields.Count; candidates=$candidates.ToArray()})
    }
    $report.paragraphs=$paragraphs.ToArray()
    $report.candidate_count=$totalCandidates
    $report.omath_count=[int]$doc.OMaths.Count
    $report.axmath_count=@(Get-AxMathShapes $doc).Count
    if((Get-SharedSha256 $inputFull) -ne $inputHash){throw 'Input changed during inspection.'}
  }else{
    $template=''
    foreach($base in @(${env:ProgramFiles(x86)},$env:ProgramFiles)){
      if($base){$candidate=Join-Path $base 'AxMath\MSOffice\AxMath.dotm';if(Test-Path -LiteralPath $candidate){$template=$candidate;break}}
    }
    if(-not $template){throw 'AxMath.dotm not found.'}
    $found=$false
    foreach($addin in $word.AddIns){
      if($addin.Name -eq 'AxMath.dotm'){$found=$true;if(-not $addin.Installed){$addin.Installed=$true}}
    }
    if(-not $found){$word.AddIns.Add($template,$true)|Out-Null}
    $doc=$word.Documents.OpenNoRepairDialog($partial,$false,$false,$false)
    if([int]$doc.OMaths.Count -ne 0){throw 'Complete the OfficeMath conversion before inserting body targets.'}
    if($doc.TrackRevisions -or [int]$doc.Revisions.Count -gt 0){throw 'Tracked changes require a separate reviewed workflow.'}
    $paragraphCount=[int]$doc.Paragraphs.Count
    $expected=@{}
    for($i=1;$i -le $paragraphCount;$i++){$expected[$i]=([string]$doc.Paragraphs.Item($i).Range.Text).ToCharArray()}
    $shapeCount=[int]$doc.InlineShapes.Count
    $axCount=@(Get-AxMathShapes $doc).Count
    $ids=@{}
    $locations=@{}
    $resolved=@()
    foreach($row in @($plan.targets)){
      $id=[string]$row.target_id
      $sourceText=[string]$row.source_text
      $tex=[string]$row.expected_tex
      $pi=[int]$row.paragraph_index
      if(-not $id -or $ids.ContainsKey($id)){throw 'Target ID is missing or duplicated.'}
      $ids[$id]=$true
      if($sourceText.Length -ne 1 -or -not [char]::IsLetterOrDigit($sourceText[0])){throw "$id is not a single BMP letter or digit."}
      if(-not $tex -or $tex.Contains('$') -or $tex.Contains("`r") -or $tex.Contains("`n")){throw "$id has invalid TeX."}
      if($sourceText -cmatch '^[A-Za-z0-9]$' -and $tex -cne $sourceText){throw "$id single ASCII character differs from its TeX."}
      if($pi -lt 1 -or $pi -gt $paragraphCount){throw "$id paragraph index is out of range."}
      $p=$doc.Paragraphs.Item($pi).Range
      $pText=[string]$p.Text
      $before=[string]$row.before
      $after=[string]$row.after
      $anchor=$before+$sourceText+$after
      foreach($ch in $anchor.ToCharArray()){if([char]::IsControl($ch)){throw "$id anchor cannot cross a field, object, or paragraph boundary."}}
      $idx=$pText.IndexOf($anchor,[StringComparison]::Ordinal)
      if($idx -lt 0 -or $pText.LastIndexOf($anchor,[StringComparison]::Ordinal) -ne $idx){throw "$id anchor is absent or ambiguous."}
      $findRange=$p.Duplicate
      $find=$findRange.Find
      $find.ClearFormatting()
      $found=$find.Execute($anchor,$true,$false,$false,$false,$false,$true,0,$false)
      if(-not $found -or [string]$findRange.Text -cne $anchor){throw "$id cannot be located by exact Word Find."}
      $start=[int]$findRange.Start+$before.Length
      if($locations.ContainsKey($start)){throw "$id overlaps another target."}
      $locations[$start]=$true
      $target=$doc.Range($start,$start+1)
      if([string]$target.Text -cne $sourceText){throw "$id Word range does not match source text."}
      for($fieldIndex=1;$fieldIndex -le [int]$p.Fields.Count;$fieldIndex++){
        $field=$p.Fields.Item($fieldIndex)
        $fieldStart=[Math]::Min([int]$field.Code.Start,[int]$field.Result.Start)-1
        $fieldEnd=[Math]::Max([int]$field.Code.End,[int]$field.Result.End)+1
        if($start -ge $fieldStart -and $start -lt $fieldEnd){throw "$id overlaps a Word field."}
      }
      for($linkIndex=1;$linkIndex -le [int]$p.Hyperlinks.Count;$linkIndex++){
        $linkRange=$p.Hyperlinks.Item($linkIndex).Range
        if($start -ge [int]$linkRange.Start -and $start -lt [int]$linkRange.End){throw "$id overlaps a hyperlink."}
      }
      for($controlIndex=1;$controlIndex -le [int]$p.ContentControls.Count;$controlIndex++){
        $controlRange=$p.ContentControls.Item($controlIndex).Range
        if($start -ge [int]$controlRange.Start -and $start -lt [int]$controlRange.End){throw "$id overlaps a content control."}
      }
      $textOffset=$idx+$before.Length
      if($textOffset -lt 0 -or $textOffset -ge $expected[$pi].Length -or [string]$expected[$pi][$textOffset] -cne $sourceText){throw "$id text position is invalid or overlaps another target."}
      $expected[$pi][$textOffset]=[char]1
      $resolved += [pscustomobject]@{target_id=$id; paragraph_index=$pi; source_text=$sourceText; expected_tex=$tex; start=$start}
    }
    $donorByTex=[System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach($row in @($resolved|Sort-Object start -Descending)){
      $start=[int]$row.start
      $target=$doc.Range($start,$start+1)
      if([string]$target.Text -cne [string]$row.source_text){throw "Target $($row.target_id) shifted before mutation."}
      $tex=[string]$row.expected_tex
      if(-not $donorByTex.ContainsKey($tex)){
        $tmp=$word.Documents.Add()
        $tmp.Content.Text='$'+$tex+'$'
        $tmp.Content.Select()
        $callback=$null
        $word.Run('AMSTeX2AM',([ref]$callback))
        $donors=@(Get-AxMathShapes $tmp)
        if($donors.Count -ne 1){throw "Target $($row.target_id) did not produce one AxMath donor."}
        $donorByTex.Add($tex,$donors[0])
        $donorDocs.Add($tmp)
        $tmp=$null
      }
      $target.Delete()|Out-Null
      $doc.Range($start,$start).FormattedText=$donorByTex[$tex].Range.FormattedText
      $currentParagraph=$doc.Paragraphs.Item([int]$row.paragraph_index).Range
      $inserted=@(Get-AxMathShapes $currentParagraph | Where-Object{[int]$_.Range.Start -eq $start})
      if($inserted.Count -ne 1){throw "Target $($row.target_id) did not insert exactly one AxMath at the approved position."}
      $row | Add-Member -NotePropertyName shape -NotePropertyValue $inserted[0]
    }
    for($i=1;$i -le $paragraphCount;$i++){
      if([string]$doc.Paragraphs.Item($i).Range.Text -cne (-join $expected[$i])){throw "Paragraph $i changed outside the approved targets."}
    }
    if([int]$doc.Paragraphs.Count -ne $paragraphCount -or [int]$doc.InlineShapes.Count -ne $shapeCount+$resolved.Count){throw 'Paragraph or inline-shape count drifted.'}
    $allAx=@(Get-AxMathShapes $doc)
    if($allAx.Count -ne $axCount+$resolved.Count){throw 'AxMath count differs from the approved target count.'}
    $ordinalByStart=@{}
    for($i=0;$i -lt $allAx.Count;$i++){
      $shapeStart=[int]$allAx[$i].Range.Start
      if($ordinalByStart.ContainsKey($shapeStart)){throw 'Multiple AxMath objects share a Word position.'}
      $ordinalByStart[$shapeStart]=$i+1
    }
    $rows=@()
    foreach($row in $resolved){
      $shapeStart=[int]$row.shape.Range.Start
      if(-not $ordinalByStart.ContainsKey($shapeStart)){throw "Cannot locate final object for $($row.target_id)."}
      $rows += [pscustomobject]@{target_id=$row.target_id; paragraph_index=$row.paragraph_index; source_text=$row.source_text; expected_tex=$row.expected_tex; final_ordinal=$ordinalByStart[$shapeStart]}
    }
    $doc.Save()
    $doc.Close($false)
    $doc=$null
    $doc=$word.Documents.OpenNoRepairDialog($partial,$false,$true,$false)
    if([int]$doc.Paragraphs.Count -ne $paragraphCount -or @(Get-AxMathShapes $doc).Count -ne $axCount+$resolved.Count){throw 'Saved output did not preserve the approved paragraphs and objects.'}
    for($i=1;$i -le $paragraphCount;$i++){
      if([string]$doc.Paragraphs.Item($i).Range.Text -cne (-join $expected[$i])){throw "Saved paragraph $i changed outside the approved targets."}
    }
    $doc.Close($false)
    $doc=$null
    if((Get-SharedSha256 $inputFull) -ne $inputHash){throw 'Input changed during the operation.'}
    [IO.File]::Move($partial,$outputFull)
    $report.output_sha256=Get-SharedSha256 $outputFull
    $report.axmath_before=$axCount
    $report.axmath_after=$allAx.Count
    $report.distinct_donor_count=$donorByTex.Count
    $report.targets=$rows
    $report.target_count=$rows.Count
  }
  $success=$true
}catch{
  $failure=$_.Exception.Message
}finally{
  if($tmp){try{$tmp.Close($false)}catch{}}
  foreach($donorDoc in $donorDocs){try{$donorDoc.Close($false)}catch{}}
  if($doc){try{$doc.Close($false)}catch{}}
  if($word -and $identity -and (Test-OwnedWordIdentity $identity)){try{$word.Quit()}catch{}}
  $report.success=$success
  if($failure){$report.error=$failure}
  $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $reportFull -Encoding UTF8
}
if(-not $success){Write-Error $failure;exit 1}
Write-Output "Body-math $Mode completed; report: $reportFull"

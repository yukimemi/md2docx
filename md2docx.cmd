@echo off
pushd "%~dp0" > nul
set tm=%time: =0%
set ps1file=%~n0___%date:~-10,4%%date:~-5,2%%date:~-2,2%_%tm:~0,2%%tm:~3,2%%tm:~6,2%%tm:~9,2%.ps1
for /f "usebackq skip=10 delims=" %%i in ("%~f0") do @echo %%i >> "%ps1file%"
powershell -NoProfile -ExecutionPolicy unrestricted -File "%ps1file%" %*
del "%ps1file%"
popd > nul
pause
exit %ERRORLEVEL%
# ========== do ps1 file as a dosbatch ==========
# markdown -> docx
param($mdFile = $(Read-Host "Enter markdown file path"))

$ErrorActionPreference = "stop"
$DebugPreference = "Continue" # Continue SilentlyContinue Stop Inquire
# for break point use $host.EnterNestedPrompt()

$commandPath = Split-Path -parent $myInvocation.MyCommand.path
$commandName = Split-Path -leaf $myInvocation.MyCommand.path
$commandBaseName = ((gci $myInvocation.MyCommand.path).BaseName -split "___")[0]

Set-Location $commandPath

<# for test
$mdFile = Join-Path $commandPath "test.md"
#>

# Const variable
$CONST = @{#{{{
  msoConnectorCurve = 3
  msoConnectorElbow = 2
  msoConnectorStraight = 1
  msoConnectorTypeMixed = -2
  msoFalse = 0
  msoLineSingle = 1
  msoLineStyleMixed = -2
  msoLineThickBetweenThin = 5
  msoLineThickThin = 4
  msoLineThinThick = 3
  msoLineThinThin = 2
  msoTrue = -1
  wdAlertsAll = -1
  wdAlertsMessageBox = -2
  wdAlertsNone = 0
  wdAlignParagraphCenter = 1
  wdAlignParagraphLeft = 0
  wdAlignParagraphRight = 2
  wdAutoFitContent = 1
  wdAutoFitFixed = 0
  wdAutoFitWindow = 2
  wdColumnBreak = 8
  wdDoNotSaveChanges = 0
  wdFieldEmpty = -1
  wdFieldPage = 33
  wdHeaderFooterEvenPages = 3
  wdHeaderFooterFirstPage = 2
  wdHeaderFooterPrimary = 1
  wdLineBreak = 6
  wdLineBreakClearLeft = 9
  wdLineBreakClearRight = 10
  wdListApplyToWholeList = 0
  wdBulletGallery = 1
  wdNumberGallery = 2
  wdOutlineNumberGallery = 3
  wdPageBreak = 7
  wdPromptToSaveChanges = -2
  wdSaveChanges = -1
  wdSectionBreakContinuous = 3
  wdSectionBreakEvenPage = 4
  wdSectionBreakNextPage = 2
  wdSectionBreakOddPage = 5
  wdSeekCurrentPageFooter = 10
  wdSeekCurrentPageHeader = 9
  wdSeekMainDocument = 0
  wdTextWrappingBreak = 11
  wdThemeColorAccent1 = 4
  wdThemeColorAccent2 = 5
  wdThemeColorAccent3 = 6
  wdThemeColorAccent4 = 7
  wdThemeColorAccent5 = 8
  wdThemeColorAccent6 = 9
  wdWord8TableBehavior = 0
  wdWord9TableBehavior = 1
  wdTableFormat3DEffects1 = 32
  wdTableFormat3DEffects2 = 33
  wdTableFormat3DEffects3 = 34
  wdTableFormatClassic1 = 4
  wdTableFormatClassic2 = 5
  wdTableFormatClassic3 = 6
  wdTableFormatClassic4 = 7
  wdTableFormatColorful1 = 8
  wdTableFormatColorful2 = 9
  wdTableFormatColorful3 = 10
  wdTableFormatColumns1 = 11
  wdTableFormatColumns2 = 12
  wdTableFormatColumns3 = 13
  wdTableFormatColumns4 = 14
  wdTableFormatColumns5 = 15
  wdTableFormatContemporary = 35
  wdTableFormatElegant = 36
  wdTableFormatGrid1 = 16
  wdTableFormatGrid2 = 17
  wdTableFormatGrid3 = 18
  wdTableFormatGrid4 = 19
  wdTableFormatGrid5 = 20
  wdTableFormatGrid6 = 21
  wdTableFormatGrid7 = 22
  wdTableFormatGrid8 = 23
  wdTableFormatList1 = 24
  wdTableFormatList2 = 25
  wdTableFormatList3 = 26
  wdTableFormatList4 = 27
  wdTableFormatList5 = 28
  wdTableFormatList6 = 29
  wdTableFormatList7 = 30
  wdTableFormatList8 = 31
  wdTableFormatNone = 0
  wdTableFormatProfessional = 37
  wdTableFormatSimple1 = 1
  wdTableFormatSimple2 = 2
  wdTableFormatSimple3 = 3
  wdTableFormatSubtle1 = 38
  wdTableFormatSubtle2 = 39
  wdTableFormatWeb1 = 40
  wdTableFormatWeb2 = 41
  wdTableFormatWeb3 = 42
}#}}}

function rgb($r, $g, $b) {#{{{
  return ($b + ($g * 256) + ($r * 65536))
}#}}}

function randomColor() {#{{{
  $r = $(0..255 | Get-Random)
  $g = $(0..255 | Get-Random)
  $b = $(0..255 | Get-Random)
  return rgb $r $g $b
}#}}}

function checkFilePath($path) {#{{{

  trap { Write-Host "[checkFilePath]: Error $($_)"; throw $_ }

  if (! (Test-Path $path)) {
    Write-Host "$($path) is not found !"
    return $false
  }

  $path = $path -replace """", ""

  # to abs path
  $path = Convert-Path $path

  Write-Debug $path

  return $path

}#}}}

function typeText($line, $word, $doc, $selection, [ref]$commandFlg, [ref]$tableMap, [ref]$listMap) {#{{{

  trap { Write-Host "[typeText]: Error $($_)"; throw $_ }

  # command
  if ($commandFlg.Value) {
    # Check command end
    if ($line -match "^-->$") {
      #Write-Debug "---------- command end ----------"
      # do command
      $commandStr = $commandList.ToArray() -join "`n"
      #Write-Debug $commandStr
      Invoke-Expression $commandStr
      $commandList = New-Object System.Collections.Generic.List[string]
      $commandFlg.Value = $false
      return
    }
    $commandList.Add($line)
    return
  }

  # table
  if ($tableMap.Value.flg) {
    # Check tale end
    if ($line -notmatch "^\|.*\|$") {
      Write-Debug "---------- table end ----------"
      Write-Debug "table row = $($tableMap.Value.row)"
      Write-Debug "table col = $($tableMap.Value.col)"

      $table = $doc.Range([ref]$tableMap.Value.rangeStart, [ref]$tableMap.Value.rangeEnd).ConvertToTable("|", $tableMap.Value.row, $tableMap.Value.col)
      $table.AutoFormat($CONST.wdTableFormatElegant)

      $tableMap.Value.row = 0
      $tableMap.Value.col = 0
      $tableMap.Value.rangeStart = 0
      $tableMap.Value.rangeEnd = 0
      $tableMap.Value.flg = $false
    }
  }

  # Bullet list
  if ($listMap.Value.bflg) {
    # Check list end
    if ($line -notmatch "^\s*\* ") {
      # Reset style
      $selection.Style = $doc.Styles.Item("標準")
      $listMap.Value.bflg = $false
      $listmap.Value.bIndent = 0
    }
  }
  # Number list
  if ($listMap.Value.nflg) {
    # Check list end
    if ($line -notmatch "^\s*[0-9]+\. ") {
      # Reset style
      $selection.Style = $doc.Styles.Item("標準")
      $listMap.Value.nflg = $false
      $listmap.Value.nIndent = 0
    }
  }

  # command
  if ($line -match "<!--!(?<command>.*)-->") {
    $inlineCommand = $matches.command
    $line = $line -replace "<!--!.*-->", ""
  } else {
    $inlineCommand = ""
  }

  switch -regex ($line) {
    # title
    "^#<!--%title--> " {
            $line = $line -replace "^#<!--%title--> ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("表題")
            break
          }
    # subtitle
    "^#<!--%subtitle--> " {
            $line = $line -replace "^#<!--%subtitle--> ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("副題")
            break
          }
    # head 1
    "^# " {
            $line = $line -replace "^# ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 1")
          }
    # head 2
    "^## " {
            $line = $line -replace "^## ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 2")
          }
    # head 3
    "^### " {
            $line = $line -replace "^### ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 3")
          }
    # head 4
    "^#### " {
            $line = $line -replace "^#### ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 4")
          }
    # bullet list
    "^(?<bIndent>\s*)\* " {
            Write-Debug "Bullet indent: $($listmap.Value.bIndent)"
            $line = $line -replace "^\s*\* ", ""
            if (! $listMap.Value.bFlg) {
              $selection.Range.ListFormat.ApplyBulletDefault()
              $listMap.Value.bFlg = $true
            }

            # indent
            $indentCnt = ($matches.bIndent).Length
            if ($indentCnt -ne 0) {
              if ($listMap.Value.indentCnt -eq 0) {
                $listMap.Value.indentCnt = $indentCnt
              }

              $indent = $indentCnt / $listMap.Value.indentCnt
              $selection.ParagraphFormat.LeftIndent = $word.MillimetersToPoints(7.4)
              for ($i = 0; $i -lt $indent; $i++) {
                $selection.Range.ListFormat.ListIndent()
              }
            }

            #if ($indentCnt -gt $listmap.Value.bIndent) {
              #$selection.Range.ListFormat.ListIndent()
            #} elseif ($indentCnt -lt $listmap.Value.bIndent) {
              #$selection.Range.ListFormat.ListOutdent()
            #}
            #$listmap.Value.bIndent = $indentCnt

            $selection.TypeText($line)
          }
    # number list
    "^(?<nIndent>\s*)[0-9]+\. " {
            Write-Debug "Number indent: $($listmap.Value.nIndent)"
            $line = $line -replace "^\s*[0-9]+\. ", ""
            if (! $listmap.Value.nFlg) {
              $selection.Range.ListFormat.ApplyNumberDefault()
              $listMap.Value.nFlg = $true
            }

            if ($listMap.Value.continuous) {
              $selection.Range.ListFormat.ApplyListTemplate($word.ListGalleries.Item($CONST.wdNumberGallery).ListTemplates.Item(1), $true)
            } else {
              $selection.Range.ListFormat.ApplyListTemplate($word.ListGalleries.Item($CONST.wdNumberGallery).ListTemplates.Item(1), $false)
              $listMap.Value.continuous = $true
            }

            # indent
            $indentCnt = ($matches.nIndent).Length
            if ($indentCnt -gt $listmap.Value.nIndent) {
              $selection.Range.ListFormat.ListIndent()
            } elseif ($indentCnt -lt $listmap.Value.nIndent) {
              $selection.Range.ListFormat.ListOutdent()
            }
            $listmap.Value.nIndent = $indentCnt

            $selection.TypeText($line)
          }
    # list continuous
    "<!--% end of list -->$" {
            $listMap.Value.continuous = $false
            $listmap.Value.nIndent = 0
            return
    }
    # image
    "!\[(?<imgName>.*)\]\((?<imgUrl>.*)\)(<!--%(?<width>[0-9]*)x(?<height>[0-9]*)-->|.*)" {
            $imgPath = $(Join-Path (Split-Path -parent $mdFile) $matches.imgUrl)
            if (Test-Path $imgPath) {
              $img = $selection.InlineShapes.AddPicture($imgPath)
              #Write-Debug "Height: [$($img.Height)] Width: [$($img.Width)]"
              $img.LockAspectRatio = $CONST.msoTrue
              if ($matches.width) {
                $img.Width = $matches.width
              }
              if ($matches.height) {
                $img.Height = $matches.height
              }
            } else {
              Write-Error "$($imgPath) is not found !"
            }
          }
    # page break
    "^<!--%(\[改ページ\]|\[PageBreak\])-->" {
            $selection.InsertBreak()
            return
          }
    # section break
    "^<!--%(\[改セクション\]|\[SectionBreak\])-->" {
            $selection.InsertBreak($CONST.wdSectionBreakNextPage)
            return
          }
    # comment command start
    "^<!--!$" {
            #Write-Debug "---------- command start ----------"
            $commandFlg.Value = $true
            return
          }
    # table
    "^\|.*\|$" {
            Write-Debug "---------- table start ----------"
            $lineArray = $line.SubString(1, $line.Length -2) -split "\|"
            $tableMap.Value.col = $lineArray.Length
            $tableMap.Value.row++

            if (! $tableMap.Value.flg) {
              $tableMap.Value.rangeStart = $selection.Start
            }

            $selection.TypeText($line.SubString(1, $line.Length -2))
            $tableMap.Value.rangeEnd = $selection.End
            $tableMap.Value.flg = $true
          }
    # comment
    "<!--[^!].*-->" {
            # do nothing
            #Write-Debug "comment"
          }
    # other
    default {
            $selection.Style = $doc.Styles.Item("標準")
            $selection.TypeText($line)
          }
  }

  # Do inline command
  if ($inlineCommand -ne "") {
    Invoke-Expression $inlineCommand
  }

  $selection.TypeParagraph()

}#}}}

function main() {#{{{

  trap { Write-Host "[main]: Error $($_)"; throw $_ }

  $commandList = New-Object System.Collections.Generic.List[string]
  $commandFlg = $false
  $tableMap = @{
    flg = $false
    row = 0
    col = 0
    rangeStart = 0
    rangeEnd = 0
  }
  $listMap = @{
    bflg = $false
    nflg = $false
    continuous = $true
    indentCnt = 0
    bIndent = 0
    nIndent = 0
  }

  # check file
  $mdFile = checkFilePath $mdFile
  if (! $mdFile) {
    return -1
  }

  Write-Host "Executing..."

  # Open docx
  try {

    $word = New-Object -ComObject Word.Application
    $word.Application.DisplayAlerts = $CONST.wdAlertsNone
    $word.Visible = $true

    $doc = $word.Documents.Add()
    $selection = $word.Selection

    gc -Encoding UTF8 $mdFile | % {
      #Write-Debug $_
      typeText $_ $word $doc $selection ([ref]$commandFlg) ([ref]$tableMap) ([ref]$listMap)
    }

    # reset style
    $selection.Style = $doc.Styles.Item("標準")

    $mdFileInfo = gci $mdFile
    $doc.SaveAs([ref]$(Join-Path $mdFileInfo.DirectoryName ($mdFileInfo.BaseName + ".docx")))

  } catch [Exception] {
    throw "[main]: Error $($_)"
  } finally {
    if ($word) {
      $word.Quit([ref]$CONST.wdDoNotSaveChanges)
    }
  }

} #}}}

# call main
Measure-Command { main }

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
$CONST = @{
  wdOutlineNumberGallery = 3
  wdListApplyToWholeList = 0
  msoFalse = 0
  msoTrue = -1
  wdAlignParagraphRight = 2
  wdSeekMainDocument = 0
  wdSeekCurrentPageHeader = 9
  wdSeekCurrentPageFooter = 10
  wdFieldEmpty = -1
  wdFieldPage = 33
}

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

function typeText($line, $word, $doc, $selection, [ref]$commandFlg) {#{{{

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

  switch -regex ($line) {
    # title
    "^#<!--!title--> " {
            $line = $line -replace "^#<!--!title--> ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("表題")
            $selection.TypeParagraph()
          }
    # subtitle
    "^#<!--!subtitle--> " {
            $line = $line -replace "^#<!--!subtitle--> ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("副題")
            $selection.TypeParagraph()
          }
    # head 1
    "^# " {
            $line = $line -replace "^# ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 1")
            $selection.TypeParagraph()
          }
    # head 2
    "^## " {
            $line = $line -replace "^## ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 2")
            $selection.TypeParagraph()
          }
    # head 3
    "^### " {
            $line = $line -replace "^### ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 3")
            $selection.TypeParagraph()
          }
    # head 4
    "^#### " {
            $line = $line -replace "^#### ", ""
            $selection.TypeText($line)
            $selection.Style = $doc.Styles.Item("見出し 4")
            $selection.TypeParagraph()
          }
    # bullet list
    "^\* " {
            $line = $line -replace "^\* ", ""
            $selection.Range.ListFormat.ApplyBulletDefault()
            $selection.TypeText($line)
            $selection.TypeParagraph()
            $selection.Style = $doc.Styles.Item("標準")
          }
    # number list
    "^[0-9]+\. " {
            $line = $line -replace "^[0-9]+\. ", ""
            $selection.Range.ListFormat.ApplyNumberDefault()
            $selection.TypeText($line)
            $selection.TypeParagraph()
            $selection.Style = $doc.Styles.Item("標準")
          }
    # image
    "!\[(?<imgName>.*)\]\((?<imgUrl>.*)\)(<!--!(?<width>[0-9]*)x(?<height>[0-9]*)-->|.*)" {
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
            $selection.TypeParagraph()
          }
    # page break
    "^<!--!(\[改ページ\]|\[PageBreak\])-->" {
            $selection.InsertBreak()
            break
          }
    # comment command start
    "^<!--!$" {
            #Write-Debug "---------- command start ----------"
            $commandFlg.Value = $true
          }
    # comment
    "<!--[^!].*-->" {
            # do nothing
            #Write-Debug "comment"
          }
    # other
    default {
            $selection.Style = $doc.Styles.Item("標準")
            # comment command
            if ($line -match "<!--!(?<command>.*)-->") {
              $selection.TypeText($($line -replace "<!--!.*-->", ""))
              Write-Debug $matches.command
              Invoke-Expression $matches.command
            } else {
              $selection.TypeText($line)
            }
            $selection.TypeParagraph()
          }
  }

}#}}}

function main() {#{{{

  trap { Write-Host "[main]: Error $($_)"; throw $_ }

  $commandList = New-Object System.Collections.Generic.List[string]
  $commandFlg = $false

  # check file
  $mdFile = checkFilePath $mdFile
  if (! $mdFile) {
    return -1
  }

  Write-Host "Executing..."

  # Open docx
  try {

    $word = New-Object -ComObject Word.Application
    #$word.Visible = $true

    $doc = $word.Documents.Add()
    $selection = $word.Selection

    gc -Encoding UTF8 $mdFile | % {
      #Write-Debug $_
      typeText $_ $word $doc $selection ([ref]$commandFlg)
    }

    $mdFileInfo = gci $mdFile
    $doc.SaveAs([ref]$(Join-Path $mdFileInfo.DirectoryName ($mdFileInfo.BaseName + ".docx")))

  } catch [Exception] {
    throw "[main]: Error $($_)"
  } finally {
    if ($word) {
      $word.Quit()
    }
  }

} #}}}

# call main
Measure-Command { main }


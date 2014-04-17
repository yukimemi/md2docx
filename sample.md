#<!--!title--> md2docx サンプル
#<!--!subtitle--> 〜MarkdownでWordを書こう〜
Author: yukimemi<!--!$selection.ParagraphFormat.Alignment = $CONST.wdAlignParagraphRight-->

# Markdownは素晴らしい！

## 概要
Wordは素晴らしい機能を持っていますが、なんかいろいろ難しいです。
そこで、Markdownを使って文書を作成し、Wordに変換します。

## 準備

1. yukimemi/md2docx https://github.com/yukimemi/md2docx からmd2docx.cmd をダウンロードします。
2. コマンドプロンプトより、md2docx.cmdを起動します。
3. docxに変換したいMarkdownファイルをドラッグアンドドロップしてEnterします。

## 特徴

* 基本的なMarkdown記法は使える。
* wordマクロがPowerShell記述で使える。
* PowerShellなので、Windowsユーザーならデフォルトで使用出来る。

## TODO

* 処理速度向上
* Wordテンプレートの使用

<!--!

  # Style change
  $doc.Styles.Item("表題").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("表題").Font.Size = 20
  $doc.Styles.Item("表題").Font.Bold = $true

  $doc.Styles.Item("副題").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("副題").Font.Size = 18

  $doc.Styles.Item("見出し 1").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("見出し 1").Font.Size = 16
  $doc.Styles.Item("見出し 1").Font.Bold = $true
  $doc.Styles.Item("見出し 1").AutomaticallyUpdate = $true

  $doc.Styles.Item("見出し 2").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("見出し 2").Font.Size = 14
  $doc.Styles.Item("見出し 2").AutomaticallyUpdate = $true

  $doc.Styles.Item("見出し 3").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("見出し 3").Font.Size = 12
  $doc.Styles.Item("見出し 3").AutomaticallyUpdate = $true

  $word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1).ListLevels.Item(1).Numberformat = "%1"
  $word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1).ListLevels.Item(2).Numberformat = "%1.%2"
  $word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1).ListLevels.Item(3).Numberformat = "%1.%2.%3"

  # Outline liked style
  $doc.Paragraphs | % {
    if ($_.Style.NameLocal -eq "見出し 1") {
      $_.Range.ListFormat.ApplyListTemplate($word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1), $true)
    } elseif ($_.Style.NameLocal -eq "見出し 2") {
      $_.Range.ListFormat.ApplyListTemplate($word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1), $true)
    } elseif ($_.Style.NameLocal -eq "見出し 3") {
      $_.Range.ListFormat.ApplyListTemplate($word.ListGalleries.Item($CONST.wdOutlineNumberGallery).ListTemplates.Item(1), $true)
    }
  }

  # Image size change
  $doc.InlineShapes | % {
    if ($_.Height -gt 200 ) {
      $_.LockAspectRatio = $CONST.msoTrue; $_.Height = 200
    }
  }

  # page setup
  $doc.PageSetup.TopMargin = $word.MillimetersToPoints(25)
  $doc.PageSetup.BottomMargin = $word.MillimetersToPoints(20)
  $doc.PageSetup.LeftMargin = $word.MillimetersToPoints(20)
  $doc.PageSetup.RightMargin = $word.MillimetersToPoints(15)
  $doc.PageSetup.Gutter = $word.MillimetersToPoints(0)

  $doc.PageSetup.HeaderDistance = $word.MillimetersToPoints(15)
  $doc.PageSetup.FooterDistance = $word.MillimetersToPoints(15)

  # header
  $word.ActiveWindow.ActivePane.View.SeekView = $CONST.wdSeekCurrentPageHeader
  $selection.TypeText("MarkdownでWordを書こう")
  $selection.ParagraphFormat.Alignment = $CONST.wdAlignParagraphRight
  $word.ActiveWindow.ActivePane.View.SeekView = $CONST.wdSeekMainDocument

  # footer
  $word.ActiveWindow.ActivePane.View.SeekView = $CONST.wdSeekCurrentPageFooter
  $selection.Fields.Add($selection.Range, $CONST.wdFieldPage, "#PAGE")
  $selection.ParagraphFormat.Alignment = $CONST.wdAlignParagraphRight
  $word.ActiveWindow.ActivePane.View.SeekView = $CONST.wdSeekMainDocument

-->

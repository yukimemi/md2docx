md2docx
=======

convert markdown to docx

markdownファイルをwordのdocxに変換するスクリプトです。
* PowerShellスクリプトなのでWindows限定です。
* 動作確認はOffice 2010でのみ行っています。
* 確認出来ているのは、Windows 7 + PowerShell V3です。

使い方
------

バッチファイル化してあるので、そちらを使用します。
コマンドプロンプトより、以下のように入力します。

```Batchfile
> md2docx.cmd file.md
```

拡張子が **docx** に変更された、 **file.docx** が同じディレクトリに出来ているはずです。

もしくは、引数に何も指定しなければ、markdown file の入力を促すプロンプトが表示されるので、そこにエクスプローラなどから markdownファイルをドラッグアンドドロップしてEnterでもよいです。

対応記述形式
------------

markdownといっても、現在対応している記述はまだまだ少ないです。
また、word独自対応するためにmarkdownにおけるコメントを使用している部分があります。

### 見出し

```Markdown
# 見出し 1
## 見出し 2
### 見出し 3
#### 見出し 4
```

それぞれ、wordのスタイル [見出し 1] 〜 [見出し 4]に対応しています。

### 表題、副題

```Markdown
#<!--!title--> 表題
#<!--!subtitle--> 副題
```

それぞれ、wordのスタイル [表題]、[副題]に対応しています。

### リスト

```Markdown
* リスト 1
* リスト 2
* リスト 3

1. リスト 1
2. リスト 2
3. リスト 3
```

通常のリストと、番号付きリストです。

### 画像

```Markdown
![代替テキスト](imgfolder/img.jpg)
![代替テキスト](imgfolder/img.jpg)<!--!200x100-->
![代替テキスト](imgfolder/img.jpg)<!--!300x-->
![代替テキスト](imgfolder/img.jpg)<!--!x200-->
```

画像の場所は、markdownファイルのある場所からの相対パスで記述します。
また、後ろに上記のように画像のサイズを指定することが可能です。
画像のサイズは縦横比固定なので、両方指定しても恐らくどちらか無視されます。
なので、片方だけ指定することも可能です。

### 改ページ

```Markdown
<!--![改ページ]-->
<!--![PageBreak]-->
```

改ページを挿入します。

### インラインマクロ

```Markdown
右寄せ文章<!--!$selection.ParagraphFormat.Alignment = $CONST.wdAlignParagraphRight-->
```

`<!--! word マクロ -->` でwordマクロ(PowerShell記述)が実行出来ます。

該当文章のみにマクロを指定したい場合に便利です。
この場合、「右寄せ文章」が右寄せになります。

提供されるwordのCOMオブジェクトは、以下の通りです。

```Markdown
$word      : Word オブジェクト
$doc       : Word Document オブジェクト
$selection : Word Selection オブジェクト
```

また、wordで使用出来る定数は、`$CONST.定数名`でアクセス出来ます。(まだアクセス出来る定数は少ない・・・)

### 複数行マクロ

```PowerShell
<!--!
  # Style change
  $doc.Styles.Item("表題").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("表題").Font.Size = 20
  $doc.Styles.Item("表題").Font.Bold = $true

  $doc.Styles.Item("見出し 1").Font.NameFarEast = "ＭＳ Ｐ明朝"
  $doc.Styles.Item("見出し 1").Font.Size = 16
  $doc.Styles.Item("見出し 1").Font.Bold = $true
  $doc.Styles.Item("見出し 1").AutomaticallyUpdate = $true

  # Page setup
  $doc.PageSetup.TopMargin = $word.MillimetersToPoints(25)
  $doc.PageSetup.BottomMargin = $word.MillimetersToPoints(20)
  $doc.PageSetup.LeftMargin = $word.MillimetersToPoints(20)
  $doc.PageSetup.RightMargin = $word.MillimetersToPoints(15)

  # Image size change
  $doc.InlineShapes | % {
    if ($_.Height -gt 200 ) {
      $_.LockAspectRatio = $CONST.msoTrue; $_.Height = 200
    }
  }
-->
```

`<!--!`で始めて、`-->`で終わる、この記述内はすべてPowerShellとして実行されます。
インラインマクロと同様、`$word`、`$doc`、`$selection`が使用出来ます。

`<!--!`と、`-->`はそれぞれ、単独で行頭に記述されている必要があります。

このように、結局書式の設定部分は、wordマクロをガリガリ書くことになります・・・。
その代わり、wordマクロを(PowerShellにて)記述出来るので、その気になればなんでも出来る！はずです。

TODO
----

* テーブル
* 引用
* 処理速度アップ


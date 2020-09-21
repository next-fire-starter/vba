Attribute VB_Name = "OutputImageList"
'----------------------------------------------------------------------------------------目的
'指定フォルダのイメージを順番に貼り付ける
'
'・イメージのサイズは元のサイズで貼り付ける
'・貼り付けたシートだけをを別名で保存する
'----------------------------------------------------------------------------------------準備
'メニューから「ツール→参照設定」とたどり、「参照可能なライブラリファイル」から
'「Microsoft Scripting Runtime」にチェックを付けて「OK」ボタンをクリックする
'次の名前を付けたシートを作成　"設定","貼り付け先"
'"設定"シートには以下の情報を入力する
'E5→イメージがあるフォルダ      (例)C:\Image\Test
'E6→Excel出力先フォルダ名       (例)C:\OutputExcel
'E7→Excel出力先ファイル名       (例)abc.xlsx
'E8→Excel出力先ファイルシート名 (例)Case001

Sub OutputImageList()
    Dim lngTop As Long
    Dim objFile As File
    Dim objFldr As FileSystemObject
    Dim wshSet  As Worksheet
    
    Set wshConf = ThisWorkbook.Sheets("設定")

    Application.DisplayAlerts = False

    Set objFldr = CreateObject("Scripting.FileSystemObject")

    '初期位置設定
    lngTop = 50
    
    '貼り付け用シート表示
    ThisWorkbook.Sheets("貼り付け先").Select
    
    'イメージ貼り付け
    For Each objFile In objFldr.GetFolder(wshConf.Range("E5").Value).Files

        Set shapePic = ActiveSheet.Shapes.AddPicture( _
          Filename:=objFile, _
          LinkToFile:=False, _
          SaveWithDocument:=True, _
          Left:=0, _
          Top:=lngTop, _
          Width:=0, _
          Height:=0)
          
        '挿入した画像に対して元画像と同じ高さ・幅にする
        With shapePic
            .ScaleHeight 1, msoTrue
            .ScaleWidth 1, msoTrue
            lngTop = lngTop + CLng(.Height) + 50
        End With
    
    Next
    
    '新規Book
    Workbooks.Add
    ThisWorkbook.Sheets("貼り付け先").Copy after:=ActiveWorkbook.Sheets(Sheets.Count)

    '初期設定シート削除
    For Each sht In ActiveWorkbook.Sheets
        If ActiveWorkbook.Sheets.Count > 1 Then
            ActiveWorkbook.Sheets(1).Delete
        End If
    Next
    
    'シート名設定
    ActiveWorkbook.Sheets(1).Name = wshConf.Range("E8").Value
    'Excel出力先へ保存
    ActiveWorkbook.SaveAs wshConf.Range("E6").Value & "\" & wshConf.Range("E7").Value
    ActiveWorkbook.Close
    
    'コピー元のシートからイメージを削除する
    ThisWorkbook.Sheets("貼り付け先").Delete
    ThisWorkbook.Sheets.Add after:=Worksheets(Worksheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "貼り付け先"

    '設定シート表示
    wshConf.Select
    
    MsgBox ("処理完了")
    Application.DisplayAlerts = True
    
End Sub



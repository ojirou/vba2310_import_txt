Attribute VB_Name = "Module1"
Sub ImportTxtTest()
    Dim TextFile As String
    Dim ws As Worksheet
    TextFile = Environ("USERPROFILE") & "\Desktop\sample.txt" ' 元のファイル
    Set ws = ThisWorkbook.Worksheets("固定資産台帳_取込用")
'    Call ClearCellRange(ws)
    Call ImportTxt(ws, TextFile)
End Sub
Sub ClearCellRangeTest()
    Dim TextFile As String
    Dim ws As Worksheet
    TextFile = Environ("USERPROFILE") & "\Desktop\sample.txt" ' 元のファイル
    Set ws = ThisWorkbook.Worksheets("固定資産台帳_取込用")
    Call ClearCellRange(ws)
'    Call ImportTxt(ws, TextFile)
End Sub
'#############################################################################
' テキストファイルを取込み
'
'　import_txt
'#############################################################################
Sub ImportTxt(ByRef ws As Worksheet, ByRef TextFile As String)
    Dim tempRow As Long
    ' 出力開始行を指定
    tempRow = 4
    ' 出力開始列を指定
    Dim strtCol As Long
    strtCol = 1
    Dim FileContent As String
    Dim NewContent As String
    Dim FileNumber As Integer
    ' テキストファイルを読み込む
    FileNumber = FreeFile
    Open TextFile For Binary As #FileNumber
    FileContent = Space$(LOF(FileNumber))
    Get #FileNumber, , FileContent
    Close #FileNumber
    ' ダブルクォーテーション3つ連続をダブルクォーテーション1つに置換
    NewContent = Replace(FileContent, "", """""")
    ' 対象のテキストファイルを読み込むために開く
    FileNumber = FreeFile ' 空いているファイル番号を取得
    Open TextFile For Input As #FileNumber
    Do Until EOF(FileNumber) ' ファイル番号を使うように変更
        ' 変数「buf」に1行分のデータを格納
        Line Input #FileNumber, buf ' ファイル番号を使うように変更
        ' ダブルクォーテーションで囲まれたテキストデータを処理
        buf = Replace(buf, """", "") ' ダブルクォーテーションを削除
        ' テキストファイル対象行のデータを配列に格納
        tmpAry = Split(buf, ",")
        ' 配列の要素の数を変数「indexNum」に定義
        indexNum = UBound(tmpAry) - LBound(tmpAry) + 1
        ' 出力セル範囲を変数「tempRng」に定義して出力
        Set tempRng = ws.Range(ws.Cells(tempRow, strtCol), ws.Cells(tempRow, strtCol + indexNum - 1))
        ' 一時的に保護を解除してから値を書き込む
        ws.Unprotect
        tempRng.Value = tmpAry
'        tempSh.Protect ' 保護を再度設定
        tempRow = tempRow + 1
    Loop
    MsgBox "最終行は " & tempRow - 1 & " です。"
    Close #FileNumber
End Sub
'#############################################################################
' 取込シートのデータをクリア
'
'　clear_cell_range
'#############################################################################
Sub ClearCellRange(ByRef ws As Worksheet)
    Dim LastRow As Long
    ' A列（4行目から始まり、V列まで）のデータを一括クリア
    With ws
        LastRow = .Cells(.Rows.Count, "V").End(xlUp).Row
        If LastRow >= 4 Then
            .Range("A4:V" & LastRow).ClearContents
        End If
    End With
End Sub

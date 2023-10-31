Attribute VB_Name = "Module1"
Sub ImportTxtTest()
    Dim TextFile As String
    Dim ws As Worksheet
    TextFile = Environ("USERPROFILE") & "\Desktop\sample.txt" ' ���̃t�@�C��
    Set ws = ThisWorkbook.Worksheets("�Œ莑�Y�䒠_�捞�p")
'    Call ClearCellRange(ws)
    Call ImportTxt(ws, TextFile)
End Sub
Sub ClearCellRangeTest()
    Dim TextFile As String
    Dim ws As Worksheet
    TextFile = Environ("USERPROFILE") & "\Desktop\sample.txt" ' ���̃t�@�C��
    Set ws = ThisWorkbook.Worksheets("�Œ莑�Y�䒠_�捞�p")
    Call ClearCellRange(ws)
'    Call ImportTxt(ws, TextFile)
End Sub
'#############################################################################
' �e�L�X�g�t�@�C�����捞��
'
'�@import_txt
'#############################################################################
Sub ImportTxt(ByRef ws As Worksheet, ByRef TextFile As String)
    Dim tempRow As Long
    ' �o�͊J�n�s���w��
    tempRow = 4
    ' �o�͊J�n����w��
    Dim strtCol As Long
    strtCol = 1
    Dim FileContent As String
    Dim NewContent As String
    Dim FileNumber As Integer
    ' �e�L�X�g�t�@�C����ǂݍ���
    FileNumber = FreeFile
    Open TextFile For Binary As #FileNumber
    FileContent = Space$(LOF(FileNumber))
    Get #FileNumber, , FileContent
    Close #FileNumber
    ' �_�u���N�H�[�e�[�V����3�A�����_�u���N�H�[�e�[�V����1�ɒu��
    NewContent = Replace(FileContent, "", """""")
    ' �Ώۂ̃e�L�X�g�t�@�C����ǂݍ��ނ��߂ɊJ��
    FileNumber = FreeFile ' �󂢂Ă���t�@�C���ԍ����擾
    Open TextFile For Input As #FileNumber
    Do Until EOF(FileNumber) ' �t�@�C���ԍ����g���悤�ɕύX
        ' �ϐ��ubuf�v��1�s���̃f�[�^���i�[
        Line Input #FileNumber, buf ' �t�@�C���ԍ����g���悤�ɕύX
        ' �_�u���N�H�[�e�[�V�����ň͂܂ꂽ�e�L�X�g�f�[�^������
        buf = Replace(buf, """", "") ' �_�u���N�H�[�e�[�V�������폜
        ' �e�L�X�g�t�@�C���Ώۍs�̃f�[�^��z��Ɋi�[
        tmpAry = Split(buf, ",")
        ' �z��̗v�f�̐���ϐ��uindexNum�v�ɒ�`
        indexNum = UBound(tmpAry) - LBound(tmpAry) + 1
        ' �o�̓Z���͈͂�ϐ��utempRng�v�ɒ�`���ďo��
        Set tempRng = ws.Range(ws.Cells(tempRow, strtCol), ws.Cells(tempRow, strtCol + indexNum - 1))
        ' �ꎞ�I�ɕی���������Ă���l����������
        ws.Unprotect
        tempRng.Value = tmpAry
'        tempSh.Protect ' �ی���ēx�ݒ�
        tempRow = tempRow + 1
    Loop
    MsgBox "�ŏI�s�� " & tempRow - 1 & " �ł��B"
    Close #FileNumber
End Sub
'#############################################################################
' �捞�V�[�g�̃f�[�^���N���A
'
'�@clear_cell_range
'#############################################################################
Sub ClearCellRange(ByRef ws As Worksheet)
    Dim LastRow As Long
    ' A��i4�s�ڂ���n�܂�AV��܂Łj�̃f�[�^���ꊇ�N���A
    With ws
        LastRow = .Cells(.Rows.Count, "V").End(xlUp).Row
        If LastRow >= 4 Then
            .Range("A4:V" & LastRow).ClearContents
        End If
    End With
End Sub

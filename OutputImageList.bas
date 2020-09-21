Attribute VB_Name = "OutputImageList"
'----------------------------------------------------------------------------------------�ړI
'�w��t�H���_�̃C���[�W�����Ԃɓ\��t����
'
'�E�C���[�W�̃T�C�Y�͌��̃T�C�Y�œ\��t����
'�E�\��t�����V�[�g��������ʖ��ŕۑ�����
'----------------------------------------------------------------------------------------����
'���j���[����u�c�[�����Q�Ɛݒ�v�Ƃ��ǂ�A�u�Q�Ɖ\�ȃ��C�u�����t�@�C���v����
'�uMicrosoft Scripting Runtime�v�Ƀ`�F�b�N��t���āuOK�v�{�^�����N���b�N����
'���̖��O��t�����V�[�g���쐬�@"�ݒ�","�\��t����"
'"�ݒ�"�V�[�g�ɂ͈ȉ��̏�����͂���
'E5���C���[�W������t�H���_      (��)C:\Image\Test
'E6��Excel�o�͐�t�H���_��       (��)C:\OutputExcel
'E7��Excel�o�͐�t�@�C����       (��)abc.xlsx
'E8��Excel�o�͐�t�@�C���V�[�g�� (��)Case001

Sub OutputImageList()
    Dim lngTop As Long
    Dim objFile As File
    Dim objFldr As FileSystemObject
    Dim wshSet  As Worksheet
    
    Set wshConf = ThisWorkbook.Sheets("�ݒ�")

    Application.DisplayAlerts = False

    Set objFldr = CreateObject("Scripting.FileSystemObject")

    '�����ʒu�ݒ�
    lngTop = 50
    
    '�\��t���p�V�[�g�\��
    ThisWorkbook.Sheets("�\��t����").Select
    
    '�C���[�W�\��t��
    For Each objFile In objFldr.GetFolder(wshConf.Range("E5").Value).Files

        Set shapePic = ActiveSheet.Shapes.AddPicture( _
          Filename:=objFile, _
          LinkToFile:=False, _
          SaveWithDocument:=True, _
          Left:=0, _
          Top:=lngTop, _
          Width:=0, _
          Height:=0)
          
        '�}�������摜�ɑ΂��Č��摜�Ɠ��������E���ɂ���
        With shapePic
            .ScaleHeight 1, msoTrue
            .ScaleWidth 1, msoTrue
            lngTop = lngTop + CLng(.Height) + 50
        End With
    
    Next
    
    '�V�KBook
    Workbooks.Add
    ThisWorkbook.Sheets("�\��t����").Copy after:=ActiveWorkbook.Sheets(Sheets.Count)

    '�����ݒ�V�[�g�폜
    For Each sht In ActiveWorkbook.Sheets
        If ActiveWorkbook.Sheets.Count > 1 Then
            ActiveWorkbook.Sheets(1).Delete
        End If
    Next
    
    '�V�[�g���ݒ�
    ActiveWorkbook.Sheets(1).Name = wshConf.Range("E8").Value
    'Excel�o�͐�֕ۑ�
    ActiveWorkbook.SaveAs wshConf.Range("E6").Value & "\" & wshConf.Range("E7").Value
    ActiveWorkbook.Close
    
    '�R�s�[���̃V�[�g����C���[�W���폜����
    ThisWorkbook.Sheets("�\��t����").Delete
    ThisWorkbook.Sheets.Add after:=Worksheets(Worksheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "�\��t����"

    '�ݒ�V�[�g�\��
    wshConf.Select
    
    MsgBox ("��������")
    Application.DisplayAlerts = True
    
End Sub



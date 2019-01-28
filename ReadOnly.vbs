'Option Explicit
'On Error Resume Next

Dim strFileName     '�t�@�C���p�X���i�[
Dim objApp          '�I�u�W�F�N�g�𐶐�
Dim objFileSys      '�I�u�W�F�N�g�𐶐��i�t�@�C���V�X�e���j
Dim strExtension    '�g���q���i�[
Dim fileName        '�t�@�C�������i�[
Dim shortFileName   '�V���[�g�p�X���i�[


'�t�@�C���V�X�e���������I�u�W�F�N�g�𐶐�
Set objFileSys = CreateObject("Scripting.FileSystemObject")

For i=0 to Wscript.Arguments.Count-1
    '�t�@�C���p�X���擾
    strFileName = Wscript.Arguments(i)
    fileName = Mid(Wscript.Arguments(i), InStrRev(Wscript.Arguments(i), "\") + 1)
    shortFileName = objFileSys.GetFile(strFileName).ShortPath

    '�g���q���擾
    strExtension = objFileSys.GetExtensionName(shortFileName)
    '�g���q��\��
    'Wscript.Echo strExtension
    'Wscript.Echo "Lpath:" &  strFileName
    'Wscript.Echo "LLen:" &  Len(strFileName)
    '�V���[�g�p�X��\��
    'Wscript.Echo "Spath:" &  strFileName
    'strExtension = objFileSys.GetExtensionName(shortFileName)


    '�t�@�C���̑��݊m�F
    If Not objFileSys.FileExists(shortFileName) Then
        Wscript.Echo "File��������܂���"
    End If

    If (strExtension = "xls") OR (strExtension = "xlsx") OR (strExtension = "xlsm") OR _
       (strExtension = "XLS") OR (strExtension = "XLSX") OR (strExtension = "XLSM") then
       'Excel�֘A����
       '�N��
       Set objApp = Wscript.CreateObject("Excel.Application")
       '��ʕ\��
       objApp.Visible = True
       '�ǂݎ���p�ŊJ��(Excel)
       Call objApp.Workbooks.Open(shortFileName,,True)

    ElseIf (strExtension = "doc") OR (strExtension = "docx") OR _
           (strExtension = "DOC") OR (strExtension = "DOCX") then
       'Word�֘A����
       '�N��
       Set objApp = WScript.CreateObject("Word.Application")
       '��ʕ\��
       objApp.Visible = True
       '�ǂݎ���p�ŊJ��(Word)
       Call objApp.Documents.Open(shortFileName,,True)

    ElseIf (strExtension = "ppt") OR (strExtension = "pptx") OR _
           (strExtension = "PPT") OR (strExtension = "PPTX") then
       'PowerPoint�֘A����
       '�N��
       Set objApp = Wscript.CreateObject("Powerpoint.Application")
       '��ʕ\��
       objApp.Visible = True
       '�ǂݎ���p�ŊJ��(PowerPoint)
       Call objApp.Presentations.Open(shortFileName,True)

    Else
        Wscript.Echo "�g���q " & strExtension & " �͑ΏۊO�̃t�@�C���ł�"
    End If

    '�I������
    Set objWshNetwork = Nothing
    Set objApp = Nothing
Next

Wscript.Quit

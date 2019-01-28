'Option Explicit
'On Error Resume Next

Dim strFileName     'ファイルパスを格納
Dim objApp          'オブジェクトを生成
Dim objFileSys      'オブジェクトを生成（ファイルシステム）
Dim strExtension    '拡張子を格納
Dim fileName        'ファイル名を格納
Dim shortFileName   'ショートパスを格納


'ファイルシステムを扱うオブジェクトを生成
Set objFileSys = CreateObject("Scripting.FileSystemObject")

For i=0 to Wscript.Arguments.Count-1
    'ファイルパスを取得
    strFileName = Wscript.Arguments(i)
    fileName = Mid(Wscript.Arguments(i), InStrRev(Wscript.Arguments(i), "\") + 1)
    shortFileName = objFileSys.GetFile(strFileName).ShortPath

    '拡張子を取得
    strExtension = objFileSys.GetExtensionName(shortFileName)
    '拡張子を表示
    'Wscript.Echo strExtension
    'Wscript.Echo "Lpath:" &  strFileName
    'Wscript.Echo "LLen:" &  Len(strFileName)
    'ショートパスを表示
    'Wscript.Echo "Spath:" &  strFileName
    'strExtension = objFileSys.GetExtensionName(shortFileName)


    'ファイルの存在確認
    If Not objFileSys.FileExists(shortFileName) Then
        Wscript.Echo "Fileが見つかりません"
    End If

    If (strExtension = "xls") OR (strExtension = "xlsx") OR (strExtension = "xlsm") OR _
       (strExtension = "XLS") OR (strExtension = "XLSX") OR (strExtension = "XLSM") then
       'Excel関連処理
       '起動
       Set objApp = Wscript.CreateObject("Excel.Application")
       '画面表示
       objApp.Visible = True
       '読み取り専用で開く(Excel)
       Call objApp.Workbooks.Open(shortFileName,,True)

    ElseIf (strExtension = "doc") OR (strExtension = "docx") OR _
           (strExtension = "DOC") OR (strExtension = "DOCX") then
       'Word関連処理
       '起動
       Set objApp = WScript.CreateObject("Word.Application")
       '画面表示
       objApp.Visible = True
       '読み取り専用で開く(Word)
       Call objApp.Documents.Open(shortFileName,,True)

    ElseIf (strExtension = "ppt") OR (strExtension = "pptx") OR _
           (strExtension = "PPT") OR (strExtension = "PPTX") then
       'PowerPoint関連処理
       '起動
       Set objApp = Wscript.CreateObject("Powerpoint.Application")
       '画面表示
       objApp.Visible = True
       '読み取り専用で開く(PowerPoint)
       Call objApp.Presentations.Open(shortFileName,True)

    Else
        Wscript.Echo "拡張子 " & strExtension & " は対象外のファイルです"
    End If

    '終了処理
    Set objWshNetwork = Nothing
    Set objApp = Nothing
Next

Wscript.Quit

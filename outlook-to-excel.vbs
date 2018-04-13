' 受信トレイ フォルダの定数定義
' https://msdn.microsoft.com/ja-jp/vba/outlook-vba/articles/oldefaultfolders-enumeration-outlook
Const FOLDER_INDEX = 6

Set config  = CreateObject("Scripting.Dictionary")

' 受信トレイの下層フォルダ。受信トレイを抽出対象としたい場合は空にする
config.Add "outlook_sub_folder", ""
' 追記するExcelファイルまでのパス
config.Add "excel_file_path", "D:\sample.xlsx"
' 追記するExcelファイルのワークシート名
config.Add "excel_worksheet_name", "Sheet1"
' 抽出対象の件名。設定した文字列が含まれるメールを対象とする
config.Add "pickup_title", "件名"
' 抽出対象件数
config.Add "max_pickup_num", "100"
' 本文抽出条件
config.Add "body_grep_rule", "■(.+)$"
' Excelに追記する最初の列値(A~Zまで)
config.Add "excel_start_row", "A"
' Excelに追記する最初の行。値が0の場合は最初にヒットした空セルの行を起点とする
config.Add "excel_start_line", 0

'
' メイン処理
'
' @params Scripting.Dictionary config 設定値
'
Function run(config)
    Set mailFolder = folder()

    If mailFolder Is Nothing Then
        Exit Function
    End If

    Set excelFile = excel(config)

    If excelFile Is Nothing Then
        Exit Function
    End If

    Set worksheet = excelFile.Worksheets(config("excel_worksheet_name"))

    If worksheet Is Nothing Then
        excelFile.Quit
        Exit Function
    End If

    line = startLine(worksheet, config)
    message = ""

    For i = 1 To config("max_pickup_num")
        If mailFolder.Items.Count < i Then
            Exit For
        End If

        Set item = mailFolder.Items(i)

        If condition(item, config) Then
            result = write(worksheet, line, contents(item.body, config), config)
            line = line + 1
            message = message + "・" + item.Subject + vbCrLf
        End If
    Next

    On Error Resume Next
    excelFile.Save
    If Err.Number Then
        MsgBox "Excelの保存に失敗しました"
    Else
        IF message = "" Then
            MsgBox "登録対象が1件も存在しませんでした"
        Else
            MsgBox "以下メールの内容をExcelに追記しました" + vbCrLf + message
        End If
    End If
    On Error Goto 0

    excelFile.Quit
End Function

'
' 抽出対象のフォルダ返却
' @return Outlook.Application Outlookの抽出対象フォルダ
'
Function folder()
    Set Application = CreateObject("Outlook.Application")
    Set Outlook = Application.GetNamespace("MAPI")
    Set folders = Outlook.GetDefaultFolder(FOLDER_INDEX)

    If config("outlook_sub_folder") = "" Then
        Set folder = folders
    Else
        Set folder = folders.Folders(config("outlook_sub_folder"))
    End If
End Function


'
' 追記対象のExcel返却
'
' @params Scripting.Dictionary config 設定値
' @return Excel.Application Excelファイル
'
Function excel(config)
    Set excelWorksheet = Nothing
    Set excel = CreateObject("Excel.Application")
    excel.Application.Workbooks.Open(config("excel_file_path"))
    excel.Application.DisplayAlerts = False
End Function

'
' 抽出対象の判定
'
' @params MailItem item メールファイル
' @params Scripting.Dictionary config 設定値
' @return boolean 抽出対象の場合true、それ以外の場合はfalse
'
Function condition(item, config)
    condition = False
    IF InStr(item.Subject, config("pickup_title")) > 0 Then
        condition = True
    END If
End Function

'
' 抽出結果のコンテンツ返却
'
' @params string body メール本文
' @params Scripting.Dictionary config 設定値
' @return Scripting.Dictionary 抽出結果のハッシュ値
'
Function contents(body, config)
    Set contents = CreateObject("Scripting.Dictionary")

    Set reg = New RegExp
    reg.Pattern = config("body_grep_rule")
    reg.Global = True
    reg.MultiLine = True

    Set matches = reg.Execute(body)

    For i = 0 To matches.Count - 1
        contents.Add Chr(Asc(config("excel_start_row")) + i), matches.Item(i).SubMatches.Item(0)
    Next
End Function

'
' Excel書込
'
Function write(worksheet, line, contents, config)
    write = False
    For Each key In contents.Keys
        cell = key + CStr(line)
        worksheet.Range(cell).Value = contents(key)
    Next
    write = True
End Function

'
' 追記対象行を返却
'
' @params Scripting.Dictionary config 設定値
' @return Integer Outlookの抽出対象フォルダ
'
Function startLine(worksheet, config)
    IF config("excel_start_line") > 0 Then
        startLine = config("excel_start_line")
        Exit Function
    END If

    startLine = 1

    Do
        IF worksheet.Range(config("excel_start_row") + CStr(startLine)).Value = "" Then
            Exit Function
        END If
        startLine = startLine + 1
    Loop
End Function

run(config)

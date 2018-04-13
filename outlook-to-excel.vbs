' ��M�g���C �t�H���_�̒萔��`
' https://msdn.microsoft.com/ja-jp/vba/outlook-vba/articles/oldefaultfolders-enumeration-outlook
Const FOLDER_INDEX = 6

Set config  = CreateObject("Scripting.Dictionary")

' ��M�g���C�̉��w�t�H���_�B��M�g���C�𒊏o�ΏۂƂ������ꍇ�͋�ɂ���
config.Add "outlook_sub_folder", ""
' �ǋL����Excel�t�@�C���܂ł̃p�X
config.Add "excel_file_path", "D:\sample.xlsx"
' �ǋL����Excel�t�@�C���̃��[�N�V�[�g��
config.Add "excel_worksheet_name", "Sheet1"
' ���o�Ώۂ̌����B�ݒ肵�������񂪊܂܂�郁�[����ΏۂƂ���
config.Add "pickup_title", "����"
' ���o�Ώی���
config.Add "max_pickup_num", "100"
' �{�����o����
config.Add "body_grep_rule", "��(.+)$"
' Excel�ɒǋL����ŏ��̗�l(A~Z�܂�)
config.Add "excel_start_row", "A"
' Excel�ɒǋL����ŏ��̍s�B�l��0�̏ꍇ�͍ŏ��Ƀq�b�g������Z���̍s���N�_�Ƃ���
config.Add "excel_start_line", 0

'
' ���C������
'
' @params Scripting.Dictionary config �ݒ�l
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
            message = message + "�E" + item.Subject + vbCrLf
        End If
    Next

    On Error Resume Next
    excelFile.Save
    If Err.Number Then
        MsgBox "Excel�̕ۑ��Ɏ��s���܂���"
    Else
        IF message = "" Then
            MsgBox "�o�^�Ώۂ�1�������݂��܂���ł���"
        Else
            MsgBox "�ȉ����[���̓��e��Excel�ɒǋL���܂���" + vbCrLf + message
        End If
    End If
    On Error Goto 0

    excelFile.Quit
End Function

'
' ���o�Ώۂ̃t�H���_�ԋp
' @return Outlook.Application Outlook�̒��o�Ώۃt�H���_
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
' �ǋL�Ώۂ�Excel�ԋp
'
' @params Scripting.Dictionary config �ݒ�l
' @return Excel.Application Excel�t�@�C��
'
Function excel(config)
    Set excelWorksheet = Nothing
    Set excel = CreateObject("Excel.Application")
    excel.Application.Workbooks.Open(config("excel_file_path"))
    excel.Application.DisplayAlerts = False
End Function

'
' ���o�Ώۂ̔���
'
' @params MailItem item ���[���t�@�C��
' @params Scripting.Dictionary config �ݒ�l
' @return boolean ���o�Ώۂ̏ꍇtrue�A����ȊO�̏ꍇ��false
'
Function condition(item, config)
    condition = False
    IF InStr(item.Subject, config("pickup_title")) > 0 Then
        condition = True
    END If
End Function

'
' ���o���ʂ̃R���e���c�ԋp
'
' @params string body ���[���{��
' @params Scripting.Dictionary config �ݒ�l
' @return Scripting.Dictionary ���o���ʂ̃n�b�V���l
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
' Excel����
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
' �ǋL�Ώۍs��ԋp
'
' @params Scripting.Dictionary config �ݒ�l
' @return Integer Outlook�̒��o�Ώۃt�H���_
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

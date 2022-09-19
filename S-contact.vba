'プログラム0｜変数設定の指定
Option Explicit

'プログラム1｜プログラム開始
Sub SendMail1()

    'プログラム2｜シート設定
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Set ws = Worksheets("Sheet1")
    Set ws2 = Worksheets("Sheet2")
         
    'プログラム3｜Outlookアプリケーションを起動
    Dim outlookObj As Outlook.Application
    Set outlookObj = CreateObject("Outlook.Application")
    
    'プログラム4｜Outlookメールを作成
    Dim mymail As Outlook.MailItem
    Set mymail = outlookObj.CreateItem(olMailItem)
    
    If Selection.Rows.Count > 1 Then
        GoTo ErrMulti
    End If
    
    'プログラム5｜メール情報を設定
    mymail.BodyFormat = 3        'リッチテキストに変更
    mymail.CC = ""   'cc宛先
    mymail.BCC = ""  'bcc宛先
    mymail.Subject = "S連絡_" & Format(Date, "m.d")    '件名
    
    'プログラム6｜メール本文を設定
    Dim mailbody As String
    Dim strSch As String
    Dim strEx As String
    Dim strSite As String
    Dim i As Integer
    Dim n() As Integer
    ReDim n(Selection.Count)
    For i = 0 To Selection.Count - 1
        If i = 0 Then
            n(i) = 1
        Else
            n(i) = 1 + n(i - 1)
        End If
        Do Until Selection(n(i)).Value <> Selection(n(i) + 1).Value Or n(i) > Selection.Count - 1
            n(i) = n(i) + 1
        Loop
        If InStr(ws.Cells(Selection.Row, Selection(n(i)).Column).Value, "#") Then
            strEx = "【暫定】"
            strSite = Mid(Selection(n(i)).Value, 2)
        Else
            strSite = Selection(n(i)).Value
        End If
        If i = 0 Then
            If ws.Cells(Selection.Row, Selection(1).Column).Value = ws.Cells(Selection.Row, Selection(1).Column - 1).Value Then
                strSch = "・" & strEx & "～" & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "：" & strSite & vbCrLf
            Else
                If ws.Cells(1, Selection(1).Column) = ws.Cells(1, Selection(n(i)).Column) Then
                    strSch = "・" & strEx & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "：" & strSite & vbCrLf
                Else
                    strSch = "・" & strEx & Format(ws.Cells(1, Selection(1).Column), "m/d(aaa)") & "～" & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "：" & strSite & vbCrLf
                End If
            End If
        Else
            If n(i) < Selection.Count + 1 Then
                If ws.Cells(Selection.Row, Selection(n(i)).Column).Value = ws.Cells(Selection.Row, Selection(n(i)).Column + 1).Value And _
                    ws.Cells(Selection.Row, Selection(n(i)).Column + 1).Value = ws.Cells(Selection.Row, Selection(n(i)).Column + 2).Value Then
                    strSch = strSch & "・" & strEx & Format(ws.Cells(1, Selection(1 + n(i - 1)).Column), "m/d(aaa)") & "～：" & strSite & vbCrLf
                ElseIf ws.Cells(Selection.Row, Selection(n(i)).Column).Value = ws.Cells(Selection.Row, Selection(n(i)).Column + 1).Value Then
                    If ws.Cells(1, Selection(1 + n(i - 1)).Column) = ws.Cells(1, Selection(n(i)).Column) Then
                        strSch = strSch & "・" & strEx & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "?：" & strSite & vbCrLf
                    Else
                        strSch = strSch & "・" & strEx & Format(ws.Cells(1, Selection(1 + n(i - 1)).Column), "m/d(aaa)") & "～" & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "?：" & strSite & vbCrLf
                    End If
                Else
                    If ws.Cells(1, Selection(1 + n(i - 1)).Column) = ws.Cells(1, Selection(n(i)).Column) Then
                        strSch = strSch & "・" & strEx & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "：" & strSite & vbCrLf
                    Else
                        strSch = strSch & "・" & strEx & Format(ws.Cells(1, Selection(1 + n(i - 1)).Column), "m/d(aaa)") & "～" & Format(ws.Cells(1, Selection(n(i)).Column), "m/d(aaa)") & "：" & strSite & vbCrLf
                    End If
                End If
            End If
        End If
        strEx = ""
        strSite = ""
    Next
    
    mailbody = strSch
    mymail.Body = mailbody & vbCrLf
    
    'プログラム7｜メールにファイルを添付
    Dim attachedfile As String
    attachedfile = ThisWorkbook.Path & "\" & ws.Range("A1").Value
    If Not attachedfile = Null Then
        mymail.Attachments.Add Source:=attachedfile
    End If
    
    On Error GoTo ErrHandl
    mymail.To = WorksheetFunction.VLookup(ws.Cells(Selection.Row, 1).Value, ws2.Range("A2:B7"), 2, False)   'To宛先
            
    'プログラム8｜メール表示
    mymail.Display     'メール表示(ここでは誤送信を防ぐために表示だけにして、メール送信はしない)
    
    'プログラム9｜メール下書き保存
    'mymail.Save
    
    'プログラム10｜メール送信
    'mymail.Send

    'プログラム11｜オブジェクト解放
    Set outlookObj = Nothing
    Set mymail = Nothing
    
    Exit Sub

ErrHandl:
    MsgBox ("対象の社員が見つかりません。")
    Exit Sub

ErrMulti:
    MsgBox ("複数行の選択は無効です。1行だけ選択してください。")
    
'プログラム12｜プログラム終了vvv
End Sub
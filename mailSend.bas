Attribute VB_Name = "mailSend"
Option Explicit
    Public addressLine As String 'bccに挿入する全部のemail Address
    Public sendSubject As String
    Public makebody As String
    Const sheetNameOne As String = "mailSendOne"
    Const sheetNameAll As String = "mailSendAll"
    
Private Sub sendAllOne()
    'メールアドレス全部をBCCに含むメールを１通、OutLook下書きに作成する
    Call addressInput
    Call createMailBcc
    MsgBox "Outlook下書きにメールを作成しました"
End Sub

Private Sub sendAllSeveral()
    'メールアドレスごとに各１通、Outlook下書きに作成する
    Dim i As Long
    Dim sendAllS As Worksheet
    Const startCellRow As Long = 9
    Const endCellRow As Long = 11
    
    Set sendAllS = ThisWorkbook.Sheets(sheetNameAll)
    For i = startCellRow To endCellRow
        Call addressUniInput(sendAllS, i)
        Call createMailSendto
    Next i
    Set sendAllS = Nothing
    MsgBox "Outlook下書きにメールを作成しました"
End Sub

Private Sub addressInput()
    '一括BCCメール作成のためのデータ読み出しルーチン
    Dim sendOneS As Worksheet
    Const startCellRow As Long = 9
    Const endCellRow As Long = 11
    Const addressCol = 5
    Dim i As Long
    addressLine = ""
    
    Set sendOneS = ThisWorkbook.Sheets(sheetNameOne) 'mailSendOneシート定義
    With sendOneS
        For i = startCellRow To endCellRow
            If i = startCellRow Then
                addressLine = .Cells(i, addressCol).Value
            Else
                addressLine = addressLine + "; " + .Cells(i, addressCol).Value
            End If
        Next i
        sendSubject = .Cells(6, 3).Value
        makebody = .Cells(7, 3).Value
    End With
    Set sendOneS = Nothing
End Sub

Private Sub createMailBcc()
    'BCC下書きメール作成ルーチン
    Dim objOutlook As New Outlook.Application
    Dim objMailitem As Outlook.MailItem
    Set objMailitem = objOutlook.CreateItem(olMailItem)
   
   'mail設定
    With objMailitem
        .To = ""            '宛先
        .CC = ""            'CC
        .BCC = addressLine  'BCC
        .Subject = sendSubject  'タイトル
        .Body = makebody    '本文
        objMailitem.Save    '下書き保存
        .Display            '新規メール画面を表示
    End With
    Set objMailitem = Nothing
End Sub

Private Sub addressUniInput(sendAllS As Worksheet, i As Long)
    '個別メールデータセットルーチン
    Const addressCol = 5
    addressLine = ""
    sendSubject = ""
    makebody = ""
    With sendAllS
        addressLine = .Cells(i, addressCol).Value
        sendSubject = .Cells(6, 3).Value
        makebody = .Cells(7, 3).Value
        makebody = Replace(makebody, "〇〇", .Cells(i, 3).Value)    'kaisha
        makebody = Replace(makebody, "××", .Cells(i, 4).Value)    'tanto
    End With
End Sub

Private Sub createMailSendto()
    'CC下書きメール作成ルーチン
    Dim objOutlook As New Outlook.Application
    Dim objMailitem As Outlook.MailItem
    Set objMailitem = objOutlook.CreateItem(olMailItem)
   
   'mail設定
    With objMailitem
        .To = ""            '宛先
        .CC = addressLine   'CC
        .BCC = ""           'BCC
        .Subject = sendSubject  'タイトル
        .Body = makebody    '本文
        objMailitem.Save    '下書き保存
        .Display            '新規メール画面を表示
    End With
    Set objMailitem = Nothing
End Sub

Sub ボタン1_Click()
    Call sendAllOne
End Sub

Sub ボタン2_Click()
    Call sendAllSeveral
End Sub


Attribute VB_Name = "mailSend"
Option Explicit
    Public addressLine As String 'bcc�ɑ}������S����email Address
    Public sendSubject As String
    Public makebody As String
    Const sheetNameOne As String = "mailSendOne"
    Const sheetNameAll As String = "mailSendAll"
    
Private Sub sendAllOne()
    '���[���A�h���X�S����BCC�Ɋ܂ރ��[�����P�ʁAOutLook�������ɍ쐬����
    Call addressInput
    Call createMailBcc
    MsgBox "Outlook�������Ƀ��[�����쐬���܂���"
End Sub

Private Sub sendAllSeveral()
    '���[���A�h���X���ƂɊe�P�ʁAOutlook�������ɍ쐬����
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
    MsgBox "Outlook�������Ƀ��[�����쐬���܂���"
End Sub

Private Sub addressInput()
    '�ꊇBCC���[���쐬�̂��߂̃f�[�^�ǂݏo�����[�`��
    Dim sendOneS As Worksheet
    Const startCellRow As Long = 9
    Const endCellRow As Long = 11
    Const addressCol = 5
    Dim i As Long
    addressLine = ""
    
    Set sendOneS = ThisWorkbook.Sheets(sheetNameOne) 'mailSendOne�V�[�g��`
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
    'BCC���������[���쐬���[�`��
    Dim objOutlook As New Outlook.Application
    Dim objMailitem As Outlook.MailItem
    Set objMailitem = objOutlook.CreateItem(olMailItem)
   
   'mail�ݒ�
    With objMailitem
        .To = ""            '����
        .CC = ""            'CC
        .BCC = addressLine  'BCC
        .Subject = sendSubject  '�^�C�g��
        .Body = makebody    '�{��
        objMailitem.Save    '�������ۑ�
        .Display            '�V�K���[����ʂ�\��
    End With
    Set objMailitem = Nothing
End Sub

Private Sub addressUniInput(sendAllS As Worksheet, i As Long)
    '�ʃ��[���f�[�^�Z�b�g���[�`��
    Const addressCol = 5
    addressLine = ""
    sendSubject = ""
    makebody = ""
    With sendAllS
        addressLine = .Cells(i, addressCol).Value
        sendSubject = .Cells(6, 3).Value
        makebody = .Cells(7, 3).Value
        makebody = Replace(makebody, "�Z�Z", .Cells(i, 3).Value)    'kaisha
        makebody = Replace(makebody, "�~�~", .Cells(i, 4).Value)    'tanto
    End With
End Sub

Private Sub createMailSendto()
    'CC���������[���쐬���[�`��
    Dim objOutlook As New Outlook.Application
    Dim objMailitem As Outlook.MailItem
    Set objMailitem = objOutlook.CreateItem(olMailItem)
   
   'mail�ݒ�
    With objMailitem
        .To = ""            '����
        .CC = addressLine   'CC
        .BCC = ""           'BCC
        .Subject = sendSubject  '�^�C�g��
        .Body = makebody    '�{��
        objMailitem.Save    '�������ۑ�
        .Display            '�V�K���[����ʂ�\��
    End With
    Set objMailitem = Nothing
End Sub

Sub �{�^��1_Click()
    Call sendAllOne
End Sub

Sub �{�^��2_Click()
    Call sendAllSeveral
End Sub


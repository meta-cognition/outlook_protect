Attribute VB_Name = "Module1"
Sub SendAllDraftEmails()
Dim xAccount As Account
Dim xDraftFld As Folder
Dim xItemCount As Integer
Dim xDraftsItems As Outlook.Items
Dim xPromptStr As String
Dim xYesOrNo As Integer
Dim i, k As Long
Dim xNewMail As MailItem
Dim xTmpPath, xFilePath As String
On Error Resume Next
xItemCount = 0
For Each xAccount In Outlook.Application.Session.Accounts
    Set xDraftFld = xAccount.DeliveryStore.GetDefaultFolder(olFolderDrafts)
    xItemCount = xItemCount + xDraftFld.Items.Count
Next xAccount
Set xDraftFld = Nothing
If xItemCount > 0 Then
    xPromptStr = "Are you sure to send out all the drafts?"
    xYesOrNo = MsgBox(xPromptStr, vbQuestion + vbYesNo, "Kutools for Outlook")
    If xYesOrNo = vbYes Then
        For Each xAccount In Outlook.Application.Session.Accounts
            Set xDraftFld = xAccount.DeliveryStore.GetDefaultFolder(olFolderDrafts)
            Set xDraftsItems = xDraftFld.Items
            For i = xDraftsItems.Count To 1 Step -1
                If i = 1 Then
                    Set xNewMail = Outlook.Application.CreateItem(olMailItem)
                    With xNewMail
                        .SendUsingAccount = xDraftsItems.Item(i).SendUsingAccount
                        .To = xDraftsItems.Item(i).To
                        .CC = xDraftsItems.Item(i).CC
                        .BCC = xDraftsItems.Item(i).BCC
                        .Subject = xDraftsItems.Item(i).Subject
                        If xDraftsItems.Item(i).Attachments.Count > 0 Then
                            xTmpPath = "C:\MyTempAttachments"
                            If Dir(xTmpPath, vbDirectory) = "" Then
                                MkDir xTmpPath
                            End If
                            For k = xDraftsItems.Item(i).Attachments.Count To 1 Step -1
                                xFilePath = xTmpPath & "\" & xDraftsItems.Item(i).Attachments.Item(k).FileName
                                xDraftsItems.Item(i).Attachments.Item(k).SaveAsFile xFilePath
                                xNewMail.Attachments.Add (xFilePath)
                                Kill xFilePath
                            Next k
                            RmDir xTmpPath
                        End If
                        .HTMLBody = xDraftsItems.Item(i).HTMLBody
                        .Send
                    End With
                    xDraftsItems.Item(i).Delete
                Else
                    xDraftsItems.Item(i).Send
                End If
            Next
        Next xAccount
        MsgBox "Successfully sent " & xItemCount & " messages", vbInformation, "Kutools for Outlook"
    End If
Else
    MsgBox "No Drafts!", vbInformation + vbOKOnly, "Kutools for Outlook"
End If
End Sub

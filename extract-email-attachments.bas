Sub ExtractEmailAttachments()
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.Folder
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim objAttachment As Outlook.Attachment
    Dim saveFolder As String
    Dim fileExt As String

    ' Set the save folder path where the attachments will be saved
    saveFolder = "C:\Attachments"

    ' Set the file extension you want to extract (e.g., ".pdf")
    fileExt = ".pdf"

    On Error Resume Next
    ' Get the selected folder in Outlook
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objFolder = Application.ActiveExplorer.CurrentFolder

    ' Loop through each selected item in the folder
    For Each objItem In objFolder.Items
        ' Check if the item is a mail item
        If TypeOf objItem Is Outlook.MailItem Then
            Set objMail = objItem

            ' Loop through each attachment in the mail item
            For Each objAttachment In objMail.Attachments
                ' Check if the attachment has the specified file extension
                If Right(objAttachment.FileName, Len(fileExt)) = fileExt Then
                    ' Save the attachment to the specified folder
                    objAttachment.SaveAsFile saveFolder & "\" & objAttachment.FileName
                End If
            Next objAttachment
        End If
    Next objItem

    ' Clean up objects
    Set objNamespace = Nothing
    Set objFolder = Nothing
    Set objMail = Nothing
    Set objAttachment = Nothing

    MsgBox "Attachments extracted successfully!"
End Sub

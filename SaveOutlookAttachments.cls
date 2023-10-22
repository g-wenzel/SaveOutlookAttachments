Public Sub ReplaceAttachmentsToLink()
    Dim objApp As Outlook.Application
    Dim aMail As Outlook.MailItem
    Dim oAttachments As Outlook.Attachments
    Dim oSelection As Outlook.Selection
    Dim i As Long
    Dim iAttachmentCount As Long
    Dim sFile As String
    Dim sDateAndSubject As String
    Dim sBaseFolderPath As String
    Dim sSubFolderPath As String
    Dim sDeletedFiles As String
    Dim regEx As Object
    Dim sIgnoreList(2) As String
    
    ' Set path to local Attachment Folder here
    sBaseFolderPath = "D:\"
    
    If Dir(sBaseFolderPath) = "" Then
        MsgBox "Error: Drive " & sBaseFolderPath & " not found."
        Exit Sub
    End If
    
    sBaseFolderPath = sBaseFolderPath & "OutlookAttachments"

    ' Add filenames to ignore here. If new filenames are added, array size above mus be increased
    sIgnoreList(0) = "ATT00001.txt"
    sIgnoreList(1) = "smime.p7s"

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True

    ' Illegal characters for folder names
    regEx.Pattern = "[<>:|/\\?*""]"


    On Error Resume Next

    ' Instantiate an Outlook Application object.
    Set objApp = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set oSelection = objApp.ActiveExplorer.Selection

    ' Set the Attachment folder.
    If Dir(sBaseFolderPath) = "" Then
        MkDir sBaseFolderPath
    End If

    ' Check each selected item for attachments. If attachments exist,
    ' save them to specified folder and strip them from the item.
    For Each aMail In oSelection

        ' Get the Attachments collection of the item.
        Set oAttachments = aMail.Attachments
        iAttachmentCount = oAttachments.Count

            
        If iAttachmentCount > 0 Then

            ' We need to use a count down loop for removing items
            ' from a collection. Otherwise, the loop counter gets
            ' confused and only every other item is removed.
            
            For i = iAttachmentCount To 1 Step -1
            
                ' Save attachment before deleting from item.
                ' Get the file name.
                sFile = oAttachments.Item(i).FileName
                If Not (IsInArray(sFile, sIgnoreList)) Then
                
                    ' Make folder name out of date and subject
                    sDateAndSubject = Format(aMail.ReceivedTime, "YYYY-MM-DD")
                    sDateAndSubject = sDateAndSubject & "_" & aMail.Subject
                    
                    'trim string to 100 characters max
                    sDateAndSubject = Left(sDateAndSubject, 100)
                    sDateAndSubject = Trim(sDateAndSubject)
                    
                    'replace illegal characters in a file path
                    sDateAndSubject = regEx.Replace(sDateAndSubject, "_")
                    
                    ' make folder named with date and subject
                    sSubFolderPath = sBaseFolderPath & "\" & sDateAndSubject
                    If Dir(sSubFolderPath) = "" Then
                            MkDir sSubFolderPath
                    End If
                    
                    ' Combine filenamne with path.
                    sFile = sSubFolderPath & "\" & sFile
                        
                    ' Save the attachment as a file.
                    oAttachments.Item(i).SaveAsFile sFile
                        
                    ' Delete the attachment.
                    oAttachments.Item(i).Delete
                        
                    'write the save as path to a string to add to the message
                    'check for html and use html tags in link
                    If aMail.BodyFormat <> olFormatHTML Then
                        sDeletedFiles = sDeletedFiles & vbCrLf & "<file://" & sFile & ">"
                    Else
                        sDeletedFiles = sDeletedFiles & "<br>" & "<a href='file://" & _
                        sFile & "'>" & sFile & "</a>"
                    End If
                        
                End If
            Next i
            
            If sDeletedFiles <> "" Then
                ' Adds the filename string to the message body and save it
                ' Check for HTML body
                If aMail.BodyFormat <> olFormatHTML Then
                    aMail.Body = aMail.Body & vbCrLf & _
                    "The file(s) were saved to " & sDeletedFiles
                Else
                    aMail.HTMLBody = aMail.HTMLBody & "<p>" & _
                    "The file(s) were saved to " & sDeletedFiles & "</p>"
                End If
                
                aMail.Save
                'sets the attachment path to nothing before it moves on to the next message.
                sDeletedFiles = ""
            End If
        End If
    Next
   
ExitSub:
 
Set oAttachments = Nothing
Set aMail = Nothing
Set oSelection = Nothing
Set objApp = Nothing
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

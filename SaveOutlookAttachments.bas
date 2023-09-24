Attribute VB_Name = "MainModule"
Sub SaveAttachmentsFromSelectedEmails()
    On Error GoTo ErrorHandler
    
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.mailItem
    Dim objAttachments As Outlook.Attachments
    Dim strBaseFolderPath As String
    Dim strFolderPath As String
    Dim strFolderName As String
    Dim strFileName As String
    Dim i As Long
    
    ' Define the base folder path
    strBaseFolderPath = "C:\EmailAttachments\"
    
    ' Create the base folder if it doesn't exist
    If Dir(strBaseFolderPath, vbDirectory) = "" Then
        CreateFolder strBaseFolderPath
    End If
    
    ' Get the selected items
    Set objSelection = Outlook.ActiveExplorer.Selection
    
    ' Loop through each selected item
    For Each objMail In objSelection
        If TypeOf objMail Is mailItem Then
            Set objAttachments = objMail.Attachments
            
            ' Continue if there are attachments
            If objAttachments.Count > 0 Then
                ' Initially set the strFolderName as the email subject
                strFolderName = objMail.Subject
                
                ' Create a RegExp object
                Dim regex As Object
                Set regex = CreateObject("VBScript.RegExp")
                regex.Global = True
                regex.Pattern = "[^a-zA-Z0-9 &._()#%-]" ' Pattern to match any character that is not alphanumeric or space
            
                ' Replace non-alphanumeric characters with underscore
                strFolderName = regex.Replace(strFolderName, "_")
                
                ' Debugging: Print each character and its ASCII value in the Immediate Window
                Dim charPos As Integer
                Debug.Print "Subject: " & strFolderName ' Print the whole subject first
                For charPos = 1 To Len(strFolderName)
                    Debug.Print "Position: " & charPos & " Char: " & Mid(strFolderName, charPos, 1) & " ASCII: " & AscW(Mid(strFolderName, charPos, 1))
                Next charPos
                
                ' Replace HTML entity &#8201; with regular space
                strFolderName = Replace(strFolderName, "&#8201;", " ")
                
                ' Replace ASCII 8201 spaces with regular spaces (ASCII 32)
                strFolderName = Replace(strFolderName, ChrW(8201), " ")
                
                ' Replace invalid characters
                strFolderName = ReplaceInvalidCharacters(strFolderName)
                
                ' Remove trailing spaces
                strFolderName = RemoveTrailingSpaces(strFolderName)
                
                ' Construct the folder path with the sanitized folder name
                strFolderPath = strBaseFolderPath & strFolderName
                
                ' Print the sanitized folder name to the Immediate Window for debugging
                Debug.Print "Sanitized Folder Name: " & strFolderName
                
                ' Create the specific folder for the email if it doesn't exist
                If Dir(strFolderPath, vbDirectory) = "" Then
                    CreateFolder strFolderPath
                End If
                
                ' Save each attachment to the folder
                For i = 1 To objAttachments.Count
                    On Error Resume Next ' Continue to the next attachment if an error occurs
                    strFileName = strFolderPath & "\" & CreateValidName(objAttachments.Item(i).FileName)
                    Debug.Print "Trying to save to: " & strFileName ' Print the path to the Immediate Window
                    objAttachments.Item(i).SaveAsFile strFileName
                    If Err.Number <> 0 Then
                        Debug.Print "Failed to save attachment to: " & strFileName & ". Error: " & Err.Description
                        MsgBox "Failed to save attachment to: " & strFileName & vbCrLf & "Error: " & Err.Description, vbCritical
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler ' Reset error handling to the general error handler after attempting to save each attachment
                Next i
            End If
        End If
    Next objMail
    
    ' Display a message when done
    MsgBox "Attachments have been saved.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Function CreateValidName(ByVal strName As String) As String
    Dim strValidName As String
    Dim ext As String
    Dim pos As Integer
    
    strValidName = strName
    
    ' Find the last period (file extension separator)
    pos = InStrRev(strValidName, ".")
    
    ' Extract the extension and the name separately if a period is found
    If pos > 0 Then
        ext = Mid(strValidName, pos)
        strValidName = Left(strValidName, pos - 1)
    Else
        ext = ""
    End If
    
    ' Remove trailing spaces from the name part
    strValidName = RemoveTrailingSpaces(strValidName)
    
    ' Replace invalid characters in the name part
    strValidName = ReplaceInvalidCharacters(strValidName)
    
    ' Combine the sanitized name with the extension
    CreateValidName = strValidName & ext
End Function

Function RemoveTrailingSpaces(ByVal str As String) As String
    Dim i As Integer
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) <> " " Then
            RemoveTrailingSpaces = Left(str, i)
            Exit Function
        End If
    Next i
    RemoveTrailingSpaces = ""
End Function

Function ReplaceInvalidCharacters(ByVal str As String) As String
    ' Replace invalid characters in the name part
    str = Replace(str, "\", "")
    str = Replace(str, ":", "")
    str = Replace(str, "*", "")
    str = Replace(str, "?", "")
    str = Replace(str, """", "")
    str = Replace(str, "<", "")
    str = Replace(str, ">", "")
    str = Replace(str, "|", "")
    ' Retaining #, ., (, and )
    ' str = Replace(str, "#", "")
    ' str = Replace(str, ".", "")
    ' str = Replace(str, "(", "")
    ' str = Replace(str, ")", "")
    
    ReplaceInvalidCharacters = str
End Function

Sub CreateFolder(ByVal strFolderPath As String)
    Dim strParentPath As String
    Dim strLastFolder As String
    Dim pos As Integer
    
    ' Check if folder already exists
    If Dir(strFolderPath, vbDirectory) <> "" Then Exit Sub
    
    ' Find the parent folder path and the last folder name in the path
    pos = InStrRev(strFolderPath, "\")
    strParentPath = Left(strFolderPath, pos - 1)
    strLastFolder = Mid(strFolderPath, pos + 1)
    
    ' Recursively create parent folders if they don't exist
    If Dir(strParentPath, vbDirectory) = "" Then
        CreateFolder strParentPath
    End If
    
    ' Create the folder
    MkDir strFolderPath
End Sub


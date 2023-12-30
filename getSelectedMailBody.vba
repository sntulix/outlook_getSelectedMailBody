' Get Bodies of Mail Selected in Outlook
Sub getMailText()
    Dim dataObj As MSForms.DataObject ' To copy Clipboard, append reference of FM20.dll
    Set dataObj = New MSForms.DataObject
    
    Dim objOL As Outlook.Application
    Dim mySelectionItems As Outlook.Selection
    Dim myItem As Outlook.MailItem
    
    Set objOL = New Outlook.Application
    Set mySelectionItems = objOL.ActiveExplorer.Selection
    
    Dim str As String

    For i = 1 To mySelectionItems.Count
        Set myItem = mySelectionItems.Item(i)
        ' date_string = Format(myItem.ReceivedTime, "yyyymmddhnnss")
        ' MsgBox myItem.Subject, vbOKOnly, date_string
        str = str & """""""" & ">" & myItem.ReceivedTime & vbCrLf & myItem.Body & vbCrLf & """""""" & "<" & vbCrLf
    Next i

    ' copy to clipboard
    dataObj.SetText str
    dataObj.PutInClipboard

End Sub

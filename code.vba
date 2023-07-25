Option Explicit

Private Declare Function ShellExecute _
Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hWnd As Long, _
ByVal Operation As String, _
ByVal fileName As String, _
Optional ByVal Parameters As String, _
Optional ByVal Directory As String, _
Optional ByVal WindowStyle As Long = vbMinimizedFocus _
) As Long
Private WithEvents Items As Outlook.Items


' This function will search the active message / draft for the search string (email address)
Public Function GetSearchString() As String
    On Error GoTo On_Error
    Dim GetCurrentItem As Object
    Dim objApp As Outlook.Application
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
    End Select
    GetSearchString = "stop"
    
    Dim fromEmail As String
    fromEmail = GetSmtpAddress(GetCurrentItem)
    Dim check As Integer
    check = checkEmail(fromEmail)
    If check = 1 Then 'Found an email address that is likely a search string
        GetSearchString = fromEmail
        GoTo Exiting
    End If
    
    'Let's look at the "Recipients" email address, and see if we can find a likely search string there.
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    Dim temp2 As String
    Set recips = GetCurrentItem.Recipients

    Dim PR_SMTP_ADDRESS As String
    PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    
    For Each recip In recips
        recip.Resolve
        Set pa = recip.PropertyAccessor
        temp2 = pa.GetProperty(PR_SMTP_ADDRESS)
        check = checkEmail(temp2)
        If check = 1 Then 'Found an email address that is likely a search string
            GetSearchString = temp2
            GoTo Exiting
        End If
    Next
        
    MsgBox "Could not find email address to search. " & GetSearchString
    Set objApp = Nothing

Exiting:
'MsgBox (GetSearchString)
GetSearchString = UCase(GetSearchString)
Set objApp = Nothing
Exit Function
On_Error:
MsgBox "error=" & Err.Number & " " & Err.Description
Resume Exiting
End Function


' This function will take a potential email address to search and verify it's format. 
' For this example, email addresses will have the format ##ssssss@xxx.com
Public Function checkEmail(fromEmail As String)
    Dim UserDomain() As String
    Dim Domain As String
    Dim Username As String
    
    If InStr(fromEmail, "@") > 0 Then 'Found Email address. Checking if matches format:
        UserDomain = Split(fromEmail, "@", 2)
        Username = UserDomain(0) 'Before the @ symbol
        Domain = UserDomain(1) 'After the @ symbol
        If Domain = "xxx.com" And Val(Left(Username, 2)) > 9 Then
            'Found an email address that is likely the correct format
            checkEmail = 1
        Else
            checkEmail = 0
        End If
    End If
End Function


Public Function getInfo(searchString, Col)
    
    searchString = LCase(searchString)
    
    Dim wb As Workbook
    Dim dataArea As Excel.Range
    Dim valuesArray() As Variant
    Dim rowIndex As Long
    Dim fileName As String
    fileName = "C:\Users\####\VBA Script\Sample.xls"
    
    Set wb = Workbooks.Open(fileName)
    Set dataArea = wb.Worksheets(1).Range("A1:E500")  ' May need to be updated if your data has more rows/columns
    valuesArray = dataArea.Value
    For rowIndex = LBound(valuesArray, 1) To UBound(valuesArray, 1)
        If valuesArray(rowIndex, 3) = searchString Then
            getInfo = valuesArray(rowIndex, Col)
            Exit Function
        End If
    Next
    wb.Close
    
End Function


Public Function GetSmtpAddress(mail As MailItem)
    Dim Report As String
    Dim Session As Outlook.NameSpace
    Set Session = Application.Session
    
    Dim PR_SMTP_ADDRESS As String
    PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
     
    If mail.SenderEmailType <> "EX" Then
        GetSmtpAddress = mail.SenderEmailAddress
    Else
        Dim senderEntryID As String
        Dim sender As AddressEntry
        Dim PR_SENT_REPRESENTING_ENTRYID As String
        PR_SENT_REPRESENTING_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x00410102"
     
        senderEntryID = mail.PropertyAccessor.BinaryToString( _
        mail.PropertyAccessor.GetProperty( _
        PR_SENT_REPRESENTING_ENTRYID))
        Set sender = Session.GetAddressEntryFromID(senderEntryID)
        If sender Is Nothing Then
            Exit Function
        End If
     
        If sender.AddressEntryUserType = olExchangeUserAddressEntry Or _
            sender.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Dim exchangeUser As exchangeUser
            Set exchangeUser = sender.GetExchangeUser()
            If exchangeUser Is Nothing Then
                Exit Function
            End If
            GetSmtpAddress = exchangeUser.PrimarySmtpAddress
        Else
            GetSmtpAddress = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        End If
    End If
End Function


Sub openSite()
    Dim searchString As String
    Dim URL As String
    searchString = GetSearchString()
    If searchString = "stop" Then Exit Sub
    URL = getInfo(searchString, 2)  ' Selects 2nd column from source .xls file
    If Len(URL) = 0 Then
        MsgBox "Could not find " & GetSearchString & " in the data file."
        Exit Sub
    End If
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", URL)
End Sub




Sub CCothers()
    Dim objApp As Outlook.Application
    Dim GetCurrentItem As Object
    Set objApp = Application
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    If (GetCurrentItem.Class = 43) And (Not GetCurrentItem.Sent) Then ' Check to make sure the message is a draft
        Dim searchString As String
        searchString = GetSearchString() ' Get Search String
        If searchString = "stop" Then Exit Sub
        Dim sResult As String
        sResult = getInfo(searchString, 3)
        If Len(sResult) = 0 Then
            MsgBox "Could not find " & GetSearchString & " in the data file"
            Exit Sub
        End If
        
        Dim recips() As String ' if more than one email address is returned, split them into an array
        recips = Split(sResult, ";")
        Dim recip As Variant
        
        For Each recip In recips
            Set recip = GetCurrentItem.Recipients.Add(recip) ' Add recips(0)
            recip.Type = Outlook.OlMailRecipientType.olCC
            recip.Resolve
        Next
    End If
End Sub




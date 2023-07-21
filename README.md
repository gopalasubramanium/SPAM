# SPAM / SPOOF Control

This VBA script for Microsoft Outlook automatically filters incoming emails, checking for any connection to a known list of spam or spoof email domains. If a match is found in the email's header information, the email is automatically moved to a separate folder named "Spoofing Emails".

The list of spam domains is fetched from two online sources during the startup of the application. The script handles potential duplicates from the two sources, ensuring only unique domains are stored.

This script aids in keeping your inbox clean and safe by filtering out potential spam or spoof emails, reducing the risk of phishing attacks.

Please note that this script only works with Microsoft Outlook and requires VBA (Visual Basic for Applications) to be enabled. Furthermore, this is a heuristic approach and might not catch all spam emails, nor does it replace a full-featured spam filter provided by security software. Always exercise caution when handling emails from unknown senders.

VBA Script to check SPAM &amp; SPOOF incoming Emails in Microsoft Outlook:

1. Open Microsoft Outlook.
2. Press Alt + F11 on your keyboard. This will open the Visual Basic for Applications (VBA) Editor.
3. Create a New Module: In the VBA Editor, click on "Insert" in the top menu >> Choose "Module." This will insert a new module into the project.
4. Delete whatever is autopolulated and insert the below code.
5. Save the VBA Project: Click on the "Save" button in the VBA Editor or press Ctrl + S.
6. Close the VBA Editor by clicking the "X" button or pressing Alt + Q.

```
Option Explicit

Private WithEvents inboxItems As Outlook.Items
Private domains As Collection

Private Sub Application_Startup()
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Dim inbox As Outlook.Folder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    Set inboxItems = inbox.Items
    
    ' Load spam domains from the web
    Set domains = New Collection
    
    Dim domain As Variant
    Dim domainsFromSource1 As Collection
    Dim domainsFromSource2 As Collection
    Set domainsFromSource1 = GetDomainsFromWeb("https://raw.githubusercontent.com/unkn0w/disposable-email-domain-list/main/domains.txt")
    Set domainsFromSource2 = GetDomainsFromWeb("https://raw.githubusercontent.com/tsirolnik/spam-domains-list/master/spamdomains.txt")
    
    For Each domain In domainsFromSource1
        domains.Add domain, domain
    Next domain
    For Each domain In domainsFromSource2
        On Error Resume Next
        domains.Add domain, domain
        On Error GoTo 0
    Next domain
End Sub

Private Sub inboxItems_ItemAdd(ByVal item As Object)
    If TypeOf item Is MailItem Then
        FilterSpamEmails item
    End If
End Sub

Sub FilterSpamEmails(item As MailItem)

    Dim spamFolder As Outlook.MAPIFolder
    Dim mail As Outlook.MailItem
    Dim headerLines As Variant
    Dim line As Variant
    Dim domain As Variant
    
    ' Create a new folder under Inbox named "Spoofing Emails"
    ' If the folder already exists, the existing one will be used
    On Error Resume Next
    Set spamFolder = Application.Session.GetDefaultFolder(olFolderInbox).Folders("Spoofing Emails")
    If spamFolder Is Nothing Then
        Set spamFolder = Application.Session.GetDefaultFolder(olFolderInbox).Folders.Add("Spoofing Emails")
    End If
    On Error GoTo 0
    
    ' Check the headers of the incoming mail
    Set mail = item
    headerLines = Split(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"), vbCrLf)
    
    ' Loop through each line in the headers
    For Each line In headerLines
        ' Check if the line contains a known spam domain
        For Each domain In domains
            If InStr(line, domain) > 0 Then
                mail.Move spamFolder
                Exit Sub
            End If
        Next domain
    Next line
    
    ' Clean up
    Set spamFolder = Nothing
    Set mail = Nothing

End Sub

Private Function GetDomainsFromWeb(url As String) As Collection
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    xmlhttp.Open "GET", url, False
    xmlhttp.send
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    stream.Open
    stream.Type = 1 ' adTypeBinary
    stream.Write xmlhttp.responseBody
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    
    Dim domains As Collection
    Set domains = New Collection
    
    Dim lines As Variant
    lines = Split(stream.ReadText, vbCrLf)
    
    Dim line As Variant
    For Each line In lines
        domains.Add Trim(line)
    Next line
    
    stream.Close
    
    Set GetDomainsFromWeb = domains
End Function
```

@gopalasubramanium

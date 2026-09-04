Option Explicit

' Outlook corporate HTML signature generator
' Works with local Active Directory and Classic Outlook for Windows.

Const SIGNATURE_SUFFIX = "-Corporate"
Const COMPANY_WEBSITE = "https://www.example.com"
Const COMPANY_ADDRESS = "Company address"
Const COMPANY_LOGO_URL = "https://www.example.com/assets/company-logo.png"
Const ACCENT_COLOR = "#1F4E78"

Dim shell, fso, sysInfo, userObj
Dim appData, signatureFolder, signatureName, signatureFile
Dim fullName, jobTitle, companyName, emailAddress
Dim officePhone, mobilePhone, html, stream
Dim wordApp, signatureOptions

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
Set sysInfo = CreateObject("ADSystemInfo")
Set userObj = GetObject("LDAP://" & sysInfo.UserName)

If Err.Number <> 0 Or userObj Is Nothing Then
    WScript.Echo "Error: Active Directory user data could not be loaded. " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

fullName = HtmlEncode(ReadAdValue(userObj, "displayName"))
jobTitle = HtmlEncode(ReadAdValue(userObj, "title"))
companyName = HtmlEncode(ReadAdValue(userObj, "company"))
emailAddress = HtmlEncode(ReadAdValue(userObj, "mail"))
officePhone = HtmlEncode(ReadAdValue(userObj, "telephoneNumber"))
mobilePhone = HtmlEncode(ReadAdValue(userObj, "mobile"))

signatureName = SafeFileName(ReadAdValue(userObj, "sAMAccountName") & SIGNATURE_SUFFIX)
appData = shell.ExpandEnvironmentStrings("%APPDATA%")
signatureFolder = appData & "\Microsoft\Signatures"
signatureFile = signatureFolder & "\" & signatureName & ".htm"

If Not fso.FolderExists(signatureFolder) Then
    fso.CreateFolder signatureFolder
End If

html = "<!DOCTYPE html>" & vbCrLf
html = html & "<html><head><meta charset='utf-8'><title>Email Signature</title></head>" & vbCrLf
html = html & "<body style='margin:0;padding:0;font-family:Arial,sans-serif;font-size:10pt;color:#333333;'>" & vbCrLf
html = html & "<table role='presentation' cellpadding='0' cellspacing='0' border='0' style='font-family:Arial,sans-serif;font-size:10pt;color:#333333;'>" & vbCrLf
html = html & "<tr>" & vbCrLf
html = html & "<td style='padding-right:14px;vertical-align:top;border-right:2px solid " & ACCENT_COLOR & ";'>"
html = html & "<a href='" & COMPANY_WEBSITE & "'><img src='" & COMPANY_LOGO_URL & "' alt='Company logo' width='120' style='display:block;border:0;width:120px;height:auto;'></a></td>" & vbCrLf
html = html & "<td style='padding-left:14px;vertical-align:top;line-height:1.45;'>" & vbCrLf
html = html & "<strong style='font-size:13pt;color:" & ACCENT_COLOR & ";'>" & fullName & "</strong><br>" & vbCrLf

If jobTitle <> "" Then html = html & "<span>" & jobTitle & "</span><br>" & vbCrLf
If companyName <> "" Then html = html & "<span>" & companyName & "</span><br>" & vbCrLf
If officePhone <> "" Then html = html & "<span><strong>ტელ:</strong> " & officePhone & "</span><br>" & vbCrLf
If mobilePhone <> "" Then html = html & "<span><strong>მობ:</strong> " & mobilePhone & "</span><br>" & vbCrLf
If emailAddress <> "" Then html = html & "<span><strong>ელფოსტა:</strong> <a href='mailto:" & emailAddress & "' style='color:" & ACCENT_COLOR & ";text-decoration:none;'>" & emailAddress & "</a></span><br>" & vbCrLf

html = html & "<a href='" & COMPANY_WEBSITE & "' style='color:" & ACCENT_COLOR & ";text-decoration:none;'>" & HtmlEncode(COMPANY_WEBSITE) & "</a><br>" & vbCrLf
html = html & "<span style='color:#666666;'>" & HtmlEncode(COMPANY_ADDRESS) & "</span>" & vbCrLf
html = html & "</td></tr></table></body></html>"

Set stream = CreateObject("ADODB.Stream")
stream.Type = 2
stream.Charset = "utf-8"
stream.Open
stream.WriteText html
stream.SaveToFile signatureFile, 2
stream.Close
Set stream = Nothing

On Error Resume Next
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "Signature file was created, but Microsoft Word could not be started: " & signatureFile
    WScript.Quit 2
End If

Set signatureOptions = wordApp.EmailOptions.EmailSignature
signatureOptions.NewMessageSignature = signatureName
signatureOptions.ReplyMessageSignature = signatureName
wordApp.Quit

If Err.Number <> 0 Then
    WScript.Echo "Signature file was created, but it could not be set as default: " & signatureFile
    WScript.Quit 3
End If
On Error GoTo 0

WScript.Echo "Signature created successfully: " & signatureFile
WScript.Quit 0

Function ReadAdValue(ByRef adObject, ByVal attributeName)
    Dim value
    value = ""

    On Error Resume Next
    value = adObject.Get(attributeName)
    If Err.Number <> 0 Or IsNull(value) Or IsEmpty(value) Then
        Err.Clear
        value = ""
    End If
    On Error GoTo 0

    ReadAdValue = CStr(value)
End Function

Function HtmlEncode(ByVal value)
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, Chr(34), "&quot;")
    value = Replace(value, "'", "&#39;")
    HtmlEncode = value
End Function

Function SafeFileName(ByVal value)
    Dim invalidChars, item
    invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")

    For Each item In invalidChars
        value = Replace(value, item, "_")
    Next

    SafeFileName = value
End Function

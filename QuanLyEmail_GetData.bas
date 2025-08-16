Attribute VB_Name = "QuanLyEmail_GetData"
' PHAN NAY LAY THONG TIN EMAIL, REPORT, MEETING CO TRONG FOLDER INBOX, SEND
' Can kiem tra truoc olItem thuoc dang: MailItem, ReportItem, MeetingItem
'
'SentOn
'ReceivedOn
'Subject
'Body
'SenderName
'SenderEmail
'CTo
'CC
'Attach
'Cate
'Importance
'Flag

'////////////////////////////////////////////
' PHAN LAY DU LIEU RIENG LE
'////////////////////////////////////////////

Private Function GetEmailFromRecipient(receip As Outlook.Recipient) As String

On Error GoTo Error_Handle
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    Dim result As String
    Dim pa As Outlook.PropertyAccessor
    Set pa = receip.PropertyAccessor
    result = pa.GetProperty(PR_SMTP_ADDRESS)
    Set pa = Nothing
    GetEmailFromRecipient = result
Exit_Func:
    Exit Function
Error_Handle:
    GetEmailFromRecipient = LamSachChuoi(receip.name) & "@mail.server"
End Function

Function GetRecipient(olItem As Object, ToOrCC As Outlook.OlMailRecipientType) As String
Dim result As String
result = ""

If (TypeName(olItem) = "MailItem") Or (TypeName(olItem) = "MeetingItem") Then
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Set recips = olItem.Recipients 'Chi co MailItem va MeetingItem co thuoc tinh nay
    For Each recip In recips
        If recip.Type = ToOrCC Then
            result = result & recip.name & " <" & GetEmailFromRecipient(recip) & ">; "
        End If
    Next
    Set receips = Nothing
   'Neu(TypeName(olItem) = "ReportItem") thi result = ""
End If

If Right(result, 2) = "; " Then
    Dim iLen As Long
    iLen = Len(result) - 2
    result = Left(result, iLen)
End If

GetRecipient = result

End Function

Function GetTo(olItem As Object) As String
    GetTo = GetRecipient(olItem, olTo)
End Function

Function GetCC(olItem As Object) As String
    GetCC = GetRecipient(olItem, olCC)
End Function

Public Function GetSenderEmail(olItem As Object) As String
Dim result As String
If TypeOf olItem Is Outlook.ReportItem Then
    result = "report.item@mail.server"
End If

If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
    Dim receip As Outlook.Recipient
    Dim replyMail As Outlook.MailItem
    If olItem.SenderEmailType = "SMTP" Then
        result = olItem.SenderEmailAddress
    Else
        If olItem.SenderEmailType = "EX" Then
            Set replyMail = olItem.Reply
            For Each receip In replyMail.Recipients
                result = GetEmailFromRecipient(receip)
            Next receip
            replyMail.Close olDiscard
            Set replyMail = Nothing
        Else
            result = "no.sender@mail.server"
        End If
    End If
End If
GetSenderEmail = result
End Function

Function GetAttachFileNames(olItem As Object) As String
    Dim result As String
    result = ""
    Dim atts As Outlook.Attachments
    Dim frd As Outlook.MailItem
        
    If TypeOf olItem Is Outlook.MailItem Then
        If olItem.BodyFormat = olFormatRichText Then
            Set frd = olItem.Forward
            frd.Display
            ActiveInspector.CommandBars.ExecuteMso ("MessageFormatHtml")
            Set atts = frd.Attachments
            'frd.BodyFormat = olFormatHTML
        End If
    End If
    
    If atts Is Nothing Then
        Set atts = olItem.Attachments
    End If
    
    Dim att As Outlook.Attachment
    For Each att In atts
        If att.Type <> olOLE Then
            result = result & att.fileName & "; "
        End If
    Next
    
    Set atts = Nothing
        
    If Not (frd Is Nothing) Then
        frd.Close olDiscard
        Set frd = Nothing
    End If
    
    If Right(result, 2) = "; " Then
        Dim iLen As Integer
        iLen = Len(result) - 2
        result = Left(result, iLen)
    End If

    GetAttachFileNames = result
End Function

Function GetBody(olItem As Object) As String

On Error GoTo Error_Handle
    'olItem.BodyFormat = olFormatHTML
    GetBody = LoaiBoNewLine(olItem.Body)
Exit_Func:
    Exit Function
Error_Handle:
    GetBody = err.Number & ":" & err.Description
End Function


'//////////////////////
'Categories la string roi
'Importace co san cho ca ba
'//////////////////////////

Function GetSendOn(olItem As Object) As String
GetSendOn = "1900-01-01 00:00:01"
If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
    GetSendOn = Format(olItem.SentOn, "yyyy-mm-dd hh:nn:ss")
End If
If TypeOf olItem Is Outlook.ReportItem Then
    GetSendOn = Format(olItem.CreationTime, "yyyy-mm-dd hh:nn:ss")
End If
End Function

Function GetReceivedTime(olItem As Object) As String
GetReceivedTime = "1900-01-01 00:00:01"
If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
    GetReceivedTime = Format(olItem.ReceivedTime, "yyyy-mm-dd hh:nn:ss")
End If
If TypeOf olItem Is Outlook.ReportItem Then
    GetReceivedTime = Format(olItem.LastModificationTime, "yyyy-mm-dd hh:nn:ss")
End If
End Function
Function GetFlagRequest(olItem As Object) As String
Dim rs As String
rs = ""
If TypeOf olItem Is Outlook.MailItem Then 'Chi co Mail co thuoc tinh nay
    rs = olItem.FlagRequest
End If
GetFlagRequest = rs
End Function
Function GetSenderName(olItem As Object) As String
If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
    GetSenderName = olItem.SenderName
Else
    GetSenderName = "MailServer"
End If
End Function
'/////////////////////////////////////////////
Private Function BuildTag(tagName As String, tagValue As String) As String
    Const UniqueTag As String = "UniqueTag19z"
    If IsNull(tagValue) Then
       tagValue = ""
    End If
    tagValue = Replace(tagValue, UniqueTag, "-")
    BuildTag = "<" & UniqueTag & tagName & ">" & tagValue & "</" & UniqueTag & tagName & ">" & vbNewLine
End Function

Private Function oItem2Text(olItem As Object) As String
    Dim result As String
    result = result & BuildTag("SentOn", GetSendOn(olItem))
    result = result & BuildTag("ReceivedOn", GetReceivedTime(olItem))
    result = result & BuildTag("SenderEmailAddress", GetSenderEmail(olItem))
    result = result & BuildTag("SenderName", GetSenderName(olItem))
    result = result & BuildTag("CTo", GetTo(olItem))
    result = result & BuildTag("CC", GetCC(olItem))
    result = result & BuildTag("Categories", olItem.Categories)
    result = result & BuildTag("Importance", olItem.Importance)
    result = result & BuildTag("FlagRequest", GetFlagRequest(olItem))
    result = result & BuildTag("Subject", olItem.Subject)
    result = result & BuildTag("Body", GetBody(olItem))
    result = result & BuildTag("AttachmentNames", GetAttachFileNames(olItem))
    oItem2Text = result
End Function

'===================================
' HAM PUBLIC DE MODULE KHAC SU DUNG
'===================================
Public Sub oItemSave(olItem As Object, sfile As String)

'Dim olItem As Object
'Set olItem = Application.ActiveExplorer.Selection(1)

Dim s As String
s = oItem2Text(olItem)

Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2 'Specify stream type - we want To save text/string data.
fsT.Charset = "utf-8" 'Specify charset For the source text data.
fsT.Open 'Open the stream And write binary data To the object
fsT.WriteText s
fsT.SaveToFile sfile, 2 'Save binary data To disk
fsT.Close
Set fsT = Nothing

End Sub

'////////////////////////////////////////////
' PHAN LAY DU LIEU VAO CUSTOM TYPE
'////////////////////////////////////////////

Public Function GetTypeMail(olItem As Object) As TypeMail

 Dim result As TypeMail
 
    result.SentOn = GetSendOn(olItem)
    'result.SentOnInt = GetSendOn(olItem)
    result.ReceivedOn = GetReceivedTime(olItem)
    result.SenderEmailAddress = GetSenderEmail(olItem)
    result.SenderName = GetSenderName(olItem)
    result.CTo = GetTo(olItem)
    result.CC = GetCC(olItem)
    result.Categories = olItem.Categories
    result.Importance = olItem.Importance
    result.FlagRequest = GetFlagRequest(olItem)
    result.Subject = olItem.Subject
    result.Body = GetBody(olItem)
    result.AttachmentNames = GetAttachFileNames(olItem)
    
GetTypeMail = result

End Function

Type TypeMail
    MailId As Long
    SentOn As String
    SentOnInt As Long
    ReceivedOn As String
    ReceivedOnInt As Long
    Subject As String
    Body As String
    SenderEmailAddress As String
    SenderEmailAddressInt As Long
    SenderName As String
    CTo As String
    CC As String
    AttachmentNames As String
    Categories As String
    Importance As String
    FlagRequest As String
    CTag As String
    LastWriteTimeInt As Long
    IsLegalRequest As Integer
    RequestStatus As Integer
    CPath As String
End Type

Public Function DateToInt(ByVal dt As Date) As Long 'Tuong ung voi ham tren C#
    Dim dRoot As Date
    dRoot = DateSerial(2020, 1, 1)
    
    DateToInt = DateDiff("d", dRoot, dt)
End Function





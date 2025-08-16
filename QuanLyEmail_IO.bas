Attribute VB_Name = "QuanLyEmail_IO"
'============================================
'
' THIET LAP CAC THONG SO
Const ConstRootSavePath As String = "D:\BACK_UP\OutlookData"  'Su dung o ham  SaveOutlookItems
Const ConstRootSavePathQuanLyCongViec As String = "D:\0-MailRequests" 'Su dung o ham  Save_1_OutlookItem_QuanLyCongViec
'
'============================================
Private Function CreateMyFolder(fso As Scripting.FileSystemObject, strPath As String)
Dim strTempPath As String
Dim lngPath As Long
Dim vPath As Variant
    vPath = Split(strPath, "\")
    strPath = vPath(0) & "\"
    For lngPath = 1 To UBound(vPath)
        strPath = strPath & vPath(lngPath) & "\"
        If Not fso.FolderExists(strPath) Then MkDir strPath
    Next lngPath
lbl_Exit:
    Exit Function
End Function

Private Function WriteLog(ts As TextStream, strText As String)
' Ham nay de ghi log cho ham SaveOutlookItem
    ts.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & CStr(strText)
End Function

Private Function ChangeCodeBase_Selection_Or_CurrentFolder(IsSelection As Boolean) As String
'Ham nay de tranh loi do email dang dat code Korea, etc dan den khong the luu file duoc

Dim objOL As Outlook.Application
Dim objItems As Object
Dim olItem As Object
 
Set objOL = Outlook.Application

If (IsSelection = True) Then
    Set objItems = objOL.ActiveExplorer.Selection
Else
    Set objItems = objOL.ActiveExplorer.CurrentFolder.Items
End If

Dim num As Long

Dim idx As Long
For idx = 1 To objItems.Count
    Set olItem = objItems.Item(idx)
    If (TypeName(olItem) = "MailItem") Then
      On Error Resume Next
      If olItem.InternetCodepage <> 65001 Then
            olItem.InternetCodepage = 65001
            olItem.Save
            num = num + 1
      End If
    End If
Next idx

ChangeCodeBase_Selection_Or_CurrentFolder = "Total mails changed to Unicode: " & CStr(num)

Set olItem = Nothing
Set objItems = Nothing
Set objOL = Nothing

End Function

Private Function PathCombine(sFirst As String, sSecond As String) As String
If Right(sFirst, 1) = "\" Then
    PathCombine = sFirst & sSecond
Else
    PathCombine = sFirst & "\" & sSecond
End If
End Function

Private Function PathFromObject(rootPath As String, olItem As Object) As String
'Ham nay de tao ra duong dan: [rootPath]\yyyy\mm\[sender]\[date]_[time]-[subject]
    Dim fPath, fPath2 As String
    If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
        fPath2 = Format(olItem.ReceivedTime, "yyyy.mm.dd") & "_" & Format(olItem.ReceivedTime, "hh.nn") & "-" & LamSachChuoi(olItem.Subject)
        fPath = PathCombine(rootPath, Format(olItem.ReceivedTime, "yyyy")) & "\" & Format(olItem.ReceivedTime, "mm") & "\" & LamSachChuoi(olItem.SenderName) & "\" & fPath2
    End If
    
    If TypeOf olItem Is Outlook.ReportItem Then
        fPath2 = Format(olItem.LastModificationTime, "yyyy.mm.dd") & "_" & Format(olItem.LastModificationTime, "hh.nn") & "-" & LamSachChuoi(olItem.Subject)
        fPath = PathCombine(rootPath, Format(olItem.LastModificationTime, "yyyy")) & "\" & Format(olItem.LastModificationTime, "mm") & "\" & "Mail_Server" & "\" & fPath2
    End If
    
    Dim lenPath As Integer
    lenPath = Len(fPath)
    If lenPath > 150 Then
       fPath = Left(fPath, 147) & "_" & CStr(lenPath - 147)
    End If
    PathFromObject = fPath
End Function

Private Function SaveOutlookItem(oItem As Object, strPath As String, _
                                    tsSucess As TextStream, tsErr As TextStream, fso As Scripting.FileSystemObject) As Integer

'Ham nay de ghi 01 OutlookItem, bao gom ca ghi log. De su dung cho ham SaveOutlookItems_Selection_Or_CurrentFolder
'Chua bao gom ChangeCodeBase

Dim result As Integer
Dim msgFileName As String
msgFileName = "outlook.msg"
Dim txtFileName As String
txtFileName = "outlook.txt"

On Error GoTo WriteLog
    If Not fso.FileExists(PathCombine(strPath, msgFileName)) Then
        oItem.SaveAs PathCombine(strPath, msgFileName), OlSaveAsType.olMSGUnicode
    End If
    
    If Not fso.FileExists(PathCombine(strPath, txtFileName)) Then
        oItemSave oItem, PathCombine(strPath, txtFileName)
    End If
    
    If oItem.Attachments.Count > 0 Then
       'Tao folder moi
       Dim fPath As String
       fPath = PathCombine(strPath, "Attachments")
       CreateMyFolder fso, fPath
            
       'Luu attach files
    Dim atts As Outlook.Attachments
    Dim frd As Outlook.MailItem
        
    If TypeOf oItem Is Outlook.MailItem Then
        If oItem.BodyFormat = olFormatRichText Then
            Set frd = oItem.Forward
            frd.Display
            ActiveInspector.CommandBars.ExecuteMso ("MessageFormatHtml")
            Set atts = frd.Attachments
            'frd.BodyFormat = olFormatHTML
        End If
    End If
    
    If atts Is Nothing Then
        Set atts = oItem.Attachments
    End If
            
       Dim att As Outlook.Attachment
       For Each att In atts
       If att.Type <> olOLE Then
       
            Dim sfile As String
            sfile = fso.GetBaseName(att.DisplayName)
            sfile = LamSachTenFile(sfile)
            sfile = PathCombine(fPath, sfile)
            Dim lenFile As Integer
            lenFile = Len(sfile)
            
            If lenFile > 250 Then
                sfile = Left(sfile, 245) & "_" & CStr(lenFile - 245)
            End If
            
            sfile = sfile & Right(att.fileName, Len(att.fileName) - InStrRev(att.fileName, ".") + 1) ' Thay cho: "." & fso.GetExtensionName(att.DisplayName)
            If Not fso.FileExists(sfile) Then
                att.SaveAsFile sfile
            End If
            
        End If
        Next att
        
        If Not (frd Is Nothing) Then
            frd.Close olDiscard
            Set frd = Nothing
        End If
        
        WriteLog tsSucess, "Luu thanh cong: " & PathCombine(strPath, msgFileName) 'Ghi log
        result = 0
    End If
    
Exit_Func:
    SaveOutlookItem = result
    Exit Function

WriteLog:
    WriteLog tsErr, err.Number & ":" & err.Description
    WriteLog tsErr, "Vui long kiem tra: " & PathCombine(strPath, msgFileName)
    WriteLog tsErr, "================================="
    result = -1
    GoTo Exit_Func

End Function

Private Function SaveOutlookItems_Selection_Or_CurrentFolder(IsSelection As Boolean) As String

'Ham nay de ghi cac OutlookItem o Folder hoac Selection

Dim objOL As Outlook.Application
Dim objItems As Object
Dim olItem As Object

Set objOL = Outlook.Application

If (IsSelection = True) Then
    Set objItems = objOL.ActiveExplorer.Selection
Else
    Set objItems = objOL.ActiveExplorer.CurrentFolder.Items
End If

If objItems.Count = 0 Then
    SaveOutlookItems_Selection_Or_CurrentFolder = "No item"
Else

Dim fso As Scripting.FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

Dim numItem As Integer
Dim numMailItem As Integer
Dim numReportItem As Integer
Dim numMeetingItem As Integer
Dim numSaveErr As Integer
Dim otherItemName As String

numItem = objItems.Count
numSaveErr = 0
numOther = 0

Dim fPath As String
fPath = ConstRootSavePath 'Thong tin cau hinh o tren dau

CreateMyFolder fso, fPath
    
'Tao file ghi log
Dim tsSucess As TextStream
Dim tsErr As TextStream
    
sFileNameLogSucess = PathCombine(fPath, "SucessLog.txt")
sFileNameLogError = PathCombine(fPath, "ErrorLog.txt")
    
If fso.FileExists(sFileNameLogSucess) Then
    Set tsSucess = fso.OpenTextFile(sFileNameLogSucess, ForAppending)
    tsSucess.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Append ==="
Else
    Set tsSucess = fso.CreateTextFile(sFileNameLogSucess, True)
    tsSucess.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Creates ==="
End If
    
If fso.FileExists(sFileNameLogError) Then
    Set tsErr = fso.OpenTextFile(sFileNameLogError, ForAppending)
    tsErr.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Append ==="
Else
    Set tsErr = fso.CreateTextFile(sFileNameLogError, True)
    tsErr.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Creates ==="
End If

Dim i As Long
For i = 1 To objItems.Count
    Set olItem = objItems.Item(i)
    
    If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Or (TypeOf olItem Is Outlook.ReportItem) Then
        'Luu ve dang [rootPath]\yyyy\MM\[Sender]\[Date]-[Subject]
        Dim fPath2 As String
        fPath2 = PathFromObject(fPath, olItem)
        CreateMyFolder fso, fPath2
               
        Dim iErr As Integer
        iErr = SaveOutlookItem(olItem, fPath2, tsSucess, tsErr, fso)
        
        'Danh dau file khong luu duoc
        If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
            If (iErr = -1) Then
                olItem.Categories = "NotSave"
                olItem.Save
            End If
                
            If iErr = 0 Then
                If InStr(1, olItem.Categories, "NotSave", vbTextCompare) Then
                    olItem.Categories = ""
                    olItem.Save
                End If
            End If
        End If
        'Thong ke loi
        numSaveErr = numSaveErr - iErr
    Else
        'Thong ke item khong nam trong danh muc
        numOther = numOther + 1
    End If
'Next olItem
 Next i
 
tsErr.Close
tsSucess.Close
    
Set tsSucess = Nothing
Set tsErr = Nothing
Set fso = Nothing

Set olItem = Nothing
Set objItems = Nothing
Set objOL = Nothing

SaveOutlookItems_Selection_Or_CurrentFolder = "Total item: " & CStr(numItem) & ", not save: " & CStr(numSaveErr) & " other items: " & CStr(numOther) & " . Please check log file for further information."

End If

End Function

Private Function Save_This_Msg() As String

'Ham nay de ghi xuong 1 mail duoc chon,
'Duoc chon duong dan va khong ghi log
'Da change InternetCodepage
'Khong tao folder luu file dinh kem
'Luu duoi ten [select folder]\[date]_[time]-[sender]-[subject]

Dim result As String

If Application.ActiveExplorer.Selection.Count = 0 Then
    result = "No selection"
Else
    Dim olItem As Object
    Set olItem = Application.ActiveExplorer.Selection.Item(1)

    If (TypeName(olItem) = "MailItem") Then
        If olItem.InternetCodepage <> 65001 Then
            olItem.InternetCodepage = 65001
            olItem.Save
        End If
    End If
    

    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fPath As String
    fPath = "D:\Mail"
    fPath = InputBox("Enter the path to save the message.", _
                     "Save Message", fPath)
    
    CreateMyFolder fso, fPath
                    
    If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Or (TypeOf olItem Is Outlook.ReportItem) Then
        
        Dim fPath2 As String
        If TypeOf olItem Is Outlook.ReportItem Then
            fPath2 = Format(olItem.LastModificationTime, "yyyy.mm.dd") & "_" & Format(olItem.LastModificationTime, "hh.nn") & "-" & "Mail_Server" & "-" & LamSachChuoi(olItem.Subject)
        Else
            fPath2 = Format(olItem.ReceivedTime, "yyyy.mm.dd") & "_" & Format(olItem.ReceivedTime, "hh.nn") & "-" & LamSachChuoi(olItem.SenderName) & "-" & LamSachChuoi(olItem.Subject)
        End If
        
        If Len(fPath2) > 150 Then
            fPath2 = Left(fPath2, 150) & "_" & CStr(Len(fPath2) - 150)
        End If
        
        'Luu duoi ten [select folder]\[date]_[time]-[SENDER]-[subject]
        Dim msgFileName As String
        msgFileName = PathCombine(fPath, fPath2) & ".msg"
        
        If Not fso.FileExists(msgFileName) Then
            olItem.SaveAs msgFileName, OlSaveAsType.olMSGUnicode
        End If

        result = msgFileName
    End If

    Set fso = Nothing
End If

Save_This_Msg = result

End Function

Private Function Save_Selection_NoLog() As String

'Ham nay de ghi xuong cac mail lua chon, duoc chon noi luu va khong ghi log
'Da change InternetCodepage
'Khong tao folder luu file dinh kem

Dim result As String

If Application.ActiveExplorer.Selection.Count = 0 Then
    result = "-"
Else
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim rootPath As String
    rootPath = "D:\BACK_UP\OutlookData"
    rootPath = InputBox("Enter the path to save the message.", _
                     "Save Message", fPath)
    CreateMyFolder fso, rootPath

    Dim olItem As Object
    Dim i As Long
    For i = Application.ActiveExplorer.Selection.Count To 1 Step -1
        Set olItem = Application.ActiveExplorer.Selection.Item(i)
        If (TypeName(olItem) = "MailItem") Then
            If olItem.InternetCodepage <> 65001 Then
                olItem.InternetCodepage = 65001
                olItem.Save
            End If
        End If
                        
        If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Or (TypeOf olItem Is Outlook.ReportItem) Then
            'Luu ve dang [rootPath]\yyyy\MM\[Sender]\[Date]-[Subject]
            Dim fPath2 As String
            Dim fPath As String
            If TypeOf olItem Is Outlook.ReportItem Then
                fPath2 = Format(olItem.LastModificationTime, "yyyy.mm.dd") & "_" & Format(olItem.LastModificationTime, "hh.nn") & "-" & LamSachChuoi(olItem.Subject)
                fPath = PathCombine(rootPath, Format(olItem.LastModificationTime, "yyyy")) & "\" & Format(olItem.LastModificationTime, "mm") & "\" & "Mail_Server"
            Else
                fPath2 = Format(olItem.ReceivedTime, "yyyy.mm.dd") & "_" & Format(olItem.ReceivedTime, "hh.nn") & "-" & LamSachChuoi(olItem.Subject)
                fPath = PathCombine(rootPath, Format(olItem.ReceivedTime, "yyyy")) & "\" & Format(olItem.ReceivedTime, "mm") & "\" & LamSachChuoi(olItem.SenderName)
            End If
            
            CreateMyFolder fso, fPath

            Dim msgFileName As String
            msgFileName = PathCombine(fPath, fPath2)
            If Len(msgFileName) > 200 Then
                fPath = Left(msgFileName, 200) & "_" & CStr(Len(msgFileName) - 200)
            End If
            msgFileName = msgFileName & ".msg"
        
            If Not fso.FileExists(msgFileName) Then
                olItem.SaveAs msgFileName, OlSaveAsType.olMSGUnicode
            End If

            result = msgFileName
  
        End If
    Next
    Set fso = Nothing
End If

Save_Selection_NoLog = result

End Function

'===================================
'
' THU TUC, HAM PUBLIC DE SU DUNG
'
'===================================

Public Function Save_New_OutlookItem(olItem As Object) As String
' Ham nay de trigger trong StartUp (mail moi gui, mail moi den)

Dim result As String
result = "Un-defined"

If IsNull(olItem) Then
    result = "No item"
Else

Dim fso As Scripting.FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

Dim fPath As String
fPath = ConstRootSavePath 'Thong tin cau hinh o tren dau
CreateMyFolder fso, fPath
    
'Tao file ghi log
Dim tsSucess As TextStream
Dim tsErr As TextStream
    
sFileNameLogSucess = PathCombine(fPath, "SucessLog.txt")
sFileNameLogError = PathCombine(fPath, "ErrorLog.txt")
    
If fso.FileExists(sFileNameLogSucess) Then
    Set tsSucess = fso.OpenTextFile(sFileNameLogSucess, ForAppending)
    tsSucess.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Append ==="
Else
    Set tsSucess = fso.CreateTextFile(sFileNameLogSucess, True)
    tsSucess.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Creates ==="
End If
    
If fso.FileExists(sFileNameLogError) Then
    Set tsErr = fso.OpenTextFile(sFileNameLogError, ForAppending)
    tsErr.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Append ==="
Else
    Set tsErr = fso.CreateTextFile(sFileNameLogError, True)
    tsErr.WriteLine Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2) & " > " & "=== LogFile Creates ==="
End If

If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Or (TypeOf olItem Is Outlook.ReportItem) Then
   
   Dim fPath2 As String
   fPath2 = PathFromObject(fPath, olItem)
   CreateMyFolder fso, fPath2
               
   Dim iErr As Integer
   iErr = SaveOutlookItem(olItem, fPath2, tsSucess, tsErr, fso)
        
   'Danh dau file khong luu duoc
   If (TypeOf olItem Is Outlook.MailItem) Or (TypeOf olItem Is Outlook.meetingItem) Then
      If (iErr = -1) Then
                olItem.Categories = "NotSave"
                olItem.Save
                result = "Saving Error"
            End If
                
            If iErr = 0 Then
                If InStr(1, olItem.Categories, "NotSave", vbTextCompare) Then
                    olItem.Categories = ""
                    olItem.Save
                End If
                result = "Saved"
            End If
        End If
End If

tsErr.Close
tsSucess.Close
    
Set tsSucess = Nothing
Set tsErr = Nothing
Set fso = Nothing

End If

Save_New_OutlookItem = result

End Function
Public Sub ExportMsg_CurrentFolder()
    Dim s1 As String
    Dim s2 As String
    
    s1 = ChangeCodeBase_Selection_Or_CurrentFolder(False)
    s2 = SaveOutlookItems_Selection_Or_CurrentFolder(False)
    MsgBox s1
    MsgBox s2
End Sub
Public Sub ExportMsg_Selection()
    Dim s1 As String
    Dim s2 As String
    
    s1 = ChangeCodeBase_Selection_Or_CurrentFolder(True)
    s2 = SaveOutlookItems_Selection_Or_CurrentFolder(True)
    MsgBox s1
    MsgBox s2
End Sub
Public Sub Save_This_Mail()
    Dim s1 As String
    s1 = Save_This_Msg
    MsgBox s1
End Sub

Public Sub Save_Selection()
    Dim s1 As String
    s1 = Save_Selection_NoLog
    MsgBox s1
End Sub

Attribute VB_Name = "SaveMail"
'
'
'
'Const ConstRootPath As String = "D:\BACK_UP\OutlookData"
Const ConstRootPath As String = "D:\Test"
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
Private Function PathCombine(sFirst As String, sSecond As String) As String
If Right(sFirst, 1) = "\" Then
    PathCombine = sFirst & sSecond
Else
    PathCombine = sFirst & "\" & sSecond
End If
End Function
'******************************************
' Ham xu ly string
'******************************************
Private Function LamSachChuoi(ByVal sContent As String) As String
     Dim i As Long
     Dim intCode As Long
     Dim sChar As String
     Dim sConvert As String
     For i = 1 To Len(sContent)
        sChar = Mid(sContent, i, 1)
        If sChar <> "" Then
            intCode = AscW(sChar)
        End If
        Select Case intCode
            Case 273
                sConvert = sConvert & "d"
            Case 272
                sConvert = sConvert & "D"
            Case 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863
                sConvert = sConvert & "a"
            Case 192, 193, 194, 195, 258, 7840, 7842, 7844, 7846, 7848, 7850, 7852, 7854, 7856, 7858, 7860, 7862
                sConvert = sConvert & "A"
            Case 232, 233, 234, 7865, 7867, 7869, 7871, 7873, 7875, 7877, 7879
                sConvert = sConvert & "e"
            Case 200, 201, 202, 7864, 7866, 7868, 7870, 7872, 7874, 7876, 7878
                sConvert = sConvert & "E"
            Case 236, 237, 297, 7881, 7883
                sConvert = sConvert & "i"
            Case 204, 205, 296, 7880, 7882
                sConvert = sConvert & "I"
            Case 242, 243, 244, 245, 417, 7885, 7887, 7889, 7891, 7893, 7895, 7897, 7899, 7901, 7903, 7905, 7907
                sConvert = sConvert & "o"
            Case 210, 211, 212, 213, 416, 7884, 7886, 7888, 7890, 7892, 7894, 7896, 7898, 7900, 7902, 7904, 7906
                sConvert = sConvert & "O"
            Case 249, 250, 361, 432, 7909, 7911, 7913, 7915, 7917, 7919, 7921
                sConvert = sConvert & "u"
            Case 217, 218, 360, 431, 7908, 7910, 7912, 7914, 7916, 7918, 7920
                sConvert = sConvert & "U"
            Case 253, 7923, 7925, 7927, 7929
                sConvert = sConvert & "y"
            Case 221, 7922, 7924, 7926, 7928
                sConvert = sConvert & "Y"
            Case Else
                sConvert = sConvert & sChar
        End Select
     Next
     
     'Thay the ky tu dac biet
    Dim strSpecialChars As String
    Dim i2 As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]-_" & Chr(10) & Chr(13) 'Thay the cac ky tu dac biet khong su dung de dat ten file/FOLDER
    For i2 = 1 To Len(strSpecialChars)
        sConvert = Replace(sConvert, Mid$(strSpecialChars, i2, 1), " ")
    Next
    
    'Xoa het cac ky tu ngoai danh muc chu hoa, chu thuong (loai tieng Han)
    Dim i3 As Long
    For i3 = 1 To Len(sConvert)
        If InStr(1, "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz. _-", Mid(sConvert, i3, 1)) Then
            LamSachChuoi = LamSachChuoi & Mid(sConvert, i3, 1)
        End If
    Next
    
    ' Loai bo khoang trang du thua
    Do Until InStr(LamSachChuoi, "   ") = 0
        LamSachChuoi = Replace(LamSachChuoi, "   ", " ")
    Loop
    
    Do Until InStr(LamSachChuoi, "  ") = 0
        LamSachChuoi = Replace(LamSachChuoi, "  ", " ")
    Loop
    LamSachChuoi = Trim(LamSachChuoi)
        
    'Thay the khoang trang
    LamSachChuoi = Replace(LamSachChuoi, " ", "_")
        
    If Len(LamSachChuoi) > 60 Then
        LamSachChuoi = Left(LamSachChuoi, 60)
    End If
    
End Function

Public Sub Save_Selection()

'Ham nay de ghi xuong cac mail lua chon, duoc chon noi luu va khong ghi log
'Da change InternetCodepage
'Khong tao folder luu file dinh kem

If Application.ActiveExplorer.Selection.Count > 0 Then
    
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim rootPath As String
    rootPath = ConstRootPath
    'rootPath = InputBox("Enter the path to save the message.", "Save Message", fPath)
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

'Save_Selection_NoLog = result

End Sub

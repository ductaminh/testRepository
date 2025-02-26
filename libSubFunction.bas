Attribute VB_Name = "libSubFunction"
Option Explicit

Public bMethodAdd As Boolean

Public Const bgDataTrue = &HC0FFC0
Public Const bgDataFalse = &HC0C0FF
Public Const dfPassword = "201519143201519143201519143"

'===========================================================================
'   SUB COLLECTIION
'===========================================================================

'   Lock & Unlock Worksheet
Public Sub lock_sheet(strSheetName As String)
    Sheets(strSheetName).Protect Password:=dfPassword, AllowFiltering:=True
End Sub

Public Sub unlock_sheet(strSheetName As String)
    Sheets(strSheetName).Unprotect Password:=dfPassword
End Sub

'   Show Alert
Public Sub show_alert_fail(strContent As String)
    On Error GoTo ErrorHandle
    Application.Assistant.DoAlert UCase(UniConvert("Thoong baso", "Telex")), UniConvert(strContent, "Telex"), msoAlertButtonOK, 1, 0, 0, 0
    On Error GoTo 0
Exit Sub
ErrorHandle:
    MsgBox UniConvert(strContent, "Telex"), vbCritical
End Sub

Public Sub show_alert_success(strContent As String)
    On Error GoTo ErrorHandle
    Application.Assistant.DoAlert UCase(UniConvert("Thoong baso", "Telex")), UniConvert(strContent, "Telex"), msoAlertButtonOK, 4, 0, 0, 0
    On Error GoTo 0
Exit Sub
ErrorHandle:
    MsgBox UniConvert(strContent, "Telex"), vbInformation
End Sub

Public Sub show_alert_warning(strContent As String)
    On Error GoTo ErrorHandle
    Application.Assistant.DoAlert UCase(UniConvert("Thoong baso", "Telex")), UniConvert(strContent, "Telex"), msoAlertButtonOK, 3, 0, 0, 0
    On Error GoTo 0
Exit Sub
ErrorHandle:
    MsgBox UniConvert(strContent, "Telex"), vbExclamation
End Sub

'   Turn on/off worksheet events/alerts
Public Sub speed_off()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .CalculateBeforeSave = True
        .Cursor = xlDefault
        .StatusBar = False
        .EnableCancelKey = xlInterrupt
    End With
End Sub

Public Sub speed_on(Optional StatusBarMsg As String = "Please wait...")
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
        .StatusBar = StatusBarMsg
        .EnableCancelKey = xlErrorHandler
    End With
End Sub


'===========================================================================
'   FUNCTION COLLECTIION
'===========================================================================

'   Bo dau Tieng Viet
Public Function ConvertToUnSign(ByVal sContent As String) As String
    Dim i As Long
    Dim intCode As Long
    Dim sChar As String
    Dim sConvert As String
    On Error Resume Next
    ConvertToUnSign = AscW(sContent)
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
            Case 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863, 7953
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
            Case 768, 769, 770, 777, 795, 803
                
            Case Else
                sConvert = sConvert & sChar
        End Select
    Next
    sConvert = Replace(sConvert, "?", "")
    ConvertToUnSign = sConvert
    On Error GoTo 0
End Function

'   Lay dong cuoi
Public Function GetLastRow(SheetName As String, ColName As String, minValue As Integer)
    GetLastRow = Sheets(SheetName).Range(ColName & Rows.Count).End(xlUp).Row
    If GetLastRow < minValue Then GetLastRow = minValue
End Function

'   Lay cot cuoi
Public Function GetLastCol(SheetName As String)
    GetLastCol = Sheets(SheetName).Cells(1, Columns.Count).End(xlToLeft).Column
End Function

'   Question Confirm
Public Function Question_YesNo(strQuestion As String) As String
    On Error GoTo ErrorHandle
    Question_YesNo = Application.Assistant.DoAlert(UCase(UniConvert("Thoong baso", "Telex")), UniConvert(strQuestion, "Telex"), msoAlertButtonYesNo, 1, 0, 0, 0)
    On Error GoTo 0
Exit Sub
ErrorHandle:
    Question_YesNo = MsgBox(UniConvert(strQuestion, "Telex"), vbYesNo + vbQuestion)
End Function

'   Go tieng Viet
Function UniConvert(text As String, InputMethod As String) As String
    Dim VNI_Type, Telex_Type, CharCode, temp, i As Long
    
    UniConvert = text
    VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
            "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
            "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
            "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
            "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
    
    Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
            "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
            "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
            "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
            "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
    CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
            ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
            ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
            ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
            ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
            ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
            ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
            ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
        
    Select Case InputMethod
        Case Is = "VNI": temp = VNI_Type
        Case Is = "Telex": temp = Telex_Type
    End Select
    
    For i = 0 To UBound(CharCode)
        UniConvert = Replace(UniConvert, temp(i), CharCode(i))
        UniConvert = Replace(UniConvert, UCase(temp(i)), UCase(CharCode(i)))
    Next i
End Function

Function SplitNumber(strValue As String, bOption As Boolean) As String
    '   bOption = True: tach so; = False: tach chu
    On Error GoTo ErrorHandle
    With CreateObject("VBScript.RegExp")
        .Pattern = IIf(bOption = True, "\d+", "\D+")
        .Global = True
        SplitNumber = .Replace(strValue, "")
    End With
    On Error GoTo 0
Exit Function
ErrorHandle:
    SplitNumber = strValue
End Function

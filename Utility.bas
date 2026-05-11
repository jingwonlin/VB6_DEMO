Attribute VB_Name = "Utility"
Option Explicit

'Sub Trace(sMsg As String, Optional nType As nyLogType, Optional sModuleName As String)
'    On Error Resume Next
'    If Not (goTrace Is Nothing) Then
'        goTrace.LogFile = App.Path & "\LogMessage\" & Format$(Date, "yyyymmdd") & ".log"
'        goTrace.LogMessage sMsg, nType, sModuleName
''        goTrace.LogMessage nType, sModuleName, sMsg
'        Kill App.Path & "\LogMessage\" & Format$(Date - 60, "yyyymmdd") & ".log"
'    End If
'End Sub

'Public Sub Trace(sMsg As String, Optional nType As nyLogType, Optional sModuleName As String)
'End Sub
Public Function IniRead(lpAppName As String, lpKeyName As String, lpFileName As String) As String
   Dim size As Long '字串總長度
   Dim valid As Long '傳回有值的長度
   Dim resoult As String '傳回字串
   'Dim lpFileName As String '路徑
   Dim lpDefault As String '如果沒有 Key的話，傳回來的初始值
   Dim lpReturnString As String '傳回的 Buffer
   lpDefault = "KeyNull"
   lpReturnString = Space$(1024)
   size = Len(lpReturnString)
   If StrComp(Trim$(lpKeyName), "NULL", 1) = 0 Then
     valid = GetPrivateProfileString(lpAppName, vbNullString, lpDefault, lpReturnString, size, lpFileName)
   Else
     valid = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnString, size, lpFileName)
   End If
   resoult = Left$(lpReturnString, valid)
   If resoult = "KeyNull" Then
      'Err.Raise 3001, , lpFileName & " 檔找不到 [" & lpAppName & "] " & lpKeyName & "= 的描述"
       xMsgBox lpFileName & " Cannot Find Description Of [" & lpAppName & "] " & lpKeyName & " = ", vbOKOnly, "Err message"
       IniRead = ""
       End
   Else
      IniRead = resoult
   End If
End Function

Public Function IsChecked(V As Variant) As Integer
    If V Then
        IsChecked = 1
    Else
        IsChecked = 0
    End If
End Function

Public Function xMsgBox(sMsg As String, Optional lClaim As Variant, Optional sTitle As Variant, Optional lTime As Variant) As Long
Dim MF As Form
   Set MF = New frmMsgBox
   If IsMissing(lClaim) Then lClaim = 0
   If IsMissing(sTitle) Then
     sTitle = App.EXEName
   Else
       If Len(CStr(sTitle)) = 0 Then sTitle = App.EXEName
   End If
   If IsMissing(lTime) Then lTime = 0
   xMsgBox = MF.TriggerForm(sMsg, CLng(lClaim), CStr(sTitle), CLng(lTime))
End Function

Public Sub CenterForm(ByRef frm As Form)
Dim lTop As Long
Dim lLeft As Long
    lTop = ((Screen.Height - frm.Height) / 2)
    If (lTop < 0) Then lTop = 0
    lLeft = ((Screen.Width - frm.Width) / 2)
    If (lLeft < 0) Then lLeft = 0
    frm.Move lLeft, lTop
End Sub

Public Sub uApp_StatusBar(iPanID As Integer, sPanTxt As String)
    On Error GoTo ErrHandler
    MDImainForm.ctlStatusBar.Panels(iPanID).Text = sPanTxt
    Err.Clear
ErrHandler:
End Sub
Public Function apiGetOS_Platform(ByRef strPlatform As String, ByRef strVersion As String, ByRef strBuildNo As String, _
                                    ByRef strServicePk As String) As Long
    Dim lngPlatform As Long
    mtpyOSVerInfo.dwOSVersionInfoSize = Len(mtpyOSVerInfo)
    GetVersionEx mtpyOSVerInfo
    lngPlatform = mtpyOSVerInfo.dwPlatformId
    If (lngPlatform = VER_PLATFORM_WIN32_NT) Then
        If (mtpyOSVerInfo.dwMajorVersion > 4) Then
            strPlatform = "Microsoft Windows 2000"
        Else
            strPlatform = "Microsoft Windows NT"
        End If
    ElseIf (lngPlatform = VER_PLATFORM_WIN32_WINDOWS) Then
        If (mtpyOSVerInfo.dwMinorVersion = 0) Then
            strPlatform = "Microsoft Windows 95"
        Else
            strPlatform = "Microsoft Windows 98"
        End If
    ElseIf (lngPlatform = VER_PLATFORM_WIN32s) Then
        strPlatform = "Microsoft Windows 3.1"
    End If
    strVersion = FormatVersion(mtpyOSVerInfo.dwMajorVersion, mtpyOSVerInfo.dwMinorVersion)
    strBuildNo = mtpyOSVerInfo.dwBuildNumber
    strServicePk = Left$(mtpyOSVerInfo.szCSDVersion, 14)
    apiGetOS_Platform = lngPlatform
End Function
Public Function FormatVersion(lngMajor As Long, lngMinor As Long, Optional Revision) As String
    If IsMissing(Revision) Then
        FormatVersion = Format(lngMajor, "#0.") & Format(lngMinor, "00")
    Else
        FormatVersion = Format(lngMajor, "#0.") & Format(lngMinor, "00.") & Format(Revision, "0000")
    End If
End Function

Public Sub CentrolForm(oF As Form)
Dim iW As Long
Dim iH As Long
  iW = (Screen.Width - oF.Width) / 2
  If iW < 0 Then iW = 0
  iH = ((Screen.Height - oF.Height) / 2) - 180
  If iH < 0 Then iH = 0
  oF.Top = iH - 900
  oF.Left = iW
End Sub

Public Sub CentrolFormL(oF As Form)
Dim iW As Long
Dim iH As Long
  iW = (Screen.Width - oF.Width) / 2
  If iW < 0 Then iW = 0
  iH = ((Screen.Height - oF.Height) / 2)
  If iH < 0 Then iH = 0
  oF.Top = iH + 200
  oF.Left = iW
End Sub

Public Function lLv(a As Variant) As String
    If IsNull(a) Then
        lLv = "0"
    Else
        lLv = CStr(Val(a))
    End If
End Function
Public Function lv(a As Variant) As String
    If IsNull(a) Then
        lv = ""
    Else
        lv = CStr(a)
    End If
End Function

Public Function usmid(sst1 As String, iStart As Long, iWLn As Long) As String
Dim ikx As Long
Dim inx As Long
Dim sret As String
Dim sc As String
    
    inx = 0
    sret = ""
    If (Len(sst1) <> 0) Then
        For ikx = 1 To Len(sst1)
            inx = inx + 1
            sc = Mid$(sst1, ikx, 1)
            'If Asc(Mid$(sst1, ikx, 1)) < 0 Then
            If (Asc(Mid$(sst1, ikx, 1)) < 0 And iWLn >= 1) Then
                   inx = inx + 1
            End If
            If inx >= iStart And inx <= iStart + iWLn - 1 Then
               sret = sret & sc
            End If
            If inx > iStart + iWLn - 1 Then
               Exit For
            End If
            'If (inx > uiLen(sst1)) Then Exit For
        Next ikx
    End If
    usmid = sret
End Function

Public Function uiLen(sstr As String) As Long
Dim im As Long
Dim ikx As Long
    im = 0
    For ikx = 1 To Len(Trim(sstr))
        If Asc(Mid$(sstr, ikx, 1)) < 0 Then
           im = im + 1
        End If
    Next
    uiLen = Len(sstr) + im
End Function

Public Sub H_Copy(ByVal Form_Temp As Form, ByVal Pic_Temp As PictureBox, ByVal Rotation)
Dim hDC As Long
Dim sx As Long, sy As Long
Dim px As Long, py As Long
Dim frame_width As Integer
Dim Printer_value As Integer
    Printer_value = Printer.Orientation
    If Rotation <> 0 Then
        Printer.Orientation = 2
    End If
    Form_Temp.ScaleMode = vbPixels
    frame_width = 4

    ' 將 Pic_Temp 設定成與 Form 同樣大小
    sx = Form_Temp.Width / Screen.TwipsPerPixelX
    sy = Form_Temp.Height / Screen.TwipsPerPixelY
    Pic_Temp.Width = sx + frame_width                      ' +frame_width以列印螢幕框
    Pic_Temp.Height = sy + frame_width
    ' 將 Form 的圖像轉移到 Pic_Temp 上面
    Pic_Temp.AutoRedraw = True
    hDC = GetWindowDC(Form_Temp.hwnd)
    If Form_Temp.BorderStyle = 0 Then
        BitBlt Pic_Temp.hDC, 0, 0, sx, sy, hDC, 0, 0, vbSrcCopy
    Else
        BitBlt Pic_Temp.hDC, 0, 0, sx + frame_width, sy + frame_width, hDC, 0, 0, vbSrcCopy ' +frame_width以列印螢幕框
    End If
    ReleaseDC Form_Temp.hwnd, hDC
    Pic_Temp.AutoRedraw = False

    ' 將 Pic_Temp 的圖像設定成 Pic_Temp 的背景圖
    Set Pic_Temp.Picture = Pic_Temp.Image
    
    ' 列印 Pic_Temp 的背景圖
    Form_Temp.ScaleMode = vbTwips
    Printer.ScaleMode = vbTwips
    px = (Printer.ScaleWidth - Pic_Temp.Width) / 2
    py = (Printer.ScaleHeight - Pic_Temp.Height) / 2
    If px < 0 Or py < 0 Then
        Pic_Temp.Width = Pic_Temp.Width * 90 \ 100      ' 長或寬超出範圍後縮小成90%
        Pic_Temp.Height = Pic_Temp.Height * 90 \ 100    ' 長或寬超出範圍後縮小成90%
        px = (Printer.ScaleWidth - Pic_Temp.Width) / 2
        py = (Printer.ScaleHeight - Pic_Temp.Height) / 2
    End If
    Printer.PaintPicture Pic_Temp.Picture, px, py, Pic_Temp.Width, Pic_Temp.Height
    Printer.EndDoc
    Printer.Orientation = Printer_value
End Sub

Public Sub delay_sec(ByVal delay_sec As Double)
Dim start_sec As Double
    start_sec = Timer
    Do While Timer < start_sec + delay_sec
        DoEvents
    Loop
End Sub

Public Sub PUT_DLE_ETX_BCC()
Dim lErrorNo As Long
Dim sErrorMsg As String
Dim sTemp As String
Dim lLen As Long
Dim lRtn As Long
Dim i As Integer
    On Error GoTo ErrorHandler
        sTemp = StrConv(AGV_TempStr, vbFromUnicode)
        lLen = LenB(sTemp)
        Send_AGV_Str = StrConv(AGV_TempStr, vbFromUnicode)
        ReDim Preserve Send_AGV_Str(UBound(Send_AGV_Str) + 3)
        Send_AGV_Str(lLen) = DLE
        Send_AGV_Str(lLen + 1) = ETX
        BCC = 0
        For i = 0 To lLen + 1
            BCC = BCC Xor Send_AGV_Str(i)
        Next
        Send_AGV_Str(lLen + 2) = BCC
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[PUT_BCC]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[PUT_BCC]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub



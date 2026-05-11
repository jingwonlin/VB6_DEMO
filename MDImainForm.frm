VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm MDImainForm 
   BackColor       =   &H8000000C&
   Caption         =   "新港玻三自動倉庫輸送設備連線"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin MSCommLib.MSComm MSCommAGV 
      Left            =   1935
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   9
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommAR 
      Index           =   8
      Left            =   2475
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   8
      DTREnable       =   -1  'True
      RThreshold      =   50
   End
   Begin VB.Timer TimerNEW 
      Left            =   4095
      Top             =   3915
   End
   Begin MSCommLib.MSComm MSCommAR 
      Index           =   7
      Left            =   1890
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   16
   End
   Begin MSCommLib.MSComm MSCommAR 
      Index           =   6
      Left            =   1305
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
      RThreshold      =   12
   End
   Begin MSCommLib.MSComm MSCommAR 
      Index           =   5
      Left            =   720
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      RThreshold      =   16
   End
   Begin MSCommLib.MSComm MSCommAR 
      Index           =   4
      Left            =   135
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
      RThreshold      =   28
   End
   Begin VB.Timer TimerAGV 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5445
      Top             =   2115
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  '對齊表單下方
      Height          =   360
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   7500
      Width           =   11880
      Begin VB.PictureBox nyTrace1 
         Height          =   495
         Left            =   1080
         ScaleHeight     =   435
         ScaleWidth      =   10785
         TabIndex        =   2
         Top             =   0
         Width           =   10850
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  '單線固定
         Caption         =   "操作記錄"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   420
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1100
      End
   End
   Begin ComctlLib.StatusBar ctlStatusBar 
      Align           =   2  '對齊表單下方
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   8
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4057
            MinWidth        =   4057
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1941
            MinWidth        =   1941
            Text            =   "操作訊息"
            TextSave        =   "操作訊息"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5609
            MinWidth        =   5609
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "None"
            TextSave        =   "None"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1059
            MinWidth        =   1059
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   742
            MinWidth        =   742
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "2018/3/7"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2364
            MinWidth        =   2364
            TextSave        =   "下午 04:27"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu QxMnu 
      Caption         =   "功能選項(&1)"
      Index           =   1
      Begin VB.Menu Item1 
         Caption         =   "連線監視"
         Index           =   1
      End
      Begin VB.Menu Item1 
         Caption         =   "............."
         Index           =   2
      End
      Begin VB.Menu Item1 
         Caption         =   ".............."
         Index           =   3
      End
      Begin VB.Menu Item1 
         Caption         =   "結束作業"
         Index           =   4
      End
   End
End
Attribute VB_Name = "MDImainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSQLStr As String
Dim lErrorNo As Long
Dim sErrorMsg As String
Dim iStartCom As Integer
Dim iPortIndex As Integer
Dim iPlcCmdIndex As Integer
Dim iCraneCnt   As Integer
Dim CraneErrNo(1 To 5) As Integer
Dim CraneTemp(1 To 5) As Integer
Dim ChangeTo As String
Dim BCC_Err_Cnt As Integer
Dim LineNo(1 To 24) As Boolean
Dim AGV_CommandCnt As Integer
Dim iDataLen As Integer

Private Sub Item1_Click(index As Integer)
    Dim iIndx As Integer
    Select Case index
        Case 1
            frmComPort.Show
        Case 4     '結束應用程式
            Unload Me
            End
    End Select
End Sub

Private Sub MDIForm_Load()
Dim sDir As String
Dim iPortNo As Integer
    On Error GoTo ErrorHandler

    Call uApp_StatusBar(1, Me.name)
    frmApStartMsg.lblApStartMsg.Caption = Me.Caption & "啟動中,請稍候..."
    frmApStartMsg.Show
    DoEvents
    If App.PrevInstance Then
        xMsgBox "[" + App.EXEName + "] is already Running!"
        End
    End If
    
    sDir = Dir$(App.Path & "\IniFile\GF3Com.ini", vbDirectory)
    If Len(sDir) = 0 Then
        MkDir App.Path & "\IniFile"
        xMsgBox "Do not have GF3Com.ini in 'IniFile' sub-directory ! ", vbOKOnly + vbCritical, "Error Message"
        End
    End If
    sDir = Dir$(App.Path & "\LogMessage", vbDirectory)
    If Len(sDir) = 0 Then
        MkDir App.Path & "\LogMessage"
    End If
'    nyTrace1.StartLogging
    Call uApp_StatusBar(3, "資料讀取中")
    Set goTrace = nyTrace1
    
    Call DataBaseConnect
    Call InitialParameter
    Call init_parameter1
    
'    AGV_Port = 9
'    MSCommAGV.Settings = "9600,e,8,1"
'    MSCommAGV.PortOpen = True
'    If MSCommAGV.PortOpen = True Then
'        giAGVOpenOK = 1
'        Call HOST_To_AGV("K")
'    Else
'        giAGVOpenOK = -1
'    End If
    
    For iPortNo = 4 To 8         '輸送設備=COM4~COM8 (原紗,經軸,織軸,胚布,成品)
        MSCommAR(iPortNo).PortOpen = True
    Next iPortNo
    iPortIndex = 4
    TimerNEW.Interval = 600
    TimerNEW.Enabled = True
    
     TimerAGV.Interval = 2000
     TimerAGV.Enabled = False
     
'     For iPortNo = 4 To 8         '輸送設備=COM4~COM8 (原紗,經軸,織軸,胚布,成品)
'        If MSCommAR(iPortNo).PortOpen = True Then
'            giComOpenOK(iPortNo + 17) = 1
'        Else
'            giComOpenOK(iPortNo + 17) = -1
'        End If
'     Next iPortNo
    
    Unload frmApStartMsg
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[MDImainForm!MDIForm_Load]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[MDImainForm!MDIForm_Load]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = -1
        xMsgBox "無法在此方式下結束作業系統", vbExclamation + vbOKOnly, "操作訊息", 5
        Exit Sub
    End If
End Sub

Private Sub MDIForm_Resize()
'    If Me.WindowState = vbMinimized Or Me.WindowState = vbNormal Then
'        Me.WindowState = vbMaximized
'        xMsgBox "限制此功能", vbExclamation + vbOKOnly, "操作訊息", 5
'    End If
End Sub

Private Sub InitialParameter()
Dim i   As Integer
    On Error GoTo ErrorHandler
    
    For i = 1 To 5
          CraneErrNo(i) = 99
    Next
   
    For i = 1 To 3
        ST21_Temp(i) = "9"
    Next
    For i = 1 To 7
        ST22_Temp(i) = "9"
    Next
    For i = 1 To 6
        ST23_Temp(i) = "9"
    Next
    For i = 1 To 5
        ST24_Temp(i) = "9"
    Next
    For i = 1 To 9
        ST25_Temp(i) = "9"
    Next
    
    PLC_LEN(21) = 28     '原紗
    PLC_LEN(22) = 20 '16 '經軸     >2 B001 B002 B003 B004 B005 B006 B007 00000 B101 B102 B103 保留 CrLf
    PLC_LEN(23) = 14 '12 '織軸     >2 C001 C002 C003 C004 C005 C006 00 C102 保留 CrLf
    PLC_LEN(24) = 16     '胚布
    PLC_LEN(25) = 50     '成品
    
    ST21_StNo(1) = "A003":   ST21_StNo(2) = "A004":   ST21_StNo(3) = "A005"
    ST21_StByte(1) = 7:      ST21_StByte(2) = 9:      ST21_StByte(3) = 11
    
    ST22_StNo(1) = "B001":   ST22_StNo(2) = "B002":   ST22_StNo(3) = "B003":   ST22_StNo(4) = "B004":   ST22_StNo(5) = "B005":   ST22_StNo(6) = "B006":   ST22_StNo(7) = "B007"
    ST22_StByte(1) = 3:      ST22_StByte(2) = 4:      ST22_StByte(3) = 5:      ST22_StByte(4) = 6:      ST22_StByte(5) = 7:      ST22_StByte(6) = 8:      ST22_StByte(7) = 9
    
    ST23_StNo(1) = "C001":   ST23_StNo(2) = "C002":   ST23_StNo(3) = "C003":   ST23_StNo(4) = "C004":   ST23_StNo(5) = "C005":   ST23_StNo(6) = "C006"
    ST23_StByte(1) = 3:      ST23_StByte(2) = 4:      ST23_StByte(3) = 5:      ST23_StByte(4) = 6:      ST23_StByte(5) = 7:      ST23_StByte(6) = 8
    
    ST24_StNo(1) = "D001":   ST24_StNo(2) = "D002":   ST24_StNo(3) = "D003":   ST24_StNo(4) = "D004":   ST24_StNo(5) = "D005"
    ST24_StByte(1) = 3:      ST24_StByte(2) = 4:      ST24_StByte(3) = 5:      ST24_StByte(4) = 6:      ST24_StByte(5) = 7
'字串Byte數 123 4  5  6  7  8  9  0  1  2  3
'           >51 E2 E1 E5 E4 E3 E7 E6 E9 E8 E10
    ST25_StNo(1) = "E001":   ST25_StNo(2) = "E002":   ST25_StNo(3) = "E003":   ST25_StNo(4) = "E004":   ST25_StNo(5) = "E005":   ST25_StNo(6) = "E006":   ST25_StNo(7) = "E007":   ST25_StNo(8) = "E008":   ST25_StNo(9) = "E009":   ST25_StNo(10) = "E010"
    ST25_StByte(1) = 5:      ST25_StByte(2) = 4:      ST25_StByte(3) = 8:      ST25_StByte(4) = 7:      ST25_StByte(5) = 6:     ST25_StByte(6) = 10:     ST25_StByte(7) = 9:     ST25_StByte(8) = 12:     ST25_StByte(9) = 11:     ST25_StByte(10) = 13

    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[InitialParameter]:Error:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[InitialParameter]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub
Private Function dr_sub(awno1 As String, PLC_Port As Integer) As Boolean
Dim id As String
Dim lRtn As Integer
Dim sSend As String
Dim sBuf As String
Dim ii As Integer, jj As Integer, LL As Integer, buf As String * 210, buf1 As String, dr_buf As String * 210, brs2 As New ADODB.Recordset
    On Error GoTo ErrorHandler
      dr_sub = False
      Call sio_flush(PLC_Port, 2)
      id = "21"
      If MSCommAR(PLC_Port).PortOpen = True Then
          sSend = id & awno1 & awno1 & Chr(13) & Chr(10)
          MSCommAR(PLC_Port).Output = sSend
      End If
      dr_buf = "": ii = 0: jj = 0
      Do
        buf = ""
        LL = MSCommAR(iPortIndex).InBufferCount
        buf = MSCommAR(iPortIndex).Input
        If LL > 0 Then
          Mid(dr_buf, jj + 1, LL) = Mid(buf, 1, LL): jj = jj + LL
        Else
          ii = ii + 1
        End If
        If ii > 5 Or jj >= 12 Then Exit Do
        Call Wait(0.1)
      Loop
      'ii=1即可讀到
      frmComPort.lblComPortLen(PLC_Port + 17) = jj: frmComPort.lblComPortString(PLC_Port + 17) = dr_buf
      If jj <> 12 Then
        frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & "," & jj & ":" & Mid(dr_buf, 1, 10)
        buf1 = "ID" & id & "," & jj & ":" & Mid(dr_buf, 1, 10)
        Call syslog("dr_sub", 0, 0, buf1)
        '下一行有時會當掉
        'rcn1.Execute "update dr set dr_stat='" & buf1 & "' where dr_awno='" & awno1 & "'"
        Call Wait(1): Exit Function
      Else
        For ii = 2 To (jj - 2)
          If Mid(dr_buf, ii, 1) < "0" Or Mid(dr_buf, ii, 1) > "9" Then
            frmComPort.lblComPortStatus(PLC_Port + 17) = "awno" & awno1 & ",ID" & id & ": 第" & ii & "字異常"
            Call Wait(1): Exit Function
          End If
        Next ii
        If True Then
          rcn1.Execute "update dr set dr_use='" & Mid(dr_buf, 1, 1) & "',dr_stat='" & Mid(dr_buf, 2, 9) & "' where dr_awno='" & awno1 & "'"
        ElseIf awno1 = "2" Then
          rcn1.Execute "update dr set dr_use='" & Mid(dr_buf, 1, 1) & "',dr_stat='" & Mid(dr_buf, 2, 7) & "' where dr_awno='" & awno1 & "'"
        ElseIf awno1 = "3" Then
          rcn1.Execute "update dr set dr_use='" & Mid(dr_buf, 1, 1) & "',dr_stat='" & Mid(dr_buf, 2, 6) & "' where dr_awno='" & awno1 & "'"
        ElseIf awno1 = "4" Then
          rcn1.Execute "update dr set dr_use='" & Mid(dr_buf, 1, 1) & "',dr_stat='" & Mid(dr_buf, 2, 5) & "' where dr_awno='" & awno1 & "'"
        ElseIf awno1 = "5" Then
          rcn1.Execute "update dr set dr_use='" & Mid(dr_buf, 1, 1) & "',dr_stat='" & Mid(dr_buf, 2, 9) & "' where dr_awno='" & awno1 & "'"
        End If
        frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & ": OK"
      End If
      Call Wait(1)
'    If Mid(dr_buf, 1, 1) = "1" Then
      'GoTo rtn5
      brs2.Open "select * from dr where dr_awno='" & awno1 & "' and (substring(dr_cvar1,1,2)='01' or substring(dr_cvar1,1,2)='11' or substring(dr_cvar1,1,2)='21')", rcn, adOpenKeyset, adLockOptimistic
      'brs2.Open "select * from cw where cw_awno='" & awno1 & "' and (substring(cw_cvar1,1,2)='01' or substring(cw_cvar1,1,2)='11')", rcn, adOpenKeyset, adLockOptimistic
      If Not brs2.EOF Then
        Call sio_flush(PLC_Port, 2)
        If Left(brs2!dr_cvar1, 2) = "01" Then '門禁系統鎖定
          id = "22"
        ElseIf Left(brs2!dr_cvar1, 2) = "11" Then '門禁系統解鎖
          id = "23"
        ElseIf Left(brs2!dr_cvar1, 2) = "21" Then '人員闖入警報解除
          id = "25"
        End If
'rtn5:
'        id = "23": awno1 = "3": PLC_Port = "16"
        If MSCommAR(PLC_Port).PortOpen = True Then
            sSend = id & awno1 & awno1 & Chr(13) & Chr(10)
            MSCommAR(PLC_Port).Output = sSend
        End If
        lib_buf = "": ii = 0: jj = 0
        Do
          buf = ""
          LL = MSCommAR(iPortIndex).InBufferCount
          buf = MSCommAR(iPortIndex).Input
          If LL > 0 Then
            Mid(lib_buf, jj + 1, LL) = Mid(buf, 1, LL): jj = jj + LL
          Else
            ii = ii + 1
          End If
          If ii > 5 Or jj >= 4 Then Exit Do
          Call Wait(0.1)
        Loop
        'ii=1即可讀到
        frmComPort.lblComPortLen(PLC_Port + 17) = jj: frmComPort.lblComPortString(PLC_Port + 17) = lib_buf
        If jj <> 4 Then
          frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & "," & jj & ":" & Mid(lib_buf, 1, 10)
          buf1 = "ID" & id & "," & jj & ":" & Mid(dr_buf, 1, 10)
          Call syslog("dr_sub1", 0, 0, buf1)
          '下一行有時會當掉
          'rcn1.Execute "update dr set dr_cvar1='" & buf1 & "' where dr_awno='" & awno1 & "'"
          Set brs2 = Nothing: Exit Function
        Else
          If Left(lib_buf, 2) = "11" Then
            If Left(brs2!dr_cvar1, 2) = "01" Then
              frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & ": 門禁已鎖定成功"
              rcn1.Execute "update dr set dr_cvar1='02' where dr_awno='" & awno1 & "'"
              'rcn1.Execute "update cw set cw_cvar1='02' where cw_awno='" & awno1 & "' and cw_palt=" & brs2!cw_palt
            ElseIf Left(brs2!dr_cvar1, 2) = "11" Then
              frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & ": 門禁已解鎖成功"
              rcn1.Execute "update dr set dr_cvar1='12' where dr_awno='" & awno1 & "'"
              'rcn1.Execute "update cw set cw_cvar1='12' where cw_awno='" & awno1 & "' and cw_palt=" & brs2!cw_palt
            ElseIf Left(brs2!dr_cvar1, 2) = "21" Then
              frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & ": 人員闖入警報解除成功"
              rcn1.Execute "update dr set dr_cvar1='22' where dr_awno='" & awno1 & "'"
              'rcn1.Execute "update cw set cw_cvar1='22' where cw_awno='" & awno1 & "' and cw_palt=" & brs2!cw_palt
            End If
            Call Wait(1)
          Else
            '若有門開著會回19
            If Left(lib_buf, 2) = "19" Then
              rcn1.Execute "update dr set dr_cvar1='19' where dr_awno='" & awno1 & "'"
            Else
              buf1 = "ID" & id & "," & jj & ":" & Mid(dr_buf, 1, 10)
              Call syslog("dr_sub2", 0, 0, buf1)
              '下一行有時會當掉
              'rcn1.Execute "update dr set dr_cvar1='" & buf1 & "' where dr_awno='" & awno1 & "'"
            End If
            'If Left(brs2!cw_cvar1, 2) = "01" Then
            '  frmComPort.lblComPortStatus(PLC_Port+17) = "ID" & id & ": 門禁鎖定失敗"
            'ElseIf Left(brs2!cw_cvar1, 2) = "11" Then
            '  frmComPort.lblComPortStatus(PLC_Port+17) = "ID" & id & ": 門禁解鎖失敗"
            'End If
            frmComPort.lblComPortStatus(PLC_Port + 17) = "ID" & id & "," & jj & ":" & Mid(lib_buf, 1, 10)
            Call Wait(1)
            'Set brs2 = Nothing:Exit Function
          End If
        End If
      End If
      Set brs2 = Nothing
'    End If
rtn1:
      dr_sub = True
    Exit Function
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[dr_sub]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[dr_sub]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Function
Private Sub TimerNEW_Timer()
Dim i As Integer
Dim gTcmd As New ADODB.Recordset
Dim lRtn As Integer
Dim sSend As String
Dim sBuf As String
Dim ii As Integer, jj As Integer, LL As Integer, buf As String * 210, lib_buf As String * 210, brs2 As New ADODB.Recordset
    On Error GoTo ErrorHandler
    If iPortIndex = 4 Then Call dr_sub("1", iPortIndex)
    If iPortIndex = 5 Then Call dr_sub("2", iPortIndex)
    If iPortIndex = 6 Then Call dr_sub("3", iPortIndex)
    If iPortIndex = 7 Then Call dr_sub("4", iPortIndex)
    If iPortIndex = 8 Then Call dr_sub("5", iPortIndex)
    'If iPortIndex = 4 Then GoTo rtn15
    
    If iPortIndex <= 8 Then             ' 輸送設備PLC (原紗,經軸,織軸,胚布,成品)
      If MSCommAR(8).PortOpen = True Then   '成品
          sSend = ">5" & Chr(13) & Chr(10)
          MSCommAR(8).Output = sSend
      End If
      If MSCommAR(iPortIndex).PortOpen = True And iPortIndex <> 8 Then    ' 輸送設備PLC (原紗,經軸,織軸,胚布)
          If iPortIndex = 4 Then      ' 原紗
              sSend = ">3" & Chr(13) & Chr(10)
          'ElseIf iPortIndex = 8 Then  ' 成品
          '    sSend = ">5" & Chr(13) & Chr(10)
          ElseIf iPortIndex = 5 Or iPortIndex = 6 Or iPortIndex = 7 Then
              sSend = ">2" & Chr(13) & Chr(10)
          End If
          MSCommAR(iPortIndex).Output = sSend
          If frmComPort.Visible = True Then
              frmComPort.lblComPortLen(iPortIndex + 17) = 0
              frmComPort.lblComPortString(iPortIndex + 17) = ""
              frmComPort.lblComPortStatus(iPortIndex + 17) = ""
          End If
      ElseIf MSCommAR(iPortIndex).PortOpen = False Then
          If frmComPort.Visible = True Then
              frmComPort.lblComPortStatus(iPortIndex + 17).ForeColor = &HFF&
              frmComPort.lblComPortStatus(iPortIndex + 17) = "通訊埠開啟失敗!!"
          End If
      End If
      If True Or MSCommAR(iPortIndex).InBufferCount = PLC_LEN(iPortIndex + 17) Then
        lib_buf = "": ii = 0: jj = 0
        Do
          buf = ""
          LL = MSCommAR(iPortIndex).InBufferCount
          buf = MSCommAR(iPortIndex).Input
          If LL > 0 Then
            Mid(lib_buf, jj + 1, LL) = Mid(buf, 1, LL): jj = jj + LL
          Else
            ii = ii + 1
          End If
          If ii > 5 Or (iPortIndex = 4 And jj >= 28) Or (iPortIndex = 5 And jj >= 20) Or (iPortIndex = 6 And jj >= 14) Or (iPortIndex = 7 And jj >= 16) Or (iPortIndex = 8 And jj >= 50) Then Exit Do
          'If ii > 5 Or (iPortIndex = 4 And jj >= 28) Or (iPortIndex = 5 And jj >= 16) Or (iPortIndex = 6 And jj >= 12) Or (iPortIndex = 7 And jj >= 16) Or (iPortIndex = 8 And jj >= 50) Then Exit Do
          Call Wait(0.2)
        Loop
        'ii=1即可讀到
        'lRtn = MSCommAR(iPortIndex).InBufferCount
        'sBuf = MSCommAR(iPortIndex).Input
        'sBuf = Trim(sBuf)
        '成品PLC每次會送出4組>51110110000090010101000000000000000000000000000CRLF(50bytes),
        '總計200bytes,所以jj=200
        If (iPortIndex = 8 And (jj = 50 Or jj = 200)) Or (iPortIndex = 4 And jj >= 0) Or (iPortIndex = 5 And jj = 20) Or (iPortIndex = 6 And jj = 14) Or (iPortIndex = 7 And jj = 16) Then
        'If (iPortIndex = 8 And (jj = 50 Or jj = 200)) Or (iPortIndex = 4 And jj = 28) Or (iPortIndex = 5 And jj = 20) Or (iPortIndex = 6 And jj = 14) Or (iPortIndex = 7 And jj = 16) Then
        'If (iPortIndex = 8 And (jj = 50 Or jj = 200)) Or (iPortIndex = 4 And jj = 28) Or (iPortIndex = 5 And jj = 16) Or (iPortIndex = 6 And jj = 12) Or (iPortIndex = 7 And jj = 16) Then
          lRtn = PLC_LEN(iPortIndex + 17) 'jj
          sBuf = Left(lib_buf, PLC_LEN(iPortIndex + 17))
          If (iPortIndex = 4 And lRtn = 28) Or Left(sBuf, 2) = ">5" Or Left(sBuf, 2) = ">2" Then
          '原紗倉收到28字都為空白
          'If Left(sBuf, 2) = ">3" Or Left(sBuf, 2) = ">5" Or Left(sBuf, 2) = ">2" Then
            If iPortIndex = 5 And jj = 20 Then
              rcn.Execute "update st set st_load='" & Mid(lib_buf, 15, 1) & "' where substring(st_stno,1,4) = 'B101'"
              rcn.Execute "update st set st_load='" & Mid(lib_buf, 16, 1) & "' where substring(st_stno,1,4) = 'B102'"
              rcn.Execute "update st set st_load='" & Mid(lib_buf, 17, 1) & "' where substring(st_stno,1,4) = 'B103'"
            End If
            If iPortIndex = 6 And jj = 14 Then
              rcn.Execute "update st set st_load='" & Mid(lib_buf, 11, 1) & "' where substring(st_stno,1,4) = 'C102'"
            End If
            Call UpdateST(iPortIndex + 17, sBuf)
            If frmComPort.Visible = True Then
                frmComPort.lblComPortLen(iPortIndex + 17) = lRtn
                frmComPort.lblComPortString(iPortIndex + 17) = sBuf
                frmComPort.lblComPortStatus(iPortIndex + 17).ForeColor = &HFF0000
                frmComPort.lblComPortStatus(iPortIndex + 17) = "資料正確!!"
            End If
          Else
            If frmComPort.Visible = True Then
                frmComPort.lblComPortLen(iPortIndex + 17) = lRtn
                frmComPort.lblComPortString(iPortIndex + 17) = sBuf
                frmComPort.lblComPortStatus(iPortIndex + 17).ForeColor = &HFF&
                frmComPort.lblComPortStatus(iPortIndex + 17) = "資料不正確!!"
            End If
            'Trace "[MSCommAR_OnComm(" & iPortIndex & ")]:資料不正確: Len=" & lRtn & "; Input=" & sBuf
            Sleep (250)
            sBuf = MSCommAR(iPortIndex).Input    '清空 Input Buffer
          End If
        Else
          GoTo rtn1
        End If
      ElseIf MSCommAR(iPortIndex).InBufferCount > PLC_LEN(iPortIndex + 17) Then
rtn1:
          If frmComPort.Visible = True Then
              frmComPort.lblComPortLen(iPortIndex + 17) = lRtn
              frmComPort.lblComPortString(iPortIndex + 17) = sBuf
              frmComPort.lblComPortStatus(iPortIndex + 17).ForeColor = &HFF&
              frmComPort.lblComPortStatus(iPortIndex + 17) = "資料長度不符!!"
          End If
          'Trace "[MSCommAR_OnComm(" & iPortIndex & ")]:資料長度不符: Len=" & lRtn & "; Input=" & sBuf
          Sleep (250)
          sBuf = MSCommAR(iPortIndex).Input        '清空 Input Buffer
      End If
    
    
    ElseIf iPortIndex = 9 Then        '檢查天車異常警示燈
        TimerNEW.Enabled = False
        If gADOCon.State = 1 Then
              sSQLStr = "Select cr_awno,cr_err,cr_mveq,cr_mvst,cr_from,cr_to From cr where cr_awno < '6'"
              gADOCmd.CommandText = sSQLStr
              Set gADORecord = gADOCmd.Execute
              Do While Not gADORecord.EOF
                   CraneErrNo(Val(gADORecord!cr_awno)) = gADORecord!cr_err
                   gADORecord.MoveNext
              Loop
              Set gADORecord = Nothing
              For iCraneCnt = 1 To 4
                    If CraneErrNo(iCraneCnt) <> CraneTemp(iCraneCnt) And MSCommAR(iCraneCnt + 3).PortOpen = True Then 'PLC Port
                          If CraneErrNo(iCraneCnt) = 0 Then
                              sSend = ">500" & Chr(13) & Chr(10)
                              MSCommAR(iCraneCnt + 3).Output = sSend
                              Sleep (500)
                              sBuf = MSCommAR(iCraneCnt + 3).Input
                         ElseIf CraneErrNo(iCraneCnt) = 67 Or CraneErrNo(iCraneCnt) >= 3000 Then    '67=天車連線中斷
                              sSend = ">510" & Chr(13) & Chr(10)
                              MSCommAR(iCraneCnt + 3).Output = sSend
                              Sleep (500)
                              sBuf = MSCommAR(iCraneCnt + 3).Input
                         End If
                         CraneTemp(iCraneCnt) = CraneErrNo(iCraneCnt)
                    End If
               Next

              If CraneErrNo(5) <> CraneTemp(5) And MSCommAR(8).PortOpen = True Then '成品 PLC Port
                   If CraneErrNo(5) = 0 Then
                        sSend = ">300" & Chr(13) & Chr(10)
                        MSCommAR(8).Output = sSend
                        Sleep (500)
                        sBuf = MSCommAR(8).Input
                   ElseIf CraneErrNo(5) = 67 Or CraneErrNo(5) >= 100 Then    '67=天車連線中斷
                        sSend = ">310" & Chr(13) & Chr(10)
                        MSCommAR(8).Output = sSend
                        Sleep (500)
                        sBuf = MSCommAR(8).Input
                   End If
                   CraneTemp(5) = CraneErrNo(5)
              End If
        End If
        TimerNEW.Enabled = True
    End If
rtn15:
    iPortIndex = iPortIndex + 1
    If iPortIndex > 9 Then iPortIndex = 4
    
    Exit Sub
ErrorHandler:
    If TimerNEW.Enabled = False Then TimerNEW.Enabled = True
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[TimerPLC_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[TimerPLC_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub TimerNEW_Timer_old()
Dim i As Integer
Dim gTcmd As New ADODB.Recordset
Dim lRtn As Integer
Dim sSend As String
Dim sBuf As String

    On Error GoTo ErrorHandler
    
    If iPortIndex <= 8 Then             ' 輸送設備PLC (原紗,經軸,織軸,胚布,成品)
        If MSCommAR(8).PortOpen = True Then   '成品
            sSend = ">5" & Chr(13) & Chr(10)
            MSCommAR(8).Output = sSend
        End If
        If MSCommAR(iPortIndex).PortOpen = True And iPortIndex <> 8 Then    ' 輸送設備PLC (原紗,經軸,織軸,胚布)
            If iPortIndex = 4 Then      ' 原紗
                sSend = ">3" & Chr(13) & Chr(10)
            'ElseIf iPortIndex = 8 Then  ' 成品
            '    sSend = ">5" & Chr(13) & Chr(10)
            ElseIf iPortIndex = 5 Or iPortIndex = 6 Or iPortIndex = 7 Then
                sSend = ">2" & Chr(13) & Chr(10)
            End If
            MSCommAR(iPortIndex).Output = sSend
            If frmComPort.Visible = True Then
                frmComPort.lblComPortLen(iPortIndex + 17) = 0
                frmComPort.lblComPortString(iPortIndex + 17) = ""
                frmComPort.lblComPortStatus(iPortIndex + 17) = ""
            End If
        ElseIf MSCommAR(iPortIndex).PortOpen = False Then
            If frmComPort.Visible = True Then
                frmComPort.lblComPortStatus(iPortIndex + 17).ForeColor = &HFF&
                frmComPort.lblComPortStatus(iPortIndex + 17) = "通訊埠開啟失敗!!"
            End If
        End If
    ElseIf iPortIndex = 9 Then        '檢查天車異常警示燈
        TimerNEW.Enabled = False
        If gADOCon.State = 1 Then
              sSQLStr = "Select cr_awno,cr_err,cr_mveq,cr_mvst,cr_from,cr_to From cr where cr_awno < '6'"
              gADOCmd.CommandText = sSQLStr
              Set gADORecord = gADOCmd.Execute
              Do While Not gADORecord.EOF
                   CraneErrNo(Val(gADORecord!cr_awno)) = gADORecord!cr_err
                   gADORecord.MoveNext
              Loop
              Set gADORecord = Nothing
              For iCraneCnt = 1 To 4
                    If CraneErrNo(iCraneCnt) <> CraneTemp(iCraneCnt) And MSCommAR(iCraneCnt + 3).PortOpen = True Then 'PLC Port
                          If CraneErrNo(iCraneCnt) = 0 Then
                              sSend = ">500" & Chr(13) & Chr(10)
                              MSCommAR(iCraneCnt + 3).Output = sSend
                              Sleep (500)
                              sBuf = MSCommAR(iCraneCnt + 3).Input
                         ElseIf CraneErrNo(iCraneCnt) = 67 Or CraneErrNo(iCraneCnt) >= 3000 Then    '67=天車連線中斷
                              sSend = ">510" & Chr(13) & Chr(10)
                              MSCommAR(iCraneCnt + 3).Output = sSend
                              Sleep (500)
                              sBuf = MSCommAR(iCraneCnt + 3).Input
                         End If
                         CraneTemp(iCraneCnt) = CraneErrNo(iCraneCnt)
                    End If
               Next

              If CraneErrNo(5) <> CraneTemp(5) And MSCommAR(8).PortOpen = True Then '成品 PLC Port
                   If CraneErrNo(5) = 0 Then
                        sSend = ">300" & Chr(13) & Chr(10)
                        MSCommAR(8).Output = sSend
                        Sleep (500)
                        sBuf = MSCommAR(8).Input
                   ElseIf CraneErrNo(5) = 67 Or CraneErrNo(5) >= 100 Then    '67=天車連線中斷
                        sSend = ">310" & Chr(13) & Chr(10)
                        MSCommAR(8).Output = sSend
                        Sleep (500)
                        sBuf = MSCommAR(8).Input
                   End If
                   CraneTemp(5) = CraneErrNo(5)
              End If
        End If
        TimerNEW.Enabled = True
    End If
    
    iPortIndex = iPortIndex + 1
    If iPortIndex > 9 Then iPortIndex = 4
    
    Exit Sub
ErrorHandler:
    If TimerNEW.Enabled = False Then TimerNEW.Enabled = True
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[TimerPLC_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[TimerPLC_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

'Private Sub MSCommAR_OnComm(Index As Integer)
'Dim lRtn As Integer
'Dim sBuf As String
'Dim sTemp As String
'On Error GoTo ErrorHandler
'
'    If MSCommAR(Index).InBufferCount = PLC_LEN(Index + 17) Then
'        lRtn = MSCommAR(Index).InBufferCount
'        sBuf = MSCommAR(Index).Input
'        'sBuf = Trim(sBuf)
'        If Left(sBuf, 2) = ">3" Or Left(sBuf, 2) = ">5" Or Left(sBuf, 2) = ">2" Then
'            Call UpdateST(Index + 17, sBuf)
'            If frmComPort.Visible = True Then
'                frmComPort.lblComPortLen(Index + 17) = lRtn
'                frmComPort.lblComPortString(Index + 17) = sBuf
'                frmComPort.lblComPortStatus(Index + 17).ForeColor = &HFF0000
'                frmComPort.lblComPortStatus(Index + 17) = "資料正確!!"
'            End If
'        Else
'            If frmComPort.Visible = True Then
'                frmComPort.lblComPortLen(Index + 17) = lRtn
'                frmComPort.lblComPortString(Index + 17) = sBuf
'                frmComPort.lblComPortStatus(Index + 17).ForeColor = &HFF&
'                frmComPort.lblComPortStatus(Index + 17) = "資料不正確!!"
'            End If
'            Trace "[MSCommAR_OnComm(" & Index & ")]:資料不正確: Len=" & lRtn & "; Input=" & sBuf
'            Sleep (250)
'            sBuf = MSCommAR(Index).Input    '清空 Input Buffer
'        End If
'    ElseIf MSCommAR(Index).InBufferCount > PLC_LEN(Index + 17) Then
'        If frmComPort.Visible = True Then
'            frmComPort.lblComPortLen(Index + 17) = lRtn
'            frmComPort.lblComPortString(Index + 17) = sBuf
'            frmComPort.lblComPortStatus(Index + 17).ForeColor = &HFF&
'            frmComPort.lblComPortStatus(Index + 17) = "資料長度不符!!"
'        End If
'        Trace "[MSCommAR_OnComm(" & Index & ")]:資料長度不符: Len=" & lRtn & "; Input=" & sBuf
'        Sleep (250)
'        sBuf = MSCommAR(Index).Input        '清空 Input Buffer
'    End If
'    Exit Sub
'ErrorHandler:
'    lErrorNo = Err.Number
'    sErrorMsg = Err.Description
'    Trace "[MSCommAR_OnComm]:Error:" & lErrorNo & "; Error Msg : " & sErrorMsg
'    xMsgBox "[InitialParameter]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
'End Sub


Private Sub TimerAGV_Timer()
Dim iResult As Integer
Dim gTcmd As New ADODB.Recordset
Dim lRtn As Long
Dim sSend As String
Dim sTemp As String
Dim sInBuf As String
Dim i As Integer
Dim bA001Flag As Boolean
Dim bA002Flag As Boolean
Dim sNxt As String
Dim bDoNext As Boolean

   ' On Error GoTo ErrorHandler

        If giAGVOpenOK = 1 And MSCommAGV.InBufferCount > 0 Then      'AGV有資料回傳時優先處理
            sInBuf = MSCommAGV.Input
            sTemp = Right(sInBuf, 1)
            If sTemp = Chr(STX) Then '最後一碼=STX
                sSend = Chr(DLE)
                MSCommAGV.Output = sSend
                Call Sleep(300)
                If MSCommAGV.InBufferCount > 0 Then
                    iDataLen = MSCommAGV.InBufferCount
                    AGV_InBuf = MSCommAGV.Input
                    Call AGV_To_HOST
                Else                    'HOST回答DLE後,AGV未傳回資料
                    If frmComPort.Visible = True Then
                        frmComPort.lblComPortLen(AGV_Port) = 0
                        frmComPort.lblComPortString(AGV_Port) = ""
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV未傳回狀態資料!!"
                    End If
                    Exit Sub
                End If
            End If

'                If frmComPort.Visible = True Then
'                    lRtn = sio_iqueue(AGV_Port)
'                    lRtn = sio_read(AGV_Port, AGV_InBuf, lRtn)
'                    frmComPort.lblComPortLen(AGV_Port) = lRtn + 1
'                    frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
'                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
'                    frmComPort.lblComPortStatus(AGV_Port) = "AGV傳回不明資料,清除後重新讀取!!"
'                End If
'                lRtn = sio_flush(AGV_Port, 2)      '清除 input & output buffer
'                Exit Sub

        ElseIf giAGVOpenOK = 1 And MSCommAGV.InBufferCount = 0 Then           'AGV無資料回傳時判斷是否須下命令至AGV
        
            If gADOCon.State = 1 Then
'               For i = 1 To 24
'                    LineNo(i) = False
'               Next
'               AGV_CommandCnt = 0
'               sSQLStr = "Select cw_now,cw_nxt From cw Where cw_awno='1' and cw_mveq like 'AGV%' and cw_stat='2'"
'               gADOCmd.CommandText = sSQLStr
'               Set gADORecord = gADOCmd.Execute
'               Do While Not gADORecord.EOF
'                    If Trim(gADORecord!cw_nxt) = "A001" Then
'                         bA001Flag = True
'                    ElseIf Trim(gADORecord!cw_nxt) = "A002" Then
'                         bA002Flag = True
'                    End If
'
'                    If Left(gADORecord!cw_now, 2) = "03" Then
'                         LineNo(Val(Mid(gADORecord!cw_now, 4, 2))) = True
'                    ElseIf Left(gADORecord!cw_nxt, 2) = "03" Then
'                         LineNo(Val(Mid(gADORecord!cw_nxt, 4, 2))) = True
'                    End If
'                    AGV_CommandCnt = AGV_CommandCnt + 1
'                    gADORecord.MoveNext
'               Loop
'               Set gADORecord = Nothing
'
'               sSQLStr = "Select st_stno,st_load From st Where st_awno='1' and (st_stno='A001' or st_stno='A002')"
'               gADOCmd.CommandText = sSQLStr
'               Set gADORecord = gADOCmd.Execute
'               Do While Not gADORecord.EOF
'                    If Trim(gADORecord!st_stno) = "A001" And gADORecord!st_load = "1" Then
'                         bA001Flag = True
'                    ElseIf Trim(gADORecord!st_stno) = "A002" And gADORecord!st_load = "1" Then
'                         bA002Flag = True
'                    End If
'                    gADORecord.MoveNext
'               Loop
'               Set gADORecord = Nothing
'
'
'               If AGV_CommandCnt < 4 Then
'                    sSQLStr = "Select cw_palt,cw_to,cw_now,cw_nxt From View_cw_Awno1_AGV order by cw_sitm"
'                    gADOCmd.CommandText = sSQLStr
'                    Set gADORecord = gADOCmd.Execute
'                    Do While ((Not gADORecord.EOF) And AGV_CommandCnt < 4)
'                         If Left(gADORecord!cw_now, 2) = "03" And LineNo(Val(Mid(gADORecord!cw_now, 4, 2))) = True Then
'                              bDoNext = True
'                         ElseIf Left(gADORecord!cw_nxt, 2) = "03" And LineNo(Val(Mid(gADORecord!cw_nxt, 4, 2))) = True Then
'                              bDoNext = True
'                         Else
'                              bDoNext = False
'                         End If
'
'                         If (Trim(gADORecord!cw_to) = "A006A008") And (bA001Flag = True And bA002Flag = True) Then   '搬運至樓上且A001 & A002 皆有載或已派工
'                              bDoNext = True
'                         End If
'
'                         If bDoNext = False Then
'                              If Trim(gADORecord!cw_to) = "A006A008" And bA001Flag = False Then
'                                   ChangeTo = "A006"
'                                   sNxt = "A001"
'                              ElseIf Trim(gADORecord!cw_to) = "A006A008" And bA002Flag = False Then
'                                   ChangeTo = "A008"
'                                   sNxt = "A002"
'                              Else
'                                   ChangeTo = ""
'                                   sNxt = ""
'                              End If
'                              Call HOST_To_AGV("T")
'                              Call Sleep(300)
'                              sInBuf = sio_getch(AGV_Port)
'                              If sInBuf = DLE Then
'                                   If Trim(gADORecord!cw_nxt) = "A001" Then bA001Flag = True
'                                   If Trim(gADORecord!cw_nxt) = "A002" Then bA002Flag = True
'                                   If ChangeTo = "A006" Then bA001Flag = True
'                                   If ChangeTo = "A008" Then bA002Flag = True
'                                   If ChangeTo = "" Then
'                                        sSQLStr = "Update cw Set cw_stat='2' Where cw_palt=" & gADORecord!cw_palt & " and cw_awno='1'"
'                                   Else
'                                        sSQLStr = "Update cw Set cw_stat='2',cw_to='" & ChangeTo & "',cw_nxt='" & sNxt & "' Where cw_palt=" & gADORecord!cw_palt & " and cw_awno='1'"
'                                   End If
'                                   gADOCmd.CommandText = sSQLStr
'                                   Set gTcmd = gADOCmd.Execute
'                                   AGV_CommandCnt = AGV_CommandCnt + 1
'                              End If
'                          End If
'                          gADORecord.MoveNext
'                      Loop
'                      Set gADORecord = Nothing
'                 End If
'            Else
       '         Call DataBaseConnect
            End If
        End If
    
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[TimerAGV_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[TimerAGV_Timer]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub


Private Sub UpdateST(ST_Index As Integer, StrBuf As String)
Dim i As Integer
Dim sTemp As String
Dim sLono As String
Dim iPalt As Integer
Dim sSitm As String
    ' On Error GoTo ErrorHandler
    Select Case ST_Index
        Case 21  '原紗
            For i = 1 To 3
                sTemp = Mid(StrBuf, ST21_StByte(i), 1)
                If ST21_Temp(i) <> sTemp And gADOCon.State = 1 Then
                    sSQLStr = "Update st Set st_load='" & sTemp & "' where st_awno='1' and st_stno='" & ST21_StNo(i) & "'"
                    gADOCmd.CommandText = sSQLStr
                    Set gADORecord = gADOCmd.Execute
                    ST21_Temp(i) = sTemp
                End If
             Next
        Case 22  '經軸
            For i = 1 To 7
                sTemp = Mid(StrBuf, ST22_StByte(i), 1)
                If ST22_Temp(i) <> sTemp And gADOCon.State = 1 Then
                    sSQLStr = "Update st Set st_load='" & sTemp & "' where st_awno='2' and st_stno='" & ST22_StNo(i) & "'"
                    gADOCmd.CommandText = sSQLStr
                    Set gADORecord = gADOCmd.Execute
                    ST22_Temp(i) = sTemp
                End If
             Next
        Case 23  '織軸
             For i = 1 To 6
                sTemp = Mid(StrBuf, ST23_StByte(i), 1)
                If ST23_Temp(i) <> sTemp And gADOCon.State = 1 Then
                    sSQLStr = "Update st set st_load='" & sTemp & "',st_err=0 where st_awno='3' and st_stno='" & ST23_StNo(i) & "'"
                    gADOCmd.CommandText = sSQLStr
                    Set gADORecord = gADOCmd.Execute
                    ST23_Temp(i) = sTemp
                End If
             Next
        Case 24  '胚布
             For i = 1 To 5
                sTemp = Mid(StrBuf, ST24_StByte(i), 1)
                If ST24_Temp(i) <> sTemp And gADOCon.State = 1 Then
                    sSQLStr = "Update st Set st_load='" & sTemp & "' where st_awno='4' and st_stno='" & ST24_StNo(i) & "'"
                    gADOCmd.CommandText = sSQLStr
                    Set gADORecord = gADOCmd.Execute
                    ST24_Temp(i) = sTemp
                End If
             Next
        Case 25  '成品
             For i = 1 To 10
                sTemp = Mid(StrBuf, ST25_StByte(i), 1)
                If gADOCon.State = 1 Then
                'E003及E006兩站有時無法反應最新狀態
                'If ST25_Temp(i) <> sTemp And gADOCon.State = 1 Then
                    sSQLStr = "Update st Set st_load='" & sTemp & "' where st_awno='5' and st_stno='" & ST25_StNo(i) & "'"
                    gADOCmd.CommandText = sSQLStr
                    Set gADORecord = gADOCmd.Execute
                    ST25_Temp(i) = sTemp
                End If
             Next
    End Select
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[UpdateST]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[UpdateST]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub AGV_To_HOST()
Dim sTemp As String
Dim lRtn As Long
Dim bOK As Boolean
Dim sShowStr As String
Dim sAGVNo As String
Dim i As Integer
Dim sErrStr As String

  '  On Error GoTo ErrorHandler
        bOK = True
        If CHK_BCC = False Then
            If frmComPort.Visible = True Then
                frmComPort.lblComPortLen(AGV_Port) = iDataLen
                frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                frmComPort.lblComPortStatus(AGV_Port) = "AGV回傳資料BCC檢查碼錯誤!!"
             End If
             sErrStr = MSCommAGV.Input
             BCC_Err_Cnt = BCC_Err_Cnt + 1
             'Trace "[AGV BCC碼錯誤(第" & BCC_Err_Cnt & "次) : " & iDataLen & " Bytes : " & AGV_InBuf
             If BCC_Err_Cnt > 10 Then
                    For i = 1 To iDataLen
                         sTemp = Mid(AGV_InBuf, i, 1)
                         If sTemp <> "" Then
                            sErrStr = sErrStr & " " & Format(Asc(sTemp))
                         Else
                            sErrStr = sErrStr & " " & "00"
                         End If
                    Next
                    
                    'Trace "AGV Error Str:" & sErrStr
                    MSCommAGV.PortOpen = False
                    Sleep (1000)
                    MSCommAGV.PortOpen = True
                    If MSCommAGV.PortOpen = True Then
                         giAGVOpenOK = 1
                    Else
                         giAGVOpenOK = -1
                         'Trace "COM " & AGV_Port & " REOPEN  &  PROTOCOL RESETTING OK"
                    End If
                    BCC_Err_Cnt = 0
             Else
                Exit Sub
             End If
        Else
             BCC_Err_Cnt = 0
        End If
        
        sTemp = Left(AGV_InBuf, 3)
        Select Case sTemp
            Case HOST + AGV + "k"
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortLen(AGV_Port) = iDataLen
                    frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                    frmComPort.lblComPortStatus(AGV_Port) = "k=AGV啟動!!"
                End If
                'Trace "AGV(k): " & iDataLen & " Bytes : " & AGV_InBuf
                'Call HOST_To_AGV("N")
            Case HOST + AGV + "e"
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortLen(AGV_Port) = iDataLen
                    frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                    If Mid(AGV_InBuf, 5, 2) = "01" Or Mid(AGV_InBuf, 5, 2) = "02" Or Mid(AGV_InBuf, 5, 2) = "03" Then
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                        frmComPort.lblComPortStatus(AGV_Port) = Mid(AGV_InBuf, 5, 2) & " AGV 至充電站!!"
                    ElseIf Mid(AGV_InBuf, 5, 2) = "11" Then
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV控制箱電源異常!!"
                    ElseIf Mid(AGV_InBuf, 5, 2) = "12" Then
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV導引頻率訊號中斷!!"
                    Else
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV ERROR STATUS訊息不明!!"
                    End If
                End If
                'Trace "AGV(e): " & iDataLen & " Bytes : " & AGV_InBuf
            Case HOST + AGV + "a"
                If Mid(AGV_InBuf, 5, 1) = "1" Then
                    If frmComPort.Visible = True Then
                        frmComPort.lblComPortLen(AGV_Port) = iDataLen
                        frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV準備就緒!!"
                    End If
                ElseIf Mid(AGV_InBuf, 5, 1) = "2" Then
                    If frmComPort.Visible = True Then
                        frmComPort.lblComPortLen(AGV_Port) = iDataLen
                        frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV異常!!"
                    End If
                ElseIf Mid(AGV_InBuf, 5, 1) = "3" Then
                    If frmComPort.Visible = True Then
                        frmComPort.lblComPortLen(AGV_Port) = iDataLen
                        frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV處於測試模式!!"
                    End If
                Else
                    If frmComPort.Visible = True Then
                        frmComPort.lblComPortLen(AGV_Port) = iDataLen
                        frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                        frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                        frmComPort.lblComPortStatus(AGV_Port) = "AGV狀況不明!!"
                    End If
                    bOK = False
                End If
                'Trace "AGV(a): " & iDataLen & " Bytes : " & AGV_InBuf
            Case HOST + AGV + "f"
                sAGVNo = Mid(AGV_InBuf, 5, 1)
                If Mid(AGV_InBuf, 4, 1) = "0" Then
                    sShowStr = "AGV" & sAGVNo & " : " & Mid(AGV_InBuf, 11, 5) & " --> " & Mid(AGV_InBuf, 16, 5) & " 作業完成 序號:" & Mid(AGV_InBuf, 6, 5)
                    If gADOCon.State = 1 Then
                        sSQLStr = "Update cw set cw_stat='0' Where cw_awno='1' and cw_palt=" & Val(Mid(AGV_InBuf, 6, 5))
                        gADOCmd.CommandText = sSQLStr
                        Set gADORecord = gADOCmd.Execute
                    End If
                ElseIf Mid(AGV_InBuf, 4, 1) = "1" Then     '決定AGV並更新SQL
                    If gADOCon.State = 1 Then
                        sSQLStr = "Update cw set cw_mveq='AGV" & Mid(AGV_InBuf, 5, 1) & "' Where cw_awno='1' and cw_mveq='AGV' And cw_palt=" & Val(Mid(AGV_InBuf, 6, 5))
                        gADOCmd.CommandText = sSQLStr
                        Set gADORecord = gADOCmd.Execute
                    End If
                    sShowStr = "AGV" & sAGVNo & " : " & Mid(AGV_InBuf, 11, 5) & " --> " & Mid(AGV_InBuf, 16, 5) & " 作業中 序號:" & Mid(AGV_InBuf, 6, 5)
                ElseIf Mid(AGV_InBuf, 4, 1) = "2" Or Mid(AGV_InBuf, 4, 1) = "4" Or Mid(AGV_InBuf, 4, 1) = "5" Then     'AGV PC強迫取消命令
                    If gADOCon.State = 1 Then
                        sSQLStr = "Update cw set cw_err=101 Where cw_awno='1' and cw_mveq like 'AGV%' and cw_palt=" & Val(Mid(AGV_InBuf, 6, 5))
                        gADOCmd.CommandText = sSQLStr
                        Set gADORecord = gADOCmd.Execute
                    End If
                    sShowStr = "AGV PC 取消 : " & Mid(AGV_InBuf, 11, 5) & " --> " & Mid(AGV_InBuf, 16, 5) & " 工作序號:" & Mid(AGV_InBuf, 6, 5)
                Else
                    'Trace "AGV f=???: " & iDataLen & " Bytes : " & AGV_InBuf
                End If

                If frmComPort.Visible = True And sAGVNo <> "0" Then
                    frmComPort.lblAGVStatus(sAGVNo) = sShowStr
                End If
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortLen(AGV_Port).Caption = iDataLen
                    frmComPort.lblComPortString(AGV_Port).Caption = AGV_InBuf
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                    frmComPort.lblComPortStatus(AGV_Port).Caption = "AGV 通訊正常!"
                    If sAGVNo = "0" Then
                         frmComPort.lblComPortStatus(AGV_Port).Caption = sShowStr
                    End If
                End If
            Case HOST + AGV + "t"
                If Mid(AGV_InBuf, 4, 1) = "9" Then      'End of transport job report
                    'Trace "AGV(t): " & iDataLen & " Bytes : " & AGV_InBuf
                    'Call HOST_To_AGV("N")
                End If
            Case Else
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortLen(AGV_Port) = iDataLen
                    frmComPort.lblComPortString(AGV_Port) = AGV_InBuf
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF&
                    frmComPort.lblComPortStatus(AGV_Port) = "AGV狀況不明!!"
                End If
                bOK = False
                'Trace "AGV ???: " & iDataLen & " Bytes : " & AGV_InBuf
        End Select
        If bOK = True Then
            MSCommAGV.Output = Chr(DLE)
        Else
            'lRtn = sio_putch(AGV_Port, NAK)
        End If
        Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[AGV_To_HOST]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[AGV_To_HOST]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub HOST_To_AGV(Str As String)
Dim lRtn As Long
Dim sStNoNow As String
Dim sStNoNext As String
Dim sStNoEnd As String
Dim sInBuf As Long
Dim i As Integer
    On Error GoTo ErrorHandler
    
        If Str = "T" Or Str = "K" Then
            lRtn = sio_flush(AGV_Port, 2)
            lRtn = sio_putch(AGV_Port, STX)
            Sleep (300)
            sInBuf = sio_getch(AGV_Port)
            If sInBuf <> DLE Then
                    Exit Sub
            End If
        End If
        Select Case Str
            Case "N"
                AGV_TempStr = AGV + HOST + "N" + "0" + DUMMY
                Call PUT_DLE_ETX_BCC
                i = Len(AGV_TempStr)
                lRtn = sio_write(AGV_Port, AGV_TempStr, i)
                lRtn = sio_putch(AGV_Port, DLE)
                lRtn = sio_putch(AGV_Port, ETX)
                lRtn = sio_putch(AGV_Port, BCC)
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortOutLen(AGV_Port) = i + 3
                    frmComPort.lblComPortString(AGV_Port) = AGV_TempStr
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                    frmComPort.lblComPortStatus(AGV_Port) = "傳送 N 命令至AGV"
                End If
            Case "K"
                AGV_TempStr = AGV + HOST + "K" + DUMMY + READY
                Call PUT_DLE_ETX_BCC
                i = Len(AGV_TempStr)
                lRtn = sio_write(AGV_Port, AGV_TempStr, i)
                lRtn = sio_putch(AGV_Port, DLE)
                lRtn = sio_putch(AGV_Port, ETX)
                lRtn = sio_putch(AGV_Port, BCC)
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortOutLen(AGV_Port) = i + 3
                    frmComPort.lblComPortString(AGV_Port) = AGV_TempStr
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                    frmComPort.lblComPortStatus(AGV_Port) = "傳送 K 命令至AGV"
                End If
            Case "T"
                If Left(gADORecord!cw_now, 2) = "03" Then
                    sStNoNow = "B1" & Mid(gADORecord!cw_now, 4, 2) & Mid(gADORecord!cw_now, 8, 1)
                Else
                    Select Case Trim(gADORecord!cw_now)
                    Case "A001"
                        sStNoNow = "X1   "
                    Case "A002"
                        sStNoNow = "X2   "
                    Case "Y001"
                        sStNoNow = "Y1   "
                    Case "Y002"
                        sStNoNow = "Y2   "
                    Case "F002"
                        sStNoNow = "F2   "
                    Case "Z001"
                        sStNoNow = "Z1   "
                    Case "Z002"
                        sStNoNow = "Z2   "
                    Case Else
                        sStNoNow = gADORecord!cw_now
                    End Select
                End If
                
                If ChangeTo = "A006" Then
                    sStNoNext = "X1   "
                ElseIf ChangeTo = "A008" Then
                    sStNoNext = "X2   "
                ElseIf Left(gADORecord!cw_nxt, 2) = "03" Then
                    sStNoNext = "B1" & Mid(gADORecord!cw_nxt, 4, 2) & Mid(gADORecord!cw_nxt, 8, 1)
                Else
                    Select Case Trim(gADORecord!cw_nxt)
                        Case "A001"
                            sStNoNext = "X1   "
                        Case "A002"
                            sStNoNext = "X2   "
                        Case "Y001"
                            sStNoNext = "Y1   "
                        Case "Y002"
                            sStNoNext = "Y2   "
                        Case "F002"
                            sStNoNext = "F2   "
                        Case "Z001"
                            sStNoNext = "Z1   "
                        Case "Z002"
                            sStNoNext = "Z2   "
                        Case Else
                            sStNoNext = gADORecord!cw_nxt
                    End Select
                End If
                AGV_TempStr = AGV + HOST + "T" + DUMMY + "0" + Format(gADORecord!cw_palt, "00000") + sStNoNow + sStNoNext + "1"
                Call PUT_DLE_ETX_BCC
                i = Len(AGV_TempStr)
                lRtn = sio_write(AGV_Port, AGV_TempStr, i)
                lRtn = sio_putch(AGV_Port, DLE)
                lRtn = sio_putch(AGV_Port, ETX)
                lRtn = sio_putch(AGV_Port, BCC)
                If frmComPort.Visible = True Then
                    frmComPort.lblComPortOutLen(AGV_Port) = i + 3
                    frmComPort.lblComPortString(AGV_Port) = AGV_TempStr
                    frmComPort.lblComPortStatus(AGV_Port).ForeColor = &HFF0000
                    frmComPort.lblComPortStatus(AGV_Port) = "傳送命令至AGV 序號:" & Format(gADORecord!cw_palt, "00000")
                End If
            Case Else
        End Select
        Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[HOST_To_AGV]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[HOST_To_AGV]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Function CHK_BCC() As Boolean
Dim sAGVTemp() As Byte
Dim lLen As Long
Dim BCC As Byte
Dim i As Integer
Dim sAgvStr As String
    On Error GoTo ErrorHandler
        sAgvStr = Mid(AGV_InBuf, 1, iDataLen)
        sAGVTemp = StrConv(sAgvStr, vbFromUnicode)
        lLen = Len(sAgvStr)
        BCC = 0
        For i = 0 To lLen - 2
            BCC = BCC Xor sAGVTemp(i)
        Next
        If BCC = sAGVTemp(lLen - 1) Then
            CHK_BCC = True
        Else
            CHK_BCC = False
        End If
    Exit Function
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[CHK_BCC]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[CHK_BCC]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Function

Private Sub DataBaseConnect()
On Error GoTo ErrorHandler
    ' SQL SERVER CONNECTION
    gsSQLUserID = IniRead("DATABASE", "UserID", App.Path & "\IniFile\GF3Com.ini")
    gsSQLPassword = IniRead("DATABASE", "Password", App.Path & "\IniFile\GF3Com.ini")
    gsSQLTable = IniRead("DATABASE", "Table", App.Path & "\IniFile\GF3Com.ini")
    gsSQLDSN = IniRead("DATABASE", "DSN", App.Path & "\IniFile\GF3Com.ini")
    gsSQLDbPvd = "SQLOLEDB.1"
    gsSQLCon = "Provider=" & gsSQLDbPvd & ";" & _
          "Password=" & gsSQLPassword & ";" & _
          "Persist Security Info=True;" & _
          "Initial Catalog=" & gsSQLTable & ";" & _
          "Data Source=" & gsSQLDSN
    Set gADOCon = New ADODB.Connection
    Set gADOCmd = New ADODB.Command
    gADOCon.Open gsSQLCon, gsSQLUserID, gsSQLPassword
    Set gADOCmd.ActiveConnection = gADOCon
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    'Trace "[DataBaseConnect]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "[DataBaseConnect]:Error No:" & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

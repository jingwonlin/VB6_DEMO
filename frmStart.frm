VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Start"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.CommandButton Command6 
      Caption         =   "成品"
      Height          =   645
      Left            =   840
      TabIndex        =   5
      Top             =   4095
      Width           =   1485
   End
   Begin VB.CommandButton Command5 
      Caption         =   "胚布"
      Height          =   645
      Left            =   840
      TabIndex        =   4
      Top             =   3150
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "織軸"
      Height          =   645
      Left            =   840
      TabIndex        =   3
      Top             =   2205
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   645
      Left            =   4830
      TabIndex        =   2
      Top             =   7035
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "經軸(二)"
      Height          =   630
      Left            =   840
      TabIndex        =   1
      Top             =   1365
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "經軸(一)"
      Height          =   630
      Left            =   840
      TabIndex        =   0
      Top             =   525
      Width           =   1470
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSQLCon As String
Dim sSQLDSN As String
Dim sSQLUserID As String
Dim sSQLPassword As String
Dim sSQLTable As String
Dim sSQLDbPvd As String

Dim oADOCon As ADODB.Connection
Dim oADOCmd As ADODB.Command
Dim oADORecord As New ADODB.Recordset

Dim sSQLStr As String
Dim lErrorNo As Long
Dim sErrorMsg As String

Dim sAwno As String
Dim sZone As String
Dim sLono As String
Dim sSgst As String
Dim sFbst As String
Dim iSisq As Integer
Dim sCran As String
Dim sSitm As String
Dim sDate As String
Dim sProg As String

Private Sub Command1_Click()
Dim sTemp As String
    On Error GoTo ErrorHandler
    sAwno = "2"
    sZone = "1"
    sSitm = Format(Now, "yyyy/mm/dd hh:mm")
    sDate = " "
    sProg = " "
    
    sSQLStr = "Select awno,lono,losz,zone,sgst,fbst,sisq From Lo Where Awno='1'"
    oADOCmd.CommandText = sSQLStr
    Set oADORecord = oADOCmd.Execute
    Do While Not oADORecord.EOF
        If (Not IsNull(oADORecord!lono)) And Len(oADORecord!lono) = 6 Then
            sLono = Left(oADORecord!lono, 2) & "0" & Mid(oADORecord!lono, 3, 2) & "0" & Right(oADORecord!lono, 2)
            
            If oADORecord!sgst = "0" Or oADORecord!sgst = "F" Or oADORecord!sgst = "S" Or oADORecord!sgst = "R" Then
                sSgst = oADORecord!sgst
            Else
                sSgst = "E"
            End If
            
            If oADORecord!fbst = "0" Then
                sFbst = "0"
            Else
                sFbst = "X"
            End If
            
            iSisq = oADORecord!sisq
            
            sTemp = Left(oADORecord!lono, 2)
            If sTemp = "01" Or sTemp = "02" Then
                sCran = "1"
            ElseIf sTemp = "03" Or sTemp = "04" Then
                sCran = "2"
            ElseIf sTemp = "05" Or sTemp = "06" Then
                sCran = "3"
            End If

            sSQLStr = "Insert Lo (lo_awno,lo_zone,lo_lono,lo_sgst,lo_fbst,lo_sisq,lo_cran,lo_sitm,lo_date,lo_prog) Values ('" & _
            sAwno & "','" & sZone & "','" & sLono & "','" & sSgst & "','" & sFbst & "'," & iSisq & ",'" & sCran & "','" & _
            sSitm & "','" & sDate & "','" & sProg & "')"
            gADOCmd.CommandText = sSQLStr
            Set gADORecord = gADOCmd.Execute
            
            oADORecord.MoveNext
        End If
    Loop
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!經軸(一)] : Error > " & "Error No : " & lErrorNo & "; Error Msg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error Msg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub Command2_Click()
Dim sTemp As String
     
    On Error GoTo ErrorHandler
    sAwno = "6"
    sZone = "1"
    sSitm = Format(Now, "yyyy/mm/dd hh:mm")
    sDate = " "
    sProg = " "
    
    sSQLStr = "Select awno,lono,losz,zone,sgst,fbst,sisq From Lo Where Awno='4'"
    oADOCmd.CommandText = sSQLStr
    Set oADORecord = oADOCmd.Execute
    Do While Not oADORecord.EOF
        If (Not IsNull(oADORecord!lono)) And Len(oADORecord!lono) = 6 Then
            sLono = Left(oADORecord!lono, 2) & "0" & Mid(oADORecord!lono, 3, 2) & "0" & Right(oADORecord!lono, 2)
            
            If oADORecord!sgst = "0" Or oADORecord!sgst = "F" Or oADORecord!sgst = "S" Or oADORecord!sgst = "R" Then
                sSgst = oADORecord!sgst
            Else
                sSgst = "E"
            End If
            
            If oADORecord!fbst = "0" Then
                sFbst = "0"
            Else
                sFbst = "X"
            End If
            
            iSisq = oADORecord!sisq
            
            sTemp = Left(oADORecord!lono, 2)
            If sTemp = "01" Or sTemp = "02" Then
                sCran = "1"
            ElseIf sTemp = "03" Or sTemp = "04" Then
                sCran = "2"
            ElseIf sTemp = "05" Or sTemp = "06" Then
                sCran = "3"
            End If

            sSQLStr = "Insert Lo (lo_awno,lo_zone,lo_lono,lo_sgst,lo_fbst,lo_sisq,lo_cran,lo_sitm,lo_date,lo_prog) Values ('" & _
            sAwno & "','" & sZone & "','" & sLono & "','" & sSgst & "','" & sFbst & "'," & iSisq & ",'" & sCran & "','" & _
            sSitm & "','" & sDate & "','" & sProg & "')"
            gADOCmd.CommandText = sSQLStr
            Set gADORecord = gADOCmd.Execute
            
            oADORecord.MoveNext
        End If
    Loop
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!經軸(二)] : Error > " & "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim sTemp As String
     
    On Error GoTo ErrorHandler
    sAwno = "3"
    sZone = "1"
    sSitm = Format(Now, "yyyy/mm/dd hh:mm")
    sDate = " "
    sProg = " "
    
    sSQLStr = "Select awno,lono,losz,zone,sgst,fbst,sisq From Lo Where Awno='2'"
    oADOCmd.CommandText = sSQLStr
    Set oADORecord = oADOCmd.Execute
    Do While Not oADORecord.EOF
        If (Not IsNull(oADORecord!lono)) And Len(oADORecord!lono) = 6 Then
            sLono = Left(oADORecord!lono, 2) & "0" & Mid(oADORecord!lono, 3, 2) & "0" & Right(oADORecord!lono, 2)
            
            If oADORecord!sgst = "0" Or oADORecord!sgst = "F" Or oADORecord!sgst = "S" Or oADORecord!sgst = "R" Then
                sSgst = oADORecord!sgst
            Else
                sSgst = "E"
            End If
            
            If oADORecord!fbst = "0" Then
                sFbst = "0"
            Else
                sFbst = "X"
            End If
            
            iSisq = oADORecord!sisq
            
            sTemp = Left(oADORecord!lono, 2)
            If sTemp = "01" Or sTemp = "02" Then
                sCran = "1"
            ElseIf sTemp = "03" Or sTemp = "04" Then
                sCran = "2"
            ElseIf sTemp = "05" Or sTemp = "06" Then
                sCran = "3"
            End If

            sSQLStr = "Insert Lo (lo_awno,lo_zone,lo_lono,lo_sgst,lo_fbst,lo_sisq,lo_cran,lo_sitm,lo_date,lo_prog) Values ('" & _
            sAwno & "','" & sZone & "','" & sLono & "','" & sSgst & "','" & sFbst & "'," & iSisq & ",'" & sCran & "','" & _
            sSitm & "','" & sDate & "','" & sProg & "')"
            gADOCmd.CommandText = sSQLStr
            Set gADORecord = gADOCmd.Execute
            
            oADORecord.MoveNext
        End If
    Loop
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!織軸] : Error > " & "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub Command5_Click()
Dim sTemp As String
     
    On Error GoTo ErrorHandler
    sAwno = "4"
    sZone = "1"
    sSitm = Format(Now, "yyyy/mm/dd hh:mm")
    sDate = " "
    sProg = " "
    
    sSQLStr = "Select awno,lono,losz,zone,sgst,fbst,sisq From Lo Where Awno='5'"
    oADOCmd.CommandText = sSQLStr
    Set oADORecord = oADOCmd.Execute
    Do While Not oADORecord.EOF
        If (Not IsNull(oADORecord!lono)) And Len(oADORecord!lono) = 6 Then
            sLono = Left(oADORecord!lono, 2) & "0" & Mid(oADORecord!lono, 3, 2) & "0" & Right(oADORecord!lono, 2)
            
            If oADORecord!sgst = "0" Or oADORecord!sgst = "F" Or oADORecord!sgst = "S" Or oADORecord!sgst = "R" Then
                sSgst = oADORecord!sgst
            Else
                sSgst = "E"
            End If
            
            If oADORecord!fbst = "0" Then
                sFbst = "0"
            Else
                sFbst = "X"
            End If
            
            iSisq = oADORecord!sisq
            
            sTemp = Left(oADORecord!lono, 2)
            If sTemp = "01" Or sTemp = "02" Then
                sCran = "1"
            ElseIf sTemp = "03" Or sTemp = "04" Then
                sCran = "2"
            ElseIf sTemp = "05" Or sTemp = "06" Then
                sCran = "3"
            End If

            sSQLStr = "Insert Lo (lo_awno,lo_zone,lo_lono,lo_sgst,lo_fbst,lo_sisq,lo_cran,lo_sitm,lo_date,lo_prog) Values ('" & _
            sAwno & "','" & sZone & "','" & sLono & "','" & sSgst & "','" & sFbst & "'," & iSisq & ",'" & sCran & "','" & _
            sSitm & "','" & sDate & "','" & sProg & "')"
            gADOCmd.CommandText = sSQLStr
            Set gADORecord = gADOCmd.Execute
            
            oADORecord.MoveNext
        End If
    Loop
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!胚布] : Error > " & "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub Command6_Click()
Dim sTemp As String
     
    On Error GoTo ErrorHandler
    sAwno = "5"
    sZone = "1"
    sSitm = Format(Now, "yyyy/mm/dd hh:mm")
    sDate = " "
    sProg = " "
    
    sSQLStr = "Select awno,lono,losz,zone,sgst,fbst,sisq From Lo Where Awno='7'"
    oADOCmd.CommandText = sSQLStr
    Set oADORecord = oADOCmd.Execute
    Do While Not oADORecord.EOF
        If (Not IsNull(oADORecord!lono)) And Len(oADORecord!lono) = 6 Then
            sLono = Left(oADORecord!lono, 2) & "0" & Mid(oADORecord!lono, 3, 2) & "0" & Right(oADORecord!lono, 2)
            
            If oADORecord!sgst = "0" Or oADORecord!sgst = "F" Or oADORecord!sgst = "S" Or oADORecord!sgst = "R" Then
                sSgst = oADORecord!sgst
            Else
                sSgst = "E"
            End If
            
            If oADORecord!fbst = "0" Then
                sFbst = "0"
            Else
                sFbst = "X"
            End If
            
            iSisq = oADORecord!sisq
            
            sTemp = Left(oADORecord!lono, 2)
            If sTemp = "01" Or sTemp = "02" Then
                sCran = "1"
            ElseIf sTemp = "03" Or sTemp = "04" Then
                sCran = "2"
            ElseIf sTemp = "05" Or sTemp = "06" Then
                sCran = "3"
            End If

            sSQLStr = "Insert Lo (lo_awno,lo_zone,lo_lono,lo_sgst,lo_fbst,lo_sisq,lo_cran,lo_sitm,lo_date,lo_prog) Values ('" & _
            sAwno & "','" & sZone & "','" & sLono & "','" & sSgst & "','" & sFbst & "'," & iSisq & ",'" & sCran & "','" & _
            sSitm & "','" & sDate & "','" & sProg & "')"
            gADOCmd.CommandText = sSQLStr
            Set gADORecord = gADOCmd.Execute
            
            oADORecord.MoveNext
        End If
    Loop
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!成品] : Error > " & "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    ' SQL SERVER CONNECTION (OROCALE DATA)
    sSQLUserID = IniRead("DATABASE", "oUserID", App.Path & "\IniFile\GF4.ini")
    sSQLPassword = IniRead("DATABASE", "oPassword", App.Path & "\IniFile\GF4.ini")
    sSQLTable = IniRead("DATABASE", "oTable", App.Path & "\IniFile\GF4.ini")
    sSQLDSN = IniRead("DATABASE", "oDSN", App.Path & "\IniFile\GF4.ini")
    sSQLDbPvd = "SQLOLEDB.1"
    sSQLCon = "Provider=" & sSQLDbPvd & ";" & _
          "Password=" & sSQLPassword & ";" & _
          "Persist Security Info=True;" & _
          "Initial Catalog=" & sSQLTable & ";" & _
          "Data Source=" & sSQLDSN
    
    Set oADOCon = New ADODB.Connection
    Set oADOCmd = New ADODB.Command
    oADOCon.Open sSQLCon, sSQLUserID, sSQLPassword
    Set oADOCmd.ActiveConnection = oADOCon
    Exit Sub
ErrorHandler:
    lErrorNo = Err.Number
    sErrorMsg = Err.Description
    Trace "[frmStart!Form_Load] : Error > " & "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg
    xMsgBox "Error No : " & lErrorNo & "; Error MSg : " & sErrorMsg, vbCritical + vbOKOnly, "Error", 5
End Sub

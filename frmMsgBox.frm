VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   1  '單線固定
   Caption         =   "Testing"
   ClientHeight    =   1860
   ClientLeft      =   1650
   ClientTop       =   1905
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3825
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   90
      Top             =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "666666"
      Height          =   555
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1125
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   90
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   90
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   90
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   90
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   90
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TimerCount As Long
Private TimerLimit As Long
Private FormTitle As String
Private Answer As Long

Private Sub ArrangeSize(s As String)
  Dim tS1 As String
  Dim sLeftMaxLen As String
  Dim TemS As String
  Dim sResult As String
  Dim i As Long, j As Long, k As Long
  Const MaxLen = 90
  
  sResult = ""
  TemS = s
  Do While Len(TemS) > 0
    'take left 100 char, j is last vbcrlf positon in sLeftMaxlen
    If Len(TemS) > MaxLen Then
      sLeftMaxLen = Left$(TemS, MaxLen + 1)
      j = 0 'not find vbcrlf
      If Right$(sLeftMaxLen, 2) = vbCrLf Then
         j = MaxLen
      Else
         sLeftMaxLen = Left$(TemS, MaxLen)
         'find last vbcrlf in sLeftMaxlen
         i = 0
         tS1 = sLeftMaxLen
         Do While Len(tS1) > 0
            i = InStr(1, tS1, vbCrLf, vbBinaryCompare)
            If i > 0 Then
               tS1 = Right$(tS1, Len(tS1) - i - 1)
               j = j + i + 1
            Else
               tS1 = ""
            End If
         Loop
      End If
      Select Case j
      Case 0
         'find last empty in sLeftMaxlen, k is the last empty
         k = 0
         i = 0
         tS1 = sLeftMaxLen
         Do While Len(tS1) > 0
            i = InStr(1, tS1, " ", vbBinaryCompare)
            If i > 0 Then
               tS1 = Right$(tS1, Len(tS1) - i)
               k = k + i
            Else
               tS1 = ""
            End If
         Loop
         If k > 0 Then
            sResult = sResult + Left$(sLeftMaxLen, k) + vbCrLf
            TemS = Right$(sLeftMaxLen, Len(sLeftMaxLen) - k) + Right$(TemS, Len(TemS) - MaxLen)
         Else
            sResult = sResult + sLeftMaxLen + vbCrLf
            TemS = Right$(TemS, Len(TemS) - MaxLen)
         End If
      Case MaxLen
         sResult = sResult + sLeftMaxLen
         TemS = Right$(TemS, Len(TemS) - MaxLen - 1)
      Case Else
         sResult = sResult + Left$(sLeftMaxLen, j + 1)
         TemS = Right$(TemS, Len(TemS) - j - 1)
      End Select
    Else
      sResult = sResult + TemS
      TemS = ""
    End If
  Loop
     
  s = sResult

  
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Command1(Index).Caption
   Case "&OK"
     Answer = vbOK
   Case "&Cancel"
     Answer = vbCancel
   Case "&Abort"
     Answer = vbAbort
   Case "&Retry"
     Answer = vbRetry
   Case "&Ignore"
     Answer = vbIgnore
   Case "&Yes"
     Answer = vbYes
   Case "&No"
     Answer = vbNo
   Case Else
     Answer = 0
   End Select
   Me.Hide '
   
End Sub

Public Function TriggerForm(ByVal sM As String, lC As Long, sT As String, lT As Long) As Long
  
  
  ArrangeSize sM
  
  Label1.Caption = sM
  
  'DoEvents

  Label1.Left = 720
  Label1.Top = 90
  
  ArrangeIcon lC
  
  Me.Width = Label1.Width + 720 + 360
  Me.Height = Label1.Top + Label1.Height + 120 + Command1(0).Height + 120 + 360
  
  ArrangeButton lC
  
  Me.Height = Command1(0).Top + Command1(0).Height + 120 + 360
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
    
  FormTitle = sT
  Me.Caption = FormTitle
  
  TimerCount = 0
  If lT > 0 Then
     TimerLimit = lT
     Timer1.Enabled = True
  Else
     Timer1.Enabled = False
     TimerLimit = 0
  End If
      
  SetFinalValue lC
  
  DoEvents
  
  On Error Resume Next
  Me.Show vbModal
  'If ERR Then
  '  ERR.Clear
  '  Me.Show
  'End If
    
  TriggerForm = Answer
End Function
Private Sub ArrangeIcon(lv As Long)
   
   If (lv And &HF0) = vbCritical Then
      Image1(0).Visible = True
      Image1(1).Visible = False
      Image1(2).Visible = False
      Image1(3).Visible = False
   ElseIf (lv And &HF0) = vbQuestion Then
      Image1(0).Visible = False
      Image1(1).Visible = True
      Image1(2).Visible = False
      Image1(3).Visible = False
   ElseIf (lv And &HF0) = vbExclamation Then
      Image1(0).Visible = False
      Image1(1).Visible = False
      Image1(2).Visible = True
      Image1(3).Visible = False
   ElseIf (lv And &HF0) = vbInformation Then
      Image1(0).Visible = False
      Image1(1).Visible = False
      Image1(2).Visible = False
      Image1(3).Visible = True
   Else
      Image1(0).Visible = False
      Image1(1).Visible = False
      Image1(2).Visible = False
      Image1(3).Visible = False
   End If
   
End Sub
Private Sub ArrangeButton(lv As Long)
   
   If (lv And &HF) = vbOKOnly Then
      LoadButton 1, "&OK", "", ""
   ElseIf (lv And &HF) = vbOKCancel Then
      LoadButton 2, "&OK", "&Cancel", ""
   ElseIf (lv And &HF) = vbAbortRetryIgnore Then
      LoadButton 3, "&Abort", "&Retry", "&Ignore"
   ElseIf (lv And &HF) = vbYesNoCancel Then
      LoadButton 3, "&Yes", "&No", "&Cancel"
   ElseIf (lv And &HF) = vbYesNo Then
      LoadButton 2, "&Yes", "&No", ""
   ElseIf (lv And &HF) = vbRetryCancel Then
      LoadButton 2, "&Retry", "&Cancel", ""
   Else
      LoadButton 1, "&OK", "", ""
   End If
End Sub
Private Sub LoadButton(ln As Long, c1 As String, c2 As String, c3 As String)
  Dim l As Long
  Const Distinct = 180
  l = (1260 + (ln - 1) * Distinct + ln * Command1(0).Width)
  If Me.Width < l Then Me.Width = l
  Command1(0).Top = Label1.Top + Label1.Height + 120
  If Command1(0).Top < 1560 Then Command1(0).Top = 1560
  Command1(0).Caption = c1
  Command1(0).Visible = True
  'command positon
  Select Case ln
  Case 1
     Command1(0).Left = (Me.ScaleWidth - Command1(0).Width) / 2
  Case 2
     Load Command1(1)
     Command1(1).Caption = c2
     Command1(1).Visible = True
     Command1(0).Left = Me.ScaleWidth / 2 - Command1(0).Width - Distinct / 2
     Command1(1).Left = Me.ScaleWidth / 2 + Distinct / 2
  Case 3
     Load Command1(1)
     Command1(1).Caption = c2
     Command1(1).Visible = True
     Load Command1(2)
     Command1(2).Caption = c3
     Command1(2).Visible = True
     Command1(1).Left = (Me.ScaleWidth - Command1(0).Width) / 2
     Command1(0).Left = Command1(1).Left - Distinct - Command1(0).Width
     Command1(2).Left = Command1(1).Left + Distinct + Command1(1).Width
  End Select
  
   
End Sub

Private Sub SetFinalValue(lv As Long)

   
   On Error GoTo errorhandle
   If (lv And &HF00) = vbDefaultButton1 Then
       Command1(0).TabIndex = 0
   ElseIf (lv And &HF00) = vbDefaultButton2 Then
       Command1(1).TabIndex = 0
   ElseIf (lv And &HF00) = vbDefaultButton3 Then
       Command1(2).TabIndex = 0
   ElseIf (lv And &HF00) = vbDefaultButton4 Then
       Command1(3).TabIndex = 0
   Else
       Command1(0).TabIndex = 0
   End If
   Exit Sub
errorhandle:
    Command1(0).TabIndex = 0
End Sub

Private Sub Form_Activate()
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE


End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub Label1_Click()
   Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
   TimerCount = TimerCount + 1
   Me.Caption = FormTitle + " - " + CStr(TimerLimit - TimerCount)
   If TimerCount >= TimerLimit Then
      Timer1.Enabled = False
      Unload Me 'Unload Me
      Set frmMsgBox = Nothing
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
Select Case UnloadMode
Case vbFormControlMenu  ' 0   使用者從表單上的控制功能表中選取「關閉」指令。
  Cancel = True
'  'Me.Visible = False
'Case vbFormCode ' 1   Unload 陳述式被程式碼呼叫。
'  Cancel = False
'  'me.visible=false
'Case vbAppWindows   ' 2   目前 Microsoft Windows 作業環境任務結束。
'  Cancel = False
'Case vbAppTaskManager   ' 3   Microsoft Windows 工作管理員正在關閉應用程式。
'  Cancel = False
'Case vbFormMDIForm '4   因為 MDI 表單正在關閉的緣故，MDI 子表單正在關閉。
'  Cancel = False
End Select
End Sub

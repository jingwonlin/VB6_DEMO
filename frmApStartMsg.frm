VERSION 5.00
Begin VB.Form frmApStartMsg 
   BorderStyle     =   3  '雙線固定對話方塊
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblApStartMsg 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "啟動中,請稍後..."
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   15.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7950
   End
End
Attribute VB_Name = "frmApStartMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
End Sub

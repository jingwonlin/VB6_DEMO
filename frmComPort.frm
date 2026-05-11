VERSION 5.00
Begin VB.Form frmComPort 
   Caption         =   "ComPort"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.Frame Frame6 
      Caption         =   "成品通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1035
      Left            =   90
      TabIndex        =   72
      Top             =   4770
      Width           =   11610
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Index           =   25
         Left            =   1248
         TabIndex        =   78
         Top             =   288
         Width           =   6048
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   25
         Left            =   7632
         TabIndex        =   77
         Top             =   288
         Width           =   792
      End
      Begin VB.Label Label26 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   76
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label24 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8544
         TabIndex        =   75
         Top             =   288
         Width           =   660
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   25
         Left            =   1245
         TabIndex        =   74
         Top             =   585
         Width           =   6045
      End
      Begin VB.Label Label14 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   192
         TabIndex        =   73
         Top             =   288
         Width           =   996
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "胚布通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1035
      Left            =   90
      TabIndex        =   56
      Top             =   3645
      Width           =   11610
      Begin VB.Label Label16 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   192
         TabIndex        =   62
         Top             =   288
         Width           =   996
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   24
         Left            =   1245
         TabIndex        =   61
         Top             =   585
         Width           =   6045
      End
      Begin VB.Label Label7 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8544
         TabIndex        =   60
         Top             =   288
         Width           =   660
      End
      Begin VB.Label Label6 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   59
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   24
         Left            =   7632
         TabIndex        =   58
         Top             =   288
         Width           =   792
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Index           =   24
         Left            =   1248
         TabIndex        =   57
         Top             =   288
         Width           =   6048
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "AGV通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   2265
      Left            =   90
      TabIndex        =   20
      Top             =   5895
      Width           =   11625
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   690
         Left            =   10575
         TabIndex        =   79
         Top             =   1170
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "離    開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   9765
         TabIndex        =   70
         Top             =   315
         Width           =   1680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "執行測試"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10035
         TabIndex        =   48
         Top             =   1665
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtPalt 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   6165
         MaxLength       =   5
         TabIndex        =   47
         Top             =   1755
         Width           =   1065
      End
      Begin VB.ComboBox cmbStNow 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7290
         TabIndex        =   46
         Top             =   1755
         Width           =   1275
      End
      Begin VB.ComboBox cmbStNext 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8610
         TabIndex        =   45
         Top             =   1740
         Width           =   1275
      End
      Begin VB.Label lblAGVStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1245
         TabIndex        =   55
         Top             =   1710
         Width           =   4770
      End
      Begin VB.Label Label15 
         Caption         =   "AGV2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   54
         Top             =   1755
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "AGV1"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   53
         Top             =   1395
         Width           =   765
      End
      Begin VB.Label lblAGVStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1245
         TabIndex        =   52
         Top             =   1350
         Width           =   4770
      End
      Begin VB.Label Label13 
         Caption         =   "工作序號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6390
         TabIndex        =   51
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "啟始站號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7515
         TabIndex        =   50
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "下一站號"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8760
         TabIndex        =   49
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label lblComPortOutString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   276
         Index           =   9
         Left            =   1260
         TabIndex        =   30
         Top             =   336
         Width           =   6048
      End
      Begin VB.Label Label25 
         Caption         =   "傳送字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   216
         TabIndex        =   29
         Top             =   336
         Width           =   996
      End
      Begin VB.Label Label23 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8592
         TabIndex        =   28
         Top             =   816
         Width           =   660
      End
      Begin VB.Label lblComPortOutLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   276
         Index           =   9
         Left            =   7680
         TabIndex        =   27
         Top             =   336
         Width           =   792
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   9
         Left            =   1260
         TabIndex        =   26
         Top             =   630
         Width           =   6045
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   9
         Left            =   7680
         TabIndex        =   25
         Top             =   768
         Width           =   792
      End
      Begin VB.Label Label22 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   24
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label21 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8592
         TabIndex        =   23
         Top             =   336
         Width           =   660
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   9
         Left            =   1260
         TabIndex        =   22
         Top             =   945
         Width           =   6045
      End
      Begin VB.Label Label20 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   630
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "織軸通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1035
      Left            =   90
      TabIndex        =   13
      Top             =   2520
      Width           =   11610
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Index           =   23
         Left            =   1248
         TabIndex        =   19
         Top             =   288
         Width           =   6048
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   23
         Left            =   7632
         TabIndex        =   18
         Top             =   288
         Width           =   792
      End
      Begin VB.Label Label10 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   17
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label9 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8544
         TabIndex        =   16
         Top             =   288
         Width           =   660
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   23
         Left            =   1245
         TabIndex        =   15
         Top             =   585
         Width           =   6045
      End
      Begin VB.Label Label8 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   192
         TabIndex        =   14
         Top             =   288
         Width           =   996
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "經軸通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1020
      Left            =   90
      TabIndex        =   6
      Top             =   1395
      Width           =   11610
      Begin VB.Label Label5 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   216
         TabIndex        =   12
         Top             =   288
         Width           =   996
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   22
         Left            =   1245
         TabIndex        =   11
         Top             =   585
         Width           =   6045
      End
      Begin VB.Label Label4 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   8544
         TabIndex        =   10
         Top             =   288
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   9
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Index           =   22
         Left            =   7632
         TabIndex        =   8
         Top             =   288
         Width           =   792
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Index           =   22
         Left            =   1248
         TabIndex        =   7
         Top             =   288
         Width           =   6048
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "關閉"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   13440
      TabIndex        =   1
      Top             =   8880
      Width           =   1365
   End
   Begin VB.Frame Frame14 
      Caption         =   "原紗通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1200
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   11610
      Begin VB.CommandButton Command3 
         Caption         =   "出入庫指示燈測試"
         Height          =   870
         Left            =   9990
         TabIndex        =   71
         Top             =   225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label37 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   44
         Top             =   1935
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblComPortStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   21
         Left            =   1260
         TabIndex        =   43
         Top             =   675
         Width           =   6045
      End
      Begin VB.Label Label36 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8460
         TabIndex        =   42
         Top             =   1935
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label35 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   41
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   7650
         TabIndex        =   40
         Top             =   1935
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   4
         Left            =   1215
         TabIndex        =   39
         Top             =   1935
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label Label33 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   38
         Top             =   1620
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label32 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8460
         TabIndex        =   37
         Top             =   1620
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   7650
         TabIndex        =   36
         Top             =   1620
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   3
         Left            =   1215
         TabIndex        =   35
         Top             =   1620
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label Label28 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   34
         Top             =   1350
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label27 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8460
         TabIndex        =   33
         Top             =   1260
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   7650
         TabIndex        =   32
         Top             =   1260
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   2
         Left            =   1215
         TabIndex        =   31
         Top             =   1305
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label lblComPortString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   21
         Left            =   1245
         TabIndex        =   5
         Top             =   315
         Width           =   6045
      End
      Begin VB.Label lblComPortLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   21
         Left            =   7635
         TabIndex        =   4
         Top             =   285
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8496
         TabIndex        =   3
         Top             =   336
         Width           =   660
      End
      Begin VB.Label Label31 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   192
         TabIndex        =   2
         Top             =   336
         Width           =   996
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "BarCode通訊埠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1035
      Left            =   90
      TabIndex        =   63
      Top             =   6075
      Visible         =   0   'False
      Width           =   11610
      Begin VB.Label lblBarCodeString 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   264
         Left            =   1248
         TabIndex        =   69
         Top             =   288
         Width           =   6048
      End
      Begin VB.Label lblBarCodeLen 
         Alignment       =   2  '置中對齊
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   7632
         TabIndex        =   68
         Top             =   288
         Width           =   792
      End
      Begin VB.Label Label19 
         Caption         =   "通訊狀態"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   195
         TabIndex        =   67
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label18 
         Caption         =   "Byte"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   8544
         TabIndex        =   66
         Top             =   288
         Width           =   660
      End
      Begin VB.Label lblBarCodeStatus 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1245
         TabIndex        =   65
         Top             =   585
         Width           =   6045
      End
      Begin VB.Label Label17 
         Caption         =   "接收字串"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   192
         TabIndex        =   64
         Top             =   288
         Width           =   996
      End
   End
End
Attribute VB_Name = "frmComPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim lRtn As Long
Dim sStNoNow As String
Dim sStNoNext As String
Dim sStNoEnd As String
Dim i As Long

'    sStNoNow = cmbStNow.Text
'    If Len(sStNoNow) <> 5 Then
'        xMsgBox "啟始站號須5位數(例BXXXX) ", vbCritical + vbOKOnly, "Error", 5
'        Exit Sub
'    End If
    
    If Left(cmbStNow.Text, 2) = "03" Then
         sStNoNow = "B1" & Mid(cmbStNow.Text, 4, 2) & Mid(cmbStNow.Text, 8, 1)
     Else
        Select Case Trim(cmbStNow.Text)
        Case "A010"
            sStNoNow = "X1   "
        Case "A011"
            sStNoNow = "X2   "
        Case "Y1"
            sStNoNow = "Y1   "
        Case "Y2"
            sStNoNow = "Y2   "
        Case "Z1"
            sStNoNow = "Z1   "
        Case "Z2"
            sStNoNow = "Z2   "
        End Select
     End If
    
    
    Select Case Trim(cmbStNext.Text)
        Case "A010"
            sStNoNext = "X1   "
        Case "A011"
            sStNoNext = "X2   "
        Case "Y1"
            sStNoNext = "Y1   "
        Case "Y2"
            sStNoNext = "Y2   "
        Case Else

            sStNoNext = cmbStNext.Text
            If Len(sStNoNext) <> 5 Then
                xMsgBox "下一站號須5位數(例BXXXX) ", vbCritical + vbOKOnly, "Error", 5
                Exit Sub
            End If
    End Select
    lRtn = sio_putch(AGV_Port, STX)
    Sleep (500)
    AGV_TempStr = AGV + HOST + "T" + DUMMY + "0" + Format(txtPalt.Text, "00000") + sStNoNow + sStNoNext + "1"
    Call PUT_DLE_ETX_BCC
End Sub

Private Sub Command2_Click()
     Unload Me
End Sub

Private Sub Command3_Click()
Dim lRtn As Long
Dim sSend As String
     sSend = ">350" & Chr(13) & Chr(10)
     lRtn = sio_write(15, sSend, 6)
     '經軸 1=st2 , 2=st4 , 3=st6 , 4=st8 , 5=st1 , 6=st3 , 7=st5 , 8=st7
     '織軸 1=st2 , 2=st1 , 3=st3 , 4=st4 , 5=st5
End Sub

Private Sub Command5_Click()
Dim iResult As Integer
Dim lRtn As Long
    lRtn = sio_putch(AGV_Port, DLE)
    If lRtn > 0 Then
    iResult = iResult + 1
    End If
    
'                    Call Sleep(300)
'                    lRtn = sio_iqueue(AGV_Port)
'                    If lRtn > 0 Then
'                        iDataLen = sio_read(AGV_Port, AGV_InBuf, lRtn)
'                        Call AGV_To_HOST
End Sub

Private Sub Form_Load()
cmbStNow.AddItem "Z1   "
cmbStNow.AddItem "Z2   "

cmbStNext.AddItem "A010"
cmbStNext.AddItem "A011"
cmbStNext.AddItem "Z2   "
End Sub


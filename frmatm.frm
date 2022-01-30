VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmatm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ATM   Machine"
   ClientHeight    =   9405
   ClientLeft      =   5535
   ClientTop       =   2220
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1935
      Left            =   2640
      TabIndex        =   68
      Top             =   2880
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "EXIT"
         Height          =   255
         Left            =   4680
         TabIndex        =   72
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtdpin 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtnpin 
         Height          =   435
         Left            =   2400
         TabIndex        =   1
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "CHANGE"
         Height          =   255
         Left            =   3120
         TabIndex        =   69
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Pin:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   71
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Pin:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         TabIndex        =   70
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Timer Timer_Font 
      Left            =   10560
      Top             =   1200
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2040
      TabIndex        =   42
      Top             =   2640
      Width           =   7815
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   720
         TabIndex        =   75
         Top             =   1800
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   450
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   2520
         Top             =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Please Wait ...     Your Transaction is being proccess ......"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   43
         Top             =   960
         Width           =   7215
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   2160
      TabIndex        =   50
      Top             =   2280
      Width           =   7575
      Begin VB.Label lbl_Reciept 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Take Your Cash and Reciept"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   3120
         Width           =   7575
      End
      Begin VB.Label lbl_Frame4_AvailBalance 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   55
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label lbl_Frame4_CurrBalance 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   54
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lbl_Frame4_AccNum 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   53
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label lbl_Frame4_Name 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   51
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Balance Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         TabIndex        =   52
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   33
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   3855
      Begin VB.CommandButton Command6 
         BackColor       =   &H8000000B&
         Caption         =   "Take"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaskColor       =   &H008080FF&
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtReciept 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   2775
      End
      Begin VB.Image Image_cash 
         Height          =   720
         Left            =   3000
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Take Your Cash and Reciept Here !!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   7680
      TabIndex        =   20
      Top             =   6120
      Width           =   4215
      Begin VB.Timer Timer_Blink 
         Left            =   2880
         Top             =   2280
      End
      Begin VB.CommandButton cmd_TakeCard 
         Caption         =   "Take Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         Caption         =   "&Insert Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   2760
         Left            =   600
         Picture         =   "frmatm.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "INSERT  ATM  CARD  HERE  !!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command18 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton withdrawal 
      Caption         =   "Withdrawal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdnewaccount 
      Caption         =   "  New Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Pin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBalInqury 
      Caption         =   "Balance Inquiry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmddeposit 
      Caption         =   "  Make Deposit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtWithdraw 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   45
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   8640
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   615
      Left            =   4080
      TabIndex        =   74
      Top             =   7560
      Visible         =   0   'False
      Width           =   3735
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   41
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6588
      _cy             =   1085
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   495
      Left            =   3840
      TabIndex        =   73
      Top             =   8640
      Visible         =   0   'False
      Width           =   3735
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   7
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   98
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6588
      _cy             =   873
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Pan African Bank"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   7875
      TabIndex        =   67
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   0
      TabIndex        =   66
      Top             =   600
      Width           =   12015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11160
      TabIndex        =   64
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl_WithAmount 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   63
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label lblCurrBalance 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   62
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   61
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   60
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   59
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label_Bal_Inquery 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Click Enter Button to Proccess Your Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3000
      TabIndex        =   58
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label lblIDeposit 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   49
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lblLName 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   48
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblCStatus 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   47
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblAccNum 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   46
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lblWthdrawAmount 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Enter  Withdrawal  Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2160
      TabIndex        =   44
      Top             =   3240
      Width           =   7695
   End
   Begin VB.Label lblBanner 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Thanks for Banking with us !!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2040
      TabIndex        =   41
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Label lblInqueryName 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   40
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label lblInqueryAccNum 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Top             =   3960
      Width           =   5415
   End
   Begin VB.Label lblInqueryAvailBalance 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   38
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label lblAccInfo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Balance Inquery"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2040
      TabIndex        =   37
      Top             =   2640
      Width           =   7695
   End
   Begin VB.Label lblSTransaction 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Select Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   36
      Top             =   4800
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label lblEnterPIN 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Enter Your [  PIN  ]  Personal Identification Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2040
      TabIndex        =   34
      Top             =   2640
      Width           =   7695
   End
   Begin VB.Label lblPassword 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   32
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label lblAddress 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   31
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblSex 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   30
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblFName 
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   29
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   5160
      Left            =   1920
      Picture         =   "frmatm.frx":246A
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   8040
   End
   Begin VB.Label lblEnterButton 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Click Enter Button to Proccess Your Transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   35
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "      FIRST BANK NIGERIA PLC"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "frmatm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ButtonClick, Cancel_Button_click  As Integer
Dim time As String
Dim Transaction As String
Dim fname, LName, Sex, CStatus, Address As String
Dim WithdrawalAmount, CurrBalance, AccNum, Password, Ideposit, NewBalance, AvailBalance As Long

Private Sub cmd_TakeCard_Click()

ins = 0
Cancel_Button_click = 0
Timer_Blink.Enabled = True
Timer_Font.Enabled = True

Label1.Caption = " INSERT  ATM  CARD  HERE  !! "

cmdnewaccount.Visible = False
cmddeposit.Visible = False
cmdBalInqury.Visible = False
withdrawal.Visible = False
Cancel.Visible = False
cmd_TakeCard.Visible = False
lblSTransaction.Visible = False
Command3.Visible = False
cmd_TakeCard.Visible = False
Command19.Enabled = True
lblSTransaction.Visible = False
'wmp.URL = App.Path & "\atm2.amr"
   'wmp2.URL = App.Path & "\Celindion.mp3"
   wmp2.Controls.play
   


Image_cash.Visible = False
Label_Bal_Inquery.Visible = False
txtReciept.Visible = False

Frame3.Visible = False
Frame4.Visible = False
lblBanner.Visible = True
lblSTransaction.Visible = False
lblAccInfo.Visible = False
lblInqueryName.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterButton.Visible = False
lblEnterPIN.Visible = False
lblEnterButton.Visible = False
lblWthdrawAmount.Visible = False
txtWithdraw.Visible = False
txtPassword.Visible = False

frmatm.lblFName = ""
frmatm.lblLName = ""
frmatm.lblSex = ""
frmatm.lblCStatus = ""
frmatm.lblAddress = ""
frmatm.lblAccNum = ""
frmatm.lblPassword = ""
frmatm.lblIDeposit = ""

End Sub

Private Sub cmdBalInqury_Click()
On Error GoTo err

If ins = 0 Then
 MsgBox " Please Insert Your ATM card first before you begin transaction", vbCritical, " ATM Card not inserted"
 
Else
Network

sql = "select * from New_Account where ATM_Pin='" & pin & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
lbl_Frame4_AccNum.Caption = "Account No: " & !Account_Number
lbl_Frame4_Name.Caption = "Names: " & !Surname & " " & !Othernames
lbl_Frame4_AvailBalance.Caption = "Available Bal.: " & !Available_Balance
lbl_Frame4_CurrBalance.Caption = "Ledger Balance: " & !balance

Label_Bal_Inquery.Visible = True
Frame3.Visible = False
Frame4.Visible = True
lbl_Frame4_AccNum.Visible = True
lbl_Frame4_AvailBalance.Visible = True
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterButton.Visible = False
lblEnterPIN.Visible = False
lblEnterButton.Visible = False
lblWthdrawAmount.Visible = False
txtWithdraw.Visible = False
txtPassword.Visible = False
lblBanner.Visible = False
lblSTransaction.Visible = False
Timer_Font.Enabled = False
Timer_Blink.Enabled = False
Shape1.Visible = False
lbl_Reciept.Visible = False
End With
End If
Exit Sub
err:
    MsgBox "If you have changed your pin please remove and re-insert your ATM Card", vbCritical, "Enquiry"
    
End Sub

Private Sub cmdchange_Click()

On Error GoTo err

If txtdpin.Text = Empty Then
MsgBox "Default Pin is Required", vbCritical, "Change Pin"
Exit Sub
End If


Network

sql = "select * from New_Account where ATM_Pin='" & Trim(txtdpin.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
!ATM_Pin = txtnpin.Text
.Update
End With
MsgBox "ATM pin changed successfully", vbInformation, "Pin changed"
txtdpin = Empty
txtnpin = Empty
Exit Sub
err:
MsgBox "There was a problem changing pin", vbCritical, "Pin change"
End Sub

Private Sub cmdEnter_Click()
On Error GoTo A

If ins = 0 Then
  MsgBox " Please Insert Your ATM card first before you begin transaction", vbCritical, " ATM Card not inserted"
Exit Sub
End If
  
        If Cancel_Button_click = 1 Then
        MsgBox " You have Cancel your Transaction .. Insert ATM card again to begin another transaction ", vbCritical, " Transaction aborted by user"
        Exit Sub
        End If

cmdBalInqury.Enabled = True
withdrawal.Enabled = True
lblBanner.Visible = True
lblSTransaction.Visible = True
lblAccInfo.Visible = False
lblInqueryName.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterButton.Visible = False
lblEnterPIN.Visible = False
cmd_TakeCard.Enabled = False

txtWithdraw.Visible = False
lblWthdrawAmount.Visible = False
Frame4.Visible = False
txtPassword.Visible = False
Label_Bal_Inquery.Visible = False

Network

sql = "select * from New_Account where ATM_Pin='" & Trim(txtPassword.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
If txtPassword.Text = !ATM_Pin Then
bool = 2
frmfinger.cmdProceed.Visible = False
frmfinger.Show vbModal
Else
GoTo A
End If
End With
 Exit Sub
 
A:
   MsgBox " Inccorect [ PIN ] Personal Identification Number ", vbCritical, "Invalid Input"
  
 txtPassword = ""
 Frame4.Visible = False
lblInqueryName.Visible = False
lblAccInfo.Visible = False
lblBanner.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterPIN.Visible = True
txtPassword.Visible = True
lblEnterButton.Visible = False
lblSTransaction.Visible = False
txtPassword.SetFocus
cmdBalInqury.Enabled = False
withdrawal.Enabled = False

End Sub

Private Sub Cancel_Click()

Cancel_Button_click = 1
Timer_Blink.Enabled = True
Timer_Font.Enabled = True

Frame6.Visible = False
cmddeposit.Visible = False
cmdnewaccount.Visible = False
cmdBalInqury.Visible = False
withdrawal.Visible = False
Command19.Enabled = False
cmd_TakeCard.Enabled = True
Cancel.Visible = False
Command3.Visible = False
Frame2.Visible = False
cmd_TakeCard.Visible = True
cmdEnter.Enabled = True

wmp.settings.playCount = 1
wmp.URL = App.Path & "\atm2.amr"
   wmp.Controls.play
   



Label1.Caption = " INSERT  ATM  CARD  HERE  !! "

lblBanner.Visible = True
lblSTransaction.Visible = False
lblAccInfo.Visible = False
lblInqueryName.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterButton.Visible = False
lblEnterPIN.Visible = False
lblEnterButton.Visible = False
txtWithdraw.Visible = False
lblWthdrawAmount.Visible = False
Frame4.Visible = False
txtPassword.Visible = False
Label_Bal_Inquery.Visible = False

Transaction = ""
frmatm.lblFName = ""
frmatm.lblLName = ""
frmatm.lblSex = ""
frmatm.lblCStatus = ""
frmatm.lblAddress = ""
frmatm.lblAccNum = ""
frmatm.lblPassword = ""
frmatm.lblIDeposit = ""

cmdBalInqury.Enabled = True
withdrawal.Enabled = True
    If trap = 0 Then
MsgBox " Transaction cancelled .. Take your ATM Card ", vbInformation, "Aborted by user "
    End If
End Sub



Private Sub cmddeposit_Click()
frmdeposit.Show
End Sub

Private Sub Command1_Click()
Frame6.Visible = False
End Sub

Private Sub Command10_Click()
txtdpin = txtdpin & "8"
txtnpin = txtnpin & "8"
txtPassword = txtPassword & "8"
txtWithdraw = txtWithdraw & "8"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin
End Sub

Private Sub Command11_Click()
txtdpin = txtdpin & "7"
txtnpin = txtnpin & "7"
txtPassword = txtPassword & "7"
txtWithdraw = txtWithdraw & "7"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin

End Sub

Private Sub Command12_Click()
txtdpin = txtdpin & "6"
txtnpin = txtnpin & "6"
txtPassword = txtPassword & "6"
txtWithdraw = txtWithdraw & "6"
txtPassword1 = txtPassword
txtWithdraw1 = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin

End Sub

Private Sub Command13_Click()
txtdpin = txtdpin & "5"
txtnpin = txtnpin & "5"
txtPassword = txtPassword & "5"
txtWithdraw = txtWithdraw & "5"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin


End Sub

Private Sub Command14_Click()
txtdpin = txtdpin & "4"
txtnpin = txtnpin & "4"
txtPassword = txtPassword & "4"
txtWithdraw = txtWithdraw & "4"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin
End Sub

Private Sub Command15_Click()
txtdpin = txtdpin & "3"
txtnpin = txtnpin & "3"
txtPassword = txtPassword & "3"
txtWithdraw = txtWithdraw & "3"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin


End Sub

Private Sub Command16_Click()
txtdpin = txtdpin & "2"
txtnpin = txtnpin & "2"
txtPassword = txtPassword & "2"
txtWithdraw = txtWithdraw & "2"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin

End Sub

Private Sub Command17_Click()

txtPassword = ""
txtWithdraw = ""



End Sub

Private Sub Command18_Click()
txtdpin = txtdpin & "0"
txtnpin = txtnpin & "0"
txtPassword = txtPassword & "0"
txtWithdraw = txtWithdraw & "0"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin


End Sub

Private Sub Command19_Click()

Command19.Enabled = False
ins = 1
Cancel_Button_click = 0
Transaction = "CheckPIN"
Timer_Blink.Enabled = False
Timer_Font.Enabled = False
Shape1.Visible = False

Label1.Caption = " CARD INSERTED "

cmdBalInqury.Enabled = False
withdrawal.Enabled = False
txtPassword = ""
lblInqueryName.Visible = False
lblAccInfo.Visible = False
lblBanner.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterPIN.Visible = True
txtPassword.Visible = True
Label_Bal_Inquery.Visible = False
lblEnterButton.Visible = False
lblSTransaction.Visible = False
lblWthdrawAmount.Visible = False
   txtWithdraw.Visible = False
   Frame4.Visible = False
txtPassword.SetFocus
'Set mm = CreateObject("SAPI.SPVoice")
'mm.speak "Welcome....."
'mm.speak "please      Enter your secret   number"

   wmp2.Controls.pause
   
wmp.settings.playCount = 1
wmp.URL = App.Path & "\atm.amr"
   wmp.Controls.play

End Sub


Private Sub cmdnewaccount_Click()
frmnewaccount.Show
End Sub

Private Sub Command5_Click()
Dim fname, LName, Sex, CStatus, Address As String
Dim WithdrawalAmount, CurrBalance, AccNum, Password, Ideposit, NewBalance, AvailBalance As Long

WithdrawalAmount = txtWithdraw
CurrBalance = lblIDeposit
NewBalance = CurrBalance - WithdrawalAmount


fname = lblFName
LName = lblLName
Sex = lblSex
CStatus = lblCStatus
Address = lblAddress
AccNum = lblAccNum
Password = lblPassword
Ideposit = NewBalance

Open App.Path & "/users/" & Password & ".txt" For Output As #1
Print #1, fname
Print #1, LName
Print #1, Sex
Print #1, CStatus
Print #1, Address
Print #1, Ideposit
Print #1, AccNum
Print #1, Password
Close #1


Frame3.Visible = True
Timer1.Enabled = True
Timer1.Interval = 20
ProgressBar1 = 0


Frame4.Visible = True
lbl_Frame4_Name = "Account Name          : " & "   " & fname & " " & LName
lbl_Frame4_AccNum = "Account Number       : " & "   " & AccNum
lbl_Frame4_CurrBalance = "Current Balance         : " & "   " & CurrBalance
lbl_Frame4_AvailBalance = "Available Balance       : " & "   " & NewBalance
lbl_Frame4_WithdrawalAmount = "Withdrawal Amount   : " & "   " & WithdrawalAmount


txtReciept = vbCrLf + "              Official Reciept " + vbCrLf + vbCrLf + " Name     :    " & fname & " " & LName + vbCrLf + " Account #      :  " & AccNum + vbCrLf + " Current Bal     :  " & CurrBalance + vbCrLf + " Withdrawal     :  " & WithdrawalAmount + vbCrLf + " Available Bal  :  " & NewBalance



End Sub

Private Sub Command3_Click()

Frame6.Visible = True

End Sub

Private Sub Command6_Click()

On Error GoTo err


If Val(txtWithdraw.Text) < 500 Then
    MsgBox "Amount should not be less than 500", vbCritical, "Withdrawal"
    txtWithdraw.Text = Empty
    txtWithdraw.SetFocus
    Exit Sub
End If

 Network
  
  sql = " select * from New_Account where ATM_Pin='" & pin & "'"
  rec.Open sql, con, adOpenDynamic, adLockOptimistic
  
With rec
If Val(txtWithdraw.Text) > Val(!Available_Balance) Then
MsgBox "Insufficient Balance", vbCritical, "Withdrawal"
txtWithdraw.Text = Empty
txtWithdraw.SetFocus
Else
Frame3.Visible = True
  Timer1.Enabled = True
   !Available_Balance = Val(!Available_Balance) - Val(txtWithdraw.Text)
  !balance = Val(!balance) - Val(txtWithdraw.Text)
  .Update
  wmp.settings.playCount = 7
wmp.URL = App.Path & "\A.T.M.mp3"
   wmp.Controls.play
   Command6.Enabled = False
   cmddeposit.Visible = False
   cmdnewaccount.Visible = False
   Command3.Visible = False
   cmdBalInqury.Visible = False
   withdrawal.Visible = False
   Cancel.Visible = False
   cmdEnter.Enabled = False
   
End If
End With
txtWithdraw.Text = Empty
txtWithdraw.SetFocus
Exit Sub
err:
MsgBox "Error in transaction", vbCritical, "Error"

End Sub

Private Sub Command7_Click()
txtdpin = txtdpin & "1"
txtnpin = txtnpin & "1"
txtPassword = txtPassword & "1"
txtWithdraw = txtWithdraw & "1"

txtdpin = txtdpin
txtnpin = txtnpin
txtPassword = txtPassword
txtWithdraw = txtWithdraw
Label3 = txtWithdraw

End Sub

Private Sub Command9_Click()
txtdpin = txtdpin & "9"
txtnpin = txtnpin & "9"
txtPassword = txtPassword & "9"
txtWithdraw = txtWithdraw & "9"
txtPassword = txtPassword
txtWithdraw = txtWithdraw
txtdpin = txtdpin
txtnpin = txtnpin


End Sub

Private Sub Form_Load()

trap = 0

wmp2.settings.playCount = 10
'wmp2.URL = App.Path & "\Kuch Kuch Huta Hai.mp3"
wmp2.URL = App.Path & "\Celindion.mp3"
wmp2.Controls.play


Me.Left = (Screen.width - Me.width) / 4
Me.Top = (Screen.height - Me.height) / 10


ins = 0
Cancel_Button_click = 0
Image_cash.Visible = False
Label_Bal_Inquery.Visible = False
txtReciept.Visible = False
Timer_Blink.Enabled = True
Timer_Blink.Interval = 50
Timer_Font.Enabled = True
Timer_Font.Interval = 300

Command3.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
lblBanner.Visible = True
'lblSTransaction.Visible = True
lblAccInfo.Visible = False
lblInqueryName.Visible = False
lblInqueryAccNum.Visible = False
lblInqueryAvailBalance.Visible = False
lblEnterButton.Visible = False
lblEnterPIN.Visible = False
lblEnterButton.Visible = False
lblWthdrawAmount.Visible = False
txtWithdraw.Visible = False
txtPassword.Visible = False

frmatm.lblFName = ""
frmatm.lblLName = ""
frmatm.lblSex = ""
frmatm.lblCStatus = ""
frmatm.lblAddress = ""
frmatm.lblAccNum = ""
frmatm.lblPassword = ""
frmatm.lblIDeposit = ""

End Sub



Private Sub Label9_Click()
Unload Me
End Sub

Private Sub Timer_Blink_Timer()

If Shape1.Visible = True Then
   Shape1.Visible = False
Else
   Shape1.Visible = True
End If

End Sub

Private Sub Timer_Font_Timer()

If lblBanner.Visible = True Then
   lblBanner.Visible = False
Else
   lblBanner.Visible = True
End If


End Sub

Private Sub Timer1_Timer()
 If ProgressBar1 = 100 Then
 Timer1.Enabled = False
 Frame3.Visible = False
 Image_cash.Visible = True
 txtReciept.Visible = True
 MsgBox "Please take your Money", vbInformation, "Transaction Completed"
    txtWithdraw.Visible = False
    lblWthdrawAmount.Visible = False
    
    time = MsgBox("Do you want  to perform another transaction", vbInformation + vbYesNo, "Transaction")
    If time = vbYes Then
    
    lblSTransaction.Visible = True
    Command19_Click
    Frame2.Visible = False
    cmdEnter.Enabled = True
    Image_cash.Visible = False
    txtReciept.Visible = False
     Exit Sub
    Else
    cmd_TakeCard.Visible = False
    trap = 1
    Cancel_Click
    Exit Sub
    End If
   
 Else
   ProgressBar1 = ProgressBar1 + 1
 End If

End Sub

Private Sub Timer2_Timer()
If Counter = 100 Then
 Timer1.Enabled = False
 Form7.Visible = True
 Else
  Counter = Counter + 1
 End If
End Sub

Private Sub txtcancel_Click()

End Sub

Private Sub txtPassword_keypress(KeyAscii As Integer)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0   ' Cancel the character.
      MsgBox "Please enter numbers only", vbCritical, "Error"
   End If
   
End Sub

Private Sub txtWithdraw_keypress(KeyAscii As Integer)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0   ' Cancel the character.
      MsgBox "Please enter numbers only", vbCritical, "Error"
   End If
   
End Sub

Private Sub withdrawal_Click()

 If ins = 0 Then
 MsgBox " Please Insert Your ATM card first before you begin transaction", vbCritical, " ATM Card not inserted"
 
 Else
    Frame2.Visible = True
   Frame4.Visible = False
   Label_Bal_Inquery.Visible = False
   lblSTransaction.Visible = False
   lblWthdrawAmount.Visible = True
   txtWithdraw.Visible = True
   lblEnterButton.Visible = False
   lblEnterPIN.Visible = False
   txtPassword.Visible = False
   lblInqueryName.Visible = False
   lblAccInfo.Visible = False
   lblBanner.Visible = False
   lblInqueryAccNum.Visible = False
   lblInqueryAvailBalance.Visible = False
   lblEnterButton.Visible = False
   lblSTransaction.Visible = False
   Timer_Font.Enabled = False
   Timer_Blink.Enabled = False
   Shape1.Visible = False
   cmdEnter.Enabled = False
   Command6.Enabled = True
      
      txtWithdraw = ""
   txtWithdraw.SetFocus

 End If
   
End Sub

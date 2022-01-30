VERSION 5.00
Begin VB.Form frmfinger 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ATM_Machine"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   3600
      Top             =   1080
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   3195
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   60
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   555
      Left            =   -3840
      TabIndex        =   6
      Top             =   3120
      Width           =   9300
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   8355
      Left            =   -3840
      TabIndex        =   5
      Top             =   -360
      Width           =   60
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   8355
      Left            =   5400
      TabIndex        =   4
      Top             =   -360
      Width           =   60
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "CAPTURE FINGERPRINT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5460
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1215
      Left            =   1920
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblBlink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place your fingerprint on the Scanner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
End
Attribute VB_Name = "frmfinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdProceed_Click()
formMain.Register
End Sub

Private Sub Form_Activate()
Load formMain
End Sub

Private Sub Timer1_Timer()
lblBlink.Visible = Not lblBlink.Visible
End Sub


VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   7500
   ClientTop       =   4860
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   360
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9600
      Top             =   3360
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "please wait......"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   720
      TabIndex        =   8
      Top             =   3360
      Width           =   5445
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7440
      Shape           =   1  'Square
      Top             =   3480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7680
      Shape           =   1  'Square
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   8040
      Shape           =   1  'Square
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   1905
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mrs. Dada O.M"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   2400
      Width           =   3390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ajayi Oluwaremilekun Racheal"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   4995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HND/12/COM/FT/078"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   3645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   9600
      TabIndex        =   1
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00004080&
      BackStyle       =   0  'Transparent
      Caption         =   "Implementation of Bio-Authentication in ATM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   480
      Left            =   705
      TabIndex        =   0
      Top             =   600
      Width           =   8925
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Private Sub Form_Load()

Timer2.Enabled = True
'Timer2.Interval = 50

End Sub



Private Sub Timer1_Timer()

   If ProgressBar1 = 100 Then
     frmSplash.Visible = False
     Timer1.Enabled = False
     frmmatric.Visible = True
       

   Else
       ProgressBar1 = ProgressBar1 + 1
       Label2.Caption = "% " & ProgressBar1
   End If
   
End Sub

Private Sub Timer2_Timer()
If (Shape1.Visible = True) And (Shape2.Visible = False) Then
    Shape2.Visible = True

ElseIf (Shape1.Visible = True) And (Shape2.Visible = True) And (Shape3.Visible = False) Then
    Shape3.Visible = True

ElseIf (Shape1.Visible = True) And (Shape2.Visible = True) And (Shape3.Visible = True) Then
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
Else
    Shape1.Visible = True
End If

k = k + 1

If k = 50 Then
    Timer2.Enabled = False
    Unload Me
    MDIForm1.Visible = True
End If
End Sub

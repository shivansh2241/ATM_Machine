VERSION 5.00
Begin VB.Form frmmatric 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   4080
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   3360
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   2880
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   2400
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   1440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PreSS ANY KEY TO CONTINUE......"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   4080
      Width           =   7575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   420
      Left            =   5640
      TabIndex        =   6
      Top             =   3360
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   420
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   75
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   3645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iya Ibeji Klass Mi"
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
      Height          =   420
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   2745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name and Matri No."
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   2730
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   3390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised By:"
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
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   0
      Picture         =   "frmmatric.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "frmmatric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
frmmatric.Visible = False
MDIForm1.Visible = True
End Sub

Private Sub Form_Load()
Label1.Top = 4680
Label2.Top = 4680
Label3.Top = 4680
Label4.Top = 4680
Label5.Top = 4680
Label6.Top = 4680
Label7.Top = 4680
Label8.Visible = False
End Sub


Private Sub Timer1_Timer()
Label1.Top = Label1.Top - 100
If Label1.Top <= 1440 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Label2.Top = Label2.Top - 100
If Label2.Top <= 1920 Then
Timer3.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Label3.Top = Label3.Top - 100
If Label3.Top <= 1440 Then
Timer4.Enabled = True
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Label4.Top = Label4.Top - 100
If Label4.Top <= 1920 Then
Timer5.Enabled = True
Timer4.Enabled = False
End If

End Sub

Private Sub Timer5_Timer()
Label5.Top = Label5.Top - 100
If Label5.Top <= 2400 Then
Timer6.Enabled = True
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
Label6.Top = Label6.Top - 100
If Label6.Top <= 2880 Then
Timer7.Enabled = True
Timer6.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
Label7.Top = Label7.Top - 100
If Label7.Top <= 3360 Then
Timer8.Enabled = True
Timer7.Enabled = False
End If
End Sub

Private Sub Timer8_Timer()
Label8.Visible = Not Label8.Visible
End Sub

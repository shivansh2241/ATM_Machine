VERSION 5.00
Begin VB.Form frmaccountcreated 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Account Created "
   ClientHeight    =   4620
   ClientLeft      =   3570
   ClientTop       =   4380
   ClientWidth     =   9555
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      MouseIcon       =   "frmaccountcreated.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1080
   End
   Begin VB.TextBox txtatmnumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   6975
   End
   Begin VB.TextBox txtatmpin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtaccountnumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtaccountname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   6975
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
      Height          =   4635
      Left            =   9480
      TabIndex        =   15
      Top             =   0
      Width           =   60
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
      Height          =   4635
      Left            =   0
      TabIndex        =   14
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
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   4440
      Width           =   9660
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "REGISTRATION SLIP"
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
      TabIndex        =   12
      Top             =   0
      Width           =   9660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: ATM Pin can be changed on the ATM Machine for security purpose"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   3720
      Width           =   9390
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "     Kindly copy this in a piece of paper ..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Account Information "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3810
      TabIndex        =   8
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ATM Default Pin:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   3
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ATM Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   570
      TabIndex        =   2
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Account  Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   405
      TabIndex        =   1
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   0
      Top             =   2520
      Width           =   1530
   End
End
Attribute VB_Name = "frmaccountcreated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdok_Click()
Unload Me
End Sub

Private Sub Form_Load()

'On Error GoTo err

Network

sql = "select * from New_Account where Account_Number='" & Keep & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
txtaccountname.Text = !Surname & " " & !Othernames
txtaccountnumber.Text = !Account_Number
txtatmpin.Text = !ATM_Pin
txtatmnumber.Text = !ATM_Number
End With
Exit Sub
err:
MsgBox "No Account", vbCritical
End Sub

Private Sub Timer1_Timer()
Label7.Left = Label7.Left - 70
    If Label7.Left <= -9480 Then
        Label7.Left = 9600
    End If
End Sub

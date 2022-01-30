VERSION 5.00
Begin VB.Form frmdeposit 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Deposit  Slip"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   9300
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      MouseIcon       =   "frmdeposit.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5520
      Top             =   3720
   End
   Begin VB.TextBox txttime 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtdate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox txtabalance 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2880
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      MouseIcon       =   "frmdeposit.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtdeposit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtaccountnumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtaccountname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   6735
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
      Left            =   9240
      TabIndex        =   19
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
      Height          =   8355
      Left            =   0
      TabIndex        =   18
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
      Left            =   0
      TabIndex        =   17
      Top             =   5400
      Width           =   9300
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "DEPOSIT FORM ENTRY"
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
      TabIndex        =   16
      Top             =   0
      Width           =   9300
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Deposited:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1665
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Deposited:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   1650
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Balance:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   1725
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Balance:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1545
   End
End
Attribute VB_Name = "frmdeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsearch_Click()

Dim fname As String
On Error GoTo Ree
fname = App.Path & "\pass\"
 
 If txtaccountnumber.Text = Empty Then
    MsgBox "Account Number is required", vbCritical, "Withdrawal"
    txtaccountnumber.SetFocus
 Exit Sub
 End If
 Network
 
sql = "select * from New_Account where Account_Number='" & Trim(txtaccountnumber.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

cmdsearch.Enabled = False
With rec
txtaccountname.Text = !Surname & " " & !Othernames
txtbalance.Text = !balance
txtabalance.Text = !Available_Balance
Image1.Picture = LoadPicture(fname & !Passport)
pin = !ATM_Pin
End With
Exit Sub
Ree:
MsgBox "There was a problem loading data", vbCritical

End Sub

Private Sub Command1_Click()
   
   
    For Each control In Me
        If TypeOf control Is TextBox Then
            If control.Text = Empty Then
                MsgBox "Deposit Slip is not Properly Filled", vbCritical, "Error"
            Exit Sub
            End If
        End If
    Next

    
Network

sql = "select * from New_Account where Account_Number='" & Trim(txtaccountnumber.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec

If txtaccountnumber.Text = !Account_Number Then
!Last_Deposit = txtdeposit.Text
!balance = Val(!balance) + txtdeposit.Text
!Available_Balance = Val(!balance) - 1000
.Update
End If
End With

Network

rec.Open "Daily_Deposit", con, adOpenDynamic, adLockOptimistic

With rec
.AddNew
!Account_Number = txtaccountnumber.Text
!ATM_Pin = Trim(pin)
!Account_Name = txtaccountname.Text
!Amount_Deposited = txtdeposit.Text
!Date_Deposited = txtdate.Text
!Time_Deposited = txttime.Text
.Save
End With

MsgBox "Depositing made successfully", vbInformation, "Deposit"
 
For Each control In Me
    If TypeOf control Is TextBox Then
        control.Text = Empty
    End If
Next

Image1.Picture = LoadPicture()
End Sub

Private Sub Form_Load()
Me.Left = (Screen.width - Me.width) / 3
Me.Top = (Screen.width - Me.width) / 3
End Sub

Private Sub Timer1_Timer()
txttime.Text = Format(time, "hh:mm:ss AMPM")
txtdate.Text = Format(Date, "dd/mm/yyyy")
End Sub

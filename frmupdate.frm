VERSION 5.00
Begin VB.Form frmupdate 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ACCOUNT UPDATE"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12675
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MouseIcon       =   "frmupdate.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MouseIcon       =   "frmupdate.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtreligion 
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
      Left            =   7440
      TabIndex        =   22
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtothername 
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
      Left            =   7440
      TabIndex        =   21
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txttype 
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
      Left            =   7440
      TabIndex        =   19
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox txtmarital 
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
      Left            =   2040
      TabIndex        =   18
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox txtpno 
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
      Left            =   7440
      TabIndex        =   17
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtocc 
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
      Left            =   2040
      TabIndex        =   16
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtDOB 
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
      Left            =   7440
      TabIndex        =   14
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox txtsex 
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
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtsurname 
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
      Left            =   2040
      TabIndex        =   12
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MouseIcon       =   "frmupdate.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtaccountnumber 
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
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label7 
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
      Height          =   6075
      Left            =   12600
      TabIndex        =   28
      Top             =   0
      Width           =   180
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
      Height          =   6075
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   60
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
      Height          =   195
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   14460
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "UPDATE ACCOUNT"
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
      TabIndex        =   25
      Top             =   0
      Width           =   12780
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5385
      TabIndex        =   20
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   10560
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Of Account:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5385
      TabIndex        =   11
      Top             =   2760
      Width           =   1875
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   4200
      Width           =   1680
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   465
      TabIndex        =   9
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   465
      TabIndex        =   7
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   465
      TabIndex        =   6
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5385
      TabIndex        =   5
      Top             =   2040
      Width           =   1470
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   465
      TabIndex        =   4
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   465
      TabIndex        =   3
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   2100
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsearch_Click()

On Error GoTo err

Dim fname As String

fname = App.Path & "\pass\"

If txtaccountnumber.Text = Empty Then
    MsgBox "Account Number is required for this operation", vbCritical, "Account Update"
    Exit Sub
End If

Network

sql = " select * from New_Account Where Account_Number='" & Trim(txtaccountnumber.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
txtsurname.Text = !Surname
txtothername.Text = !Othernames
txtsex.Text = !Sex
txtDOB.Text = !Date_Of_Birth
txtaddress.Text = !Address
txtpostaladdress.Text = !Postal_Address
txtocc.Text = !Occupation
txtmarital.Text = !Marital_Status
txtpno.Text = !Mobile_Phone
txttype.Text = !Account_Type
txtreligion.Text = !Religion
Image1.Picture = LoadPicture(fname & !Passport)
MsgBox "Record Found", vbInformation, "Account Creation"
End With
cmdsearch.Enabled = False
txtaccountnumber.Locked = True
cmdupdate.Enabled = True
Exit Sub


err:
    MsgBox "Record not Found", vbCritical, "Account Creation"
    
End Sub

Private Sub cmdupdate_Click()
On Error GoTo err

For Each control In Me
    If TypeOf control Is TextBox Then
        If control.Text = Empty Then
            MsgBox control.Name & " " & "is empty", vbCritical, "Account Update"
            Exit Sub
        End If
    End If
Next

Network

sql = "select * from New_Account where Account_Number='" & Trim(txtaccountnumber.Text) & "'"
rec.Open sql, con, adOpenDynamic, adLockOptimistic

With rec
If Trim(txtaccountnumber.Text) = !Account_Number Then
!Surname = txtsurname.Text
!Othernames = txtothername.Text
!Sex = txtsex.Text
!Date_Of_Birth = txtDOB.Text
!Address = txtaddress.Text
!Postal_Address = txtpostaladdress.Text
!Occupation = txtocc.Text
!Marital_Status = txtmarital.Text
!Mobile_Phone = txtpno.Text
!Account_Type = txttype.Text
!Religion = txtreligion.Text
!Passport = Trim(txtaccountnumber.Text) & ".jpg"
.Update
MsgBox "Record Updated Successfully", vbInformation, "Account Update"
End If
End With
For Each control In Me
    If TypeOf control Is TextBox Then
        control.Text = Empty
    End If
Next
Image1.Picture = LoadPicture()
Exit Sub

err:
    MsgBox "Record not Found", vbCritical, "Account Update"
End Sub

Private Sub Form_Load()
cmdupdate.Enabled = False
End Sub


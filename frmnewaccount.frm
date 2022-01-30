VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmnewaccount 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "New Account  [  Application Form  ]"
   ClientHeight    =   8730
   ClientLeft      =   8040
   ClientTop       =   5640
   ClientWidth     =   14475
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmnewaccount.frx":0000
      Left            =   3960
      List            =   "frmnewaccount.frx":000A
      TabIndex        =   43
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbaccounttype 
      Height          =   315
      Left            =   7680
      TabIndex        =   12
      Top             =   3240
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtkphone 
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
      Height          =   420
      Left            =   7740
      TabIndex        =   20
      Top             =   7800
      Width           =   2775
   End
   Begin VB.TextBox txtphone 
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
      Height          =   420
      Left            =   1680
      TabIndex        =   11
      Top             =   4440
      Width           =   3735
   End
   Begin VB.ComboBox cmbrelationship 
      Height          =   315
      Left            =   7740
      TabIndex        =   16
      Top             =   6360
      Width           =   4935
   End
   Begin VB.TextBox txtksurname 
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
      Height          =   420
      Left            =   1680
      TabIndex        =   13
      Top             =   5760
      Width           =   3855
   End
   Begin VB.TextBox txtkoccupation 
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
      Height          =   420
      Left            =   1710
      TabIndex        =   19
      Top             =   7800
      Width           =   3735
   End
   Begin VB.TextBox txtkothernames 
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
      Height          =   420
      Left            =   7740
      TabIndex        =   14
      Top             =   5760
      Width           =   4935
   End
   Begin VB.TextBox txtkaddress 
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
      Left            =   1710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   6840
      Width           =   3735
   End
   Begin VB.ComboBox cmbksex 
      Height          =   315
      Left            =   1710
      TabIndex        =   15
      Top             =   6360
      Width           =   2895
   End
   Begin VB.TextBox txtkpostaladdress 
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
      Height          =   660
      Left            =   7740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   6960
      Width           =   3975
   End
   Begin VB.CommandButton cmdupload 
      Caption         =   "UPLOAD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   31
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ComboBox cmbreligion 
      Height          =   315
      Left            =   7680
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cmbmarital 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox cmbyear 
      Height          =   315
      Left            =   9960
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbmonth 
      Height          =   315
      Left            =   8520
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbday 
      Height          =   315
      Left            =   7680
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cmbsex 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
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
      Height          =   1020
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtinitialdepo 
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
      Height          =   420
      Left            =   7680
      TabIndex        =   10
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtothernames 
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
      Height          =   420
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin VB.TextBox txtoccupation 
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
      Height          =   420
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtsurname 
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
      Height          =   420
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000B&
      Caption         =   "&Add Record"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label24 
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
      Left            =   14400
      TabIndex        =   47
      Top             =   360
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
      TabIndex        =   46
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label13 
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
      Height          =   75
      Left            =   0
      TabIndex        =   45
      Top             =   8640
      Width           =   14460
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "NEXT OF KIN RECORDS"
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
      TabIndex        =   44
      Top             =   5040
      Width           =   14460
   End
   Begin VB.Label Label22 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5760
      TabIndex        =   42
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Phone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5805
      TabIndex        =   41
      Top             =   7800
      Width           =   1515
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Phone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   40
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "APPLICANT RECORD ENTRY"
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
      TabIndex        =   39
      Top             =   0
      Width           =   14460
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relationship:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5805
      TabIndex        =   38
      Top             =   6360
      Width           =   1380
   End
   Begin VB.Label Label19 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   150
      TabIndex        =   37
      Top             =   6360
      Width           =   465
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   150
      TabIndex        =   36
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Label Label17 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   30
      TabIndex        =   35
      Top             =   7800
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Names:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5805
      TabIndex        =   34
      Top             =   5760
      Width           =   1485
   End
   Begin VB.Label Label15 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   150
      TabIndex        =   33
      Top             =   6720
      Width           =   945
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5805
      TabIndex        =   32
      Top             =   6840
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label12 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5760
      TabIndex        =   30
      Top             =   2160
      Width           =   945
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   29
      Top             =   3840
      Width           =   1515
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5760
      TabIndex        =   28
      Top             =   1560
      Width           =   1470
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initial  Deposit:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5760
      TabIndex        =   26
      Top             =   2640
      Width           =   1635
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   25
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Names:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   5760
      TabIndex        =   24
      Top             =   840
      Width           =   1485
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   23
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   22
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   90
      TabIndex        =   21
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "frmnewaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsave_Click()
On Error GoTo err

    For Each control In Me
        If TypeOf control Is TextBox Or TypeOf control Is ComboBox Then
            If control.Text = Empty Then
                MsgBox "one of the Field(s) Is Empty", vbCritical
            Exit Sub
            End If
        End If
    Next

 If Image1.Picture = LoadPicture("") Then
 MsgBox "Picture is Missing", vbCritical
Exit Sub
End If

    If cmbaccounttype.Text = "DOMICILIARY ACCOUNT" And Val(txtinitialdepo.Text) < 100000 Then
    MsgBox "Amount too small for account creation", vbCritical, "Account creation"
    Exit Sub
    End If
    
    If cmbaccounttype.Text = "FIXED ACCOUNT" And Val(txtinitialdepo.Text) < 20000 Then
    MsgBox "Amount too small for account creation", vbCritical, "Account creation"
    Exit Sub
    End If
    
    If cmbaccounttype.Text = "SAVINGS ACCOUNT" And Val(txtinitialdepo.Text) < 2000 Then
    MsgBox "Amount too small for account creation", vbCritical, "Account creation"
    Exit Sub
    End If
    
    If cmbaccounttype.Text = "CURRENT ACCOUNT" And Val(txtinitialdepo.Text) < 10000 Then
    MsgBox "Amount too small for account creation", vbCritical, "Account creation"
    Exit Sub
    End If
    
    If cmbaccounttype.Text = "JOINT ACCOUNT" And txtinitialdepo.Text < 50000 Then
    MsgBox "Amount too small for account creation", vbCritical, "Account creation"
    Exit Sub
    End If
    
    
Randomize Timer
accountnumber = Int((9999999999# - 1111111111 + 1) * Rnd) + 1111111111
Randomize Timer
atmnumber = Int((999999999999999# - 111111111111111# + 1) * Rnd) + 111111111111111#
Randomize Timer
atmpin = Int((9999 - 1111 + 1) * Rnd) + 1111
    bool = 1
    frmFinger.Show vbModal
Exit Sub

err:
MsgBox "There was a problem creating account", vbCritical, "Account creation"

End Sub

Private Sub cmdupload_Click()
CommonDialog1.Filter = "JPEG Passport (*.jpg)|*.jpg"
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Form_Load()
Me.Left = (Screen.width - Me.width) / 6
Me.Top = (Screen.width - Me.width) / 17
cmbsex.AddItem "MALE"
cmbsex.AddItem "FEMALE"

cmbksex.AddItem "MALE"
cmbksex.AddItem "FEMALE"

cmbreligion.AddItem "ISLAM"
cmbreligion.AddItem "CHRISTIANITY"
cmbreligion.AddItem "TRADITIONAL"

cmbrelationship.AddItem "BROTHER"
cmbrelationship.AddItem "SISTER"
cmbrelationship.AddItem "HUSBAND"
cmbrelationship.AddItem "WIFE"
cmbrelationship.AddItem "FATHER"
cmbrelationship.AddItem "MOTHER"

cmbaccounttype.AddItem "CURRENT ACCOUNT"
cmbaccounttype.AddItem "SAVINGS ACCOUNT"
cmbaccounttype.AddItem "JOINT ACCOUNT"
cmbaccounttype.AddItem "FIXED ACCOUNT"
cmbaccounttype.AddItem "DOMICILIARY ACCOUNT"

cmbmarital.AddItem "SINGLE"
cmbmarital.AddItem "MARRIED"
cmbmarital.AddItem "DIVORCED"

Dim year As Integer
Dim i As Integer
year = 1900
For i = i To 3000
year = year + 1
cmbyear.AddItem year
Next i

cmbmonth.AddItem "JANUARY"
cmbmonth.AddItem "FEBRUARY"
cmbmonth.AddItem "MARCH"
cmbmonth.AddItem "APRIL"
cmbmonth.AddItem "MAY"
cmbmonth.AddItem "JUNE"
cmbmonth.AddItem "JULY"
cmbmonth.AddItem "AUGUST"
cmbmonth.AddItem "SEPTEMBER"
cmbmonth.AddItem "OCTOBER"
cmbmonth.AddItem "NOVEMBER"
cmbmonth.AddItem "DECEMBER"

Dim day As Integer
Dim k As Integer
day = 0
For k = 1 To 31
day = day + 1
cmbday.AddItem day
Next k


End Sub


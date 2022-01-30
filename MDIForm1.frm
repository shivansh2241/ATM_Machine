VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFC0C0&
   Caption         =   " ATM Machine "
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   2011
      ButtonWidth     =   2725
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&New Account"
            Key             =   ""
            Object.ToolTipText     =   "New Account"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Deposit"
            Key             =   ""
            Object.ToolTipText     =   "Make Deposit"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Show ATM"
            Key             =   ""
            Object.ToolTipText     =   "Withdraw from ATM"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Update Account"
            Key             =   ""
            Object.ToolTipText     =   "Update Account Information"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&About"
            Key             =   ""
            Object.ToolTipText     =   "Information About The Developer"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Close"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "MDIForm1.frx":43907
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":43C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4EBD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":50725
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":158237
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":15A039
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":15BB8B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNew_Acct_Click()
frmnewaccount.Show
End Sub

Private Sub Command2_Click()
frmdeposit.Show
End Sub

Private Sub Command3_Click()
frmatm.Show
End Sub

Private Sub Command4_Click()
frmupdate.Show
End Sub

Private Sub Command5_Click()
frmabt.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1:
    frmnewaccount.Show
    frmdeposit.Visible = False
    Unload frmatm
    frmupdate.Visible = False
    frmabt.Visible = False
Case 2:
    frmdeposit.Show
    frmnewaccount.Visible = False
    Unload frmatm
    frmupdate.Visible = False
    frmabt.Visible = False
Case 3:
    frmatm.Show
    frmdeposit.Visible = False
    frmnewaccount.Visible = False
    frmupdate.Visible = False
    frmabt.Visible = False
Case 4:
    frmupdate.Show
    frmdeposit.Visible = False
    frmnewaccount.Visible = False
    Unload frmatm
    frmabt.Visible = False
Case 5:
    frmabt.Show
    frmdeposit.Visible = False
    frmnewaccount.Visible = False
    Unload frmatm
    frmupdate.Visible = False
Case 6:
    response = MsgBox("Do you want to terminate the application", vbQuestion + vbYesNo)
    If response = vbYes Then
        End
    Else
        Exit Sub
    End If
End Select
End Sub

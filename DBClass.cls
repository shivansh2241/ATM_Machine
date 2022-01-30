VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'GrFinger Sample
'(c) 2005 - 2010 Griaule Biometrics Ltda.
'http://www.griaulebiometrics.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Biometrics makes no representations concerning either the merchantability
'of this software or the suitability of this sample for any particular purpose.
'
'THIS SAMPLE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR
'IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
'IN NO EVENT SHALL GRIAULE BE LIABLE FOR ANY DIRECT, INDIRECT,
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'You can download the trial version of GrFinger directly from Griaule website.
'
'These notices must be retained in any copies of any part of this
'documentation and/or sample.
'
'-------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Database routines
' -----------------------------------------------------------------------------------

Option Explicit
Option Base 0
' Add template to database. Returns added template ID.
Public Function AddTemplate(ByRef template() As Byte) As Long
Network
    Dim DOB As String
    Dim ret As Integer, score As Long
    
    score = 0
    ' identify it
    ret = Identify(score)
    ' if template has not being previously added, adds it
     
    ' Set rec = New ADODB.Recordset
rec.Open "New_Account", con, adOpenDynamic, adLockOptimistic
With rec
.AddNew
!template = template
!Surname = frmnewaccount.txtsurname.Text
!Othernames = frmnewaccount.txtothernames.Text
!Sex = frmnewaccount.cmbsex.Text
!Date_Of_Birth = frmnewaccount.cmbday.Text & "-" & frmnewaccount.cmbmonth.Text & "-" & frmnewaccount.cmbyear.Text
!Address = frmnewaccount.txtaddress.Text
!Postal_Address = frmnewaccount.txtpostaladdress.Text
!Occupation = frmnewaccount.txtoccupation.Text
!Marital_Status = frmnewaccount.cmbmarital.Text
!Religion = frmnewaccount.cmbreligion.Text
!Initial_Deposit = frmnewaccount.txtinitialdepo.Text
!Mobile_Phone = frmnewaccount.txtphone.Text
!Account_Type = frmnewaccount.cmbaccounttype.Text
!Passport = accountnumber & ".jpg" 'txtsurname.Text & ".jpg"
!Kin_Surname = frmnewaccount.txtksurname.Text
!Kin_Othernames = frmnewaccount.txtkothernames.Text
!Kin_Sex = frmnewaccount.cmbksex.Text
!Relationship = frmnewaccount.cmbrelationship.Text
!Kin_Address = frmnewaccount.txtkaddress.Text
!Kin_Postal_Address = frmnewaccount.txtkpostaladdress.Text
!Kin_Occupation = frmnewaccount.txtkoccupation.Text
!Kin_Phone = frmnewaccount.txtkphone.Text
!Account_Number = Int(accountnumber)
!ATM_Number = Int(atmnumber)
!ATM_Pin = Int(atmpin)
!Last_Deposit = frmnewaccount.txtinitialdepo.Text
!balance = frmnewaccount.txtinitialdepo.Text
!Available_Balance = Val(frmnewaccount.txtinitialdepo.Text) - Val(1000)
.Save
End With

Dim fname As String
fname = App.Path & "\pass\"
FileCopy frmnewaccount.CommonDialog1.FileName, fname & accountnumber & ".jpg"

MsgBox "Account created successfully", vbInformation, "Account creation"
Dim control As control
For Each control In frmnewaccount
    If TypeOf control Is TextBox Or TypeOf control Is ComboBox Then
        control.Text = Empty
    End If
Next


frmnewaccount.Image1.Picture = LoadPicture("")
Keep = accountnumber
Unload frmnewaccount
Unload frmFinger
frmaccountcreated.Show
    
    
    ' return ID
    AddTemplate = rec("ID")
    'retholder = rec("ID")
       
End Function

' Created By Me
Public Function CheckTemplate(ByRef template() As Byte) As Long
    Dim Nen As String
    ' Get query response
    Set rec = New ADODB.Recordset
    Nen = "select * from New_Account where template = " & template
    rec.Open Nen, con, adOpenDynamic, adLockReadOnly
    ' No results?
    If rec.EOF And rec.BOF Then
        Dim tptVazio(0) As Byte
        rec.Close
        getTemplate = tptVazio
    Else
        DB.AddTemplate (template.tpt)
    End If
End Function
' Returns a ADODB.recordset with all enrolled templates from database.
Public Function getTemplates() As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.Open "select * from New_Account", con, adOpenDynamic, adLockReadOnly
    Set getTemplates = rec
End Function

' returns template with the supplied ID.
Public Function getTemplate(id As Long) As Byte()
    ' Get query response
    Set rec = New ADODB.Recordset
    rec.Open "select * from New_Account where ID = " & id, con, adOpenDynamic, adLockReadOnly
    ' No results?
    If rec.EOF And rec.BOF Then
        Dim tptVazio(0) As Byte
        rec.Close
        getTemplate = tptVazio
    Else
        ' Deserialize template and return it
        Dim tpt() As Byte
        tpt = rec("template")
        rec.Close
        getTemplate = tpt
    End If
End Function

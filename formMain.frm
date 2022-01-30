VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A9995C7C-77BF-4E27-B581-A4B5BBD90E50}#1.0#0"; "GrFingerX.dll"
Begin VB.Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GrFingerX - Sample application - Visual Basic 6.0"
   ClientHeight    =   7755
   ClientLeft      =   3735
   ClientTop       =   1455
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   517
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ckBoxAutoExtract 
      Caption         =   "Auto Extract"
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   6240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton btExtract 
      Caption         =   "Extract template"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ListBox lbLog 
      Height          =   1035
      Left            =   120
      OLEDragMode     =   1  'Automatic
      TabIndex        =   6
      Top             =   6600
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmap (*.bmp)|*.bmp"
   End
   Begin VB.CommandButton btClearDB 
      Caption         =   "Clear database"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton btClearLog 
      Caption         =   "Clear log"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton btVerify 
      Caption         =   "Verify"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton btIdentify 
      Caption         =   "Identify"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton btEnroll 
      Caption         =   "Enroll"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox ckBoxAutoIdentify 
      Caption         =   "Auto identify"
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   5880
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin GrFingerXLibCtl.GrFingerXCtrl GrFingerXCtrl1 
      Left            =   6480
      OleObjectBlob   =   "formMain.frx":0000
      Top             =   1200
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fixed Single
      Height          =   6375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5775
   End
   Begin VB.Menu menuImage 
      Caption         =   "&Image"
      Begin VB.Menu menuImgSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu menuImgFromFile 
         Caption         =   "&Load From File..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options..."
   End
   Begin VB.Menu menuVersion 
      Caption         =   "&Version"
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
'GrFinger Sample
'(c) 2005 Griaule Tecnologia Ltda.
'http://www.griaule.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Tecnologia makes no representations concerning either the merchantability
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
' GUI routines: main form
' -----------------------------------------------------------------------------------

Option Explicit

' Application startup code
Private Sub Form_Load()
    Dim err As Integer
    
    ' Initialize GrFingerX Library
    err = InitializeGrFinger()
    ' Print result in log
    If err < 0 Then
        writeError (err)
        Exit Sub
    Else
        writeLog ("**GrFingerX Initialized Successfull**")
    End If
End Sub

' Application finalization code
Private Sub Form_Terminate()
    Call FinalizeGrFinger
End Sub

' Add a fingerprint to database
Private Sub btEnroll_Click()
    Dim id As Integer
    
    ' add fingerprint
    id = Enroll
    ' write result to log
    If id >= 0 Then
        writeLog ("Fingerprint enrolled with id = " & id)
    Else
        writeLog ("Error: Fingerprint not enrolled")
    End If
End Sub
Sub Register()
Dim id As Integer
    
    ' add fingerprint
    id = Enroll
    ' write result to log
    If id >= 0 Then
        writeLog ("Fingerprint enrolled with id = " & id)
    Else
        writeLog ("Error: Fingerprint not enrolled")
    End If
End Sub
Public Sub btIdentify_Click()
'On Error GoTo Errs
    Network
    Dim ret As Integer, score As Long
    score = 0
    
    ' identify it
    ret = Identify(score)
    
    ' write result to log
    If ret > 0 Then
        writeLog ("Fingerprint identified. ID = " & ret & ". Score = " & score & ".")
        Call PrintBiometricDisplay(True, GR_DEFAULT_CONTEXT)
        sql = "select * from New_Account where ID =" & ret
        rec.Open sql, con, adOpenDynamic, adLockOptimistic
        With rec
        If !id = ret Then
            If bool = 1 Then
         frmFinger.imgfinger.Picture = img.Picture
         frmFinger.cmdProceed.Enabled = True
            ElseIf bool = 2 Then
            frmatm.cmdnewaccount.Visible = True
            frmatm.cmddeposit.Visible = True
            frmatm.cmdBalInqury.Visible = True
            frmatm.withdrawal.Visible = True
            frmatm.cancel.Visible = True
            frmatm.cmd_TakeCard.Visible = True
            frmatm.lblSTransaction.Visible = True
            frmatm.Command3.Visible = True
            pin = !ATM_Pin
            frmFinger.imgfinger.Picture = img.Picture
            Unload frmFinger
            End If
            
            Exit Sub
        End If
        End With
    ElseIf ret = 0 Then
        writeLog ("Fingerprint not Found.")
        If bool = 1 Then
        frmFinger.imgfinger.Picture = img.Picture
        frmFinger.cmdProceed.Enabled = True
        ElseIf bool = 2 Then
        MsgBox "The Fingerprint does not match", vbCritical, "Not Match"
        Exit Sub
        End If
    Else
        writeError (ret)
    End If
    Exit Sub
End Sub

' Check a fingerprint
Private Sub btVerify_Click()
    Dim id As Integer
    Dim ret As Integer
    Dim score As Long
    Dim sID As String

    ' ask target fingerprint ID
    score = 0
    sID = InputBox("Enter the ID to verify", "Verify", "")
    If sID <> "" Then
        ' compare fingerprints
        ret = Verify(Val(sID), score)
        ' write result to log
        If ret < 0 Then
            writeError (ret)
        ElseIf ret = GR_NOT_MATCH Then
            writeLog ("Did not match with score = " & score)
        Else
            writeLog ("Matched with score = " & score)
            ' if they match, display matching minutiae/segments/directions
            Call PrintBiometricDisplay(True, GR_DEFAULT_CONTEXT)
        End If
    End If
    
End Sub

' Extract a template from a fingerprint image
Public Sub btExtract_Click()
    Dim ret As Integer

    ' extract template
    ret = ExtractTemplate()
    ' write template quality to log
    If ret = GR_BAD_QUALITY Then
        writeLog ("Template extracted successfully. Bad quality.")
    ElseIf ret = GR_MEDIUM_QUALITY Then
        writeLog ("Template extracted successfully. Medium quality.")
    ElseIf ret = GR_HIGH_QUALITY Then
        writeLog ("Template extracted successfully. High quality.")
    End If
    If ret >= 0 Then
        ' if no error, display minutiae/segments/directions into the image
        Call PrintBiometricDisplay(True, GR_NO_CONTEXT)
        ' enable operations we can do over extracted template
        btExtract.Enabled = False
        btEnroll.Enabled = True
        btIdentify.Enabled = True
        btVerify.Enabled = True
    Else
        ' write error to log
        writeError (ret)
    End If
End Sub

' Clear database
Private Sub btClearDB_Click()
    ' clear database
    DB.clearDB
    ' write result to log
    writeLog ("Database is clear...")
End Sub

' Clear log
Private Sub btClearLog_Click()
    lbLog.Clear
End Sub

' Save fingerprint image to a file
Private Sub menuImgSave_Click()
    ' we need an image
    If raw.height < 1 Or raw.width < 1 Then
        MsgBox "There is no image to save."
        Exit Sub
    End If
        
    ' open "save" dialog
    CommonDialog.Filter = "BMP files (*.bmp)|*.bmp|All files (*.*)|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.FileName = ""
    CommonDialog.ShowSave
    
    ' Save image.
    If Not CommonDialog.CancelError And CommonDialog.FileName <> "" Then
        If formMain.GrFingerXCtrl1.CapSaveRawImageToFile(raw.img, raw.width, raw.height, CommonDialog.FileName, GRCAP_IMAGE_FORMAT_BMP) <> GR_OK Then
            writeLog ("Fail to save the file.")
        End If
    End If
End Sub

' Load a fingerprint image from a file
Private Sub menuImgFromFile_Click()
    ' open "load" dialog
    CommonDialog.Filter = "BMP files (*.bmp)|*.bmp|All files (*.*)|*.*"
    CommonDialog.FilterIndex = 1
    CommonDialog.FileName = ""
    CommonDialog.ShowOpen
    
    ' load image
    If Not CommonDialog.CancelError And CommonDialog.FileName <> "" Then
       Dim res As Long
       ' Getting resolution.
        res = Val(InputBox("Enter the resolution of the selected image", "Resolution"))
        ' Checking if action was canceled, no value or an invalid value was entered.
        If res <> 0 Then
            If GrFingerXCtrl1.CapLoadImageFromFile(CommonDialog.FileName, res) <> GR_OK Then
                writeLog ("Fail to load the file.")
            End If
        End If
    End If
End Sub

' Open "Options" window
Private Sub menuOptions_Click()
    Dim ret As Integer
    Dim thresholdId As Long
    Dim rotationMaxId As Long
    Dim thresholdVr As Long
    Dim rotationMaxVr As Long
    Dim minutiaeColor As Long
    Dim minutiaeMatchColor As Long
    Dim segmentsColor As Long
    Dim segmentsMatchColor As Long
    Dim directionsColor As Long
    Dim directionsMatchColor As Long
    Dim ok As Boolean

    Do
        ' get current identification/verification parameters
        GrFingerXCtrl1.GetIdentifyParameters thresholdId, rotationMaxId, GR_DEFAULT_CONTEXT
        GrFingerXCtrl1.GetVerifyParameters thresholdVr, rotationMaxVr, GR_DEFAULT_CONTEXT
        ' set current identification/verification parameters on options form
        Call formOptions.setParameters(thresholdId, rotationMaxId, thresholdVr, rotationMaxVr)

        ok = True
        ' show form with match, display and colors options
        ' and get new parameters
        If Not (formOptions.getParameters(thresholdId, rotationMaxId, thresholdVr, rotationMaxVr, _
                    minutiaeColor, minutiaeMatchColor, segmentsColor, segmentsMatchColor, directionsColor, directionsMatchColor)) Then
            Exit Sub
        End If
        If ((thresholdId < GR_MIN_THRESHOLD) Or _
         (thresholdId > GR_MAX_THRESHOLD) Or _
         (rotationMaxId < GR_ROT_MIN) Or _
         (rotationMaxId > GR_ROT_MAX)) Then
            MsgBox ("Invalid identify parameters values!")
            ok = False
        End If
        If (thresholdVr < GR_MIN_THRESHOLD Or _
         thresholdVr > GR_MAX_THRESHOLD Or _
         rotationMaxVr < GR_ROT_MIN Or _
         rotationMaxVr > GR_ROT_MAX) Then
            MsgBox ("Invalid verify parameters values!")
            ok = False
        End If
        ' set new identification parameters
        If ok Then
            ret = GrFingerXCtrl1.SetIdentifyParameters(thresholdId, rotationMaxId, GR_DEFAULT_CONTEXT)
            ' error?
            If ret = GR_DEFAULT_USED Then
                MsgBox ("Invalid identify parameters values. Default values will be used.")
                ok = False
            End If
            ' set new verification parameters
            ret = GrFingerXCtrl1.SetVerifyParameters(thresholdVr, rotationMaxVr, GR_DEFAULT_CONTEXT)
            ' error?
            If ret = GR_DEFAULT_USED Then
                MsgBox ("Invalid verify parameters values. Default values will be used.")
                ok = False
            End If
            ' if everything ok
            If ok Then
                ' accept new parameters
                Call formOptions.AcceptChanges
                ' set new colors
                GrFingerXCtrl1.SetBiometricDisplayColors minutiaeColor, minutiaeMatchColor, segmentsColor, segmentsMatchColor, directionsColor, directionsMatchColor
                Exit Sub
            End If
        End If
    Loop
End Sub

' Display GrFinger version
Private Sub menuVersion_Click()
    Call MessageVersion
End Sub

' -----------------------------------------------------------------------------------
' GrFingerX events
' -----------------------------------------------------------------------------------

' A fingerprint reader was plugged on system
Private Sub GrFingerXCtrl1_SensorPlug(ByVal idSensor As String)
    writeLog ("Sensor: " & idSensor & ". Event: Plugged.")
    GrFingerXCtrl1.CapStartCapture (idSensor)
End Sub

' A fingerprint reader was unplugged from system
Private Sub GrFingerXCtrl1_SensorUnplug(ByVal idSensor As String)
    writeLog ("Sensor: " & idSensor & ". Event: Unplugged.")
    GrFingerXCtrl1.CapStopCapture (idSensor)
End Sub

' A finger was placed on reader
Private Sub GrFingerXCtrl1_FingerDown(ByVal idSensor As String)
    writeLog ("Sensor: " & idSensor & ". Event: Finger Placed.")
End Sub

' A finger was removed from reader
Private Sub GrFingerXCtrl1_FingerUp(ByVal idSensor As String)
    writeLog ("Sensor: " & idSensor & ". Event: Finger removed.")
End Sub

' An image was acquired from reader
Private Sub GrFingerXCtrl1_ImageAcquired(ByVal idSensor As String, ByVal width As Long, ByVal height As Long, rawImage As Variant, ByVal res As Long)
    ' Copying aquired image
    With raw
        .img = rawImage
        .height = height
        .width = width
        .res = res
    End With

    ' Signaling that an Image Event occurred.
    writeLog ("Sensor: " & idSensor & ". Event: Image captured.")
    ' display fingerprint image
    Call PrintBiometricDisplay(False, GR_DEFAULT_CONTEXT)
    
    ' now we have a fingerprint, so we can extract template
    formMain.btExtract.Enabled = True
    formMain.btEnroll.Enabled = False
    formMain.btIdentify.Enabled = False
    formMain.btVerify.Enabled = False
    
    ' extracting template from image
    If formMain.ckBoxAutoExtract.Value Then
        formMain.btExtract_Click
        
        ' identify fingerprint
        If formMain.ckBoxAutoIdentify.Value Then
            formMain.btIdentify_Click
        End If
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' set date to april 05 2013
    '''''''''''''''''''''''''''''''''''''''''''''''''
    

End Sub


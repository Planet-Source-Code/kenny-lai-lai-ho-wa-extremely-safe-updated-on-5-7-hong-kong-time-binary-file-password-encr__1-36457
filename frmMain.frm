VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "Binary Encryptor"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Frame Frame2 
      Caption         =   "Decrypt"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "&Decrypt"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtDPassword 
         Appearance      =   0  '¥­­±
         Height          =   285
         IMEMode         =   3  '¼È¤î
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtDFilename 
         Appearance      =   0  '¥­­±
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdDOpen 
         Caption         =   "Open"
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbldPercent 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "0%"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt"
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "&Encrypt"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtEPassword 
         Appearance      =   0  '¥­­±
         Height          =   285
         IMEMode         =   3  '¼È¤î
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtEFilename 
         Appearance      =   0  '¥­­±
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdEOpen 
         Caption         =   "&Open"
         Height          =   285
         Left            =   3000
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblePercent 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackStyle       =   0  '³z©ú
         Caption         =   "0%"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim eSource As String, eDestination As String, ePassword As String
Dim dSource As String, dDestination As String, dPassword As String

Dim WithEvents EN As clsBinaryEncryptor
Attribute EN.VB_VarHelpID = -1

Private Sub cmdEncrypt_Click()
Dim b As Boolean
Screen.MousePointer = vbHourglass
    b = EN.EncryptFile(eSource, eDestination, IIf(txtEPassword.Text = "", "default", txtEPassword.Text))

If b = True Then
    MsgBox "The file is encrypted successfully." & vbCrLf & "Please save your password.", vbInformation, "BinaryEncryptor"
Else
    MsgBox "Error occured while encrypting the file. Please contact the software developer.", vbCritical, "Kenny Lai, assw@hkem.com"
End If

Screen.MousePointer = 0
End Sub

Private Sub cmdEOpen_Click()

With cd1
    .CancelError = True
    .Filter = "All Files *.*|*.*"
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo OpenError
    .DialogTitle = "Open a file to encrypt."
    .ShowOpen
    eSource = .Filename
    
    On Error GoTo SaveError
    .DialogTitle = "Save the encrypted file."
    .ShowSave
    txtEFilename.Text = .Filename
    eDestination = .Filename
    
    txtEPassword.SetFocus
    txtEPassword.SelStart = 0
    txtEPassword.SelLength = Len(txtEPassword.Text)
    
    cmdEncrypt.Enabled = True
    
End With

Exit Sub
OpenError:
SaveError:

End Sub

Private Sub cmdDecrypt_Click()
Screen.MousePointer = vbHourglass
Dim b As Boolean
    b = EN.DecryptFile(dSource, dDestination, IIf(txtDPassword.Text = "", "default", txtDPassword.Text))

If b = True Then
    Dim m As Integer
    m = MsgBox("The file is decrypted successfully." & vbCrLf & "Do you want to view the file now?", vbInformation + vbYesNo, "BinaryEncryptor")
    If m = vbYes Then Browser dDestination, Me.hwnd
Else
    MsgBox "Error occured while decrypting the file. Please contact the software developer.", vbCritical, "Kenny Lai, assw@hkem.com"
End If
Screen.MousePointer = 0
End Sub

Private Sub cmdDOpen_Click()

With cd1
    .CancelError = True
    .Filter = "All Files *.*|*.*"
    .Flags = cdlOFNFileMustExist
    
    On Error GoTo OpenError
    .DialogTitle = "Open a file to decrypt."
    .ShowOpen
    dSource = .Filename
    
    On Error GoTo SaveError
    .DialogTitle = "Save the decrypted file."
    .ShowSave
    txtDFilename.Text = .Filename
    dDestination = .Filename
    
    txtDPassword.SetFocus
    txtDPassword.SelStart = 0
    txtDPassword.SelLength = Len(txtDPassword.Text)
    
    cmdDecrypt.Enabled = True
    
End With

Exit Sub
OpenError:
SaveError:

End Sub

Private Sub EN_DecryptProgress(Progress As Long, ProgressMax As Long)
lbldPercent.Caption = Round(Progress / ProgressMax * 100, 2) & "%"
End Sub

Private Sub EN_EncryptProgress(Progress As Long, ProgressMax As Long)
lblePercent.Caption = Round(Progress / ProgressMax * 100, 2) & "%"
End Sub

Private Sub Form_Load()
Set EN = New clsBinaryEncryptor
End Sub

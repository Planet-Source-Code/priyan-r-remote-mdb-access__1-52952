VERSION 5.00
Object = "{BB006AF6-6CFD-41CA-8D2E-CA332DAFEF3C}#1.1#0"; "remotemdb.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan 'S Remote MDB User Registration Sample"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin mdbaccess.remotemdbaccess remotemdbaccess1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   3175
      _ExtentY        =   344
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Login"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3855
      Begin VB.TextBox txtloginusername 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtloginpassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "User name "
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pass Word"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registration"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
      Begin VB.CommandButton cmdregister 
         Caption         =   "Register ME"
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtpassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtusername 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pass Word"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User name "
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTANT:Put the userregsample.mdb in the root of web server"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a sample that introduce the use of this Active-X Control"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)



Private Sub cmdlogin_Click()
On Error GoTo ext:
If Trim(Me.txtloginusername.Text) = "" Then
    MsgBox "Type User Name", vbCritical
    txtloginusername.SetFocus
    Exit Sub
ElseIf Trim(Me.txtloginpassword.Text) = "" Then
    MsgBox "Type Password", vbCritical
    txtloginpassword.SetFocus
    Exit Sub
End If
cmdlogin.Enabled = False
    With Me.remotemdbaccess1
        .executequery "Select * from users where userid='" & txtloginusername.Text & "'" & " and password='" & txtloginpassword.Text & "'"
        If .recordcount = 0 Then
            MsgBox "Login Unsuccessfull Check you password or username", vbCritical
        Else
            MsgBox "Login Successfull", vbInformation
        End If
    End With
cmdlogin.Enabled = True
Exit Sub
ext:
MsgBox Err.Description, vbCritical
cmdlogin.Enabled = True
End Sub

Private Sub cmdregister_Click()
On Error GoTo ext:
If Trim(Me.txtusername.Text) = "" Then
    MsgBox "Type User Name", vbCritical
    txtusername.SetFocus
    Exit Sub
ElseIf Trim(Me.txtpassword.Text) = "" Then
    MsgBox "Type Password", vbCritical
    txtpassword.SetFocus
    Exit Sub
End If
cmdregister.Enabled = False
    With Me.remotemdbaccess1
        .executequery "Select * from users where userid='" & txtusername.Text & "'"
        If .recordcount = 0 Then
            .addnew "userid=" & txtusername.Text & "|~|password=" & txtpassword.Text
            MsgBox "Registered Successfully", vbInformation
        Else
            MsgBox "User Name '" & txtusername.Text & "' Allready Registered", vbInformation
            txtusername.SetFocus
        End If
    End With
cmdregister.Enabled = True
Exit Sub
ext:
MsgBox Err.Description, vbCritical
cmdregister.Enabled = True
End Sub

Private Sub Form_Load()
'change the following according to the address of you server
Me.remotemdbaccess1.scripturl = "http://localhost/remotemdb.asp"
Me.remotemdbaccess1.mdbfile = "userregsample.mdb"


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'if not run within ide call exit process else the app will take a little time to close.
If App.LogMode <> 0 Then
    ExitProcess 1
End If
End Sub

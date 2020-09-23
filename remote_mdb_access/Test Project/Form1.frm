VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BB006AF6-6CFD-41CA-8D2E-CA332DAFEF3C}#1.1#0"; "remotemdb.ocx"
Begin VB.Form form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's Remote MDB Active-X Test"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin mdbaccess.remotemdbaccess remotemdbaccess1 
      Left            =   3240
      Top             =   2280
      _ExtentX        =   3175
      _ExtentY        =   344
   End
   Begin VB.TextBox txtserverscripturl 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtmdbpassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   15
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdaddnew 
      Caption         =   "Add New"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtmdbfile 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Text            =   "a.mdb"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtsql 
      Height          =   495
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtpos 
      Height          =   405
      Left            =   6000
      TabIndex        =   5
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtrecordstofetch 
      Height          =   405
      Left            =   6000
      TabIndex        =   2
      Text            =   "0"
      Top             =   1320
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdexecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblabout 
      BackColor       =   &H8000000B&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblvote 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Vote For This Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   1650
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   480
   End
   Begin VB.Label Label6 
      Caption         =   "Server Script URL"
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MS-Access File"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "SQL"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Pos"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Records To Fetch"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long





Private Sub cmdaddnew_Click()
On Error GoTo ext:
Dim str$
str = InputBox("Enter Values Format: field1=value|~|field2=value|~|field3=value")
If Trim(str) = "" Then Exit Sub
Me.remotemdbaccess1.addnew str
list
Exit Sub
ext:
MsgBox Err.Description, vbCritical, "Error!"

End Sub

Private Sub cmdupdate_Click()
On Error GoTo ext:
If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
Dim str$
str = InputBox("Enter Values Format: field1=value|~|field2=value|~|field3=value")
If Trim(str) = "" Then Exit Sub
Me.remotemdbaccess1.Update Me.ListView1.SelectedItem.Index, str
list
Exit Sub
ext:
MsgBox Err.Description, vbCritical, "Error!"
End Sub

Private Sub cmdexecute_Click()
On Error GoTo ext:
Dim i%, str() As String, litem As ListItem
Me.remotemdbaccess1.mdbfile = txtmdbfile.Text
Me.remotemdbaccess1.recordstofetch = txtrecordstofetch.Text
Me.remotemdbaccess1.dbpassword = txtmdbpassword.Text
Me.remotemdbaccess1.scripturl = txtserverscripturl.Text
Me.remotemdbaccess1.executequery txtsql.Text, txtpos.Text
    Me.ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    'If remotemdbaccess1.recordcount = 0 Then Exit Sub
    For i = 1 To Me.remotemdbaccess1.fieldcount
        Me.ListView1.ColumnHeaders.Add , , Me.remotemdbaccess1.getfield(i)
    Next
    list
Exit Sub
ext:
MsgBox Err.Description, vbCritical
End Sub



Private Sub cmddelete_Click()
On Error GoTo ext:
If Not Me.ListView1.SelectedItem Is Nothing Then
    Me.remotemdbaccess1.Delete Me.ListView1.SelectedItem.Index
    Me.ListView1.ListItems.Remove Me.ListView1.SelectedItem.Index
End If
Exit Sub
ext:
MsgBox Err.Description, vbCritical
End Sub

Public Sub list()
Dim i&, j%, str() As String, litem As ListItem
Me.ListView1.ListItems.Clear
For i = 1 To Me.remotemdbaccess1.recordcount
    DoEvents
    Me.lblstatus.Caption = "Listing " & i & " Of " & Me.remotemdbaccess1.recordcount & " Records"
        str = Split(Me.remotemdbaccess1.getrow(i), "|$|")
        If UBound(str) = -1 Then Exit For
        Set litem = Me.ListView1.ListItems.Add(, , str(0))
        For j = 1 To UBound(str)
            litem.ListSubItems.Add , , str(j)
        Next
Next
Me.lblstatus.Caption = "Listed " & Me.remotemdbaccess1.recordcount & " Records"
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Me.txtrecordstofetch = Me.remotemdbaccess1.recordstofetch
txtserverscripturl.Text = GetSetting("priyan", App.Title, "scripturl", Me.remotemdbaccess1.scripturl)
txtmdbfile.Text = GetSetting("priyan", App.Title, "mdbfile", Me.remotemdbaccess1.mdbfile)
txtsql.Text = GetSetting("priyan", App.Title, "sql", "")
lblstatus.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting "priyan", App.Title, "scripturl", txtserverscripturl.Text
SaveSetting "priyan", App.Title, "mdbfile", txtmdbfile.Text
SaveSetting "priyan", App.Title, "sql", txtsql.Text
'app.logmode always rturn 0 when run from VB IDE
If App.LogMode <> 0 Then
'call exit process api instead of using END statement becuse
'the app will not close soon after unload becuse
' it's a problem of Microsoft Internet Transfer Control
'do't call exitprocess asp when run from VB IDE because the app and ide will be closed
'So i used app.logmode
    ExitProcess (1)
End If
End Sub

Private Sub lblvote_Click()
    vote
End Sub

Private Sub lblabout_Click()
Me.remotemdbaccess1.about
End Sub

Private Sub remotemdbaccess1_status(ByVal status As String)
lblstatus.Caption = "Status : " & status
End Sub
Public Sub vote()
Const url = "http://unni.europe.webmatrixhosting.net/pscredirect/redir.aspx?appid="
'Const url = "http://priyan/home/pscredirect/redir.asp?appname="
Const appid = "Remote MDB"
ShellExecute 0, "open", url & appid, "", "", 1
End Sub


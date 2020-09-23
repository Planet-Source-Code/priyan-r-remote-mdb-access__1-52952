VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Caption         =   "This Active-X Control Allows you to Access MS-Access Database on Your Webserver"
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Mail Me At"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "Visit Me at"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      Caption         =   "admin@priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Label lblsite 
      Caption         =   "http://www.priyan.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmabout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Programmed by Priyan"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   1245
      Left            =   120
      Picture         =   "frmabout.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
lblmail.MouseIcon = lblsite.MouseIcon

End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblmail_Click()
ShellExecute 0, "open", "mailto:" & lblmail.Caption & "?subject=vb:" & App.Title, "", "", 1
End Sub

Private Sub lblsite_Click()
ShellExecute 0, "open", lblsite.Caption, "", "", 1
End Sub


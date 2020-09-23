VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl remotemdbaccess 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   3300
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remote MDB Access"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "remotemdbaccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'=================================
'  Remote Access MDB
'
' Programmed by Priyan
' Visit me at http://www.priyan.tk
' mail me at vb@priyan.tk
' If you found this code useful Please Vote For ME!!!
'=================================
'Default Property Values:
Const m_def_dbpassword = ""
Const m_def_mdbfile = ""
Const m_def_scripturl = ""
Const m_def_recordstofetch = 100
'Property Variables:
Dim canceled As Boolean
Dim m_dbpassword As String
Dim m_mdbfile As String
Dim m_scripturl As String
Dim fields As New Collection
Dim rows As New Collection
Dim m_recordstofetch As Long
Dim last_query$
'Dim m_startpos As Variant
Public Event status(ByVal status$)

Private Sub UserControl_Initialize()
Label1.Top = 0
Label1.Left = 0
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = Label1.Width
UserControl.Height = Label1.Height
End Sub
Public Sub executequery(ByVal query$, Optional ByVal pos&)
Dim data$, temp() As String, str() As String, i&
Dim obj
last_query = ""
Inet1.Cancel
RaiseEvent status("Downloading records...")
canceled = False
'
'To execute and get the result of query download data by this format
' scripturl?mdb file,password|$|query|$|Sql query|$|record start pos|$|Records to fetch
data = Inet1.OpenURL(Me.scripturl & "?" & mdbfile & "," & dbpassword & "|$|query|$|" & query & "|$|" & pos & "|$|" & recordstofetch)
If data = "" Or geterrordescription <> "" Then
    RaiseEvent status(geterrordescription)
    Err.Raise 1, , geterrordescription
    Inet1.Cancel
    Exit Sub
End If
Inet1.Cancel
'split the data get from the script to an array
str = Split(data, "|vbcrlf|")
'if a error occured in the server first member of array will return error,error desciption
'if no error it gives success
If Left(str(0), 5) = "error" Then
    RaiseEvent status("Error Occured")
    Err.Raise 1, , Mid(str(0), 7, Len(str(0)) - 5)
    Exit Sub
ElseIf str(0) <> "success" Then
    RaiseEvent status("An unknown status from the server script")
    Err.Raise 1, , "An unknown status from the server script"
    Exit Sub
End If
'the second meber of array will contain the recordset field names
If UBound(str) < 2 Then Exit Sub
temp = Split(str(2), ",")
clearfields
clearrows
For Each obj In temp
    fields.Add obj, obj
Next
If fields.Count = 0 Then Exit Sub
 If UBound(str) >= 3 Then
    If canceled = True Then
        clearfields
        clearrows
        last_query = ""
        RaiseEvent status("Canceled")
        Exit Sub
    End If
'form the third member of array it will give the records fields seperated by |$|
    For i = 3 To UBound(str) - 1
        rows.Add str(i)
    Next
End If
RaiseEvent status("Execute Successfull")
last_query = query
End Sub

Public Sub delete(ByVal pos&)
    If last_query = "" Then
        Err.Raise 1, , "No Query is executed"
        Exit Sub
    End If
    If pos = 0 Or pos > recordcount Then
        Err.Raise 1, , "Invalid Record Position"
        Exit Sub
    End If
Dim data$, temp() As String, str() As String, i%
Dim obj
Inet1.Cancel
RaiseEvent status("Downloading records...")
'To delete a record use this format
' scripturl?mdb file,password|$|delete|$|Sql query|$|record pos
data = Inet1.OpenURL(Me.scripturl & "?" & mdbfile & "," & dbpassword & "|$|delete|$|" & last_query & "|$|" & pos)
If data = "" Or geterrordescription <> "" Then
    RaiseEvent status(geterrordescription)
    Err.Raise 1, , geterrordescription
    Inet1.Cancel
    Exit Sub
End If
Inet1.Cancel
'if a error occured in the server first member of array will return error,error desciption
'if no error it gives success
str = Split(data, "|vbcrlf|")
If Left(str(0), 5) = "error" Then
    RaiseEvent status("Error Occured")
    Err.Raise 1, , Mid(str(0), 7, Len(str(0)) - 5)
    Exit Sub
ElseIf str(0) <> "success" Then
    RaiseEvent status("An unknown status from the server script")
    Err.Raise 1, , "An unknown status from the server script"
    Exit Sub
End If
RaiseEvent status("Record Deleted Successfully")
rows.Remove pos
End Sub
Public Sub update(ByVal pos&, values$)
Dim obj, i&
If last_query = "" Then
        Err.Raise 1, , "No Query is executed"
        Exit Sub
    End If
    If pos = 0 Or pos > recordcount Then
        Err.Raise 1, , "Invalid Record Position"
        Exit Sub
    End If
Dim data$, temp() As String, str() As String
Inet1.Cancel
RaiseEvent status("Downloading records...")
'Toupdate a record use this format
' scripturl?mdb file,password|$|update|$|Sql query|$|record pos|$|field1=value|~|field2=value|~|field3=value
data = Inet1.OpenURL(Me.scripturl & "?" & mdbfile & "," & dbpassword & "|$|update|$|" & last_query & "|$|" & pos & "|$|" & values)
If data = "" Or geterrordescription <> "" Then
    RaiseEvent status(geterrordescription)
    Err.Raise 1, , geterrordescription
    Inet1.Cancel
    Exit Sub
End If
Inet1.Cancel
If data = "" Then Exit Sub
str = Split(data, "|vbcrlf|")
'if a error occured in the server first member of array will return error,error desciption
'if no error it gives success
If Left(str(0), 5) = "error" Then
    RaiseEvent status("Error Occured")
    Err.Raise 1, , Mid(str(0), 7, Len(str(0)) - 5)
    Exit Sub
ElseIf str(0) <> "success" Then
    RaiseEvent status("An unknown status from the server script")
    Err.Raise 1, , "An unknown status from the server script"
    Exit Sub
End If
str = Split(str(1), "|$|")
data = ""
For i = 0 To UBound(str)
    If i <> 0 Then data = data & "|$|"
    data = data & str(i)
Next
'rows.Remove pos
'If rows.Count <> 0 Then
'    rows.Add data, , pos - 1
'Else
'     rows.Add data
'End If
rows.Add data, , , pos
rows.Remove pos
RaiseEvent status("Updated Successfully")
'If rows.Count <> 0 Then
'    rows.Add data, , pos - 1
'Else
'    rows.Add data
'End If
End Sub
Public Sub addnew(values$)
Dim obj, i&
If last_query = "" Then
        Err.Raise 1, , "No Query is executed"
        Exit Sub
    End If
    
Dim data$, temp() As String, str() As String
RaiseEvent status("Downloading records...")
Inet1.Cancel
'to add a record use
' scripturl?mdb file,password|$|addnew|$|Sql query|$|0|$|field1=value|~|field2=value|~|field3=value
data = Inet1.OpenURL(Me.scripturl & "?" & mdbfile & "," & dbpassword & "|$|addnew|$|" & last_query & "|$|0|$|" & values)
If data = "" Or geterrordescription <> "" Then
    RaiseEvent status(geterrordescription)
    Err.Raise 1, , geterrordescription
    Inet1.Cancel
    Exit Sub
End If
Inet1.Cancel
str = Split(data, "|vbcrlf|")
'if a error occured in the server first member of array will return error,error desciption
'if no error it gives success
If Left(str(0), 5) = "error" Then
    RaiseEvent status("Error Occured")
    Err.Raise 1, , Mid(str(0), 7, Len(str(0)) - 5)
    Exit Sub
ElseIf str(0) <> "success" Then
    RaiseEvent status("An unknown status from the server script")
    Err.Raise 1, , "An unknown status from the server script"
    Exit Sub
End If
str = Split(str(1), "|$|")
data = ""
For i = 0 To UBound(str)
    If i <> 0 Then data = data & "|$|"
    data = data & str(i)
Next
'rows.Remove pos
rows.Add data
RaiseEvent status("Record Added Successfully")

End Sub


Private Function geterrordescription() As String
On Error GoTo ext:
    Select Case Mid$(Inet1.GetHeader, 10, 3)
         Case 403
           geterrordescription = "Server : Can't Open Script File Access Denied"
         Case 404
           geterrordescription = "Server : Script File Not Found"
         Case 401
           geterrordescription = "Server :Can't Open Script File Unauthorized Access"
         Case Else
            geterrordescription = ""
    End Select
Exit Function
ext:
geterrordescription = "Server Error Or Connection failed to the host"

End Function
Public Function Cancel()
Inet1.Cancel
canceled = True
End Function
Public Function getfield(ByVal index&)
    If index = 0 Or index > fields.Count Then
        Err.Raise 1, , "Invalid Index"
    End If
    getfield = fields(index)
End Function

Public Function getrow(ByVal index&)
If index = 0 Or index > rows.Count Then
        Err.Raise 1, , "Invalid Index"
End If
getrow = rows(index)
End Function

Private Sub clearrows()
Dim obj
For Each obj In rows
    rows.Remove 1
Next
End Sub
Private Sub clearfields()
Dim obj
For Each obj In fields
    fields.Remove 1
Next
End Sub
Public Property Get fieldcount() As Long
    fieldcount = fields.Count
End Property

Public Property Get recordcount() As Long
    recordcount = rows.Count
End Property


'===================================================
'    Following are the properties created by using activex-x Control Interface Wizard
'===================================================


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get recordstofetch() As Long
    recordstofetch = m_recordstofetch
End Property

Public Property Let recordstofetch(ByVal New_recordstofetch As Long)
    m_recordstofetch = New_recordstofetch
    PropertyChanged "recordstofetch"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_recordstofetch = m_def_recordstofetch
'    m_startpos = m_def_startpos
    m_scripturl = m_def_scripturl
    m_mdbfile = m_def_mdbfile
    m_dbpassword = m_def_dbpassword
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_recordstofetch = PropBag.ReadProperty("recordstofetch", m_def_recordstofetch)
'    m_startpos = PropBag.ReadProperty("startpos", m_def_startpos)
    m_scripturl = PropBag.ReadProperty("scripturl", m_def_scripturl)
    m_mdbfile = PropBag.ReadProperty("mdbfile", m_def_mdbfile)
    m_dbpassword = PropBag.ReadProperty("dbpassword", m_def_dbpassword)
End Sub



'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("recordstofetch", m_recordstofetch, m_def_recordstofetch)
'    Call PropBag.WriteProperty("startpos", m_startpos, m_def_startpos)
    Call PropBag.WriteProperty("scripturl", m_scripturl, m_def_scripturl)
    Call PropBag.WriteProperty("mdbfile", m_mdbfile, m_def_mdbfile)
    Call PropBag.WriteProperty("dbpassword", m_dbpassword, m_def_dbpassword)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get scripturl() As String
    scripturl = m_scripturl
End Property

Public Property Let scripturl(ByVal New_scripturl As String)
    m_scripturl = New_scripturl
    PropertyChanged "scripturl"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get mdbfile() As String
    mdbfile = m_mdbfile
End Property

Public Property Let mdbfile(ByVal New_mdbfile As String)
    m_mdbfile = New_mdbfile
    PropertyChanged "mdbfile"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get dbpassword() As String
    dbpassword = m_dbpassword
End Property

Public Property Let dbpassword(ByVal New_dbpassword As String)
    m_dbpassword = New_dbpassword
    PropertyChanged "dbpassword"
End Property


Public Sub about()
Attribute about.VB_UserMemId = -552
On Error Resume Next
    frmabout.Show vbModal
End Sub

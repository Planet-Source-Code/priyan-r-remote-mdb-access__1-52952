<%@ Language=VBScript %>
<%
'=================================
'  Remote Access MDB
'
' Programmed by Priyan
' Visit me at http://www.priyan.tk
' mail me at vb@priyan.tk
' If you found this code useful Please Vote For ME!!!
'=================================
on error resume  next

dim str,obj,i,file,password,rs,con,pos,recordstofetch

str=replace(Request.QueryString,"%20"," ")
file=extractstring(extractstring(str,"|$|",0),",",0 )
password=extractstring(extractstring(str,"|$|",0),",",1 )
pos=extractstring(str,"|$|",3)
recordstofetch=extractstring(str,"|$|",4)
if recordstofetch="" then recordstofetch=100
if pos="" then pos=0
'Response.Write(pos &"<br>" & recordstofetch)
'=================================
set con = server.createobject("ADODB.Connection")
set RS = server.createobject("ADODB.Recordset")
con.provider="Microsoft.Jet.OLEDB.4.0;jet oledb:database password=" & password 
con.open(server.MapPath(file))
if err then
	Response.Write("error|" &  err.Description)
	Response.End 
end if
select case extractstring(str,"|$|",1)
	case "query"
		query
	case "delete"				
		delete
	case "update"
		update		
	case "addnew"		
		addnew
end select
con.close
rs.close
set con=nothing
set rs=nothing
Response.End 
sub query()
on error resume  next
temp=0
	RS.Open extractstring(str,"|$|",2) , con, 1, 3	
	if err then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	if pos<>0 then rs.AbsolutePosition=pos

Response.Write("success" & "|vbcrlf|")
Response.Write(rs.RecordCount & "|vbcrlf|" )
if rs.Fields.count=0 then exit sub
for i=0 to rs.Fields.count-1
	if i<>0 then Response.Write(",")
	Response.Write(rs.Fields(cint(i)).Name)
next
Response.Write("|vbcrlf|")
do until rs.EOF	
	for i=0 to rs.Fields.count-1
	if i<>0 then Response.Write("|$|")
		Response.Write(rs.Fields(cint(i)))			
	next
	Response.Write("|vbcrlf|")	
	rs.MoveNext
	temp=temp+1 
	if clng(temp)=clng(recordstofetch) then exit sub
		
loop
end sub
sub delete()
on error resume next
	RS.Open extractstring(str,"|$|",2) , con, 1, 3	
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	if pos<>0 then rs.AbsolutePosition=pos
	rs.Delete 
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	Response.Write("success|vbcrlf|deleted")		
end sub
sub update()
on error resume next
	dim arr,obj
	'Response.Write(extractstring(str,"|$|",2))
	RS.Open extractstring(str,"|$|",2) , con, 1, 3
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	if pos<>0 then rs.AbsolutePosition=pos
	arr=split(extractstring(str,"|$|",4),"|~|" )
	for each obj in  arr
		rs.Fields(extractstring(obj,"=",0))=extractstring(obj,"=",1)
		'Response.Write(extractstring(obj,"=",0) &"<br>" & extractstring(obj,"=",1))
	next
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	rs.Update 
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 		
	Response.Write("success|vbcrlf|")		
	for i=0 to rs.Fields.count-1
			if i<>0 then Response.Write("|$|")
			Response.Write(rs.Fields(cint(i)))
	next
	
end sub
sub addnew()
on error resume next
	dim arr,obj
	'Response.Write(extractstring(str,"|$|",2))
	RS.Open extractstring(str,"|$|",2) , con, 1, 3
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	arr=split(extractstring(str,"|$|",4),"|~|" )
	rs.AddNew 
	for each obj in  arr
		rs.Fields(extractstring(obj,"=",0))=extractstring(obj,"=",1)
	next	
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 
	rs.Update 
	if err.number<>0 then
		Response.Write("error|" &  err.Description)
		Response.End 
	end if 	
	Response.Write("success|vbcrlf|")		
	for i=0 to rs.Fields.count-1
			if i<>0 then Response.Write("|$|")
			Response.Write(rs.Fields(cint(i)))
	next
	
end sub
'====================
Function extractstring(ByVal str, ByVal cmp, ByVal no)
Dim arr
arr = Split(str, cmp)
If no <= UBound(arr) Then
    extractstring = arr(no)
Else
    extractstring = ""
End If

end Function
%>

<%@ Language=VBScript %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="default.css" rel="stylesheet">
<title>Database Administration</title>
</head>
<body>
<!--#include file=config.asp -->
<!--#include file=inc_protect.asp -->
<!--#include file=inc_functions.asp -->
<table WIDTH="100%" ALIGN="center">
	<tr>
		<td width="180px" valign="top"><!--#include file=inc_nav.asp --></td>
		<td>
		
<%
On Error Resume Next
dim con, rec, sTableName, fld, abspage, i, script, pk, sSQL, oIndexes, key
pk = ""
script = Request.ServerVariables("SCRIPT_NAME")
sTableName = Request("table")
set con = Server.CreateObject("ADODB.Connection")
con.Open strProvider & Session("DBAdminDatabase") & Session("DBAdminAuth")
IsError
set oIndexes = new TableIndexes
oIndexes.OpenTable sTableName

'delete record
if Request.QueryString("action") = "delete" then
	sSQL = "DELETE FROM [" & sTableName & "] WHERE"
	fld = ""
	pk = Split(Request.QueryString("pk"), ".")
	key = Split(Request.QueryString("key"), ".")
	for i=0 to UBound(pk)-1
		if Len(pk(i)) > 0 and Len(key(i)) > 0 then
			sSQL = sSQL & fld & " [" & pk(i) & "]=" & key(i)
			fld = " AND"
		end if
	Next
	con.Execute sSQL, adExecuteNoRecords
end if

set rec = Server.CreateObject("ADODB.Recordset")
rec.CursorLocation = adUseClient
if Request("pagesize").Count > 0 and IsNumeric(Request("pagesize")) then
	rec.CacheSize = CInt(Request("pagesize"))
	rec.PageSize = CInt(Request("pagesize"))
else
	rec.CacheSize = 10
	rec.PageSize = 10
end if

sSQL = "SELECT * FROM [" & sTableName & "]"
if Len(Request.QueryString("order")) > 0 then sSQL = sSQL & " ORDER BY " & Request.QueryString("order")
rec.Open sSQL, con, adOpenForwardOnly, adLockReadOnly
if Err then
	Response.Write "<P class=Error>Error opening table: " & Err.Description & "</P>"
else

	if rec.PageCount > 0 then
		if Request("page").Count = 0 or not IsNumeric(Request("page")) then
			rec.AbsolutePage = 1
		else
			if rec.PageCount < CInt(Request("page")) then
				rec.AbsolutePage  = rec.PageCount 
			else
				rec.AbsolutePage = CInt(Request("page"))
			end if
		end if
	end if
	abspage = rec.AbsolutePage
%>
<h1>Data for table: <%=sTableName%></h1>
<p align="center">
<%if Len(oIndexes.GetPrimaryKeyName) > 0 then%>
*&nbsp;<img src="images/add.gif" border="0" WIDTH="16" HEIGHT="16"><a href="recedit.asp?table=<%=Server.URLEncode(sTableName)%>&amp;pk=<%=Server.URLEncode(Join(oIndexes.GetPrimaryKeys,"."))%>&amp;page=<%=Request("page")%>&amp;pagesize=<%=Request("pagesize")%>">Add new record</a>&nbsp;
<%end if%>
*&nbsp;<img src="images/refresh.gif" border="0" WIDTH="16" HEIGHT="16"><a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>">Refresh table</a>&nbsp;
*&nbsp;<img src="images/xml.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_xml.asp?sql=<%=Server.URLEncode(sSQL)%>" alt="Export table content as XML file">XML Export</a>&nbsp;
*&nbsp;<img src="images/excel.gif" border="0" WIDTH="16" HEIGHT="16"><a href="export_csv.asp?sql=<%=Server.URLEncode(sSQL)%>" alt="Export as delimited text file">Excel Export</a>&nbsp;*
</p>
<%if Len(oIndexes.GetPrimaryKeyName) = 0 then%>
<P align=center class=Error>You cannot add/update data in this table. Please set the primary key first</A>
<%end if%>
	<p align="left">
	<%if abspage > 1 then%>
		<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">&laquo;&nbsp;Prev</font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">Next&nbsp;&raquo;</font></a>
	<%end if
	i = 0
	%>
	</p>

<table align="center" border="1" width="100%" class="DataTable">
<tr>
	<th class="DataTH">*</th>
<%for each fld in rec.Fields%>
	<th class="DataTH">
		<%if oIndexes.IsPrimaryKey(fld.Name) then%>
			<img src="images/key.gif" border="0" WIDTH="16" HEIGHT="16">
		<%end if%>
		<A href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode(fld.Name & " ASC")%>" title="Order ascending"><font color=white><%=fld.Name%></font></A>&nbsp;<A href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&order=<%=Server.URLEncode(fld.Name & " DESC")%>" title="Order descending"><font color=white>&darr;</font></A>
	</th>
<%next%>
</tr>

<%do while not rec.EOF and i < rec.PageSize and Err = 0%>
<tr onmouseover="bgColor='#DDDDDD'" onmouseout="bgColor=''">
	<td valign="top" class="DataTD">
	<%if Len(oIndexes.GetPrimaryKeyName) > 0 then
		sSQL = ""
		fld = oIndexes.GetPrimaryKeys()
		for each pk in fld
			if Len(pk) > 0 then
				Select Case rec(pk).Type 
					Case adBSTR,adChar,adLongVarChar,adLongVarWChar,adVarChar,adVarWChar,adWChar
						sSQL = sSQL & "'" & Replace(rec(pk), "'", "''") & "'"
					Case adDate,adDBDate, adDBTime, adDBTimeStamp,adFileTime
						sSQL = sSQL & "CDate('" & FormatDateTime(rec(pk), vbLongDate) & " " & FormatDateTime(rec(pk), vbLongTime) & "')"
					Case Else
						sSQL = sSQL & rec(pk)
				End Select
				sSQL = sSQL & "."
			end if
		Next
	%>
		<a href="recedit.asp?action=edit&amp;pk=<%=Server.URLEncode(Join(oIndexes.GetPrimaryKeys,"."))%>&amp;key=<%=Server.URLEncode(sSQL)%>&amp;table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=Request("page")%>&amp;pagesize=<%=Request("pagesize")%>"><img src="images/edit.gif" alt="Edit the record" border="0" WIDTH="16" HEIGHT="16"></a>
		<a href="javascript:deleteRecord('<%=Server.URLEncode(Replace(Join(oIndexes.GetPrimaryKeys,"."), "'", "\'"))%>','<%=Server.URLEncode(Replace(sSQL, "'", "\'"))%>')"><img src="images/delete.gif" alt="Delete the record" border="0" WIDTH="16" HEIGHT="16"></a>
	<%end if%>
	</td>
	<%for each fld in rec.Fields%>
		<td valign="top" align="center" class="DataTD">
		<%if fld.Type <> adBinary then%>
			<%if fld.Value <> "" then%>
				<%=Replace(fld.Value, "<", "&lt;")%>
			<%else%>
				&nbsp;
			<%end if%>
		<%else%>
			&lt;Binary data&gt;
		<%end if%>
		</td>
	<%next%>
</tr>
<%	rec.MoveNext
	i = i + 1 
loop%>

</table>		

	<p align="left">
	<%if abspage > 1 then%>
		<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage - 1)%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">&laquo;&nbsp;Prev</font></a>
	<%end if%>
	<%for i=1 to rec.PageCount
		if i = abspage then%>
			<font size="2">[<%=i%>]</font>&nbsp;
	<%	else%>
			<font size="1">&nbsp;[<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=i%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><%=i%></a>]&nbsp;</font>
	<%	end if
	Next
	if abspage < rec.PageCount and abspage > 0 then%>
		<a href="<%=script%>?table=<%=Server.URLEncode(sTableName)%>&amp;page=<%=(abspage + 1)%>&amp;pagesize=<%=Request("pagesize")%>&order=<%=Server.URLEncode(Request.QueryString("order"))%>"><font size="1">Next&nbsp;&raquo;</font></a>
	<%end if%>
	</p>

		</td>
	</tr>
</table>

<p>&nbsp;</p>

</body>
<script LANGUAGE="javascript">
<!--
function deleteRecord(pk,key){
	if(confirm("Are you sure you want to delete record with Primary Key(s) " + key + "?")){
		document.location.replace("<%=script%>?action=delete&key=" + escape(key) + "&pk=" + escape(pk) + "&table=<%=sTableName%>&page=<%=Request("page")%>&pagesize=<%=Request("pagesize")%>");
	}
}
//-->
</script>

</html>

<%
	rec.Close
end if
con.Close
set rec = nothing
set con = nothing
set oIndexes = nothing
%>

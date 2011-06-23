<!--#include file="../conn.asp"-->
var subsmallclassid = new Array();
<%
set ors=Conn.Execute("select ClassID,ClassName,ParentID FROM KS_AskClass WHERE parentid<>0 order by rootid,orders")
dim n:n=0
do while not ors.eof
%>
subsmallclassid[<%=n%>] = new Array(<%=ors(2)%>,<%=ors(0)%>,'<%=trim(ors(1))%>')
<%
ors.movenext
n=n+1
loop
ors.close
set ors=nothing
%>
function changesmallclassid(selectValue)
{
document.getElementById('smallclassid').length = 0;
document.getElementById('smallclassid').options[0] = new Option('-选择-','');
for (i=0; i<subsmallclassid.length; i++)
{
if (subsmallclassid[i][0] == selectValue)
{
document.getElementById('smallclassid').options[document.getElementById('smallclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
}
}
}
function changesmallerclassid(selectValue)
{
document.getElementById('smallerclassid').length = 0;
document.getElementById('smallerclassid').options[0] = new Option('-选择-','');
for (i=0; i<subsmallclassid.length; i++)
{
if (subsmallclassid[i][0] == selectValue)
{
	document.getElementById('smallerclassid').style.display='';
	document.getElementById('smallerclassid').options[document.getElementById('smallerclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
}
}
}

<%
exec="select ClassID,ClassName from KS_AskClass where parentid=0 order by rootid"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
%>
document.write ("<select name='classid' id='classid' style='width:120px' size='8' onChange='changesmallclassid(this.value)'>");
document.write ("<option value='' selected>-选择-</option>");
<%
do while not rs.eof%>
document.write ("<option value=<%=rs(0)%>><%=rs(1)%></option>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
document.write ("</select>")

document.write ("  <select name='smallclassid' size='8' onChange='changesmallerclassid(this.value)' style='width:120px' id='smallclassid'>");
document.write ("<option value='' selected>-选择-</option>");
document.write ("</select>")
document.write ("  <select name='smallerclassid' size='8' style='display:none;width:120px' id='smallerclassid'>");
document.write ("<option value='' selected>-选择-</option>");
document.write ("</select>")

<!--#include file="../conn.asp"-->
var subcity = new Array();
<%
set ors=Conn.Execute("select a.ParentID,a.City,b.City FROM KS_Province a inner join KS_Province b on b.id=a.parentid WHERE a.parentid<>0 order by a.orderid")
dim n:n=0
do while not ors.eof
%>
subcity[<%=n%>] = new Array('<%=ors(2)%>','<%=trim(ors(1))%>')
<%
ors.movenext
n=n+1
loop
ors.close
set ors=nothing
%>
function changecity(selectValue)
{
 try{
 setCookie("pid",selectValue);
 }catch(e)
 {
 }
document.getElementById('City').length = 0;
document.getElementById('City').options[0] = new Option('请选择','');
for (i=0; i<subcity.length; i++)
{
if (subcity[i][0] == selectValue)
{
document.getElementById('City').options[document.getElementById('City').length] = new Option(subcity[i][1], subcity[i][1]);
}
}
}

<%
exec="select ID,City from KS_Province where parentid=0 order by orderid"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
%>
document.write ("<select name='Province' id='Province' onChange='changecity(this.value)'>");
document.write ("<option value='' selected>选择省份</option>");
<%
do while not rs.eof%>
document.write ("<option value=<%=rs(1)%>><%=rs(1)%></option>");
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
document.write ("</select>")

document.write (" <select name='City' id='City'>");
document.write ("<option value='' selected>请选择</option>");
document.write ("</select>")
<%
CloseConn
%>
try
{changecity(getCookie("pid"));
}catch(e)
{}

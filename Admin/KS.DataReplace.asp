<%Option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.Asp"-->
<%
Dim KSCls
Set KSCls = New Admin_Replace
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Replace
        Private KS,Action,I,BeginTime,EndTime
		Private Sub Class_Initialize()
		   Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 If Not KS.ReturnPowerResult(0, "KMST10008") Then                '�������ִ��SQL���
				  Call KS.ReturnErr(1, "")
			  Response.End
         End If
%>
<html>
<head>
<title>���ݿ������滻����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="include/admin_style.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
if trim(request("type1"))="GetChird" then
	Call ShowChird()
else
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick=""location.href='KS.DataReplace.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ͨģʽ</span></li>"
		Response.Write "<li class='parent' onclick=""location.href='?Action=Main2';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>�߼�ģʽ</span></li>"
		Response.Write "<li class='parent' onclick=""location.href='?Action=Main3';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>��ֵģʽ</span></li>"
		Response.Write "</ul>"

		i=0
		Action=trim(request("action"))
		Select Case Action
		Case "Replace1","Replace2","Replace3"
			call step1()
		Case "Main2"
			call Main2()
		Case "Main3"
			call Main3()
		Case Else
			call Main1()
		End Select
end if
%>
</body>
</html>
<%End Sub

Sub Main1()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm(){
  if (document.myform.TableName.value==''){
    alert('���ݱ�������Ϊ�գ�');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('�ֶ�������Ϊ�գ�');
    //document.myform.ColumnName.focus();
    return false;
  }
  if (document.myform.strOld.value==''){
    alert('�滻�ַ�����Ϊ�գ�');
    document.myform.strOld.focus();
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="ctable">
<form method="post" name="myform" onSubmit="return CheckForm();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="CLeftTitle"><strong>���ݱ�����</strong></td>
	<td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� ����</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� ����</strong></td>
	<td height="30"><textarea name="strOld" cols="60" rows="4"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� �ɣ�</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="4"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>ע�����</strong></td><td height="30">	1��ִ�в���ǰ���뱸�����ݿ��ļ���<br>2���������ĸ���ʱ���������ݵĶ����Լ����������򱾵ػ����������þ�����������ݺܶ࣬���¿��ܺ������������ǧ����ˢ��ҳ���ر��������������ֳ�ʱ���ߴ�����ʾ����ʹ�ñ����������½��в�����</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" align="center"><input type="hidden" name="Action" value="Replace1"><input class="button" type="submit" name="Submit" value="��ʼ�滻"></td></tr>
</table>
</form>
<%
End Sub

Sub Main2()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm1(){
  if (document.myform.TableName.value==''){
    alert('���ݱ�������Ϊ�գ�');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('�ֶ�������Ϊ�գ�');
    //document.myform.ColumnName.focus();
    return false;
  }
  if (document.myform.strOld.value==''){
    alert('�滻��ʼ���벻��Ϊ�գ�');
    document.myform.strOld.focus();
    return false;
  }
   if (document.myform.strOld1.value==''){
    alert('�滻�������벻��Ϊ�գ�');
    document.myform.strOld1.focus();
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="CTable">
<form method="post" name="myform" onSubmit="return CheckForm1();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="CLeftTitle"><strong>���ݱ�����</strong></td>
	<td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� ����</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�ַ���ʼ��</strong></td>
	<td height="30"><textarea name="strOld" cols="60" rows="3"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�ַ�������</strong></td>
	<td height="30"><textarea name="strOld1" cols="60" rows="3"></textarea><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� �ɣ�</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="3"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><b>�ر�����</b>��</td><td height="30">1��<font color=red>�߼�ģʽ���׳������滻�ٶȽ�������ռ������CPU���ڴ���Դ,�����ʹ��!����������Ҫ,����ʹ����ͨģʽ!</font><br>	2��ִ�в���ǰ���뱸�����ݿ��ļ���<br>3���������ĸ���ʱ���������ݵĶ����Լ����������򱾵ػ����������þ�����������ݺܶ࣬���¿��ܺ������������ǧ����ˢ��ҳ���ر��������������ֳ�ʱ���ߴ�����ʾ����ʹ�ñ����������½��в�����</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" align="center"><input type="hidden" name="Action" value="Replace2"><input type="submit" name="Submit" value="��ʼ�滻" class="button"></td></tr>
</form>
</table>
<%
End Sub

Sub Main3()
%>
<script language = "JavaScript">
function changedb(){
    tablechird.location.href="KS.DataReplace.asp?type1=GetChird&type="+document.myform.TableName.value;
}
function CheckForm1(){
  if (document.myform.TableName.value==''){
    alert('���ݱ�������Ϊ�գ�');
    document.myform.TableName.focus();
    return false;
  }
  if (document.myform.ColumnName.value==''){
    alert('�ֶ�������Ϊ�գ�');
    return false;
  }
  return true;  
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="CTable">
<form method="post" name="myform" onSubmit="return CheckForm1();" action="KS.DataReplace.asp" target="_self">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" width=80 class="CLeftTitle"><strong>���ݱ�����</strong></td><td height="30"><%Call ShowMain()%><font color=red> *</font></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� ����</strong></td>
	<td height="30"><input name="ColumnName" type="hidden" value="" size="40"><iframe style="top:2px;" ID="tablechird" src="KS.DataReplace.asp?type1=GetChird" frameborder=0 scrolling=no width="100%" height="22"></iframe></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><strong>�� �� �ɣ�</strong></td>
	<td height="30"><textarea name="strNew" cols="60" rows="3"></textarea></td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
	<td height="30" align="right" class="CLeftTitle"><b>�ر�����</b>��</td><td height="30">1��<font color=red>��ֵģʽ�����ԭ������,���ǰ�ѡ���ֶε�ֱֵ���޸�Ϊ�µ�ֵ,��ᵼ��ԭ���ݶ�ʧ����Ϊ������!�����ʹ��!</font>	<br>2��ִ�в���ǰ���뱸�����ݿ��ļ���<br>3���������ĸ���ʱ���������ݵĶ����Լ����������򱾵ػ����������þ�����������ݺܶ࣬���¿��ܺ������������ǧ����ˢ��ҳ���ر��������������ֳ�ʱ���ߴ�����ʾ����ʹ�ñ����������½��в�����</td></tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"><td height="30" colspan="2" align="center"><input type="hidden" name="Action" value="Replace3"><input class="button" type="submit" name="Submit" value="��ʼ�滻"></td></tr>
</form>
</table>
<%
End Sub


'**************************************************
'��������Step1
'��  �ã������滻�����ó���
'��  ������
'**************************************************
Sub Step1()
	dim rs,sql
	BeginTime=timer
	dim TableName,ColumnName,strOld,strOld1,strNew
	TableName	= KS.R(trim(request("TableName")))
	ColumnName	= KS.R(trim(request("ColumnName")))
	strOld		= trim(request("strOld"))
	strOld1		= trim(request("strOld1"))
	strNew		= trim(request("strNew"))
	if TableName="" then
		response.write "������Ҫ�滻�����ݱ�����"
		exit sub
	end if
	if ColumnName="" then
		response.write "������Ҫ�滻���ֶ�����"
		exit sub
	end if
	if action="Replace1" then
		if strOld="" then
			response.write "������Ҫ�滻���ַ���"
			exit sub
		end if
	else 
		if action="Replace2" then
			if strOld="" then
				response.write "������Ҫ�滻���ַ���ʼ���룡"
				exit sub
			end if
			if strOld1="" then
				response.write "������Ҫ�滻���ַ��������룡"
				exit sub
			end if
		End if		
	End if
	on error resume next
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "select " & ColumnName & "  from " & TableName
	OpenConn : rs.open sql,conn,1,1
	if err.number<>0 then
		response.write "<font color=red>���ݿ����ʧ�ܣ��������ݱ������ֶ�����д�Ƿ���ȷ</font>"
		exit sub
	end if
	set rs=nothing
	on error GoTo 0
	response.write "<br>�����滻�й����ݡ���<font color=red>�ڴ˹���������ˢ��ҳ���رմ��ڣ�</font><br>"
	call ReplaceData(TableName,ColumnName,strOld,strOld1,strNew)	
	EndTime=timer	

    Response.Write "<table border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1""  class=""ctable"" width=""90%""><tr align='center' class='tdbg'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & "���滻��<font color='#0000ff'>"&i&"</font> �����ݡ�<br>����ʱ��<font color='#0000ff'>"&FormatNumber((EndTime-BeginTime)*1000,2)&"</font> ���롣" & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg' height='30'><td><input type='button' class='button' onclick=history.go(-1)  value='������һҳ'></td></tr></table>"& vbCrLf

End Sub

'**************************************************
'��������ReplaceData
'��  �ã���ʾ���ݱ��б�
'��  ����TableName	----���ݱ���
'        ColumnName	----�ֶ���
'        strOld		----�����ַ�(����ҿ�ʼ����)
'        strOld1	----���ҽ�������
'        strNew		----�滻�ַ�
'**************************************************
Sub ReplaceData(TableName,ColumnName,strOld,strOld1,strNew)
	dim rs,sql,tt
	Response.Write "<li>�����滻����</li>&nbsp;&nbsp;"
	Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "select " & ColumnName & "  from " & TableName 
	rs.open sql,conn,1,3
	Select Case Action
	Case "Replace1"
		do while not rs.eof
			if instr(rs(ColumnName),strOld)<>0 then
				rs(ColumnName)=replace(rs(ColumnName),strOld,strNew)
				i=i+1
				Response.Write "."
			end if
			rs.update
			rs.movenext
		loop
	Case "Replace2"
		Set tt=new RegExp
  		tt.IgnoreCase =true
  		tt.Global=True
  		tt.Pattern=strOld&"[^"&strOld1&"]*"&strOld1
		response.write strOld&"[^"&strOld1&"]*"&strOld1
		do while not rs.eof
 			rs(ColumnName) = tt.Replace(rs(ColumnName),strNew) 
  			i=i+1
			Response.Write "."
			rs.update
			rs.movenext
		loop
		Set tt=Nothing
	Case "Replace3"
		do while not rs.eof
			Response.Write strNew
			rs(ColumnName)=strNew
			i=i+1
			Response.Write "."
			rs.update
			rs.movenext
		loop
	End Select
	rs.close:set rs=nothing
	response.write "&nbsp;&nbsp;<font color='#009900'>�滻���ݳɹ���</font>"
End Sub

'**************************************************
'��������ShowMain
'��  �ã���ʾ���ݱ��б�
'��  ������
'**************************************************
Sub ShowMain()
	dim rs,tablename,temptable
	OpenConn : Set rs = Conn.OpenSchema(4)
	tablename=""
	response.write "<select name='TableName' onChange='changedb()'><option value=''>��ѡ��һ�����ݱ�</option>"
	Do Until rs.EOF
		temptable=rs("Table_name")
		if temptable <> tablename and temptable <> "KS_Admin" and temptable <> "KS_NotDown" and temptable <> "MSysAccessXML" and temptable <> "MSysAccessObjects" then
			Response.write "<option value='" & temptable & "'>" & temptable & "</option>"
			Tablename = temptable
		end if
	rs.MoveNext
	Loop
	response.write "</select>"
	rs.close:set rs=nothing
End Sub

'**************************************************
'��������ShowChird
'��  �ã���ʾָ�����ݱ���ֶ��б�
'��  ������
'**************************************************
Sub ShowChird()
	dim rs
	response.write "<body class='tdbg'><form method='post' name='myform11' action='KS.DataReplace.asp'><select name='dbname2' onChange=parent.document.myform.ColumnName.value=document.myform11.dbname2.value><option value=''>��ѡ��һ���ֶΡ�</option>"
	if trim(request("type"))<>"" then	
		OpenConn : Set rs = Conn.OpenSchema(4)	
		Do Until rs.EOF or rs("Table_name") = trim(request("type"))
			rs.MoveNext
		Loop
		Do Until rs.EOF or rs("Table_name") <> trim(request("type"))
			response.write "<option value='"&rs("column_Name")&"'>"&rs("column_Name")&"</option>"
			rs.MoveNext
		loop
		rs.close:set rs=nothing
	End if
	response.write "</select><font color=red> *</font></form><script language = 'JavaScript'>parent.document.myform.ColumnName.value=document.myform11.dbname2.value;</script></body>"
End Sub
End Class
%> 

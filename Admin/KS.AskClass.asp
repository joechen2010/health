<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%dim channelid
Dim KSCls
Set KSCls = New Admin_Ask_Class
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Class
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
         If Not KS.ReturnPowerResult(0, "WDXT10002") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 KS.Die ""
		 End If
%>
<html>
<head>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="../KS_Inc/common.js" language="JavaScript"></script>
</head>
<body>
<%
    Response.Write "<ul id='menu_top'>"
	Response.Write "<li onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('�ʴ�ϵͳ >> <font color=red>����ʴ����</font>')+'&ButtonSymbol=Go';location.href='?action=add';"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>��ӷ���</span></li>"
	Response.Write "<li onclick='location.href=""?action=orders""' class='parent' onclick='MoveClassInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/set.gif' border='0' align='absmiddle'>һ����������</span></li>"
	Response.Write "<li onclick='location.href=""?action=total""' class='parent' onclick='MoveClassInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unite.gif' border='0' align='absmiddle'>���·���ͳ��</span></li>"
	Response.Write "<li class='parent' onclick=""location.href='?';"""
	if KS.G("Action")="" Then Response.Write " disabled"
	Response.Write"><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>����һ��</span></li>"
	Response.Write "</ul>"

	Dim Action:Action = LCase(Request("action"))
	Select Case Trim(Action)
	Case "savenew"
		Call savenew()
	Case "savedit"
		Call savedit()
	Case "add"
		Call addCategory()
	Case "edit"
		Call editCategory()
	Case "del"
		Call delCategory()
	Case "orders"
		Call ClassOrders()
	Case "updatorders"
		Call UpdateOrders()
	Case "restore"
		Call Restoration()
	Case "total"
	    Call ClassTotal()
	Case Else
		Call showmain()
	End Select
End Sub

Sub showmain()
	Dim Rs,SQL,i
	Dim tdstyle
	Response.Write " <table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
	Response.Write " <tr class='sort'>"
	Response.Write " <td width=""35%"">�ʴ�������� </td>"
	Response.Write " <td width=""43%"">����ѡ��</td>"
	Response.Write "</tr>" & vbNewLine
	SQL = "SELECT * FROM KS_AskClass ORDER BY rootid,orders"
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open SQL, Conn, 1, 1
	If Rs.BOF And Rs.EOF Then
		Response.Write " <tr> <td align=""center"" colspan=""2"" class=""tablerow1"">����û������κη��࣡</td></tr>"
	End If
	i = 0
	Do While Not Rs.EOF
		Response.Write " <tr>"
		Response.Write " <td class='splittd'>"
		Response.Write " "
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;<font color=""#666666"">��</font>"
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;<font color=""#666666"">��</font>"
			Next
			Response.Write "&nbsp;&nbsp;<font color=""#666666"">��</font> "
		End If
		If Rs("parentid") = 0 Then Response.Write ("<b>")
		Response.Write Rs("ClassName")
		If Rs("child") > 0 Then Response.Write "(" & Rs("child") & ")"
		Response.Write " </td>" & vbNewLine
		Response.Write " <td class='splittd' align=""center"">"
		Response.Write "<a onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('�ʴ�ϵͳ >> <font color=red>����ʴ����</font>')+'&ButtonSymbol=Go';"" href=""?action=add&editid="
		Response.Write Rs("classid")
		Response.Write """>��ӷ���</a>"
		Response.Write " | <a onclick=""parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('�ʴ�ϵͳ >> <font color=red>�ʴ��������</font>')+'&ButtonSymbol=GoSave';"" href=""?action=edit&editid="
		Response.Write Rs("classid")
		Response.Write """>��������</a>"
		Response.Write " |"
		Response.Write " "
		If Rs("child") < 1 Then
			Response.Write " <a href=""?action=del&ChannelID="&ChannelID&"&editid="
			Response.Write Rs("classid")
			Response.Write """ onclick=""{if(confirm('ɾ���������÷����������Ϣ��ȷ��ɾ����?')){return true;}return false;}"">ɾ������</a>"
		Else
			Response.Write " <a href=""#"" onclick=""{if(confirm('�÷��ຬ���������࣬������ɾ�����������෽��ɾ�������࣡')){return true;}return false;}"">"
			Response.Write " ɾ������</a>"
		End If
		Response.Write " </td>" & vbNewLine
		Response.Write "</tr>" & vbNewLine
		Rs.movenext
		i = i + 1
	Loop
	Rs.Close
	Set Rs = Nothing
	Response.Write "</table>"
End Sub

Sub addCategory()
	Dim NewClassID
	Dim Rs,SQL,i
	SQL = "SELECT MAX(ClassID) FROM KS_AskClass"
	Set Rs = Conn.Execute(SQL)
	If Rs.BOF And Rs.EOF Then
		NewClassID = 1
	Else
		NewClassID = Rs(0) + 1
	End If
	If IsNull(NewClassID) Then NewClassID = 1
	Rs.Close
%>
<script language="javascript">
function CheckForm(){ 
 if ($F('ClassName')=='')
 {
   alert('�������������!');
   $Foc('ClassName');
   return false;
 }
 $("myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">����ʴ����</div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
	<form name="myform" method="POST" action="?action=savenew">
	<input type="hidden" name="NewClassID" value="<%=NewClassID%>">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>�������ࣺ</strong></td>
		<td>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">��Ϊһ������</option>"
	SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
	Set Rs = Conn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("classid") & """ "
		If Request("editid") <> "" And CLng(Request("editid")) = Rs("classid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;�� "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;��"
			Next
			Response.Write "&nbsp;&nbsp;�� "
		End If
		Response.Write Rs("ClassName") & "</option>" & vbCrLf
		Rs.movenext
	Loop
	Rs.Close
	Response.Write "</select>"
	Set Rs = Nothing
%>
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td width="20%" class="clefttitle" align="right"><strong>�������ƣ�</strong><br/>
		<font color="red">��Ӷ���������ûس��ֿ�</font></td>
		<td width="80%">
		<textarea name="ClassName" cols="50" rows="5"></textarea>
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>����˵����</strong></td>
		<td>
		<textarea name="Readme" cols="50" rows="5"></textarea></td>
	</tr>

	</form>
</table>
<%
End Sub

Sub editCategory()
	Dim RsObj
	Dim Rs,SQL,i
	Set Rs = Conn.Execute("SELECT * FROM KS_AskClass WHERE classid = " & KS.ChkClng(Request("editid")))
	If Rs.BOF And Rs.EOF Then
		FoundErr = True
		ErrMsg = "���ݿ���ִ���,û�д�վ�����!"
		Rs.Close
		Set Rs = Nothing
		Exit Sub
	End If
%>
<script language="javascript">
function CheckForm(){ 
 if ($F('ClassName')=='')
 {
   alert('�������������!');
   $Foc('ClassName');
   return false;
 }
 $("myform").submit();
}
</script>
<div style="text-align:center;height:30px;line-height:30px;font-weight:bold">�༭�ʴ����</div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
	<form name="myform" method="POST" action="?action=savedit">
	<input type="hidden" name="editid" value="<%=Request("editid")%>">
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>�������ࣺ</strong></td>
		<td>
<%
	Response.Write " <select name=""class"">"
	Response.Write "<option value=""0"">��Ϊһ������</option>"
	SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
	Set RsObj = Conn.Execute(SQL)
	Do While Not RsObj.EOF
		Response.Write "<option value=""" & RsObj("classid") & """ "
		If CLng(Rs("parentid")) = RsObj("classid") Then Response.Write "selected"
		Response.Write ">"
		If RsObj("depth") = 1 Then Response.Write "&nbsp;&nbsp;�� "
		If RsObj("depth") > 1 Then
			For i = 2 To RsObj("depth")
				Response.Write "&nbsp;&nbsp;��"
			Next
			Response.Write "&nbsp;&nbsp;�� "
		End If
		Response.Write RsObj("ClassName") & "</option>" & vbCrLf
		RsObj.movenext
	Loop
	RsObj.Close
	Response.Write "</select>"
	Set RsObj = Nothing
%>
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td width="20%" class="clefttitle" align="right"><strong>�������ƣ�</strong></td>
		<td width="80%">
		<input type="text" name="ClassName" id="ClassName" size="35" value="<% = Rs("ClassName")%>">
		</td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>����˵����</strong></td>
		<td >
		<textarea name="Readme" cols="50" rows="5"><%=Server.HTMLEncode(Rs("readme")&"")%></textarea></td>
	</tr>
	<tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td class="clefttitle" align="right"><strong>����ͳ�ƣ�</strong></td>
		<td>
		δ�����<input type="text" name="AskPendNum" size="10" value="<%=Rs("AskPendNum")%>">
		�ѽ����<input type="text" name="AskDoneNum" size="10" value="<%=Rs("AskDoneNum")%>">
		<span style="display:none">
		ͶƱ��<input type="text" name="AskVoteNum" size="10" value="<%=Rs("AskVoteNum")%>">
		����<input type="text" name="AskshareNum" size="10" value="<%=Rs("AskshareNum")%>">
		</span>
		</td>
	</tr>
	</form>
</table>
<%
Set Rs = Nothing
End Sub

Sub savenew()
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("�������������!",-1)
		Exit Sub
	End If
	If Not IsNumeric(Request.Form("class")) Then
		Call KS.AlertHistory("��ѡ����������!",-1)
		Exit Sub
	End If
	'If Trim(Request.Form("Readme")) = "" Then
	'	Call KS.AlertHistory("���������˵��!",-1)
	'	Exit Sub
	'End If
	Dim Rs,SQL,i
	Dim newclassid,rootid,ParentID,depth,orders
	Dim maxrootid,Parentstr,neworders
	Dim m_strClassname,m_arrClassname,strClassname

	m_strClassname = Replace(Trim(Request("classname")), vbCrLf, "$$$")
	m_arrClassname = Split(m_strClassname, "$$$")

	If Request("class") <> "0" Then
		SQL = "SELECT rootid,classid,depth,orders,Parentstr FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("class"))
		Set Rs = Conn.Execute (SQL)
		rootid = Rs(0)
		ParentID = Rs(1)
		depth = Rs(2)
		orders = Rs(3)
		If depth > 3 Then
			Call KS.AlertHistory("��ϵͳ����3������",-1)
			Exit Sub
		End If
		Parentstr = Rs(4)
		Set Rs = Nothing
	Else
		SQL = "SELECT MAX(rootid) FROM KS_AskClass"
		Set Rs = Conn.Execute (SQL)
		maxrootid = KS.ChkClng(Rs(0)) + 1
		If maxrootid =0 Then maxrootid = 1
		Set Rs = Nothing
	End If

	SQL = "SELECT classid FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("newclassid"))
	Set Rs = Conn.Execute (SQL)
	If Not (Rs.EOF And Rs.BOF) Then
		Call KS.AlertHistory("������ָ���ͱ�ķ���һ�������!",-1)
		Exit Sub
	Else
		newclassid = KS.ChkClng(Request("newclassid"))
	End If
	Set Rs = Nothing
	
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_AskClass"
	Rs.Open SQL, Conn, 1, 3
	For i = 0 To UBound(m_arrClassname)
		strClassname = KS.R(Trim(m_arrClassname(i)))
		If strClassname <> "" Then
			Rs.addnew
			If Request("class") <> "0" Then
				Rs("depth") = depth + 1
				Rs("rootid") = rootid
				Rs("parentid") = Request.Form("class")
				'If Parentstr = "0" Then
				'	Rs("Parentstr") = Request.Form("class")
				'Else
				'	Rs("Parentstr") = parentstr & "," & KS.ChkClng(Request.Form("class"))
				'End If
			Else
				Rs("depth") = 0
				Rs("rootid") = maxrootid
				Rs("parentid") = 0
				Rs("ParentStr") = 0
			End If
            Rs("parentstr")=parentstr & newclassid & ","
			Rs("child") = 0
			Rs("classid") = newclassid
			Rs("orders") = newclassid
			Rs("classname") = strClassname
			Rs("readme") = Trim(Request.Form("readme"))
			Rs("Askmaster") = ""
			Rs("c_setting") = ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
			Rs("AskPendNum") = 0
			Rs("AskDoneNum") = 0
			Rs("AskVoteNum") = 0
			Rs("AskshareNum") = 0
			Rs.Update
			Rs.MoveNext
			newclassid = newclassid + 1
			maxrootid = maxrootid + 1
		End If
	Next
	Rs.Close
	Set Rs = Nothing

	CheckAndFixClass 0,1
	Call KS.Confirm("��ϲ��������µķ���ɹ�,���������?","?action=add","?")
End Sub

Sub savedit()
	If CLng(Request.Form("editid")) = CLng(Request.Form("class")) Then
		Call KS.AlertHistory("�������಻��ָ���Լ�",-1)
		Exit Sub
	End If
	If Trim(Request.Form("classname")) = "" Then
		Call KS.AlertHistory("�������������!",-1)
		Exit Sub
	End If
	
	Dim newclassid,maxrootid,readme
	Dim parentid,depth,child,ParentStr,rootid,iparentid,iParentStr
	Dim trs,mrs
	Dim Rs,SQL,nParentStr
	Set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT ParentStr FROM KS_AskClass Where ClassID=" & KS.ChkClng(KS.G("Class")),conn,1,1
	If Not RS.Eof Then
	 nParentStr=Rs(0)
	End If
	Rs.Close
	SQL = "SELECT * FROM KS_AskClass WHERE classid="& KS.ChkClng(Request("editid"))
	Rs.Open SQL,Conn,1,3
	newclassid = Rs("classid")
	parentid = Rs("parentid")
	iparentid = Rs("parentid")
	ParentStr = Rs("ParentStr")
	depth = Rs("depth")
	child = Rs("child")
	rootid = Rs("rootid")
	
	'�ж���ָ���ķ����Ƿ�����������
	If ParentID=0 Then
		If CLng(Request("class"))<>0 Then
		Set trs=Conn.Execute("SELECT rootid FROM KS_AskClass WHERE classid="&KS.ChkClng(Request("class")))
		If rootid=trs(0) Then
			Call KS.AlertHistory("������ָ�����ʴ������������Ϊ��������",-1)
			Exit Sub
		End If

		End If
	Else
		Set trs=Conn.Execute("SELECT classid FROM KS_AskClass WHERE ParentStr like '%"&ParentStr&","&newclassid&"%' And classid="&KS.ChkClng(Request("class")))
		If Not (trs.EOF And trs.BOF) Then
			Call KS.AlertHistory("������ָ�����ʴ������������Ϊ��������",-1)
			Exit Sub
		End If
	End If
	If parentid = 0 Then
		parentid = Rs("classid")
		iparentid=0
	End If
	Rs("parentstr")=nParentStr & rs("classid") & ","
	Rs("classname") = Trim(Request.Form("classname"))
	Rs("parentid") = KS.ChkClng(Request.Form("class"))
	Rs("readme") =Trim( Request("readme"))
	Rs("AskPendNum") = KS.ChkClng(Request.Form("AskPendNum"))
	Rs("AskDoneNum") = KS.ChkClng(Request.Form("AskDoneNum"))
	Rs("AskVoteNum") = KS.ChkClng(Request.Form("AskVoteNum"))
	Rs("AskshareNum") = KS.ChkClng(Request.Form("AskshareNum"))
	Rs.Update 
	Rs.Close
	Set Rs=nothing
	
	Set mrs=Conn.Execute("SELECT MAX(rootid) FROM KS_AskClass")
	Maxrootid=mrs(0)+1
	mrs.close:Set mrs=nothing
	CheckAndFixClass 0,1
	Call KS.Alert("��ϲ���������޸ĳɹ�!","?")
End Sub

Sub delCategory()
	Dim Rs,SQL,i
	Dim ChildStr,nChildStr
	Dim Rss,Rsc
	On Error Resume Next
	Set Rs = Conn.Execute("SELECT ParentStr,child,depth,parentid FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("editid")))
	If Not (Rs.EOF And Rs.BOF) Then
		If Rs(1) > 0 Then
			Call KS.AlertHistory("�÷��ຬ���������࣬��ɾ��������������ٽ���ɾ��������Ĳ���!",-1)
			Exit Sub
		End If

		If Rs(2) > 0 Then
			Conn.Execute ("UPDATE KS_AskClass Set child=child-1 WHERE classid in (" & Rs(0) & ")")
		End If
		For i = 0 To Ubound(AllPostTable)
			SQL = "DELETE FROM " & AllPostTable(i) & " WHERE classid=" & KS.ChkClng(Request("editid"))
			Conn.Execute(SQL)
		Next
		Conn.Execute("DELETE FROM KS_AskAnswer WHERE classid=" & KS.ChkClng(Request("editid")))
		Conn.Execute("DELETE FROM KS_AskTopic WHERE classid=" & KS.ChkClng(Request("editid")))
		Conn.Execute("DELETE FROM KS_AskClass WHERE classid=" & KS.ChkClng(Request("editid")))
		
	End If
	Set Rs = Nothing
	Conn.Execute("UPDATE KS_AskClass Set child=0 WHERE child<0")
	CheckAndFixClass 0,1
	UpdateClassTotal
	Call KS.Alert("��ϲ��������ɾ���ɹ���","?")
End Sub

Sub Restoration()
	CheckAndFixClass 0,1
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub CheckAndFixClass(ParentID,orders)
	Dim Rs,Child,ParentStr
	If ParentID=0 Then
		Conn.Execute("UPDATE KS_AskClass Set Depth=0 WHERE ParentID=0")
	End If
	Set Rs=Conn.Execute("SELECT classid,rootid,ParentStr,Depth FROM KS_AskClass WHERE ParentID="&ParentID&" ORDER BY rootid,orders")
	Do while Not Rs.EOF
		Conn.Execute "UPDATE KS_AskClass Set Depth="&Rs(3)+1&",rootid="&Rs(1)&" WHERE ParentID="&Rs(0)&"",Child
		Conn.Execute("UPDATE KS_AskClass Set Child="&Child&",orders="&orders&" WHERE classid="&Rs(0)&"")
		orders=orders+1
		CheckAndFixClass Rs(0),orders
		Rs.MoveNext
	Loop
	Set Rs=Nothing
	Application(KS.SiteSN&"_askclasslist")=empty
End Sub


Sub ClassTotal()
 UpdateClassTotal()
 Call KS.AlertHistory("��ϲ,����������ͳ�Ƴɹ�!",-1)
End Sub

Sub UpdateClassTotal()
 Dim Rs:Set Rs=Server.CreateObject("ADODB.RECORDSET")
 Rs.Open "Select * From KS_AskClass Order By Rootid,orders",conn,1,3
 do while not rs.Eof 
   Rs("AskPendNum")=Conn.Execute("select count(topicid) From KS_AskTopic WHERE classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&rs("classid")&",%') and topicmode=0")(0)
   Rs("AskDoneNum")=Conn.Execute("select count(topicid) From KS_AskTopic WHERE classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&rs("classid")&",%') and topicmode<>0")(0)
   Rs.Update
  Rs.MoveNext
 Loop
 Rs.Close
 Set RS=Nothing
 Application(KS.SiteSN&"_askclasslist")=empty
End Sub

Sub ClassOrders()
%>
<br>
<table border="0" cellspacing="1" cellpadding="3" align="center"  class="Ctable">
	<tr> 
	<th class="sort" colspan=2>�ʴ�һ���������������޸�(������Ӧ������������������Ӧ���������)</th>
	</tr>
	<tr>
<%
	Dim Rs,SQL,i
	Set Rs=Server.CreateObject("ADODB.Recordset")
	SQL="SELECT * FROM KS_AskClass WHERE ParentID=0 ORDER BY rootid"
	Rs.Open SQL,Conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "��û����Ӧ���ʴ���ࡣ"
	Else
		Do While Not Rs.Eof
		Response.Write "<form action=""?action=updatorders"" method=""post""><tr class='tdbg'>"
		Response.Write "<td align=""right"" class=""clefttitle"">" & rs("ClassName") & "</td><td><input type=""text"" name=""OrderID"" size=""4"" value="""&rs("rootid")&"""><input type=""hidden"" name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=""submit"" name=""Submit"" value=""�޸�"" class=""button""></td></tr></form>"
		Rs.Movenext
		Loop
%>
</table>
<%
	End If
	Rs.Close
	Set Rs=Nothing
%>
	</td>
	</tr>
</table>
<%
End Sub

Sub UpdateOrders()
	Dim cID,OrderID,Rs
	cID = Replace(Request.Form("cID"),"'","")
	OrderID = Replace(Request.Form("OrderID"),"'","")
	Set Rs = Conn.Execute("SELECT classid FROM KS_AskClass WHERE rootid="&orderid)
	If Rs.EOF And Rs.BOF Then
		Conn.Execute("UPDATE KS_AskClass SET rootid="&OrderID&" WHERE rootid="&cID)
		Call KS.AlertHintScript("���óɹ�!")
	Else
		Call KS.AlertHistory("�벻Ҫ����������������ͬ�����",-1)
		Response.End
	End If
End Sub
End Class
%>
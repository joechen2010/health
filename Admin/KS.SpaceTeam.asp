<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SpaceTeam
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceTeam
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10004") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeam.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>Ȧ�ӹ���</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=topic';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>���ӹ���</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeamSkin.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>ģ�����</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>Ȧ�ӷ���</span></li>"
			  .Write	" </ul>"
			End With
		
		
		maxperpage = 30 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("�����ϵͳ����!����������")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		Select Case KS.G("action")
		 Case "Del"
		  Call TeamDel()
		 Case "lock"
		  Call TeamLock()
		 Case "unlock"
		  Call TeamUnLock()
		 Case "verific"
		  Call TeamVerific()
		 Case "recommend"
		  Call Blogrecommend()
		 Case "Cancelrecommend"
		  Call BlogCancelrecommend()
		 case "topic" topicshow
		 case "topicdel" topicdel
		 case "class" classshow
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>Ȧ������</th>
	<td nowrap>������</th>
	<td nowrap>����ʱ��</th>
	<td nowrap>Ȧ��״̬</th>
	<td nowrap>�������</th>
</tr>
<%
		totalPut = Conn.Execute("Select Count(blogID) from KS_Blog")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Team order by id desc"
		Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û���û�����Ȧ�ӣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="KS.SpaceTeam.asp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="../space/group.asp?id=<%=rs("id")%>" target="_blank"><%=Rs("Teamname")%></a>
	<%If rs("recommend")=1 then response.write "<font color=red>��</font>"%>
	</td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("verific")
	 case 0
	  response.write "δ��"
	 case 1
	  response.write "<font color=red>����</font>"
	 case 2
	  response.write "<font color=blue>����</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../space/group.asp?id=<%=rs("id")%>" target="_blank">���</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('ɾ��Ȧ�ӽ�ɾ��Ȧ���µ�������Ϣ��ȷ��ɾ����'));">ɾ��</a> &nbsp;<%if rs("verific")=0 then%><a href="?Action=verific&id=<%=rs("id")%>">���</a> <%elseif rs("verific")=1 then%><a href="?Action=lock&id=<%=rs("id")%>">����</a><%elseif rs("verific")=2 then%><a href="?Action=unlock&id=<%=rs("id")%>">����</a><%end if%>
	<%if rs("recommend")="0" then%>
	<a href="?Action=recommend&id=<%=rs("id")%>">��Ϊ�Ƽ�</a>
	<%else%>
	<a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>ȡ���Ƽ�</font></a>
	<%end if%>
	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input type="hidden" name="action" value="Del">
	<input class="button" type="submit" name="Submit2" value="����ɾ��" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){document.getElementById('action').value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" value="�������" onclick="document.getElementById('action').value='verific';">
	<input class="button" type="submit" value="��������" onclick="document.getElementById('action').value='lock';">
	<input class="button" type="submit" value="��������" onclick="document.getElementById('action').value='unlock';">
	</td>
</tr>
</form>
<tr>
	<td  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.SpaceTeam.asp", True, "��", CurrentPage, "Action=" & Action)
	%></td>
</tr>
</table>

<%
End Sub

	'ɾ��
	Sub TeamDel()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 Dim tid
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
	 RS.Open "Select * from KS_Team Where ID in(" & id & ")",conn,1,1
	 do while not rs.eof
	  tid=rs("id")
  	Conn.execute("delete from ks_uploadfiles where channelid=1030 and infoid in(" & tid& ")")
  	Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where teamid in(" & tid& "))")
	  Conn.execute("Delete From KS_TeamTopic Where teamid=" & tid)
	  Conn.Execute("Delete From KS_TeamUsers Where teamid=" & tid)
	  rs.movenext
	 loop
	 rs.close:set rs=nothing
	 Conn.execute("Delete From KS_Team Where ID In("& id & ")")
	 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	'����
	Sub TeamLock()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=2 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'����
	Sub TeamUnLock()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	'���
	Sub TeamVerific
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set verific=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	End Sub
	
	sub Blogrecommend()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set recommend=1 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	end sub
	
	sub BlogCancelrecommend()
	 Dim ID:ID=replace(KS.G("ID")," ","")
	 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	 Conn.execute("Update KS_Team Set recommend=0 Where ID In("& id & ")")
	 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
	end sub
     
	'���ӹ���
    Sub topicshow()
%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>����</th>
	<td nowrap>�� �� ��</th>
	<td nowrap>�� �� ʱ ��</th>
	<td nowrap>״ ̬</th>
	<td nowrap>�� �� �� ��</th>
</tr>
<%
		totalPut = Conn.Execute("Select Count(ID) from KS_TeamTopic")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_TeamTopic order by id desc"
	If DataBaseType = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û���˷���Ȧ�����⣡</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=topicdel>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd">
	<%if rs("parentid")=0 then
	 response.write "<font color=red>[��]</font>"
	 end if
	 %>
	<a href="../space/group.asp?action=showtopic&id=<%=rs("teamid")%>&tid=<%=rs("id")%>" target="_blank"><%=Rs("title")%></a><% if rs("isbest")="1" then response.write "<img src=""../images/jh.gif"" align=""absmiddle"">"%></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "δ��"
	 case 1
	  response.write "<font color=red>����</font>"
	 case else
	  response.write "����"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../space/group.asp?action=showtopic&id=<%=rs("teamid")%>&tid=<%=rs("id")%>" target="_blank">���</a> <a href="?Action=topicdel&ID=<%=RS("ID")%>" onclick="return(confirm('ȷ��ɾ����������'));">ɾ��</a> </td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�е����� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "?", True, "ƪ", CurrentPage, KS.QueryParam("page"))
	%></td>
</tr>
</table>

<%
End Sub

'ɾ������
Sub topicdel()
 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
 If ID="" Then KS.Die "<script>alert('�Բ�����û��ѡ��!');history.back();</script>"

		 dim rst:set rst=server.createobject("adodb.recordset")
		 rst.open "select * from ks_teamtopic where id in(" & id & ")",conn,1,3
		 if not rst.eof then
		   do while not rst.eof
			 Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(" & rst("id")& ")")
			 Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where parentid=" & rst("id")& "))")
		     conn.execute("delete from ks_teamtopic where parentid=" & rst("id"))
			 rst.delete
			 rst.movenext
		   loop
		 end if
		 rst.close:set rst=nothing
		 response.write "<script>alert('ɾ���ɹ�');location.href='"& request.servervariables("http_referer") & "';</script>"
End Sub

'�������
Sub ClassShow()
%>		
		  
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>���</strong></td>
			<td width="217"><strong>��������</strong></td>
			<td width="197"><strong>����</strong></td>
			<td width="196"><strong>�������</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_TeamClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">��û������κε�Ȧ�ӷ���!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?action=class&x=a">
				<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" align="center"><%=rs("ClassID")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rs("ClassID")%>"></td>
				  <td class="splittd" width="217" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="Submit" class="button" type="submit"value=" �޸� ">&nbsp;
				  <a onclick="return(confirm('ȷ��ɾ����?'))" href="?action=class&x=c&classid=<%=rs("classid")%>">ɾ��</a></td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?action=class&x=b" method="post" name="myform" id="form">
		    <tr>
		      <td class="spltitd" colspan="4" height="25">&nbsp;&nbsp;<strong>&gt;&gt;����Ȧ�ӷ���<<</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="spltitd" height="25"></td>
			  <td class="spltitd" height="25" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td class="spltitd" height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td class="spltitd" height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,�ύ"></td>
			</tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_TeamClass set ClassName='" & KS.G("ClassName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "��ϲ,�����޸ĳɹ�!"
		   case "b"
		       If KS.G("ClassName")="" Then Response.Write "<script>alert('��������������!');history.back();</script>":response.end
			   
				conn.execute("Insert into KS_TeamClass(ClassName,orderid)values('" & KS.G("ClassName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				KS.AlertHintScript "��ϲ,������ӳɹ�!"
		   case "c"
				conn.execute("Delete from KS_TeamClass where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "��ϲ,����ɾ���ɹ�!"
		End Select
		%></body>
		</html>
<%End Sub

End Class
%> 

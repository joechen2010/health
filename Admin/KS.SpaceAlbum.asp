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
Set KSCls = New Admin_Photoxc
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Photoxc
        Private KS
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		With Response
					If Not KS.ReturnPowerResult(0, "KSMS10003") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		    .Write "<script src='../KS_Inc/common.js'></script>"
		    .Write "<script src='../KS_Inc/jquery.js'></script>"
		    .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='KS.SpaceAlbum.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>������</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=showzp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>��Ƭ����</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=photoclass';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>������</span></li>"
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
		 Case "Del" PhotoDel
		 Case "lock" PhotoLock
		 Case "unlock" PhotoUnLock
		 Case "verific" Photoverific
		 Case "recommend" Photorecommend
		 Case "Cancelrecommend" PhotoCancelrecommend
		 case "showzp" showzp
		 case "delzp" delzp
		 case "photoclass" photoclass
		 Case Else
		  Call showmain
		End Select
End Sub

Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='4%' nowrap align="center">ѡ��</th>
	<td width="27%" nowrap>�������
	  </th>
	<td width="8%" nowrap>�� �� ��</th>
	<td width="18%" nowrap>����ʱ��</th>
	<td width="9%" nowrap>״ ̬
	  </th>
	<td width="11%" nowrap>�� ��    
	<td width="23%" nowrap>�������</th></tr>
<%
		totalPut = Conn.Execute("Select Count(id) from KS_photoxc")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Photoxc order by id desc"
		Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>û���û�������ᣡ</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="KS.SpaceAlbum.asp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd">
	<img src="<%=rs("photourl")%>" width="32" height="32" style="padding:2px;border:1px solid #f1f1f1">
	<a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>" target="_blank"><%=Rs("xcname")%>(<font color=red><%=Rs("xps")%></font>)</a></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "δ��"
	 case 1
	  response.write "<font color=red>����</font>"
	 case 2
	  response.write "<font color=blue>����</font>"
	end select
	%></td>
	<td class="splittd" align="center">
	<font color=red>
	<% select case rs("flag")
	    case 1 :response.write "��ȫ����"
		Case 2 :response.write "��Ա����"
		case 3 :response.write "���빲��"
		case 4 :response.write "������˽"
	   end select
	%></font></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>" target="_blank">���</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('ɾ����Ὣɾ��������������Ƭ��ȷ��ɾ����'));">ɾ��</a> <%IF rs("recommend")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>ȡ���Ƽ�</font></a><%else%><a href="?Action=recommend&id=<%=rs("id")%>">��Ϊ�Ƽ�</a><%end if%>&nbsp;<%if rs("status")=0 then%><a href="?Action=verific&id=<%=rs("id")%>">���</a> <%elseif rs("status")=1 then%><a href="?Action=lock&id=<%=rs("id")%>">����</a><%elseif rs("status")=2 then%><a href="?Action=unlock&id=<%=rs("id")%>">����</a><%end if%></td>
</tr>
<%
		  Rs.movenext
		  i = i + 1:If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input type="hidden" name="action" value="Del" />
	<input class=Button type="submit" name="Submit2" value="����ɾ��" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){document.getElementById('action').value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" name="vbutton" value="�������" onclick="document.getElementById('action').value='verific';">
	<input class="button" type="submit" name="vbutton" value="��������" onclick="document.getElementById('action').value='lock';">
	<input class="button" type="submit" name="vbutton" value="��������" onclick="document.getElementById('action').value='unlock';">
	</td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=8 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

'�鿴��Ƭ
Sub ShowZP()
		totalPut = Conn.Execute("Select Count(id) from KS_Photozp")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='4%' nowrap>ѡ��</th>
	<td width="27%" nowrap>�� Ƭ �� ��
	  </th>
	<td width="8%" nowrap>�� �� ��</th>
	<td width="18%" nowrap>�� �� ʱ ��</th>
	<td width="9%" nowrap>�� С
	  </th>
	<td width="11%" nowrap>�� �� �� ��    
	<td width="23%" nowrap>�� �� �� ��</th></tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Photozp order by id desc"
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
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>û���û�������Ƭ��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?action=delzp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd">
	<img src="<%=rs("photourl")%>" width="32" height="32" style="padding:2px;border:1px solid #f1f1f1">
	<a href="<%=rs("photourl")%>" target="_blank" title="<%=rs("title")%>"><%=Rs("title")%></a></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%=round(rs("photosize")/1024,2)%> kb
	</td>
	<td class="splittd" align="center">
	<a href="../space/?<%=rs("username")%>/showalbum/<%=rs("xcid")%>" target="_blank">
	<font color=red>
	<%=conn.execute("select xcname from ks_photoxc where id=" & rs("xcid"))(0)%>
	</font></a></td>
	<td class="splittd" align="center"><a href="<%=rs("photourl")%>" target="_blank" title="<%=rs("title")%>">���</a> <a href="?Action=delzp&ID=<%=rs("id")%>" onclick="return(confirm('ɾ����Ƭ��ɾ����Ƭ���������Ƭ��ȷ��ɾ����'));">ɾ��</a> </td>
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
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�е���Ƭ" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=8 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

'ɾ����Ƭ
Sub DelZP()
	Dim ID:ID=KS.FilterIDS(KS.G("ID"))
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ������Ƭ!",ComeUrl):Response.End
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where id in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   KS.DeleteFile(rs("photourl"))
	   Conn.execute("update ks_photoxc set xps=xps-1 where id=" & rs("xcid"))
	   rs.movenext
	   loop
	end if
	Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid in(" & id& ")")
	Conn.execute("delete from ks_photozp where id in(" & id& ")")
	rs.close:set rs=nothing
   Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'ɾ�����
Sub PhotoDel()
	Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
	Conn.Execute("Delete From KS_Photoxc Where ID In(" & ID & ")")
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where xcid in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid=" & rs("id"))
	   KS.DeleteFile(rs("photourl"))
	   rs.movenext
	   loop
	end if
	Conn.execute("delete from ks_uploadfiles where channelid=1028 and infoid in(" & id& ")")
	Conn.execute("delete from ks_photozp where xcid in(" & id& ")")
	rs.close:set rs=nothing
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'��Ϊ����
Sub Photorecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set recommend=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'ȡ������
Sub PhotoCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set recommend=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'����
Sub PhotoLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=2 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'����
Sub PhotoUnLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'���
Sub Photoverific
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'������
sub photoclass()
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>���</strong></td>
			<td width="217"><strong>��������</strong></td>
			<td width="197"><strong>����</strong></td>
			<td width="196"><strong>�������</strong></td>
		  </tr>
		   <form name="form1" id='from1' method="post" action="?">
			 <input type="hidden" name="action" value="photoclass">
             <input name="ClassID" type="hidden" id="ClassID" value="">
             <input name="x" type="hidden" id="x" value="a">
		  <%dim orderid
		  set rs = conn.execute("select * from KS_PhotoClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">��û������κε�������!</td></tr>"
			else
			   do while not rs.eof%>
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" height="22" align="center"><%=rs("ClassID")%> </td>
				  <td class="splittd" width="217" align="center"><input name="ClassName<%=rs("classid")%>" type="text" class="textbox" id="ClassName<%=rs("classid")%>" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID<%=rs("classid")%>" type="text" class="textbox" id="OrderID<%=rs("classid")%>" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="button" onclick="$('#x').val('a');$('#ClassID').val('<%=rs("classid")%>');" class="button" type="submit"value=" �޸� ">&nbsp;<input  onclick='if (confirm("ȷ��ɾ����")==true){$("#x").val("c");$("#ClassID").val("<%=rs("classid")%>");}' name="Submit2" type="submit" class="button" value=" ɾ�� "></td>
				</tr>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
		</form>
			<form action="?x=b" method="post" name="myform" id="form">
			<input type="hidden" name="action" value="photoclass">
		    <tr>
		      <td height="22" colspan="4" class="splittd">&nbsp;&nbsp;<strong>&gt;&gt;����������<<</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="splittd">&nbsp;</td>
			  <td class="splittd" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td class="splittd" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td class="splittd" align="center"><input name="Submit3" class="button" type="submit" value="OK,�ύ"></td>
			</tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_PhotoClass set ClassName='" & KS.G("ClassName" & KS.G("ClassID")) & "',orderid='" & KS.ChkClng(KS.G("OrderID" & KS.G("ClassID"))) &"' where ClassID="&KS.ChkClng(KS.G("ClassID"))&"")
				Response.Redirect Request.ServerVariables("http_referer")
		   case "b"
		       If KS.G("ClassName")="" Then Response.Write "<script>alert('��������������!');history.back();</script>":response.end
			   
				conn.execute("Insert into KS_PhotoClass(ClassName,orderid)values('" & KS.G("ClassName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				Response.Redirect Request.ServerVariables("http_referer")
		   case "c"
				conn.execute("Delete from KS_PhotoClass where ClassID="&KS.G("ClassID")&"")
				Response.Redirect Request.ServerVariables("http_referer")
		End Select
		%></body>
		</html>
<%End Sub
End Class
%> 

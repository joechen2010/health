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
					If Not KS.ReturnPowerResult(0, "KSMS10003") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		    .Write "<script src='../KS_Inc/common.js'></script>"
		    .Write "<script src='../KS_Inc/jquery.js'></script>"
		    .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='KS.SpaceAlbum.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>相册管理</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=showzp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>照片管理</span></li>"
			.Write "<li class='parent' onclick=""location.href='?action=photoclass';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>相册分类</span></li>"
			.Write	" </ul>"
		End With	
		
		
		maxperpage = 30 '###每页显示数
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("错误的系统参数!请输入整数")
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
	<td width='4%' nowrap align="center">选择</th>
	<td width="27%" nowrap>相册名称
	  </th>
	<td width="8%" nowrap>创 建 者</th>
	<td width="18%" nowrap>创建时间</th>
	<td width="9%" nowrap>状 态
	  </th>
	<td width="11%" nowrap>类 型    
	<td width="23%" nowrap>管理操作</th></tr>
<%
		totalPut = Conn.Execute("Select Count(id) from KS_photoxc")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Photoxc order by id desc"
		Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>没有用户创建相册！</td></tr>"
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
	  response.write "未审"
	 case 1
	  response.write "<font color=red>已审</font>"
	 case 2
	  response.write "<font color=blue>锁定</font>"
	end select
	%></td>
	<td class="splittd" align="center">
	<font color=red>
	<% select case rs("flag")
	    case 1 :response.write "完全公开"
		Case 2 :response.write "会员开见"
		case 3 :response.write "密码共享"
		case 4 :response.write "个人稳私"
	   end select
	%></font></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>" target="_blank">浏览</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('删除相册将删除相册里的所有照片，确定删除吗？'));">删除</a> <%IF rs("recommend")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>取消推荐</font></a><%else%><a href="?Action=recommend&id=<%=rs("id")%>">设为推荐</a><%end if%>&nbsp;<%if rs("status")=0 then%><a href="?Action=verific&id=<%=rs("id")%>">审核</a> <%elseif rs("status")=1 then%><a href="?Action=lock&id=<%=rs("id")%>">锁定</a><%elseif rs("status")=2 then%><a href="?Action=unlock&id=<%=rs("id")%>">解锁</a><%end if%></td>
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
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input type="hidden" name="action" value="Del" />
	<input class=Button type="submit" name="Submit2" value="批量删除" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){document.getElementById('action').value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" name="vbutton" value="批量审核" onclick="document.getElementById('action').value='verific';">
	<input class="button" type="submit" name="vbutton" value="批量锁定" onclick="document.getElementById('action').value='lock';">
	<input class="button" type="submit" name="vbutton" value="批量解锁" onclick="document.getElementById('action').value='unlock';">
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

'查看照片
Sub ShowZP()
		totalPut = Conn.Execute("Select Count(id) from KS_Photozp")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '得到总页数
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='4%' nowrap>选择</th>
	<td width="27%" nowrap>相 片 名 称
	  </th>
	<td width="8%" nowrap>上 传 者</th>
	<td width="18%" nowrap>上 传 时 间</th>
	<td width="9%" nowrap>大 小
	  </th>
	<td width="11%" nowrap>归 属 相 册    
	<td width="23%" nowrap>管 理 操 作</th></tr>
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
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>没有用户创建照片！</td></tr>"
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
	<td class="splittd" align="center"><a href="<%=rs("photourl")%>" target="_blank" title="<%=rs("title")%>">浏览</a> <a href="?Action=delzp&ID=<%=rs("id")%>" onclick="return(confirm('删除照片将删除照片里的所有照片，确定删除吗？'));">删除</a> </td>
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
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
	<input class=Button type="submit" name="Submit2" value=" 删除选中的照片" onclick="{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"></td>
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

'删除照片
Sub DelZP()
	Dim ID:ID=KS.FilterIDS(KS.G("ID"))
	If ID="" Then Call KS.Alert("你没有选中要删除的照片!",ComeUrl):Response.End
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
   Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'删除相册
Sub PhotoDel()
	Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
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
 Response.Write "<script>alert('删除成功！');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'设为精华
Sub Photorecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set recommend=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'取消精华
Sub PhotoCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set recommend=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'锁定
Sub PhotoLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=2 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'解锁
Sub PhotoUnLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'审核
Sub Photoverific
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('对不起，您没有选择!');history.back();</script>":response.end
 Conn.execute("Update KS_photoxc Set status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'相册分类
sub photoclass()
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td width="217"><strong>类型名称</strong></td>
			<td width="197"><strong>排序</strong></td>
			<td width="196"><strong>管理操作</strong></td>
		  </tr>
		   <form name="form1" id='from1' method="post" action="?">
			 <input type="hidden" name="action" value="photoclass">
             <input name="ClassID" type="hidden" id="ClassID" value="">
             <input name="x" type="hidden" id="x" value="a">
		  <%dim orderid
		  set rs = conn.execute("select * from KS_PhotoClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">还没有添加任何的相册分类!</td></tr>"
			else
			   do while not rs.eof%>
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" height="22" align="center"><%=rs("ClassID")%> </td>
				  <td class="splittd" width="217" align="center"><input name="ClassName<%=rs("classid")%>" type="text" class="textbox" id="ClassName<%=rs("classid")%>" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID<%=rs("classid")%>" type="text" class="textbox" id="OrderID<%=rs("classid")%>" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="button" onclick="$('#x').val('a');$('#ClassID').val('<%=rs("classid")%>');" class="button" type="submit"value=" 修改 ">&nbsp;<input  onclick='if (confirm("确定删除吗？")==true){$("#x").val("c");$("#ClassID").val("<%=rs("classid")%>");}' name="Submit2" type="submit" class="button" value=" 删除 "></td>
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
		      <td height="22" colspan="4" class="splittd">&nbsp;&nbsp;<strong>&gt;&gt;新增相册分类<<</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="splittd">&nbsp;</td>
			  <td class="splittd" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td class="splittd" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td class="splittd" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_PhotoClass set ClassName='" & KS.G("ClassName" & KS.G("ClassID")) & "',orderid='" & KS.ChkClng(KS.G("OrderID" & KS.G("ClassID"))) &"' where ClassID="&KS.ChkClng(KS.G("ClassID"))&"")
				Response.Redirect Request.ServerVariables("http_referer")
		   case "b"
		       If KS.G("ClassName")="" Then Response.Write "<script>alert('请输入类型名称!');history.back();</script>":response.end
			   
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

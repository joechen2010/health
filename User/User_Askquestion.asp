<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

Dim KSCls
Set KSCls = New User_myask
KSCls.Kesion()
Set KSCls = Nothing

Class User_myask
        Private KS,KSUser
		Private CurrentPage,totalPut,i,PageNum
		Private RS,MaxPerPage,SQL,tablebody,action
		Private ComeUrl,TotalPages
		Private Sub Class_Initialize()
			MaxPerPage =10
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  KS.Die "<script>top.location.href='Login';</script>"
		End If
		Action=Request("action")
		
		If KS.S("Action")="cancel" Then	 Call FavCancel() : KS.Die ""
		
		CurrentPage=KS.ChkClng(Request("page"))
		if CurrentPage=0 Then CurrentPage=1
		Call KSUser.Head()
		Call KSUser.InnerLocation("我的问题")
		KSUser.CheckPowerAndDie("s09")
		call info()

	  End Sub

		
	  sub info()
		%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr>
				<td colspan="10" height="45" align="center"><%=KSUser.UserName%>  等级头衔:<%=KSUser.GradeTitle%></td>
			</tr>
		    <tr class="title" align="center">
			 <td height="25" >
			  <font color=#ff6600>总分</font></td>
			 <td>回答总数</td>
			 <td>回答被采纳</td>
			 <td>回答被采纳率</td>
			 <td width="50"></td>
			 <td>提问总数</td>
			 <td>已解决</td>
			 <td>解决中</td>
			 <td>被推荐</td>
			 <td>被关闭</td>
		    </tr>				
		    <tr align="center">
			 <td height="25" >
			  <font color=#ff6600>
			 <%=KSUser.Score%>
			  </font>
			 </td>
			 <td>
			 <%
			  Dim AnswerTotal:AnswerTotal=Conn.Execute("Select count(answerid) From KS_AskAnswer Where UserName='" & KSUser.UserName & "'")(0)
			  Response.Write AnswerTotal
			 %>
			 </td>
			 <td>
			 <%
			  Dim AnswerTotalCN:AnswerTotalCN=Conn.Execute("Select count(answerid) From KS_AskAnswer Where UserName='" & KSUser.UserName & "' AND Answermode=1")(0)
			  Response.Write AnswerTotalCN
			 %>
			 </td>
			 <td>
			 <%
			 If AnswerTotal<>0 Then
			 Response.Write formatpercent(AnswerTotalCN/AnswerTotal,2)
			 Else
			  Response.Write "0%"
			 End If
			 %>
			 </td>
			 <td width="50"></td>
			 <td><%=Conn.Execute("Select count(topicid) From KS_AskTopic Where UserName='" & KSUser.UserName & "'")(0)%></td>
			 <td><%=Conn.Execute("Select count(topicid) From KS_AskTopic Where UserName='" & KSUser.UserName & "' and topicmode=1")(0)%></td>
			 <td><%=Conn.Execute("Select count(topicid) From KS_AskTopic Where UserName='" & KSUser.UserName & "' and topicmode=0")(0)%></td>
			 <td><%=Conn.Execute("Select count(topicid) From KS_AskTopic Where UserName='" & KSUser.UserName & "' and recommend=1")(0)%></td>
			 <td><%=Conn.Execute("Select count(topicid) From KS_AskTopic Where UserName='" & KSUser.UserName & "' and closed=1")(0)%></td>
		    </tr>				
            </table>
			
		<div class="tabs">	
			<ul>
				<li<%If action="" then KS.Echo " class='select'"%>><a href="?">我的问题</a></li>
				<li<%If action="answer" Then KS.Echo " class='select'"%>><a href="?action=answer">我的回答</a></li>
				<li<%If action="fav" Then KS.Echo " class='select'"%>><a href="?action=fav">我的收藏</a></li>
			</ul>
		</div>
			<table height='400' width="99%" align="center">
			<tr>
			<td valign="top">
		
   <%
          select Case Action
		   case "answer" answer
		   case "fav" fav
		   case else quesion
		  end select
		  
    Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
   %>
			 </td>
			 </tr>
		    </table>
		<%
	end sub
	
	Sub Quesion()
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">标题</td>
				<td width="10%" align="center">悬赏分</td>
				<td width="15%" align="center">状态</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select b.classname,a.* from KS_asktopic a inner join KS_AskClass b on a.classid=b.classid where a.Username='"&KSUser.UserName&"' order by a.topicid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有提问题！</td>
			</tr>
		<%else
		          totalPut = RS.RecordCount
				  If CurrentPage < 1 Then	CurrentPage = 1
								
			      If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrentPage - 1) * MaxPerPage
				  Else
					  CurrentPage = 1
				  End If
				  i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td height="25" class="splittd">
							<div class="ContentTitle">
							·<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>提问时间:[<%=rs("dateandtime")%>]
							  是否审核:[<%if rs("locktopic")="1" then response.write "未审核" else response.write "已审核"%>]
							  分类:[<%=rs(0)%>]
							 </span>
							 </div>
							</td>

							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>无悬赏</span>
							<%end if%>
							</td>
							<td class="splittd" align=center>
							<img src="../<%=KS.Asetting(1)%>images/ask<%=rs("topicmode")%>.gif">
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			loop
			end if
			rs.close
			set rs=Nothing
		%>
</table>
	<%
	End Sub
	
	Sub answer()
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">标题</td>
				<td width="10%" align="center">提问者</td>
				<td width="10%" align="center">悬赏分</td>
				<td width="15%" align="center">状态</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select a.*,b.reward from KS_AskAnswer a inner join KS_AskTopic b on a.topicid=b.topicid where a.Username='"&KSUser.UserName&"' order by a.answerid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有回答过问题！</td>
			</tr>
		<%else
		
		                       totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td height="25" class="splittd">
							<div class="ContentTitle">
							·<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>回答时间:[<%=rs("answertime")%>]
							  分类:[<%=rs("classname")%>]
							 </span>
							 
							 </div>
							</td>
                            <td class="splittd" align=center><%=rs("postusername")%></td>
							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>无悬赏</span>
							<%end if%>
							</td>
							
							<td class="splittd" align=center>
							<img src="../<%=KS.Asetting(1)%>images/ask<%=rs("topicmode")%>.gif">
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
		</table>
		<%
	End Sub
	
	Sub Fav()
	%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">标题</td>
				<td width="10%" align="center">提问者</td>
				<td width="10%" align="center">悬赏分</td>
				<td width="15%" align="center">状态</td>
			</tr>
			<form name="myform" action="?action=cancel" method="post">
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select b.classname,b.username,a.*,b.title,b.topicmode,b.reward from KS_askfavorite a inner join KS_AskTopic b on a.topicid=b.topicid where a.Username='"&KSUser.UserName&"' order by a.favorid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有收藏问题！</td>
			</tr>
		<%else
		
		                       totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td height="25" class="splittd">
							<div class="ContentTitle">
							 <input type="checkbox" name="favorid" value="<%=rs("favorid")%>">
							·<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>收藏时间:[<%=rs("FavorTime")%>]
							 &nbsp;分类:[<%=rs(0)%>]
							 </span>
							 
							 </div>
							</td>
                            <td class="splittd" align=center><%=rs(1)%></td>
							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>无悬赏</span>
							<%end if%>
							</td>
							
							<td class="splittd" align=center>
							<img src="../<%=KS.Asetting(1)%>images/ask<%=rs("topicmode")%>.gif">
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
		<tr>
		 <td><input type="submit" value="取消收藏" class="button" onClick="return(confirm('确定取消收藏吗?'))"></td>
		</tr>
		</form>
	 </table>
	 <%
	End Sub
		
	Sub FavCancel()
		 Dim FavorID:Favorid=KS.FilterIDS(KS.S("favorid"))
		 if FavorID="" Then KS.AlertHintScript "对不起,您没有选择记录!"
		 Conn.Execute("Delete From KS_AskFavorite Where Favorid in(" & Favorid & ") and username='" & KSUser.UserName & "'")
		 Response.Redirect ComeUrl
	End Sub	
End Class
%> 

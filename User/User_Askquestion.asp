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
		Call KSUser.InnerLocation("�ҵ�����")
		KSUser.CheckPowerAndDie("s09")
		call info()

	  End Sub

		
	  sub info()
		%>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr>
				<td colspan="10" height="45" align="center"><%=KSUser.UserName%>  �ȼ�ͷ��:<%=KSUser.GradeTitle%></td>
			</tr>
		    <tr class="title" align="center">
			 <td height="25" >
			  <font color=#ff6600>�ܷ�</font></td>
			 <td>�ش�����</td>
			 <td>�ش𱻲���</td>
			 <td>�ش𱻲�����</td>
			 <td width="50"></td>
			 <td>��������</td>
			 <td>�ѽ��</td>
			 <td>�����</td>
			 <td>���Ƽ�</td>
			 <td>���ر�</td>
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
				<li<%If action="" then KS.Echo " class='select'"%>><a href="?">�ҵ�����</a></li>
				<li<%If action="answer" Then KS.Echo " class='select'"%>><a href="?action=answer">�ҵĻش�</a></li>
				<li<%If action="fav" Then KS.Echo " class='select'"%>><a href="?action=fav">�ҵ��ղ�</a></li>
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
				<td height="25" align="center">����</td>
				<td width="10%" align="center">���ͷ�</td>
				<td width="15%" align="center">״̬</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select b.classname,a.* from KS_asktopic a inner join KS_AskClass b on a.classid=b.classid where a.Username='"&KSUser.UserName&"' order by a.topicid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">��û�������⣡</td>
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
							��<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>����ʱ��:[<%=rs("dateandtime")%>]
							  �Ƿ����:[<%if rs("locktopic")="1" then response.write "δ���" else response.write "�����"%>]
							  ����:[<%=rs(0)%>]
							 </span>
							 </div>
							</td>

							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>������</span>
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
				<td height="25" align="center">����</td>
				<td width="10%" align="center">������</td>
				<td width="10%" align="center">���ͷ�</td>
				<td width="15%" align="center">״̬</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select a.*,b.reward from KS_AskAnswer a inner join KS_AskTopic b on a.topicid=b.topicid where a.Username='"&KSUser.UserName&"' order by a.answerid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">��û�лش�����⣡</td>
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
							��<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>�ش�ʱ��:[<%=rs("answertime")%>]
							  ����:[<%=rs("classname")%>]
							 </span>
							 
							 </div>
							</td>
                            <td class="splittd" align=center><%=rs("postusername")%></td>
							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>������</span>
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
				<td height="25" align="center">����</td>
				<td width="10%" align="center">������</td>
				<td width="10%" align="center">���ͷ�</td>
				<td width="15%" align="center">״̬</td>
			</tr>
			<form name="myform" action="?action=cancel" method="post">
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select b.classname,b.username,a.*,b.title,b.topicmode,b.reward from KS_askfavorite a inner join KS_AskTopic b on a.topicid=b.topicid where a.Username='"&KSUser.UserName&"' order by a.favorid desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=3 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">��û���ղ����⣡</td>
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
							��<a href="../<%=KS.ASetting(1)%><%if KS.ASetting(16)="1" then KS.Echo "show-" & rs("topicid") & KS.ASetting(17) Else KS.Echo "q.asp?id=" & rs("topicid")%>" target="_blank"><%=KS.GotTopic(rs("title"),60)%></a> 
							</div>
							<div class="Contenttips">
			                 &nbsp;<span>�ղ�ʱ��:[<%=rs("FavorTime")%>]
							 &nbsp;����:[<%=rs(0)%>]
							 </span>
							 
							 </div>
							</td>
                            <td class="splittd" align=center><%=rs(1)%></td>
							<td class="splittd" align=center>
							<%if rs("reward")>0 then%>
							<img src="../<%=KS.Asetting(1)%>images/ask_xs.gif" align="absmiddle">
							<font color=red><%=rs("reward")%></font>
							<%else%>
							 <span>������</span>
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
		 <td><input type="submit" value="ȡ���ղ�" class="button" onClick="return(confirm('ȷ��ȡ���ղ���?'))"></td>
		</tr>
		</form>
	 </table>
	 <%
	End Sub
		
	Sub FavCancel()
		 Dim FavorID:Favorid=KS.FilterIDS(KS.S("favorid"))
		 if FavorID="" Then KS.AlertHintScript "�Բ���,��û��ѡ���¼!"
		 Conn.Execute("Delete From KS_AskFavorite Where Favorid in(" & Favorid & ") and username='" & KSUser.UserName & "'")
		 Response.Redirect ComeUrl
	End Sub	
End Class
%> 

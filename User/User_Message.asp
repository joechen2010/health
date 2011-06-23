<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_Message
KSCls.Kesion()
Set KSCls = Nothing

Class User_Message
        Private KS,KSUser
		Private Max_sEnd
        Private Max_sms
		Private Max_Num
        Private Action
        Private RS,SqlStr,ComeUrl
		Private FoundErr,Errmsg
		Private i
		Private CurrentPage,TotalPut,MaxPerPage
		Private Sub Class_Initialize()
		   MaxPerPage=10
		   Set KS=New PublicCls
		   Set KSUser = New UserCls
		   Max_sEnd=KS.Setting(49)	'群发限制人数
		   Max_sms=KS.Setting(48)	'内容最多字符数
		   Max_Num=KS.Setting(47)  '最多允许存放条数
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
				IF Cbool(KSUser.UserLoginChecked)=false Then
				  Response.Write "<script>top.location.href='Login';</script>"
				  Exit Sub
				End If
			  Action=Lcase(request("action"))
			  ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
			  If ComeUrl="" Then ComeUrl="User_Message.asp"
			  Call KSUser.Head()
		%>
				<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Touser.value=='')
				{
				   alert('请输入收信人!')
				   document.myform.Touser.focus();
				   return false;
				 }
				if (document.myform.title.value=='')
				{
				   alert('请输入信件主题!')
				   document.myform.title.focus();
				   return false;
				 }

				if (frames["MessageContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
                document.myform.message.value=frames["MessageContent"].KS_EditArea.document.body.innerHTML;
	
				if (document.myform.message.value=='')
					{
					alert("请输入信件内容！");
					frames["MessageContent"].KS_EditArea.focus();
					return false;
					}
				 return true;  
				}
				</script>
				
				<div class="tabs">	
			<ul>
				<li<%if Action="" or Action="inbox" or action="read" or action="fw" or Action="outbox" or Action="issend" or Action="recycle" or Action="new" then response.write " class='select'"%>><a href="User_Message.asp">短消息(<span class="red"><%=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)%></span>)</a></li>
				<li<%If action="friendrequest" then KS.Echo " class='select'"%>><a href="?action=friendrequest">好友请求(<font color=red><%=conn.execute("select count(id) from ks_friend where friend='" & ksuser.username & "' and accepted=0")(0)%></font>)</a></li>
				<li<%if Action="message" or Action="replaymessage" or Action="savemessagereplay" then response.write " class='select'"%>><a href="?action=Message">空间留言(<span class="red"><%=Conn.Execute("Select Count(ID) From KS_BlogMessage Where username='" &KSUser.UserName &"' And readtf=0")(0)%></span>)</a></li>
				<li<%if Action="comment" or Action="replaycomment" or Action="savecommentreplay" then response.write " class='select'"%>><a href="?action=Comment">日志评论(<span class="red"><%=Conn.Execute("Select Count(ID) From KS_BlogComment Where username='" &KSUser.UserName &"' And readtf=0")(0)%></span>)</a></li>

			</ul>
        </div>
		<div class="clear"></div>
		<%
		IF Action="" or Action="inbox" or Action="outbox" or Action="issend" or Action="recycle" or Action="new" Then
		 %>
		<div align="center" style="height:30">
						<a href="User_Message.asp?action=inbox"><img src="Images/inbox.gif" border=0 alt="收件箱"></a> &nbsp;
						<a href="User_Message.asp?action=outbox"><img src="Images/outbox.gif" border=0 alt="发件箱"></a> &nbsp; 		<a href="User_Message.asp?action=issend"><img src="Images/issend.gif" border=0 alt="已发送邮件"></a> &nbsp;
						<a href="User_Message.asp?action=recycle"><img src="Images/recycle.gif" border=0 alt="废件箱"></a> &nbsp; 		
						<a href="User_Message.asp?action=new"><img src="Images/write.gif" border=0 alt="发送消息"></a> 
		
		<table width="98%" border="0" align="center" class="border" cellpadding="2" cellspacing="1" style="display:nowrap">
<tr class="tdbg">
<td width="127" align="right">您的邮箱容量：</td>
<td width="602" ><img src="images/bar.gif" width="0" height="16" id="Sms_bar" align="absmiddle" /></td>
<td width="211"  align="center" id="Sms_txt">100%</td>
</tr></table>

		</div>
		 <%
response.write showtable("Sms_bar","Sms_txt",conn.execute("select count(*) from KS_Message where Incept='"&KSUser.UserName&"'")(0),Max_Num)

		Else
		 Response.Write "<br>"
		End IF
		
		Select Case Action
		Case "new" : sendMessage
		Call KSUser.InnerLocation("发送消息")
		Case "read" : read
		Call KSUser.InnerLocation("阅读消息")
		Case "outread" : read
		Case "delet" : delete
		Case "newmsg" : newmsg
		Case "send" : savemsg
		Case "fw" : fw
		Case "edit" : edit
		Case "savedit" : savedit
		Case "删除收件" : delinbox
		Case "清空收件箱" : AllDelinbox
		Case "删除草稿" : deloutbox
		Case "清空草稿箱" : AllDeloutbox
		Case "删除已发送的消息" : DelIsSend
		Case "清空已发送的消息" : AllDelIsSend
		Case "删除垃圾" : delrecycle
		Case "清空垃圾箱" : AllDelrecycle
		Case "message" : Message
		Case "replaymessage" : ReplayMessage
		Case "savemessagereplay" :  SaveMessageReplay
        Case "messagedel" : MessageDel
		Case "comment" : Comment
		Case "replaycomment" :ReplayComment
		Case "savecommentreplay" : SaveCommentReplay
		Case "friendrequest" : friendrequest
		Case "accepta" : friendAcceptA
		Case "accept" : friendaccept
		Case "delfriend" : FriendDel
		Case "deletefriend" : FriendDelete
		Case "commentdel" : CommentDel
		Case Else : MessageMain
		End Select

		End Sub
		
		'处理好友请求
		Sub friendrequest()
		        If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
				Else
								  CurrentPage = 1
				End If
                Dim Accepted                  
				Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
				Dim Sql:sql = "select * from KS_Friend Where Friend='" &KSUser.UserName & "' and accepted<>1 order by id DESC" 
				  Call KSUser.InnerLocation("好友请求")
		  %>
	      <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
		  <form name="myform" id="myform" action="User_Message.asp" method="post">
		  <input type="hidden" name="action" id="action" value="accepta">
                             <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' class='splittd' align='center' colspan='6' height='30' valign='top'>还没有用户给您发邀请，要加油哦!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
								   If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Dim XML,Node
								Set XML=KS.ArrayToxml(RS.GetRows(maxperpage),rs,"row","root")
								If IsObject(XML) Then
								  For Each Node In XML.DocumentElement.SelectNodes("row")
								    Accepted=Node.SelectSingleNode("@accepted").text
								    Response.Write "<tr>"
									Response.Write " <td width='45' align='center' class='splittd'><input type='checkbox' name='id' value='" & Node.SelectSingleNode("@id").text & "'></td>"
									Response.Write " <td height='45' class='splittd'><img src='../images/user/log/106.gif'/>朋友：<a href='../space?" & Node.SelectSingleNode("@username").text & "' target='_blank'>" & Node.SelectSingleNode("@username").text & "</a>请求您成为他的好友!"
									if accepted="2" then response.write "<font color=#ff6600>(已拒绝)</font>"
									Response.Write "<br/>附言：" & Node.SelectSingleNode("@message").text & "</td>"
									Response.Write " <td class='splittd' align='center' width='240'>"
									If Accepted="2" Then
									Response.Write "<a href='?action=deletefriend&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('此操作不可逆，确定删除该请求吗？'))"" class='box'>删除</a>"
									Else
									Response.Write "<a href='?action=AcceptA&id=" & Node.SelectSingleNode("@id").text & "' class='box'>接受并加为好友</a> <a href='?action=Accept&id=" & Node.SelectSingleNode("@id").text & "' class='box'>接受</a> <a href='?action=delfriend&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('此操作不可逆，确定拒绝该请求吗？'))"" class='box'>拒绝</a>"
									Response.Write " <a href='?action=deletefriend&id=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('此操作不可逆，确定删除该请求吗？'))"" class='box'>删除</a>"
									End If
									Response.Write "</td>"
									Response.Write "</tr>"
								  Next
								End If
								XML=Empty
				End If
           %>   
		     <tr>
			   <td colspan='4' height='35' class='splittd'>
			     <table borer='0' width='100%'>
				  <td>
			    &nbsp;&nbsp;<label><input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有</label><input class="button" onClick="$('#action').val('accepta')" type="submit" value="接受并加为好友" name=submit1> <input class="button" onClick="$('#action').val('accept')" type="submit" value=" 接 受 " name=submit1> <input class="button" onClick="$('#action').val('delfriend');return(confirm('此操作不可逆,确定拒绝选中的请求吗？'));" type="submit" value=" 拒 绝 " name=submit1> <input class="button" onClick="$('#action').val('deltefriend');return(confirm('此操作不可逆,确定删除吗?'));" type="submit" value=" 接 受 " name=submit1>
			      </td>

				 </tr>
				</table>
				<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
			  </td>
			 </tr>                  
             </table>
		  <%
		End Sub
		
		'同意加为好友，并加他
		Sub friendAcceptA()
		 Dim ID:ID=KS.S("ID")
		 If ID="" Then Call KS.AlertHistory("对不起，您没有选择!",-1)
		 ID=KS.FilterIDs(ID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select UserName,Accepted From KS_Friend Where ID in(" & ID & ")",conn,1,3
		 Do While Not RS.Eof 
		   RS("Accepted")=1
		   RS.Update
		   Conn.Execute("insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&KSUser.UserName&"','"&RS("UserName")&"',"&SqlNowString&",1,'',1)")
		   Call KS.SendInfo(rs("username"),KS.Setting(0),KSUser.UserName & "已通过您的好友请求!","亲爱的" & RS("UserName") & "!<br />&nbsp;&nbsp;&nbsp;&nbsp;恭喜您！<br/><br/>本站会员：<a href=""../space?" & KSUser.UserName & """ target=""_blank"">" & KSUser.UserName & "</a>已接受您的加为好友请求！并将您加为好友了。<br /><br />备注：此信息由系统自动发布，请不要回复！！！")
		   Call KSUser.AddLog(KSUser.UserName,"同意了<a href=""{$GetSiteUrl}space/?" & RS("UserName") & """ target=""_blank"">" & RS("UserName") & "</a>将自己加为好友，并将他也加为了好友!",106)
		  RS.MoveNext
		 Loop
		 RS.Close
		 Set RS=Nothing
		 KS.AlertHintScript("恭喜，操作成功!")
		End Sub
		'同意好邀请
		Sub friendaccept()
         Dim ID:ID=KS.S("ID")
		 If ID="" Then Call KS.AlertHistory("对不起，您没有选择!",-1)
		 ID=KS.FilterIDs(ID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select UserName,Accepted From KS_Friend Where ID in(" & ID & ")",conn,1,3
		 Do While Not RS.Eof 
		   RS("Accepted")=1
		   RS.Update
		   Call KS.SendInfo(rs("username"),KS.Setting(0),KSUser.UserName & "已通过您的好友请求!","亲爱的" & RS("UserName") & "!<br />&nbsp;&nbsp;&nbsp;&nbsp;恭喜您！<br/><br/>本站会员：<a href=""../space?" & KSUser.UserName & """ target=""_blank"">" & KSUser.UserName & "</a>已接受您的加为好友请求！<br /><br />备注：此信息由系统自动发布，请不要回复！！！")
		   Call KSUser.AddLog(KSUser.UserName,"同意了<a href=""{$GetSiteUrl}space/?" & RS("UserName") & """ target=""_blank"">" & RS("UserName") & "</a>将自己加为好友!",106)
		  RS.MoveNext
		 Loop
		 RS.Close
		 Set RS=Nothing
		 KS.AlertHintScript("恭喜，操作成功!")		
		End Sub
		
		'拒绝好友请求
		Sub FriendDel()
         Dim ID:ID=KS.S("ID")
		 If ID="" Then Call KS.AlertHistory("对不起，您没有选择!",-1)
		 ID=KS.FilterIDs(ID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select UserName,Accepted From KS_Friend Where Accepted=0 and ID in(" & ID & ")",conn,1,3
		 Do While Not RS.Eof 
		   RS("Accepted")=2
		   RS.Update
		   Call KS.SendInfo(rs("username"),KS.Setting(0),KSUser.UserName & "拒绝您的好友请求!","亲爱的" & RS("UserName") & "!<br /><br/>本站会员：<a href=""../space?" & KSUser.UserName & """ target=""_blank"">" & KSUser.UserName & "</a>已拒绝了您的加为好友请求！<br /><br />备注：此信息由系统自动发布，请不要回复！！！")
		   Call KSUser.AddLog(KSUser.UserName,"拒绝了<a href=""{$GetSiteUrl}space/?" & RS("UserName") & """ target=""_blank"">" & RS("UserName") & "</a>的好友请求!",106)
		  RS.MoveNext
		 Loop
		 RS.Close
		 Set RS=Nothing
		 KS.AlertHintScript("恭喜，操作成功!")		
		End Sub
		
		'删除好友请求
		Sub FriendDelete()
         Dim ID:ID=KS.S("ID")
		 If ID="" Then Call KS.AlertHistory("对不起，您没有选择!",-1)
		 ID=KS.FilterIDs(ID)
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select UserName,Accepted From KS_Friend Where  ID in(" & ID & ")",conn,1,3
		 Do While Not RS.Eof 
		   Call KSUser.AddLog(KSUser.UserName,"删除了<a href=""{$GetSiteUrl}space/?" & RS("UserName") & """ target=""_blank"">" & RS("UserName") & "</a>的好友请求操作!",106)
		   RS.Delete
		  RS.MoveNext
		 Loop
		 RS.Close
		 Set RS=Nothing
		 KS.AlertHintScript("恭喜，操作成功!")		
		End Sub
		
		
		 '评论管理
	   Sub Comment()
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									Dim Sql:sql = "select * from KS_BlogComment "& Param &" order by AddDate DESC" 
								    Call KSUser.InnerLocation("日志评论")
			
								  %>
								     
				                    <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
												  <td width="12%" height="22" align="center">发表人</td>
                                                  <td width="33%" height="22" align="center">评论标题</td>
                                                  <td width="12%" height="22" align="center">发表时间</td>
                                                  <td width="8%" height="22" align="center">主页</td>
                                                  <td width="8%" height="22" align="center">回复</td>
                                                  <td width="21%" height="22" align="center" nowrap>管理操作</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>没有用户给你评论!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
								   If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowComment
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowComment
									Else
										CurrentPage = 1
										Call ShowComment
									End If
								End If
				End If
     %>                     
                        </table>
		  <%
	   End Sub
	   
	   Sub ShowComment()
	        Dim I
    Response.Write "<FORM Action=""?Action=CommentDel"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
           <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
               <td class="splittd" height="22" align="center">
				<INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID">
			 </td>
				<td class="splittd" align="center"><%=RS("AnounName")%></td>
                <td class="splittd" align="left"><a href="../space/?<%=KSUser.UserName%>/log/<%=rs("logid")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("title")),35)%></a>
											
				</td>
                <td class="splittd" align="center"><%=KS.GetTimeFormat(rs("adddate"))%></td>
                <td class="splittd" align="center">
				  <%if rs("homepage")="" or lcase(rs("homepage"))="http://" then%>
				     ---
				  <%else%>
					 <a href="<%=rs("homepage")%>" target="_blank">访问</a>
				  <%end if%>
				  </td>
				  <td class="splittd" align="center">
				  <%if KS.IsNul(rs("replay")) Then
				     response.write "<span style='color:red'>未回复</font>"
					else
					 response.write "<span style='color:green'>已回复</font>"
					end if
				  %>
				  </td>
                <td class="splittd" height="22" align="center">
											<a href="?id=<%=rs("id")%>&Action=ReplayComment&page=<%=CurrentPage%>" class="box">回复</a> <a href="?action=CommentDel&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除评论吗?'))" class="box">删除</a>
				</td>
            </tr>
                  
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=6 valign=top>
								&nbsp; <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中本页显示的所有评论<INPUT class='button' onClick="return(confirm('确定删除选中的评论吗?'));" type=submit value=删除选定的评论 name=submit1>  
								  </td>
								  </FORM>
								</tr>
								<tr><td colspan="6">
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								</td>
								</tr>
								<% 

	   End Sub
	   '回复评论
	   Sub ReplayComment()
	     Call KSUser.InnerLocation("回复评论")
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_BlogComment Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
		   If KS_A_RS_Obj.Eof And KS_A_RS_Obj.Bof Then
		    Response.Write "<script>alert('参数出错!');history.back();</script>"
			Response.end
		   End If
		   KS_A_RS_Obj("readtf")=1
		   KS_A_RS_Obj.update
		   Dim Title:Title=KS_A_RS_Obj("Title")
		   Dim Content:Content=KS_A_RS_Obj("Content")
		   Dim Replay:Replay=KS_A_RS_Obj("Replay"):If IsNull(Replay) Then Replay=""
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				  if (FCKeditorAPI.GetInstance('Replay').GetXHTML(true)=="")
					{
					  alert("请输入回复内容！");
					  FCKeditorAPI.GetInstance('Replay').Focus();
					  return false;
					}
				
				 return true;  
				}
				</script>
				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?Action=SaveCommentReplay&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>回 复 评 论</td>
					</tr>

                      <tr class="tdbg">
                           <td  height="25" align="center"><span>评论标题：</span></td>
                              <td>  <input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                    </tr>
							 
                              <tr class="tdbg">
                                  <td  height="25" align="center"><span>评论内容：</span></td>
                                  <td>
								  <textarea name="Content" style="display:none"><%=Server.HtmlEncode(Content)%></textarea><iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic" width="98%" height="150" frameborder="0" scrolling="no"></iframe> 
								  </td>
                            </tr>
                              <tr class="tdbg">
                                  <td  height="25" align="center"><span>回复内容：</span></td>
                                  <td>
								  <textarea name="Replay" style="display:none"><%=Server.HtmlEncode(Replay)%></textarea><iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Replay&amp;Toolbar=Basic" width="98%" height="150" frameborder="0" scrolling="no"></iframe> 
								  </td>
                            </tr>
								
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,立即回复 " />
                      <input type="reset" name="Submit2"   class="Button" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
	   End Sub
	   
	   '保存评论回复
	   Sub SaveCommentReplay()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
		Dim Title:Title=KS.S("Title")
		Dim Content:Content=Request.Form("Content")
		Dim Replay:Replay=Request.Form("Replay")
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_BlogComment Where ID=" & ID,conn,1,3
		If Not RS.Eof Then
		  RS("Title")=Title
		  RS("Content")=Content
		  RS("Replay")=Replay
		  RS("ReplayDate")=Now
		 RS.Update
		End If
		RS.Close:Set RS=Nothing
		Call KSUser.AddLog(KSUser.UserName,"回复了评论操作!评论内容:" & Content & "回复:" & Replay & "",100)
		Response.Write "<script>alert('恭喜,您已成功回复！');location.href='?Action=Comment';</script>"
	   End Sub 
	     '删除评论
	  Sub CommentDel()
		Dim ID:ID=KS.S("ID")
		ID=KS.FilterIDs(ID)
		If ID="" Then Call KS.Alert("你没有选中要删除的日志评论!",ComeUrl):Response.End
		Conn.Execute("Delete From KS_BlogComment Where UserName='" & KSUser.UserName & "' and ID In(" & ID & ")")
		Call KSUser.AddLog(KSUser.UserName,"对用户的评论进行删除操作!",100)
		Response.Redirect ComeUrl
	  End Sub
	  
		
		
		 '留言管理
	   Sub Message()
			    If KS.S("page") <> "" Then
					 CurrentPage = KS.ChkClng(KS.S("page"))
				Else
					 CurrentPage = 1
				End If
                Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
				Dim Sql:sql = "select * from KS_BlogMessage "& Param &" order by AddDate DESC" 
				  Call KSUser.InnerLocation("留言管理")
			
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
												  <td width="12%" height="22" align="center">发表人</td>
                                                  <td width="41%" height="22" align="center">留言标题</td>
                                                  <td width="12%" height="22" align="center">发表时间</td>
                                                  <td width="8%" height="22" align="center">主页</td>
                                                  <td width="21%" height="22" align="center" nowrap>管理操作</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>没有用户给你留言!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
											If CurrentPage < 1 Then
												CurrentPage = 1
											End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowMessage
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowMessage
									Else
										CurrentPage = 1
										Call ShowMessage
									End If
								End If
				End If
     %>                     
                        </table>
		  <%
	   End Sub
	   
	   Sub ShowMessage()
	        Dim I
    Response.Write "<FORM Action=""?Action=MessageDel"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td width="5%" height="30" class="splittd" align="center">
											<INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
											<td width="10%" class="splittd" align="center"><%=RS("AnounName")%></td>
                                            <td width="35%" class="splittd" align="left"><a href="../space/?<%=KSUser.UserName%>/message#<%=rs("id")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("title")),35)%></a>
											<%if Not IsNull(RS("Replay")) or rs("replay")<>"" Then
											response.write "<font color=#ff6600>(已回复)</font>"
											end if
											%>
											</td>
                                            <td width="18%" class="splittd" align="center"><%=KS.GetTimeFormat(rs("adddate"))%></td>
                                            <td width="10%" class="splittd" align="center">
											 <a href="<%=rs("homepage")%>" target="_blank">访问</a></td>
                                            <td class="splittd" align="center">
											<a href="?id=<%=rs("id")%>&Action=ReplayMessage&page=<%=CurrentPage%>" class="box">回复</a> <a href="?action=MessageDel&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除留言吗?'))" class="box">删除</a>
											</td>
                                          </tr>
                  
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=2 valign=top><label><INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">选中</label><INPUT class='button' onClick="return(confirm('确定删除选中的留言吗?'));" type=submit value=删除留言 name=submit1> 
								   </td>
								   <td colspan='10' align='right'>    
				<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								  </FORM>
								</tr>
								<% 

	   End Sub
		

	   '回复留言
	   Sub ReplayMessage()
	     Call KSUser.InnerLocation("回复留言")
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_BlogMessage Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
		   If KS_A_RS_Obj.Eof And KS_A_RS_Obj.Bof Then
			KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		    Response.Write "<script>alert('参数出错!');history.back();</script>"
			Response.end
		   End If
		   KS_A_RS_Obj("readtf")=1
		   KS_A_RS_Obj.update
		   Dim Title:Title=KS_A_RS_Obj("Title")
		   Dim Content:Content=KS_A_RS_Obj("Content")
		   Dim Replay:Replay=KS_A_RS_Obj("Replay"):If IsNull(Replay) Then Replay=""
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		%>
		<script language = "JavaScript">
			function CheckForm()
			{
				if (FCKeditorAPI.GetInstance('Replay').GetXHTML(true)=="")
					{
					  alert("请输入回复内容！");
					  FCKeditorAPI.GetInstance('Replay').Focus();
					  return false;
					}
				
				 return true;  
			}
		</script>
				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form action="?Action=SaveMessageReplay&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>回 复 留 言</td>
					</tr>

                      <tr class="tdbg">
                           <td  height="25" align="center"><span>留言标题：</span></td>
                              <td>  <input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                    </tr>
							 
                              <tr class="tdbg">
                                  <td  height="25" align="center"><span>留言内容：</span></td>
                                  <td><input type="hidden" value="<%=Server.HtmlEncode(Content)%>" name="Content" id="Content">
                                <iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic" width="98%" height="150" frameborder="0" scrolling="no"></iframe>     </td>
                            </tr>
                              <tr class="tdbg">
                                  <td  height="25" align="center"><span>回复内容：</span></td>
                                  <td><input type="hidden" value="<%=Server.HtmlEncode(Replay)%>" name="Replay" id="Replay">
                              <iframe id="Replay___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Replay&amp;Toolbar=Basic" width="98%" height="150" frameborder="0" scrolling="no"></iframe></td>
                            </tr>
								
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,立即回复 " />
                      <input type="reset" name="Submit2"   class="Button" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
	   End Sub		
		
		
	   
	   '保存留言回复
	   Sub SaveMessageReplay()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
		Dim Title:Title=KS.S("Title")
		Dim Content:Content=Request.Form("Content")
		Dim Replay:Replay=Request.Form("Replay")
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_BlogMessage Where ID=" & ID,conn,1,3
		If Not RS.Eof Then
		  RS("Title")=Title
		  RS("Content")=Content
		  RS("Replay")=Replay
		  RS("ReplayDate")=Now
		 RS.Update
		End If
		RS.Close:Set RS=Nothing
		Call KSUser.AddLog(KSUser.UserName,"回复了留言操作!内容:" & Content & "回复:" & Replay & "",100)
		Response.Write "<script>alert('恭喜,您已成功回复！');location.href='?Action=Message';</script>"
	   End Sub
	   '删除留言
	  Sub MessageDel()
		Dim ID:ID=KS.S("ID")
		ID=KS.FilterIDs(ID)
		If ID="" Then Call KS.Alert("你没有选中要删除的留言!",ComeUrl):Response.End
		Conn.Execute("Delete From KS_BlogMessage Where UserName='" & KSUser.UserName & "' and ID In(" & ID & ")")
		Call KSUser.AddLog(KSUser.UserName,"删除了用户的留言操作!",100)
		Response.Redirect ComeUrl
	  End Sub
	  
		
		
		'发送信息
		Sub sendMessage()
			dim SendTime,title,content
			If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then
				Set rs=server.createobject("adodb.recordSet")
				SqlStr="Select SendTime,title,content from KS_Message where Incept='"&KSUser.UserName&"' and id="&Clng(KS.S("ID"))
				RS.open SqlStr,Conn,1,1
				If not(RS.eof and RS.bof) Then
					SendTime=rs("SendTime")
					Title="RE " & rs("title")
					Content=rs("content")
				End If
				RS.close
				Set rs=Nothing
			End If
		%>
		<table width="98%" align="center" cellpadding="3" cellspacing="1" class="border">
				<form action="User_Message.asp"  name="myform" method="post" id="myform" onSubmit="return CheckForm();">
				  <tr> 
					<td colspan=2 align=center class="Title">
					发送短消息
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td width="100" align="right" valign=middle><b>收件人：</b></td>
					<td valign=middle>
					  <input type=hidden name="action" value="sEnd">
					  <input class="textbox" type=text name="Touser" value="<%=KS.S("Touser")%>" size=60>
					  <Select class="textbox" name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
					  <OPTION selected value="">选择</OPTION>
						<%
						Set rs=server.createobject("adodb.recordSet")
						SqlStr="Select friend from KS_Friend where Username='"&KSUser.UserName&"' order by Addtime desc"
						RS.open SqlStr,Conn,1,1
						Do While not RS.eof
						%>
						<OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION>
						<%
						RS.movenext
						loop
						RS.close:Set rs=Nothing
						%>
					  </Select>
					  <a href="User_Friend.asp?action=addF">添加好友</a>
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td width="100" align="right" valign=top><b>标　题：</b></td>
					<td valign=middle>
					  <input class="textbox" type=text name="title" size=70 maxlength=90 value="<%=title%>">
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td width="100" align="right" valign=top><b>内　容：</b></td>
					<td valign=middle>
					  <textarea cols=50 rows=16 name="message" style="display:none" title="Ctrl+Enter发送">
					  <%If KS.S("ID")<>"" Then%>
						============= 在 <%=SendTime%> 您来信中写道： ==============<br>
						<%=server.htmlencode(content)%>
						<br>=======================================================
					<%End If%>
		             </textarea>
					 <iframe id="message___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=message&amp;Toolbar=Basic" width="93%" height="200" frameborder="0" scrolling="no"></iframe>
					
					</td>
				  </tr>
				  <tr class='tdbg'> 
				    <td align="right"><b>说明</b>：</td>
					<td colspan=2>
		① 可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=max_sEnd%></b>个用户<br>
		② 标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br>
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td valign=middle colspan=2 align=center> 
					  <input class="Button" type=Submit value=" 发 送 " name=Submit>
					  &nbsp; 
					  <input class="Button" type=Submit value=" 保 存 " name=Submit>
					  &nbsp; 
					  <input class="Button" type="reSet" name="Clear" value=" 清 除 ">
					  &nbsp; 
		<%If request("reaction")="chatlog" Then%>
					  <input  class="Button" type=button value="关闭聊天记录" name="chatlog" onClick="location.href='?action=new&id=<%=KS.S("ID")%>&Touser=<%=KS.S("Touser")%>'">
		<%Else
		    If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then      
		     %>
					  <input class="Button" type=button value="查看聊天记录" name="chatlog" onClick="location.href='?action=new&id=<%=KS.S("ID")%>&Touser=<%=KS.S("Touser")%>&reaction=chatlog'">
		  <%Else%>
					  <input class="Button" type=button value="查看聊天记录" name="chatlog" disabled>
		<% End IF
		End If%>
					  &nbsp; 
					  <input class="Button" type="button" name="close" value=" 关 闭 " onClick="window.close()">
					</td>
				  </tr>
		<%If request("reaction")="chatlog" Then%>
				  <tr Class=title> 
					<td colspan=3>我与<%=KS.S("Touser")%>的聊天记录</td>
				  </tr>
		<%If KSUser.UserName=KS.S("Touser") Then%>
				  <tr> 
					<td colspan=3>自己跟自己的聊天记录没什么好看的：）</td>
				  </tr>
		<%Else%>
		<%
			Set rs=server.createobject("adodb.recordSet")
			SqlStr="Select * from KS_Message where ((Incept='"&KSUser.UserName&"' and Incept='"&replace(KS.S("Touser"),"'","")&"') or (sEnder='"&replace(KS.S("Touser"),"'","")&"' and Incept='"&KSUser.UserName&"')) and delS=0 order by SendTime desc"
			RS.open SqlStr,Conn,1,1
			If RS.eof and RS.bof Then
		%>
				  <tr> 
					<td colspan=3>还没有任何聊天记录！</td>
				  </tr>
		<%
			Else
			Do While not RS.eof
		%>
						<tr>
							<td height=25 colspan=3>
		<%If rs("sEnder")=KSUser.UserName Then%>
							在<b><%=rs("SendTime")%></b>，您发送此消息给<b><%=KS.HTMLcode(rs("Incept"))%></b>！
		<%Else%>
					在<b><%=rs("SendTime")%></b>，<b><%=KS.HTMLcode(rs("sEnder"))%></b>给您发送的消息！
		<%End If%></td>
						</tr>
						<tr>
							<td valign=top align=left colspan=2>
							<b>消息标题：<%=KS.HTMLcode(rs("title"))%></b><hr size=1>
							<%=KS.HTMLcode(rs("content"))%>
					</td>
						</tr>
		<%
			RS.movenext
			loop
			End If
			RS.close:Set rs=Nothing
		%>
		<%End If%>
		<%End If%>
				</form>
</table>
		<%
			DoTitleJs
		End Sub
		'读取信息
		Sub read()
			If KS.S("id")=0 Then
				Response.Write "<script>alert('请指定正确的参数。');history.back();</script>"
			End If
			Set rs=server.createobject("adodb.recordSet")
			If request("action")="read" Then
				Conn.Execute("Update KS_Message Set flag=1 where ID="&Clng(KS.S("id")))
			End If
			SqlStr="Select * from KS_Message where (Incept='"&KSUser.UserName&"' or sEnder='"&KSUser.UserName&"') and id="&Clng(KS.S("ID"))
			RS.open SqlStr,Conn,1,1
			If RS.eof and RS.bof Then
				RS.close:Set rs=Nothing
				Response.Write "<script>alert('你是不是跑到别人的信箱啦、或者该信息已经被收件人删除。');history.back();</script>"
			Else
		%>
		<table width="98%" align=center cellpadding=3 cellspacing=1>
					<tr>
						<th class="tdbg" colspan=3>欢迎使用短消息接收，<%=KSUser.UserName%></th>
					</tr>
					<tr>
						<td valign=middle align=center colspan=3><a href="User_Message.asp?action=delet&id=<%=rs("id")%>&ComeUrl=<%=ComeUrl%>"><img src="images/delete.gIf" border=0 alt="删除消息"></a> &nbsp; <a href="User_Message.asp?action=new"><img src="images/write.gIf" border=0 alt="发送消息"></a> &nbsp;<a href="User_Message.asp?action=new&Touser=<%=KS.HTMLEncode(rs("sEnder"))%>&id=<%=KS.S("ID")%>"><img src="images/reply.gIf" border=0 alt="回复消息"></a>&nbsp;<a href="User_Message.asp?action=fw&id=<%=KS.S("ID")%>"><img src="images/fw.gIf" border=0 alt=转发消息></a></td>
					</tr>
						<tr>
							<td height=25>
		<%If request("action")="outread" Then%>
							在<b><%=rs("SendTime")%></b>，您发送此消息给<b><%=KS.HTMLEncode(rs("Incept"))%></b>！
		<%Else%>
					在<b><%=rs("SendTime")%></b>，<b><%=KS.HTMLEncode(rs("sEnder"))%></b>给您发送的消息！
		<%End If%></td>
						</tr>
						<tr>
							<td valign=top align=left>
							<b>消息标题：<%=KS.HTMLencode(rs("title"))%></b><hr size=1>
							<%=rs("content")%>
					</td>
						</tr>
		<%
			RS.close:Set rs=Nothing
			SqlStr="Select id,sEnder from KS_Message where Incept='"&KSUser.UserName&"' and flag=0 and IsSend=1 and id>"&KS.ChkClng(KS.S("ID")&" order by SendTime")
			Set rs=Conn.Execute(SqlStr)
			If not (RS.eof and RS.bof) Then
		%>
						<tr>
							<td valign=top align=right><a href=User_Message.asp?action=read&id=<%=rs(0)%>&sEnder=<%=rs(1)%>>[读取下一条信息]</a>
					</td>
						</tr>
		<%
		End If
		RS.close:Set rs=Nothing
		%>
</table>
		<%
			End If
		End Sub
		'转发信息
		Sub fw()
			dim title,content,sEnder
			If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then
				Set rs=server.createobject("adodb.recordSet")
				SqlStr="Select title,content,sEnder from KS_Message where (Incept='"&KSUser.UserName&"' or sEnder='"&KSUser.UserName&"') and id="&Clng(KS.S("ID"))
				RS.open SqlStr,Conn,1,1
				If RS.eof and RS.bof Then
					RS.close:Set rs=Nothing
					Response.Write "<script>alert('请指定正确的参数。');history.back();</script>"
				Else
				title=rs("title"):content=rs("content"):sEnder=rs("sEnder")
				End If
				RS.close:Set rs=Nothing
			End If
		%>
		<table width="100%" align=center cellpadding=3 cellspacing=1 class=border>
				<form action="User_Message.asp"  name="myform" method="post" id="myform" onSubmit="return CheckForm();">
				  <tr class="Title"> 
					<td colspan=2 align=center height=25>
					 发送短消息
				    </td>
				  </tr>
				  <tr class='tdbg'> 
					<td valign=middle width=100><b>收件人：</b></td>
					<td valign=middle>
					  <input type="hidden" name="action" value="sEnd">
					  <input class='textbox' type=text name="Touser" value="<%=KS.S("Touser")%>" size=70>
					  <Select name="font" onChange="DoTitle(this.options[this.selectedIndex].value)">
					  <OPTION selected value="">选择</OPTION>
						<%
						Set rs=server.createobject("adodb.recordSet")
						SqlStr="Select friend from KS_Friend where Username='"&KSUser.UserName&"' order by Addtime desc"
						RS.open SqlStr,Conn,1,1
						Do While not RS.eof
						%>
						<OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION>			
						<%
						RS.movenext
						loop
						RS.close:Set rs=Nothing
						%>
					  </Select>
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td width="100" valign=top><b>标　题：</b></td>
					<td valign=middle>
					  <input class='textbox' type=text name="title" size=80 maxlength=90 value="Fw：<%=title%>">&nbsp;
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td width="100" valign=top><b>内　容：</b></td>
					<td valign=middle>
					  <textarea cols=76 rows=16 name="message" id="message" title="Ctrl+Enter发送" style="display:none">
		
		
		======================== 下面是转发信息 =====================<br>
		原发件人：<%=sEnder%><br>
		<%=server.htmlencode(content)%>
		=======================================================</textarea>
		<iframe id="message___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=message&amp;Toolbar=Basic" width="93%" height="200" frameborder="0" scrolling="no"></iframe>
		
	
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td colspan=2>
		<b>说明</b>：<br>
		① 可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=max_sEnd%></b>个用户<br>
		② 标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br>
					</td>
				  </tr>
				  <tr class='tdbg'> 
					<td valign=middle colspan=2 align=center> 
					  <input class="Button" type=Submit value=" 发 送 " name=Submit>
					  &nbsp; 
					  <input class="Button" type=Submit value=" 保 存 " name=Submit>
					  &nbsp; 
					  <input class="Button" type="reSet" name="Clear" value=" 清 除 ">
					  &nbsp; 
					  <input class="Button" type="button" name="close" value=" 关 闭 " onClick="window.close()">
					</td>
				  </tr>
			　</form>
		</table>
		<%
			DoTitleJs
		End Sub
		
		Sub savemsg()
			dim Incept,title,message,Subtype,i,sUname
			If KS.S("Touser")="" Then
				Response.Write("<script>alert('您忘记填写发送对象了吧。');history.back();</script>")
			Else
				Incept=KS.S("Touser")
				Incept=split(Incept,",")
			End If
			If KS.S("Title")="" Then
				Response.Write("<script>alert('您还没有填写标题呀。');history.back();</script>")
			ElseIf KS.strLength(KS.S("title"))>50 Then
				Response.Write("<script>alert('标题限定最多50个字符。');history.back();</script>")
			Else
				title=KS.S("title")
			End If
			If KS.S("Message")="" Then
				Response.Write("<script>alert('内容是必须要填写的噢。');history.back();</script>")
			ElseIf KS.strLength(KS.S("Message"))>Cint(max_sms) Then
				Response.Write("<script>alert('内容限定最多"&max_sms&"个字符。');history.back();</script>")
			Else
				message=KS.S("message")
			End If
		
			for i=0 to ubound(Incept)
				sUname=replace(Incept(i),"'","")
				SqlStr="Select UserName from KS_User where UserName='"&sUname&"'"
				Set rs=Conn.Execute(SqlStr)
				If RS.eof and RS.bof Then
					RS.close:Set rs=Nothing
					call KS.AlertHistory("系统没有这个用户，看看你的发送对象写对了嘛？",-1)
					response.end
				End If
				RS.Close
				rs.open "select username from ks_friend where username='" & sUname & "' and friend='" & ksuser.username & "' and flag=3",conn,1,1
				if not rs.eof then
					RS.close:Set rs=Nothing
					call KS.AlertHistory("对不起，你被" & sUname & "列为黑名单,不能发送短信给他！",-1)
					response.end
				end if
				RS.close:Set rs=Nothing
						
				Select Case KS.S("Submit")
				Case "发送"
					SqlStr="insert into KS_Message (Incept,sEnder,title,content,SendTime,flag,IsSend,DelR,DelS) values ('"&sUname&"','"&KSUser.UserName&"','"&title&"','"&message&"','"&Now()&"',0,1,0,0)"
					Subtype="已发送信息"
				Case "保存"
					SqlStr="insert into KS_Message (Incept,sEnder,title,content,SendTime,flag,IsSend,DelR,DelS) values ('"&sUname&"','"&KSUser.UserName&"','"&title&"','"&message&"','"&Now()&"',0,0,0,0)"
					Subtype="发件箱"
				Case Else
					SqlStr="insert into KS_Message (Incept,sEnder,title,content,SendTime,flag,IsSend,DelR,DelS) values ('"&sUname&"','"&KSUser.UserName&"','"&title&"','"&message&"','"&Now()&"',0,1,0,0)"
					Subtype="已发送信息"
				End Select
				
				'判断对方信箱是否已满
				If conn.execute("select count(*) from KS_Message where Incept='"&sUname&"'")(0)>=Max_Num Then
					Response.Write("<script>alert('由于[" & sUname & "]的信箱已满，发送没有成功！');history.back();</script>")
				Else
				   Conn.Execute(SqlStr)
				End If

				
				If i>Cint(max_sEnd)-1 Then
					Response.Write("<script>alert('最多只能发送给"&max_sEnd&"个用户，您的名单"&max_sEnd&"位以后的请重新发送');history.back();</script>")
					exit for
				End If
			next
		Response.Write("<script>alert('恭喜您，发送短信息成功。发送的消息同时保存在您的"&Subtype&"中。');location.href='User_Message.asp';</script>")
		
		End Sub
		
		'更改信息
		Sub edit()
			dim Incept,title,content,id
			If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then
				Set rs=server.createobject("adodb.recordSet")
				SqlStr="Select id,Incept,title,content from KS_Message where sEnder='"&KSUser.UserName&"' and IsSend=0 and id="&Clng(KS.S("ID"))
				RS.open SqlStr,Conn,1,1
				If not(RS.eof and RS.bof) Then
					Incept=rs("Incept"):title=rs("title"):content=rs("content"):id=rs("id")
				Else
					Response.Write("<script>alert('没有找到您要编辑的信息。');history.back();</script>")
				End If
				RS.close:Set rs=Nothing
			Else
				Response.Write("<script>alert('请指定相关参数。');history.back();</script>")
			End If
		%>
			<table width="100%" align=center cellpadding=3 cellspacing=1 class=border>
			<form action="User_Message.asp" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				  <tr> 
					<th colspan=2 height=25> 
					  <input type=hidden name="action" value="savedit"> 
					  <input type=hidden name="id" value="<%=id%>">
					  发送短消息--请完整输入下列信息</th>
				  </tr>
				  <tr> 
					<td valign=middle><b>收件人：</b></td>
					<td valign=middle>
					  <input type=text name="Touser" value="<%=Incept%>" size=80>
					</td>
				  </tr>
				  <tr> 
					<td valign=top><b>标题：</b></td>
					<td valign=middle>
					  <input type=text name="title" size=80 maxlength=80 value="<%=title%>">
					</td>
				  </tr>
				  <tr> 
					<td valign=top><b>内容：</b></td>
					<td valign=middle>
					  <input type="hidden" value="<%=server.htmlencode(content)%>" name="message" title="" style="display:none">
					  <iframe id='MessageContent' name='MessageContent' src='Editor.asp?ID=message&style=0&ChannelID=9998' frameborder=0 scrolling=no width='100%' height='280'></iframe>
					</td>
				  </tr>
				  <tr> 
					<td colspan=2>
		<b>说明</b>：<br>
		① 您可以使用<b>Ctrl+Enter</b>键快捷发送短信<br>
		② 标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br>
					</td>
				  </tr>
				  <tr> 
					<td valign=middle colspan=2 align=center> 
					  <input class="Button" type=Submit value=" 发 送 " name="Submit">
					  &nbsp; 
					  <input class="Button" type=Submit value=" 保 存 " name="Submit">
					  &nbsp; 
					  <input class="Button" type="reSet" name="Clear" value=" 清 除 ">
					  &nbsp; 
					  <input class="Button" type="button" name="close" value=" 关 闭 " onClick="window.close()">
					</td>
				  </tr>
			  </form>
</table>
			  </td>
			</tr>
			
</table>
		
		<%
		End Sub
		
		Sub savedit()
			dim Incept,title,message,Subtype
			If KS.S("ID")="" or not isNumeric(KS.S("ID")) Then
				Response.Write("<script>alert('请指定相关参数。');history.back();</script>")
			End If
			If KS.S("Touser")="" Then
				Response.Write("<script>alert('您忘记填写发送对象了吧。');history.back();</script>")
			Else
				Incept=KS.S("Touser")
			End If
			If KS.S("Title")="" Then
				Response.Write("<script>alert('您还没有填写标题呀!');history.back();</script>")
			Else
				title=KS.S("title")
			End If
			If KS.S("Message")="" Then
			   Response.Write("<script>alert('内容是必须要填写的噢!');history.back();</script>")
			Else
				message=KS.S("message")
			End If
		
			SqlStr="Select UserName from KS_User where UserName='"&Incept&"'"
			Set rs=Conn.Execute(SqlStr)
			If RS.eof and RS.bof Then
				Set rs=Nothing
				Response.Write("<script>alert('系统没有这个用户，看看你的发送对象写对了嘛？');history.back();</script>")
			End If
			Set rs=Nothing
		
			If KS.S("Submit")="发送" Then
				SqlStr="Update KS_Message Set Incept='"&Incept&"',sEnder='"&KSUser.UserName&"',title='"&title&"',content='"&message&"',SendTime=Now(),flag=0,IsSend=1 where id="&Clng(KS.S("ID"))
				Subtype="已发送信息"
			Else
				SqlStr="Update KS_Message Set Incept='"&Incept&"',sEnder='"&KSUser.UserName&"',title='"&title&"',content='"&message&"',SendTime=Now(),flag=0,IsSend=0 where id="&Clng(KS.S("ID"))
				Subtype="发件箱"
			End If
			Set rs=Conn.Execute(SqlStr)
		   
		   Response.Write("<script>alert('恭喜您，发送短信息成功。发送的消息同时保存在您的"&Subtype&"中。');location.href='User_Message.asp';</script>")
		  
		End Sub
		
		'收件置于回收站，参数字段delR，可用于批量及单个删除
		Sub delinbox()
			dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
				Response.Write "<script>alert('请选择相关参数!');history.go(-1);</script>"
				Exit Sub
			Else
				Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' and id in ("&DelID&")")
				Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
			
			End If
		End Sub
		
		Sub AllDelinbox()
			Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' and delR=0")
			Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
			Response.End
		End Sub
		
		'发件逻辑删除，置于回收站，入口字段delS，可用于批量及单个删除
		Sub deloutbox()
			dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
				Response.Write "<script>alert('请选择相关参数!');history.go(-1);</script>"
			Else
				Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' and IsSend=0 and id in ("&DelID&")")
				Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
				Response.End
			End If
		End Sub
		
		Sub AllDeloutbox()
			Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' and delS=0 and IsSend=0")
			Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
			Response.End
		End Sub
		
		'已发送置于回收站，入口字段delS，可用于批量及单个删除
		'delS：0未操作，1发送者删除，2发送者从回收站删除
		Sub DelIsSend()
			dim DelID
			DelID=KS.S("ID")
			'Response.Write delid
			'Response.End()
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(replace(Replace(DelID,",","")," ","")) Then
				Response.Write "<script>alert('请选择相关参数!');history.go(-1);</script>"
			Else
				Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' and IsSend=1 and id in ("&DelID&")")
				Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
				Response.End
			End If
		End Sub
		
		Sub AllDelIsSend()
			Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' and delS=0 and IsSend=1")
			Response.Write "<script>alert('删除短信息成功。删除的消息将转移到您的回收站!');location.href='" & ComeUrl & "';</script>"
			Response.End
		End Sub
		
		'用户能完全删除收到信息和逻辑删除所发送信息，逻辑删除所发送信息设置入口字段delS参数为2
		Sub delrecycle()
			dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
				Response.Write "<script>alert('请选择相关参数!');history.go(-1);</script>"
			Else
				Conn.Execute("delete from KS_Message where Incept='"&KSUser.UserName&"' and id in ("&DelID&")")
				Conn.Execute("Update KS_Message Set delS=2 where Sender='"&KSUser.UserName&"' and delS=1 and id in ("&DelID&")")          
			Response.Write "<script language=""javascript"">alert('删除短信息成功。删除的消息将不可恢复');location.href='"&ComeUrl&"';</script>"
		    Response.End

				
			End If
		End Sub
		Sub AllDelrecycle()
			Conn.Execute("delete from KS_Message where Incept='"&KSUser.UserName&"' and delR=1")	
			Conn.Execute("Update KS_Message Set delS=2 where Sender='"&KSUser.UserName&"' and delS=1")
			Response.Write "<script language=""javascript"">alert('删除短信息成功。删除的消息将不可恢复');location.href='"&ComeUrl&"';</script>"
			Response.End
		End Sub
		
		Sub delete()
			dim DelID
			DelID=KS.S("id")
			ComeUrl=Request("ComeUrl")
			'Response.End()
			If ComeUrl="" Then ComeUrl="User_Message.asp"
			If not isNumeric(DelID) or DelID="" or isnull(DelID) Then
				Response.Write "<script>alert('请选择相关参数!');history.go(-1);</script>"
			Else
				Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' and id="&Clng(DelID))
				Conn.Execute("Update KS_Message Set delS=1 where sEnder='"&KSUser.UserName&"' and id="&Clng(DelID))
				Response.Write "<script language=""javascript"">alert('删除短信息成功。删除的消息将置于您的回收站内。');location.href='"&ComeUrl&"';</script>"
				Response.End
			End If
		End Sub
		
		Sub MessageMain()
		   currentpage=ks.chkclng(request("page"))
		   if currentpage=0 then currentpage=1
			dim SqlStr,boxName,smstype,readaction,turl
			dim keyword,param
			keyword=KS.S("KeyWord")
			if keyword<>"" then
			  if ks.s("searcharea")=1 then
			   param=" and title like '%" & keyword & "%'"
			  else
			   param=" and content like '%" & keyword & "%'"
			  end if
			end if
			Select Case Action
			Case "inbox"
				boxName="收件箱":smstype="inbox":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where Incept='"&KSUser.UserName&"'" & param & " and IsSend=1 and delR=0 order by flag,SendTime desc"
			Case "outbox"
				boxName="草稿箱":smstype="outbox":readaction="edit":turl="sms"
				SqlStr="select * from KS_Message where Sender='"&KSUser.UserName&"'" & param & " and IsSend=0 and delS=0 order by SendTime desc"
			Case "issend"
				boxName="已发送的消息":smstype="IsSend":readaction="outread":turl="readsms"
				SqlStr="select * from KS_Message where Sender='"&KSUser.UserName&"'" & param & " and IsSend=1 and delS=0 order by SendTime desc"
			Case "recycle"
				boxName="垃圾箱":smstype="recycle":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where ((Sender='"&KSUser.UserName&"'" & param & " and delS=1) or (Incept='"&KSUser.UserName&"' and delR=1)) and not delS=2 order by SendTime desc"
			Case Else
				boxName="收件箱":smstype="inbox":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where Incept='"&KSUser.UserName&"'" & param & " and IsSend=1 and delR=0 order by flag,SendTime desc"
			End Select
		Call KSUser.InnerLocation("我的" & boxname)
		%>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1"  class="border">
		<form action="User_Message.asp" method="post" name="inbox">
		<tr height='23' class="Title">
		<td align="center">已读</td>
		<td align="center">主题</td>
		<td  height="26" align="center">
		<%if smstype="inbox" or smstype="recycle" then Response.Write "发件人" else Response.Write "收件人"%></td>
		<td align="center">日期</td>
		<td align="center">大小</td>
		<td align="center">操作</td>
		</tr>
		<%
			Dim RS:Set RS=server.createobject("adodb.recordset")
			OpenConn
			RS.open SqlStr,Conn,1,1
			if RS.eof and RS.bof then
		%>
		<tr>
		<td colspan=6 align=center valign=middle class='tdbg'>您的<%=boxname%>中没有任何内容。</td>
		</tr>
		<%else
		
		         totalPut = RS.RecordCount
						
				   If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
		 dim i:i=0
		
		
		Do While not RS.eof
		%>
		<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		<td height="25" align=center valign=middle>
		<%
		select case smstype
		case "inbox"
			if rs("flag")=0 then
				Response.Write "<img src=""images/news.gif"">"
			else
				Response.Write "<img src=""images/olds.gif"">"
			end if
		case "outbox"
			Response.Write "<img src=""images/IsSend_2.gif"">"
		Case "issend"
			Response.Write "<img src=""images/IsSend_1.gif"">"
		case "recycle"
			if rs("flag")=0 then
				Response.Write "<img src=""images/news.gif"">"
			else
				Response.Write "<img src=""images/olds.gif"">"
			end if
		end select
		%>
		</td>
		<td height="25" align=left><a href="User_Message.asp?action=<%=readaction%>&id=<%=rs("id")%>&sender=<%=rs("sender")%>"><%=KS.HTMLEncode(rs("title"))%></a>	</td>
		<td height="25" align=center valign=middle>
		<%if smstype="inbox" or smstype="recycle" then%>
		<%=KS.HTMLEncode(rs("sender"))%>
		<%else%>
		<%=KS.HTMLEncode(rs("Incept"))%>
		<%end if%>
		</td>
		<td height="25"><%=formatdatetime(rs("SendTime"),2)%></td>
		<td height="25"><%=len(rs("content"))%>Byte</td>
		<td width=30 height="25" align=center valign=middle><input type=checkbox name=id value=<%=rs("id")%>></td>
		</tr>
		<%
		  i=I+1
		 if i>maxperpage or rs.eof then exit do
			RS.movenext
			loop
			end if
			RS.close:set rs=Nothing
		%>
		<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'"> 
		<td height="26" colspan=6 align=right valign=middle>节省每一分空间，请及时删除无用信息&nbsp;
		  <input type=checkbox name=chkall value=on onClick="CheckAll(this.form)">选中所有显示记录&nbsp;<input class="button" type=submit name=action onClick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除<%=replace(boxname,"箱","")%>">&nbsp;
		  <input type=submit class="button" name=action onClick="{if(confirm('确定清除<%=boxname%>所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空<%=boxname%>"></td>
		</tr>
		</form>
		<tr>
		<td colspan=6>
		 <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>

		</td>
		</tr>
</table>
		<br>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1"  class="border">
		<form action="User_Message.asp" method="post" name="myform">
		 <tr>
		  <td><strong>短消息搜索:</strong>
		   <select name="action">
		    <option value="inbox"<%if ks.s("action")="inbox" then response.write " selected"%>>收件箱</option>
		    <option value="outbox"<%if ks.s("action")="outbox" then response.write " selected"%>>发件箱</option>
		    <option value="issend"<%if ks.s("action")="issend" then response.write " selected"%>>已发送</option>
		    <option value="recycle"<%if ks.s("action")="recycle" then response.write " selected"%>>废件箱</option>
		   </select>
		   <select name="searcharea">
		    <option value="1">短消息主题</option>
			<option value="2">短消息内容</option>
		   </select>
		   <input type="text" name="keyword" value="关键字" onFocus="this.value='';" onBlur="if (this.value=='') this.value='关键字';">
		   <input type="submit" value=" 搜 索 " name="submit1" class="button">
		  </td>		  
		 </tr>
		 </form>
		</table>
		<script language=javascript>
		function CheckAll(form)
		{
		for (var i=0;i<form.elements.length;i++)    {
		var e = form.elements[i];
		if (e.name != 'chkall')       e.checked = form.chkall.checked; 
		}
		}
		</script>
		<%
		end sub
		
		Sub DoTitleJs()
		%>
		<script language="javascript"> 
		function DoTitle(addTitle) {  
		 var revisedTitle;  
		 var currenttitle = document.myform.Touser.value; 
		
		 if(currenttitle=="") revisedTitle = addTitle; 
		 else { 
		  var arr = currenttitle.split(","); 
		  for (var i=0; i < arr.length; i++) { 
		   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
		  } 
		  revisedTitle = currenttitle+","+addTitle; 
		 } 
		
		 document.myform.Touser.value=revisedTitle;  
		 document.myform.Touser.focus(); 
		 return; 
		} 
		function document.onkeydown()
		{
			if(event.ctrlKey && window.event.keyCode==13)
			{
				CheckForm();
			}
			
		}
		</script>
		<%
		End Sub
		 '（图片对象名称，标题对象名称，更新数，总数）
		Function ShowTable(SrcName,TxtName,str,c)
		Dim Tempstr,Src_js,Txt_js,TempPercent
		If C = 0 Then C = 99999999
		Tempstr = str/C
		TempPercent = FormatPercent(tempstr,0,-1)
		Src_js = "document.getElementById(""" + SrcName + """)"
		Txt_js = "document.getElementById(""" + TxtName + """)"
			ShowTable = VbCrLf + "<script>"
			ShowTable = ShowTable + Src_js + ".width=""" & FormatNumber(tempstr*600,0,-1) & """;"
			ShowTable = ShowTable + Src_js + ".title=""容量上限为："&c&"条，总共已储存（"&str&"）条站内短信！"";"
			ShowTable = ShowTable + Txt_js + ".innerHTML="""
			If FormatNumber(tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable + "已使用:" & TempPercent & """;"
			Else
				ShowTable = ShowTable + "<font color=\""red\"">已使用:" & TempPercent & ",请赶快清理！</font>"";"
			End If
			ShowTable = ShowTable + "</script>"
		End Function

End Class
%> 

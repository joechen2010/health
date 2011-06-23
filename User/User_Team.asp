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
Set KSCls = New User_Team
KSCls.Kesion()
Set KSCls = Nothing

Class User_Team
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather,PhotoUrls,Note
		Private XCID,Title,Tags,UserName,Face,Content,Verific,PicUrl,Action,I,ClassID,Point
		Private Sub Class_Initialize()
		  MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		If KS.SSetting(0)=0 Then
		 Call KS.Alert("对不起，本站关闭个人空间功能！","")
		 Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("圈子管理")
		KSUser.CheckPowerAndDie("s06")
		%>
		
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">圈子首页</a></li>
				<li<%If KS.S("Action")="MyTeam" then response.write " class='select'"%>><a href="User_Team.asp?Action=MyTeam">我建的圈子</a></li>
				<li<%If KS.S("Action")="MyJoinTeam" then response.write " class='select'"%>><a href="User_Team.asp?Action=MyJoinTeam">我加入的圈子</a></li>
				<li><a href="User_Blog.asp?Action=Template&Flag=3">圈子模板</a></li>
			</ul>
		</div>
			 <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="?action=CreateTeam"><span style="font-size:14px;color:#ff3300">创建圈子</span></a></div>
		<%

			Select Case KS.S("Action")
			 Case "MyTeam"
			  Call MyTeam()
			 Case "MyJoinTeam"
			  Call MyJoinTeam()
			 Case "ReplayComment"
			  Call ReplayComment()
			 Case "SaveCommentReplay"
			  Call SaveCommentReplay()
			 Case "CommentDel"
			  Call CommentDel()
			 Case "VerificApply"
			  Call VerificApply()
			 Case "AcceptApply"
			  Call AcceptApply()
			 Case "ApplyDel" '拒绝申请
			  Call ApplyDel()
			 Case "Del"
			  Call DelTeam()
			 Case "EditTeam","CreateTeam"
			  Call Managexc()
			 Case "Teamsave"
			  Call Teamsave()
			 Case "OutTeam"
			  Call OutTeam()
			 Case Else
			  Call TeamIndex()
			End Select
	   End Sub
	
	    '圈子，添加／修改
	   Sub Managexc()
	    Dim TeamName,ClassID,Note,PhotoUrl,Point,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,JoinTF,Announce
		Dim ID:ID=KS.ChkCLng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Team Where ID=" & ID,conn,1,1
		If Not RS.EOF Then
		Call KSUser.InnerLocation("修改圈子")
		 TeamName=RS("TeamName")
		 ClassID=RS("ClassID")
		 Note=RS("Note")
		 JoinTF=RS("JoinTF")
		 PhotoUrl=RS("PhotoUrl")
		 Point=RS("Point")
		 Announce=RS("Announce")
		 OpStr="OK了，确定修改":TipStr="修 改 我 的 圈 子"
		Else
	   	 if KS.ChkClng(ks.SSetting(6))<>0 then
		   if conn.execute("select count(id) from ks_team where username='"& ksuser.username & "'")(0)>=KS.ChkClng(ks.SSetting(6)) then
		   rs.close:set rs=nothing
		    response.write "<script>alert('对不起，每个用户最多只能建 " & KS.SSetting(6) & " 个圈子!');history.back();</script>"
		    response.end
		   end if
		  end if
		 Call KSUser.InnerLocation("创建圈子")
		 Point="10"
		 ClassID="0"
		 JoinTF="1"
		 Announce="暂无公告!"
		 PhotoUrl="../images/user/group.gif"
		 OpStr="OK了，立即创建":TipStr="创 建 我 的 圈 子"
		End if
		RS.Close:Set RS=Nothing
	    %>
		<script>
		 function CheckForm()
		 {
		  if (document.myform.TeamName.value=='')
		  {
		   alert('请输入圈子名称!');
		   document.myform.TeamName.focus();
		   return false;
		  }
		  if (document.myform.ClassID.value=='0')
		  {
		   alert('请选择圈子类型!');
		   document.myform.ClassID.focus();
		   return false;
		  }
		  return true;
		 }

		</script>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
          <form  action="User_Team.asp?Action=Teamsave&id=<%=id%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
            <tr class="title">
              <td colspan=2 align=center><%=TipStr%></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>圈子名称：</strong><br>
              请给你的圈子取个合适的名称。
              </div></td>
              <td> 　
                  <input class="textbox" name="TeamName" type="text" id="TeamName" style="width:250px; " value="<%=TeamName%>" maxlength="100" />
              <span style="color: #FF0000">*</span></td>
            </tr>
<tr class="tdbg">
              <td width="24%"  height="25" align="center"><div align="left"><strong>圈子分类：</strong><br>
      圈子分类，以便查找浏览</div></td>
              <td width="76%">　
                  <select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-请选择类别-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_TeamClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select>               </td>
            </tr>
			<tr class="tdbg"> 
                  <td height="30"><div align="left"><strong>加入条件：</strong><br>
                  在这里设置会员加入到此圈子的条件。</div></td>
                  <td>&nbsp;
                       <input type="radio" value="1" name="JoinTF"<%if JoinTF="1" then response.write " checked"%>>任意加入
                       <input type="radio" value="2" name="JoinTF"<%if JoinTF="2" then response.write " checked"%>>申请加入
                       <input type="radio" value="3" name="JoinTF"<%if JoinTF="3" then response.write " checked"%>>仅可邀请
                       <br>&nbsp;&nbsp;<input type="radio" value="4" name="JoinTF"<%if JoinTF="4" then response.write " checked"%>>积分要求
                       积分大小等于:<input type="text" class="textbox" name="Point" style="width:40px" maxlength="16" value="<%=Point%>" size="10">分        </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>圈子图片：</strong><br>
                  您可以上传您喜欢的图片做为圈子的图片标志。</div></td>
              <td>　<div style="margin-left:10px"><img style="border:1px solid #ccc" src="<%=photourl%>" width="100" height="100" border="1" name="showimages" align="left">
                  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9996' frameborder="0" scrolling="No" align="center" width='95%' height='30'></iframe><br>&nbsp;只支持jpg、gif、png，小于50k，默认尺寸为120*90 &nbsp;&nbsp;<br><br>
                自定义图片地址：
                  <input class="textbox" name="PhotoUrl" type="text" id="PhotoUrl" style="width:250px; " value="<%=PhotoUrl%>" />                  </div>
              </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>圈子申请说明：</strong></span></div>
                <br>
                用户申请加入该圈子的说明文字。</td>
              <td>　
                  
                  <textarea class="textbox" name="Note" id="Note" style="height:80" cols=50 rows=5><%=Note%></textarea>              </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>圈子公告：</strong></span></div>
                <br>
                群主首页的活动公告。</td>
              <td>　
                  
                  <textarea class="textbox" name="Announce" id="Announce" style="height:80" cols=50 rows=5><%=Announce%></textarea>              </td>
            </tr>
            <tr class="tdbg">
              <td height="30" align="center" colspan=2>
                <input type="submit" name="Submit3"  class="Button" value="<%=OpStr%>" />
                <input type="reset" name="Submit22"   class="Button" value=" 重 来 " />              </td>
            </tr>
          </form>
</table>
		<%
	   End Sub
	   '保存圈子
	   Sub Teamsave()
	     Dim TeamName:TeamName=KS.S("TeamName")
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Note:Note=KS.S("Note")
		 Dim JoinTF:JoinTF=KS.S("JoinTF")
		 Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
		 Dim Point:Point=KS.ChkCLng(KS.S("Point"))
		 Dim ID:ID=KS.Chkclng(KS.S("id"))
		 Dim Announce:Announce=KS.S("Announce")
		 If TeamName="" Then Response.Write "<script>alert('请输入圈子名称!');history.back();</script>"
		 If ClassID=0 Then Response.Write "<script>alert('请选择圈子类型!');history.back();</script>"
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Team Where id=" & id ,conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now
			if ks.SSetting(5)=1 then
			RS("Verific")=0
			else
			RS("Verific")=1 '设为已审
			end if
		    RS("UserName")=KSUser.UserName
		 End If
		    RS("TeamName")=TeamName
			RS("ClassID")=ClassID
			RS("Note")=Note
			RS("JoinTF")=JoinTF
			RS("Point")=Point
			RS("PhotoUrl")=PhotoUrl
			RS("Announce")=Announce
			RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=3 and IsDefault='true'")(0))
		  RS.Update
		  If KS.ChkClng(id)=0 then
		   rs.movelast
		   id=rs("id")
		  rs.close
		  rs.open "select * from ks_teamusers",conn,1,3
		  rs.addnew
		   rs("teamid")=conn.execute("select max(id) from ks_team")(0)
		   rs("username")=KSUser.UserName
		   rs("power")=2
		   rs("status")=3
		   rs("applydate")=now
		   rs("adddate")=now
		   rs("reason")="创建圈子"
		   rs.update
		     RS.Close:Set RS=Nothing
			 
			  if not KS.IsNul(PhotoUrl) Then
			  Call KS.FileAssociation(1030,id,PhotoUrl,0)
			  End If
	  		 Call KSUser.AddLog(KSUser.UserName,"创建了圈子:<a href=""{$GetSiteUrl}space/group.asp?id=" & id &""" target=""_blank"">" & TeamName & "</a>",105)
			  Response.Write "<script>alert('圈子创建成功!');location.href='User_Team.asp';</script>"

          else
		     RS.Close:Set RS=Nothing
			  if not KS.IsNul(PhotoUrl) Then
			  Call KS.FileAssociation(1030,id,PhotoUrl,1)
			  End If
	  		 Call KSUser.AddLog(KSUser.UserName,"修改了圈子:<a href=""{$GetSiteUrl}space/group.asp?id=" & id &""" target=""_blank"">" & TeamName & "</a>",105)
			  Response.Write "<script>alert('圈子修改成功!');location.href='User_Team.asp';</script>"
		  end if
		 
	   End Sub
	   '我建的圈子
	   Sub MyTeam()
	   	 
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									Dim Sql:sql = "select * from KS_Team "& Param &" order by AddDate DESC" 
								    Call KSUser.InnerLocation("我创建的圈子")
			
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
												  <td width="27%" align="center">圈子名称</td>
												  <td width="8%" align="center">创建人</td>
                                                  <td width="11%" align="center">成员数</td>
                                                  <td width="10%" align="center">主题/回复</td>
                                                  <td width="17%" align="center">创建时间</td>
                                                  <td width="20%" align="center" nowrap>管理操作</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=7 height=30 valign=top>你没有创建圈子!</td></tr>"
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
									Call ShowMyTeam
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowMyTeam
									Else
										CurrentPage = 1
										Call ShowMyTeam
									End If
								End If
				End If
     %>                     
                        </table>
		  <%
	   End Sub
	   
	   Sub ShowMyTeam()
	        Dim I
    Response.Write "<FORM Action=""User_Team.asp?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td height="22" align="center">
											<INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
											<td height="22"><a href="../space/group.asp?id=<%=rs("id")%>" target="_blank"><%=KS.GotTopic(RS("teamName"),35)%></a></td>
                                            <td height="22" align="left"><%=rs("username")%>
											</td>
                                            <td height="22" align="center"><%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%></td>
                                            <td height="22" align="center">
											 <%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>/<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%></td>
											 <td align="center"><%=formatdatetime(rs("adddate"),2)%></td>
                                            <td height="22" align="center">
											<a href="../space/group.asp?id=<%=rs("id")%>" target="_blank" class="link3">访问</a> 
											<a href="?action=VerificApply&id=<%=rs("id")%>" class="link3">审核成员[<font color=red><%=conn.execute("select count(username) from ks_teamusers where status=2 and teamid=" & rs("id"))(0)%></font>]</a>
											<a href="User_Team.asp?action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除该圈子吗?'))" class="link3">删除</a>
											</td>
                                          </tr>
                  
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=2 valign=top>
								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;全选&nbsp;<INPUT class='button' onClick="return(confirm('确定删除选中的圈子吗?'));" type=submit value=删除选定的圈子 name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;         
								  <td>
								  <td colspan="5">
								  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								  </FORM>
								</tr>
								<% 

	   End Sub
	   '我加入的圈子
	   Sub MyJoinTeam()
	   	 
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									Dim Sql:sql = "select b.username,a.id,b.teamname,b.photourl,b.adddate from ks_teamusers a, ks_team b where a.status=3 and a.teamid=b.id and a.username='" & ksuser.username & "' and b.username<>'" & ksuser.username & "' order by a.Adddate desc"
								    Call KSUser.InnerLocation("我加入的圈子")
			
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
												  <td width="27%" align="center">圈子名称</td>
												  <td width="13%" height="22" align="center">创建人</td>
                                                  <td width="11%" height="22" align="center">成员数</td>
                                                  <td width="10%" height="22" align="center">主题/回复</td>
                                                  <td width="17%" height="22" align="center">创建时间</td>
                                                  <td width="16%" height="22" align="center" nowrap>管理操作</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=7 height=30 valign=top>你没有加入任何圈子!</td></tr>"
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
									Call ShowMyJoinTeam
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowMyJoinTeam
									Else
										CurrentPage = 1
										Call ShowMyJoinTeam
									End If
								End If
				End If
     %>                     
                        </table>
		  <%
	   End Sub
	   
	   Sub ShowMyJoinTeam()
	        Dim I
    Response.Write "<FORM Action=""User_Team.asp?Action=OutTeam"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td height="22" align="center">
											<INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
											<td height="22"><a href="../space/group.asp?id=<%=rs("id")%>" target="_blank"><%=KS.GotTopic(RS("teamName"),35)%></a></td>
                                            <td height="22" align="left"><%=rs("username")%>
											</td>
                                            <td height="22" align="center"><%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%></td>
                                            <td height="22" align="center">
											 <%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>/<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%></td>
											 <td align="center"><%=formatdatetime(rs("adddate"),2)%></td>
                                            <td height="22" align="center">
											<a href="../space/group.asp?action=info&id=<%=rs("id")%>" target="_blank" class="link3">圈子信息</a> <a href="../space/group.asp?action=post&id=<%=rs("id")%>" target="_blank" class="link3">发表新帖</a>
											</td>
                                          </tr>
                  
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=2 valign=top>
								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;全选&nbsp;<INPUT class='button' onClick="return(confirm('确定删除选中的圈子吗?'));" type=submit value=删除选定的圈子 name=submit1> </td>
								<td colspan="2"><%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								  </FORM>
								</tr>
								<% 

	   End Sub
	   
	  
	   '群祖首页
	   Sub TeamIndex()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									Dim Sql:sql = "select top 5 * from KS_Team "& Param &" order by AddDate DESC"

								    Call KSUser.InnerLocation("圈子列表")
								  %>
								     
				                     <table height="500" width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="50%" height="22">我创建的圈子</td>
                                                  <td width="50%" height="22">我参与讨论的圈子</td>
                                                </tr>
												<tr class='tdbg'>
												  <td valign="top">
                                 <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "您还没有创建圈子!"
								 Else
								    Dim I,K
								   Do While Not RS.Eof
							 %>
							   <table class="border" cellSpacing=0 cellPadding=0 style="margin:3px" width="99%" border=0>
								<tr>
								 <td class="splittd" style="padding:10px" width="29%" align=center>
									   <table style="BORDER-COLLAPSE: collapse" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>
										<tr>
										 <td>
										 <a href="../space/group.asp?id=<%=rs("id")%>" title="<%=rs("teamname")%>" target="_blank">
										 <div style="width:100px;height:100px;background:transparent url(<%=rs("photourl")%>) no-repeat scroll 50% 50%;cursor:pointer;"></div>
										 </a>
										 
										 </td>
									    </tr>
								  </table>
								  </td>
									 <td class="splittd" style="padding:10px" width="71%"><a href="../space/group.asp?id=<%=rs("id")%>" title="<%=rs("teamname")%>" target="_blank"><%=rs("TeamName")%></a> (成员:<%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%>人)<br>
									   <font color="#a7a7a7">创建时间:<%=rs("adddate")%></font><br>
								     主题/回复：<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>/<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%> <br><a href="?Action=VerificApply&id=<%=rs("id")%>">审核成员(<font color=red><%=conn.execute("select count(username)  from ks_teamusers where status=2 and teamid=" & rs("id"))(0)%></font>)</a>&nbsp;<a href="?Action=EditTeam&id=<%=rs("id")%>">
								     修改</a>&nbsp;<a href="?Action=Del&id=<%=rs("id")%>" onClick="return(confirm('删除圈子将删除该圈子里的所有照片，确定删除吗？'))">删除</a>&nbsp;
                                    </td>
							   </tr>
								</table>
								  <%
										RS.MoveNext
										I=I+1
									   Loop

				                 End If
								 rs.close
     %>  
												  </td>
												  <td valign="top">
                                         <% dim sqlstr:sqlstr="select top 5 b.username,b.id,b.teamname,b.photourl,b.adddate from ks_teamusers a, ks_team b where a.status=3 and a.teamid=b.id and a.username='" & ksuser.username & "' and b.username<>'" & ksuser.username & "' order by a.Adddate desc"
										 rs.open sqlstr,conn,1,1
										 if rs.eof then
										  response.write "你没有加入任何圈子!"
										 end if
										 do while not rs.eof 
										  %>
										  <table class="border" cellSpacing=0 cellPadding=0 style="margin:3px" width="99%" border=0>
								<tr>
									 <td width="29%" align=center class="splittd" style="padding:10px">
									   <table style="BORDER-COLLAPSE: collapse" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>
										<tr>
										 <td><a href="../space/group.asp?id=<%=rs("id")%>" title="<%=rs("teamname")%>" target="_blank"><div style="width:100px;height:100px;background:transparent url(<%=rs("photourl")%>) no-repeat scroll 50% 50%;cursor:pointer;"></div></a></td>
									    </tr>
								  </table></td>
									 <td width="71%" class="splittd" style="padding:10px"><a class="teamname" href="../space/group.asp?id=<%=rs("id")%>" title="<%=rs("teamname")%>" target="_blank"><%=rs("TeamName")%></a><br>
									   <font color="#a7a7a7">群主:<%=rs("username")%></font><br>
									   <font color="#a7a7a7">创建时间:<%=rs("adddate")%></font><br>
								     主题/回复：<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>/<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%>&nbsp;&nbsp;&nbsp;成员:<%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%>人
                                    </td>
							   </tr>
								</table>
										  <%
										 rs.movenext
										 loop
										 %>
                         
												  
												  </td>
												</tr>
            </table>
		  <%
  End Sub
'审核成员
	   Sub VerificApply()
	
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
									Dim Sql:sql = "select * from KS_TeamUsers where status=2  order by AddDate DESC" 
								    Call KSUser.InnerLocation("审核申请加入圈子")
			
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
												  <td width="12%" height="22" align="center">申请人</td>
                                                  <td width="41%" height="22" align="center">加入理由</td>
                                                  <td width="10%" height="22" align="center">申请时间</td>
                                                  <td width="15%" height="22" align="center">圈子名称</td>
                                                  <td width="18%" height="22" align="center" nowrap>管理操作</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>没有用户申请加入该圈子!</td></tr>"
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
									Call ShowApply
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowApply
									Else
										CurrentPage = 1
										Call ShowApply
									End If
								End If
				End If
     %>                     
                        </table>
		  <%
	   End Sub
	   
	   Sub ShowApply()
	        Dim I
    Response.Write "<FORM Action=""User_Team.asp?Action=ApplyDel"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td height="22" align="center">
											<INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
											<td height="22" align="center"><a href="../space/?<%=conn.execute("select username from ks_user where username='" & rs("username") & "'")(0)%>" target="_blank"><%=RS("username")%></a></td>
                                            <td height="22" align="left"><%=RS("reason")%>
											
											</td>
                                            <td height="22" align="center"><%=formatdatetime(rs("applyDate"),2)%></td>
                                            <td height="22" align="center">
											<a href="../space/group.asp?id=<%=rs("teamid")%>" target="_blank"><%=conn.execute("select teamname from ks_team where id=" & rs("teamid"))(0)%></a>
											</td>
                                            <td height="22" align="center">
											<a href="User_Team.asp?id=<%=rs("id")%>&Action=AcceptApply" class="link3">接受申请</a> <a href="User_Team.asp?action=ApplyDel&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除此申请吗?'))" class="link3">拒绝</a>
											</td>
                                          </tr>
                                          <tr><td colspan=6 background='images/line.gif'></td></tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=2 valign=top>
								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;全选<INPUT class='tdbg' onClick="return(confirm('确定拒绝选中的申请吗?'));" type=submit value=拒绝选定的申请 name=submit1> </td>
								<td colspan="4">
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								  </FORM>
								</tr>
								<% 

 End Sub
 '接受请求
 Sub AcceptApply()
  Dim id:id=KS.chkclng(KS.S("id"))
  Dim RS:Set rs=server.createobject("adodb.recordset")
  rs.open "select * from ks_teamusers where id=" & id ,conn,1,3
  if not rs.eof then
     rs("status")=3
	 rs("adddate")=now
	 rs.update
    Call KS.SendInfo(rs("username"),Ksuser.username,"通过加入圈子的申请!",rs("username") & "您好!<br>您加入圈子[<a href=""../space/group.asp?id=" & rs("teamid") &""" target=""_blank"">" & conn.execute("select teamname from ks_team where id=" & rs("teamid"))(0) & "</a>]的申请已于<font color=red>" & now & "</font>通过审核，现在您可以参与该圈子的讨论!")
  end if
  rs.close:set rs=nothing
  response.redirect request.servervariables("http_referer")
 End Sub
 
 '拒绝申请
 Sub ApplyDel()
  Dim ID:id=KS.S("id")
  ID=KS.FilterIDs(ID)
  Dim rs:set rs=server.createobject("adodb.recordset")
  rs.open "select * from ks_teamusers where id in(" & id & ")",conn,1,3
  if not rs.eof then
    Call KS.SendInfo(rs("username"),Ksuser.username,"申请加入圈子被拒绝!",rs("username") & "您好!<br>您加入圈子[<a href=""../space/group.asp?id=" & rs("teamid") &""" target=""_blank"">" & conn.execute("select teamname from ks_team where id=" & rs("teamid"))(0) & "</a>]的申请已于<font color=red>" & now & "</font>被群主拒绝!")
  end if
  rs.close:set rs=nothing
  conn.execute("delete from ks_teamusers where id in(" & id & ")")
  response.redirect request.servervariables("http_referer")
 End Sub

  '删除圈子
  Sub DelTeam()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("你没有选中要删除的圈子!",ComeUrl):Response.End
  	Conn.execute("delete from ks_uploadfiles where channelid=1030 and infoid in(" & id& ")")
  	Conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where teamid in(" & id& "))")
	Conn.Execute("Delete From KS_Team Where ID In(" & ID & ")")
	Conn.execute("delete from ks_teamusers where teamid in(" & id& ")")
	Conn.execute("delete from ks_teamtopic where teamid in("&id&")")
	Response.Redirect ComeUrl
  End Sub
  Sub OutTeam()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("你没有选中要退出的圈子!",ComeUrl):Response.End
  	Conn.execute("delete from ks_teamusers where id in(" & id& ")")
	Response.Redirect ComeUrl
  End Sub
 
End Class
%> 

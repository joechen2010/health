<%@ Language="VBSCRIPT" codepage="936" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KSCls
Set KSCls = New AjaxCls
KSCls.Kesion()
Set KSCls = Nothing

Class AjaxCls
      Private KS,KSBCls
	  Private Action,UserName,UserType
	  Private CurrentPage,totalPut,MaxPerPage,PageNum
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
       Set KSBCls=New BlogCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KS=Nothing
	  Set KSBCls=Nothing
	  CloseConn()
	 End Sub

     Sub Kesion()
      Action=KS.S("Action")
      UserName=KS.S("UserName")
	  If UserName="" Then KS.Die "username null!"
	  UserName=KS.UrlDeCode(UserName)
	  UserType=KS.ChkClng(Conn.Execute("Select UserType From KS_User Where UserName='" & UserName & "'")(0))

	   
	   Select Case Action
		Case "friend" FriendList
		Case "group" GroupList
		Case "photo" PhotoList
		Case "log" LogList
		Case "guest" GuestList
		Case "xx" xxList
		Case "listxx" Listxx
	   End Select	
	 End Sub	
	 
	 Sub FriendList()
	     Response.Write KSBcls.Location("<strong>首页 >> 我的好友</strong>")
		 MaxPerPage =20
		 response.write "           <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select friend,u.userface,u.username from ks_friend f inner join ks_user u on f.friend=u.username where f.username='" & username & "' and f.accepted=1",Conn,1,1
		                 If RSObj.EOF and RSObj.Bof  Then
						  response.write "<tr><td style=""BORDER: #efefef 1px dotted;text-align:center"" colspan=3>没有加好友！</td></tr>"
						 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If

			
								If CurrentPage = 1 Then
									call showfriend(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showfriend(RSObj)
									Else
										CurrentPage = 1
										call showfriend(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|位||2"
		 RSObj.Close:Set RSObj=Nothing
		End Sub
		
		sub showfriend(RS)
		    Dim I,k
			  Do While Not RS.Eof 
                 response.write "<tr height=""20""> " &vbNewLine
				 for k=1 to 4
				 	   Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
					   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
                     response.write  "<td width=""25%"" style=""border: #efefef 1px dotted;"" align=""center""><a target=""_blank"" href=""" & KS.GetDomain & "space/?" & RS("username") & """><img width=""80"" height=""80"" src=""" & UserFaceSrc & """ border=""0""></a><div align=""center""><a target=""_blank"" href=""blog.asp?username=" & RS("username") & """ target=""_blank"">" &RS(0) & "</a></div><a href=""javascript:void(0)"" onclick=""ksblog.addF(event,'" & rs("UserName") & "');""><img src=""images/adfriend.gif"" border=""0"" align=""absmiddle"" title=""加为好友"">好友</a> <a href=""javascript:void(0)"" onclick=""ksblog.sendMsg(event,'" & rs("username") & "')""><img src=""images/sendmsg.gif"" border=""0"" align=""absmiddle"" title=""发小纸条"">消息</a></td>" & vbnewline
			     RS.MoveNext
			     I = I + 1
				 If I >= MaxPerPage or rs.eof Then Exit for
				 next 
				 do while k<4
				  response.write  "<td width=""25%"">&nbsp</td>"
				  k=K+1
				 loop
                 response.write  "</tr> " & vbcrlf
				If I >= MaxPerPage Then Exit Do
			 Loop
	end sub
			
	Sub GroupList()
	     If UserType=1 Then
	     Response.Write KSBcls.Location("<strong>首页 >> 公司圈子</strong>")
		 Else
	     Response.Write KSBcls.Location("<strong>首页 >> 我的圈子</strong>")
		 End If
		 MaxPerPage =10
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select * from KS_team where username='" & username & "' and verific=1 order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>没有创建圈子！</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showgroup(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showgroup(RSObj)
									Else
										CurrentPage = 1
										call showgroup(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		 Response.Write "<tr style=""margin:2px;border-bottom:#9999CC dotted 1px;"">"
		   Response.Write "<td width=""20%"" style=""border-bottom:#9999CC dotted 1px;"">"& vbcrlf
			Response.Write " <table style=""BORDER-COLLAPSE: collapse"" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>"
			Response.Write "	<tr>"
			Response.Write "		<td><a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img src=""" & rs("photourl") & """ width=""110"" height=""80"" border=""0""></a></td>"
			Response.Write "	 </tr>"
			Response.Write " </table>"
			Response.Write "</td>"
			Response.Write " <td style=""border-bottom:#9999CC dotted 1px;""><a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br><font color=""#a7a7a7"">创建者：" & rs("username") & "</font><br><font color=""#a7a7a7"">创建时间:" &rs("adddate") & "</font><br>主题/回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;成员:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "人  </td>"
			Response.Write "</tr>"
			Response.Write "<tr><td height='2'></td></tr>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	 Sub PhotoList()
	     If UserType=1 Then
	     Response.Write KSBcls.Location("<strong>首页 >> 公司相册</strong>")
		 Else
	     Response.Write KSBcls.Location("<strong>首页 >> 我的相册</strong>")
		 End If
		 MaxPerPage =10
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_Photoxc Where username='" & username & "' and status=1 order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>没有创建相册！</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showphoto(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showphoto(RSObj)
									Else
										CurrentPage = 1
										call showphoto(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
	 
	 Sub showphoto(rs)
	 	 Dim I,k
		 Do While Not RS.Eof 
		  Response.Write "<tr>"
		   for k=1 to 3
		   %>
		  <td width="33%" height="22" align="center">
						<table borderColor=#b2b2b2 height=149 cellSpacing=0 cellPadding=0 width="110%" border=0>
							  <tr>
								 <td align=middle width="100%"><B><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>"><%=rs("xcname")%></a></B></td>
							  </tr>
							  <tr>
									  <td align=middle width="100%">
														<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0>
														  <tr>
															<td background="images/pic.gif" width="136" height="106" valign="top"><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>" target="_blank"><img style="margin-left:6px;margin-top:5px" src="<%=rs("photourl")%>" width="120" height="90" border=0></a></td>
														  </tr>
														</table>
									  </td>
								</tr>
								<tr>
								  <td align=middle width="100%" height=20><%=rs("xps")%>张/<%=rs("hits")%>次<font color=red>[<%=GetStatusStr(rs("flag"))%>]</font></td>
						      </tr>
			  </table>
			 </td>
		   <%
			rs.movenext
			I = I + 1
			if rs.eof or i>=cint(MaxPerPage) then exit for
		   Next
		   Response.Write "</tr>"
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	 Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
	 End Function
	 
	 Sub LogList()
	     If UserType=1 Then
	     Response.Write KSBcls.Location("<strong>首页 >> 公司日志</strong>")
		 Else
	     Response.Write KSBcls.Location("<strong>首页 >> 我的日志</strong>")
		 End If
		 MaxPerPage =KSBCls.GetUserBlogParam(UserName,"ListBlogNum")
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Param:Param=" UserName='" & UserName &"'"
		 Dim KeyWord:KeyWord=KS.S("Date") '日历搜索
		 Dim Key:Key=KS.R(KS.S("Key")) '表单搜索
		 Dim Tag:Tag=KS.R(KS.S("Tag")) 
		 If IsDate(KeyWord) Then
		       If CInt(DataBaseType) = 1 Then
			   Param=Param & " And AddDate>='" & KeyWord & " 00:00:00' and AddDate<='" &KeyWord & " 23:59:59'"
			 else
			   Param=Param & " And AddDate>=#" & KeyWord & " 00:00:00# and AddDate<=#" &KeyWord & " 23:59:59#"
			 end if
		 End If
		 If ClassID<>0 Then Param=Param & " And ClassID=" & ClassID
		 If Key <>"" Then Param=Param & " And Title Like '%" & Key & "%'"
		 If Tag <>"" Then Param=Param & " And Tags Like '%" & Tag & "%'"
		 
		 If KS.S("Date")<>"" Then Response.Write "<h2>搜索日期:<font color=red>" & KS.S("Date") & "</font>的日志</h2></br>"
		 If Tag<>"" Then Response.Write "<h2>搜索Tag:<font color=red>" & Tag& "</font>的日志</h2></br>"
		 iF Key<>"" Then Response.Write "<h2>搜索标题含有""<font color=red>" & Key & "</font>""的日志</h2></br>"
		 iF ClassID<>0 Then Response.Write "<h2>搜索自定义分类ID""<font color=red>" & ClassID & "</font>""的日志</h2></br>"
		 
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogInfo Where " & Param & " and Status=0 Order By AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 
				  	  If Key<>"" Then
							response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到日志标题含有<font color=red>""" & key & """</font>的日志!</p></td></tr>"
						Else
						  If KeyWord="" And ClassID=0 Then
							response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>您还没有写日志！</p></td></tr>"
						  ElseIf ClassID<>0 Then
							response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到该分类的日志!</p></td></tr>"
						  Else
							response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>日期：<font color=red>" & KeyWord & "</font>,您没有写日志！</p></td></tr>"
						  End If
					   End if
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showlog(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showlog(RSObj)
									Else
										CurrentPage = 1
										call showlog(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	End Sub
    Sub showlog(RS)
		 Dim I
		 Do While Not RS.Eof 
		  response.write KSBCls.ReplaceLogLabel(UserName,LFCls.GetConfigFromXML("space","/labeltemplate/label","log"),RS)
		 RS.MoveNext
		 I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	End Sub
		
		Sub GuestList()
		 Response.Write KSBcls.Location("<div align=""left""><strong>首页 >> 留言板</strong>(<a href=""#write"">签写留言</a>)</div>")
		 MaxPerPage =5
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogMessage Where UserName='" & UserName & "' and status=1 Order By AddDate Desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>找不到任何留言信息！</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showguest(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showguest(RSObj)
									Else
										CurrentPage = 1
										call showguest(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
	 
		Function showguest(RS)
		 Dim I,CommentStr,n
		  CommentStr="&nbsp;&nbsp;共有 <font color=red>" & totalPut & " </font> 条留言信息，共分 <font color=red>" & pagenum & "</font> 页,第 <font color=red>" & CurrentPage & "</font> 页<br />"
			CommentStr=CommentStr & "<table  width='99%' border='0' align='center' cellpadding='0' cellspacing='1'>"
			If CurrentPage=1 Then
			 N=TotalPut
			 Else
			 N=totalPut-MaxPerPage*(CurrentPage-1)
			 End IF
		  Dim FaceStr,Publish
		  Do While Not RS.Eof 
		   FaceStr=KS.Setting(3) & "images/face/0.gif"
		
			Publish=RS("AnounName")
			If not Conn.Execute("Select UserFace From KS_User Where UserName='"& Publish & "'").eof Then
			  FaceStr=Conn.Execute("Select UserFace From KS_User Where UserName='"& Publish & "'")(0)
			  If lcase(left(FaceStr,4))<>"http" then FaceStr=KS.Setting(3) & FaceStr
		    End IF
			
		   CommentStr=CommentStr & "<tr>"
		   CommentStr=CommentStr & "<td width='70' rowspan='3' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><img width=""60"" height=""60"" src=""" & facestr & """ border=""1""></td>"
		   CommentStr=CommentStr & "<td height='25' width=""70%"">"
		   CommentStr=CommentStr & RS("Title")
		   CommentStr=CommentStr  & "  </td><td width=""30"" align=""right""><font style='font-size:32px;font-family:""Arial Black"";color:#EEF0EE'> " & N & "</font></td>"
		   CommentStr=CommentStr & "</tr>"
		   CommentStr=CommentStr & "<tr>"
		   CommentStr=CommentStr & "<td height='25' colspan='2'>" & RS("Content")
				 If Not IsNull(RS("Replay")) Or Rs("Replay")<>"" Then
				 CommentStr=CommentStr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & RS("Replay") & "<br><div align=right>时间:" & rs("replaydate") &"</div></div>"
				 End If
		   CommentStr=CommentStr & "	 </td>"
		   CommentStr=CommentStr & "</tr>"
		   CommentStr=CommentStr & "<tr>"
			 Dim MoreStr,KSUser,LoginTF
			 Set KSUser=New UserCls
			 LoginTF=Cbool(KSUser.UserLoginChecked)
			 IF LoginTF=true and KSUser.UserName=UserName Then
			  MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> | <a href='../User/user_message.asp?Action=MessageDel&id=" & RS("ID") & "' onclick=""return(confirm('确定删除该留言吗?'))"">删除</a> | <a href='../User/?User_message.asp?id=" & RS("ID") & "&Action=ReplayMessage' target='_blank'>回复</a>"
             Else
			  MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a>"
			 End If
			 Set KSUser=Nothing
		
		   CommentStr=CommentStr & "<td align='right' colspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><font color='#999999'>(" & publish & " 发表于：" & RS("AddDate") &")</font>&nbsp;&nbsp;" & MoreStr & " </td>"
		   CommentStr=CommentStr & "</tr>"
		   N=N-1
		   RS.MoveNext
				I = I + 1
			  If I >= MaxPerPage Then Exit Do
		  loop
		 CommentStr=CommentStr & "</table>"
		
		 response.write CommentStr
		End Function
			
		Sub listxx()
		Dim Channelid:Channelid=KS.ChkClng(KS.S("ChannelID"))
		if channelid=0 then channelid=1
		Dim SQL,K,OPStr,RSC:Set RSC=Conn.Execute("Select ChannelID,itemName From KS_Channel Where ChannelStatus=1 and channelid<>6 And ChannelID<>9 order by channelid")
		SQL=RSc.GetRows(-1)
		RSc.Close:set RSc=Nothing
		For K=0 To Ubound(SQL,2)
		 if channelid=sql(0,k) then
		 OpStr=OpStr & "<option value='" & SQL(0,K) & "' selected>" & SQL(1,K) & "</option>"
		 else
		 OpStr=OpStr & "<option value='" & SQL(0,K) & "'>" & SQL(1,K) & "</option>"
		 end if
		Next
		
	    Response.Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td></td><td align=right>信息分类<select name='channelid' onchange=""ksblog.ajaxLoadPage(ksblog._url,'action=listxx&username=" & username & "&channelid='+this.value,'post','ksblog._setxx');"">" & opstr & "</select>&nbsp;&nbsp;&nbsp;</td></tr></table>"
		
		 Dim Sqlstr,Max
		 if KS.C_S(ChannelID,6) =2 then
		 Max=12
		 else
		 Max=10
		 end if
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 2
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,PhotoUrl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
          Case 3
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 4
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 5
  		   SQLStr="Select top " & Max & " ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 7
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 8
  		   SQLStr="Select top " & Max & " ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where UserName='" & UserName & "' Order By AddDate Desc,id desc"
		 End Select


		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SqlStr,Conn,1,1
		         If RS.EOF and RS.Bof  Then
					response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到您要的信息！</p></td></tr>"
				 Else
				  if KS.C_S(ChannelID,6) =2 then
				   Response.Write GetUserPhoto(RS,Max,ChannelID)
				  else
					  do while not rs.eof 
						   Response.write "<tr><td style=""border-bottom: #efefef 1px dotted;height:22""><img src=""../images/arrow_r.gif"" align=""absmiddle""> [" & KS.GetClassNP(RS(2)) & "] <a href='" & KS.GetItemUrl(channelid,RS(2),RS(0),RS(5)) & "' target='_blank'>" & RS(1) & "</a>&nbsp;&nbsp;(" & RS(7) & ")</td></tr>"
						  rs.movenext
						  loop
					  end if
				 End If	
		 response.write "  </table>" & vbcrlf
		 rs.close:set rs=nothing
		End Sub
		
			'===========9-30========================
			Function GetUserPhoto(RS,totalPut,ChannelID)
		    Dim I,K,Url
			Dim PerLineNum:PerLineNum=4   '每行显示作品数
			  Do While Not RS.Eof 
              GetUserPhoto=GetUserPhoto & "<tr height=""20""> " &vbNewLine
			  
			  For K=1 To PerLineNum
			  If ChannelID=2 Then
			   Url="../space/?" & UserName & "/showphoto/" & RS(0)
			  Else
			   Url=KS.GetItemUrl(channelid,RS(2),RS(0),RS(5))
			  End If
              GetUserPhoto=GetUserPhoto & "  <td style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center""><a href=""" & Url & """ target=""_blank""><img style='border:1px #efefef solid' width=120 height=80 src=""" & rs("photourl") & """ border=""0""></a><br><a href=""" & Url & """ target=""_blank"">" & KS.Gottopic(RS(1),15) & "</a></td>" & vbnewline
             RS.MoveNext
			    I = I + 1
				If rs.eof or I >= totalPut Then Exit For
			  Next
			   For K=K+1 To PerLineNum
            GetUserPhoto=GetUserPhoto & "   <td width=120 style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted;text-align:center"">&nbsp;</td> " & vbcrlf
			   Next
            GetUserPhoto=GetUserPhoto & "   </tr> " & vbcrlf
				If I >= totalPut Then Exit Do
			 Loop

			End Function

		
			
		Sub xxlist()
		Dim Channelid:Channelid=KS.ChkClng(KS.S("ChannelID"))
		if channelid=0 then channelid=1
		Dim SQL,K,OPStr,RSC:Set RSC=Conn.Execute("Select ChannelID,itemName From KS_Channel Where ChannelStatus=1 and channelid<>6  And ChannelID<>9 And ChannelID<>10 order by channelid")
		SQL=RSc.GetRows(-1)
		RSc.Close:set RSc=Nothing
		For K=0 To Ubound(SQL,2)
		 if channelid=sql(0,k) then
		 OpStr=OpStr & "<option value='" & SQL(0,K) & "' selected>" & SQL(1,K) & "</option>"
		 else
		 OpStr=OpStr & "<option value='" & SQL(0,K) & "'>" & SQL(1,K) & "</option>"
		 end if
		Next
	    Response.Write KSBcls.Location("<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td><strong>首页 >> 信息集</strong></td><td align=right>信息分类<select name='channelid' onchange=""ksblog.loadxx(this.value,'" & UserName & "')"">" & opstr & "</select>&nbsp;&nbsp;&nbsp;</td></tr></table>")
		 MaxPerPage =20
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim Sqlstr
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 2
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,photourl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
          Case 3
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 4
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 5
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 7
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 8
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where UserName='" & UserName & "' Order By AddDate Desc"
		 End Select

		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open SqlStr,Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
							response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3><p>找不到您要的信息！</p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showxx(RSObj,channelid)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showxx(RSObj,channelid)
									Else
										CurrentPage = 1
										call showxx(RSObj,channelid)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	End Sub	
	
	Sub showxx(rs,channelid)
		if KS.C_S(ChannelID,6) =2 then       '图片显示不同
		   Response.Write GetUserPhoto(RS,MaxPerPage,ChannelID)
		Else
			 Dim K,SQL
			 Do While Not RS.Eof
				Response.write "<tr><td style=""border-bottom: #efefef 1px dotted;height:22""><img src=""../images/arrow_r.gif"" align=""absmiddle""> [" & KS.GetClassNP(RS(2)) & "] <a href='" & KS.GetItemUrl(channelid,RS(2),RS(0),RS(5)) & "' target='_blank'>" & RS(1) & "</a>&nbsp;&nbsp;(" & RS(7) & ")</td></tr>"
				K=K+1
				If K>=MaxPerPage Then Exit Do
				RS.MoveNext
			 Loop
		 End if
	End Sub
					 
	
	
 End Class 
%>
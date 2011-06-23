<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
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
Set KSCls = New UserAjax
KSCls.Kesion()
Set KSCls = Nothing

Class UserAjax
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		  
		  Select Case KS.S("Action")
		   Case "GetNewMessage" Call GetNewMessage()
		   Case "GetAdminMessage" Call GetAdminMessage()
		   Case "Info" Call GetUserBasicInfo()
		   Case "ModelMenu" Call GetModelMenu()
		   Case "SpaceMenu" Call SpaceMenu()
		   Case "Space" Call TurnToSpace()
		   Case "JobMenu" Call GetJobMenu()
		   Case "TopTips" Call TopTips()
		  End Select
		End Sub
		
		Sub TurnToSpace()
		  IF Cbool(KSUser.UserLoginChecked)=false Then
		   Response.Write "<script>alert('请先登录!');window.close();</script>"
		  Exit Sub
		  End If
		  If KS.SSetting(14)="1" and not conn.execute("select top 1 username from ks_blog where username='" & ksuser.username & "'").eof Then
		   dim predomain:predomain=conn.execute("select top 1 [domain] from ks_blog where username='" & ksuser.username & "'")(0)
		  end if
		  if predomain<>"" then
		    Response.Redirect("http://" & predomain &"." & KS.SSetting(16))
		   else
		    If KS.SSetting(21)="1" Then
		     Response.Redirect "../space/" & KSUser.UserName
			Else
		     Response.Redirect "../space/?" & KSUser.UserName
			End If
		   end if		
		End Sub
		
		Sub GetUserBasicInfo()
		%>
		<div class="mem_left_top">
			<div class="mem_left_photo">
			<%
			IF Cbool(KSUser.UserLoginChecked)=false Then
			   KS.Echo "<img src=""../images/face/0.gif"" width=""55"" height=""53"" alt=""个人形象"" />"
			Else
			  Dim UserFaceSrc:UserFaceSrc=KSUser.UserFace
			  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
			  KS.Echo escape("<img src=""" & UserFaceSrc & """ alt=""" & KSUser.RealName & """ border=""1"" width=""55"" height=""53"">")
			End If
			%>
			</div>
			<div class="mem_left_name">
			 <ul>
				<%
				 KS.Echo escape("<li><strong>帐号：</strong>" & KSUser.UserName & "</li><li><strong>组别：</strong><span style=""curson:pointer"" title=""" & KS.U_G(KSUser.GroupID,"groupname") & """>" & KS.Gottopic(KS.U_G(KSUser.GroupID,"groupname"),12) & "</span></li>")
				 %>
				 <li><a href="User_EditInfo.asp?Action=face" target="main">更新我的头像</a></li>
			 </ul>
		  </div>
		</div>
		<%
		End Sub
		
		Sub TopTips()
		    KSUser.UserLoginChecked
		    Dim Str
			If (Hour(Now) < 6) Then
            Str= "凌晨好,"
			ElseIf (Hour(Now) < 9) Then
			Str= "早上好,"
			ElseIf (Hour(Now) < 12) Then
			Str= "上午好,"
			ElseIf (Hour(Now) < 14) Then
			Str= "中午好,"
			ElseIf (Hour(Now) < 17) Then
			Str= "下午好,"
			ElseIf (Hour(Now) < 18) Then
			Str= "傍晚好,"
			Else
			Str= "晚上好,"
			End If
			KS.Echo escape("<strong>" & str & KSUser.UserName & "</strong>")
		 %>
<%
							  dim spacedomain,predomain
							  If KS.SSetting(14)="1" and not conn.execute("select top 1 username from ks_blog where username='" & ksuser.username & "'").eof Then
							   predomain=conn.execute("select top 1 [domain] from ks_blog where username='" & ksuser.username & "'")(0)
							  end if
							  if predomain<>"" then
							   spacedomain="http://" & predomain & "." & KS.SSetting(16)
							  else
							   If KS.SSetting(21)="1" Then
								 spacedomain=KS.GetDomain & "space/" & KSUser.UserName
								Else
								 spacedomain=KS.GetDomain & "space/?" & KSUser.UserName
								End If
							  end if
		   KS.Echo escape("&nbsp;&nbsp;您的空间地址:<a href='" & spacedomain & "' target='_blank'>" & spacedomain & "</a>")
	    End Sub
		
		Sub GetNewMessage()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "站内消息(0)"
		  Exit Sub
		End If
		Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)

		Response.write "站内消息(<font color='#ff0000'>" & MyMailTotal&"</font>)"
		If MyMailTotal>0 Then Response.Write "<bgsound src=""images/mail.wmv"" border=0>"
		End Sub

		Sub GetAdminMessage()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "(0)"
		  Exit Sub
		End If
		Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
		Response.write "(<font color='#ff0000'>" & MyMailTotal&"</font>)"
		If MyMailTotal>0 Then Response.Write "<bgsound src=""../User/images/mail.wmv"" border=0>"
		End Sub
		
		Sub GetJobMenu()
		 KSUser.UserLoginChecked
		 If KSUser.UserType=1 Then
		  Response.Write "<li><a href='user_jobcompany.asp' target='main'>招聘单位资料</a></li>"
		  Response.Write "<li><a href='User_JobCompanyZW.asp' target='main'>职位管理</a> <a href='User_JobCompanyZW.asp?action=Add' target='main'>发布</a></li>"
		  Response.Write "<li><a href='User_JobCompanyAccept.asp' target='main'>已收到的应聘简历</a></li>"
		  Response.Write "<li><a href='User_JobInterview.asp' target='main'>发出的面试函管理</a></li>"
		  Response.Write "<li><a href='User_JobResumeSC.asp' target='main'>简历收藏夹</a></li>"
		  Response.Write "<li><a href='User_JobResumeSearch.asp' target='main'>在线搜索简历</a></li>"
		 Else
		  Response.Write "<li><a href='user_jobresume.asp' target='main'>发布/修改求职简历</a></li>"
		  Response.Write "<li><a href='user_jobresumeedu.asp' target='main'>设置教育背景</a></li>"
		  Response.Write "<li><a href='user_jobresumetemp.asp' target='main'>设置简历模板</a></li>"
		  Response.Write "<li><a href='user_jobresumelist.asp' target='main'>简历投递记录</a></li>"
		  Response.Write "<li><a href='user_jobinter.asp' target='main'>查看面试通知</a></li>"
		  Response.Write "<li><a href='user_jobresumeletter.asp' target='main'>我的求职信</a></li>"
		 End If
		End Sub
		
		Sub GetModelMenu()
		    KSUser.UserLoginChecked
		    if KSUser.CheckPower("s18")=false  then KS.Die "<div>---</div>"
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "Select ChannelID,BasicType,ItemName,ChannelName,ChannelTable From KS_Channel Where ChannelStatus=1 and ChannelID<>6 and channelid<>5 and usertf>0",Conn,1,1
			If RS.Eof Then 
			 RS.Close:Set RS=Nothing
			Else
			Dim K,SQL:SQL=RS.GetRows(-1)
			RS.Close:Set RS=Nothing
			For K=0 To Ubound(SQL,2)
			 KS.Echo "<li class=""fs"">"
			 select Case SQL(1,K)
			  Case 1
			    KS.echo "<a href='User_MyArticle.asp?ChannelID=" & SQL(0,K) & "' onFocus=""this.blur()"" target='main'>" & SQL(3,K) & "管理(" & conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0) & ")</a>"
			Case 2%>
			  <!--<img src="../images/user/log/<%=SQL(1,K)%>.gif" align="absmiddle" />--><a onFocus="this.blur()" href="User_MyPhoto.asp?ChannelID=<%=SQL(0,K)%>" target="main"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 3%>
			  <a href="User_MySoftWare.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 4%>
			<a href="User_MyFlash.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 5%>
			 <a href="User_MyShop.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 7%>
			  <a href="User_MyMovie.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 8%>
			 <a href="User_MySupply.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
		    <%Case 9%>
			 <a href="User_MyExam.asp" target="main" onFocus="this.blur()"><%=SQL(3,K)%>管理</a>
			<%
			 End Select
			 KS.Echo "</li>"
			Next
		  End If
			 KS.Echo "<hr size=""1"" style=""width:90%;color:green;margin:4px""/>"
			If KS.C_S(5,21)=1 Then
						 KS.Echo "<li class=""fss""><a href=""user_order.asp"" target=""main"">已买到的商品</A> <a href=""user_order.asp?action=coupon"" style=""color:red"" target=""main"">优惠券</a></li>"
			End If 
			If KS.C_S(10,21)=1 Then
			If KSUser.UserType=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						 KS.Echo "<li class=""fss""><a href=""User_JobResume.asp"" target=""main"">求职中心</A> <a href=""User_JobResume.asp"" target=""main"">简历</a></li>"
					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
				   KS.Echo "<li class=""fss""><a href=""user_Enterprise.asp?action=job"" target=""main"">企业招聘管理</a> <a href=""User_JobCompanyZW.asp?Action=Add"" target=""main"">发布</a></li>"
			 end if
            End If
		   End If
		   If KSUser.CheckPower("s09")=true Then %>
		   <li class="fss">
			<a href="User_Askquestion.asp" target="main" class=f_size title="" onFocus="this.blur()" style="text-align:left;">我的提问/回答</a> <a href="../ask/a.asp" target="_blank">提问</a>
		   </li>
		   <%end if%>
		   <li class="fss">
			<a href="User_ItemSign.asp" target="main" class=f_size>待我签收的文档</a>
		   </li>
		   <%if KSUser.CheckPower("s16")=true then%>
		   <li class="fss">
			<a href="User_favorite.asp" target="main" class=f_size>我的收藏夹</a>
		   </li>
		   <%end if%>
		   <li class="fss">
		   <a href="user_feedback.asp" target="main">我要投诉或提建议</a>			
		   </li>
			<%
		End Sub

		
		Sub SpaceMenu()
		  KSUser.UserLoginChecked
		  If KS.SSetting(0)=0 Then
		   KS.Echo "<dl>没有开通此功能</dl>"
		  Else
		    Dim Str:Str=""
		   If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then
		  
						If KSUser.CheckPower("s01")=true then
						   str= "<dl><a href=""User_Blog.asp?Action=BlogEdit"" target=""main"">空间设置</a> <span><a href=""user_Enterprise.asp"" target=""main"" title=""升级为企业空间"">升级</a></span></dl>"
						End If
						If KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
						   str=str & "<dl><a href=""User_Blog.asp"" target=""main"">我的日志</a> <span><a href=""User_Blog.asp?Action=Add"" target=""main"">写日志</a></span></dl>"
						End If
						If KSUser.CheckPower("s03")=true Then
						   str=str & "<dl><a href=""User_Friend.asp"" target=""main"">我的好友</A> <span><a href=""User_Friend.asp?action=addF"" target=""main"">寻找</a></span></dl>"
						End If
						If KSUser.CheckPower("s04")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Music.asp"" target=""main"">我的音乐</A> <span><a href=""User_Music.asp?Action=addlink"" target=""main"">添加</a></span></dl>"
						End If
						If KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Photo.asp"" target=""main"">我的相册</A> <span><a href=""User_Photo.asp?Action=Add"" target=""main"">上传</a></span></dl>"
						End If
						If KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Team.asp"" target=""main"">我的圈子</A> <span><a href=""User_Team.asp?action=CreateTeam"" target=""main"">创建</a></span></dl>"
						End If
						If KSUser.CheckPower("s10")=true Then 
						 If KS.C_S(5,21)="1" Then
						  str=str & "<dl><a href=""User_MyShop.asp?ChannelID=5"" target=""main"">我的宝贝</a> <span><a href=""User_MyShop.asp?action=Add&ChannelID=5"" target=""main"">出售</A></span></dl>"
						 end if
					    End If
						
						

					   If KSUser.CheckPower("s07")=true Then 
						str=str & "<dl><a href=""User_Class.asp"" target=""main"">我的专栏</A> <span><a href=""User_Class.asp?action=Add"" target=""main"">创建</a></span></dl>"
                       End If
			else
			     if KSUser.CheckPower("s01")=true then
			      str=str & "<dl><a href=""User_Blog.asp?Action=BlogEdit"" target=""main"">企业空间设置</A></dl>"
				 End If
				 
				  str=str & "<dl><a href=""user_Enterprise.asp"" target=""main"">企业基本信息</a></dl>"
				 
				 if KSUser.CheckPower("s10")=true then
				   str=str & "<dl><a href=""User_MyShop.asp?ChannelID=5"" target=""main"">企业产品管理</a> <span><a href=""User_MyShop.asp?action=Add&ChannelID=5"" target=""main"">发布</A></span></dl>"
				   end if
				   if KSUser.CheckPower("s11")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_EnterpriseNews.asp"" target=""main"">企业新闻管理</a></dl>"
				   end if
				  if KSUser.CheckPower("s12")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_EnterpriseAD.asp"" target=""main"">关键词广告管理</a></dl>"
				  end if
				 
				  if KSUser.CheckPower("s13")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_Enterprisezs.asp"" target=""main"">企业荣誉证书</a></dl>"
				  end if
				  if KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Blog.asp"" target=""main"">企业日志管理</A></dl>"
				  end if
				  If KSUser.CheckPower("s03")=true then
				   str=str & "<dl><a href=""User_Friend.asp"" target=""main"">我的商友</A> <span><a href=""User_Friend.asp?action=addF"" target=""main"">寻找</a></span></dl>"
				  End If
				  if KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Photo.asp"" target=""main"">企业相册管理</A></dl>"
				  end if
				  if KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Team.asp"" target=""main"">企业商圈管理</A></dl>"
				  end if

				  if KSUser.CheckPower("s07")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Class.asp"" target=""main"">专栏分类管理</A></dl>"
				  end if
			 End If
			 if KS.IsNul(str) Then str="<dl>对不起,您没有此功能功权限!</dl>"
			 KS.Echo str
		 End If
		End Sub
End Class
%> 

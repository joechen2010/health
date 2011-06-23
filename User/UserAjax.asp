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
		   Response.Write "<script>alert('���ȵ�¼!');window.close();</script>"
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
			   KS.Echo "<img src=""../images/face/0.gif"" width=""55"" height=""53"" alt=""��������"" />"
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
				 KS.Echo escape("<li><strong>�ʺţ�</strong>" & KSUser.UserName & "</li><li><strong>���</strong><span style=""curson:pointer"" title=""" & KS.U_G(KSUser.GroupID,"groupname") & """>" & KS.Gottopic(KS.U_G(KSUser.GroupID,"groupname"),12) & "</span></li>")
				 %>
				 <li><a href="User_EditInfo.asp?Action=face" target="main">�����ҵ�ͷ��</a></li>
			 </ul>
		  </div>
		</div>
		<%
		End Sub
		
		Sub TopTips()
		    KSUser.UserLoginChecked
		    Dim Str
			If (Hour(Now) < 6) Then
            Str= "�賿��,"
			ElseIf (Hour(Now) < 9) Then
			Str= "���Ϻ�,"
			ElseIf (Hour(Now) < 12) Then
			Str= "�����,"
			ElseIf (Hour(Now) < 14) Then
			Str= "�����,"
			ElseIf (Hour(Now) < 17) Then
			Str= "�����,"
			ElseIf (Hour(Now) < 18) Then
			Str= "�����,"
			Else
			Str= "���Ϻ�,"
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
		   KS.Echo escape("&nbsp;&nbsp;���Ŀռ��ַ:<a href='" & spacedomain & "' target='_blank'>" & spacedomain & "</a>")
	    End Sub
		
		Sub GetNewMessage()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "վ����Ϣ(0)"
		  Exit Sub
		End If
		Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
		MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)

		Response.write "վ����Ϣ(<font color='#ff0000'>" & MyMailTotal&"</font>)"
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
		  Response.Write "<li><a href='user_jobcompany.asp' target='main'>��Ƹ��λ����</a></li>"
		  Response.Write "<li><a href='User_JobCompanyZW.asp' target='main'>ְλ����</a> <a href='User_JobCompanyZW.asp?action=Add' target='main'>����</a></li>"
		  Response.Write "<li><a href='User_JobCompanyAccept.asp' target='main'>���յ���ӦƸ����</a></li>"
		  Response.Write "<li><a href='User_JobInterview.asp' target='main'>���������Ժ�����</a></li>"
		  Response.Write "<li><a href='User_JobResumeSC.asp' target='main'>�����ղؼ�</a></li>"
		  Response.Write "<li><a href='User_JobResumeSearch.asp' target='main'>������������</a></li>"
		 Else
		  Response.Write "<li><a href='user_jobresume.asp' target='main'>����/�޸���ְ����</a></li>"
		  Response.Write "<li><a href='user_jobresumeedu.asp' target='main'>���ý�������</a></li>"
		  Response.Write "<li><a href='user_jobresumetemp.asp' target='main'>���ü���ģ��</a></li>"
		  Response.Write "<li><a href='user_jobresumelist.asp' target='main'>����Ͷ�ݼ�¼</a></li>"
		  Response.Write "<li><a href='user_jobinter.asp' target='main'>�鿴����֪ͨ</a></li>"
		  Response.Write "<li><a href='user_jobresumeletter.asp' target='main'>�ҵ���ְ��</a></li>"
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
			    KS.echo "<a href='User_MyArticle.asp?ChannelID=" & SQL(0,K) & "' onFocus=""this.blur()"" target='main'>" & SQL(3,K) & "����(" & conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0) & ")</a>"
			Case 2%>
			  <!--<img src="../images/user/log/<%=SQL(1,K)%>.gif" align="absmiddle" />--><a onFocus="this.blur()" href="User_MyPhoto.asp?ChannelID=<%=SQL(0,K)%>" target="main"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 3%>
			  <a href="User_MySoftWare.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 4%>
			<a href="User_MyFlash.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 5%>
			 <a href="User_MyShop.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 7%>
			  <a href="User_MyMovie.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
			<%Case 8%>
			 <a href="User_MySupply.asp?ChannelID=<%=SQL(0,K)%>" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����(<%=conn.execute("select count(1) from " & SQL(4,K) &" Where Inputer='" & KSUser.UserName & "'")(0)%>)</a>
		    <%Case 9%>
			 <a href="User_MyExam.asp" target="main" onFocus="this.blur()"><%=SQL(3,K)%>����</a>
			<%
			 End Select
			 KS.Echo "</li>"
			Next
		  End If
			 KS.Echo "<hr size=""1"" style=""width:90%;color:green;margin:4px""/>"
			If KS.C_S(5,21)=1 Then
						 KS.Echo "<li class=""fss""><a href=""user_order.asp"" target=""main"">���򵽵���Ʒ</A> <a href=""user_order.asp?action=coupon"" style=""color:red"" target=""main"">�Ż�ȯ</a></li>"
			End If 
			If KS.C_S(10,21)=1 Then
			If KSUser.UserType=0 Then
				If KSUser.CheckPower("s14")=true Then 
					 If KS.C_S(10,21)="1" Then
						 KS.Echo "<li class=""fss""><a href=""User_JobResume.asp"" target=""main"">��ְ����</A> <a href=""User_JobResume.asp"" target=""main"">����</a></li>"
					 End If
				End If
			Else			 
			if KSUser.CheckPower("s14")=true  then
				   KS.Echo "<li class=""fss""><a href=""user_Enterprise.asp?action=job"" target=""main"">��ҵ��Ƹ����</a> <a href=""User_JobCompanyZW.asp?Action=Add"" target=""main"">����</a></li>"
			 end if
            End If
		   End If
		   If KSUser.CheckPower("s09")=true Then %>
		   <li class="fss">
			<a href="User_Askquestion.asp" target="main" class=f_size title="" onFocus="this.blur()" style="text-align:left;">�ҵ�����/�ش�</a> <a href="../ask/a.asp" target="_blank">����</a>
		   </li>
		   <%end if%>
		   <li class="fss">
			<a href="User_ItemSign.asp" target="main" class=f_size>����ǩ�յ��ĵ�</a>
		   </li>
		   <%if KSUser.CheckPower("s16")=true then%>
		   <li class="fss">
			<a href="User_favorite.asp" target="main" class=f_size>�ҵ��ղؼ�</a>
		   </li>
		   <%end if%>
		   <li class="fss">
		   <a href="user_feedback.asp" target="main">��ҪͶ�߻��Ὠ��</a>			
		   </li>
			<%
		End Sub

		
		Sub SpaceMenu()
		  KSUser.UserLoginChecked
		  If KS.SSetting(0)=0 Then
		   KS.Echo "<dl>û�п�ͨ�˹���</dl>"
		  Else
		    Dim Str:Str=""
		   If Conn.Execute("Select top 1 ID From ks_EnterPrise Where UserName='" &KSUser.UserName & "'").eof Then
		  
						If KSUser.CheckPower("s01")=true then
						   str= "<dl><a href=""User_Blog.asp?Action=BlogEdit"" target=""main"">�ռ�����</a> <span><a href=""user_Enterprise.asp"" target=""main"" title=""����Ϊ��ҵ�ռ�"">����</a></span></dl>"
						End If
						If KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
						   str=str & "<dl><a href=""User_Blog.asp"" target=""main"">�ҵ���־</a> <span><a href=""User_Blog.asp?Action=Add"" target=""main"">д��־</a></span></dl>"
						End If
						If KSUser.CheckPower("s03")=true Then
						   str=str & "<dl><a href=""User_Friend.asp"" target=""main"">�ҵĺ���</A> <span><a href=""User_Friend.asp?action=addF"" target=""main"">Ѱ��</a></span></dl>"
						End If
						If KSUser.CheckPower("s04")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Music.asp"" target=""main"">�ҵ�����</A> <span><a href=""User_Music.asp?Action=addlink"" target=""main"">���</a></span></dl>"
						End If
						If KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Photo.asp"" target=""main"">�ҵ����</A> <span><a href=""User_Photo.asp?Action=Add"" target=""main"">�ϴ�</a></span></dl>"
						End If
						If KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true Then
						   str=str & "<dl><a href=""User_Team.asp"" target=""main"">�ҵ�Ȧ��</A> <span><a href=""User_Team.asp?action=CreateTeam"" target=""main"">����</a></span></dl>"
						End If
						If KSUser.CheckPower("s10")=true Then 
						 If KS.C_S(5,21)="1" Then
						  str=str & "<dl><a href=""User_MyShop.asp?ChannelID=5"" target=""main"">�ҵı���</a> <span><a href=""User_MyShop.asp?action=Add&ChannelID=5"" target=""main"">����</A></span></dl>"
						 end if
					    End If
						
						

					   If KSUser.CheckPower("s07")=true Then 
						str=str & "<dl><a href=""User_Class.asp"" target=""main"">�ҵ�ר��</A> <span><a href=""User_Class.asp?action=Add"" target=""main"">����</a></span></dl>"
                       End If
			else
			     if KSUser.CheckPower("s01")=true then
			      str=str & "<dl><a href=""User_Blog.asp?Action=BlogEdit"" target=""main"">��ҵ�ռ�����</A></dl>"
				 End If
				 
				  str=str & "<dl><a href=""user_Enterprise.asp"" target=""main"">��ҵ������Ϣ</a></dl>"
				 
				 if KSUser.CheckPower("s10")=true then
				   str=str & "<dl><a href=""User_MyShop.asp?ChannelID=5"" target=""main"">��ҵ��Ʒ����</a> <span><a href=""User_MyShop.asp?action=Add&ChannelID=5"" target=""main"">����</A></span></dl>"
				   end if
				   if KSUser.CheckPower("s11")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_EnterpriseNews.asp"" target=""main"">��ҵ���Ź���</a></dl>"
				   end if
				  if KSUser.CheckPower("s12")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_EnterpriseAD.asp"" target=""main"">�ؼ��ʹ�����</a></dl>"
				  end if
				 
				  if KSUser.CheckPower("s13")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""user_Enterprisezs.asp"" target=""main"">��ҵ����֤��</a></dl>"
				  end if
				  if KSUser.CheckPower("s02")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Blog.asp"" target=""main"">��ҵ��־����</A></dl>"
				  end if
				  If KSUser.CheckPower("s03")=true then
				   str=str & "<dl><a href=""User_Friend.asp"" target=""main"">�ҵ�����</A> <span><a href=""User_Friend.asp?action=addF"" target=""main"">Ѱ��</a></span></dl>"
				  End If
				  if KSUser.CheckPower("s05")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Photo.asp"" target=""main"">��ҵ������</A></dl>"
				  end if
				  if KSUser.CheckPower("s06")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Team.asp"" target=""main"">��ҵ��Ȧ����</A></dl>"
				  end if

				  if KSUser.CheckPower("s07")=true And KSUser.CheckPower("s01")=true then
				   str=str & "<dl><a href=""User_Class.asp"" target=""main"">ר���������</A></dl>"
				  end if
			 End If
			 if KS.IsNul(str) Then str="<dl>�Բ���,��û�д˹��ܹ�Ȩ��!</dl>"
			 KS.Echo str
		 End If
		End Sub
End Class
%> 

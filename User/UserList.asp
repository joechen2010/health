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
Set KSCls = New UserList
KSCls.Kesion()
Set KSCls = Nothing

Class UserList
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 CloseConn
		End Sub
		Public Sub Kesion()
         KSUser.Head()
		 Call KSUser.InnerLocation("所有注册会员")	
		 %>
		 <div class="tabs">	
			<ul>
				<li class='select'>所有注册会员</li>
			</ul>
			<div style="padding-top:8px">
			·<a href="?ListType=1">按会员ID排序</a> ·<a href="?ListType=2">按注册日期排序</a> ·<a href="?ListType=3">按登录次数排序</a>
			</div>
		</div>
	
      <%
		 Response.Write GetUserList	   
	   End Sub
  
  Function GetUserList()
  		Dim  CurrentPage,totalPut,RS,MaxPerPage,SqlStr,ListType,Param
		ListType=KS.ChkClng(KS.S("ListType"))
		MaxPerPage =15
		If KS.S("page") <> "" Then
					CurrentPage = KS.ChkClng(KS.S("page"))
		 Else
					CurrentPage = 1
		 End If
         GetUserList= " <table cellspacing=""1"" class=""border"" cellpadding=""1"" width=""98%"" align=""center"" border=""0"">" & vbnewline
         GetUserList= GetUserList &"     <tr class=""title"" height=""20"">" & vbnewline
         GetUserList= GetUserList &"       <td width=""9%"" align=""center"">用户名</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""14%"" align=""center"">会员组</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""7%"" align=""center"">性别</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""17%"" align=""center"" nowrap=""nowrap"">邮箱</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""24%"" align=""center"" nowrap=""nowrap"">主页</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""20%"" align=""center"" nowrap=""nowrap"">最后登录</td>" & vbnewline
         GetUserList= GetUserList &"       <td width=""11%"" align=""center"" nowrap=""nowrap"">登录数</td>" & vbnewline
         GetUserList= GetUserList &"     </tr>" & vbnewline
			  
			  Set RS=Server.CreateObject("Adodb.Recordset")
			  
			  If ListType=1 Then
			   Param="Order By UserID Desc"
			  ElseIF ListType=2 Then
			   Param="Order By LastLoginTime Desc"
			  ElseIF ListType=3 Then
			   Param="Order By LoginTimes Desc"
			  End IF
			  if KS.S("Username")<>"" then
			  SqlStr="Select * From KS_User where groupid<>4 and username like '%" & ks.s("username") & "%' " & Param
			  else
			  SqlStr="Select * From KS_User where groupid<>4 " & Param
			  end if
			  RS.Open SqlStr,Conn,1,1
			       If Not RS.EOF  Then
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
									GetUserList= GetUserList & showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										GetUserList= GetUserList &showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
									Else
										CurrentPage = 1
										GetUserList= GetUserList &showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
									End If
								End If
				           End If
        GetUserList= GetUserList &"  </table>" & vbnewline
		GetUserList= GetUserList & "<form action='userlist.asp' name='myform' method='pose'>" & vbcrlf
		GetUserList= GetUserList & "&nbsp;&nbsp;&nbsp;快速查找用户->&nbsp;用户名:<input type=""text"" name=""username"" size=""20"" maxlength=""30"">" & vbcrlf
		GetUserList= GetUserList & "<input type='submit' value='搜索'>" & vbcrlf
		GetUserList= GetUserList & "</form>" & vbcrlf
		  RS.Close:Set RS=Nothing
		  End Function
		  
		Function ShowContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
		    Dim I,Privacy
			  Do While Not RS.Eof 
			   Privacy=RS("Privacy")
              ShowContent = ShowContent & "<tr class='tdbg'  onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> " &vbNewLine
              ShowContent = ShowContent & "  <td width=""9%""><img src=""images/m_list_49.gif"" align=""absmiddle""><a href=""../space/space.asp?username=" & RS("Username") & """ target=""_blank"">" & RS("UserName") & "</a></td>" & vbnewline
              ShowContent = ShowContent & "  <td style=""text-align:center"">"& KS.GetUserGroupName(RS("GroupID")) & "</td>" & vbcrlf
             ShowContent = ShowContent & "   <td style=""text-align:center"">" & vbnewline
			 If Privacy=2 Then ShowContent = ShowContent & "保密"  Else ShowContent = ShowContent & RS("Sex")  &"</td>" & vbcrlf
             ShowContent = ShowContent & "   <td>" & vbcrlf
			 If Privacy=2 Then ShowContent = ShowContent & "保密"  Else ShowContent = ShowContent & "<a href=""mailto:" & RS("Email") &""">" & RS("Email") & "</a>" & "</td>" & vbcrlf
             ShowContent = ShowContent & "   <td style=""text-align:center"">" & vbcrlf
			 If Privacy=2 Then ShowContent = ShowContent & "保密"  Else ShowContent = ShowContent & "<a href=""" & RS("HomePage") & """ target=""_blank"">" & RS("HomePage") & "</a>" & "</td>" & vbcrlf
             ShowContent = ShowContent & "   <td style=""text-align:center"">" & vbcrlf
			 If Privacy=2 Then ShowContent = ShowContent & "保密"  Else ShowContent = ShowContent &  RS("LastLoginTime") & "</td>" & vbcrlf
            ShowContent = ShowContent & "    <td style=""text-align:center"">" & vbcrlf
			If Privacy=2 Then ShowContent = ShowContent & "保密"  Else ShowContent = ShowContent &  RS("LoginTimes") & "</td>" & vbcrlf
           ShowContent = ShowContent & "   </tr> <tr><td colspan=7 background='images/line.gif'></td></tr>" & vbcrlf
             RS.MoveNext
			I = I + 1
				If I >= MaxPerPage Then Exit Do
			 Loop
			 
			 ShowContent = ShowContent & "<tr><td colspan=7 align=""right"">" & vbcrlf
			 ShowContent = ShowContent & KS.ShowPagePara(totalPut, MaxPerPage, "", True, "位", CurrentPage, "ListType=" & ListType)
			 ShowContent = ShowContent & "</td></tr>" & vbcrlf
			End Function

End Class
%> 

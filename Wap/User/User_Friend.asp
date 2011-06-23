<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的好友">
<p>
<%
Dim KSCls
Set KSCls = New User_Friend
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Friend
        Private KS,Prev,DomainStr
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,SQL,TableBody,strErr,Action,BoxName,smsCount,smsType,readAction,TUrl
		Private ArticleStatus,TotalPages
		Private Sub Class_Initialize()
			MaxPerPage = 16
		    Set KS=New PublicCls
			DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect DomainStr&"User/Login/"
			   Exit Sub
			End If
			%>
            <a href="User_Friend.asp?ListType=1&amp;<%=KS.WapValue%>">好友</a>
            <a href="User_Friend.asp?ListType=2&amp;<%=KS.WapValue%>">陌生人</a>
            <a href="User_Friend.asp?ListType=3&amp;<%=KS.WapValue%>">黑名单</a>
            <br/>
			<%
			Action=Trim(Request("Action"))
			CurrentPage=Trim(Request("page"))
			If Isnumeric(CurrentPage) Then
			   CurrentPage=Clng(CurrentPage)
			Else
			   CurrentPage=1
			End If
			Select Case Action
			    Case "add":Call AddFriend()'添加好友
				Case "edit":Call AddFriend()'修改好友资料
				Case "addsave":Call addsave()
				Case "del":Call Del()
				Case "note":Call note()'查看备注
				Case "info":Call info()'我的好友
				Case "addF":Call addF()'添加好友
				Case "saveF":Call saveF()
				Case "DelFriend":Call DelFriend()'删除
				Case "AllDelFriend":Call AllDelFriend()'清空好友
				Case Else:Call info()'我的好友
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
	    End Sub
	    '删除好友
	    Sub Del()
	        Conn.Execute("delete from ks_friend where id=" & KS.chkclng(KS.S("id")))
		    Response.redirect DomainStr&"User/User_Friend.asp?" & KS.WapValue & ""
	    End Sub
		Sub note()
		    Prev=True
		    Dim ID:ID=KS.Chkclng(KS.S("ID"))
			Dim RS:set RS=server.createobject("adodb.recordset")
			RS.Open "select * from KS_Friend where ID=" & ID,Conn,1,1
			If RS.EOF And RS.BOF Then
			   RS.Close:set RS=Nothing:Response.Write "参数传递出错!<br/>":Exit Sub
			Else
			%>
            【查看备注】<br/>
            用 户 名：<%=RS("Friend")%><br/>
            真实姓名：<%=RS("RealName")%><br/>
            联系电话：<%=RS("Phone")%><br/>
            手机号码：<%=RS("Mobile")%><br/>
            Q Q 号码：<%=RS("QQ")%><br/>
            电子邮箱：<%=RS("Email")%><br/>
            备注信息：<%=RS("Note")%><br/>
            <a href="User_Message.asp?Action=new&amp;ToUser=<%=KS.HTMLEncode(RS(1))%>&amp;<%=KS.WapValue%>">发送短信</a>
            <a href="User_Friend.asp?Action=edit&amp;id=<%=RS("id")%>&amp;<%=KS.WapValue%>">修改</a>
            <a href="User_Friend.asp?Action=del&amp;id=<%=RS(0)%>&amp;<%=KS.WapValue%>">移除</a>
            <br/><br/>
            <%
			End If
		End Sub
		'添加好友
		Sub AddFriend()
		    Dim flag,username,realname,phone,mobile,qq,msn,email,note
		    Dim ID:ID=KS.Chkclng(KS.S("ID"))
			If KS.S("Action")="edit" Then
			   Dim RS:set RS=server.createobject("adodb.recordset")
			   RS.Open "select * from ks_friend where id=" & id,Conn,1,1
			   If RS.EOF And RS.BOF Then
			      RS.Close:set RS=Nothing
				  Response.Write "参数传递出错!<br/>"
				  Prev=True
				  Exit Sub
			   Else
			      UserName=RS("Friend")
				  Flag=RS("Flag")
				  RealName=RS("RealName")
				  Phone=RS("Phone")
				  Mobile=RS("Mobile")
				  QQ=RS("QQ")
				  Msn=RS("Msn")
				  Email=RS("Email")
				  Note=RS("Note")
			   End If
			   RS.Close:set RS=Nothing
 			Else
			   Flag=KS.S("Flag")
			End If
			%>
            【添加好友】<br/>
            用户名,登录会员中心的用户名，必须填写。<br/>
            <input type="text" name="UserName<%=Minute(Now)%><%=Second(Now)%>" value="<%=UserName%>" /><br/>
            类 型:<select name="Flag" value="<%=Flag%>"><option value="1">好朋友</option><option value="2">陌生人</option><option value="3">黑名单</option></select><br/>
            真实姓名:<input type="text" value="<%=RealName%>" name="RealName<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            联系电话:<input type="text" value="<%=Phone%>" name="Phone<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            手机号码:<input type="text" value="<%=Mobile%>" name="Mobile<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            Q Q 号码:<input type="text" value="<%=QQ%>" name="QQ<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            MSN 号码:<input type="text" value="<%=Msn%>" name="Msn<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            电子邮箱:<input type="text" value="<%=Email%>" name="Email<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            备注信息:<input type="text" value="<%=Note%>" name="Note<%=Minute(Now)%><%=Second(Now)%>" /><br/>
            <anchor>确定保存<go href="User_Friend.asp?Action=addsave&amp;<%=KS.WapValue%>" method="post">
            <postfield name="ID" value="<%=ID%>"/>
            <postfield name="UserName" value="$(UserName<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Flag" value="$(Flag)"/>
            <postfield name="RealName" value="$(RealName<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Phone" value="$(Phone<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Mobile" value="$(Mobile<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="QQ" value="$(QQ<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Msn" value="$(Msn<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Email" value="$(Email<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Note" value="$(Note<%=Minute(Now)%><%=Second(Now)%>)"/>
            </go></anchor><br/>
			<%
		End Sub
		'保存
		Sub addsave()
		    Dim Flag:Flag=KS.Chkclng(KS.S("Flag"))
			Dim UserName:UserName=KS.R(KS.S("UserName"))
			Dim RealName:RealName=KS.R(KS.S("RealName"))
			Dim Phone:Phone=KS.R(KS.S("Phone"))
			Dim Mobile:Mobile=KS.R(KS.S("Mobile"))
			Dim QQ:QQ=KS.R(KS.S("QQ"))
			Dim Msn:Msn=KS.R(KS.S("Msn"))
			Dim Email:Email=KS.R(KS.S("Email"))
			Dim Note:Note=KS.S("Note")
			If UserName="" Then Response.Write "用户名必须填写!<br/>":Prev=True:Exit Sub
			If UserName=KSUser.UserName Then Response.Write "不能将自己加为好友!<br/>":Prev=True:Exit Sub
			If Len(Note)>255 Then Response.Write "备注信息必须小于255个字符!<br/>":Prev=True:Exit Sub
			
			Dim RS:set RS=server.createobject("adodb.recordset")
			RS.open "select username from ks_user where username='" & username & "'",conn,1,1
			if RS.eof and RS.bof then
		       RS.close:set RS=nothing
			   Response.Write "对不起，你输入的用户名不存在!<br/>"
			   Prev=True
			   exit sub
		    end if
			RS.Close
			RS.Open "select * from ks_friend where friend='" & UserName & "'",conn,1,3
			If RS.EOF Then
			   RS.Addnew
			End If
			RS("UserName")=KSUser.UserName
			RS("Friend")=UserName
			RS("AddTime")=Now
			RS("RealName")=RealName
			RS("Phone")=Phone
			RS("Mobile")=Mobile
			RS("qq")=QQ
			RS("Msn")=Msn
			RS("Email")=Email
			RS("Note")=Note
			RS("Flag")=Flag
			RS.Update
			RS.Close:set RS=Nothing
			If KS.chkclng(KS.S("ID"))<>0 Then
			   Response.Write "好友资料修改成功！<br/>"
		    Else
			   Response.Write "好友添加成功，继续添加吗?"
			   Response.Write "<a href=""User_Friend.asp?Action=add&Flag=" & Flag & "&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_Friend.asp?"&KS.WapValue&""">取消</a><br/>"
		    End If
		End Sub
		
		Sub info()
		    Dim Param,I
			Select Case KS.Chkclng(KS.S("listtype"))
			    Case 1:response.write "【好 朋 友】<br/>"
			    Case 2:response.write "【陌 生 人】<br/>"
				Case 3:response.write "【黑 名 单】<br/>"
				Case Else:response.write "【我的好友】<br/>"
			End Select
		    If KS.Chkclng(KS.S("listtype"))<>0 Then Param=Param & " And flag=" & KS.Chkclng(KS.S("listtype"))
			set RS=server.createobject("adodb.recordset")
			sql="select F.id,f.friend,f.flag,f.phone,f.mobile,f.note,f.email,f.QQ,f.msn,U.Username,f.Email,f.RealName,U.HomePage from KS_Friend F inner join KS_User U on F.Friend=U.UserName where F.Username='"&KSUser.UserName&"' " & Param & " order by F.addtime desc"
			RS.Open sql,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Select Case KS.S("listtype")
			       Case "2":Response.Write "你没有添加陌生人。<br/>"
				   Case "3":Response.Write "你没有添加黑名单。<br/>"
				   Case Else:Response.Write "你没有添加好朋友。<br/>"
			   End Select
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
			   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Do while Not RS.EOF
				  %>
                  <%=((i+1)+CurrentPage*MaxPerPage)-MaxPerPage%>.<a href="User_Friend.asp?Action=note&amp;id=<%=RS("id")%>&amp;<%=KS.WapValue%>"><%=KS.HTMLEncode(RS(1))%></a>
                  <a href="User_Message.asp?Action=new&amp;ToUser=<%=KS.HTMLEncode(RS(1))%>&amp;<%=KS.WapValue%>">发送短信</a>
                  <br/>
                  <%
				  Dim DelID:DelID=DelID&RS(0)&","
				  RS.Movenext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   loop
			 End If
			RS.Close:set RS=Nothing
			%>
            <br/>
            <a href="User_Friend.asp?Action=DelFriend&amp;ID=<%=DelID%>&amp;<%=KS.WapValue%>">删除本页纪录</a><br/>
            <a href="User_Friend.asp?Action=addF&amp;<%=KS.WapValue%>">快速添加好友</a><br/>
            <a href="User_Friend.asp?Action=AllDelFriend&amp;ID=<%=DelID%>&amp;<%=KS.WapValue%>">清空所有纪录</a><br/>
		<%
		End Sub
		
		Sub DelFriend()
		    Dim DelID
			DelID=Replace(KS.S("ID"),"'","")
			If KS.S("Checked")="ok" Then
			   DelID=KS.FilterIDs(DelID)
			   If DelID="" or isnull(DelID) Then
			      Response.Write "您没有要删除好友名单。<br/>":Prev=True:Exit Sub
			   Else
			      Conn.Execute("delete from KS_Friend where UserName='"&KSUser.UserName&"' And id in ("&DelID&")")
				  Response.Write "您已经删除成功。<br/>"
			  End If
			Else
			   Response.Write "删除的好友名单将不可恢复。确定要删除吗？"
			   Response.Write "<a href=""User_Friend.asp?Action=DelFriend&amp;ID="&DelID&"&amp;Checked=ok&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_Friend.asp?"&KS.WapValue&""">取消</a><br/>"
			End If
		End Sub
		
		Sub AllDelFriend()
		    If KS.S("Checked")="ok" Then
			   Conn.Execute("delete from KS_Friend where UserName='"&KSUser.UserName&"'")
			   Response.Write "您已经删除了所有好友列表。<br/>"
			Else
			   Response.Write "删除的好友名单将不可恢复。确定要删除吗？"
			   Response.Write "<a href=""User_Friend.asp?Action=AllDelFriend&amp;Checked=ok&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_Friend.asp?"&KS.WapValue&""">取消</a><br/>"
			End If
		End Sub
		
		Sub addF()
		    Call UserList()
			%>
            【批量添加】<br/>
            <input type="text" name="ToUser" value="<%=Request("MyFriend")%>" />
            <anchor>保存<go href="User_Friend.asp?Action=saveF&amp;<%=KS.WapValue%>" method="post">
            <postfield name="ToUser" value="$(ToUser)"/>
            </go></anchor><br/>
            用户之间使用逗号(,)分开<br/>
			<%
		End Sub
		
		Sub saveF()
		    Dim InCept,i
			If Request("ToUser")="" Then
			   Response.Write "请填写对象。<br/>":Prev=True:Exit Sub
			Else
			   InCept=KS.R(Request("ToUser"))
			   InCept=Split(incept,",")
			End If
			
			For i=0 To Ubound(InCept)
			    set RS=server.createobject("adodb.recordset")
				sql="select UserName from KS_User where UserName='"&incept(i)&"'"
				set RS=Conn.Execute(sql)
				If RS.EOF And RS.BOF Then
				   Response.Write "系统没有（"&InCept(i)&"）这个用户，操作未成功。<br/>":Prev=True:Exit Sub
				End If
				set RS=Nothing
				If KSUser.UserName=Trim(InCept(i)) Then
				   Response.Write "不能把自已添加为好友。<br/>":Prev=True
				End If
				sql="select friend from KS_Friend where UserName='"&KSUser.UserName&"' and  friend='"&InCept(i)&"'"
				set RS=Conn.Execute(sql)
				If RS.EOF And RS.BOF Then
				   sql="insert into KS_Friend (UserName,Friend,AddTime,Flag) values ('"&KSUser.UserName&"','"&Trim(InCept(i))&"',"&SqlNowString&",1)"
				   set RS=Conn.Execute(sql)
				End If
				'If i>5 Then
				   'Response.Write "每次最多只能添加5位用户，您的名单5位以后的请重新填写。<br/>":Prev=True:Exit Sub:Exit For
				'End If
		    Next
			Response.Write "恭喜您，好友添加成功。<br/>":Prev=True:Exit Sub
		End Sub
		
		Sub UserList()
		    Dim i,n
		    Response.Write "【管理员组】<br/>"
			sql="select UserName,Sex,QQ,Email from KS_User where GroupID=4 order by UserID"
			set RS=Conn.Execute(sql)
			i=0
			Do while not RS.EOF
			   If KSUser.UserName=RS(0) Then
			      Response.Write "<a href=""User_Friend.asp?Action=saveF&amp;ToUser="&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</a> "
			   Else
			      Response.Write "<a href=""User_Friend.asp?Action=saveF&amp;ToUser="&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</a> "
			   End If
			   i=i+1
			   If i>=6 Then
			      Response.Write "<br/>"
			      i=0
			   End If
			   RS.Movenext
			Loop
			Response.Write "<br/>"
			set RS=Nothing
			Response.Write "【网站会员】<br/>"
			
			sql="select UserName,Sex,QQ,Email from KS_User where GroupID<>4 order by UserID"
			set RS=Server.CreateObject("adodb.recordSet")
			RS.Open sql,Conn,1,1
			i=0:n=0:TotalPut=0
			If Not (RS.EOF And RS.BOF) Then
			   TotalPut=RS.recordcount
			   If (TotalPut Mod MaxPerPage)=0 Then
			      TotalPages = TotalPut \ MaxPerPage
			   Else
			      TotalPages = TotalPut \ MaxPerPage + 1
			   End If
			   If CurrentPage > TotalPages Then CurrentPage=TotalPages
			   If CurrentPage < 1 Then CurrentPage=1
			   RS.Move (CurrentPage-1)*MaxPerPage
			   Do while not RS.EOF
			      If KSUser.UserName=RS(0) Then
				     Response.Write "<a href=""User_Friend.asp?action=saveF&amp;touser="&RS(0)&"&amp;" & KS.WapValue & """>"&RS(0)&"</a> "
				  Else
				     Response.Write "<a href=""User_Friend.asp?action=saveF&amp;touser="&RS(0)&"&amp;" & KS.WapValue & """>"&RS(0)&"</a> "
				  End If
				  i=i+1
				  If i>=6 Then 
				     If i=6 Then Response.Write "<br/>"
					 i=0
				  End If
				  n=n+1
				  If n>= MaxPerPage Then Exit Do
				  RS.Movenext
			   loop
			   Response.Write "<br/>"
			   Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
			   Call KS.ShowPageParamter(totalPut, MaxPerPage,"User_Friend.asp", false, "个用户", CurrentPage, "Action=addF&amp;flag=" &KS.S("flag")& "&amp;" & KS.WapValue & "")
			Else
			   Response.Write "无任何用户<br/>"
			End If
			Response.Write "<br/>"
			set RS=Nothing
		end sub
End Class
%> 

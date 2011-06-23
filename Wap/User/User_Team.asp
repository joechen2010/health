<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="UpFileSave.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的圈子">
<p>
<%
Set KS=New PublicCls
Action=Trim(Request("Action"))

IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If
%>
<%
id=KS.G("id")
If KS.SSetting(0)=0 Then
   Response.write "对不起，本站关闭个人空间功能！<br/>"
ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)=0 Then
   Response.write "您还没有开通个人空间！<br/>"
ElseIf Conn.Execute("Select status From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)<>1 Then
   Response.write "对不起，你的空间还没有通过审核或被锁定！<br/>"
Else

          Select Case Action
			 Case "MyJoinTeam"
			  Call MyJoinTeam()'我加入的圈子
			 Case "VerificApply"
			  Call VerificApply()'审核成员
			 Case "AcceptApply"
			  Call AcceptApply()'接受请求
			 Case "ApplyDel" 
			  Call ApplyDel()'拒绝申请
			 Case "DelTeam"
			  Call DelTeam()'删除圈子
			 Case "EditTeam","CreateTeam"
			  Call Managexc()'圈子，添加／修改
			 Case "Teamsave"
			  Call Teamsave()'保存圈子
			 Case "upfileImage"
			  Call upfileImage()'圈子图片
			 Case Else
			  Call MyTeam() '我建的圈子
			End Select

End If
'圈子，添加／修改=================================================================================

Sub Managexc()
    If KS.ChkClng(KS.SSetting(6))<>0 Then
	   If Conn.Execute("select count(id) from ks_team where UserName='"& KSUser.UserName & "'")(0)>=KS.ChkClng(KS.SSetting(6)) Then
	      Response.write "对不起，每个用户最多只能建 " & KS.SSetting(6) & " 个圈子！<br/>"
	   Else
	      If id<>"" Then
	         Set rs=Server.CreateObject("ADODB.RECORDSET")
			 rs.Open "Select * From KS_Team Where id="&id&"",conn,1,1
			 If Not rs.EOF Then
		        TeamName=rs("TeamName")'圈子名称
				ClassID=rs("ClassID")'圈子分类
				Note=RS("Note")'圈子说明
				JoinTF=rs("JoinTF")'加入条件
				Announce=rs("Announce")'圈子公告
			 End if
			 rs.Close:Set rs=Nothing
			 OpStr="确定修改"
			 TipStr="修改我的圈子"
		  Else
		     Announce="暂无公告!"
		     OpStr="立即创建"
		     TipStr="创建我的圈子"
		  End if

%>
=<%=TipStr%>=<br/>
圈子名称：<input name="TeamName<%=minute(now)%><%=second(now)%>" type="text" maxlength="40" size="20" value="<%=TeamName%>"/><br/>
圈子分类：<select name='ClassID'>
<option value="0">-请选择类别-</option>
<%
Set rs=Server.CreateObject("ADODB.RECORDSET")
    rs.Open "Select * From KS_TeamClass order by orderid",conn,1,1
	If Not rs.EOF Then
	   Do While Not rs.Eof 
	      Response.Write "<option value=""" & rs("ClassID") & """>" & rs("ClassName") & "</option>"
		  rs.MoveNext
	   Loop
	End If
	rs.Close
	Set rs=Nothing
%></select><br/>      
加入条件：<select name='JoinTF'>
<option value="1">任意加入</option>
<option value="2">申请加入</option>
<option value="3">仅可邀请</option>
</select>  <br/> 
圈子说明：<input name="Note<%=minute(now)%><%=second(now)%>" type="text" maxlength="40" size="20" value="<%=Note%>"/><br/>
圈子公告：<input name="Announce<%=minute(now)%><%=second(now)%>" type="text" maxlength="40" size="20" value="<%=Announce%>"/><br/>
<anchor><%=OpStr%><go href='User_Team.asp?Action=Teamsave&amp;id=<%=id%>&amp;<%=KS.WapValue%>' method='post' accept-charset="utf-8">
<postfield name='TeamName' value='$(TeamName<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ClassID' value='$(ClassID)'/>
<postfield name='JoinTF' value='$(JoinTF)'/>
<postfield name='Note' value='$(Note<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Announce' value='$(Announce<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>
<br/>
<%
          End If
    End If
End Sub

'保存圈子=================================================================================
Sub Teamsave()
    TeamName=KS.S("TeamName")
	ClassID=KS.S("ClassID")
	JoinTF=KS.S("JoinTF")
	Note=KS.S("Note")
	Announce=KS.S("Announce")
	PhotoUrl="/images/nopic.gif"
	Point=0
    If TeamName="" Then
	   Response.write "出错提示，请输入圈子名称！<br/>"
	   Response.Write "<a href='User_Team.asp?action=Managexc&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
    ElseIF ClassID=0 Then
	   Response.write "出错提示，请选择圈子类型！<br/>"
	   Response.Write "<a href='User_Team.asp?action=Managexc&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"     
    Else
           Set rs=Server.CreateObject("ADODB.RECORDSET")
		      If id="" Then
			     sql="Select * From KS_Team"
			  Else
			     sql="Select * From KS_Team Where id="&id&""
			  End If
			  rs.Open sql,conn,1,3
                 If id="" Then
				    rs.AddNew
					rs("AddDate")=now
					If KS.SSetting(5)=1 Then
					   RS("Verific")=0
					Else
					   RS("Verific")=1 '设为已审
					End If
					rs("UserName")=KSUser.UserName
				 End If
				    rs("TeamName")=TeamName
					rs("ClassID")=ClassID
					rs("Note")=Note
					rs("JoinTF")=JoinTF
					rs("Point")=Point
					rs("PhotoUrl")=PhotoUrl
					rs("Announce")=Announce
					RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=3 and IsDefault='true'")(0))
					rs.Update
					rs.Close:Set rs=Nothing
                  If id="" Then
		          Set rs1=Server.CreateObject("ADODB.RECORDSET")
				      rs1.open "select * from ks_teamusers",conn,1,3
					  rs1.addnew
					  rs1("teamid")=conn.execute("select max(id) from ks_team")(0)
					  rs1("username")=KSUser.UserName
					  rs1("power")=2
					  rs1("status")=3
					  rs1("applydate")=now
					  rs1("adddate")=now
					  rs1("reason")="创建圈子"
					  rs1.update
					  rs1.Close:Set rs1=Nothing
				   End If		  
		   Response.write "圈子创建/修改成功。<br/>"
		   Response.Write "<a href='User_Team.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"
	End If
End Sub

'我建的圈子=================================================================================
Sub MyTeam()
%>
=我建的圈子=<br/>
<%=KS.GetReadMessage%>
<a href="User_Team.asp?Action=CreateTeam&amp;<%=KS.WapValue%>">创建圈子</a>
<a href="User_Team.asp?Action=MyJoinTeam&amp;<%=KS.WapValue%>">我加入的圈子</a><br/>

<%
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If
set rs=server.createobject("adodb.recordset")
sql="select * from KS_Team where UserName='"& KSUser.UserName &"' order by AddDate DESC"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "您还没有创建圈子!<br/>"
else
   MaxPerPage =2
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   
   Do While Not RS.Eof
   %>
   ---------<br/>
   <img src="<%=rs("photourl")%>" width="80" height="60" /><br/>
   <a href="User_UpFile.asp?id=<%=rs("id")%>&amp;ChannelID=9996&amp;<%=KS.WapValue%>">修改圈子图片</a><br/>
   名称:<a href="../Space/Group.asp?id=<%=rs("id")%>&amp;<%=KS.WapValue%>"><%=RS("teamName")%></a><br/>
   成员:<%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%>
   创建人:<%=KS.GetUserRealName(rs("username"))%>
   <br/>
   主题:<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>
   回复:<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%>
   时间:<%=formatdatetime(rs("adddate"),2)%><br/>
   管理:<a href="User_Team.asp?Action=VerificApply&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">审核成员(<%=conn.execute("select count(username)  from ks_teamusers where status=2 and teamid=" & rs("id"))(0)%>)</a> 
   <a href="User_Team.asp?action=EditTeam&amp;ID=<%=rs("id")%>&amp;<%=KS.WapValue%>">修改</a>
   <a href="User_Team.asp?action=DelTeam&amp;ID=<%=rs("id")%>&amp;<%=KS.WapValue%>">删除</a>
   <a href="../Space/Group.asp?id=<%=rs("id")%>&amp;<%=KS.WapValue%>">访问</a>
   <br/>
   <%
      RS.MoveNext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   Loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Team.asp", False, "个", CurrentPage, "Action=MyTeam&amp;" & KS.WapValue & "")
   Response.Write "<br/>"
end if
rs.Close:Set rs=Nothing
End Sub

'我加入的圈子=================================================================================
Sub MyJoinTeam()
%>
=我加入的圈子=<br/>
<%=KS.GetReadMessage%>
<a href="User_Team.asp?Action=CreateTeam&amp;<%=KS.WapValue%>">创建圈子</a>
<a href="User_Team.asp?Action=MyTeam&amp;<%=KS.WapValue%>">我建的圈子</a><br/>
<%
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If
set rs=server.createobject("adodb.recordset")
sql = "select b.username,b.id,b.teamname,b.photourl,b.adddate from ks_teamusers a, ks_team b where a.status=3 and a.teamid=b.id and a.username='" & KSUser.UserName & "' and b.username<>'" & KSUser.UserName & "' order by a.Adddate desc"
rs.Open sql,Conn,1,1
if rs.bof and rs.eof then
   response.write "你没有加入任何圈子!<br/>"
else
   MaxPerPage =3
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   
   Do While Not RS.Eof
   %>
---------<br/>
<img src="<%=rs("photourl")%>" width="80" height="60" /><br/>
圈子名称:<a href="../Space/Group.asp?id=<%=rs("id")%>&amp;<%=KS.WapValue%>"><%=RS("teamName")%></a><br/>
创建人:<%=rs("username")%> 成员数:<%=conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0)%> <br/>
主题/回复:<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0)%>/<%=conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0)%><br/>
创建时间:<%=formatdatetime(rs("adddate"),2)%><br/>
管理操作:<a href="User_Team.asp?Action=OutTeam&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">删除</a> <a href="../Group.asp?action=ShowGroupInfo&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">信息</a> <a href="Group.asp?action=PostTopic&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">发表新贴</a>
<br/>
   <%
      RS.MoveNext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   Loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Team.asp", False, "个", CurrentPage, "Action=MyJoinTeam&amp;" & KS.WapValue & "")
   Response.Write "<br/>"
End If
rs.Close:Set rs=Nothing
End Sub

'审核成员=================================================================================
Sub VerificApply()
Response.write "=审核成员=<br/>"
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If
set rs=server.createobject("adodb.recordset")
sql = "select * from KS_TeamUsers where TeamID="&id&" and status=2  order by AddDate DESC"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "没有用户申请加入该圈子!<br/>"
else
   MaxPerPage =10
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   
   Do While Not RS.Eof
   %>
   申请人:<%=rs("username")%><br/>
   申请时间:<%=formatdatetime(rs("applyDate"),2)%><br/>
   加入理由:<%=RS("reason")%><br/>
   操作:<a href="User_Team.asp?id=<%=rs("id")%>&amp;Action=AcceptApply&amp;<%=KS.WapValue%>">接受申请</a> <a href="User_Team.asp?action=ApplyDel&amp;ID=<%=rs("id")%>&amp;<%=KS.WapValue%>">拒绝</a><br/>
   -----------<br/> 
   <%
      RS.MoveNext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   Loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Team.asp", True, "个", CurrentPage, "Action=Apply&amp;" & KS.WapValue & "")
end if
rs.Close:Set rs=Nothing
End Sub

Sub upfileImage()
    sTemp = Kesion
	sTempArr=split(sTemp,"|||")
	If Cbool(sTempArr(0))=False Then
	   Response.Write sTempArr(1)
	Else
	   PhotoUrl = Replace(sTempArr(1),"|","")
	   Conn.Execute("update KS_Team set PhotoUrl='"&PhotoUrl&"' where id="&id&"")
	   response.Write "修改圈子图片成功。<br/>"
	   response.Write "<a href=""User_Team.asp?" & KS.WapValue & """>我的圈子</a><br/>"
	End If
End Sub

'接受请求=================================================================================
Sub AcceptApply()
Set rs=server.createobject("adodb.recordset")
  rs.open "select * from ks_teamusers where id=" & id ,conn,1,3
  if not rs.eof then
     rs("status")=3
	 rs("adddate")=now
	 rs.update
title=""&KSUser.UserName&"通过加入圈子的申请!"
message=""&rs("username")&"您好!<br>您加入圈子[<a href=""Group.asp?id="&rs("teamid")&""">" & conn.execute("select teamname from ks_team where id="&rs("teamid"))(0)&"</a>]的申请已于"&now&"通过审核，现在您可以参与该圈子的讨论!"
SqlStr="insert into KS_Message (Incept,sEnder,title,content,SendTime,flag,IsSend,DelR,DelS) values ('"&rs("username")&"','"&KSUser.UserName&"','"&title&"','"&message&"','"&Now()&"',0,1,0,0)"
conn.Execute(SqlStr)
  end if
Response.write "审核成功。<br/>"
response.Write "<a href=""User_Team.asp?action=VerificApply&amp;id="&id&"&amp;" & KS.WapValue & """>审核成员</a><br/>"
rs.Close:Set rs=Nothing
End Sub

'拒绝申请=================================================================================
Sub ApplyDel()
     set rs=server.createobject("adodb.recordset")
	     rs.open "select * from ks_teamusers where id in(" & id & ")",conn,1,3
	 if not rs.eof then
	     title=""&KSUser.UserName&"申请加入圈子被拒绝!"
		 message=""&rs("username")&"您好!<br>您加入圈子[<a href=""Group.asp?id="&rs("teamid")&""">" & conn.execute("select teamname from ks_team where id="&rs("teamid"))(0)&"</a>]的申请已于"&now&"被群主拒绝!"
		 SqlStr="insert into KS_Message (Incept,sEnder,title,content,SendTime,flag,IsSend,DelR,DelS) values ('"&rs("username")&"','"&KSUser.UserName&"','"&title&"','"&message&"','"&Now()&"',0,1,0,0)"
		 conn.Execute(SqlStr)
	 end if
	 rs.Close:Set rs=Nothing
conn.execute("delete from ks_teamusers where id in(" & id & ")")
Response.write "拒绝申请成功。<br/>"
response.Write "<a href=""User_Team.asp?action=VerificApply&amp;id="&id&"&amp;" & KS.WapValue & """>审核成员</a><br/>"
End Sub

'删除圈子=================================================================================
Sub DelTeam()
    Conn.Execute("Delete From KS_Team Where id In("&id&")")
	Conn.execute("delete from ks_teamusers where TeamID in("&id&")")
	Conn.execute("delete from ks_teamtopic where TeamID in("&id&")")
	Response.write "删除圈子成功。<br/>"
	Response.Write "<a href=""User_Team.asp?action=&amp;" & KS.WapValue & """>审核成员</a><br/>"
End Sub

Sub OutTeam()
  	Conn.execute("delete from ks_teamusers where id in(" & id& ")")
	Response.write "退出的圈子成功。<br/>"
	Response.Write "<a href=""User_Team.asp?action=&amp;" & KS.WapValue & """>审核成员</a><br/>"
End Sub
%>
<br/>
<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>

<%
Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p>
</card>
</wml>

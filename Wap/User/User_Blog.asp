<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的日志">
<p>
<%Set KS=New PublicCls%>
<%
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If
%>
<%
id=KS.S("id")
Action=Trim(KS.S("Action"))

If KS.SSetting(0)=0 Then
   Response.write "对不起，本站关闭个人空间功能！<br/>"
ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)=0 Then
   'Response.write "您还没有开通个人空间！<br/>"
   Select Case Action
      Case "BlogEdit"
	  Call ApplyBlog()'申请日志
	  Case "ApplyBlogSave"
	  Call ApplyBlogSave()
	End Select
ElseIf Conn.Execute("Select status From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)<>1 Then
   Response.write "对不起，你的空间还没有通过审核或被锁定！<br/>"
Else

		Select Case Action
			 Case "Comment"
			  Call Comment()'评论管理
			 Case "ReplayComment"
			  Call ReplayComment()'回复评论
			 Case "SaveCommentReplay"
			  Call SaveCommentReplay()'保存评论回复
			 Case "CommentDel"
			  Call CommentDel()'删除评论
			 Case "Message"
			  Call Message()'留言管理
			 Case "ReplayMessage"
			  Call ReplayMessage()'回复留言
			 Case "SaveMessageReplay"
			  Call SaveMessageReplay()'保存留言回复
			 Case "MessageDel"
			  Call MessageDel()'删除留言
			 Case "ArticleDel"
			  Call ArticleDel()'删除日志
			 Case "Add","Edit"
			  Call ArticleAdd()'添加日志
			 Case "AddSave"
			  Call AddSave()'发布日志
			 Case "EditSave"
			  Call EditSave()'日志修改
			 Case "BlogEdit"
			  Call ApplyBlog()'申请日志
			 Case "ApplyBlogSave"
			  Call ApplyBlogSave()'保存申请日志
			 Case "rizhixuxie"
			  Call rizhixuxie()'日志续写
			 Case "xuxiebaocun"
			  Call xuxiebaocun()'续写保存
			 Case "rizhichakan"
			  Call rizhichakan()'日志查看
			 Case Else
			  Call BlogList()'日志列表
			End Select
End If


'申请日志=================================================================================
Sub ApplyBlog()
	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_Blog Where UserName='" & KSUser.UserName &"'",Conn,1,1
	If Not RS.EOF Then
	   BlogName=RS("BlogName")
	   domain=RS("domain")
	   ClassID=RS("ClassID")
	   Descript=RS("Descript")
	   Announce=RS("Announce")
	   ContentLen=RS("ContentLen")
	   ListBlogNum=RS("ListBlogNum")
	   ListLogNum=RS("ListLogNum")
	   ListReplayNum=RS("ListReplayNum")
	   ListGuestNum=RS("ListGuestNum")
	   OpStr="OK了,确定修改":TipStr="修改个人空间参数"
	Else
	  
	   BlogName=KSUser.UserName & "的个人空间"
	   domain=KSUser.UserName
	   ClassID="0"
	   ContentLen=100
	   ListBlogNum=3
	   ListLogNum=3
	   ListReplayNum=3
	   ListGuestNum=3
	   Announce="没有公告!"
	   OpStr="OK了,立即申请":TipStr="申请开通个人空间"
    End if
	RS.Close:Set RS=Nothing
%>

<%=KS.GetReadMessage%>
<%=TipStr%><br/>
空间名称:<input name="BlogName<%=minute(now)%><%=second(now)%>" type="text" value="<%=BlogName%>" maxlength="100" /><br/>
空间分类:<select name='ClassID'>
        <option value="0">-请选择类别-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
								  RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select><br/>
站点描述:<input name="Descript<%=minute(now)%><%=second(now)%>" type="text" value="<%=BlogName%>" maxlength="500" /><br/>
空间公告:<input name="Announce<%=minute(now)%><%=second(now)%>" type="text" value="<%=BlogName%>" maxlength="500" /><br/>
日志默认部分显示字数:<input name="ContentLen<%=minute(now)%><%=second(now)%>" type="text" value="<%=ContentLen%>" /><br/>
每页显示日志篇数:<input name="ListBlogNum<%=minute(now)%><%=second(now)%>" type="text" value="<%=ListBlogNum%>" /><br/>
显示最新回复条数:<input name="ListReplayNum<%=minute(now)%><%=second(now)%>" type="text" value="<%=ListReplayNum%>" /><br/>
显示最新日志篇数:<input name="ListLogNum<%=minute(now)%><%=second(now)%>" type="text" value="<%=ListLogNum%>" /><br/>
显示最新留言条数:<input name="ListGuestNum<%=minute(now)%><%=second(now)%>" type="text" value="<%=ListGuestNum%>" /><br/>

<anchor><%=OpStr%><go href="User_Blog.asp?Action=ApplyBlogSave&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
<postfield name='BlogName' value='$(BlogName<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ClassID' value='$(ClassID)'/>
<postfield name='Descript' value='$(Descript<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Announce' value='$(Announce<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ContentLen' value='$(ContentLen<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ListBlogNum' value='$(ListBlogNum<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ListReplayNum' value='$(ListReplayNum<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ListLogNum' value='$(ListLogNum<%=minute(now)%><%=second(now)%>)'/>
<postfield name='ListGuestNum' value='$(ListGuestNum<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>
<br/>
<%
End Sub

'保存个人空间申请=================================================================================
Sub ApplyBlogSave()
    BlogName=KS.GotTopic(KS.S("BlogName"),16)'博客名称
	ClassID=KS.S("ClassID")''站点类型
	Descript=KS.GotTopic(KS.S("Descript"),500)'站点描述
	Announce=KS.GotTopic(KS.S("Announce"),500)'公告
	ContentLen=500'日志默认部分显示字数
	ListBlogNum=10'每页显示日志篇数
	ListLogNum=10'显示最新日志篇数
	ListReplayNum=10'显示最新回复条数
	ListGuestNum=10'显示最新留言条数
	
	Dim BlogName:BlogName=KS.S("BlogName")
	'Dim Domain:Domain=KS.S("Domain")
	Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
	Dim Descript:Descript=KS.S("Descript")
	Dim Announce:Announce=KS.S("Announce")
	Dim ContentLen:ContentLen=KS.ChkClng(KS.S("ContentLen"))
	Dim ListBlogNum:ListBlogNum=KS.ChkClng(KS.S("ListBlogNum"))
	Dim ListLogNum:ListLogNum=KS.ChkClng(KS.S("ListLogNum"))
	Dim ListReplayNum:ListReplayNum=KS.ChkClng(KS.S("ListReplayNum"))
	Dim ListGuestNum:ListGuestNum=KS.ChkClng(KS.S("ListGuestNum"))


	TemplateID=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))'博客模板
	If BlogName="" Then
	   Response.write "出错提示，请输入博客站点名称！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=ApplyBlog&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	ElseIF ClassID=0 Then
	   Response.write "出错提示，请选择站点类型！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=ApplyBlog&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	'ElseIF domain<>"" And not Conn.Execute("select username from ks_Blog where username<>'" & KSUser.UserName & "' and [domain]='" & domain  &"'").eof Then
	   'Response.Write "对不起，你注册的二级域名已被其它用户使用！<br/>"
	   'Response.Write "<a href='User_Blog.asp?action=ApplyBlog&amp;" & KS.WapValue & "'>返回重写</a><br/>"
    Else
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	      RS.Open "Select * From KS_Blog Where UserName='"&KSUser.UserName&"'",conn,1,3
		  If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now'创建时间
			RS("TemplateID")=TemplateID'博客模板
			If KS.SSetting(2)=1 Then'站点状态：0未审1已审2锁定
			   RS("Status")=0
			Else
			   RS("Status")=1
			End If
		  End If
		    RS("UserName")=KSUser.UserName
		    RS("BlogName")=BlogName
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Announce")=Announce
			RS("ContentLen")=ContentLen
			RS("ListLogNum")=ListLogNum
			RS("ListBlogNum")=ListBlogNum
			RS("ListReplayNum")=ListReplayNum
			RS("ListGuestNum")=ListGuestNum
			RS.Update
			Response.write "操做成功！博客站点申请/修改成功。<br/>"
			Response.Write "<a href='User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"
			RS.Close:Set RS=Nothing
    End If
End Sub

'日志列表=================================================================================
Sub BlogList()
%>
<%=KS.GetReadMessage%>
<a href="User_Blog.asp?Action=Add&amp;<%=KS.WapValue%>">发布日志</a>
<a href="User_Blog.asp?Action=Comment&amp;<%=KS.WapValue%>">日志评论</a>
<br/>
<%
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If
Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
Status=KS.S("Status")
If Status<>"" and isnumeric(Status) Then 
   Param= Param & " and Status=" & Status
End If

IF KS.S("Flag")<>"" Then
   IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
   IF KS.S("Flag")=1 Then Param=Param & " And Tags like '%" & KS.S("KeyWord") & "%'"
End if
If KS.S("TypeID")<>"" And KS.S("TypeID")<>"0" Then Param=Param & " And TypeID=" & KS.ChkClng(KS.S("TypeID")) & ""
Dim Sql:sql = "select * from KS_BlogInfo "& Param &" order by AddDate DESC"
Select Case ks.s("Status")
    Case "0" 
	response.write "=已审日志=<br/>"
	Case "1"
	response.write "=草稿日志=<br/>"
	Case "2"
	response.write "=未审日志=<br/>"
	Case Else
	response.write "=所有日志=<br/>"
End Select
%>
<a href="User_Blog.asp?<%=KS.WapValue%>">所有</a>
<a href="User_Blog.asp?Status=0&amp;<%=KS.WapValue%>">已审[<%=conn.execute("select count(id) from KS_BlogInfo where Status=0 and UserName='"& KSUser.UserName &"'")(0)%>]</a>
<a href="User_Blog.asp?Status=2&amp;<%=KS.WapValue%>">未审[<%=conn.execute("select count(id) from KS_BlogInfo where Status=2 and UserName='"& KSUser.UserName &"'")(0)%>]</a>
<a href="User_Blog.asp?Status=1&amp;<%=KS.WapValue%>">草稿[<%=conn.execute("select count(id) from KS_BlogInfo where Status=1 and UserName='"& KSUser.UserName &"'")(0)%>]</a>
<br/>

<%
set rs=server.createobject("adodb.recordset")
sql = "select * from KS_BlogInfo "& Param &" order by AddDate DESC"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "没有你要的日志!<br/>"
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
   %>
   搜索:<select name="Flag">
   <option value="0">标题</option>
   <option value="1">标签</option>
   </select>
   <select size='1' name='TypeID'>
   <option value="0">请选择分类</option>
   <%
   Dim RS1:Set RS1=Server.CreateObject("ADODB.RECORDSET")
   RS1.Open "Select * From KS_BlogType order by orderid",conn,1,1
   If Not RS1.EOF Then
      Do While Not RS1.Eof 
	  Response.Write "<option value=""" & RS1("TypeID") & """>" & RS1("TypeName") & "</option>"
	  RS1.MoveNext
	  Loop
   End If
   RS1.Close:Set RS1=Nothing
   %>
   </select>
   关键字:<input type="text" name="KeyWord" class="textbox" value="关键字" size="20"/>
   <anchor>搜索<go href="User_Blog.asp?Status=<%=Status%>&amp;<%=KS.WapValue%>" method="post">
   <postfield name="Flag" value="$(Flag)"/>
   <postfield name="TypeID" value="$(TypeID)"/>
   <postfield name="KeyWord" value="$(KeyWord)"/>
   </go></anchor>
   <br/>
   <%
   Do While Not RS.Eof
      Response.write "---------<br/>"
	  If Status="2" or Status="1" Then
      Response.write "<a href='User_Blog.asp?action=rizhichakan&amp;id="&rs("ID")&"&amp;" & KS.WapValue & "'>"&KS.GotTopic(trim(RS("title")),35)&"</a>"
	  Else
      Response.write "<a href='../Space/List.asp?id="&rs("id")&"&amp;UserName="&KSUser.UserName&"&amp;" & KS.WapValue & "'>"&KS.GotTopic(trim(RS("title")),35)&"</a>"
	  End If
	  Select Case rs("Status")
	      Case 0
		  Response.Write "正常<br/>"
		  Case 1
		  Response.Write "草稿<br/>"
		  Case 2
		  Response.Write "未审<br/>"
	  end select
	  Response.write "分类:"&Conn.Execute("Select TypeName From KS_BlogType Where TypeID=" & RS("TypeID"))(0)&" "&formatdatetime(rs("AddDate"),2)&"<br/>"
	  Response.Write "<a href='User_Blog.asp?action=rizhixuxie&amp;id="&rs("ID")&"&amp;" & KS.WapValue & "'>续写</a> "
	  Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&rs("ID")&"&amp;" & KS.WapValue & "'>修改</a> "
	  Response.Write "<a href='User_Blog.asp?action=ArticleDel&amp;id="&rs("ID")&"&amp;" & KS.WapValue & "'>删除</a><br/>"
      RS.MoveNext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   Loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Blog.asp", True, "篇日志", CurrentPage, "Status=" & Status &"&amp;" & KS.WapValue & "")


	

end if
rs.Close:Set rs=Nothing
response.Write "---------<br/>"
End Sub

'日志查看=================================================================================
Sub rizhichakan()
Set rs=conn.execute("Select top 1 * From KS_BlogInfo Where ID="&ID&"")
If Not rs.Eof Then
Dim KSB:Set KSB=New BlogCls
%>
标题:<%=rs("Title")%><br/>
心情:<img src="../User/images/face/<%=rs("Face")%>.gif" />
天气:<%Call KSB.GetWeather(rs)%><br/>
内容:<%=KS.ContentPage("User_Blog.asp?action=rizhichakan&amp;id="&rs("ID")&"&amp;" & KS.WapValue & "",KS.LoseHtml(rs("content")),80,False)%><br/>
时间:<%=rs("AddDate")%><br/>
<a href='User_Blog.asp?action=rizhixuxie&amp;id=<%=rs("ID")%>&amp;<%=KS.WapValue%>'>续写</a> 
<a href='User_Blog.asp?action=Edit&amp;id=<%=rs("ID")%>&amp;<%=KS.WapValue%>'>修改</a>
<a href='User_Blog.asp?action=ArticleDel&amp;id=<%=rs("ID")%>&amp;<%=KS.WapValue%>'>删除</a><br/>
<a href='User_Blog.asp?action=&amp;<%=KS.WapValue%>'>日志列表</a><br/>
<%
Set KSB=Nothing
Else
   Response.write "非法参数!<br/>"
End If
rs.close:Set rs=Nothing
End Sub

'删除日志=================================================================================
Sub ArticleDel()
    Conn.Execute("Delete From KS_BlogInfo Where Status<>1 And ID="&ID&"")
	Response.write "日志删除成功。<br/>"
	Response.Write "<a href='User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"
End Sub

'日志续写=================================================================================
Sub rizhixuxie()
Set rs=conn.execute("Select Content From KS_BlogInfo Where ID="&ID&"")
If Not rs.Eof Then
%>
=日志续写=<br/>
尾部内容:<%=Right(Rs("Content"),20)%><br/>
追加内容:<input name="Content<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value=""/><br/>
<anchor>确定<go href="User_Blog.asp?Action=xuxiebaocun&amp;id=<%=id%>&amp;<%=KS.WapValue%>" method="post">
<postfield name="Content" value="$(Content<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>
<% 
Else
   Response.write "非法参数!<br/>"
End If
rs.close:Set rs=Nothing
End Sub
'续写保存=================================================================================
Sub xuxiebaocun()
Set rs=conn.execute("Select Content From KS_BlogInfo Where ID="&ID&"")
If rs.Eof Then
   Response.write "非法参数!<br/>"
Else
   Content=KS.GotTopic(KS.S("Content"),500)
   If Content="" Then
      Response.write "出错提示，你没有输入续写内容！<br/>"
      Response.Write "<a href='User_Blog.asp?action=rizhixuxie&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
   Else
      Set RSObj=Server.CreateObject("Adodb.Recordset")
	  RSObj.Open "Select Content From KS_BlogInfo Where ID="&ID,Conn,1,3
	  RSObj("Content")=rs("Content")&Content
	  RSObj.Update:RSObj.Close:Set RSObj=Nothing
	  Response.write "续写成功。<br/>"
	  Response.Write "<a href='User_Blog.asp?action=rizhichakan&amp;ID="&ID&"&amp;" & KS.WapValue & "'>日志查看</a><br/>"
	  Response.Write "<a href='User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"
   End IF
End If
rs.close:Set rs=Nothing
End Sub
'添加日志=================================================================================
Sub ArticleAdd()
  		if Action="Edit" Then
		Set rs=Server.CreateObject("ADODB.RECORDSET")
		   rs.Open "Select top 1 * From KS_BlogInfo Where ID="&ID,Conn,1,1
		   If Not rs.Eof Then
		     TypeID=rs("TypeID")
			 ClassID=rs("ClassID")
			 Title=rs("Title")
			 Face=rs("Face")
			 weather=rs("Weather")
			 Content=rs("Content")
			 PassWord=rs("PassWord")
		   End If
		   rs.Close:Set rs=Nothing
		   Action="EditSave"
		   OpStr="确定修改"
		   titl="日志修改"
		Else
		  Action="AddSave"
		  OpStr="确定发布"
		  titl="日志发布"
		End If
		%>
=<%=titl%>=<br/>
<%=KS.GetReadMessage%>

你的心情：<select name="face">
	<option value="1">不错</option>
    <option value="2">茫然</option>
    <option value="3">开心</option>
    <option value="9">激动</option>
	<option value="7">郁闷</option>
	<option value="8">难受</option>
	<option value="11">寂寞</option>
	<option value="10">变态</option>
	<option value="4">其他</option>
</select><br />
当天天气：<select name="Weather">
<option value="sun.gif">晴天</option>
<option value="sun2.gif">和煦</option>
<option value="yin.gif">阴天</option>
<option value="qing.gif">清爽</option>
<option value="yun.gif">多云</option>
<option value="wu.gif">有雾</option>
<option value="xiaoyu.gif">小雨</option>
<option value="yinyu.gif">中雨</option>
<option value="leiyu.gif">雷雨</option>
<option value="caihong.gif">彩虹</option>
<option value="hexu.gif">酷热</option>
<option value="feng.gif">寒冷</option>
<option value="xue.gif">小雪</option>
<option value="daxue.gif">大雪</option>
<option value="moon.gif">月圆</option>
<option value="moon2.gif">月缺</option>
</select><br/>

日志分类：<select name='TypeID'>
<option value="0">请选择日志类别</option>
<%Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select * From KS_BlogType order by orderid",conn,1,1
If Not RS.EOF Then
Do While Not RS.Eof 
Response.Write "<option value=""" & RS("TypeID") & """>" & RS("TypeName") & "</option>"
RS.MoveNext
Loop
End If
RS.Close
Set RS=Nothing
%></select><br/>	
我的专栏：<select name='ClassID'>
<option value="0">选择我的专栏</option>
<%=KSUser.UserClassOption(2,ClassID)%>
</select>			
<br/>				
日志标题：<input name="Title<%=minute(now)%><%=second(now)%>" type="text" maxlength="40" value="<%=KS.LoseHtml(Title)%>"/><br/>
日志内容：<input name="Content<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" value="<%=KS.LoseHtml(Content)%>"/><br/>
日志密码：<input name="Password<%=minute(now)%><%=second(now)%>" type="text" value="<%=PassWord%>" /><br/>
<anchor><%=OpStr%><go href='User_Blog.asp?Action=<%=Action%>&amp;id=<%=id%>&amp;<%=KS.WapValue%>' method='post'>
<postfield name='face' value='$(face)'/>
<postfield name='Weather' value='$(Weather)'/>
<postfield name='TypeID' value='$(TypeID)'/>
<postfield name='ClassID' value='$(ClassID)'/>
<postfield name='Title' value='$(Title<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Content' value='$(Content<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Password' value='$(Password<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>


<anchor>放入草稿箱<go href='User_Blog.asp?Action=<%=Action%>&amp;id=<%=id%>&amp;<%=KS.WapValue%>' method='post'>
<postfield name='Status' value='1'/>
<postfield name='face' value='$(face)'/>
<postfield name='Weather' value='$(Weather)'/>
<postfield name='TypeID' value='$(TypeID)'/>
<postfield name='ClassID' value='$(ClassID)'/>
<postfield name='Title' value='$(Title<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Content' value='$(Content<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Password' value='$(Password<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>

<br/>
<%
response.Write "---------<br/>"
End Sub

'日志修改=================================================================================
Sub EditSave()
    face=KS.S("face")
	Weather=KS.S("Weather")
	TypeID=KS.S("TypeID")
	ClassID=KS.S("ClassID")
	Title=KS.GotTopic(KS.S("Title"),60)
	'AddDate=S("AddDate")
	Content=KS.GotTopic(KS.S("Content"),500)
	Password=KS.S("Password")
	Status=KS.ChkClng(KS.S("Status"))
	
	If TypeID=0 Then
	   Response.write "出错提示，你没有选择日志分类！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	'ElseIF ClassID=0 Then
	'   Response.write "出错提示，你没有选择我的专栏！<br/>"
	'   Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	ElseIF Title="" Then
	   Response.write "出错提示，你没有输入日志标题！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	'ElseIF not isdate(adddate) Then
	   'Response.write "出错提示，你输入的日期不正确！<br/>"
	   'Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	ElseIF Content="" Then
	   Response.write "出错提示，你没有输入日志内容！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Edit&amp;id="&id&"&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	Else
	        Set RSObj=Server.CreateObject("Adodb.Recordset")
			    RSObj.Open "Select top 1 * From KS_BlogInfo Where ID="&ID,Conn,1,3
				RSObj("Title")=Title
				RSObj("TypeID")=TypeID
				RSObj("ClassID")=ClassID
				RSObj("UserName")=KSUser.UserName
				RSObj("Face")=Face
				RSObj("Weather")=weather
				RSObj("Adddate")=now()
				RSObj("Content")=Content
				RSObj("Password")=Password
				  if status=1 then
				  RSObj("Status")=1
				  elseif KS.SSetting(3)=1 Then
				  RSObj("Status")=2
				  Else
				  RSObj("Status")=0
				  end if
				RSObj.Update
				RSObj.Close:Set RSObj=Nothing
			Response.write "操作成功。<br/>"
			Response.Write "<a href='User_Blog.asp?action=rizhichakan&amp;ID="&ID&"&amp;" & KS.WapValue & "'>日志查看</a><br/>"
			Response.Write "<a href='User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"
    End IF
End Sub

'发布日志=================================================================================
Sub AddSave()
    face=KS.S("face")
	Weather=KS.S("Weather")
	TypeID=KS.S("TypeID")
	ClassID=KS.S("ClassID")
	Title=KS.S("Title")
	Tags=Trim(KS.S("Tags"))
	'AddDate=S("AddDate")
	Content=KS.S("Content")
	Password=KS.S("Password")
	Status=KS.ChkClng(KS.S("Status"))
	
	If TypeID=0 Then
	   Response.write "出错提示，你没有选择日志分类！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Add&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	'ElseIF ClassID=0 Then
	'   Response.write "出错提示，你没有选择我的专栏！<br/>"
	 '  Response.Write "<a href='User_Blog.asp?action=Add&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	ElseIF Title="" Then
	   Response.write "出错提示，你没有输入日志标题！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Add&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	'ElseIF not isdate(adddate) Then
	   'Response.write "出错提示，你输入的日期不正确！<br/>"
	   'Response.Write "<a href='User_Blog.asp?action=Add&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	ElseIF Content="" Then
	   Response.write "出错提示，你没有输入日志内容！<br/>"
	   Response.Write "<a href='User_Blog.asp?action=Add&amp;" & KS.WapValue & "'>返回重写</a><br/>"
	Else
            Set RSObj=Server.CreateObject("Adodb.Recordset")
			    RSObj.Open "Select top 1 * From KS_BlogInfo",Conn,1,3
				RSObj.AddNew
				RSObj("Title")=Title
				RSObj("TypeID")=TypeID
				RSObj("ClassID")=ClassID
				RSObj("Tags")=Tags
				RSObj("UserName")=KSUser.UserName
				RSObj("Face")=Face
				RSObj("Weather")=weather
				RSObj("Adddate")=now()
				RSObj("Content")=Content
				RSObj("Password")=Password
				  if status=1 then
				  RSObj("Status")=1
				  elseif KS.SSetting(3)=1 Then
				  RSObj("Status")=2
				  Else
				  RSObj("Status")=0
				  end if
				RSObj("Hits")=0
				ID=RSObj("ID")
				RSObj.Update
				RSObj.Close:Set RSObj=Nothing
			Response.write "操作成功。<br/>"
			Response.Write "<a href='User_Blog.asp?action=rizhichakan&amp;ID="&ID&"&amp;" & KS.WapValue & "'>日志查看</a><br/>"
			Response.Write "<a href='User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"	
	End If
End Sub

'评论管理=================================================================================
Sub Comment()
%>
=评论管理=<br/>
<%
Response.write KS.GetReadMessage

If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If

set rs=server.createobject("adodb.recordset")
Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
Dim Sql:sql = "select * from KS_BlogComment "& Param &" order by AddDate DESC" 
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "没有用户给你评论!<br/>"
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
   
   do while not rs.eof
   %>

      发表人:<%=RS("AnounName")%><br/>
      时间:<%=formatdatetime(rs("AddDate"),2)%><br/>
      评论标题:<%=KS.GotTopic(trim(RS("title")),35)%>
	  <%
	  if Not IsNull(RS("Replay")) or rs("replay")<>"" Then
	  response.write "(已回复)"
	  end if
	  %><br/>
	  <%if Not IsNull(RS("Replay")) or rs("replay")<>"" Then%>
      <a href="User_Blog.asp?id=<%=rs("id")%>&amp;Action=ReplayComment&amp;<%=KS.WapValue%>">修改回复</a>
	  <%else%>
      <a href="User_Blog.asp?id=<%=rs("id")%>&amp;Action=ReplayComment&amp;<%=KS.WapValue%>">回复评论</a>
	  <%End If%>
      <a href="User_Blog.asp?action=CommentDel&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">删除评论</a><br/>
      ----------<br/>
   <%
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Blog.asp", True, "条评论", CurrentPage, "Action=Comment&amp;" & KS.WapValue & "")
   
end if

rs.close

End Sub

'回复评论=================================================================================
Sub ReplayComment()
   Set rs=Server.CreateObject("ADODB.RECORDSET")
   rs.Open "Select * From KS_BlogComment where id="&id,Conn,1,1
   If rs.Eof And rs.Bof Then
      Response.Write "参数出错!<br/>"
	  Response.end
   End If
   Title=rs("Title")
   Content=rs("Content")
   Replay=rs("Replay")
   If IsNull(Replay) Then Replay=""
   rs.Close:Set rs=Nothing	
%>
=回复评论=<br/>
评论标题：<%=Title%><br/>
评论内容：<%=KS.LoseHtml(Content)%><br/>
<%if Not IsNull(Replay) or replay<>"" Then%>
回复内容：<%=KS.LoseHtml(Replay)%><br/>
回复修改：<input name="Replay<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value="<%=KS.LoseHtml(Replay)%>"/><br/>
<%else%>
回复内容：<input name="Replay<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value=""/><br/>
<%End If%>
<anchor>立即回复<go href="User_Blog.asp?action=SaveCommentReplay&amp;id=<%=id%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
<postfield name="Replay" value="$(Replay)"/>
</go></anchor>
<br/>
<%
response.Write "---------<br/>"
End Sub

'保存评论回复=================================================================================
Sub SaveCommentReplay()
    Replay=KS.S("Replay")
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_BlogComment Where ID="&ID,conn,1,3
    If Not RS.Eof Then
	   RS("Replay")=Replay
	   RS("ReplayDate")=Now
	   RS.Update
	End If
	Response.write "评论成功回复。<br/>"
	Response.Write "<a href='User_Blog.asp?action=Comment&amp;" & KS.WapValue & "'>评论管理</a><br/>"
	RS.Close:Set RS=Nothing
End Sub

'删除评论=================================================================================
Sub CommentDel()
        Conn.Execute("Delete From KS_BlogComment Where ID In("&ID&")")
		Response.write "评论删除成功。<br/>"
		Response.Write "<a href='User_Blog.asp?action=Comment&amp;" & KS.WapValue & "'>评论管理</a><br/>"
End Sub

'留言管理=================================================================================
Sub Message()
%>
=留言管理=<br/>
<%
Response.write KS.GetReadMessage

If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If
Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
Dim Sql:sql = "select * from KS_BlogMessage "& Param &" order by AddDate DESC" 
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "没有用户给你留言!<br/>"
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
   
   do while not rs.eof
   %>
   发表人:<%=RS("AnounName")%><br/>
   时间:<%=rs("AddDate")%><br/>
   标题:<%=trim(rs("title"))%>
   <%
   if Not IsNull(RS("Replay")) or rs("replay")<>"" Then
   response.write "(已回复)"
   end if
   %><br/>
   <%if Not IsNull(RS("Replay")) or rs("replay")<>"" Then%>
   <a href="User_Blog.asp?id=<%=rs("id")%>&amp;Action=ReplayMessage&amp;<%=KS.WapValue%>">修改回复</a>
   <%else%>
   <a href="User_Blog.asp?id=<%=rs("id")%>&amp;Action=ReplayMessage&amp;<%=KS.WapValue%>">回复留言</a>
   <%End If%>
   <a href="User_Blog.asp?action=MessageDel&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>">删除留言</a><br/>
   ----------<br/>
   <%
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Blog.asp", True, "条留言", CurrentPage, "Action=Message&amp;" & KS.WapValue & "")
end if
rs.close

End Sub

'回复留言=================================================================================
Sub ReplayMessage()
    Set rs=Server.CreateObject("ADODB.RECORDSET")
	rs.Open "Select * From KS_BlogMessage where id="&id,Conn,1,1
	If rs.Eof And rs.Bof Then
	   Response.Write "参数出错!<br/>"
	   Response.end
	End If
	Title=rs("Title")
	Content=rs("Content")
	Replay=rs("Replay")
	If IsNull(Replay) Then Replay=""
	rs.Close:Set rs=Nothing
%>
=回复留言=<br/>
留言标题：<%=Title%><br/>
留言内容：<%=KS.LoseHtml(Content)%><br/>
<%if Not IsNull(Replay) or replay<>"" Then%>
回复内容：<%=KS.LoseHtml(Replay)%><br/>
回复修改：<input name="Replay<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value="<%=KS.LoseHtml(Replay)%>"/><br/>
<%else%>
回复内容：<input name="Replay<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value=""/><br/>
<%End If%>
<anchor>立即回复<go href="User_Blog.asp?action=SaveMessageReplay&amp;id=<%=id%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
<postfield name="Replay" value="$(Replay<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>
<%
End Sub

'保存留言回复=================================================================================
Sub SaveMessageReplay()
    Replay=KS.S("Replay")
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_BlogMessage where id="&id,Conn,1,3
	If Not RS.Eof Then
	   RS("Replay")=Replay
	   RS("ReplayDate")=Now
	   RS.Update
	End If
	Response.write "留言成功回复。<br/>"
	Response.Write "<a href=""User_Blog.asp?action=Message&amp;" & KS.WapValue & """>留言管理</a><br/>"
	RS.Close:Set RS=Nothing
End Sub

'删除留言=================================================================================
Sub MessageDel()
    Conn.Execute("Delete From KS_BlogMessage where id In("&id&")")
	Response.Write "留言删除成功。<br/>"
	Response.Write "<a href='User_Blog.asp?action=Message&amp;" & KS.WapValue & "'>留言管理</a><br/>"
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

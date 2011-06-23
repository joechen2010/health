<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：转载文章到个人日记
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>﻿
<card id="main" title="文章转载到个人日记">
<p>
<%
Dim KSCls
Set KSCls = New User_Blog
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%

Class User_Blog
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
            Call CloseConn
            Set KSUser=Nothing
            Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    Dim ID,ChannelID,Action
			Dim Title,ArticleContent
		    ID=KS.S("ID")
			ChannelID=KS.S("ChannelID")
			Action=KS.S("Action")
			
			IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.write "对不起，您还没有注册或登录！<br/>"
			ElseIf KS.SSetting(0)=0 Then
			   Response.write "对不起，本站关闭个人空间功能！<br/>"
			ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)=0 Then
			   Response.write "对不起，你还没申请个人空间,<a href=""../User/User_Blog.asp?Action=BlogEdit&amp;" & KS.WapValue & """>现在开通</a>吧！<br/>"
			ElseIf Conn.Execute("Select status From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)<>1 Then
			   Response.write "对不起，你的空间还没有通过审核或被锁定！<br/>"
			ElseIf Conn.Execute("Select Count(ClassID) From KS_UserClass Where UserName='"&KSUser.UserName&"'")(0)=0 Then
			   Response.write "对不起，你没有添加专栏目,<a href=""../User/User_Class.asp?" & KS.WapValue & """>现在添加</a>吧！<br/>"
			Else
			   Title=Conn.Execute("select Title from "&KS.C_S(ChannelID,2)&" where ID="&ID&"")(0)
			   ArticleContent=Conn.Execute("select ArticleContent from "&KS.C_S(ChannelID,2)&" where ID="&ID&"")(0)
			   Select Case Action
			       Case "AddSave"
				   Call AddSave()
				   Case Else
				   Call ArticleAdd()
			   End Select
			End If
			Response.write "---------<br/>"
			Response.write "<a href=""../Show.asp?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">返回"&KS.C_S(ChannelID,3)&"页</a><br/><br/>"
            Response.write "<a href="""&KS.GetGoBackIndex&""">网站首页</a><br/>"
	    End Sub
		
		Sub ArticleAdd()
		%>
将:<%=Title%>转载到个人日记<br/>
心情：<select name="face" value="1">
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
天气：<Select Name="Weather" value="1">
     <Option value="sun.gif">晴天</Option>
     <Option value="sun2.gif">和煦</Option>
     <Option value="yin.gif">阴天</Option>
     <Option value="qing.gif">清爽</Option>
     <Option value="yun.gif">多云</Option>
     <Option value="wu.gif">有雾</Option>
     <Option value="xiaoyu.gif">小雨</Option>
     <Option value="yinyu.gif">中雨</Option>
     <Option value="leiyu.gif">雷雨</Option>
     <Option value="caihong.gif">彩虹</Option>
     <Option value="hexu.gif">酷热</Option>
     <Option value="feng.gif">寒冷</Option>
     <Option value="xue.gif">小雪</Option>
     <Option value="daxue.gif">大雪</Option>
     <Option value="moon.gif">月圆</Option>
     <Option value="moon2.gif">月缺</Option>
     </Select><br/>
分类：<select name='TypeID'>
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
专栏：<select name='ClassID'>
<option value="0">请选择我的专栏</option>
<%
Set RS=Server.CreateObject("ADODB.RECORDSET")
    RS.Open "select * from KS_UserClass Where UserName='"&KSUser.UserName&"'",conn,1,1
	If Not RS.EOF Then
	    Do While Not RS.Eof 
		Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
		RS.MoveNext
		Loop
	End If
RS.Close:Set RS=Nothing
%></select><br/>
密码：<input name="Password<%=minute(now)%><%=second(now)%>" type="text" value="<%=PassWord%>" /><br/>
<anchor>确定转载<go href='TurnToCarry.asp?Action=AddSave&amp;ID=<%=ID%>&amp;ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>' method='post'>
<postfield name='Status' value='2'/>
<postfield name='face' value='$(face)'/>
<postfield name='Weather' value='$(Weather)'/>
<postfield name='TypeID' value='$(TypeID)'/>
<postfield name='ClassID' value='$(ClassID)'/>
<postfield name='Password' value='$(Password<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>
<br/>
<%
End Sub

		Sub AddSave()
		    Dim face,Weather,TypeID,ClassID,Title,Content,Password,Status
		    face=KS.S("face")
			Weather=KS.S("Weather")
			TypeID=KS.S("TypeID")
			ClassID=KS.S("ClassID")
			Title=KS.LoseHtml(Title)
			Content=KS.LoseHtml(ArticleContent)
			Password=KS.S("Password")
			Status=KS.S("Status")
			If TypeID=0 Then
			   Response.write "出错提示，你没有选择日志分类！<br/>"
			   Response.Write "<anchor>返回重写<prev/></anchor><br/>"
			ElseIF ClassID=0 Then
			   Response.write "出错提示，你没有选择我的专栏！<br/>"
			   Response.Write "<anchor>返回重写<prev/></anchor><br/>"
			ElseIF Title="" Then
			   Response.write "出错提示，你日志标题为空！<br/>"
			   Response.Write "<anchor>返回重写<prev/></anchor><br/>"
			ElseIF Content="" Then
			   Response.write "出错提示，你日志内容为空！<br/>"
			   Response.Write "<anchor>返回重写<prev/></anchor><br/>"
			Else
			   Set RSObj=Server.CreateObject("Adodb.Recordset")
			   RSObj.Open "Select * From KS_BlogInfo",Conn,1,3
			   RSObj.AddNew
			   RSObj("Title")=Title
			   RSObj("TypeID")=TypeID
			   RSObj("ClassID")=ClassID
			   RSObj("UserName")=KSUser.UserName
			   RSObj("Face")=Face
			   RSObj("Weather")=weather
			   RSObj("Adddate")=now()
			   RSObj("Content")=Content
			   RSObj("Password")=Password
			   If status=1 Then
				  RSObj("Status")=1
			   Elseif KS.SSetting(3)=1 Then
				  RSObj("Status")=2
			   Else
				  RSObj("Status")=0
			   End If
			   RSObj("Hits")=0
			   RSObj.Update
			   RSObj.MoveLast
			   Response.write "转载文章到个人空间日记成功。<br/>"
			   If DataBaseType=0 then
			      Response.Write "<a href='../User/User_Blog.asp?action=rizhichakan&amp;ID="&RSObj("ID")&"&amp;" & KS.WapValue & "'>日志查看</a><br/>"
			   End If
			   Response.Write "<a href='../User/User_Blog.asp?action=&amp;" & KS.WapValue & "'>日志列表</a><br/>"	
			   RSObj.Close:Set RSObj=Nothing
			End If
	    End Sub
End Class
%>


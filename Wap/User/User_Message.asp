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
<card id="main" title="短消息服务">
<p>
<%
Dim KSCls
Set KSCls = New User_Message
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Message
        Private KS
		Private Max_sEnd,Max_sms,Max_Num,DomainStr
        Private Action
        Private RS,SqlStr,Prev
		Private FoundErr,Errmsg
		Private i
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Max_sEnd=KS.Setting(49)	'群发限制人数
		    Max_sms=KS.Setting(48)	'内容最多字符数
		    Max_Num=KS.Setting(47)  '最多允许存放条数
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
			Action=Lcase(Request("Action"))
			
			IF Action<>"read" And Action<>"outread" Then
			%>
            <a href="User_Message.asp?Action=new&amp;<%=KS.WapValue%>">发送短消息</a><br/>
            <a href="User_Message.asp?Action=inbox&amp;<%=KS.WapValue%>">收件箱</a>
            <a href="User_Message.asp?Action=outbox&amp;<%=KS.WapValue%>">发件箱</a>
            <a href="User_Message.asp?Action=issend&amp;<%=KS.WapValue%>">已发送</a>
            <a href="User_Message.asp?Action=recycle&amp;<%=KS.WapValue%>">废件箱</a><br/>
            <%
			End IF
			
			Select Case Action
			    Case "new" : sendMessage'发送消息
				Case "read" : read'阅读消息
				Case "outread" : read
				Case "delet" : delete
				Case "newmsg" : newmsg
				Case "send" : SavEmsg
				Case "fw" : fw
				Case "edit" : Edit
				Case "savedit" : SavEdit
				Case "delinbox" : Delinbox'删除收件
				Case "alldelinbox" : AllDelinbox'清空收件箱
				Case "deloutbox" : Deloutbox'删除草稿
				Case "alldeloutbox" : AllDeloutbox'清空草稿箱
				Case "delissend" : DelIsSend'删除已发送的消息
				Case "alldelissend" : AllDelIsSend'清空已发送的消息
				Case "delrecycle" : Delrecycle'删除垃圾
				Case "alldelrecycle" : AllDelrecycle'清空垃圾箱
				Case Else : MessageMain
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		'发送信息
		Sub sendMessage()
			Dim SendTime,Title,Content
			Dim ToUser:ToUser=KS.S("ToUser")
			If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then
				Set rs=server.createobject("adodb.recordSet")
				SqlStr="Select SendTime,title,content from KS_Message where Incept='"&KSUser.UserName&"' and id="&Clng(KS.S("ID"))
				RS.open SqlStr,Conn,1,1
				If not(RS.EOF And RS.BOF) Then
					SendTime=RS("SendTime")
					Title="RE " & RS("Title")
					Content=RS("Content")
				End If
				RS.Close:Set RS=Nothing
			End If
			%>
            【发送信息】<br/>
            收件人：<input name="ToUser<%=Minute(Now)%><%=Second(Now)%>" title="收件人" type="text" maxlength="10" size="10" value="<%=ToUser%>"/>
            <select name="font">
            <option>选择好友...</option>
            <option onpick="User_Friend.asp?Action=addF&amp;<%=KS.WapValue%>">添加好友...</option>
			<%
			Set RS=server.createobject("adodb.recordSet")
			SqlStr="select friend from KS_Friend where Username='"&KSUser.UserName&"' order by Addtime desc"
			RS.Open SqlStr,Conn,1,1
			Do While not RS.EOF
			   If ToUser="" Then
				  Response.Write "<option onpick=""User_Message.asp?Action=new&amp;ToUser="&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</option>"
			   Else
				  Response.Write "<option onpick=""User_Message.asp?Action=new&amp;ToUser="&ToUser&","&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</option>"
			   End If
			   RS.Movenext
			Loop
			RS.Close:Set RS=Nothing
			%>
            </select>
            <br/>
            标　题：<input name="Title<%=Minute(Now)%><%=Second(Now)%>" title="标题" type="text" maxlength="30" size="30" value="<%=Title%>"/><br/>
            <%
			If KS.S("ID")<>"" Then
			   Content="在"&SendTime&"您来信中写道：<br/>"&Content&"<br/>"
			Else
			   Content=""
			End If
			%>
            内　容：<input name="Message<%=Minute(Now)%><%=Second(Now)%>" title="内容" type="text" maxlength="500" size="30" value="<%=Server.Htmlencode(Content)%>"/><br/>
            <anchor>发送<go href="User_Message.asp?Action=sEnd&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Message" value="$(Message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="发送"/>
            </go></anchor>
            <anchor>保存<go href="User_Message.asp?Action=sEnd&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Message" value="$(Message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="保存"/>
            </go></anchor>
            <br/><br/>
            可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=max_sEnd%></b>个用户<br/>
            标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br/>
            <%
		End Sub
		
		'读取信息
		Sub read()
		    Prev=True
			If KS.S("ID")=0 Then
			   Response.Write "请指定正确的参数。<br/>"
			   Exit Sub
			End If
			Set RS=Server.Createobject("adodb.recordSet")
			If Request("Action")="read" Then
			   Conn.Execute("Update KS_Message Set flag=1 where ID="&Clng(KS.S("id")))
			End If
			SqlStr="Select * from KS_Message where (Incept='"&KSUser.UserName&"' or sEnder='"&KSUser.UserName&"') And ID="&Clng(KS.S("ID"))
			RS.open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "你是不是跑到别人的信箱啦、或者该信息已经被收件人删除。<br/>"
			   Exit Sub
			Else
			%>
               欢迎使用短消息接收，<%=KSUser.UserName%><br/>
               <a href="User_Message.asp?Action=delet&amp;id=<%=RS("ID")%>&amp;<%=KS.WapValue%>">删除</a>
               <a href="User_Message.asp?Action=new&amp;<%=KS.WapValue%>">发送</a>
               <a href="User_Message.asp?Action=new&amp;ToUser=<%=KS.HTMLEncode(RS("sEnder"))%>&amp;id=<%=KS.S("ID")%>&amp;<%=KS.WapValue%>">回复</a>
               <a href="User_Message.asp?Action=fw&amp;id=<%=KS.S("ID")%>&amp;<%=KS.WapValue%>">转发</a>
               <br/>
			   <%
			   If Request("Action")="outread" Then
			      Response.Write "在<b>"&RS("SendTime")&"</b>，您发送此消息给<b>"&KS.HTMLEncode(RS("Incept"))&"</b>！<br/>"
			   Else
			      Response.Write "在<b>"&RS("SendTime")&"</b>，<b>"&KS.HTMLEncode(RS("sEnder"))&"</b>给您发送的消息！<br/>"
			   End If
			   Dim Content
			   Content=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.HTMLCode(RS("Content")))))
			   If InStr(Content, "Shop/Show.asp") <> 0 Then
			      Content= Replace(Content,KS.Setting(2)&KS.Setting(3)&"Shop/Show.asp?",DomainStr&"Show.asp?ChannelID=5&amp;" & KS.WapValue & "&amp;")
			   End If
			   %>
               消息标题：<%=KS.HTMLencode(RS("Title"))%><br/>
			   <%=KS.ContentPagination(Content,"200","User_Message.asp?Action=read&amp;ID="&KS.S("ID")&"&amp;" & KS.WapValue & "",False,False)%><br/>
			   <%
			   RS.Close:Set RS=Nothing
			   SqlStr="Select id,sEnder from KS_Message where Incept='"&KSUser.UserName&"' and flag=0 and IsSend=1 and id>"&KS.ChkClng(KS.S("ID")&" order by SendTime")
			   Set RS=Conn.Execute(SqlStr)
			   If not (RS.EOF And RS.BOF) Then
			      Response.Write "<a href=""User_Message.asp?Action=read&amp;id="&RS(0)&"&amp;sEnder="&RS(1)&"&amp;" & KS.WapValue & """>[读取下一条信息]</a><br/><br/>"
			   End If
			   RS.Close:Set RS=Nothing
			End If
		End Sub
		
		'转发信息
		Sub fw()
			Dim Title,Content,sEnder
			Dim ToUser:ToUser=KS.S("ToUser")
			If KS.S("ID")<>"" And isNumeric(KS.S("ID")) Then
			   Set RS=Server.Createobject("adodb.recordSet")
			   SqlStr="Select title,content,sEnder from KS_Message where (Incept='"&KSUser.UserName&"' or sEnder='"&KSUser.UserName&"') and id="&Clng(KS.S("ID"))
			   RS.Open SqlStr,Conn,1,1
			   If RS.EOF And RS.BOF Then
			      RS.Close:Set RS=Nothing
				  Response.Write "请指定正确的参数。<br/>"
			      Prev=True
			      Exit Sub
			   Else
				  Title=RS("Title"):Content=RS("Content"):sEnder=RS("sEnder")
			   End If
			   RS.Close:Set RS=Nothing
			End If
			%>
            【转发信息】<br/>
            收件人：<input name="ToUser<%=Minute(Now)%><%=Second(Now)%>" type="text" size="10" value="<%=ToUser%>"/>
            <select value="0">
            <option>选择好友...</option>
            <option onpick="User_Friend.asp?Action=addF&amp;<%=KS.WapValue%>">添加好友...</option>
			<%
			Set RS=server.createobject("adodb.recordSet")
			SqlStr="Select friend from KS_Friend where Username='"&KSUser.UserName&"' order by Addtime desc"
			RS.Open SqlStr,Conn,1,1
			Do While not RS.eof
			   If ToUser="" Then
				  Response.Write "<option onpick=""User_Message.asp?Action=fw&amp;ID="&KS.S("ID")&"&amp;ToUser="&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</option>"
			   Else
				  Response.Write "<option onpick=""User_Message.asp?Action=fw&amp;ID="&KS.S("ID")&"&amp;ToUser="&ToUser&","&RS(0)&"&amp;"&KS.WapValue&""">"&RS(0)&"</option>"
			   End If
			   RS.Movenext
			Loop
			RS.Close:Set RS=Nothing
			%>
            </select><br/>
            标　题：<input type="text" name="Title<%=Minute(Now)%><%=Second(Now)%>"  maxlength="90" value="Fw：<%=Title%>"/><br/>
            <%
			Content="下面是转发信息<br/> 原发件人："&sEnder&"<br/>"&Content&""
			%>
            内　容：<input type="text" name="Message<%=Minute(Now)%><%=Second(Now)%>" maxlength="300" value="<%=Server.Htmlencode(Content)%>"/><br/>
            <anchor>发送<go href="User_Message.asp?Action=sEnd&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Message" value="$(Message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="发送"/>
            </go></anchor>
            <anchor>保存<go href="User_Message.asp?Action=sEnd&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Message" value="$(Message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="保存"/>
            </go></anchor>
            <br/><br/>
            可以用英文状态下的逗号将用户名隔开实现群发，最多<b><%=max_sEnd%></b>个用户<br/>
            标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br/>
		<%
		End Sub
		
		Sub savemsg()
			Dim Incept,title,message,Subtype,i,sUname
			If KS.S("ToUser")="" Then
			   Response.Write "您忘记填写发送对象了吧。<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Incept=KS.S("ToUser")
			   Incept=split(Incept,",")
			End If
			If KS.S("Title")="" Then
			   Response.Write "您还没有填写标题呀。<br/>"
			   Prev=True
			   Exit Sub
			ElseIf KS.strLength(KS.S("Title"))>50 Then
			   Response.Write "标题限定最多50个字符。<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Title=KS.S("Title")
			End If
			If KS.S("Message")="" Then
			   Response.Write "内容是必须要填写的噢。<br/>"
			   Prev=True
			   Exit Sub
			ElseIf KS.strLength(KS.S("Message"))>Cint(max_sms) Then
			   Response.Write "内容限定最多"&max_sms&"个字符。<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Message=KS.S("Message")
			End If
		
			For i=0 To Ubound(Incept)
				sUname=replace(Incept(i),"'","")
				SqlStr="Select UserName from KS_User where UserName='"&sUname&"'"
				Set rs=Conn.Execute(SqlStr)
				If RS.EOF And RS.BOF Then
					RS.Close:Set RS=Nothing
					Response.Write "系统没有这个用户，看看你的发送对象写对了嘛？<br/>"
			        Prev=True
					Exit Sub
				End If
				RS.Close
				RS.Open "select username from ks_friend where username='" & sUname & "' and friend='" & ksuser.username & "' and flag=3",Conn,1,1
				If not rs.eof Then
				   RS.close:Set RS=Nothing
				   Response.Write "对不起，你被" & sUname & "列为黑名单,不能发送短信给他！<br/>"
			       Prev=True
				   Exit Sub
				End If
				RS.Close:Set RS=Nothing
						
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
				If Conn.Execute("select count(*) from KS_Message where Incept='"&sUname&"'")(0)>=Max_Num Then
					Response.Write "由于[" & sUname & "]的信箱已满，发送没有成功！<br/>"
				Else
				   Conn.Execute(SqlStr)
				End If
				If i>Cint(max_sEnd)-1 Then
					Response.Write "最多只能发送给"&max_sEnd&"个用户，您的名单"&max_sEnd&"位以后的请重新发送！<br/>"
					Exit For
				End If
			Next
			'Response.Write "恭喜您，发送短信息成功。发送的消息同时保存在您的"&Subtype&"中。<br/>"
			Response.redirect DomainStr&"User/User_Message.asp?" & KS.WapValue & ""
		End Sub
		
		'更改信息
		Sub Edit()
			dim Incept,Title,Content,ID
			If KS.S("ID")<>"" and isNumeric(KS.S("ID")) Then
				Set rs=server.createobject("adodb.recordSet")
				SqlStr="Select id,Incept,title,content from KS_Message where sEnder='"&KSUser.UserName&"' and IsSend=0 and id="&Clng(KS.S("ID"))
				RS.open SqlStr,Conn,1,1
				If not(RS.eof and RS.bof) Then
				   Incept=rs("Incept"):title=rs("title"):content=rs("content"):id=rs("id")
				Else
				   Response.Write "没有找到您要编辑的信息。<br/>"
				   Prev=True
				   Exit Sub
				End If
				RS.Close:Set RS=Nothing
			Else
			   Response.Write "请指定相关参数。<br/>"
			   Prev=True
			   Exit Sub
			End If
			%>
            【更改信息】<br/>
            请完整输入下列信息<br/>
            收件人：<input name="ToUser<%=Minute(Now)%><%=Second(Now)%>" type="text" size="10" value="<%=Incept%>"/><br/>
            标　题：<input type="text" name="Title<%=Minute(Now)%><%=Second(Now)%>"  maxlength="90" value="<%=Title%>"/><br/>
            内　容：<input type="text" name="Message<%=Minute(Now)%><%=Second(Now)%>" maxlength="300" value="<%=Server.Htmlencode(Content)%>"/><br/>
            <anchor>发送<go href="User_Message.asp?Action=SavEdit&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="title" value="$(title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="message" value="$(message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="发送"/>
            </go></anchor>
            <anchor>保存<go href="User_Message.asp?Action=SavEdit&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8"> 
            <postfield name="ToUser" value="$(ToUser<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Message" value="$(Message<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Submit" value="保存"/>
            </go></anchor>
            <br/><br/>
		    标题最多<b>50</b>个字符，内容最多<b><%=max_sms%></b>个字符<br/>
		<%
		End Sub
		
		Sub SavEdit()
			Dim Incept,title,message,Subtype
			If KS.S("ID")="" or not isNumeric(KS.S("ID")) Then
			   Response.Write "请指定相关参数。<br/>"
			   Prev=True
			   Exit Sub
			End If
			If KS.S("ToUser")="" Then
			   Response.Write "您忘记填写发送对象了吧。<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Incept=KS.S("ToUser")
			End If
			If KS.S("Title")="" Then
			   Response.Write "您还没有填写标题呀!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Title=KS.S("Title")
			End If
			If KS.S("Message")="" Then
			   Response.Write "内容是必须要填写的噢!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Message=KS.S("Message")
			End If
			SqlStr="Select UserName from KS_User where UserName='"&Incept&"'"
			Set RS=Conn.Execute(SqlStr)
			If RS.EOF And RS.BOF Then
			   Set RS=Nothing
			   Response.Write "系统没有这个用户，看看你的发送对象写对了嘛？<br/>"
			   Prev=True
			   Exit Sub
			End If
			Set RS=Nothing
		
			If KS.S("Submit")="发送" Then
			   SqlStr="Update KS_Message Set Incept='"&Incept&"',sEnder='"&KSUser.UserName&"',title='"&Title&"',content='"&Message&"',SendTime="&SqlNowString&",flag=0,IsSend=1 where id="&Clng(KS.S("ID"))
			   Subtype="已发送信息"
			Else
			   SqlStr="Update KS_Message Set Incept='"&Incept&"',sEnder='"&KSUser.UserName&"',title='"&Title&"',content='"&Message&"',SendTime="&SqlNowString&",flag=0,IsSend=0 where id="&Clng(KS.S("ID"))
			   Subtype="发件箱"
			End If
			Set RS=Conn.Execute(SqlStr)
		    Response.Write "恭喜您，发送短信息成功。发送的消息同时保存在您的"&Subtype&"中。<br/>"
		End Sub
		
		'收件置于回收站，参数字段delR，可用于批量及单个删除
		Sub Delinbox()
			Dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
			   Response.Write "请选择相关参数!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' and id in ("&DelID&")")
			   Response.Write "短信息成功转移到您的回收站!<br/>"
			End If
		End Sub
		
		Sub AllDelinbox()
			Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' And delR=0")
			Response.Write "短信息成功转移到您的回收站!<br/>"
		End Sub
		
		'发件逻辑删除，置于回收站，入口字段delS，可用于批量及单个删除
		Sub Deloutbox()
			Dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
			   Response.Write "请选择相关参数!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' And IsSend=0 And id in ("&DelID&")")
			   Response.Write "短信息成功转移到您的回收站!<br/>"
			End If
		End Sub
		
		Sub AllDeloutbox()
			Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' And delS=0 And IsSend=0")
			Response.Write "短信息成功转移到您的回收站!<br/>"
		End Sub
		
		'已发送置于回收站，入口字段delS，可用于批量及单个删除
		'delS：0未操作，1发送者删除，2发送者从回收站删除
		Sub DelIsSend()
			Dim DelID
			DelID=KS.S("ID")
			DelID=KS.FilterIDs(DelID)
			If DelID="" or isnull(DelID) or Not IsNumeric(replace(Replace(DelID,",","")," ","")) Then
			   Response.Write "请选择相关参数!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' And IsSend=1 And id in ("&DelID&")")
			   Response.Write "短信息成功转移到您的回收站!<br/>"
			End If
		End Sub
		
		Sub AllDelIsSend()
			Conn.Execute("Update KS_Message Set delS=1 where Sender='"&KSUser.UserName&"' And delS=0 And IsSend=1")
			Response.Write "短信息成功转移到您的回收站!<br/>"
		End Sub
		
		'用户能完全删除收到信息和逻辑删除所发送信息，逻辑删除所发送信息设置入口字段delS参数为2
		Sub Delrecycle()
			Dim DelID
			DelID=KS.S("ID")
			If KS.S("Checked")="ok" Then
			   DelID=KS.FilterIDs(DelID)
			   If DelID="" or isnull(DelID) or Not IsNumeric(Replace(Replace(DelID,",","")," ","")) Then
			      Response.Write "请选择相关参数!<br/>"
				  Prev=True
				  Exit Sub
			   Else
			      Conn.Execute("delete from KS_Message where Incept='"&KSUser.UserName&"' And id in ("&DelID&")")
				  Conn.Execute("Update KS_Message Set delS=2 where Sender='"&KSUser.UserName&"' And delS=1 And id in ("&DelID&")")          
				  Response.Write "删除短信息成功。<br/>"
			   End If
			Else
			   Response.Write "删除的消息将不可恢复。确定要删除短信息吗？"
			   Response.Write "<a href=""User_Message.asp?Action=Delrecycle&amp;ID="&DelID&"&amp;Checked=ok&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_Message.asp?"&KS.WapValue&""">取消</a><br/>"
			End If 
		End Sub
		
		Sub AllDelrecycle()
		    If KS.S("Checked")="ok" Then
			   Conn.Execute("delete from KS_Message where Incept='"&KSUser.UserName&"' And delR=1")	
			   Conn.Execute("Update KS_Message Set delS=2 where Sender='"&KSUser.UserName&"' And delS=1")
			   Response.Write "删除短信息成功。<br/>"
			Else
			   Response.Write "删除的消息将不可恢复。确定要删除短信息吗？"
			   Response.Write "<a href=""User_Message.asp?Action=AllDelrecycle&amp;Checked=ok&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_Message.asp?"&KS.WapValue&""">取消</a><br/>"
			End if
		End Sub
		
		Sub delete()
			Dim DelID
			DelID=KS.S("id")
			If not isNumeric(DelID) or DelID="" or isnull(DelID) Then
			   Response.Write "请选择相关参数!<br/>"
			   Prev=True
			   Exit Sub
			Else
			   Conn.Execute("Update KS_Message Set delR=1 where Incept='"&KSUser.UserName&"' And id="&Clng(DelID))
			   Conn.Execute("Update KS_Message Set delS=1 where sEnder='"&KSUser.UserName&"' And id="&Clng(DelID))
			   Response.Write "删除短信息成功。删除的消息将置于您的回收站内。<br/>"
			End If
		End Sub
		
		Sub MessageMain()
			Dim SqlStr,boxName,smstype,readaction,turl,DelID
			Dim keyword,param
			keyword=KS.S("KeyWord")
			If keyword<>"" Then
			   If KS.S("searcharea")=1 Then
			      param=" and title like '%" & keyword & "%'"
			   Else
			      param=" and content like '%" & keyword & "%'"
			   End If
			End If
			Dim CurrentPage,MaxPerPage,TotalPut
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Select Case Action
			Case "inbox"
				BoxName="收件箱":smstype="inbox":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where Incept='"&KSUser.UserName&"'" & param & " and IsSend=1 and delR=0 order by flag,SendTime desc"
			Case "outbox"
				BoxName="草稿箱":smstype="outbox":readaction="edit":turl="sms"
				SqlStr="select * from KS_Message where Sender='"&KSUser.UserName&"'" & param & " and IsSend=0 and delS=0 order by SendTime desc"
			Case "issend"
				BoxName="已发送":smstype="IsSend":readaction="outread":turl="readsms"
				SqlStr="select * from KS_Message where Sender='"&KSUser.UserName&"'" & param & " and IsSend=1 and delS=0 order by SendTime desc"
			Case "recycle"
				BoxName="垃圾箱":smstype="recycle":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where ((Sender='"&KSUser.UserName&"'" & param & " and delS=1) or (Incept='"&KSUser.UserName&"' and delR=1)) and not delS=2 order by SendTime desc"
			Case Else
				BoxName="收件箱":smstype="inbox":readaction="read":turl="readsms"
				SqlStr="select * from KS_Message where Incept='"&KSUser.UserName&"'" & param & " and IsSend=1 and delR=0 order by flag,SendTime desc"
			End Select
			Response.Write "【我的" & Boxname & "】<br/>"
			Dim RS:Set RS=server.createobject("adodb.recordset")
			RS.Open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "您的" & Boxname & "中没有任何内容。<br/>"
			Else
			   MaxPerPage =15
			   TotalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (totalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = totalPut \ MaxPerPage
				  Else
				     CurrentPage = totalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Do While not RS.EOF
			      Select Case smstype
				      Case "inbox"
					  If RS("flag")=0 Then
					     Response.Write "<img src=""Images/news.gif"" alt="".""/>"
					  Else
					     Response.Write "<img src=""Images/olds.gif"" alt="".""/>"
					  End If
					  Case "outbox"
					  Response.Write "<img src=""Images/IsSend_2.gif"" alt="".""/>"
					  Case "IsSend"
					  Response.Write "<img src=""Images/IsSend_1.gif"" alt="".""/>"
					  Case "recycle"
					  If RS("flag")=0 Then
					     Response.Write "<img src=""Images/news.gif"" alt="".""/>"
					  Else
					     Response.Write "<img src=""Images/olds.gif"" alt="".""/>"
					  End If
				  End Select
                  Response.Write "<a href=""User_Message.asp?Action="&ReadAction&"&amp;ID="&RS("ID")&"&amp;sender="&RS("sender")&"&amp;"&KS.WapValue&""">"&KS.HTMLEncode(RS("Title"))&KS.DateFormat(RS("SendTime"),37)&"</a>"
                  Response.Write "<br/>"
				  DelID=DelID&RS("ID")&","
			      RS.Movenext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   Call  KS.ShowPageParamter(TotalPut, MaxPerPage, "User_Message.asp", False, "个消息", CurrentPage, "Action="&Action&"&amp;" & KS.WapValue & "")
			End If
			RS.Close:set RS=Nothing
			%>
            <br/>
            <%
			Response.Write ShowTable(Conn.Execute("select Count(*) from KS_Message where Incept='"&KSUser.UserName&"'")(0),Max_Num)
			%>
            <a href="User_Message.asp?Action=Del<%=smstype%>&amp;ID=<%=DelID%>&amp;<%=KS.WapValue%>">删除本页纪录</a><br/>
            <a href="User_Message.asp?Action=AllDel<%=smstype%>&amp;<%=KS.WapValue%>">清空所有纪录</a><br/>
            
            
            搜索:	<select name="Action">
            <option value="inbox">收件箱</option>
            <option value="outbox">发件箱</option>
            <option value="issend">已发送</option>
            <option value="recycle">废件箱</option>
            </select>
            <select name="searcharea">
            <option value="1">短消息主题</option>
            <option value="2">短消息内容</option>
            </select>
            <input type="text" name="keyword" value="关键字"/>
            <anchor>搜索<go href="User_Message.asp?<%=KS.WapValue%>" method="post">
            <postfield name="action" value="$(action)"/>
            <postfield name="searcharea" value="$(searcharea)"/>
            <postfield name="keyword" value="$(keyword)"/>
            </go></anchor><br/>
      		<%
		End Sub

		 '更新数，总数
		Function ShowTable(str,c)
		    Dim Tempstr,TempPercent
			If C = 0 Then C = 99999999
			Tempstr = str/C
			TempPercent = FormatPercent(Tempstr,0,-1)
			ShowTable = "消息容量:"&C&"/"&str&""
			If FormatNumber(Tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable &"已使用:" & TempPercent & "，请及时删除无用信息！<br/>"
			Else
				ShowTable = ShowTable &"<b>已使用:" & TempPercent & ",请赶快清理！</b><br/>"
			End If
		End Function
End Class
%> 

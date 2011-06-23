<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Md5.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="找回密码">
<p>
<%
Dim KSCls
Set KSCls = New Admin_GetPass
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class Admin_GetPass
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    Dim Step:Step=KS.S("Step")
			IF Step="" Then Step=1
			IF Step=1 Then
			   %>
               取回密码第一步,输入用户名<br/>
               您的用户名：<input name="UserName" type="text" maxlength="30" value=""/>
               <anchor>下一步<go href="index.asp?Step=2" method="post" accept-charset="utf-8"><postfield name="UserName" value="$(UserName)"/></go></anchor><br/><br/>
               <a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
			<%
			End IF
			
			IF Step=2 Then
			   Dim RS
			   Dim UserName:UserName=KS.R(KS.S("UserName"))
			   If UserName = "" Then
			      Response.Write "请输入用户名!<br/>"
			   Else
			      Set RS=Server.CreateObject("Adodb.RecordSet")
			      RS.Open "Select Question From KS_User Where UserName='" & UserName & "'",Conn,1,1
				  IF RS.Eof And RS.Bof Then
				     Response.Write "对不起,您输入的用户名不存在！<br/>"
				  Else
				  %>
                  取回密码第二步,回答密码问题<br/>
                  密码问题：<%=RS(0)%><br/>
                  您的答案：<input name="Answer<%=minute(now)%><%=second(now)%>" type="text" size="20" maxlength="30" value=""/><br/>
				  <%
				  Response.Write "验证码：<input name=""verifycode"&minute(now)&second(now)&""" type=""text"" size=""4"" maxlength=""4"" value="""" /><b>" & KS.GetVerifyCode & "</b><br/>"
				  %>
                  <anchor>下一步<go href="Index.asp?Step=3" method="post" accept-charset="utf-8">
                  <postfield name="UserName" value="<%=UserName%>"/>
                  <postfield name="Answer" value="$(Answer<%=minute(now)%><%=second(now)%>)"/>
                  <postfield name="verifycode" value="$(verifycode<%=minute(now)%><%=second(now)%>)"/>
                  </go></anchor>
                  <br/><br/>
                  <a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
				  <%
				  End IF
			   End IF
			End IF
			
			IF Step=3 Then
			   Dim Verifycode:Verifycode=	KS.R(KS.S("Verifycode"))
			   UserName=KS.R(KS.S("UserName"))
			   Dim Answer:Answer=KS.S("Answer")
			   IF Trim(Verifycode)<>Trim(Session("Verifycode")) Then 
			      Response.Write("验证码有误，请重新输入！<br/>")
			   ElseIf Answer="" Then
			      Response.write "请输入问题答案！<br/>"
			   Else
			      Dim RSC:Set RSC=Conn.Execute("Select Answer From KS_User Where UserName='" & UserName & "' and Answer='" & Answer & "'")
				  IF RSC.EOF AND RSC.Bof Then
				     Response.Write "对不起,您输入的答案不正确！<br/>"
				  Else
				  %>
                  取回密码第三步,设置新密码<br/>
                  用户名：<%=UserName%><br/>
                  新密码：<input name="PassWord<%=minute(now)%><%=second(now)%>" type="password" size="20" maxlength="30" value=""/><br/>
                  确认密码：<input name="RePassWord<%=minute(now)%><%=second(now)%>" type="password" size="20" maxlength="30" value=""/><br/>
                  <anchor>完成设置<go href="Index.asp?Step=4" method="post" accept-charset="utf-8">
                  <postfield name="UserName" value="<%=UserName%>"/>
                  <postfield name="Answer" value="<%=Answer%>"/>
                  <postfield name="PassWord" value="$(PassWord<%=minute(now)%><%=second(now)%>)"/>
                  <postfield name="RePassWord" value="$(RePassWord<%=minute(now)%><%=second(now)%>)"/>
                  </go></anchor><br/><br/>
                  <a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
				  <%
				  End If
			   End If
			End IF
			
			IF Step=4 Then
			   UserName=KS.DelSql(Replace(Replace(Request.Form("UserName"), "'", ""), """", ""))
			   Answer=KS.S("answer")
			   Dim PassWord:PassWord=KS.DelSql(Replace(Replace(Request.Form("PassWord"), "'", ""), """", ""))
			   Dim RePassWord:RePassWord=KS.DelSql(Replace(Replace(Request.Form("RePassWord"), "'", ""), """", ""))
			   If UserName="" Then
			      Response.Write "操作非法!<br/>"
			   ElseIF PassWord = "" Then
			      Response.Write "请输入登录密码!<br/>"
			   ElseIF RePassWord="" Then
			      Response.Write "请输入确认密码<br/>"
			   ElseIF PassWord<>RePassWord Then
			      Response.Write "两次输入的密码不一致<br/>"
			   Else
			      Set RS=Server.CreateObject("Adodb.RecordSet")
				  RS.Open "Select PassWord From KS_User Where UserName='" & UserName & "' and answer='" & answer &"'",Conn,1,3
				  If Not RS.Eof Then
			      RS(0)=MD5(PassWord,16)
				  RS.Update
				  %>
                  恭喜你,密码取回成功。<br/>您的新密码是:<%=PassWord%>,请牢记新密码。<br/>
                  <anchor>自动登陆<go href="CheckLogin.asp" method="post">
                  <postfield name="Way" value="1"/>
                  <postfield name="UserName" value="<%=UserName%>"/>
                  <postfield name="PassWord" value="<%=PassWord%>"/>
                  <postfield name="verifycode" value="<%=Session("Verifycode")%>"/>
                  </go></anchor>
                  <br/><br/>
                  <a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
				  <%
				  Else
			         Response.write "操作失败!<br/>"
				  End If
				  RS.Close:Set RS=Nothing
			   End IF
		    End IF
        End Sub
End Class
%>

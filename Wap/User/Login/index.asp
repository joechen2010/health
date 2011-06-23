<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS,KSUser
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    Response.ContentType="text/vnd.wap.wml; charset=utf-8"
			Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
			%>
            <!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 
            <wml>
            <head>
            <meta http-equiv="Cache-Control" content="max-age=0"/>
            <meta http-equiv="Cache-Control" content="no-cache"/>
            </head>
            <card id="card1" title="会员中心-会员登录">  
            <p align="left">
            【会员登录】<br/>
            用户名称:<input name="UserName" maxlength="30" value="" emptyok="false"/><br/>
            登陆密码:<input name="PassWord<%=Minute(Now)%><%=Second(Now)%>" type="password" maxlength="30" value="" emptyok="false"/><br/>
			<%
			If KS.Setting(34)=1 Then
			   Response.write "验证码:<input name=""verifycode"&Minute(Now)&Second(Now)&""" type=""text"" size=""4"" maxlength=""4"" value="""" emptyok=""false"" format=""*N""/>" & KS.GetVerifyCode & "<br/>"
			End If
			Dim ToUrl
			ToUrl = Request.ServerVariableS("QUERY_STRING")
			ToUrl = Replace(Replace(ToUrl,"&amp;","&"),"&","&amp;")
			%>
            <anchor>会员登录<go href="../CheckUserlogin.asp" method="post">
            <postfield name="UserName" value="$(UserName)"/>
            <postfield name="PassWord" value="$(PassWord<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="verifycode" value="$(verifycode<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="ToUrl" value="<%=ToUrl%>"/>
            </go></anchor>
            <br/><br/>
            忘记密码？ 如果你忘记密码请点击<a href="../GetPass/">找回密码</a>。<br/>
            请先注册成为会员<a href="../reg/">注册会员</a><br/>
            ----------<br/>
            <anchor>返回来源页<prev/></anchor><br/>
            </p>
            </card>
            </wml>
			<%
        End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="内容管理">
<p>
<%
Dim KSCls
Set KSCls = New User_MyModelMenuCls
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_MyModelMenuCls
        Private KS,DomainStr
		Private Sub Class_Initialize()
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
			Response.Write "【内容管理】<br/>"
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "Select ChannelID,BasicType,ItemName From KS_Channel Where ChannelStatus=1 And usertf>0",Conn,1,1
			If RS.Eof Then RS.Close:Set RS=Nothing
			Dim K,SQL:SQL=RS.GetRows(-1)
			RS.Close:Set RS=Nothing
			For K=0 To Ubound(SQL,2)
			 select Case SQL(1,K)
			  Case 1%>
              <a href="User_MyArticle.asp?ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>"><%=SQL(2,K)%>管理</a>
              <a href="User_MyArticle.asp?action=Add&amp;ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">发布</a><br/>
			<%Case 2%>
              <a href="User_MyPhoto.asp?ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>"><%=SQL(2,K)%>管理</a>
              <a href="User_MyPhoto.asp?action=Add&amp;ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">发布</a><br/>
			<%Case 3%>
			  <a href="User_MySoftWare.asp?ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>"><%=SQL(2,K)%>管理</a>
              <a href="User_MySoftWare.asp?action=Add&amp;ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">发布</a><br/>
			<%Case 7%>
			  <a href="User_MyMovie.asp?ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>"><%=SQL(2,K)%>管理</a>
              <a href="User_MyMovie.asp?action=Add&amp;ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">发布</a><br/>
			<%Case 8%>
			  <a href="User_MySupply.asp?ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>"><%=SQL(2,K)%>管理</a>
              <a href="User_MySupply.asp?action=Add&amp;ChannelID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">发布</a><br/>
			<%
			 End Select
			Next

			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
End Class
%>
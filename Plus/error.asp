<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>操作提示消息！</title>
<style>
BODY {
	text-align:center; SCROLLBAR-FACE-COLOR: #f6f6f6; FONT-SIZE: 9pt; MARGIN: 0px; 
}
A {
	TEXT-DECORATION: none
}
A:link {
	COLOR: #000000
}
A:visited {
	COLOR: #333333
}
A:hover {
	COLOR: #ae0927
}
A:active {
	COLOR: #0000ff
}

TD {
	FONT-WEIGHT: normal; FONT-SIZE: 9pt; LINE-HEIGHT: 150%; FONT-FAMILY: "宋体"
}
</style>
</head>
<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<p>&nbsp;</p>
<p>&nbsp;</p>
<%
Dim KS:Set KS=New PublicCls
Dim action,Message
action = KS.R(KS.S("action"))
Message = KS.CheckXSS(KS.S("message"))
Select Case action
        Case "error"
                Call Error_Msg()
        Case "succeed"
                Call Succeed_Msg()
        Case Else
                Call Error_Msg()
End Select
Set KS=Nothing
Sub Error_Msg()
        Response.Write "<br><br><br><table width=""523"" style=""border:1px solid #cccccc""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td background=""../images/bottombg.gif"" width=""523"" height=""30"">&nbsp;<b>操作提示：</b></td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td height=""160"" background=""images/img_r2_c2.gif""><table width=""92%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        Response.Write "      <tr>"& vbCrLf
        Response.Write "        <td width=""22%"" align=""center""><img src=""../images/error.gif"" width=""95"" height=""97""></td>"& vbCrLf
        Response.Write "        <td width=""78%""><b>　　产生错误的可能原因：</b><br>" & Message &"<br><font color=red>时间：" & Now() & "</font></td>"& vbCrLf
        Response.Write "      </tr>"& vbCrLf
        Response.Write "    </table></td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td align=""center"" height=""30"" background=""../images/bottombg.gif""><a href=""" & Request.ServerVariables("HTTP_REFERER") & """> << 确定返回</a>&nbsp;</td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "</table><p>&nbsp;</p>"& vbCrLf
End Sub

'********成功提示信息****************
Sub Succeed_Msg()
        Response.Write "<br><br><br><table width=""523"" style=""border:1px solid #cccccc""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td background=""../images/bottombg.gif"" width=""523"" height=""30"">&nbsp;<b>操作提示：</b></td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td height=""160"" background=""images/img_r2_c2.gif""><table width=""92%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"& vbCrLf
        Response.Write "      <tr>"& vbCrLf
        Response.Write "        <td width=""22%"" align=""center""><img src=""../images/succeed.gif"" width=""95"" height=""97""></td>"& vbCrLf
        Response.Write "        <td width=""78%""><b>　　成功提示信息：</b><br>" & Message &"<br><font color=red>时间：" & Now() & "</font></td>"& vbCrLf
        Response.Write "      </tr>"& vbCrLf
        Response.Write "    </table></td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "  <tr>"& vbCrLf
        Response.Write "    <td align=""center"" height=""30"" background=""../images/bottombg.gif""><a href=""" & Request.ServerVariables("HTTP_REFERER") & """> << 确定返回</a>&nbsp;</td>"& vbCrLf
        Response.Write "  </tr>"& vbCrLf
        Response.Write "</table><p>&nbsp;</p>"& vbCrLf
End Sub

%>
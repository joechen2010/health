<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS
Set KS=New PublicCls
Dim ChannelID,ID,RS,ArticleUrl,WebName,WebUrl
Dim ReturnInfo,Subject,MyName,MyMail,FrName,FrMail,MailBody
ChannelID=KS.ChkClng(KS.S("m"))
ID=KS.ChkClng(KS.S("ID"))
ArticleUrl=Request.ServerVariables("HTTP_REFERER")
if ID=0 or ChannelID=0 then
	Response.Write"<script>alert(""����Ĳ�����"");location.href=""javascript:history.back()"";</script>"
    Response.End
end if
Set RS=Server.CreateObject("Adodb.Recordset")
RS.Open "Select * From " & KS.C_S(ChannelID,2) & " Where ID=" & ID,conn,1,1
IF RS.EOF And RS.BoF Then
  RS.CLOSE
 SET RS=NOthing
 Call CloseConn()
 Set KS=Nothing
	Response.Write"<script>alert(""����Ĳ�����"");location.href=""javascript:history.back()"";</script>"
    Response.End
End if
WebName=KS.Setting(0)
WebUrl=KS.Setting(2)
MailServerAddress=KS.Setting(12)
IF KS.S("Action")="Send" Then
FrName=KS.S("FrName")
MyName=KS.S("MyName")
FrMail=KS.S("FrMail")
IF FrMail="" Then
		Response.Write"<script>alert(""���������ַ����Ϊ�գ�"");location.href=""javascript:history.back()"";</script>"
        Response.End
End IF
IF KS.IsValidEmail(FrMail)=false then
		Response.Write"<script>alert(""���������ַ��ʽ����" & FrMail & """);location.href=""javascript:history.back()"";</script>"
        Response.End
End if
MyMail=KS.S("MyMail")
IF MyMail="" Then
		Response.Write"<script>alert(""���������ַ����Ϊ�գ�"");location.href=""javascript:history.back()"";</script>"
        Response.End
End IF
if KS.IsValidEmail(MyMail)=false then
		Response.Write"<script>alert(""���������ַ��ʽ����"");location.href=""javascript:history.back()"";</script>"
        Response.End
End if
Content=KS.S("Content")


Subject="����" & KS.S("FrName") & ",��������"&KS.S("MyName")&"��" & KS.S("SiteName") & "����������һƪ��Ϣ����"
	MailBody=MailBody &"<style>A:visited {	TEXT-DECORATION: none	}"
	MailBody=MailBody &"A:active  {	TEXT-DECORATION: none	}"
	MailBody=MailBody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	MailBody=MailBody &"A:link 	  {	text-decoration: none;}"
	MailBody=MailBody &"A:visited {	text-decoration: none;}"
	MailBody=MailBody &"A:active  {	TEXT-DECORATION: none;}"
	MailBody=MailBody &"A:hover   {	TEXT-DECORATION: underline overline}"
	MailBody=MailBody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	MailBody=MailBody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"

	MailBody=MailBody &"<table border='0' width='90%' align='center'><tr>"
	MailBody=MailBody &"<td valign='middle' align='top'>"
    MailBody=MailBody &Content & "<br>��������Ϣ����<br>" & RS("ArticleContent") 
	MailBody=MailBody &"</td></tr></table>"

'��ʼ����
ReturnInfo=KS.SendMail(MailServerAddress,KS.Setting(13), KS.Setting(14),Subject,FrMail,KS.S("MyName"),MailBody,MyMail)
  IF ReturnInfo="OK" Then
    Response.Write ("<script>alert('�ż��ɹ�����!');window.close();</script>")
	 Response.End
  Else
    Response.Write ("<script>alert('�ż�����ʧ��!ʧ��ԭ��:\n" & ReturnInfo & "');window.close();</script>")
	Response.End
  End if
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>���͵����ʼ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link href="/images/style.css" rel="stylesheet">
<body>
<table width="770" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<form action="?action=Send" name="myform" method="post">
  <tr>
    <td><table width="60%" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr bgcolor="#FFFFFF">
        <td width="107" height="30" align="center"> ����������</td>
        <td width="345" height="30"><input name="FrName" type="text" id="FrName" size="15" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="29" align="center">�������䣺</td>
        <td height="30"><input name="FrMail" type="text" id="FrMail" maxlength="50" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="29" align="center">����������</td>
        <td height="30"><input name="MyName" type="text" id="MyName" size="15" maxlength="20" /></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="26" align="center">�������䣺</td>
        <td height="30"><input name="MyMail" type="text" id="MyMail" /></td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td height="23" colspan="2" > �ʼ�����</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td height="56" colspan="2"><br />
          ����!<br />
          ����<a href="<%=WebUrl%>" target="_blank">[<%=WebName%>]</a>�Ͽ���һƪ����Ϊ<font color="#FF0000"><%=RS("Title")%></font>����Ϣ��ϣ���ܶ�������������<br />
          ��ַΪ��<a href="<%=ArticleUrl%>" target="_blank"><%=ArticleUrl%></a><br />
          <input name="Content" type="hidden" id="Content" value="&lt;br&gt;����!&lt;br&gt;����&lt;a href=<%=WebUrl%> target=_blank&gt;[<%=WebName%>]&lt;/a&gt;�Ͽ���һƪ����Ϊ&lt;font color=#FF0000&gt;<%=RS("Title")%>&lt;/font&gt;����Ϣ��ϣ���ܶ�������������&lt;br&gt;��ַΪ��&lt;a href=<%=ArticleUrl%> target=_blank&gt;<%=ArticleUrl%>&lt;/a&gt;&lt;br&gt;" />
        </td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td height="28" colspan="2"><input type="hidden" value="<%=WebName%>" name="SiteName">
            <input type="hidden" name="ID" value="<%=ID%>">
            <input type="hidden" name="m" value="<%=channelid%>">
            <input type="submit" name="Submit" class="fmbtn" value="�����͡��ʡ���" /></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</body>
</html> 

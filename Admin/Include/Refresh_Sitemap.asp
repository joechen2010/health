<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin_Style.CSS" rel="stylesheet" type="text/css">
<title>Google Sitemap</title>

<%


dim xmlstr,lastmod
dim sql_KS_Class,SqlStr,rs,rsclass,i,Classpath
dim sitemappath
Dim KS:Set KS=New PublicCls


'=========================������==========================
If KS.G("Action")<>"" Then
    Dim changefreq:changefreq=KS.G("changefreq")
	Dim prioritynum:prioritynum=KS.ChkCLng(KS.G("prioritynum"))
	dim tmFile,objFso,smw
	sitemappath=KS.Setting(3)&"sitemap.xml"
	Set objFso = KS.InitialObject(KS.Setting(99))

	if KS.G("Action")="creategoogle" then
		If prioritynum=0 then prioritynum=15
		Dim big:big=KS.G("Big")
		Dim SQL,K
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select BasicType,ChannelTable,ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And Channelid<>9 And ChannelID<>10 Order By ChannelID",Conn,1,1
		SQL=RS.GetRows(-1)
		RS.Close

		xmlstr="<?xml version=""1.0"" encoding=""UTF-8""?>"&vbcrlf
		xmlstr=xmlstr&"<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"&vbcrlf
	
		For K=0 To Ubound(SQL,2)
		 Select Case  SQL(0,K)
		  Case 1 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate"
		  Case 2 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 3 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 4 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate"
		  Case 5 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 7 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		  Case 8 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate"
		 End Select
		
		SqlStr=SqlStr & " from "& SQL(1,K) & " where verific=1 and deltf=0 order by id desc"
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&"    <url>"&vbcrlf
			xmlstr=xmlstr&"        <loc><![CDATA["&KS.GetItemUrl(SQL(2,K),RS(2),RS(0),RS(5))&"]]></loc>"&vbcrlf
			xmlstr=xmlstr&"        <lastmod>" & GetDate(rs(7)) & "</lastmod>"&vbcrlf
			xmlstr=xmlstr&"        <changefreq>"&changefreq&"</changefreq>"&vbcrlf
			xmlstr=xmlstr&"        <priority>"&big&"</priority>"&vbcrlf
			xmlstr=xmlstr&"    </url>"&vbcrlf
			rs.movenext 
		next
		rs.close
	  Next
	'=sitemap===============================================================================================================
		xmlstr=xmlstr&"</urlset>"
	
	
		'==============д��sitemap======================
		Call KS.WriteTOFile(sitemappath,xmlstr)
	   '===========sitemap================================
	
	response.write("<script language='JavaScript' type='text/JavaScript'>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>��ϲ,sitemap.xml������ϣ�<br><br><a href=" & KS.Setting(3) & "sitemap.xml target=_blank>����鿴���ɺõ�sitemap.xml�ļ�</a></div>'; }")
	response.write("</script>")
	
	elseif  KS.G("Action")="createbaidu" then
	
		xmlstr="<?xml version=""1.0"" encoding=""gb2312""?>"&vbcrlf
	    xmlstr=xmlstr & "<document>"
		xmlstr=xmlstr & "<webSite>" & Replace(KS.Setting(2),"http://","") & "</webSite>"
		xmlstr=xmlstr & "<webMaster>" & KS.Setting(11) &"</webMaster>"
		xmlstr=xmlstr & "<updatePeri>" &changefreq & "</updatePeri>"
		Dim Num:Num=0
		Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select BasicType,ChannelTable,ChannelID From KS_Channel Where ChannelStatus=1 And ChannelID<>6 And Channelid<>9 And ChannelID<>10 Order By ChannelID",Conn,1,1
		SQL=RS.GetRows(-1)
		RS.Close

		
		For K=0 To Ubound(SQL,2)
		 Select Case  SQL(0,K)
		  Case 1 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Intro,ArticleContent,PhotoUrl,author,origin"
		  Case 2 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,PictureContent,PictureContent,photourl,author,origin"
		  Case 3 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,downcontent,downcontent,photourl,author,origin"
		  Case 4 :SqlStr="select top " & prioritynum & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,flashcontent,flashcontent,photourl,author,origin"
		  Case 5 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,prointro,prointro,photourl,ProducerName,TrademarkName"
		  Case 7 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,moviecontent,moviecontent,photourl,MovieAct,MovieDQ"
		  Case 8 :SqlStr="select top " & prioritynum & " ID,Title,Tid,0,0,Fname,0,AddDate,gqcontent,gqcontent,photourl,inputer,ContactMan"
		 End Select

		
		 SqlStr=SqlStr & " from "& SQL(1,K) & " where verific=1 and deltf=0 order by id desc"
		 
		 
		rs.Open SqlStr,conn,1,1
		for i=1 to rs.RecordCount
			xmlstr=xmlstr&"    <item>"&vbcrlf
			xmlstr=xmlstr&"        <title>" & rs(1) &"</title>"
			xmlstr=xmlstr&"        <link><![CDATA["&KS.GetItemUrl(SQL(2,K),RS(2),RS(0),RS(5))&"]]></link>"&vbcrlf
			xmlstr=xmlstr&"        <description><![CDATA[" & Replace(KS.LoseHtml(rs(8)),"&nbsp;","") & "]]></description>"&vbcrlf
			xmlstr=xmlstr&"        <text><![CDATA[" &Replace(KS.LoseHtml(rs(9)),"&nbsp;","") & "]]></text>"&vbcrlf
			if Not KS.IsNul(RS(10)) Then
			 Dim PhotoUrl:PhotoUrl=RS(10)
			 If Left(Lcase(PhotoUrl),4)<>"http" Then
			   PhotoUrl=KS.Setting(2) & PhotoUrl
			 End If
			xmlstr=xmlstr&"        <image><![CDATA["&PhotoUrl&"]]></image>"&vbcrlf
			End If
			xmlstr=xmlstr&"        <category><![CDATA["&KS.C_C(RS(2),0)&"]]></category>"&vbcrlf
			xmlstr=xmlstr&"        <author><![CDATA["&rs(11)&"]]></author>"&vbcrlf
			xmlstr=xmlstr&"        <source><![CDATA["&rs(12)&"]]></source>"&vbcrlf
			xmlstr=xmlstr&"        <pubDate>"&GetDate(rs(7))&"</pubDate>"&vbcrlf
			xmlstr=xmlstr&"    </item>"&vbcrlf
			Num=Num+1
			If Num>=100 Then Exit For
			rs.movenext 
		next
		rs.close
		If Num>=100 Then Exit For
	  Next
	'=sitemap===============================================================================================================
		
		
		xmlstr=xmlstr & "</document>"
	
		'==============д��news.xml======================
		Dim NewsPath:NewsPath=KS.Setting(3) &"news.xml"
		Call KS.WriteTOFile(NewsPath,xmlstr)
	   '===========sitemap================================

	
	response.write("<script language='JavaScript' type='text/JavaScript'>")
	response.write("function yy() {")
	response.write("overstr.innerHTML='<div align=center>��ϲ,news.xml������ϣ�<br><br><a href=" & KS.Setting(3) & "news.xml target=_blank>����鿴���ɺõ�news.xml�ļ�</a></div>'; }")
	response.write("</script>")
	end if
	
	'===================================================
		set rs=nothing
End If


response.write("<script language='JavaScript' type='text/JavaScript'>")
response.write("function ll() { ")
response.write("overstr.innerHTML='<div align=center>�������ɣ������ĵȴ�������<br></div>'; } ")
response.write("</script>")

set rs=nothing
conn.Close:set conn=nothing
'===================================================����
Function GetDate(DateStr)
	if KS.G("Action")="creategoogle" then
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
	else
	GetDate=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)& " " & Right("0" &hour(DateStr),2) &":" & Right("0" &minute(DateStr),2)& ":" & Right("0" & Second(DateStr),2)
	end if
End Function
%>


</head>

<body onLoad="yy()">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td height="25" class="Sort">
 <div align="center"><strong>XML��ͼ���ɲ���</strong></div></td>
</tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td><div id="overstr"></div></td>
  </tr>
</table>

<form id="form1" name="bqsitemapform" method="post" action="?action=creategoogle">

<table width="600" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
  <tr class="Title">
    <td>��XML��ͼ���ɲ���</td>
  </tr>
  <tr class="tdbg">
    <td height="17" align="center">
	<a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'><img border=0 src="../images/GoogleSiteMaplogo.gif" /></a>���ɷ���GOOGLE�淶��XML��ʽ��ͼҳ��
	<br /></td>
  </tr>
  <tr class="tdbg">
    <td height="18">����Ƶ�ʣ�
      <select name="changefreq" id="changefreq">
        <option value="always ">Ƶ���ĸ���</option>
        <option value="hourly">ÿСʱ����</option>
        <option value="daily" selected="selected">ÿ�ո���</option>
        <option value="weekly">ÿ�ܸ���</option>
        <option value="monthly">ÿ�¸���</option>
        <option value="yearly">ÿ�����</option>
        <option value="never">�Ӳ�����</option>
      </select></td>
  </tr>
  <tr class="tdbg">
    <td height="35">ÿ��ϵͳ���ã�
      <input name="prioritynum" type="text" id="prioritynum" value="15" size="6" />����Ϣ����Ϊ���ע���
	 </td>
  </tr>
  <tr class="tdbg">
    <td height="35">ע �� �ȣ�
      <input name="big" type="text" id="big" value="0.5" size="6" />0-1.0֮��,�Ƽ�ʹ��Ĭ��ֵ

	  <br>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="��ʼ����sitemap" /></td>
  </tr>
</table>
</form>


<form id="form1" name="bqsitemapform" method="post" action="?action=createbaidu">

<table width="600" border="0" align="center" cellpadding="6" cellspacing="0" class="border">
  <tr class="Title">
    <td>��ٶ����ſ���Э��XML���ɲ���</td>
  </tr>
  <tr class="tdbg">
    <td height="17" align="center">
	<a href='http://news.baidu.com/newsop.html#kg' target='_blank'><img border=0 src="../images/baidulogo.gif" /></a>���ɷ��ϰٶ�XML��ʽ�Ŀ�������Э��
	<br /></td>
  </tr>
  <tr class="tdbg">
    <td height="18">�������ڣ�      
      <input name="changefreq" type="text" id="changefreq" value="15" size="8"> 
      ���� </td>
  </tr>
  <tr class="tdbg">
    <td height="35">ÿ��ϵͳ���ã�
      <input name="prioritynum" type="text" id="prioritynum" value="50" size="6" />
      ����Ϣ����Ϊ���ע���(���100��)	 </td>
  </tr>
</table>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td height="45" align="center"><input name="Submit1"  class="button" onClick="ll();" type="submit" id="Submit1" value="��ʼ����sitemap" /></td>
  </tr>
</table>
</form>

<br />
</body>
</html>

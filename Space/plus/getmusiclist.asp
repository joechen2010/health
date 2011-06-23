<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
on error resume next
Dim KS,UserName,SQL,I
Set KS=New PublicCls
response.expires=0
response.ContentType="text/xml"

UserName=KS.R(KS.S("UserName"))
Dim RS:Set RS=Server.CreateObject("adodb.recordset")
rs.open "select top 100 songname,url from ks_blogmusic where username='" & username & "' order by adddate desc",conn,1,1
If Not RS.Eof Then SQL=RS.GetRows(-1)
RS.Close:Set rs=nothing
closeconn()
set KS=Nothing
Response.CodePage=65001
Response.Charset="utf-8"
%>

<?xml version="1.0" encoding="utf-8" ?>
<playlist version="1" xmlns="http://xspf.org/ns/0/">
    <title>music-box</title>
    <info>http://www.jeroenwijering.com/?item=Flash_MP3_Player</info>
    <trackList>
          <%
		  for i=0 to ubound(sql,2)
		  dim u:u=sql(1,i)
          response.write "<track>" & vbcrlf
          response.write "<annotation>" & sql(0,i) & "</annotation>" & vbcrlf
          response.write "<location>" & sql(1,i) & "</location>"& vbcrlf
          response.write "<info>http://www.kesion.com/</info>" & vbcrlf
          response.write "</track>" & vbcrlf
		  next
        %>
        
    </trackList>
</playlist>

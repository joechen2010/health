<!--#include file="../conn.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'==============================================================
'����������Ҫ�����޸����´���
'���ļ����÷�ʽ:<script src='/plus/total.asp'></script>
'=================================================================
Response.Write("document.writeln('����������<font color=""red"">"&Conn.Execute("Select Count(id) From ks_Article")(0) & "</font> ƪ<br />');")
Response.Write("document.writeln('���������<font color=""red"">"&Conn.Execute("Select Count(id) From ks_download")(0) & "</font> ��<br />');") 
if DataBaseType=1 Then
Response.Write("document.writeln('���ո������£�<font color=""red"">" & conn.execute("select count(id) from ks_article where datediff(d,adddate,getdate())<1 ")(0) & "</font> ƪ<br />');")
Else
Response.Write("document.writeln('���ո������£�<font color=""red"">" & conn.execute("select count(id) from ks_article where datediff('d',adddate,now())<1 ")(0) & "</font> ƪ<br />');")

End If
If DataBaseType=1 Then
Response.Write("document.writeln('���ո��������<font color=""red"">" & conn.execute("select count(id) from ks_download where datediff(d,adddate,getdate())<1 ")(0) & "</font> ��<br />');")
Else
Response.Write("document.writeln('���ո��������<font color=""red"">" & conn.execute("select count(id) from ks_download where datediff('d',adddate,now())<1 ")(0) & "</font> ��<br />');")

End If
Response.Write("document.writeln('���������������<font color=""red"">" & conn.execute("select sum(hits) from ks_article")(0) & "</font> ��<br />');")
Response.Write("document.writeln('�����ܴ�����<font color=""red"">" & conn.execute("select sum(hits) from ks_download")(0) & "</font> ��<br />');")
%>
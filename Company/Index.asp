<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,astr
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Dim Template
		   Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "��ҵ�ռ�/company_index.html")
		   FCls.RefreshType = "enterprise" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
		   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		   call getclasslist()
		   Template=Replace(Template,"{$ShowClass}",str)
		   call getarealist()
		   Template=Replace(Template,"{$ShowAreaList}",astr)
		   Template=KSR.KSLabelReplaceAll(Template)
		 Response.Write Template  
		End Sub
		Sub GetClassList()
		 Dim RS,I,RSS
		 Set RS=Conn.Execute("select id,classname from ks_enterpriseclass where parentid=0 order by orderid")
		 str="<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
		 Do While Not RS.Eof
		 str=str & "<tr>" & vbcrlf
		 for i=1 to 2
		   str=str & "<td width=""50%"" style=""padding:5px"">" & vbcrlf
		   str=str & "<div style=""height:20px;"" class=""classname""><span style=""font-weight:bold""><img src=""../images/arrow_r.gif""> <a href=""list.asp?pid=" & rs(0) & """>" & rs(1) &"</a></span>(" & conn.execute("select count(id) from ks_enterprise where status=1 and classid=" & rs(0))(0) &") </div>" & vbcrlf
		   str=str & "<div style=""height:50px;padding-left:12px"">"
		   dim xml,node,num,n
		   set rss=conn.execute("select id,classname from ks_enterpriseclass where parentid=" & rs(0))
		   if not rss.eof then set xml=KS.RsToXml(rss,"row","") else xml=empty
		   rss.close:set rss=nothing
		   if isobject(xml) then
		       num=xml.DocumentElement.SelectNodes("row").length : n=0
			   for each node in xml.DocumentElement.SelectNodes("row") 
				str=str & "<a href='list.asp?id=" & node.selectsinglenode("@id").text & "'>" & node.selectsinglenode("@classname").text & "</a>"
				n=n+1
				if num<>n then str=str & " | "
			   next
			 xml=empty : set node=nothing
		   end if
		   str=str & "</div>"
		   str=str & "</td>" & vbcrlf
		   rs.movenext
		   if rs.eof then exit for
		 next
		 str=str & "</tr>"
		 Loop
		 str=str & "</table>" & vbcrlf
		 rs.close:set rs=nothing
		End Sub
		
		Sub getarealist()
		  Dim RS,I,SQL,K,N
		  Set RS=Conn.Execute("Select id,city from KS_Province where parentid=0 order by orderid")
		  IF Not RS.Eof Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		  If IsArray(SQL) Then
			  astr="<table border='0' width='100%'>" &vbcrlf
			  N=0
			  For i=0 To Ubound(SQL,2)
				astr=astr & "<tr>" &vbcrlf
				For K=1 To 3
				astr=astr & "<td><img src='../images/arrow_r.gif'> <a href=""list.asp?province=" & sql(1,n) & "&provinceid=" & SQL(0,n) & """>" & sql(1,n) & "</a></td>"
				n=n+1
				if n>Ubound(SQL,2) then Exit For
				Next
				astr=astr & "</tr>" &vbcrlf
				if n>Ubound(SQL,2) then Exit For
			 Next
			 astr=astr & "</table>" & vbcrlf
		 End If
		End Sub
End Class
%>

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
        Private KS, KSR,str,c_str,curr_tips,Template
		Private TotalPut,MaxPerPage,CurrentPage,UserName
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If KS.S("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
			Else
			  CurrentPage = 1
			End If
			UserName=KS.S("UserName")

				   Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "��ҵ�ռ�/company_show.html")
				   FCls.RefreshType = "enterpriselist" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   call getcompany()
				   call getsupply()
				   Template=KSR.KSLabelReplaceAll(Template)
		           Response.Write Template  
		End Sub
		
		Sub GetCompany()
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open "select * from ks_enterprise where status=1 and username='" & KS.S("UserName") & "'",conn,1,1
		 IF RS.Eof And RS.Bof Then
			 Call KS.ShowTips("error","<li>�������ݳ���!</li>")
			 Exit Sub
		 Else
		    on error resume next
			template=replace(template,"{$ShowCompanyName}",rs("companyname"))
			template=replace(template,"{$ShowCompanyIntro}",KS.HtmlCode(rs("intro")))
			template=replace(template,"{$ShowIndustry}",conn.execute("select classname from ks_enterpriseclass where id=" & rs("classid"))(0) & "&nbsp;" &conn.execute("select classname from ks_enterpriseclass where id=" & rs("smallclassid"))(0) )
			template=replace(template,"{$ShowLegalPeople}",LFCls.ReplaceDBNull(RS("legalpeople"),"---"))
			template=replace(template,"{$ShowCompanyScale}",LFCls.ReplaceDBNull(RS("companyscale"),"---"))
			template=replace(template,"{$ShowRegisteredCapital}",LFCls.ReplaceDBNull(RS("RegisteredCapital"),"---"))
			template=replace(template,"{$ShowProvince}",LFCls.ReplaceDBNull(RS("province"),"---"))
			template=replace(template,"{$ShowCity}",LFCls.ReplaceDBNull(RS("city"),"---"))
			template=replace(template,"{$ShowContactMan}",LFCls.ReplaceDBNull(RS("contactman"),"---"))
			template=replace(template,"{$ShowAddress}",LFCls.ReplaceDBNull(RS("address"),"---"))
			template=replace(template,"{$ShowZipCode}",LFCls.ReplaceDBNull(RS("zipcode"),"---"))
			template=replace(template,"{$ShowTelphone}",LFCls.ReplaceDBNull(RS("telphone"),"---"))
			template=replace(template,"{$ShowFax}",LFCls.ReplaceDBNull(RS("fax"),"---"))
			template=replace(template,"{$ShowBankAccount}",LFCls.ReplaceDBNull(RS("bankaccount"),"---"))
			template=replace(template,"{$ShowAccountNumber}",LFCls.ReplaceDBNull(RS("accountnumber"),"---"))
			if rs("weburl")="" or rs("weburl")="http://" then
			template=replace(template,"{$ShowWebSite}","��")
			else
			template=replace(template,"{$ShowWebSite}","<a href='" & rs("weburl") & "' target='_blank'>" & rs("weburl") & "</a>")
			end if
							
		 End IF
			RS.Close:Set RS=Nothing
		End Sub
		
		Sub GetSupply()
		 Dim rs,I,logo,n,url,c_str
		 c_str="<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">" & vbcrlf
         c_str=c_str & "<tr bgcolor=""#E7E7E7"">"
         c_str=c_str & "<td width=""111"" height=""26"" align=""center"">ͼƬ</td>"
         c_str=c_str & "<td width=""300"" align=""center"">����/��Ҫ����</td>"
         c_str=c_str & "<td width=""85"" align=""center"">����</td>"
         c_str=c_str & "<td width=""90"" align=""center"">����</td>"
         c_str=c_str & "<td width=""150"" align=""center"">��������</div></td>"
         c_str=c_str & "</tr>"
		 Set RS=Server.CreateOBject("ADODB.RECORDSET")
		 RS.Open "Select top 10 typename,typecolor,a.* From KS_GQ a inner join ks_gqtype b on a.typeid=b.typeid Where a.verific=1 and inputer='" & UserName & "'",conn,1,1
		 Do While Not RS.Eof
		 logo=rs("photourl")
		 if KS.IsNul(logo) then logo="/images/logo.jpg"
		 url=KS.GetItemUrl(8,RS("TID"),RS("id"),RS("Fname"))
         n=n+1
		 if n mod 2=0 then
		 c_str=c_str & "<tr bgcolor=""#f6f6f6"">"
		 else
         c_str=c_str & "<tr>"
		 end if
         c_str=c_str & "<td width=""111"" height=""80"" align=""center""><div style=""border:1px solid #666666;padding:5px""><img src=""" & logo & """ width=88 height=50></div></td>"
         c_str=c_str & "<td width=""300"" style=""WORD-BREAK: break-all""><a href=""" & url & """ target=""_blank""><div style='font-weight:bold;font-size:14px;text-decoration:underline;margin:2px;'>" & RS("Title") &"</div></a>" & KS.Gottopic(KS.LoseHtml(KS.HtmlCode(RS("GQContent"))),120) &"...</td>"
         c_str=c_str & "<td width=""85"" align=""center""><font color='" & rs(1) & "'>" & RS(0) & "</font></td>"
         c_str=c_str & "<td width=""90"" align=""center"">" & RS("province") & "&nbsp;" & RS("City") & "</td>"
         c_str=c_str & "<td width=""150"" align=""center"">" & RS("AddDate") & "</div></td>"
         c_str=c_str & "</tr>"
		 I=I+1
		If I >= MaxPerPage Then Exit Do
		 RS.MoveNext
		 Loop
         c_str=c_str & "</table>"
		 RS.Close:Set RS=Nothing
		 Template=Replace(Template,"{$ShowCompanySupply}",c_str)
		End Sub
		
End Class
%>

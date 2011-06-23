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
        Private KS, KSR,str,c_str,pid
		Private TotalPut,MaxPerPage,CurrentPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  MaxPerPage=20
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
			Pid=KS.ChkClng(KS.S("Pid"))

		           Dim Template
				   Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "企业空间/news_list.html")
				   FCls.RefreshType = "enterpriselist" '设置刷新类型，以便取得当前位置导航等
				   call getnewslist()
				   Template=Replace(Template,"{$ShowNewsList}",c_str)
				   Template=KSR.KSLabelReplaceAll(Template)
		 Response.Write Template  
		End Sub


		
		Sub getnewslist()
		 Dim Param:Param=" where a.status=1 order by a.adddate desc" 

		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open "select b.[Domain],a.* from ks_enterprisenews a inner join ks_blog b on a.username=b.username" & Param,conn,1,1
		 IF RS.Eof And RS.Bof Then
			  totalput=0
			  exit sub
		  Else
							TotalPut= Conn.Execute("Select count(*) from KS_EnterpriseNews a inner join ks_blog b on a.username=b.username where a.status=1")(0)
							If CurrentPage < 1 Then CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (TotalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call ShowContent(RS)
			End IF
			
			c_str =c_str & "<div style='text-align:right'>" &  KS.ShowPagePara(totalPut, MaxPerPage, "", true, "条", CurrentPage, "") & "</div>"
			
			RS.Close
			Set RS=Nothing
		End Sub
		
		Sub ShowContent(RS)
		 on error resume next
		 Dim I,logo,n,url

		 c_str=c_str & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">" & vbcrlf
		 Do While Not RS.Eof
         'If KS.SSetting(14)="1" and rs(1)<>"" then 
		 '	  url="http://" & rs(0) &"." & KS.SSetting(16) & "/Space/Show_News.asp?username=" & RS("UserName") & "&id=" &RS("ID")
		 '	 else
			  url="../space/?" & RS("UserName") & "/shownews/" &RS("ID")
		 '	 end if
         n=n+1
		 if n mod 2=0 then
		 c_str=c_str & "<tr bgcolor=""#f6f6f6"">"
		 else
         c_str=c_str & "<tr>"
		 end if
         c_str=c_str & "<td style=""height:22px;border-bottom:1px dashed #cccccc""><img src='../images/h_arrow.gif' align='absmiddle'> <a href=""" & URL & """ target=""_blank"">" & RS("Title") & "</a></td>"
         c_str=c_str & "<td style=""height:22px;border-bottom:1px dashed #cccccc"" width=""285"" align=""center""><a href='../space/?" & RS("UserName")& "' target='_blank'>" & Conn.Execute("Select top 1 CompanyName From KS_EnterPrise Where UserName='" & RS("UserName") & "'")(0) & "</a></td>"
         c_str=c_str & "<td  style=""height:22px;border-bottom:1px dashed #cccccc"" width=""80"" align=""center"">" & month(RS("AddDate")) & "-" & day(rs("adddate")) & "</td>"
         c_str=c_str & "</tr>"
		 I=I+1
		If I >= MaxPerPage Then Exit Do
		 RS.MoveNext
		 Loop
         c_str=c_str & "</table>"
		End Sub
		
End Class
%>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Link
KSCls.Kesion()
Set KSCls = Nothing

Class Link
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 Dim Template,KSR
		 FCls.RefreshType = "LinkIndex"   '���õ�ǰλ��Ϊ����������ҳ
		 Set KSR = New Refresh
			If KS.Setting(113)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
		    Template = KSR.LoadTemplate(KS.Setting(113))
			Template = ReplaceListContent(Template)     '�滻��������ҳ��ǩΪ����
			Template = KSR.KSLabelReplaceAll(Template)
		   Set KSR = Nothing
		   
		   Response.Write Template   
	End Sub
	    '*********************************************************************************************************
		'��������ReplaceLinkContent
		'��  �ã��滻��������ҳ��ǩΪ����
		'��  ����Template���滻������
		'*********************************************************************************************************
		Function ReplaceListContent(Template)
		   '  on error resume next
			 Dim Domain, ClassLinkStr, DetailListStr, KeyWord
			 Dim RClassID, ClassID, LinkType, ViewKind
			 Dim ObjRS:Set ObjRS=Server.CreateObject("ADODB.Recordset")
			 
			   Domain = KS.GetDomain()
			   RClassID = KS.ChkClng(KS.S("ClassID"))
			   LinkType = KS.ChkClng(KS.S("LinkType"))
			   ViewKind = KS.ChkClng(KS.S("ViewKind"))
			   KeyWord = KS.S("KeyWord")
			   IF ViewKind=0 Then ViewKind=1
			   If LinkType = 0 Then LinkType = 2
			   If Not IsNumeric(RClassID) Then
				Call KS.Alert("�Ƿ�����!", "")
				Set KS = Nothing:Exit Function
			   End If
			   
			   If InStr(Template, "{$GetLinkCommonInfo}") <> 0 Then
				  Template = Replace(Template, "{$GetLinkCommonInfo}", "<a href=""" & Domain & "plus/link/"">����鿴</a> | <a href=""" & Domain & "plus/link/?ViewKind=1"">��������鿴</a> | <a href=""" & Domain & "plus/link/?ViewKind=2"">�����鿴</a> | <a href=""" & Domain & "plus/link/?ViewKind=3"">�����Ƽ�վ��</a> | <a href=""" & Domain & "plus/link/reg"">������������</a> ")
			   End If
			   If InStr(Template, "{$GetClassLink}") <> 0 Then
				  
				  ClassLinkStr = "<table width=""100%"" border=""0""><form action=""?"" method=""get"" name=""SearchLink""><tr><td>������ʾ��<select name='LinkType' id='LinkType' onchange=""if(this.options[this.selectedIndex].value!=''){location='?LinkType='+this.options[this.selectedIndex].value;}"">"
				  If LinkType = 2 Then
				  ClassLinkStr = ClassLinkStr & "<option value='2' selected>��������</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='2'>��������</option>"
				  End If
				  If LinkType = 1 Then
				  ClassLinkStr = ClassLinkStr & "<option value='1' selected>LOGO����</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='1'>LOGO����</option>"
				  End If
				  If LinkType = 0 Then
				  ClassLinkStr = ClassLinkStr & "<option value='0' selected>��������</option>"
				  Else
				  ClassLinkStr = ClassLinkStr & "<option value='0'>��������</option>"
				  End If
				  ClassLinkStr = ClassLinkStr & "</select>"
				  
				  ClassLinkStr = ClassLinkStr & "&nbsp;<select name='ViewClassID' id='ViewClassID' onchange=""if(this.options[this.selectedIndex].value!=''){location='?LinkType=" & LinkType & "&ClassID='+this.options[this.selectedIndex].value;}""><option value='0'>���з���վ��</option>"
				  ObjRS.Open "Select FolderID,FolderName From KS_LinkFolder Order BY OrderID,FolderID Desc", Conn, 1, 1
				  If Not ObjRS.EOF Then
				   Do While Not ObjRS.EOF
					ClassID = ObjRS(0)
					If CStr(RClassID) = CStr(ClassID) Then
					ClassLinkStr = ClassLinkStr & "<option value='" & ClassID & "' selected>" & ObjRS(1) & "</option>"
					Else
					ClassLinkStr = ClassLinkStr & "<option value='" & ClassID & "'>" & ObjRS(1) & "</option>"
					End If
					ObjRS.MoveNext
				   Loop
				  End If
				  
				  ObjRS.Close
				  ClassLinkStr = ClassLinkStr & "</select>&nbsp;&nbsp;�ؼ��֣�<input class=""textbox"" type=""text"" size=""22"" name=""KeyWord""> &nbsp;<input class=""inputbutton"" type=""submit"" value="" �� �� ""></td></tr></form></table>"
				  Template = Replace(Template, "{$GetClassLink}", ClassLinkStr)
			   End If
			   
			   If InStr(Template, "{$GetLinkDetail}") <> 0 Then
					Dim totalPut, CurrentPage,Para
					
				  If ViewKind = 2 Then      '������鿴
					 If RClassID = 0 Then
					   Dim CRS:Set CRS=Server.CreateObject("ADODB.Recordset")
					   CRS.Open "Select FolderID,FolderName From KS_LinkFolder Order BY AddDate Desc", Conn, 1, 1
					   If CRS.EOF And CRS.BOF Then
						 DetailListStr = "��û���κ���������վ�����!"
					   Else
						 DetailListStr = "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" Class=""table_border""><tr><td>"
						 DetailListStr = DetailListStr & "<table width='100%' cellSpacing=2 cellPadding=1 border=0>"
						  Do While Not CRS.EOF
							DetailListStr = DetailListStr & "<tr><td Class=""link_table_title""><a href='Index.asp?ViewKind=2&ClassID=" & CRS(0) & "'><b>" & CRS(1) & "</b></a></td></tr>"
							DetailListStr = DetailListStr & GetClassSiteList(CRS(0))
							CRS.MoveNext
						  Loop
						  DetailListStr = DetailListStr & "</table></td></tr></table>"
					   End If
						CRS.Close
						Set CRS = Nothing
					 Else
						  DetailListStr = "<table width=""100%""  cellpadding=""0"" cellspacing=""0""  Class=""table_border""><tr><td>"
						  DetailListStr = DetailListStr & "<table width='100%' cellSpacing=2 cellPadding=1 border=0>"
						  DetailListStr = DetailListStr & "<tr><td  Class=""link_table_title""><b>"
						  
						  Dim ClassRS
						  Set ClassRS = Conn.Execute("Select FolderName From KS_LinkFolder Where FolderID=" & RClassID)
						  DetailListStr = DetailListStr & ClassRS(0)
						  ClassRS.Close
						  Set ClassRS = Nothing
						  
						  DetailListStr = DetailListStr & "</b></td></tr>"
						  DetailListStr = DetailListStr & GetClassSiteList(RClassID)
						  DetailListStr = DetailListStr & "</table></td></tr></table>"
					 End If
				  Else                      '������ȷ�ʽ�鿴
				  
					 Const MaxPerPage = 20   'ÿҳ��ʾ����
					If KS.S("page") <> "" Then
					   CurrentPage = KS.ChkClng(KS.S("page"))
					Else
					  CurrentPage = 1
					End If
					
					DetailListStr = "<TABLE WIDTH=""100%""  Cellpadding=""0"" Cellspacing=""0"" Class=""table_border""><TR><TD>"
					
					  Para = " Where Verific=1 And Locked=0"
					If LinkType = 0 Or LinkType = 1 Then
					  Para = Para & " And LinkType=" & LinkType
					End If
					If RClassID <> 0 Then
					  Para = Para & " And FolderID=" & RClassID
					End If
					If KeyWord <> "" Then
					  Para = Para & " And SiteName like '%" & KeyWord & "%' Or Description like '%" & KeyWord & "%'"
					End If
					If ViewKind = 3 Then
					  Para = Para & " And Recommend=1 Order By Hits Desc"
					ElseIf ViewKind = 1 Then
					  Para = Para & " Order By Hits Desc"
					Else
					  Para = Para & " Order By AddDate Desc"
					End If
					ObjRS.Open "Select * From KS_Link" & Para, Conn, 1, 1
					If ObjRS.EOF And ObjRS.BOF Then
					   If RClassID = 0 Then
						  DetailListStr = DetailListStr & "��û�м����κ���������!"
					   Else
						  DetailListStr = DetailListStr & "û�и�������������վ��!"
					   End If
					Else
					   totalPut = ObjRS.RecordCount
							If CurrentPage < 1 Then CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
							If CurrentPage = 1 Then
								  DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									ObjRS.Move (CurrentPage - 1) * MaxPerPage
								   DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
								Else
									CurrentPage = 1
								   DetailListStr = DetailListStr & GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
								End If
							End If
				   End If
					ObjRS.Close
					Set ObjRS = Nothing
					DetailListStr = DetailListStr & "</TD></TR></TABLE>"
			   End If
				  Template = Replace(Template, "{$GetLinkDetail}", DetailListStr)
			   End If
			   ReplaceListContent = Template
		End Function
		'�������ReplaceListContent����ʹ��
		Function GetDetailListStr(ObjRS, totalPut, MaxPerPage, CurrentPage, RClassID)
			  Dim AddDate, I, RecommendStr,LinkID
				  Do While Not ObjRS.EOF
					   AddDate = ObjRS("AddDate")
					   LinkID = ObjRS("LinkID")
					   If ObjRS("Recommend") = 1 Then
						RecommendStr = " <font color=""red"">�Ƽ�</font>"
					   Else
						RecommendStr = ""
					   End If
					   GetDetailListStr = GetDetailListStr & "<TABLE cellSpacing=1 cellPadding=4 width=100% align=center bgColor=#ffffff border=0>"
					   GetDetailListStr = GetDetailListStr & "<TR Class=""link_table_title"" height=20>"
					   If ObjRS("LinkType") = 0 Then
					   GetDetailListStr = GetDetailListStr & "<TD width=""14%""><a href=""Index.asp?LinkType=0"" title=""���������Ӳ鿴"">��������</a></TD>"
					   Else
					   GetDetailListStr = GetDetailListStr & "<TD width=""14%""><a href=""Index.asp?LinkType=1"" title=""��LOGO���Ӳ鿴"">LOGO����</a></TD>"
					   End If
					   GetDetailListStr = GetDetailListStr & "<TD width=""36%""><A href = ""to?" & LinkID & """ target=""_blank"" title=""��վ����""><B>" & ObjRS("SiteName") & "</B>  " & RecommendStr & "</A></TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"">"
					   
					   on error resume next
					   Dim ClassRS:Set ClassRS = Conn.Execute("Select FolderID,FolderName From KS_LinkFolder Where FolderID=" & ObjRS("FolderID"))
					   GetDetailListStr = GetDetailListStr & "<a href=""Index.asp?ViewKind=2&ClassID=" & ClassRS(0) & """  Title=""��վ���"">" & ClassRS(1) & "</a>"
					   ClassRS.Close:Set ClassRS = Nothing
					   
					   GetDetailListStr = GetDetailListStr & "</TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""12%"" nowrap><a href=""mailto:" & ObjRS("Email") & """ Title=""��վվ��"">" & ObjRS("WebMaster") & "</a></TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"" nowrap>" & Year(AddDate) & "-" & Month(AddDate) & "-" & Day(AddDate) & "</TD>"
					   GetDetailListStr = GetDetailListStr & "<TD width=""15%"" nowrap>��� <B>" & ObjRS("Hits") & "</B> ��</TD>"
					   GetDetailListStr = GetDetailListStr & "</TR>"
					   GetDetailListStr = GetDetailListStr & "<TR height=40>"
					   GetDetailListStr = GetDetailListStr & "<TD Style = ""BORDER-RIGHT: #efefef 1px dotted; BORDER-LEFT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted"" align=middle width=""14%""><table border=0><tr><td>"
					   
					   If ObjRS("LinkType") = 0 Then
						GetDetailListStr = GetDetailListStr & "<A href = ""to?" & LinkID & """ target=""_blank""><IMG height=31 src=""../../Images/Default/NoLinkLogo.gif"" alt=" & ObjRS("SiteName") & " width=88 border=0></A></td></tr>"
					   Else
						GetDetailListStr = GetDetailListStr & "<A href = ""to?" & LinkID & """ target=""_blank""><IMG height=31 src=""" & ObjRS("Logo") & """ alt=" & ObjRS("SiteName") & " width=88 border=0></A></td></tr>"
					   End If
					   GetDetailListStr = GetDetailListStr & "<tr><td align=""center""><a href=""modify/?LinkID=" & LinkID & """>�޸�</a> <a href=""del/?id=" & LinkID & """>ɾ��</a></td></tr></table></TD>"
					   GetDetailListStr = GetDetailListStr & "<TD style=""BORDER-RIGHT: #efefef 1px dotted; BORDER-BOTTOM: #efefef 1px dotted"" title=""��վ���"" colSpan=5>"
					   If Trim(ObjRS("Description")) = "" Then
						 GetDetailListStr = GetDetailListStr & "���޼��"
					   Else
						 GetDetailListStr = GetDetailListStr & KS.HtmlCode(ObjRS("Description"))
					   End If
					   GetDetailListStr = GetDetailListStr & "</TD></TR><TR><TD colSpan=6 height=3></TD></TR>"
					   GetDetailListStr = GetDetailListStr & "</TABLE>"
					 ObjRS.MoveNext
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
					 Loop
					 GetDetailListStr = GetDetailListStr & "<table width=""100%"" aling=""center""><tr><td align=right>" & KS.ShowPagePara(totalPut, MaxPerPage, "Index.asp", True, "��վ��", CurrentPage, "ClassID=" & RClassID & "&LinkType=" & KS.S("LinkType") & "&ViewKind=" & KS.S("ViewKind")) & "</td></tr></table>"
		End Function
		'�������ReplaceListContent����ʹ��
		Function GetClassSiteList(FolderID)
				Dim ObjRS:Set ObjRS=Server.CreateObject("ADODB.Recordset")
				Dim SiteName,I
				
				FolderID = KS.ChkClng(FolderID)
				GetClassSiteList = "<tr><td>"
				ObjRS.Open "Select LinkID,sitename From KS_Link Where FolderID=" & FolderID & " And Verific=1 And Locked=0", Conn, 1, 1
					If ObjRS.EOF And ObjRS.BOF Then
						GetClassSiteList = GetClassSiteList & "�������û���κ�վ��!"
					Else
						 GetClassSiteList = GetClassSiteList & "<table width=""100%"" border=""0"">"
						Do While Not ObjRS.EOF
							GetClassSiteList = GetClassSiteList & "<tr>"
							For I = 1 To 6
								SiteName = ObjRS(1)
								GetClassSiteList = GetClassSiteList & "<td><a href = ""to?" & ObjRS(0) & """ target='blank' title='" & SiteName & "'>" & SiteName & "</a></td>"
								ObjRS.MoveNext
								If ObjRS.EOF Then Exit For
							Next
							GetClassSiteList = GetClassSiteList & "</tr>"
						 Loop
						 GetClassSiteList = GetClassSiteList & "</table>"
				 End If
				 GetClassSiteList = GetClassSiteList & "</td></tr>"
				 ObjRS.Close:Set ObjRS = Nothing
		End Function
End Class
%>

 

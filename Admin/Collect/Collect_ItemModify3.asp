<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Collect_ItemModify3
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify3
        Private KS
		Private KMCObj
		Private ConnItem,ChannelID,ThumbType,TbsString,TboString
		Private SqlItem, RsItem, ItemID, FoundErr, ErrMsg, Action
		Private ListStr, LsString, LoString, ListPageType, LPsString, LPoString, ListPageStr1, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3,CharsetCode
		Private LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, LoginResult, LoginData
		Private HsString, HoString, HttpUrlType, HttpUrlStr
		Private ListUrl, ListCode, ListPageNext
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
	Sub Kesion()
		ItemID = Trim(Request("ItemID"))
		Action = Trim(Request("Action"))
		
		If ItemID = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "�����������ĿID����Ϊ�գ�\n"
		Else
		   ItemID = CLng(ItemID)
		End If
		If Action = "SaveEdit" And FoundErr <> True Then
		   Call SaveEdit
		End If
		
		If FoundErr <> True Then
		   Call GetTest
		End If
		
		If FoundErr = True Then
		   Call KS.AlertHistory(ErrMsg,-1)
		Else
		   Call Main
		End If
		End Sub
		
		Sub Main()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>�ɼ�ϵͳ</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(3,ItemID) &"</div>"

		Response.Write "<form method=""post"" action=""Collect_ItemModify4.asp"" name=""form1"">"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"" >"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""20%"" align=""center"" class='clefttitle'>���ӿ�ʼ��ǣ�</td>"
		Response.Write "      <td width=""75%"">"
		Response.Write "      <textarea name=""HsString"" cols=""49"" rows=""5"">" & HsString & "</textarea></td>"
		Response.Write "    </tr>"
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'>���ӽ�����ǣ�</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""HoString"" cols=""49"" rows=""5"">" & HoString & "</textarea></td>"
		 Response.Write "   </tr>"
		 
		 '==============�б��Զ����ֶβɼ�================================
		 Dim RS,SQL,I,BeginStr,EndStr
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "Select FieldID,FieldTitle,FieldName,BeginStr,EndStr From KS_FieldItem Where ShowType=1 and ChannelID=" &ChannelID & " order by orderid",ConnItem,1,1
	 
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) Then
		   For I=0 To Ubound(SQL,2)
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>" & SQL(1,I) & "��ʼ��ǣ�<br/><br/>"
			 Response.Write "    " & SQL(1,I) &"������ǣ�<br/></td>"
			 Response.Write "    <td width=""75%"">"
			   Dim RSV:Set RSV=Server.CreateObject("ADODB.RECORDSET")
			   RSV.Open "Select BeginStr,EndStr From KS_FieldRules Where ItemID=" & ItemID & " And channelid=" & ChannelID & " and FieldName='" & SQL(2,I) &"'",ConnItem,1,1
			   If Not RSV.Eof Then
			     BeginStr=RSV(0)
				 EndStr=RSV(1)
			   Else
			     BeginStr=""
				 EndStr=""
			   End If
			   RSV.Close:Set RSV=Nothing
			 
			 Response.Write "     <textarea name=""begin" & SQL(2,I) & """ cols=""49"" rows=""3"">" & BeginStr & "</textarea><br>"
			 Response.Write "    <textarea name=""end" & SQL(2,I) & """ cols=""49"" rows=""3"">" & EndStr & "</textarea></td>"
			 Response.Write "   </tr>"
		   Next
		 End If
         '==========================================================
		 
		 
		 '�ɼ�ͼƬ���������ж�
	    ' If ChannelID=4 or ChannelID=2 Then
		 %>	 
		 <tr class='tdbg'>
		  <td align=center class="clefttitle">�б�����ͼ���ã�</td>
		  
		  <td height="25"><input type="radio" name="ThumbType" value="0" <%If ThumbType=0 Then Response.Write " checked"%> onClick="picl1.style.display='none'">
	��������
	  <input type="radio" name="ThumbType" value="1" <%If ThumbType=1 Then Response.Write " checked"%> onClick="picl1.style.display=''">
	�б��ǩ</td>
		</tr>
		<tbody id="picl1" style="display:<%If ThumbType=0 Then Response.Write "none"%>">
			  <tr>
				<td align="center" class='clefttitle'>�б�����ͼ��ʼ��ǣ�
				<br /><br />�б�����ͼ������ǣ�
				</td>
				<td><textarea name="TbsString" cols="49" rows="3" id="TbsString"><%=TbsString%></textarea>
				
				<br />
				<textarea name="TboString" cols="49" rows="3" id="TboString"><%=TboString%></textarea
				></td>
			  </tr>
		</tr>
		</tbody>
		 
	<%	'End If
	 
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""20%""  class=""clefttitle"" align=""center""> �������⴦��</td>"
		  Response.Write "    <td width=""75%"">"
		  Response.Write "      <input type=""radio"" value=""0"" name=""HttpUrlType"" "
		  If HttpUrlType = 0 Then Response.Write "checked"
		  Response.Write " onClick=""HttpUrl1.style.display='none'""> �Զ�����&nbsp;"
		  Response.Write "      <input type=""radio"" value=""1"" name=""HttpUrlType"" "
		  If HttpUrlType = 1 Then Response.Write "checked"
		  Response.Write " onClick=""HttpUrl1.style.display=''""> ���¶�λ      </td>"
		 Response.Write "   </tr>"
		  Response.Write "  <tr  class='tdbg' id=""HttpUrl1"" style=""display:'"
		  If HttpUrlType = 0 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center"">���������ַ���</td>"
		   Response.Write "   <td width=""75%"">"
		  Response.Write "      <input name=""HttpUrlStr"" type=""text"" size=""49"" maxlength=""200"" value=""" & HttpUrlStr & """></td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'>"
		  Response.Write "    <td height=""30"" colspan=""2"" align=""center"">"
		  Response.Write "      <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
		   Response.Write "     <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
		  Response.Write "      <input class='button' type=""button"" name=""button1"" value=""��&nbsp;һ&nbsp;��"" onClick=""window.location.href='javascript:history.go(-1)'""  >"
		  Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;"
		  Response.Write "    <input  type=""submit"" class='button' name=""Submit"" value=""��&nbsp;һ&nbsp;��""""></td>"
		  Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "</form>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"" >"
		Response.Write "  <tr>"
		Response.Write "    <td height=""25"" colspan=""2"" class=""sort""><div align=""center""><strong>�� �� �� ȡ �� ��</strong></div></td>"
		Response.Write "  </tr>"
		 Response.Write " <tr>"
		Response.Write "    <td height=""22"" colspan=""2"">" & ListCode & " </td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		
		If ListPageType = 1 Then
		Response.Write "<br>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"" >"
		 Response.Write " <tr>"
		 Response.Write "   <td height=""22"" colspan=""2"" >"
		   Response.Write "<br>��һҳ�б�<a  href='" & ListPageNext & "' target=_blank><font  color=red>" & ListPageNext & "</font></a>"
		   
		 Response.Write "   </td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "<br>"
		End If
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		Sub SaveEdit()
		   ListStr = Trim(Request.Form("ListStr"))
		   LsString = Request.Form("LsString")
		   LoString = Request.Form("LoString")
		   ListPageType = Request.Form("ListPageType")
		   LPsString = Request.Form("LPsString")
		   LPoString = Request.Form("LPoString")
		   ListPageStr1 = Trim(Request.Form("ListPageStr1"))
		   ListPageStr2 = Trim(Request.Form("ListPageStr2"))
		   ListPageID1 = Request.Form("ListPageID1")
		   ListPageID2 = Request.Form("ListPageID2")
		   ListPageStr3 = Request.Form("ListPageStr3")
		
		If ItemID = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "��������������Ч���ӽ���\n"
		Else
		   ItemID = CLng(ItemID)
		End If
		If LsString = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "���б�ʼ��ǲ���Ϊ��\n"
		End If
		If LoString = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "���б������ǲ���Ϊ��\n"
		End If
		If ListPageType = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "����ѡ���б�������ҳ����\n"
		Else
		   ListPageType = CLng(ListPageType)
		   Select Case ListPageType
		   Case 0, 1
					If ListStr = "" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "���б�����ҳ����Ϊ��\n"
					Else
					   ListStr = Trim(ListStr)
					End If
			  If ListPageType = 1 Then
					If LPsString = "" Or LPoString = "" Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "��������ҳ��ʼ/������ǲ���Ϊ��\n"
					End If
					If ListPageStr1 <> "" And Len(ListPageStr1) < 15 Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "��������ҳ�ض������ò���ȷ(����15���ַ�)\n"
					End If
			  End If
		   Case 2
			  If ListPageStr2 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "�����������ַ�����Ϊ��\n"
			  End If
			  If IsNumeric(ListPageID1) = False Or IsNumeric(ListPageID2) = False Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���������ɵķ�Χֻ��������\n"
			  Else
				 ListPageID1 = CLng(ListPageID1)
				 ListPageID2 = CLng(ListPageID2)
				 If ListPageID1 = 0 And ListPageID2 = 0 Then
					FoundErr = True
					ErrMsg = ErrMsg & "���������ɷ�Χ���ò���ȷ\n"
				 End If
			  End If
		   Case 3
			  If ListPageStr3 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���б�������ҳ����Ϊ�գ����ֶ����\n"
			  Else
				 ListPageStr3 = Replace(ListPageStr3, Chr(13), "|")
			  End If
		   Case Else
			  FoundErr = True
			  ErrMsg = ErrMsg & "����ѡ���б�������ҳ����\n"
		   End Select
		End If
		
		If FoundErr <> True Then
		   SqlItem = "Select * From KS_CollectItem Where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 2, 3
		
		   RsItem("LsString") = LsString
		   RsItem("LoString") = LoString
		   RsItem("ListPageType") = ListPageType
		   RsItem("ListStr") = ListStr
		   Select Case ListPageType
		   Case 0, 1
			  If ListPageType = 1 Then
				 RsItem("LPsString") = LPsString
				 RsItem("LPoString") = LPoString
				 RsItem("ListPageStr1") = ListPageStr1
			  End If
		   Case 2
			  RsItem("ListPageStr2") = ListPageStr2
			  RsItem("ListPageID1") = ListPageID1
			  RsItem("ListPageID2") = ListPageID2
		   Case 3
			  RsItem("ListPageStr3") = ListPageStr3
		   End Select
		   RsItem.Update
		   RsItem.Close
		   Set RsItem = Nothing
		End If
		End Sub
		
		
		'==================================================
		'��������GetTest
		'��  �ã�����
		'��  ������
		'==================================================
		Sub GetTest()
		   SqlItem = "Select * From KS_CollectItem Where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "�����������ĿID����Ϊ��\n"
		   Else
		      ChannelID=RsItem("ChannelID")
			  '--------�б�����ͼ----
			  ThumbType=RsItem("ThumbType")
			  TbsString=RsItem("TbsString")
			  TboString=RsItem("TboString")
			  '----------------------
			  LoginType = RsItem("LoginType")
			  LoginUrl = RsItem("LoginUrl")
			  LoginPostUrl = RsItem("LoginPostUrl")
			  LoginUser = RsItem("LoginUser")
			  LoginPass = RsItem("LoginPass")
			  LoginFalse = RsItem("LoginFalse")
			  ListStr = RsItem("ListStr")
			  LsString = RsItem("LsString")
			  LoString = RsItem("LoString")
			  ListPageType = RsItem("ListPageType")
			  LPsString = RsItem("LPsString")
			  LPoString = RsItem("LPoString")
			  ListPageStr1 = RsItem("ListPageStr1")
			  ListPageStr2 = RsItem("ListPageStr2")
			  ListPageID1 = RsItem("ListPageID1")
			  ListPageID2 = RsItem("ListPageID2")
			  ListPageStr3 = RsItem("ListPageStr3")
			  HsString = RsItem("HsString")
			  HoString = RsItem("HoString")
			  HttpUrlType = RsItem("HttpUrlType")
			  HttpUrlStr = RsItem("HttpUrlStr")
			  CharsetCode =RsItem("CharsetCode")
		   End If
		   RsItem.Close
		   Set RsItem = Nothing

		   If LsString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "���б�ʼ��ǲ���Ϊ�գ�\n"
		   End If
		   If LoString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "���б������ǲ���Ϊ�գ�\n"
		   End If
		   If ListPageType = 0 Or ListPageType = 1 Then
			  If ListStr = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���б�����ҳ����Ϊ�գ�\n"
			  End If
			  If ListPageType = 1 Then
				 If LPsString = "" Or LPoString = "" Then
					FoundErr = True
					ErrMsg = ErrMsg & "��������ҳ��ʼ/������ǲ���Ϊ�գ�\n"
				 End If
				 If ListPageStr1 <> "" And Len(ListPageStr1) < 15 Then
					FoundErr = True
					ErrMsg = ErrMsg & "��������ҳ�����������ò���ȷ(�����ջ����ַ�>15��)��\n"
				 End If
			  End If
		   ElseIf ListPageType = 2 Then
			  If ListPageStr2 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "����������ԭ�ַ�������Ϊ�գ�\n"
			  End If
			  If IsNumeric(ListPageID1) = False Or IsNumeric(ListPageID2) = False Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���������ɵķ�Χ����ȷ����\n"
			  Else
				 ListPageID1 = CLng(ListPageID1)
				 ListPageID2 = CLng(ListPageID2)
				 If ListPageID1 = 0 And ListPageID2 = 0 Then
					FoundErr = True
					ErrMsg = ErrMsg & "���������ɵķ�Χ����ȷ��\n"
				 End If
			  End If
		   ElseIf ListPageType = 3 Then
			  If ListPageStr3 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "��������ҳ����Ϊ�գ�\n"
			  End If
		   Else
			  FoundErr = True
			  ErrMsg = ErrMsg & "�����������ѡ��������ҳ����\n"
		   End If
		 
		   If LoginType = 1 Then
			  If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���뽫��¼��Ϣ��д����\n"
			  End If
		   End If
		
		   If FoundErr <> True Then
			  Select Case ListPageType
			  Case 0, 1
				 ListUrl = ListStr
			  Case 2
				 ListUrl = ListStr
				 'ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1))
			  Case 3
				 If InStr(ListPageStr3, "|") > 0 Then
					ListUrl = Left(ListPageStr3, InStr(ListPageStr3, "|") - 1)
				 Else
					ListUrl = ListPageStr3
				 End If
			  End Select
		
			  If LoginType = 1 Then
				 LoginData = KMCObj.UrlEncoding(LoginUser & "&" & LoginPass)
				 LoginResult = KMCObj.PostHttpPage(LoginUrl, LoginPostUrl, LoginData)
				 If InStr(LoginResult, LoginFalse) > 0 Then
					FoundErr = True
					ErrMsg = ErrMsg & "���¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�\n"
				 End If
			  End If
		   End If

		   If FoundErr <> True Then
			  ListCode = KMCObj.GetHttpPage(ListUrl,CharsetCode)
			  If ListCode <> "Error" Then
				 If ListPageType = 1 Then
					ListPageNext = KMCObj.GetPage(ListCode, LPsString, LPoString, False, False)
					If ListPageNext <> "Error" Then
					   If ListPageStr1 <> "" Then
						  ListPageNext = Replace(ListPageStr1, "{$ID}", ListPageNext)
					   Else
						  ListPageNext = KMCObj.DefiniteUrl(ListPageNext, ListUrl)
					   End If
					End If
				 End If

				 ListCode = KMCObj.GetBody(ListCode, LsString, LoString, False, False)
				 If ListCode = "Error" Then
					FoundErr = True
					ErrMsg = ErrMsg & "���ڽ�ȡ�б�ʱ��������\n"
				 End If
			  Else
				 FoundErr = True
				 ErrMsg = ErrMsg & "���ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������\n"
			  End If
		   End If
		End Sub	
End Class
%> 

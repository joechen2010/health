<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
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
Set KSCls = New Collect_ItemAttribute
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemAttribute
        Private KS,KSCls
		Private KMCObj
		Private ConnItem
		Private SqlItem, RsItem, Action, FoundErr, ErrMsg
		Private ItemID, ItemName, ChannelID, ClassID, SpecialID
		Private PaginationType, MaxCharPerPage, ReadLevel, Stars, ReadPoint, Hits, UpDateType, UpDateTime, PicNews, Rolls
		Private Comment, Recommend, Popular, FnameType, TemplateID
		Private Script_Iframe, Script_Object, Script_Script, Script_Div, Script_Class, Script_Span, Script_Img, Script_Font, Script_A, Script_Html, Script_Table, Script_Tr, Script_Td
		Private CollecListNum, CollecNewsNum,RepeatInto, IntoBase, BeyondSavePic, CollecOrder, Verific, InputerType, Inputer, EditorType, Editor, ShowComment
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
			FoundErr = False
			ItemID = Trim(Request("ItemID"))
			Action = Trim(Request("Action"))
			Verific = 1
			Recommend=1
			IntoBase = 1
			
			If ItemID = "" Then
			   FoundErr = True
			   ErrMsg = ErrMsg & "<br><li>����������ĿID����Ϊ�գ�</li>"
			Else
			   ItemID = CLng(ItemID)
			End If
			
			If FoundErr <> True Then
				  Call GetTest
			End If
			If FoundErr <> True Then
			   Call KMCObj.GetClassList
			   Call Main
			End If
			Response.Write "<script>"
			Response.Write "function SelectClass(ChannelID)"
			Response.Write "{"
			Response.Write " document.all.ClassArea.innerHTML='<select ID=""ClassID"" name=""ClassID"" style=""Width:200"">'+ClassArr[ChannelID]+'</select>';"
			Response.Write "}"
			Response.Write "function CheckForm(myform)"
			Response.Write "{ if (myform.ItemName.value=='')"
			Response.Write "  {"
			Response.Write "   alert('��������Ŀ����');"
			Response.Write "   myform.ItemName.focus();"
			Response.Write "   return false;"
			Response.Write "  }"
			Response.Write " if (myform.ChannelID.value=='0')"
			Response.Write "  {"
			Response.Write "    alert('��ѡ��ϵͳģ��!');"
			Response.Write "    return false;"
			Response.Write "  }"
			 Response.Write "  if (myform.ClassID.value=='0')"
			Response.Write "  {"
			Response.Write "    alert('��ѡ����Ŀ!');"
			Response.Write "    return false;"
			Response.Write "  }"
			Response.Write "  if (myform.WebName.value=='')"
			 Response.Write " {"
			Response.Write "   alert('��������վ����');"
			Response.Write "   myform.WebName.focus();"
			Response.Write "   return false;"
			Response.Write "  }"
			 Response.Write "return true;"
			Response.Write "}"
			Response.Write "</script>"
			End Sub
			
			Sub Main()
			
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<title>�ɼ�ϵͳ</title>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
			Response.Write "<script language='JavaScript' src='../../KS_Inc/common.js'></script>"
			Response.Write "<style type=""text/css"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write ".STYLE1 {color: #FF0000}" & vbCrLf
			Response.Write ".STYLE4 {color: #0000FF}" & vbCrLf
			Response.Write "-->" & vbCrLf
			Response.Write "</style>"
			Response.Write "</head>"
			Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(6,ItemID) &"</div>"

			Response.Write "<table align='center' width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"">"
			Response.Write "<form method=""post"" action=""Collect_ItemSuccess.asp"" name=""myform"">"
			 Response.Write " <br>"
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""20"" width=""20%"" align=""right"" class='clefttitle'>��Ŀ���ƣ�</td>"
			 Response.Write "     <td><input name='ItemName' type='text' id='ItemName' value='" & ItemName & "' size='27' maxlength='30'></td>"
			 Response.Write "   </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> ����ģ�ͣ�</td>"
			   Response.Write "   <td height=""20""><input ID=""ChannelID"" name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"" style=""Width:200"">"
			   Response.Write "   <font color=""red"">" & KS.C_S(ChannelID,1) & "</font>     </td>"
			   Response.Write " </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> ������Ŀ��</td>"
			  Response.Write "    <td height=""20"" ID=""ClassArea""><select name=""ClassID"" ID=""ClassID"" style=""Width:200"">" & Replace(KS.LoadClassOption(ChannelID),"value='" & ClassID & "'","value='" & ClassID &"' selected") & "</select>      </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg' style=""display:none"">"
			  Response.Write "    <td height=""20"" width=""20%"" align=""right"" class='clefttitle'> ����ר�⣺</td>"
			  Response.Write "    <td><input type=""hidden"" value=""0"" name=""specialid"">"
			  'call KMCObj.Collect_ShowSpecial_Option(ChannelID,SpecialID)
			  Response.write"     </td>"
			  Response.Write "   </tr>"
			  Response.Write "      <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>��¼��ʱ�䣺</td>"
			   Response.Write "   <td><input name=""UpdateType"" type=""radio"" value=""0"" "
			   If UpDateType = 0 Then Response.Write "checked"
			   Response.Write ">��ǰʱ��"
			   Response.Write "    &nbsp;<input name=""UpdateType"" type=""radio"" value=""1"" "
			   If UpDateType = 1 Then Response.Write "checked"
			   Response.Write ">��ǩ�е�ʱ��"
			   Response.Write "    &nbsp;<input name=""UpdateType"" type=""radio"" value=""2"" "
			   If UpDateType = 2 Then Response.Write "checked"
			   Response.Write ">�Զ��壺"
			   Response.Write "    <input name=""UpdateTime"" type=""text"" value=""" & UpDateTime & """>"
			   Response.Write "   ��</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>¼��Ա��</td>"
			   Response.Write "   <td><input name=""InputerType"" type=""radio"" value=""0"" "
			   If InputerType = 0 Then Response.Write "checked"
			   Response.Write ">��ǰ�û�"
			   Response.Write "    &nbsp;<input name=""InputerType"" type=""radio"" value=""1"" "
			   If InputerType = 1 Then Response.Write "checked"
			   Response.Write ">ָ���û�"
				Response.Write "   <input name=""Inputer"" type=""text"" value=""" & Inputer & """>"
				Response.Write "  ��</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
				Response.Write "  <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>���α༭��</td>"
			   Response.Write "   <td><input name=""EditorType"" type=""radio"" value=""0"" "
			   If EditorType = 0 Then Response.Write "checked"
			   Response.Write ">��ǰ�û�"
			   Response.Write "    &nbsp;<input name=""EditorType"" type=""radio"" value=""1"" "
			   If EditorType = 1 Then Response.Write "checked"
			   Response.Write ">ָ���û�"
			   Response.Write "    <input name=""Editor"" type=""text"" value=""" & Editor & """>"
			   Response.Write "   ��</td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>�������ԣ�</td>"
			   Response.Write "   <td>"
			  
			   Response.Write "     <input name=""Rolls"" type=""checkbox"" value=""1"" "
			   If Rolls = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     �� ��"
			   Response.Write "     <input name=""Comment"" type=""checkbox"" value=""1"" "
			   If Comment = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     ��������"
			   Response.Write "     <input name=""Recommend"" type=checkbox value=""1"" "
			   If Recommend = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     �� ��"
			   Response.Write "     <input name=""Popular"" type=""checkbox"" value=""1"" "
			   If Popular = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     �� ��"
			   
			   'Response.Write "     <input name=""Verific"" type=""checkbox"" value=""1"" "
			   'If Verific = 1 Then Response.Write "checked"
			   'Response.Write "checked"
			   'Response.Write ">�����"
			Response.Write "</td>"
			 Response.Write "   </tr>"
			 Response.Write "   <tr class='tdbg' style='display:none'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>��������ʾ�������ӣ�</td>"
			 Response.Write "     <td>"
			 Response.Write "        <input name=""ShowComment"" type=""radio"" id=""ShowComment"" value=""1"" "
			 If ShowComment = 1 Then Response.Write "Checked"
			 Response.Write ">"
			 Response.Write "        ��ʾ"
			 Response.Write "        <input name=""ShowComment"" type=""radio"" id=""ShowComment"" value=""0"" "
			 If ShowComment = 0 Then Response.Write "Checked"
			 Response.Write ">         ����ʾ      </td>"
			 Response.Write "   </tr>"
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>���ݷ�ҳ��ʽ��</td>"
			 Response.Write "     <td><select name=""PaginationType"">"
			 Response.Write "           <option value=""0"" "
			 If PaginationType = 0 Then Response.Write "selected"
			 Response.Write ">����ҳ</option>"
			 Response.Write "           <option value=""1"" "
			 If PaginationType = 1 Then Response.Write "selected"
			 Response.Write ">�Զ���ҳ</option>"
			 Response.Write "           <option value=""2"" "
			 If PaginationType = 2 Then Response.Write "selected"
			 Response.Write ">����ԭ�ķ�ҳ</option>"
			 Response.Write "         </select>"
			  Response.Write "      �Զ���ҳʱ��ÿҳ��Լ�ַ���������HTML��ǣ���"
			  Response.Write "<input name=""MaxCharPerPage"" type=""text"" value=""" & MaxCharPerPage & """ size=""8"" maxlength=""8"">      "
			  Response.Write "  </td></tr>"
			  
			  
			  Response.Write "  <tr class='tdbg' style=""display:none"">"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>�Ķ��ȼ���</td>"
			 Response.Write "     <td><input type='hidden' value='0' name='ReadLevel'></td>"
			 Response.Write "   </tr>"
			 Response.Write "     <tr  class='tdbg' style=""display:none"">"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>�Ķ�������</td>"
			 Response.Write "     <td><input name='ReadPoint' type='text' id='ReadPoint' value='" & ReadPoint & "' size='5' maxlength='3'>"
			 Response.Write "     <font color='#0000FF'>�������0�����û��Ķ�������ʱ��������Ӧ�����������οͺ͹���Ա��Ч��</font>      </td>"
			 Response.Write "   </tr>"
			 
			 
			 
			 
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>���ֵȼ���</td>"
			 Response.Write "     <td><select name=""Stars"">"
			 Response.Write "           <option value=""������"" "
			 If Stars = "������" Then Response.Write "selected"
			 Response.Write ">������</option>"
			 Response.Write "           <option value=""�����"""
			 If Stars = "�����" Then Response.Write "selected"
			 Response.Write ">�����</option>"
			 Response.Write "           <option value=""����"" "
			 If Stars = "����" Then Response.Write "selected"
			 Response.Write ">����</option>"
			 Response.Write "           <option value=""���"" "
			 If Stars = "���" Then Response.Write "selected"
			 Response.Write ">���</option>"
			 Response.Write "           <option value=""��"" "
			 If Stars = "��" Then Response.Write "selected"
			 Response.Write ">��</option>"
			 Response.Write "         </select>      </td>"
			 Response.Write "   </tr>"
			  
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>�������ʼֵ��</td>"
			 Response.Write "     <td><input name=""Hits"" type=""text"" id=""Hits"" value=""" & Hits & """ size=""10"" maxlength=""10"">"
			 Response.Write "       <span class=""STYLE4"">�������������</span></td>"
			 Response.Write "   </tr>"
			  
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>������չ����</td>"
			 Response.Write "     <td>" & KSCls.GetFsoTypeStr(1)
			 Response.Write "   </td>"
			 Response.Write "   </tr>"
			 	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)

			  Response.Write "   <tr class='tdbg' style='display:none'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>��ģ�壺</td>"
			  Response.Write "    <td><input type='text' size='25' name='TemplateID' id='TemplateID' value='" & templateid & "'> <input type='button' name=""Submit"" class=""button"" value=""ѡ��ģ��..."" onClick=""OpenThenSetValue('../KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle='+escape('ѡ��ģ��')+'&CurrPath=" & Server.URLEncode(CurrPath) & "',450,350,window,TemplateID);"">"
			  Response.Write "  </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'>��ǩ���ˣ�</td>"
			  Response.Write "    <td>"
			  Response.Write "      <input name=""Script_Iframe"" type=""checkbox"" value=""yes"" "
			  If Script_Iframe = -1 Then Response.Write "checked"
			  Response.Write ">"
			  Response.Write "      Iframe"
			  Response.Write "      <input name=""Script_Object"" type=""checkbox"" value=""yes"" "
			  If Script_Object = -1 Then Response.Write "checked "
			  Response.Write "onclick='return confirm(""ȷ��Ҫѡ��ñ�����⽫ɾ�������е�����Object��ǣ���������¸������е����ж���������ɾ����"");'>"
			  Response.Write "      Object"
			  Response.Write "      <input name=""Script_Script"" type=""checkbox"" value=""yes"" "
			  If Script_Script = -1 Then Response.Write "checked"
			  Response.Write ">"
			   Response.Write "     Script"
			   Response.Write "     <input name=""Script_Div"" type=""checkbox""  value=""yes"" "
			   If Script_Div = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Div"
			   Response.Write "     <input name=""Script_Class"" type=""checkbox""  value=""yes"" "
			   If Script_Class = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Class"
			   Response.Write "     <input name=""Script_Table"" type=""checkbox""  value=""yes"" "
			   If Script_Table = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Table"
			   Response.Write "     <input name=""Script_Tr"" type=""checkbox""  value=""yes"" "
			   If Script_Tr = -1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     Tr"
			   Response.Write "     <br>"
			   Response.Write "     <input name=""Script_Span"" type=""checkbox""  value=""yes"" "
			   If Script_Span = -1 Then Response.Write "checked"
			   Response.Write ">"
				Response.Write "    Span&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_Img"" type=""checkbox"" value=""yes"" "
				If Script_Img = -1 Then Response.Write "checked"
				Response.Write ">"
				Response.Write "    Img&nbsp;&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_Font"" type=""checkbox""  value=""yes"" "
				If Script_Font = -1 Then Response.Write "checked"
				Response.Write ">"
				 Response.Write "   Font&nbsp;&nbsp;"
				Response.Write "    <input name=""Script_A"" type=""checkbox"" value=""yes"" "
				If Script_A = -1 Then Response.Write "checked"
				Response.Write ">"
				 Response.Write "   A&nbsp;&nbsp;"
				 Response.Write "   <input name=""Script_Html"" type=""checkbox"" value=""yes"" "
				 If Script_Html = -1 Then Response.Write "checked"
				 Response.Write " onclick='return confirm(""ȷ��Ҫѡ��ñ�����⽫ɾ�������е�����Html��ǣ���������¸����µĿ��Ķ��Խ��ͣ�"");'>"
				 Response.Write "   Html&nbsp;"
				 Response.Write "   <input name=""Script_Td"" type=""checkbox""  value=""yes"" "
				 If Script_Td = -1 Then Response.Write "checked"
				 Response.Write ">"
				Response.Write "    Td      </td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> �б���ȣ�</td>"
			   Response.Write "   <td>"
			   Response.Write "     <input name=""CollecListNum"" type=""text"" id=""CollecListNum"" value=""" & CollecListNum & """ size=""10"" maxlength=""10"">&nbsp;&nbsp;&nbsp;"
			   Response.Write "     <font color='#0000FF'>0Ϊ���е��б�</font></td>"
			   Response.Write " </tr>"
			   Response.Write " <tr class='tdbg'>"
			   Response.Write "   <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> �ɼ���Ϣ������</td>"
			   Response.Write "   <td>"
			  Response.Write "      <input name=""CollecNewsNum"" type=""text"" id=""CollecNewsNum"" value=""" & CollecNewsNum & """ size=""10"" maxlength=""10"">"
			  Response.Write "      &nbsp;&nbsp;&nbsp;"
			  Response.Write "      <font color='#0000FF'>0Ϊ���е�����<span lang=""en-us"">(</span>ÿһ�б��������������<span lang=""en-us"">)</span></font></td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> �ɼ�ѡ�</td>"
			  Response.Write "    <td>"
			   Response.Write "     <input name=""PicNews"" type=""checkbox"" value=""1"" "
			   If PicNews = 1 Then Response.Write "checked"
			   Response.Write ">"
			   Response.Write "     �Զ�ת��ΪͼƬ����"

			 Response.Write "       <input name=""RepeatInto"" type=""checkbox"" value=""1"" "
			  If RepeatInto="1" Then Response.Write "checked"
			  Response.Write ">�ظ���¼���"
			  Response.Write "      <input name=""BeyondSavePic"" type=""checkbox"" value=""1"" "
			  If BeyondSavePic = 1 Then Response.Write "checked"
			  If KMCObj.IsObjInstalled(KS.Setting(99)) = False Then Response.Write "disabled"
			  Response.Write ">"
			  Response.Write "      ����ͼƬ"
			  Response.Write "      <input name=""CollecOrder"" type=""checkbox"" value=""yes"" "
			  If CollecOrder = -1 Then Response.Write "checked"
			  Response.Write ">"
			  Response.Write "      ����ɼ�        </td>"
			  Response.Write "  </tr>"
			  Response.Write "  <tr class='tdbg'>"
			  Response.Write "    <td height=""22"" width=""20%"" align=""right"" class='clefttitle'> ���ѡ�</td>"
			  Response.Write "    <td>"

			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""0"" "
			  If IntoBase = 0 Then Response.Write "checked"
			  Response.Write ">  ��ֱ����⣬��Ҫ���(<font color=red>���Ƽ�</font>)<br/>"
			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""1"" "
			  If IntoBase = 1 Then Response.Write "checked"
			  Response.Write ">  ����д�������ݿ�<br/>"
			  Response.Write "      <input name=""IntoBase"" type=""radio"" value=""2"" "
			  If IntoBase = 2 Then Response.Write "checked"
			  Response.Write ">  ����д�������ݿⲢֱ����������ҳ<br/>"
	
			  Response.Write "              </td>"
			  Response.Write "  </tr>"
			  
			 Response.Write " <tr class='tdbg'>"
			  Response.Write "  <td height=""30"" width=""20%"" align=""right""></td>"
			  Response.Write "  <td align-""center""><center>"
			  Response.Write "     <input type=""hidden"" value=""" & ItemID & """ name=""ItemID"">"
			  Response.Write "     <input type=""submit""  class='button' value="" ��&nbsp;&nbsp;�� "" name=""submit"">  </center>      </td>"
			  Response.Write "  </tr>"
			Response.Write "</form>"
			Response.Write "</table>"
			Response.Write "</body>"
			Response.Write "</html>"
			End Sub
			Sub GetTest()
			   SqlItem = "Select top 1 * From KS_CollectItem Where ItemID=" & ItemID
			   Set RsItem = Server.CreateObject("adodb.recordset")
			   RsItem.Open SqlItem, ConnItem, 1, 1
			   If RsItem.EOF And RsItem.BOF Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "<br><li>���������Ҳ�������Ŀ</li>"
			   Else
				  ItemName = RsItem("ItemName")
				  ChannelID = RsItem("ChannelID")
				  ClassID = RsItem("ClassID")
				  SpecialID = RsItem("SpecialID")
				  PaginationType = RsItem("PaginationType")
				  MaxCharPerPage = RsItem("MaxCharPerPage")
				  ReadLevel = RsItem("ReadLevel")
				  Stars = RsItem("Stars")
				  ReadPoint = RsItem("ReadPoint")
				  Hits = RsItem("Hits")
				  UpDateType = RsItem("UpdateType")
				  UpDateTime = RsItem("UpdateTime")
				  PicNews = RsItem("PicNews")
				  Rolls = RsItem("Rolls")
				  Comment = RsItem("Comment")
				  Recommend = RsItem("Recommend")
				  Popular = RsItem("Popular")
				  FnameType = RsItem("FnameType")
				  TemplateID = RsItem("TemplateID")
				  Script_Iframe = RsItem("Script_Iframe")
				  Script_Object = RsItem("Script_Object")
				  Script_Script = RsItem("Script_Script")
				  Script_Div = RsItem("Script_Div")
				  Script_Class = RsItem("Script_Class")
				  Script_Span = RsItem("Script_Span")
				  Script_Img = RsItem("Script_Img")
				  Script_Font = RsItem("Script_Font")
				  Script_A = RsItem("Script_A")
				  Script_Html = RsItem("Script_Html")
				  IntoBase = RsItem("IntoBase")
				  RepeatInto = RsItem("RepeatInto")
				  BeyondSavePic = RsItem("BeyondSavePic")
				  CollecOrder = RsItem("CollecOrder")
				  Verific = RsItem("Verific")
				  CollecListNum = RsItem("CollecListNum")
				  CollecNewsNum = RsItem("CollecNewsNum")
				  InputerType = RsItem("InputerType")
				  Inputer = RsItem("Inputer")
				  EditorType = RsItem("EditorType")
				  Editor = RsItem("Editor")
				  ShowComment = RsItem("ShowComment")
				  Script_Table = RsItem("Script_Table")
				  Script_Tr = RsItem("Script_Tr")
				  Script_Td = RsItem("Script_Td")
			   End If
			   RsItem.Close
			   Set RsItem = Nothing
			End Sub
End Class
%> 

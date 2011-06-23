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
Set KSCls = New Collect_ItemModify
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify
        Private KS
		Private KMCObj
		Private ConnItem
		Private SqlItem, RsItem, FoundErr, ErrMsg
		
		Private ItemID, ItemName, WebName, WebUrl, ChannelID, ClassID, SpecialID, LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, ItemDemo,CharsetCode
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
		If ItemID = "" Then
		  ChannelID=KS.ChkClng(KS.G("ChannelID")):CharsetCode="gb2312"
		Else
		   ItemID = CLng(ItemID)
		   SqlItem = "select ItemID,ItemName,CharsetCode,WebName,WebUrl,ChannelID,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse,ItemDemo From KS_CollectItem where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "<br><li>��������û���ҵ�����Ŀ��</li>"
		   Else
			  ItemName = RsItem("ItemName")
			  CharsetCode=RsItem("CharsetCode")
			  ItemDemo = RsItem("ItemDemo")
			  WebName = RsItem("WebName")
			  WebUrl = RsItem("WebUrl")
			  ChannelID = RsItem("ChannelID")
			  ClassID = RsItem("ClassID")
			  SpecialID = RsItem("SpecialID")
			  LoginType = RsItem("LoginType")
			  LoginUrl = RsItem("LoginUrl")
			  LoginPostUrl = RsItem("LoginPostUrl")
			  LoginUser = RsItem("LoginUser")
			  LoginPass = RsItem("LoginPass")
			  LoginFalse = RsItem("LoginFalse")
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		End If
		
		If FoundErr = True Then
		  Call KS.AlertHistory(ErrMsg,-1)
		Else
		   'Call KMCObj.GetClassList
		   Call Main
		End If
		
		End Sub
		Sub Main()
		With KS
		  .echo "<html>"
		  .echo "<head>"
		  .echo "<title>�ɼ�ϵͳ</title>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		  .echo "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		  .echo "<script src=""../../ks_inc/jquery.js""></script>"
		  .echo "</head>"
		  .echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		  .echo "<div class=""topdashed"">"
		  .echo  KMCObj.GetItemLocation(1,ItemID)
		  .echo "</div>"
		  .echo "<br>"
		  .echo "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""ctable"" >"
		  .echo "<form method=""post"" action=""Collect_ItemModify2.asp"" name=""myform""  onSubmit=""return(CheckForm(this))"">"
		  .echo "    <tr class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'>��Ŀ���ƣ�</td>"
		   .echo "     <td width=""796"">"
		   .echo "     <input name=""ItemName"" type=""text"" size=""27"" maxlength=""30"" value=""" & ItemName & """>&nbsp;&nbsp;<font color=red>*</font>�磺����������������</td>"
		    .echo "  </tr>"
		    .echo "  <tr class='tdbg'>"
		    .echo "    <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ����ģ�ͣ�</td>"
		    .echo "    <td width=""796""><select ID=""ChannelID"" name=""ChannelID"" onChange=""SelectClass(this.value)"" style=""Width:200"">"
		    .echo KMCObj.Collect_ShowChannel_Option(ChannelID) & "</select>      </td>"
		    .echo "  </tr>"
		   .echo "   <tr class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ������Ŀ��</td>"
		   .echo "     <td width=""796"" ID=""ClassArea""><select name=""ClassID"" ID=""ClassID"" style=""Width:200"">"
		   .echo Replace(KS.LoadClassOption(ChannelID),"value='" & ClassID & "'","value='" & ClassID &"' selected") & "</select>      </td>"
		   .echo "   </tr>"
		  .echo "    <tr style=""display:none"">"
		  .echo "      <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ����ר�⣺</td>"
		  .echo "      <td width=""796""><input type=""hidden"" value=""0"" name=""specialid"">"
		'call KMCObj.Collect_ShowSpecial_Option(1,0)
		  .echo "      </td>"
		   .echo "   </tr>"
		   .echo "   <tr  class='tdbg' class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ��վ���ƣ�</td>"
		   .echo "     <td width=""796"">"
		    .echo "      <input name=""WebName"" type=""text"" size=""27"" maxlength=""30"" value=""" & WebName & """>      </td>"
		    .echo "  </tr>"
		     .echo "   <tr  class='tdbg' class='tdbg'>"
		   .echo "     <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ���뷽ʽ��</td>"
		   .echo "     <td width=""796"">"
		   .echo " <select name=""CharsetCode"">"
	       .echo " <option value='auto'>�Զ����</option>"
	       .echo " <option value='utf-8' "
		 if CharsetCode="utf-8" then   .echo("selected")
		   .echo " >utf-8</option>"
	       .echo "<option value='gb2312' "
		 if CharsetCode="gb2312" then   .echo("selected")
		   .echo ">gb2312</option>"
	       .echo " </select>"
	        .echo "   </td>"
		    .echo "  </tr>"
		    .echo "  <tr class='tdbg' style=""display:none"">"
		    .echo "    <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ��վ��ַ��</td>"
		    .echo "    <td width=""796""><input name=""WebUrl"" type=""text"" size=""49"" maxlength=""150"" value=""" & WebUrl & """>      </td>"
		    .echo "  </tr>"
		    .echo " <tr class='tdbg'>"
		    .echo "    <td width=""20%"" height=""25"" align=""center"" class='clefttitle'> ��վ��¼��</td>"
		    .echo "    <td>"
		    .echo "      <input type=""radio"" value=""0"" name=""LoginType"" "
		  If LoginType = 0 Then   .echo "checked"
		    .echo " onClick=""Login.style.display='none'"">����Ҫ��¼<span lang=""en-us"">&nbsp;"
		    .echo "      </span>"
		    .echo "      <input type=""radio"" value=""1"" name=""LoginType"" "
		  If LoginType = 1 Then   .echo "checked"
		    .echo " onClick=""Login.style.display=''"">���ò���      </td>"
		    .echo "   </tr>"
		    .echo " <tr  class='tdbg' id=""Login"""
		  If LoginType = 0 Then   .echo " style=""display:none"" "
		   .echo "      ><td width=""20%"" height=""25"" align=""center""> ��¼������</td>"
		   .echo "     <td>"
		   .echo "       ��¼��ַ��<input name=""LoginUrl"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginUrl & """><br>"
		   .echo "       �ύ��ַ��<input name=""LoginPostUrl"" type=""text"" size=""40"" maxlength=""150"" value=""" & LoginPostUrl & """><br>"
		   .echo "       �û�������<input name=""LoginUser"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginUser & """><br>"
		   .echo "       ���������<input name=""LoginPass"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginPass & """><br>"
		   .echo "       ʧ����Ϣ��<input name=""LoginFalse"" type=""text"" size=""30"" maxlength=""150"" value=""" & LoginFalse & """></td>"
		   .echo "   </tr>"
		   .echo "   <tr class='tdbg'>"
		    .echo "    <td  width=""20%"" height=""25"" align=""center"" class='clefttitle'>��Ŀ��ע��</td>"
		   .echo "     <td width=""796""><textarea name=""ItemDemo"" cols=""49"" rows=""5"">" & ItemDemo & "</textarea></td>"
		   .echo "   </tr>"
		    .echo "  <tr class='tdbg'>"
		    .echo "    <td height=""35"" colspan=""2"" align=""center"">"
		    .echo "      <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
		     .echo "     <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
		     .echo "     <input class='button' name=""Cancel"" type=""button"" id=""Cancel"" value="" ��&nbsp;&nbsp;�� "" onClick=""window.location.href='javascript:history.back();'"">"
		     .echo "     &nbsp;"
		     .echo "   <input  class='button' type=""submit"" name=""Submit"" value=""��&nbsp;һ&nbsp;��""></td>"
		     .echo " </tr>"
		  .echo "</form>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		  .echo "<script>"
		  .echo "function SelectClass(ChannelID)"
		  .echo "{"
		  .echo " if (ChannelID!=0){"
	      .echo "$(parent.frames[""FrameTop""].document).find(""#ajaxmsg"").toggle();" 
	      .echo "$.get(""../../plus/ajaxs.asp"",{action:""GetClassOption"",channelid:ChannelID},function(data){"
	      .echo "$(parent.frames[""FrameTop""].document).find(""#ajaxmsg"").toggle();"
	      .echo "$(""select[name=ClassID]"").empty();"
		  .echo "$(""select[name=ClassID]"").append(unescape(data));"
	      .echo " });"
	      .echo "}"
		  .echo "}"
		  .echo "function CheckForm(myform)"
		  .echo "{ if (myform.ItemName.value=='')"
		  .echo "  {"
		  .echo "   alert('��������Ŀ����');"
		  .echo "   myform.ItemName.focus();"
		  .echo "   return false;"
		  .echo "  }"
		   .echo "if (myform.ChannelID.value=='0')"
		    .echo "{"
		   .echo "   alert('��ѡ��ϵͳģ��!');"
		   .echo "   return false;"
		   .echo " }"
		   .echo "  if (myform.ClassID.value=='0')"
		   .echo " {"
		   .echo "   alert('��ѡ����Ŀ!');"
		   .echo "   return false;"
		   .echo " }"
		   .echo " if (myform.WebName.value=='')"
		   .echo " {"
		   .echo "  alert('��������վ����');"
		   .echo "  myform.WebName.focus();"
		   .echo "  return false;"
		   .echo " }"
		 
		   .echo "return true;"
		  .echo "}"
		  .echo "</script>"
		 End With
		End Sub
End Class
%> 

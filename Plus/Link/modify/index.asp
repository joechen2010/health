<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New FriendLinkModify
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkModify
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Dim LinkID, SiteName, WebMaster, Email, PassWord, Url, LinkType, Logo, Descript, FolderID
			
			LinkID = KS.ChkClng(KS.S("LinkID"))
			
			If LinkID=0 Then
			 Response.Write ("<script>alert('�������ݳ���!');history.back();</script>")
			 response.end
			End If
			   Dim RSObj
			  Set RSObj = Conn.Execute("Select * From KS_Link Where LinkID=" & LinkID)
			  If Not RSObj.EOF Then
				 SiteName = Trim(RSObj("SiteName"))
				 WebMaster = Trim(RSObj("WebMaster"))
				 Email = Trim(RSObj("Email"))
				 Url = Trim(RSObj("Url"))
				 Logo = Trim(RSObj("Logo"))
				 LinkType = Trim(RSObj("LinkType"))
				 Descript = Trim(RSObj("Description"))
				 FolderID = RSObj("FolderID")
			  End If
			   RSObj.Close
			   Set RSObj = Nothing
			Response.Write "<html>" & vbCrLf
			Response.Write "<head>" & vbCrLf
			Response.Write "<title>�޸���������</title>" & vbCrLf
			Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
			Response.Write "<link href=""../../../images/style.css"" rel=""stylesheet"" type=""text/css"">" & vbCrLf
			Response.Write "</head>" & vbCrLf
			Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf
			'Response.Write "<br>" & vbCrLf
			Response.Write "  <table width=""778"" height=""100%"" bgcolor=""#ffffff"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
				  
			Response.Write "    <td align=""center""><br>" & vbCrLf
			Response.Write "        <table width=""500"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""table_border"">" & vbCrLf
			Response.Write "          <tr class=""link_table_title"">" & vbCrLf
			Response.Write "          <td>�޸���������</td>" & vbCrLf
			Response.Write "          </tr>"
			Response.Write "          <tr><td>" & vbCrLf
				 
			Response.Write "  <form action=""ModifySave.asp"" name=""LinkForm"" method=""post"">" & vbCrLf
			Response.Write "   <input name=""LinkID"" type=""hidden"" value=""" & LinkID & """>" & vbCrLf
			Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td>" & vbCrLf
			Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td width=""20%"" height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">��վ����</div></td>" & vbCrLf
			Response.Write "            <td width=""542"" height=""25"">" & vbCrLf
			Response.Write ("<input name=""SiteName"" class=""textbox"" value=""" & SiteName & """ type=""text"" id=""SiteName"" size=""38"" >")
			Response.Write "              <font color=""red"">*</font></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">�������</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <select Name=""FolderID"" >" & vbCrLf
						   
						Dim GRS
						Set GRS = Conn.Execute("Select FolderID,FolderName From KS_LinkFolder Order BY AddDate Desc")
						 Do While Not GRS.EOF
						   If CStr(FolderID) = CStr(GRS(0)) Then
							Response.Write ("<Option value=" & GRS(0) & " selected>" & GRS(1) & "</OPTION>")
						   Else
							Response.Write ("<Option value=" & GRS(0) & ">" & GRS(1) & "</OPTION>")
						   End If
						   GRS.MoveNext
						 Loop
						 GRS.Close
						 Set GRS = Nothing
					   
			 Response.Write "             </Select> </td>" & vbCrLf
			 Response.Write "         </tr>"
			 Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">��վվ��</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""WebMaster"" class=""textbox"" type=""text"" size=""38"" value=""" & WebMaster & """ > <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">վ������</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""Email"" class=""textbox"" type=""text"" size=""38"" value=""" & Email & """ ></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">ԭ������</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""OriPassWord"" class=""textbox"" type=""password"" size=""42"" > <font color=""red"">* ��������</font> </td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">��������</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""PassWord"" class=""textbox"" type=""password"" size=""42"" > <font color=green>�����޸ģ��뱣��Ϊ��</font> </td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">ȷ������</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "            <input name=""ConPassWord"" class=""textbox"" type=""password"" size=""42"" > </td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "            <td height=""25"" align=""center"">��վ��ַ</td>" & vbCrLf
			Response.Write "            <td height=""25""><input class=""textbox"" name=""Url"" type=""text""  value=""" & Url & """ id=""Url"" size=""38""> <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "            <td height=""25"" align=""center"">��������</td>" & vbCrLf
			Response.Write "            <td height=""25"">"
						 
						 If Trim(LinkType) = "1" Then
							  Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0""> �������� ")
							  Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"" checked>  LOGO���� ")
						  Else
							  Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0"" checked> �������� ")
							  Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"">  LOGO���� ")
						  End If
					   
			Response.Write "             </td>" & vbCrLf
			Response.Write "          </tr>"
			If Trim(LinkType) = "1" Then
			Response.Write "          <tr ID=""LinkArea"">" & vbCrLf
			Else
			Response.Write ("         <tr Style=""display:none"" ID=""LinkArea"">") & vbCrLf
			End If
			Response.Write "            <td height=""25"" align=""center"">Logo ��ַ</td>" & vbCrLf
			Response.Write "            <td height=""25""><input name=""Logo"" class=""textbox"" type=""text""  value=""" & Logo & """ id=""Logo"" size=""38""> <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">"
			Response.Write "              <div align=""center"">��վ���</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <textarea name=""Description"" rows=""6"" id=""Description"" style=""width:80%;border-style: solid; border-width: 1"">" & Descript & "</textarea></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "        </table>" & vbCrLf
			Response.Write "       </td>"
			Response.Write "    </tr>" & vbCrLf
			Response.Write "    </table>" & vbCrLf
			Response.Write "  <table width=""100%"" height=""38"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
			Response.Write "        <input type=""button"" class=""inputbutton"" name=""Submit"" Onclick=""CheckForm()"" value="" ȷ �� "">" & vbCrLf
			Response.Write "        <input type=""reset"" class=""inputbutton"" name=""Submit2""  value="" �� �� "">" & vbCrLf
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "  </form>" & vbCrLf
			Response.Write "<Script Language=""javascript"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write "function is_email(str)" & vbCrLf
			Response.Write "{ if((str.indexOf('@')==-1)||(str.indexOf('.')==-1)){" & vbCrLf
			Response.Write "    return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "function SetLogoArea(Value)" & vbCrLf
			Response.Write "{"
			Response.Write "   document.all.LinkArea.style.display=Value;"
			Response.Write "}" & vbCrLf
			Response.Write "function CheckForm()" & vbCrLf
			Response.Write "{ var form=document.LinkForm;" & vbCrLf
			Response.Write "   if (form.SiteName.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""��������վ����!"");" & vbCrLf
			Response.Write "     form.SiteName.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.WebMaster.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""��������վվ��!"");" & vbCrLf
			Response.Write "     form.WebMaster.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    if ((form.Email.value!='@')&&(is_email(form.Email.value)==false))" & vbCrLf
			Response.Write "    {"
			Response.Write "    alert('�Ƿ���������!');" & vbCrLf
			Response.Write "     form.Email.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    if (form.OriPassWord.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""��������վ��ԭ������!"");" & vbCrLf
			Response.Write "     form.OriPassWord.focus();" & vbCrLf
			Response.Write "     return false;"
			Response.Write "    }" & vbCrLf
			Response.Write "    if (form.PassWord.value!='' && form.PassWord.value.length<6)" & vbCrLf
			Response.Write "    {"
			Response.Write "      alert(""��վ���벻������6λ!"");" & vbCrLf
			Response.Write "     form.PassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if(form.ConPassWord.value!='' && form.ConPassWord.value.length<6)" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""ȷ�����벻������6λ!"");" & vbCrLf
			Response.Write "     form.ConPassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.PassWord.value!=form.ConPassWord.value)" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""������������벻һ��!"");" & vbCrLf
			Response.Write "     form.PassWord.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.Url.value=='' || form.Url.value=='http://')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""��������վ��ַ"");" & vbCrLf
			Response.Write "     form.Url.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    form.submit();" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "//-->" & vbCrLf
			Response.Write "</Script>"
			Response.Write "</td></tr></table>" & vbCrLf
			Response.Write "        <br>"
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "</form>" & vbCrLf
			Response.Write "</body>" & vbCrLf
			Response.Write "</html>"
			End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../Plus/md5.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New FriendLinkDel
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkDel
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Dim LinkID, RSCheck, SiteName, WebMaster, Email, OriPassWord, Url, LinkType, Logo, Descript, Action, FolderID
			
			Action = Replace(Replace(Request("Action"), """", ""), "'", "")
			LinkID = KS.ChkClng(Request("id"))
			
			If LinkID=0 Then
			 Set KS = Nothing
			 Response.Write ("<script>alert('�������ݳ���!');history.back();</script>")
			End If
			If Action = "Del" Then
			 OriPassWord = MD5(KS.R(Request.Form("OriPassWord")),16)
			 If OriPassWord = "" Then
				  Call KS.AlertHistory("�޸�����������Ϣ��������ԭ������!", -1)
				  Set KS = Nothing
			End If
			Set RSCheck = Server.CreateObject("Adodb.Recordset")
			   RSCheck.Open " Select LinkID From KS_Link Where PassWord='" & OriPassWord & "'", Conn, 1, 1
			   If RSCheck.EOF And RSCheck.BOF Then
				  RSCheck.Close
				  Set RSCheck = Nothing
				  Call KS.AlertHistory("�Բ���,�������ԭ����������!", -1)
				  Set KS = Nothing
				  Response.End
			  End If
			  Conn.Execute ("Delete From KS_Link Where LinkID=" & LinkID)
			  RSCheck.Close
			  Set RSCheck = Nothing
			  Conn.Close
			  Set Conn = Nothing
			  Response.Write ("<script>alert('��������ɾ���ɹ�!');location.href='../';</script>")
			End If
			   Dim RSObj:Set RSObj = Conn.Execute("Select * From KS_Link Where LinkID=" & LinkID)
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
			   RSObj.Close: Set RSObj = Nothing
			Response.Write ("<html>") & vbCrLf
			Response.Write ("<head>") & vbCrLf
			Response.Write ("<title>ɾ����������</title>") & vbCrLf
			Response.Write ("<meta http-equiv=""Content-Language"" content=""zh-cn"">") & vbCrLf
			Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">") & vbCrLf
			Response.Write ("<link href=""../../../images/style.css"" rel=""stylesheet"" type=""text/css"">") & vbCrLf
			Response.Write ("</head>") & vbCrLf
			Response.Write ("<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">") & vbCrLf
			Response.Write ("<br>") & vbCrLf
			Response.Write ("  <table width=""770"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">") & vbCrLf
			Response.Write ("    <tr>") & vbCrLf
				  
			Response.Write ("    <td align=""center""><br>") & vbCrLf
			Response.Write ("        <table width=""500"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""table_border"">")
			Response.Write ("          <tr class=""link_table_title""> ") & vbCrLf
			Response.Write ("          <td>ɾ����������</td>") & vbCrLf
			Response.Write ("          </tr>") & vbCrLf
			Response.Write ("          <tr><td>") & vbCrLf
			
			Response.Write "  <form action=""?"" name=""LinkForm"" method=""post"">" & vbCrLf
			Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""Del"">" & vbCrLf
			Response.Write "   <input name=""ID"" type=""hidden"" value=""" & LinkID & """>" & vbCrLf
			Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td>" & vbCrLf
			Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td width=""20%"" height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">��վ����</div></td>" & vbCrLf
			Response.Write "            <td width=""542"" height=""25"">" & vbCrLf
			Response.Write SiteName & "</td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">�������</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			 on error resume next			   
			 Response.Write(Conn.Execute("Select FolderName From KS_LinkFolder Where FolderID=" & FolderID)(0))
			 
					   
			 Response.Write "         </td>" & vbCrLf
			 Response.Write "         </tr>"
			 Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">��վվ��</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write WebMaster & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">վ������</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write Email & "</td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "            <td height=""25"" align=""center"">��վ��ַ</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & Url & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>"
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">"
			Response.Write "              <div align=""center"">��վ���</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write Descript & "</td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr>" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">ԭ������</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""OriPassWord"" type=""password"" size=""42"" > <font color=""red"">* ��������</font> </td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "        </table>" & vbCrLf
			Response.Write "       </td>"
			Response.Write "    </tr>" & vbCrLf
			Response.Write "    </table>" & vbCrLf
			Response.Write "  <table width=""100%"" height=""38"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
			Response.Write "        <input type=""button"" name=""Submit"" Onclick=""CheckForm()"" value="" ȷ �� "">" & vbCrLf
			Response.Write "        <input type=""reset"" name=""Submit2""  value="" �� �� "">" & vbCrLf
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "  </form>" & vbCrLf
			Response.Write "<Script Language=""javascript"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write "function CheckForm()" & vbCrLf
			Response.Write "{ var form=document.LinkForm;" & vbCrLf
			Response.Write "if (form.OriPassWord.value=='')"
			Response.Write "    {"
			Response.Write "     alert(""��������վ��ԭ������!"");" & vbCrLf
			Response.Write "     form.OriPassWord.focus();" & vbCrLf
			Response.Write "     return false;"
			Response.Write "    }" & vbCrLf
			Response.Write " if (confirm('ȷ��ɾ����վ����Ϣ��?'))"
			Response.Write "  {  form.submit();" & vbCrLf
			Response.Write "    return true;}" & vbCrLf
			Response.Write "else" & vbCrLf
			Response.Write " {location.href='Index.asp'}"
			Response.Write "}" & vbCrLf
			Response.Write "//-->" & vbCrLf
			Response.Write "</Script>"
			Response.Write ("</td></tr></table>") & vbCrLf
			Response.Write ("        <br>") & vbCrLf
			Response.Write ("      </td>") & vbCrLf
			Response.Write ("    </tr>") & vbCrLf
			Response.Write ("  </table>") & vbCrLf
			Response.Write ("</form>") & vbCrLf
			Response.Write ("</body>") & vbCrLf
			Response.Write ("</html>") & vbCrLf
			End Sub
End Class
%>

 

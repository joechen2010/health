<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New AddWordJS
KSCls.Kesion()
Set KSCls = Nothing

Class AddWordJS
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'���岿��
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, JSConfig, JSConfigArr, Action, JSID, Str, Descript, FolderID
		Dim JSFileName, WordCss, OpenType, ArticleListNumber, RowHeight, TitleLen, ContentLen, ColNumber, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, DateCss, ContentCss, BGCss
		
		CurrPath = KS.GetCommonUpFilesDir()
		
		'�ж��Ƿ�༭
		JSID = Trim(Request.QueryString("JSID"))
		FolderID = Trim(Request.QueryString("FolderID"))
		If JSID = "" Then
		  Action = "Add"
		Else
		  Action = "Edit"
		  Dim JSRS, JSName
		  Set JSRS = Server.CreateObject("Adodb.Recordset")
		  JSRS.Open "Select * From KS_JSFile Where JSID='" & JSID & "'", Conn, 1, 1
		  If JSRS.EOF And JSRS.BOF Then
			 JSRS.Close
			 Set JSRS = Nothing
			 Response.Write ("<Script>alert('�������ݳ���!');history.back();</Script>")
			 Response.End
		  End If
			FolderID = JSRS("FolderID")
			JSName = Replace(Replace(JSRS("JSName"), "{JS_", ""), "}", "")
			JSFileName = Trim(JSRS("JSFileName"))
			JSID = JSRS("JSID")
			Descript = Trim(JSRS("Description"))
			JSConfig = Trim(JSRS("JSConfig"))
			JSRS.Close
			Set JSRS = Nothing
			JSConfig = Replace(JSConfig, """", "") 'ע:ȥ������˫����"
			JSConfigArr = Split(JSConfig, ",")
			WordCss = JSConfigArr(1)
			ColNumber = JSConfigArr(2)
			OpenType = JSConfigArr(3)
			ArticleListNumber = JSConfigArr(4)
			RowHeight = JSConfigArr(5)
			TitleLen = JSConfigArr(6)
			ContentLen = JSConfigArr(7)
			NavType = JSConfigArr(8)
			Navi = JSConfigArr(9)
			MoreLinkType = JSConfigArr(10)
			MoreLink = JSConfigArr(11)
			SplitPic = JSConfigArr(12)
			DateRule = JSConfigArr(13)
			DateAlign = JSConfigArr(14)
			TitleCss = JSConfigArr(15)
			DateCss = JSConfigArr(16)
			ContentCss = JSConfigArr(17)
			BGCss = JSConfigArr(18)
		End If
		If WordCss = "" Then WordCss = "A"
		If ArticleListNumber = "" Then ArticleListNumber = 5
		If RowHeight = "" Then RowHeight = 20
		If TitleLen = "" Then TitleLen = 20
		If ContentLen = "" Then ContentLen = 50
		If ColNumber = "" Then ColNumber = 1
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		Response.Write "<script src=""../../../ks_inc/jquery.js"" language=""JavaScript""></script>"
		Response.Write "<link href=""../Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		%>
		<script language="javascript">
		function SelectPicStyle(ObjValue)
		{
				document.all.ViewStylePicArea.innerHTML='<img src="../../Images/View/'+ObjValue+'.gif" border="0">';
		}
		function SetNavStatus()
		{
		  if (document.all.NavType.value==0)
		   {document.all.NavWord.style.display="";
			document.all.NavPic.style.display="none";}
		  else
		  {
		   document.all.NavWord.style.display="none";
		   document.all.NavPic.style.display="";}
		}
		function SetMoreLinkStatus()
		{
		if (document.all.MoreLinkType.value==0)
		   {document.all.LinkWord.style.display="";
			document.all.LinkPic.style.display="none";}
		  else
		  {
		   document.all.LinkWord.style.display="none";
		   document.all.LinkPic.style.display="";}
		}
		function CheckForm()
		{   
			if (document.myform.JSName.value=='')
			 {
			  alert('������JS����');
			  document.myform.JSName.focus(); 
			  return false
			  }
			  if (document.myform.JSFileName.value=='')
			  {
			   alert('������JS�ļ���');
			  document.myform.JSFileName.focus(); 
			  return false
			  }
			 if (CheckEnglishStr(document.myform.JSFileName,"JS�ļ���")==false) 
			   return false;
			 if (!IsExt(document.myform.JSFileName.value,'JS'))
			   { alert('JS�ļ�������չ��������.js');
				  document.myform.JSFileName.focus(); 
				  return false;
			   }
			var WordCss='"'+document.myform.WordCss.value+'"';
			var NavType=1;
			var ColNumber=document.myform.ColNumber.value;
			var OpenType='"'+document.myform.OpenType.value+'"';
			var ArticleListNumber=document.myform.ArticleListNumber.value;
			var RowHeight=document.myform.RowHeight.value;
			var TitleLen=document.myform.TitleLen.value;
			var ContentLen=document.myform.ContentLen.value;
			var Nav,NavType=document.myform.NavType.value;
			var MoreLink,MoreLinkType=document.myform.MoreLinkType.value;
			var SplitPic='"'+document.myform.SplitPic.value+'"';
			var DateRule=document.myform.DateRule.value;
			var DateAlign='"'+document.myform.DateAlign.value+'"';
			var TitleCss='"'+document.myform.TitleCss.value+'"';
			var DateCss='"'+document.myform.DateCss.value+'"';
			var ContentCss='"'+document.myform.ContentCss.value+'"';
			var BGCss='"'+document.myform.BGCss.value+'"';
		
			if  (ArticleListNumber=='')  ArticleListNumber=5;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav='"'+document.myform.TxtNavi.value+'"'
			 else  Nav='"'+document.myform.NaviPic.value+'"';
			if  (MoreLinkType==0) MoreLink='"'+document.myform.MoreLinkWord.value+'"'
			else  MoreLink='"'+document.myform.MoreLinkPic.value+'"';
			document.myform.JSConfig.value=	'GetWordJS,'+WordCss+','+ColNumber+','+OpenType+','+ArticleListNumber+','+RowHeight+','+TitleLen+','+ContentLen+','+NavType+','+Nav+','+MoreLinkType+','+MoreLink+','+SplitPic+','+DateRule+','+DateAlign+','+TitleCss+','+DateCss+','+ContentCss+','+BGCss;
			document.myform.submit();
		}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"">"
		Response.Write "<div align=""center"">"
		Response.Write "<form  method=""post"" name=""myform"" action=""AddJSSave.asp"">"
		Response.Write " <input type=""hidden"" name=""JSConfig"">"
		Response.Write " <input type=""hidden"" name=""JSType"" value=""1"">"
		Response.Write " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		Response.Write " <input type=""hidden"" name=""Page"" value=""" & Request("Page") & """>"
		Response.Write "  <input type=""hidden"" name=""JSID"" value=""" & JSID & """>"
		Response.Write " <input type=""hidden"" name=""FileUrl"" value=""AddWordJS.asp"">"
		Response.Write KS.ReturnJSInfo(JSID, JSName, JSFileName, FolderID, 3, Descript)
		Response.Write "<br>"
		Response.Write "    <table width=""96%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "      <tr>"
		Response.Write "        <td> <FIELDSET align=center>"
		Response.Write "          <LEGEND align=left>��������JS��������</LEGEND>"
		Response.Write "          <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "            <tr>"
		Response.Write "            <td width=""69%""><table width=""96%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" height=""28"">������ʽ"
		Response.Write "<select name=""WordCss"" class=""textbox"" id=""WordCss"" onchange=""SelectPicStyle(this.value)"">"
							   Dim SelStr
								   If WordCss = "A" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""A""" & SelStr & ">��ʽA</option>")
								If WordCss = "B" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""B""" & SelStr & ">��ʽB</option>")
								 If WordCss = "C" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								Response.Write ("<option value=""C""" & SelStr & ">��ʽC</option>")
								If WordCss = "D" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""D""" & SelStr & ">��ʽD</option>")
								 If WordCss = "E" Then
								   SelStr = " Selected"
								   Else
								   SelStr = ""
								   End If
								 Response.Write ("<option value=""E""" & SelStr & ">��ʽE</option>")
							  
		Response.Write "                      </select>"
		Response.Write "                      ��������"
		 
		 Response.Write "                     <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'��������');""  style=""width:60;"" value=""" & ColNumber & """ name=""ColNumber"">"
		 Response.Write "                   </td>"
		 Response.Write "                   <td width=""50%"" height=""28"">"
		 
		Response.Write KS.ReturnOpenTypeStr(OpenType)
		
		Response.Write "       </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" height=""28"">��������"
		Response.Write "                      <input name=""ArticleListNumber"" class=""textbox"" type=""text"" id=""ArticleListNumber""    style=""width:20%;"" onBlur=""CheckNumber(this,'��������');"" value=""" & ArticleListNumber & """>"
		Response.Write "                      ȡ<font color=""#FF0000"">0</font>ʱ,���г�ȫ������</td>"
		Response.Write "                    <td width=""50%"" height=""28"">�����о�"
		Response.Write "                      <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight""    style=""width:70%;"" onBlur=""CheckNumber(this,'�����о�');"" value=""" & RowHeight & """></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" height=""28"">��������"
		Response.Write "                      <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		Response.Write "                    </td>"
		Response.Write "                    <td width=""50%"" height=""28"">��������"
		Response.Write "                      <input name=""ContentLen"" class=""textbox"" type=""text"" id=""ContentLen""    style=""width:70%;"" onBlur=""CheckNumber(this,'��������');"" value=""" & ContentLen & """></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td width=""50%"" height=""28"">��������"
		Response.Write "                      <select name=""NavType"" class=""textbox"" style=""width:70%;"" onchange=""SetNavStatus()"">"
					
					If JSID = "" Or CStr(NavType) = "0" Then
					Response.Write ("<option value=""0"" selected>���ֵ���</option>")
					Response.Write ("<option value=""1"">ͼƬ����</option>")
				   Else
					Response.Write ("<option value=""0"">���ֵ���</option>")
					Response.Write ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
		Response.Write "                      </select></td>"
		Response.Write "                    <td width=""50%"" height=""28"">"
				 If JSID = "" Or CStr(NavType) = "0" Then
				  Response.Write ("<div align=""left"" id=""NavWord""> ")
				  Response.Write ("<input type=""text"" class=""textbox"" name=""TxtNavi"" onBlur='CheckBadChar(this,""���ֵ���"");' style=""width:90%;"" value=""" & Navi & """> ")
				  Response.Write ("</div>")
				  Response.Write ("<div align=""left"" id=NavPic style=""display:none""> ")
				  Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""NaviPic"" name=""NaviPic"">")
				  Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  Response.Write ("</div>")
				Else
				  Response.Write ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  Response.Write ("<input type=""text"" class=""textbox"" name=""TxtNavi"" onBlur='CheckBadChar(this,""���ֵ���"");' style=""width:90%;""> ")
				  Response.Write ("</div>")
				  Response.Write ("<div align=""left"" id=NavPic> ")
				  Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  Response.Write ("</div>")
				End If
		Response.Write "        </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr id=""MoreLinkArea"">"
		Response.Write "                    <td width=""50%"" height=""28"">��������"
		 Response.Write "                     <select name=""MoreLinkType"" style=""width:70%;"" class=""textbox"" onchange=""SetMoreLinkStatus()"">"
					If JSID = "" Or CStr(MoreLinkType) = "0" Then
					Response.Write ("<option value=""0"" selected>��������</option>")
					Response.Write ("<option value=""1"">ͼƬ����</option>")
				   Else
					Response.Write ("<option value=""0"">��������</option>")
					Response.Write ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
		Response.Write "                      </select></td>"
		Response.Write "                    <td width=""50%"" height=""28"">"
				  
				  If JSID = "" Or CStr(MoreLinkType) = "0" Then
					Response.Write ("<div align=""left"" id=""LinkWord""> ")
					Response.Write ("  <input type=""text"" class=""textbox"" onBlur='CheckBadChar(this,""��������"");' name=""MoreLinkWord"" style=""width:90%;"" value=""" & MoreLink & """>")
					Response.Write ("</div>")
					Response.Write ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
					Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
					Response.Write ("</div>")
				Else
				   Response.Write ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   Response.Write ("<input type=""text"" class=""textbox"" onBlur='CheckBadChar(this,""��������"");' name=""MoreLinkWord"" style=""width:90%;"">")
				   Response.Write ("</div>")
				   Response.Write ("<div align=""left"" id=""LinkPic""> ")
				   Response.Write ("<input type=""text"" readonly class=""textbox"" style=""width:100"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   Response.Write ("<input type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				   Response.Write ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				   Response.Write ("</div>")
				End If
		 Response.Write "       </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""28"" colspan=""2"">�ָ�ͼƬ"
		Response.Write "                      <input name=""SplitPic"" type=""text"" class=""textbox"" id=""SplitPic2"" style=""width:58%;"" value=""" & SplitPic & """ readonly>"
		Response.Write "                      <input name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic"" value=""ѡ��ͼƬ..."">"
		Response.Write "                      <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>"
		Response.Write "                     </td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""28"">���ڸ�ʽ"
		Response.Write "                      <select  style=""width:70%;"" class=""textbox"" name=""DateRule"" id=""select"">"
		Response.Write KS.ReturnDateFormat(DateRule)
		Response.Write "                      </select> </td>"
		Response.Write "                    <td height=""28""> <div align=""left"">���ڶ���"
		Response.Write "                        <select name=""DateAlign"" id=""select4"" style=""width:70%;"">"
					
					If JSID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""left""" & Str & ">�����</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""center""" & Str & ">���ж���</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 Response.Write ("<option value=""right""" & Str & ">�Ҷ���</option>")
					
		 Response.Write "                       </select>"
		 Response.Write "                     </div></td>"
		 Response.Write "                 </tr>"
		 Response.Write "                 <tr>"
		 Response.Write "                   <td height=""28"">������ʽ"
		 Response.Write "                     <input name=""TitleCss"" type=""text"" class=""textbox"" id=""TitleCss"" onBlur=""CheckBadChar(this,'������ʽ');"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		 Response.Write "                   <td height=""28"">������ʽ<font color=""#FF0000"">"
		 Response.Write "                     <input name=""DateCss"" type=""text"" class=""textbox""  id=""DateCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'������ʽ');"" value=""" & DateCss & """>"
		 Response.Write "                     </font></td>"
		 Response.Write "                 </tr>"
		 Response.Write "                 <tr>"
		 Response.Write "                   <td height=""28"">������ʽ"
		 Response.Write "                     <input name=""ContentCss"" type=""text"" class=""textbox"" id=""ContentCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'������ʽ');"" value=""" & ContentCss & """></td>"
		 Response.Write "                   <td height=""28"">������ʽ"
		 Response.Write "                      <input name=""BGCss"" type=""text"" class=""textbox"" id=""BGCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'������ʽ');"" value=""" & BGCss & """></td>"
		 Response.Write "                 </tr>"
		 Response.Write "               </table></td>"
		Response.Write "              <td width=""31%"" align=""center""><table width=""90%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "                  <tr>"
		 Response.Write "                   <td height=""25"" align=""center""><strong>��ʽԤ��</strong></td>"
		Response.Write "                  </tr>"
		Response.Write "                  <tr>"
		Response.Write "                    <td height=""100%"" id=""ViewStylePicArea"">&nbsp;</td>"
		Response.Write "                  </tr>"
		Response.Write "                </table></td>"
		Response.Write "            </tr>"
		Response.Write "          </table>"
		Response.Write "          </FIELDSET></td>"
		Response.Write "      </tr>"
		Response.Write "    </table>"
		Response.Write "    </form>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<script>"
		Response.Write "SelectPicStyle('" & WordCss & "');"
		Response.Write "</script>"
		End Sub
End Class
%> 

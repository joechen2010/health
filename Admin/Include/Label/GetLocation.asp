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
Set KSCls = New GetLocation
KSCls.Kesion()
Set KSCls = Nothing

Class GetLocation
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
		Dim InstallDir, CurrPath, FolderID, LabelContent,Action, LabelID, Str, Descript
		Dim Bold, StartTag, OpenType, NavType, Navi, MoreLinkType, TitleCss,ShowTitle
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		
		With KS
		'�ж��Ƿ�༭
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetLocation", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				Bold = Node.getAttribute("bold")
				StartTag = Node.getAttribute("starttag")
				NavType = Node.getAttribute("navtype")
				Navi = Node.getAttribute("nav")
				OpenType = Node.getAttribute("opentype")
				TitleCss = Node.getAttribute("titlecss")
				ShowTitle= Node.getAttribute("showtitle")
			End If
			Set Node=Nothing
			XMLDoc=Empty
		End If
		If Navi = "" Then Navi = " >> "
		If StartTag="" Then StartTag="��ǰλ�ã�"
		If ShowTitle="" Or IsNull(ShowTitle) Then ShowTitle=true
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		%>
		<script language="javascript">
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   { $("#NavWord").show();
			 $("#NavPic").hide();
		  }else{
		     $("#NavWord").hide();
		     $("#NavPic").show();
		 }
		}
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('�������ǩ����');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			  if ($("input[name=StartTag]").val()=='')
			 {
			  alert('������λ�õ����Ŀ�ʼ��־');
			  $("input[name=StartTag]").focus(); 
			  return false
			  }
			var StartTag=$("input[name=StartTag]").val();
			var Bold=false; 
			if ($("#Bold").attr("checked")==true){Bold=true}
			var ShowTitle=false; 
			if ($("#ShowTitle").attr("checked")==true){ShowTitle=true}
			var OpenType=$("#OpenType").val();
			var Nav,NavType=$("select[name=NavType]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val()
			 else  Nav=$("input[name=NaviPic]").val();
			 var tagVal='{Tag:GetLocation labelid="0" bold="'+Bold+'" starttag="'+StartTag+'" navtype="'+NavType+'" nav="'+Nav+'" opentype="'+OpenType+'" titlecss="'+TitleCss+'" showtitle="'+ShowTitle+'"}{/Tag}'
			$("input[name=LabelContent]").val(tagVal);
			
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetLocation.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">��ʼ��־"
		.echo "                <input type=""text"" class=""textbox"" name=""StartTag"" style=""width:200;"" value=""" & StartTag & """>"
		.echo "                <font color=""#FF0000"">"
		.echo "               <input name=""Bold"" type=""checkbox"" id=""Bold"" value=""true""��"
		  If CBool(Bold) = True Then .echo " checked"
		.echo ">"
		.echo "                </font>�Ӵ�<font color=""#FF0000"">���� �磺&quot;��ǰλ�ã�&quot; �ȵ�</font></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""30"">��������"
		.echo "                <select class=""textbox"" name=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
					
					If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>���ֵ���</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">���ֵ���</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
				   
		.echo "                </select>"
		.echo "              </td>"
		.echo "              <td>"
				
				If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """>")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=""NavPic"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"">")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
				  .echo ("</div>")
				End If
				
		 .echo "             </td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">"
				  
		 .echo KS.ReturnOpenTypeStr(OpenType)
		 
		.echo "              </td>"
		.echo "              <td height=""30"">&nbsp;</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">������ʽ"
		.echo "                <input name=""TitleCss"" class='textbox' type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""30""><font color=""#FF0000"">�Ѷ����CSS ,Ҫ��һ������ҳ��ƻ���</font></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">����ҳ��ʾ����"
		.echo "                <input name=""ShowTitle"" type=""checkbox"" id=""ShowTitle"" value=""true"""
		If Cbool(ShowTitle)=true Then .echo " checked"
		.echo "></td>"
		.echo "              <td height=""30""></td>"
		.echo "            </tr>"
		.echo "                  </table>"	
		.echo "  </form>"
		  
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 

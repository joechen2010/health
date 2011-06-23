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
Set KSCls = New GetAnnounceList
KSCls.Kesion()
Set KSCls = Nothing

Class GetAnnounceList
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
		Dim FolderID, CurrPath, InstallDir, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim AnnounceType, OWidth, OHeight, Width, Height, Speed, ShowStyle, OpenType, ListNumber, TitleLen, ShowAuthor, ContentLen, NavType, Navi, TitleCss,ChannelID,AjaxOut
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetAnnounceList", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				AnnounceType = Node.getAttribute("announcetype")
				OWidth = Node.getAttribute("owidth")
				OHeight = Node.getAttribute("oheight")
				Width = Node.getAttribute("width")
				Height = Node.getAttribute("height")
				Speed = Node.getAttribute("speed")
				ShowStyle = Node.getAttribute("showstyle")
				OpenType =  Node.getAttribute("opentype")
				ListNumber = Node.getAttribute("listnumber")
				TitleLen = Node.getAttribute("titlelen")
				ShowAuthor = Node.getAttribute("showauthor")
				ContentLen = Node.getAttribute("contentlen")
				NavType = Node.getAttribute("navtype")
				Navi = Node.getAttribute("nav")
				TitleCss = Node.getAttribute("titlecss")
				ChannelID= Node.getAttribute("channelid")
				AjaxOut   = Node.getAttribute("ajaxout")
			End If
			Set Node=Nothing
			XMLDoc=Empty
			
		End If
		If ShowAuthor = "" Then ShowAuthor = 1
		If ListNumber = "" Then ListNumber = 1
		If OpenType = "" Then OpenType = 1
		If OWidth = "" Then OWidth = 450
		If OHeight = "" Then OHeight = 400
		If Width = "" Then Width = 350
		If Height = "" Then Height = 20
		If TitleLen = "" Then TitleLen = 30
		If ContentLen = "" Then ContentLen = 100
		If Speed = "" Then Speed = 1
		If ChannelID="" Then ChannelID=0
		If AjaxOut="" Or IsNull(AjaxOut) Then AjaxOut=true
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		%>
		<script>
		function ChangeArea(Num)
		{
		if (Num==2)
		{
		 $("#DisplayMode").show();
		 $("#MarqueeArea").show();
		 $("#OpenArea").hide();
		}
		else if (Num==1)
		{
		 $("#DisplayMode").hide();
		 $("#MarqueeArea").hide();
		 $("#OpenArea").show();
		}
		else
		{
		 $("#DisplayMode").show();
		 $("#MarqueeArea").hide();
		 $("#OpenArea").hide();
		}
		}
		function ChangeOpenType(Num)
		{
		if (Num==0)
		 {
		 $("#OpenArea").show();
		 }
		else
		{
		$("#OpenArea").hide();
		}
		}
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
			var AnnounceType,ShowStyle,OpenType,ShowAuthor;
			var ChannelID=$("#ChannelID").val();
			var OWidth=$("input[name=OWidth]").val();
			var OHeight=$("input[name=OHeight]").val();
			var Width=$("input[name=Width]").val();
			var Height=$("input[name=Height]").val();
			var Speed=$("input[name=Speed]").val();
			var ListNumber=$("input[name=ListNumber]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var ContentLen=$("input[name=ContentLen]").val();
			var Nav,NavType=$("#NavType").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			AnnounceType=$("input[name=AnnounceType][checked=true]").val();
            ShowStyle=$("input[name=ShowStyle][checked=true]").val();
			OpenType=$("input[name=OpenType][checked=true]").val();
			ShowAuthor=$("input[name=ShowAuthor][checked=true]").val();
			if  (NavType==0) Nav=$("#TxtNavi").val()
			 else  Nav=$("#NaviPic").val();
			 		
            var tagVal='{Tag:GetAnnounceList labelid="0" announcetype="'+AnnounceType+'" owidth="'+OWidth+'" oheight="'+OHeight+'" width="'+Width+'" height="'+Height+'" speed="'+Speed+'" showstyle="'+ShowStyle+'" opentype="'+OpenType+'" listnumber="'+ListNumber+'" titlelen="'+TitleLen+'" showauthor="'+ShowAuthor+'" contentlen="'+ContentLen+'" navtype="'+NavType+'" nav="'+Nav+'" titlecss="'+TitleCss+'" channelid="'+ChannelID+'" ajaxout="'+AjaxOut+'"}{/Tag}'
			$("input[name=LabelContent]").val(tagVal);
			
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetAnnounceList.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=2> ����ģ��" &ReturnChannelList(ChannelID)
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">����Ajax���</label> <font color=red>�������ajax���,�����¹���,�Ͳ���Ҫ���·�������ҳ��</font>"				

		.echo "</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2""> ��������"
					   
					   If AnnounceType = 0 Or AnnounceType = "" Then
						 .echo ("<input onclick=""ChangeArea(0);"" type=""radio"" name=""AnnounceType"" value=""0"" Checked>��ͨ ")
						Else
						 .echo ("<input onclick=""ChangeArea(0);"" type=""radio"" name=""AnnounceType"" value=""0"">��ͨ ")
						End If
						If AnnounceType = 1 Then
						 .echo ("<input onclick=""ChangeArea(1);"" type=""radio"" name=""AnnounceType"" value=""1"" Checked>���� ")
						Else
						 .echo ("<input onclick=""ChangeArea(1);"" type=""radio"" name=""AnnounceType"" value=""1"">���� ")
						End If
						If AnnounceType = 2 Then
						 .echo ("<input onclick=""ChangeArea(2);"" type=""radio"" name=""AnnounceType"" value=""2"" Checked>���� ")
						Else
						 .echo ("<input onclick=""ChangeArea(2);"" type=""radio"" name=""AnnounceType"" value=""2"">���� ")
						End If
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr  class='tdbg' id=""OpenArea"" style=""display:none"">"
		.echo "              <td height=""30"" colspan=""2"">�������"
		.echo "                <input name=""OWidth"" class=""textbox""  onBlur=""CheckNumber(this,'���');"" type=""text"" id=""OWidth"" style=""width:50;"" value=""" & OWidth & """> ���� �����߶�"
		.echo "                <input name=""OHeight"" class=""textbox""  onBlur=""CheckNumber(this,'�߶�');"" type=""text"" id=""OHeight"" style=""width:50;"" value=""" & OHeight & """> ����</td>"
		.echo "            </tr>"
		.echo "            <tr  class='tdbg' id=""MarqueeArea"" style=""display:none"">"
		.echo "              <td height=""30"" colspan=""2"">�������"
		.echo "                <input name=""Width"" class=""textbox""  onBlur=""CheckNumber(this,'���');"" type=""text"" id=""Width"" style=""width:50;"" value=""" & Width & """>"
		.echo "                ���� �����߶�"
		.echo "                <input name=""Height"" class=""textbox""  onBlur=""CheckNumber(this,'�߶�');"" type=""text"" id=""Height"" style=""width:50;"" value=""" & Height & """>"
		.echo "                ���� �����ٶ�"
		.echo "                <input name=""Speed"" class=""textbox""  onBlur=""CheckNumber(this,'�ٶ�');"" type=""text"" id=""Speed"" style=""width:50;"" value=""" & Speed & """></td>"
		.echo "            </tr>"
		.echo "            <tr  class='tdbg' ID=""DisplayMode"">"
		.echo "              <td width=""50%"" height=""30"">��ʾ��ʽ"
						
						If ShowStyle = 1 Or ShowStyle = "" Then
						 .echo ("<input type=""radio"" name=""ShowStyle"" value=""1"" Checked>���� ")
						Else
						 .echo ("<input type=""radio"" name=""ShowStyle"" value=""1"">���� ")
						End If
						If ShowStyle = 2 Then
						 .echo ("<input type=""radio"" name=""ShowStyle"" value=""2"" Checked>���� ")
						Else
						 .echo ("<input type=""radio"" name=""ShowStyle"" value=""2"">���� ")
						End If
						
		.echo "              </td>"
		.echo "              <td>�򿪷�ʽ"
						
						If OpenType = 0 Then
						 .echo ("<input type=""radio"" onclick=""ChangeOpenType(0);"" name=""OpenType"" value=""0"" Checked>���� ")
						Else
						 .echo ("<input type=""radio"" onclick=""ChangeOpenType(0);"" name=""OpenType"" value=""0"">���� ")
						End If
						If OpenType = 1 Then
						 .echo ("<input type=""radio"" onclick=""ChangeOpenType(1);"" name=""OpenType"" value=""1"" Checked>��ͨ���� ")
						Else
						 .echo ("<input type=""radio"" onclick=""ChangeOpenType(1);"" name=""OpenType"" value=""1"">��ͨ���� ")
						End If
						
		.echo "                </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">��������"
		.echo "                <input name=""ListNumber"" class=""textbox""  onBlur=""CheckNumber(this,'��������');"" type=""text"" id=""ListNumber"" style=""width:70%;"" value=""" & ListNumber & """></td>"
		.echo "              <td height=""30""><font color=""#FF0000"">����Ϊ0ʱ���г����й���</font></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">�������� <input name=""TitleLen"" class=""textbox""  onBlur=""CheckNumber(this,'��������');"" type=""text"" id=""TitleLen"" style=""width:70%;"" value=""" & TitleLen & """></td>"
		.echo "              <td height=""24"">��������"
					   
					   If CInt(ShowAuthor) = 1 Then
						 .echo ("<input type=""radio"" name=""ShowAuthor"" value=""1"" Checked>��ʾ ")
						Else
						 .echo ("<input type=""radio"" name=""ShowAuthor"" value=""1"">��ʾ ")
						End If
						If CInt(ShowAuthor) = 0 Then
						 .echo ("<input type=""radio"" name=""ShowAuthor"" value=""0"" Checked>����ʾ ")
						Else
						 .echo ("<input type=""radio"" name=""ShowAuthor"" value=""0"">����ʾ ")
						End If
					   
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">��������"
		.echo "                <input name=""ContentLen"" class=""textbox""  onBlur=""CheckNumber(this,'��������');"" type=""text"" id=""ContentLen"" style=""width:70%;"" value=""" & ContentLen & """></td>"
		.echo "              <td height=""24""><font color=""#FF0000"">����Ϊ0ʱ����ʾ��������</font></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">��������"
		.echo "                <select name=""NavType"" id=""NavType"" class=""textbox"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				  
				  If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>���ֵ���</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">���ֵ���</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""24"">"
				
				If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;""> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				End If
				
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "             <td height=""30"">������ʽ"
		.echo "               <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""30""><font color=""#FF0000"">�Ѷ����CSS ,Ҫ��һ������ҳ��ƻ���</font></td>"
		.echo "            </tr>"
		.echo "                  </table>"	
		.echo "  </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		
		
		.echo "<script>"
		.echo "ChangeArea(" & AnnounceType & ");"
		.echo "</script>"
		End With
		End Sub
		Public Function ReturnChannelList(SelectChannelID)
	  Dim ChannelRS:Set ChannelRS=Server.CreateObject("ADODB.Recordset")
	  Dim ChannelStr:ChannelStr = ""
	   ChannelRS.Open "Select * From [KS_Channel] Where ChannelStatus=1", Conn, 1, 1
	   If ChannelRS.EOF And ChannelRS.BOF Then
		  ChannelRS.Close:Set ChannelRS = Nothing:Exit Function
	  Else
		  ChannelStr = "<select class='textbox' id=""ChannelID"" name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1"">"
		  ChannelStr = ChannelStr & "<option value=0>��ָ��ģ��</option>"
		  If SelectChannelID=9999 Then
		  ChannelStr = ChannelStr & "<option value=9999 selected style='color:red'>��ǰģ��ͨ�ñ�ǩ</option>"
		  Else
		  ChannelStr = ChannelStr & "<option value=9999 style='color:red'>��ǰģ��ͨ�ñ�ǩ</option>"
		  End If
		 If SelectChannelID=9998 Then
		  ChannelStr = ChannelStr & "<option value=9998 style='color:blue' selected>��վ��ҳ����</option>"
		 else
		  ChannelStr = ChannelStr & "<option value=9998 style='color:blue'>��վ��ҳ����</option>"
		 End If
		 If SelectChannelID=9990 Then
		  ChannelStr = ChannelStr & "<option value=9990 selected style='color:red'>��Ա���Ĺ���</option>"
		  Else
		  ChannelStr = ChannelStr & "<option value=9990 style='color:red'>��Ա���Ĺ���</option>"
		  End If
  
	   Do While Not ChannelRS.EOF
		 If cstr(ChannelRS("ChannelID")) = cstr(SelectChannelID) Then
		  ChannelStr = ChannelStr & "<option selected value=" & ChannelRS("ChannelID") & ">" & ChannelRS("ChannelName") & "</option>"
		 Else
		   ChannelStr = ChannelStr & "<option value=" & ChannelRS("ChannelID") & ">" & ChannelRS("ChannelName") & "</option>"
		 End If
		ChannelRS.MoveNext
		Loop
	   ChannelRS.Close:Set ChannelRS = Nothing
	  End If
		 ChannelStr = ChannelStr & "</Select>"
	   ReturnChannelList = ChannelStr
	End Function

End Class
%> 

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
Set KSCls = New GetCirSpecialList
KSCls.Kesion()
Set KSCls = Nothing

Class GetCirSpecialList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim ClassCol, ClassCss, MenuBgType, MenuBg
		Dim ShowClassName, OpenType, Num, IntroLen, TitleLen, RowHeight,SpecialSort, ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, PhotoCss,ShowStyle,PicWidth,PicHeight,PrintType,Col
		Dim ClassPrintType,LabelStyleW,LabelStyle,AjaxOut
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
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetCirSpecialList", ""),"}" & LabelStyle &"{/Tag}", "")
			
			'response.write labelcontent
			LabelStyleW        = Split(LabelStyle,"��")(0)
			LabelStyle         = Split(LabelStyle,"��")(1)
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
                ClassCol     = Node.getAttribute("classcol")
				ClassCss     = Node.getAttribute("classcss")
				MenuBgType   = Node.getAttribute("menubgtype")
				MenuBg       = Node.getAttribute("menubg")
				Num          = Node.getAttribute("num")
				IntroLen     = Node.getAttribute("introlen")
				TitleLen     = Node.getAttribute("titlelen")
				RowHeight    = Node.getAttribute("rowheight")
				Col          = Node.getAttribute("col")
				OpenType     = Node.getAttribute("opentype")
				NavType      = Node.getAttribute("navtype")
				Navi         = Node.getAttribute("nav")
				MoreLinkType = Node.getAttribute("morelinktype")
				MoreLink     = Node.getAttribute("morelink")
				SplitPic     = Node.getAttribute("splitpic")
				DateRule     = Node.getAttribute("daterule")
				DateAlign    = Node.getAttribute("datealign")
				TitleCss     = Node.getAttribute("titlecss")
				PhotoCss     = Node.getAttribute("photocss")
				ShowStyle    = Node.getAttribute("showstyle")
				PicWidth     = Node.getAttribute("picwidth")
				PicHeight    = Node.getAttribute("picheight")
				PrintType    = Node.getAttribute("printtype")
				AjaxOut      = Node.getAttribute("ajaxout")
                ClassPrintType    = Node.getAttribute("classprinttype")
			End If
			Set Node=Nothing
			Set XMLDoc=Nothing
		End If
		If ShowStyle="" Then ShowStyle=1
		If PrintType="" Then PrintType=2
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If Col="" Then Col=1
		If Num = "" Then Num = 10
		If IntroLen = "" Then IntroLen = 200
		If TitleLen = "" Then TitleLen = 30
		If RowHeight= "" Then RowHeight= 22
		If ClassCol = "" Then ClassCol = 2
		If AjaxOut="" Or IsNull(AjaxOut) Then AjaxOut=false
		If ClassPrintType="" Then ClassPrintType=2
		If LabelStyleW="" Then LabelStyleW="<div class=""col"">" & vbcrlf & " <div class=""t""><span><a href=""{@specialclassurl}"" target=""_blank"">����...</a></span>{@specialclassname}</div>" & vbcrlf & " <ul>{$InnerText}</ul>" & vbcrlf & "</div>"
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@specialurl}"" target=""_blank"">{@specialname}</a></li>" & vbcrlf & "[/loop]"
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		%>
		<style type="text/css">
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
		<script language="javascript">
		$(document).ready(function(){
		 ChangeClassPrintOutArea($("#ClassPrintType>option[selected=true]").val());
		 ChangeOutArea($("#PrintType>option[selected=true]").val());
		});
		
		
		function ChangeClassPrintOutArea(Val)
		{
		   if (Val==1)
		   {$("#ClassTable").show();
		    $("#ClassDiy").hide();
		   }else{
		    $("#ClassTable").hide();
		    $("#ClassDiy").show();
		   }
		}
       function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		var tag=null;
		 function setPos(Tag)
		 {   tag=Tag;
		     if (document.all){
				$("#"+Tag).focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("#"+Tag).selectionStart;
			  }
			
		 }
		 //����
		function InsertValue(Val)
		{  if (pos==null||tag==null) {alert('���ȶ�λҪ�����λ��!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#"+tag);
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }		
		
		function ChangeOutArea(Val)
		{
		 if (Val==2){
		  $("#DiyArea").show();
		  $("#TableArea").hide();
		 }
		 else{
		  $("#DiyArea").hide();
		  $("#TableArea").show();
		 }
		}
		function SetMenuBg()
		{if ($("#MenuBgType").val()==0)
		   {
		    $("#MenuBgColor").show();
			$("#MenuBgPic").hide();}
		  else
		  {
		    $("#MenuBgColor").hide();
		    $("#MenuBgPic").show();}
		   }
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   {$("#NavWord").show();
			$("#NavPic").hide();
			}else{
		   $("#NavWord").hide();
		   $("#NavPic").show();}
		}
		function SetMoreLinkStatus()
		{
		  if ($("select[name=MoreLinkType]").val()==0){
		    $("#LinkWord").show();
			$("#LinkPic").hide();
			}else{
		   $("#LinkWord").hide();
		   $("#LinkPic").show();}
		}
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('�������ǩ����');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ClassCol=$("#ClassCol").val();
			var ClassCss=$("#ClassCss").val();
			var MenuBgType=1,NavType=1;
			var MenuBg,MenuBgType=$("#MenuBgType").val();
			var OpenType=$("#OpenType").val();
			var Num=$("input[name=Num]").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var Col=$("input[name=Col]").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var Nav,NavType=$("#NavType").val();
			var MoreLink,MoreLinkType=$("#MoreLinkType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule=$("#DateRule").val();
			var DateAlign=$("#DateAlign").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var PhotoCss=$("input[name=PhotoCss]").val();
			var ShowStyle=$("#ShowStyle").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PrintType=$("#PrintType").val();
	    	var ClassPrintType=$("#ClassPrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			
			if  (Num=='')  Num=10;
			if (IntroLen=='') IntroLen=20
			if  (TitleLen=='') TitleLen=30;
			if  (ClassCol=='') ClassCol=2;
			if  (MenuBgType==0) MenuBg=$("#ColorMenuBg").val()
			 else  MenuBg=$("#PicMenuBg").val();	
			if  (NavType==0) Nav=$("#TxtNavi").val()
			 else  Nav=$("#NaviPic").val();
			if  (MoreLinkType==0) MoreLink=$("#MoreLinkWord").val()
			else  MoreLink=$("#MoreLinkPic").val();
			
			var tagVal='{Tag:GetCirSpecialList labelid="0" classid="0" classprinttype="'+ClassPrintType+'" ajaxout="'+AjaxOut+'" classcol="'+ClassCol+'" classcss="'+ClassCss+'" menubgtype="'+MenuBgType+'" menubg="'+MenuBg+'" num="'+Num+'" introlen="'+IntroLen+'" titlelen="'+TitleLen+'" rowheight="'+RowHeight+'" col="'+Col+'" opentype="'+OpenType+'" navtype="'+NavType+'" nav="'+Nav+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" photocss="'+PhotoCss+'" showstyle="'+ShowStyle+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" printtype="'+PrintType+'"}';
			tagVal  +=$("#LabelStyleW").val()+'��'+$("#LabelStyle").val()+'{/Tag}';

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
		.echo " <input type=""hidden"" name=""LabelContent"" ID=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""1"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetCirSpecialList.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" colspan='2' height=""24"">&nbsp;&nbsp;&nbsp;&nbsp;<strong>��Ŀ�����ʽ</strong>&nbsp;"
		.echo " <select class='textbox'  name=""ClassPrintType"" id=""ClassPrintType"" onChange=""ChangeClassPrintOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If ClassPrintType=1 Then .echo " selected"
		.echo ">��ͨ(Table)</option>"
        .echo "  <option value=""2"""
		If ClassPrintType=2 Then .echo " selected"
		.echo ">�Զ��������ʽ</option>"
        .echo "</select>"
		.echo "            <font color=green>���ڸ��õĿ���,����ѡ���Զ��������ʽ</font>"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">����Ajax���</label>"

		.echo "</td></tr>"
		
		.echo "         <tbody id=""ClassTable"">"
		.echo "              <tr class='tdbg'>"
		.echo "                <td width=""50%"" align='right' height=""20"">��Ŀ����"
		.echo "                  <input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'��������');""  style=""width:70%;"" value=""" & ClassCol & """ name=""ClassCol"" id=""ClassCol"">"
		.echo "                </td>"
		.echo "                <td width=""50%"" height=""20"">��ĿCSS����"
						  
		.echo "                <input name=""ClassCss"" class=""textbox"" type=""text"" id=""ClassCss"" value=""" & ClassCss & """></td>"
		.echo "              </tr>"
		.echo "              <tr class='tdbg'>"
		.echo "                <td width=""50%""  align='right' height=""20""> ��ͷ����"
		.echo "                  <select name=""MenuBgType"" id=""MenuBgType"" class=""textbox"" style=""width:70%;"" onchange=""SetMenuBg()"">"
				  
				  If LabelID = "" Or MenuBgType = "0" Then
					.echo ("<option value=""0"" selected>������ɫ</option>")
					.echo ("<option value=""1"">����ͼƬ</option>")
				   Else
					.echo ("<option value=""0"">������ɫ</option>")
					.echo ("<option value=""1"" selected>����ͼƬ</option>")
				   End If
		.echo "                  </select></td>"
		.echo "                <td width=""50%"" height=""20"">"
				
				If LabelID = "" Or MenuBgType = "0" Then
				  .echo ("<div align=""left"" id=""MenuBgColor""> ")
				  .echo ("<input type=""text"" class=""textbox"" id=""ColorMenuBg"" name=""ColorMenuBg"" style=""width:120;"" value=""" & MenuBg & """>")
				  .echo " <img border=0 id=""ColorMenuBgShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MenuBg & ";"" onClick=""Getcolor(this,'../../../ks_editor/SelectColor.asp','ColorMenuBg');"" title=""ѡȡ��ɫ"">"
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=""MenuBgPic"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""PicMenuBg"" name=""PicMenuBg"">")
				  .echo ("<input  class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.PicMenuBg);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.PicMenuBg.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""MenuBgColor"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""ColorMenuBg"" id=""ColorMenuBg"" style=""width:120;""> ")
				  .echo " <img border=0 id=""ColorMenuBgShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & MenuBg & ";"" onClick=""Getcolor(this,'../../../ks_editor/SelectColor.asp','ColorMenuBg');"" title=""ѡȡ��ɫ"">"
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=""MenuBgPic"">")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""PicMenuBg"" name=""PicMenuBg"" value=""" & MenuBg & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.PicMenuBg);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.PicMenuBg.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				End If
				
		.echo "                </td>"
		.echo "              </tr>"
		.echo "              <tr><td colspan=2><hr color=#ff6600 size=1></td></tr>"
		.echo "          </tbody>"
		
	    .echo "           <tbody id=""ClassDiy"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='field'>"
		.echo "               <table border='0' width='100%'>"
		.echo "                <tr><td align='center' width='100'><strong>���ñ�ǩ:</strong></td>"
		.echo "                <td><li onclick=""InsertLabel('{@autoid}')"">�� ��</li><li onclick=""InsertLabel('{@specialclassname}')"">ר���������</li><li onclick=""InsertLabel('{@specialclassurl}')"">ר�����URL</li><li onclick=""InsertLabel('{@specialclassintro}')"">ר��������(200��)</li><li onclick=""InsertLabel('{@classid}')"">ר�����СID</li></td>"
		.echo "                 </tr></table>"
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "                <table border='0' width='100%'><tr><td width='100' align='center'><strong>��ѭ��(����)</strong><br><font color=blue>���������ǩ{$InnerText}</font></td>"
		.echo "                <td><textarea name='LabelStyleW' onkeyup='setPos(""LabelStyleW"")' onclick='setPos(""LabelStyleW"")' id='LabelStyleW' style='width:100%;height:120px'>" & LabelStyleW & "</textarea></td>"
		.echo "                </tr>"
		.echo "               </table>"
		.echo "             </td>"
		.echo "            </tr>"
		.echo "           </tbody>"		
		
		
		
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">&nbsp;&nbsp;&nbsp;&nbsp;<strong>ר�������ʽ</strong>&nbsp;"
		.echo " <select class='textbox'  name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">�ı��б���ʽ(Table)</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">�Զ��������ʽ</option>"
        .echo "</select>"
		.echo "             </td> <td><span id='ShowDiyDate'></span> </td>"
		.echo "            </tr>"
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' align='center'><strong>���ñ�ǩ</strong></td>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@specialurl}')"">ר������URL</li> <li onclick=""InsertLabel('{@specialid}')"">ר��ID</li><li onclick=""InsertLabel('{@specialname}')"">ר������</li><li onclick=""InsertLabel('{@specialphotourl}')"">ר��ͼƬ</li><li onclick=""InsertLabel('{@classid}')"">����ID</li><li onclick=""InsertLabel('{@specialclassname}')"">��������</li><li onclick=""InsertLabel('{@specialclassurl}')"">����URL</li> <li onclick=""InsertLabel('{@intro}')"">��Ҫ����</li><li onclick=""InsertLabel('{@photourl}')"">ͼƬ��ַ</li><li onclick=""InsertLabel('{@adddate}')"">���ʱ��</li><li onclick=""InsertLabel('{@creater}')"">������</li></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'>"
		.echo "               <table border='0' width='100%'><tr><td width='100' align='center'><strong>��ѭ��(ר��)</strong></td>"
		.echo "               <td><textarea name='LabelStyle' onkeyup='setPos(""LabelStyle"")' onclick='setPos(""LabelStyle"")' id='LabelStyle' style='width:100%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "               </tr>"
		.echo "              </table>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>ʹ��˵�� :</font></strong><br />1��ѭ����ǩ[loop=n][/loop]�Կ���ʡ��,Ҳ����ƽ�г��ֶ�ԣ�<br /></font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"	
		
		
	
		.echo "            <tr class='tdbg'>"
		.echo "              <td colspan='2' height=""25"">ר������"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num2""    style=""width:50px;"" onBlur=""CheckNumber(this,'ר������');"" value=""" & Num & """> ��������"
		.echo "                <input name=""IntroLen"" class=""textbox"" type=""text"" id=""IntroLen"" style=""width:50px;"" onBlur=""CheckNumber(this,'��������');"" value=""" & IntroLen & """> �и�<input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight"" style=""width:50px;"" onBlur=""CheckNumber(this,'�и�');"" value=""" & RowHeight & """> ����<input name=""Col"" class=""textbox"" type=""text"" id=""Col"" style=""width:50px;"" onBlur=""CheckNumber(this,'����');"" value=""" & Col & """></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""25"">��������"
		.echo "                <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""25"">"
		.echo KS.ReturnOpenTypeStr(OpenType)
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""25"">���ڸ�ʽ"
		.echo "                <select  style=""width:70%;"" name=""DateRule"" id=""DateRule"" class=""textbox"">"
		.echo KS.ReturnDateFormat(DateRule)
		.echo "</select> </td>"
		.echo "              <td height=""25""> <div align=""left"">���ڶ���"
		.echo "                  <select name=""DateAlign"" class=""textbox"" id=""DateAlign"" style=""width:70%;"">"
					
					If LabelID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""left""" & Str & ">�����</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""center""" & Str & ">���ж���</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""right""" & Str & ">�Ҷ���</option>")
				   
		.echo "                  </select>"
		.echo "                </div></td>"
		.echo "            </tr>"
		
		.echo "       <tbody id=""TableArea"">"
		.echo "             <tr class='tdbg'>"
		.echo "               <td width=""50%"" height=""24"">" &KS.ReturnSpecialStyle(ShowStyle)
		.echo "               </td>"
		.echo "               <td width=""50%"" height=""24"">ͼƬ���� ��"
		.echo "<input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ���');"">"
		.echo "                ���� ��"
		.echo "<input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ�߶�');"">"
		.echo "                ����</td>"
		.echo "             </tr>"
		.echo "             <tr class='tdbg'>"
		.echo "               <td height=""24"">���� CSS"
		.echo "                 <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" onBlur=""CheckBadChar(this,'������ʽ');"" value=""" & TitleCss & """></td>"
		.echo "               <td height=""24"">ͼƬ CSS"
		.echo "                 <input name=""PhotoCss"" class=""textbox"" type=""text"" style=""width:70%;"" onBlur=""CheckBadChar(this,'ͼƬ��ʽ');"" value=""" & PhotoCss & """></td>"
		.echo "             </tr>"		
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""25"">��������"
		.echo "                <select name=""NavType"" id=""NavType"" class=""textbox"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>���ֵ���</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">���ֵ���</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
				   
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""25""> "
				 
				If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;""> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "           <tr class='tdbg'>"
		.echo "             <td width=""50%"" height=""25"">��������"
		.echo "               <select name=""MoreLinkType"" id=""MoreLinkType"" class=""textbox"" style=""width:70%;"" onchange=""SetMoreLinkStatus()"">"
				  
				  If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<option value=""0"" selected>��������</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">��������</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
				   
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""25""> "
				
				If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<div align=""left"" id=""LinkWord""> ")
					.echo ("  <input type=""text"" class=""textbox"" id=""MoreLinkWord"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """>")
					.echo ("</div>")
					.echo ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					.echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					.echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
					.echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
					.echo ("</div>")
				Else
				   .echo ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   .echo ("<input type=""text"" class=""textbox"" name=""MoreLinkWord"" id=""MoreLinkWord"" style=""width:70%;"">")
				   .echo ("</div>")
				   .echo ("<div align=""left"" id=""LinkPic""> ")
				   .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				   .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				   .echo ("</div>")
				End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""25"" colspan=""2"">�ָ�ͼƬ"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""ѡ��ͼƬ..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>"
		.echo "                <div align=""left""> </div></td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		.echo "                  </table>"	
		.echo "</form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 

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
Set KSCls = New GetNotRuleList
KSCls.Kesion()
Set KSCls = Nothing

Class GetNotRuleList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ChannelID,ClassID, IncludeSubClass,  OpenType, ArticleProperty, RowNumber, RowHeight, ShowNumPerRow, OrderStr,ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic,  TitleCss, PrintType,DocProperty,SpecialID,AjaxOut
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()

		With KS
		'�ж��Ƿ�༭
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  ClassID = "0"
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Conn.Close
			 Set Conn = Nothing
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			Conn.Close
			Set Conn = Nothing


			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetNotRuleList", ""),"}{/Tag}", "")
             'response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = Node.getAttribute("modelid")
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  OpenType           = Node.getAttribute("opentype")
			  DocProperty        = Node.getAttribute("docproperty")
			  RowNumber          = Node.getAttribute("rownumber")
			  ShowNumPerRow      = Node.getAttribute("shownumperrow")
			  RowHeight          = Node.getAttribute("rowheight")
			  OrderStr           = Node.getAttribute("orderstr")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			  MoreLinkType       = Node.getAttribute("morelinktype")
			  MoreLink           = Node.getAttribute("morelink")
			  SplitPic           = Node.getAttribute("splitpic")
			  TitleCss           = Node.getAttribute("titlecss")
			  PrintType          = Node.getAttribute("printtype")
			  AjaxOut            = Node.getAttribute("ajaxout")
            End If
			XMLDoc=Empty
			Set Node=Nothing
		End If
		If PrintType="" Then PrintType=1
		If RowNumber = "" Then RowNumber = 10
		If RowHeight = "" Then RowHeight = 20
		If ShowNumPerRow = "" Then ShowNumPerRow = 60
		If SpecialID=""  Then SpecialID=0
		If DocProperty = "" Or IsNull(DocProperty) Then DocProperty = "00000"
		If AjaxOut="" Then AjaxOut=false
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../KS_Inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../KS_Inc/jquery.js"" language=""JavaScript""></script>"
		%>
		<script language="javascript">
        $(document).ready(function(){
		 $("#ChannelID").change(function(){
		    $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			$.get('../../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty();
			  $("#ClassList").append("<option value='-1' style='color:red'>-��ǰ��Ŀ(ͨ��)-</option>");
			  $("#ClassList").append("<option value='0'>-��ָ����Ŀ-</option>");
			  $("#ClassList").append(unescape(data));
			  $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			 })
		   })	
		  $("#MutileClass").click(function(){
		    if ($(this).attr("checked")==true){
		      $("#ClassList").attr("multiple","multiple");
		      $("#ClassList").attr("style","height:60px");
		    }else{
			   $("#ClassList").removeAttr("multiple");
			}
		  });
		   <%if Instr(ClassID,",")<>0 Then%>
		   var searchStr="<%=ClassID%>";
		   $("#MutileClass").attr("checked",true);
		   $("#ClassList").attr("multiple","multiple");
		   $("#ClassList").attr("style","height:60px");
		   $("#ClassList>option").each(function(){
		     if($(this).val()=='-1' || $(this).val()=='0')
			  $(this).attr("selected",false)
			 else if (searchStr.indexOf($(this).val())!=-1)
			 {
			   $(this).attr("selected",true);
			 }
		   });
		  <%end if%>

		});		
		
		function SetStatus(Obj)
		{
		  if (Obj.value=='0')
		   {
			$("#MoreLinkArea").show();
		   }
		   else
		   $("#MoreLinkArea").hide();
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
			var ChannelID=$("#ChannelID").val();
			var ClassList='';
		    if ($("#MutileClass").attr("checked")==true){
				$("#ClassList>option[selected=true]").each(function(){
					if ($(this).val()!='0' && $(this).val()!='-1')
						if (ClassList=='') 
						 ClassList=$(this).val() 
						else
						 ClassList+=","+$(this).val();
					})
			 }else{
			    ClassList=$("#ClassList").val();
			 }
			 var DocProperty='';
			 $("input[name=DocProperty]").each(function(){
			     if ($(this).attr("checked")==true){
				  DocProperty=DocProperty+'1'
				 }else{
				  DocProperty=DocProperty+'0'
				 }      
			 })

			var SpecialID=$("select[name=SpecialID]").val();
			if (SpecialID==-1) ClassList=0;
			var OpenType=$("#OpenType").val();
			var ShowNumPerRow=$("#ShowNumPerRow").val();
			var RowHeight=$("#RowHeight").val();
			var RowNumber=$("#RowNumber").val();
			var OrderStr=$("#OrderStr").val();
			var Nav,NavType=$("#NavType").val();
			var MoreLink,MoreLinkType=$("#MoreLinkType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var PrintType=$("#PrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			
			if  (ShowNumPerRow=='')  ShowNumPerRow=10;
			if (RowHeight=='') RowHeight=20
			if  (RowNumber=='') RowNumber=30;
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			if  (MoreLinkType==0) MoreLink=$("input[name=MoreLinkWord]").val()
			else  MoreLink=$("input[name=MoreLinkPic]").val();
			
			
			var tagVal='{Tag:GetNotRuleList labelid="0" ajaxout="'+AjaxOut+'" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'"  docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" opentype="'+OpenType+'" rownumber="'+RowNumber+'" shownumperrow="'+ShowNumPerRow+'" rowheight="'+RowHeight+'" titlecss="'+TitleCss+'" navtype="'+NavType+'" nav="'+Nav+'" splitpic="'+SplitPic+'" printtype="'+PrintType+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'"}{/Tag}';
			$("#LabelContent").val(tagVal);
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
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetNotRuleList.asp"">"
	    .echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">�����ʽ"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">��ͨTable��ʽ</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">LI��ʽ</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">����Ajax���</label>"
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td colspan=""2"" height=""24"">ѡ��Χ"
		.echo "                <select name=""ChannelID"" id=""ChannelID"">"
		.echo "                 <option value=""0"">-����ģ��-</option>"
        .LoadChannelOption ChannelID
		.echo "                </select>"
		.echo "                <select class=""textbox"" name=""ClassList"" id=""ClassList"" onChange=""SetLabelFlag(this)"">"
		.echo "                 <option selected value=""-1"" style=""color:red"">- ��ǰ��Ŀ(ͨ��)-</option>"
						
						If ClassID = "0" Then
						   .echo ("<option  value=""0"" selected>- ��ָ����Ŀ -</option>")
						Else
						  .echo ("<option  value=""0"">- ��ָ����Ŀ -</option>")
					   End If
						  .echo Replace(KS.LoadClassOption(ChannelID),"value='" & ClassID & "'","value='" & ClassID &"' selected")
						  .echo "</select>"

						  
					If cbool(IncludeSubClass) = True Or LabelID = "" Then
					  Str = " Checked"
					Else
					  Str = ""
					End If
					  .echo "<input type='checkbox' name='MutileClass' id='MutileClass' value='1'>ָ������Ŀ"
					  .echo ("<input name=""IncludeSubClass"" type=""checkbox"" id=""IncludeSubClass"" value=""true""" & Str & ">")
			
		.echo "                  ��������Ŀ</div></td>"
		.echo "</td>"
		.echo "            </tr>"
		.echo "           <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">����ר��"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:70%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- ��ǰר��(ר��ҳͨ��)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- ��ָ��ר�� -</option>")
						   Else
						  .echo ("<option  value=""0"">- ��ָ��ר�� -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
		
		.echo "           </td><td>���Կ���"
		.echo "                <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">�Ƽ�</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox""  value=""2"""
		If mid(DocProperty,2,1) = 1 Then .echo (" Checked")
		  .echo ">����</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">ͷ��</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">����</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">�õ�</label>"
		.echo "                    <td>"
		.echo "              </td>"
		.echo "           </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">��ʾ����"
		.echo "                <input name=""RowNumber"" class=""textbox"" type=""text"" id=""RowNumber"" style=""width:70%;"" onBlur=""CheckNumber(this,'��ѯ����');"" value=""" & RowNumber & """></td>"
		.echo "              <td width=""50%"" height=""24"">ÿ������"
		.echo "                <input name=""ShowNumPerRow"" class=""textbox"" id=""ShowNumPerRow"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:30%;"" value=""" & ShowNumPerRow & """>&nbsp;<font color=red>��������֮��Ŀո�</font>"
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">�ĵ��о�"
		.echo "                <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight""    style=""width:70%;"" onBlur=""CheckNumber(this,'�ĵ��о�');"" value=""" & RowHeight & """></td>"
		.echo "              <td width=""50%"" height=""24"">���򷽷�"
		.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>�ĵ�ID(����)</option>")
					Else
					.echo ("<option value='ID Desc'>�ĵ�ID(����)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>�ĵ�ID(����)</option>")
					Else
					.echo ("<option value='ID Asc'>�ĵ�ID(����)</option>")
					End If
					If OrderStr = "Rnd" Then
					.echo ("<option value='Rnd' style='color:blue' selected>�����ʾ</option>")
					Else
					.echo ("<option value='Rnd' style='color:blue'>�����ʾ</option>")
					End If
					
					If OrderStr = "AddDate Asc" Then
					.echo ("<option value='AddDate Asc' selected>����ʱ��(����)</option>")
					Else
					.echo ("<option value='AddDate Asc'>����ʱ��(����)</option>")
					End If
					If OrderStr = "AddDate Desc" Then
					 .echo ("<option value='AddDate Desc' selected>����ʱ��(����)</option>")
					Else
					 .echo ("<option value='AddDate Desc'>����ʱ��(����)</option>")
					End If
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>�����(����)</option>")
					Else
					 .echo ("<option value='Hits Asc'>�����(����)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>�����(����)</option>")
					Else
					  .echo ("<option value='Hits Desc'>�����(����)</option>")
					End If

		.echo "         </select></td>"
		 .echo "           </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "              <td width=""50%"" height=""24"">��������"
		.echo "                <select class='textbox' name=""NavType"" id=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
				   If LabelID = "" Or CStr(NavType) = "0" Then
					.echo ("<option value=""0"" selected>���ֵ���</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">���ֵ���</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
		 .echo "               </select></td>"
		 .echo "             <td width=""50%"" height=""24"">"
			   If LabelID = "" Or CStr(NavType) = "0" Then
				  .echo ("<div align=""left"" id=""NavWord""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"">")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				Else
				  .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
				  .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" style=""width:70%;""> ֧��HTML�﷨")
				  .echo ("</div>")
				  .echo ("<div align=""left"" id=NavPic> ")
				  .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
				  .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				  .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.NaviPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				  .echo ("</div>")
				End If
		 .echo "             </td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg' id=""MoreLinkArea"""
		 .echo ">"
		 .echo "             <td width=""50%"" height=""24"">��������"
		 .echo "               <select class='textbox' name=""MoreLinkType"" id=""MoreLinkType"" style=""width:70%;"" onchange=""SetMoreLinkStatus()"">"
				  If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<option value=""0"" selected>��������</option>")
					.echo ("<option value=""1"">ͼƬ����</option>")
				   Else
					.echo ("<option value=""0"">��������</option>")
					.echo ("<option value=""1"" selected>ͼƬ����</option>")
				   End If
		.echo "                </select></td>"
		.echo "              <td width=""50%"" height=""24"">"
				If LabelID = "" Or CStr(MoreLinkType) = "0" Then
					.echo ("<div align=""left"" id=""LinkWord""> ")
					.echo ("  <input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """>")
					.echo ("</div>")
					.echo ("<div align=""left"" id=""LinkPic"" style=""display:none""> ")
					.echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"">")
					.echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
					.echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
					.echo ("</div>")
				Else
				   .echo ("<div align=""left"" id=""LinkWord"" style=""display:none""> ")
				   .echo ("<input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"">")
				   .echo ("</div>")
				   .echo ("<div align=""left"" id=""LinkPic""> ")
				   .echo ("<input type=""text"" class=""textbox"" readonly style=""width:120;"" id=""MoreLinkPic"" name=""MoreLinkPic"" value=""" & MoreLink & """>")
				   .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.MoreLinkPic);"" name=""Submit3"" value=""ѡ��ͼƬ..."">")
				   .echo ("&nbsp;<span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.MoreLinkPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>")
				   .echo ("</div>")
				End If
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">�ָ�ͼƬ"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:150px;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""ѡ��ͼƬ..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>"
		.echo "              </td><td>  " & KS.ReturnOpenTypeStr(OpenType) & " </td>"
		.echo "            </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""24"">������ʽ"
		.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""24""></td>"
		.echo "            </tr>"

		.echo "                  </table>"	
		.echo "    </form>"
		 
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 

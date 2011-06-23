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
Set KSCls = New GetRolls
KSCls.Kesion()
Set KSCls = Nothing

Class GetRolls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag, PicBorderColor
		Dim ClassID, IncludeSubClass, MarqueeDirection, MarqueeWidth, MarqueeHeight, PicWidth, PicHeight, PicStyle, OpenType, Num, TitleLen, ShowTitle, OrderStr, MarqueeSpeed, TitleCss,SpecialID,DocProperty
		Dim CurrPath, InstallDir
		Dim ChannelID:ChannelID=KS.G("ChannelID")
		CurrPath = KS.GetCommonUpFilesDir()
		FolderID = Request("FolderID")
		
		With KS
		'�ж��Ƿ�༭
		LabelID = Trim(KS.G("LabelID"))
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
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetRolls", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');</Script>")
			 response.End()
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			  ChannelID          = Node.getAttribute("modelid")
			  ClassID            = Node.getAttribute("classid")
			  IncludeSubClass    = Node.getAttribute("includesubclass")
			  MarqueeDirection   = Node.getAttribute("marqueedirection")
			  SpecialID          = Node.getAttribute("specialid")
			  DocProperty        = Node.getAttribute("docproperty")
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  OrderStr           = Node.getAttribute("orderstr")
			  MarqueeWidth       = Node.getAttribute("marqueewidth")
			  MarqueeHeight      = Node.getAttribute("marqueeheight")
			  OpenType           = Node.getAttribute("opentype")
			  ShowTitle          = Node.getAttribute("showtitle")
			  MarqueeSpeed       = Node.getAttribute("marqueespeed")
			  Num                = Node.getAttribute("num")
			  TitleLen           = Node.getAttribute("titlelen")
			  TitleCss           = Node.getAttribute("titlecss")
			  PicBorderColor     = Node.getAttribute("picbordercolor")
			End If
			Set Node=Nothing
			Set XMLDoc=Nothing		
		End If
		If MarqueeWidth = "" Then MarqueeWidth = 450
		If MarqueeHeight = "" Then MarqueeHeight = 120
		If MarqueeSpeed = "" Then MarqueeSpeed = 30
		If PicWidth = "" Then PicWidth = 130
		If PicHeight = "" Then PicHeight = 90
		If Num = "" Then Num = 10
		If TitleLen = "" Then TitleLen = 30
		If LabelID = "" Then ShowTitle = True
		If SpecialID="" Then SpecialID=0
		If ChannelID="" Then ChannelID=0
		If ShowTitle="" Then ShowTitle=True
		If DocProperty = "" Then DocProperty = "01000"
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/jQuery.js"" language=""JavaScript""></script>"
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
		   SetStatus($("input[name=ShowTitle]:checked").val());
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
		
		function SetStatus(Value)
		{ 
		 if (Value=='true'|| Value==true)
		  {
		   $("#titleArea").show()
		   }
		 else
		 {
		   $("#titleArea").hide()
		 }
		}
		
		function SetLabelFlag(Obj)
		{
		 if (Obj.value=='-1')
		  $("#LabelFlag").val(1);
		  else
		  $("#LabelFlag").val(0);
		}
		function SpecialChange(SpecialID)
		{
			if (SpecialID==-1) 
			  $("#ClassArea").hide();
			else
			  $("#ClassArea").show();	
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
			var SpecialID=$("select[name=SpecialID]").val();
			if (SpecialID==-1) ClassList=0;
			var DocProperty='';
			 $("input[name=DocProperty]").each(function(){
			     if ($(this).attr("checked")==true){
				  DocProperty=DocProperty+'1'
				 }else{
				  DocProperty=DocProperty+'0'
				 }      
			 })
			var MarqueeDirection=$("#MarqueeDirection").val();
			var OrderStr=$("#OrderStr").val();
			var MarqueeWidth=$("#MarqueeWidth").val();
			var MarqueeHeight=$("#MarqueeHeight").val();
			var OpenType=$("#OpenType").val();
			var PicWidth=$("#PicWidth").val();
			var PicHeight=$("#PicHeight").val();
			var MarqueeSpeed=$("#MarqueeSpeed").val();
			var Num=$("#Num").val();
			var TitleLen=$("#TitleLen").val();
			var TitleCss=$("#TitleCss").val();
			var PicBorderColor=$("#PicBorderColor").val();
			 
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
            var ShowTitle=$("input[name=ShowTitle]:checked").val();
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			var tagVal='{Tag:GetRolls labelid="0" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'"  docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" marqueewidth="'+MarqueeWidth+'" marqueeheight="'+MarqueeHeight+'" opentype="'+OpenType+'" showtitle="'+ShowTitle+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" num="'+Num+'" titlelen="'+TitleLen+'" titlecss="'+TitleCss+'" marqueedirection="'+MarqueeDirection+'" marqueespeed="'+MarqueeSpeed+'" picbordercolor="'+PicBorderColor+'"}{/Tag}';
		 
			$("#LabelContent").val(tagVal);
			$("#myform").submit();
			
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" onload=""SpecialChange(" & SpecialID &");"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """> "
		.echo " <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetRolls.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr id=""ClassArea"" class=tdbg>"
		.echo "              <td width=""50%"" height=""24"" colspan=""2"">ѡ��Χ"
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
		.echo "            </tr>"
		 .echo "            <tr class='tdbg'>"
		.echo "              <td  width=""50%"" height=""26"">����ר��"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:70%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- ��ǰר��(ר��ҳͨ��)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- ��ָ��ר�� -</option>")
						   Else
						  .echo ("<option  value=""0"">- ��ָ��ר�� -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
		.echo " </td>"
		.echo "              <td height=""26"" valign=""top"">���Կ���"
		.echo "                <label><input name=""DocProperty"" type=""checkbox"" value=""1"""
		If mid(DocProperty,1,1) = 1 Then .echo (" Checked")
		.echo ">�Ƽ�</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" Checked disabled value=""2"">����</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""3"""
		If mid(DocProperty,3,1) = 1 Then .echo (" Checked")
		  .echo ">ͷ��</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""4"""
		If mid(DocProperty,4,1) = 1 Then .echo (" Checked")
		  .echo ">����</label>"
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">�õ�</label>"
		.echo "</td>"
		.echo "            </tr>"

		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""26"">��������"
		 .echo "               <select class=""textbox"" name=""MarqueeDirection"" id=""MarqueeDirection"" style=""width:70%;"">"
					   If MarqueeDirection = "left" Then
						.echo ("<option value=""left"" selected>�������</option>")
					   Else
						.echo ("<option value=""left"">�������</option>")
					   End If
					   If MarqueeDirection = "right" Then
						.echo ("<option value=""right"" selected>���ҹ���</option>")
					   Else
						.echo ("<option value=""right"">���ҹ���</option>")
					   End If
					   If MarqueeDirection = "up" Then
						.echo ("<option value=""up"" selected>���Ϲ���</option>")
						Else
						.echo ("<option value=""up"">���Ϲ���</option>")
						End If
						If MarqueeDirection = "down" Then
						.echo ("<option value=""down"" selected>���¹���</option>")
						Else
						.echo ("<option value=""down"">���¹���</option>")
						End If
						
		.echo "                </select></td>"
		.echo "              <td height=""26"" valign=""top"">���򷽷�"
		.echo "                <select class='textbox' name='OrderStr' id='OrderStr' style=""width:75%;"">"
					
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
				   
		.echo "                </select></td>"
		.echo "            </tr>"
					
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""26"">"
		.echo KS.ReturnOpenTypeStr(OpenType)
		.echo "</td>"
		.echo "              <td height=""26"" valign=""top"">��ʾ����"
				   
				   If Cbool(ShowTitle) = True Then
					.echo ("<input name=""ShowTitle"" onclick=""SetStatus(true)"" type=""radio"" value=""true"" checked>��")
					.echo ("<input name=""ShowTitle"" onclick=""SetStatus(false)"" type=""radio"" value=""false"">��")
					Else
					  .echo ("<input type=""radio"" onclick=""SetStatus(true)""  value=""true"" name=""ShowTitle"">��")
					  .echo ("<input type=""radio"" onclick=""SetStatus(false)"" value=""false"" name=""ShowTitle"" checked>��")
				   End If
		.echo "        </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""26"">�������� ��� <input name=""MarqueeWidth"" class=""textbox"" type=""text"" id=""MarqueeWidth"" value=""" & MarqueeWidth & """ size=""6"" onBlur=""CheckNumber(this,'ռ�ݿ��');"">���� �߶�"
		.echo "                <input name=""MarqueeHeight"" class=""textbox"" type=""text"" id=""MarqueeHeight"" value=""" & MarqueeHeight & """ size=""6"" onBlur=""CheckNumber(this,'ռ�ݸ߶�');"">����"
		.echo " </td>"
		.echo "              <td height=""26"" valign=""top"">ͼƬ��С ͼƬ���"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ���');"">����  ͼƬ�߶�"
		.echo "                <input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ�߶�');"">"
		.echo "����</td>"
		.echo "            </tr>"
		  .echo "              <tr class='tdbg'>"
		  .echo "                <td colspan='2' height=""20"">�߿���ɫ"
		  .echo (" <input type=""text"" class=""textbox"" name=""PicBorderColor"" id=""PicBorderColor"" style=""width:120;"" value=""" & PicBorderColor & """>")
		  .echo (" <img border=0 id=""PicBorderColorShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & PicBorderColor & ";"" onClick=""Getcolor(this,'../../../ks_editor/SelectColor.asp','PicBorderColor');"" title=""ѡȡ��ɫ""> ������")
				
		  .echo "                </td>"
		  .echo "              </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""26"">�����ٶ�"
		.echo "              <input name=""MarqueeSpeed"" type=""text"" class=""textbox"" id=""MarqueeSpeed""    style=""width:75%;"" onBlur=""CheckNumber(this,'�����ٶ�');"" value=""" & MarqueeSpeed & """></td>"
		.echo "              <td height=""26"" valign=""top"">�г�����"
		.echo "              <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:75%;"" onBlur=""CheckNumber(this,'�г�����');"" value=""" & Num & """></td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg' id='titleArea'>"
		.echo "              <td height=""26"">��������"
		.echo "                <input name=""TitleLen"" id=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:75%;"" value=""" & TitleLen & """>              </td>"
		.echo "              <td height=""26"" valign=""top"">������ʽ"
		.echo "              <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:75%;"" value=""" & TitleCss & """></td>"
		.echo "            </tr>"
		.echo "         </table>"	
		.echo "  </form>"
		  
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"

		End With
		End Sub
End Class
%> 

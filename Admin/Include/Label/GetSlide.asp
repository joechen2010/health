<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@EasyTool.CN . QQ:111394,9537636
' Web: http://www.EasyTool.CN http://www.KeSion.cn
' Copyright (C) KeSion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetSlide
KSCls.KeSion()
Set KSCls = Nothing

Class GetSlide
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub KeSion()
		Dim TempClassList, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript, LabelFlag
		Dim ClassID, IncludeSubClass, PicWidth, PicHeight, Num, OpenType, ShowTitle,  TitleLen, TitleCss, ChangeTime,SlideType,SpecialID,DocProperty
		FolderID = Request("FolderID")
		Dim ChannelID:ChannelID=KS.G("ChannelID")
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
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			Descript = LabelRS("Description")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent = Replace(Replace(LabelContent, "{Tag:GetSlide", ""),"}{/Tag}", "")
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
			  SpecialID          = Node.getAttribute("specialid")
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  Num                = Node.getAttribute("num")
			  OpenType           = Node.getAttribute("opentype")
			  ShowTitle          = Node.getAttribute("showtitle")
			  TitleLen           = Node.getAttribute("titlelen")
			  TitleCss           = Node.getAttribute("titlecss")
			  ChangeTime         = Node.getAttribute("changetime")
			  SlideType          = Node.getAttribute("slidetype")
			  DocProperty        = Node.getAttribute("docproperty")
		   End If
		   Set Node=Nothing
		   Set XMLDoc=Nothing
		End If
		If ChannelID="" Then ChannelID=0
		If Num = "" Then Num = 5
		If TitleLen = "" Then TitleLen = 30
		If PicWidth = "" Then PicWidth = 200
		If PicHeight = "" Then PicHeight = 200
		If ChangeTime = "" Then ChangeTime = 5000
		If SlideType=0 Then SlideType=2
		If SpecialID="" Then SpecialID=0
		If ShowTitle="" Then ShowTitle=true
		If DocProperty = "" Then DocProperty = "00001"
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../KS_Inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
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
		  $("#SlideType>option[value=<%=SlideType%>]").attr("selected",true);
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
		{   if ($("input[name=LabelName]").val()=='')
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
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var Num=$("input[name=Num]").val();
			var OpenType=$("#OpenType").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var ChangeTime=$("input[name=ChangeTime]").val();
			var SlideType=$("#SlideType").val();
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
		    
			var ShowTitle=$("input[name=ShowTitle]:checked").val();
			if  (Num=='')  Num=10;
			if  (TitleLen=='') TitleLen=30;
			var tagVal='{Tag:GetSlide labelid="0" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" docproperty="'+DocProperty+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" num="'+Num+'" opentype="'+OpenType+'" showtitle="'+ShowTitle+'" titlelen="'+TitleLen+'" titlecss="'+TitleCss+'" changetime="'+ChangeTime+'" slidetype="'+SlideType+'"}{/Tag}';
		 
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
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """> "
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSlide.asp"">"
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
		.echo "              <td height=""30"">����ר��"
		.echo "                <select class=""textbox"" onchange=""SpecialChange(this.value)"" style=""width:35%;"" name=""SpecialID"" id=""SpecialID"">"
		.echo "                <option selected value=""-1"" style=""color:red"">- ��ǰר��(ר��ҳͨ��)-</option>"
						 If SpecialID = "0" Then
						   .echo ("<option  value=""0"" selected>- ��ָ��ר�� -</option>")
						   Else
						  .echo ("<option  value=""0"">- ��ָ��ר�� -</option>")
						  End If
		.echo KS.ReturnSpecial(SpecialID)
		.echo "</Select>"
        .echo "</td>"
		.echo "              <td width=""50%"" height=""24"">���Կ���"
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
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"" checked disabled>�õ�</label>"
		
		.echo "              </td>"
		.echo "</tr>"
		.echo "  <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">�õ�����"
		.echo " <select name=""SlideType"" id=""SlideType"">"
		.echo "<option value=""1"">��ͨJS�õ�</option>"
		.echo "<option value=""2"">flash�õ�1</option>"
		.echo "<option value=""3"">flash�õ�2(Sina)</option>"
		.echo "<option value=""4"">flash�õ�3(SOHU)</option>"
		.echo "<option value=""5"" style=""color:red"">flash�õ�4(6.5����)</option>"
		.echo "</select>"
					
        .echo " </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">��ѯ����"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:50;text-align:center"" onBlur=""CheckNumber(this,'ͼƬ����');"" value=""" & Num & """> ��</td>"
		.echo "              <td height=""30"">ͼƬ��С ��"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth2"" value=""" & PicWidth & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ���');"">"
		.echo "                ���� ��"
		.echo "                <input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight2"" value=""" & PicHeight & """ size=""6"" onBlur=""CheckNumber(this,'ͼƬ�߶�');"">"
		.echo "                ����</td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">��ʾ����"
					  
					If cbool(ShowTitle) = true Then
					.echo ("<input name=""ShowTitle"" id=""ShowTitle"" type=""radio"" value=""true"" checked>��ʾ��")
					.echo ("<input name=""ShowTitle"" id=""ShowTitle"" type=""radio"" value=""false"">����ʾ")
					Else
					  .echo ("<input type=""radio"" id=""ShowTitle"" value=""true"" name=""ShowTitle"">��ʾ��")
					  .echo ("<input type=""radio"" id=""ShowTitle"" value=""false"" name=""ShowTitle"" checked>����ʾ")
				   End If
				
		.echo "              </td>"
		 .echo "             <td height=""30"">" & KS.ReturnOpenTypeStr(OpenType) & "</td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">��������"
		 .echo "               <input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:70%;"" value=""" & TitleLen & """ > "
		 .echo "             </td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">һ������=����Ӣ���ַ�</font></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">������ʽ"
		 .echo "               <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">�Ѷ����CSS ,Ҫ��һ������ҳ��ƻ���</font></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">Ч���任���ʱ��"
		 .echo "               <input name=""ChangeTime"" class=""textbox"" type=""text"" id=""ChangeTime2"" value=""" & ChangeTime & """  onBlur=""CheckNumber(this,'���ʱ��');"">"
		 .echo "             </td>"
		 .echo "             <td height=""30""><font color=""#FF0000"">��λ:����</font></td>"
		 .echo "           </tr>"
		.echo "                  </table>"	
		.echo "  </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
	    End With
		End Sub
End Class
%> 

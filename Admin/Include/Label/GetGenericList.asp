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
Set KSCls = New GetGenericList
KSCls.Kesion()
Set KSCls = Nothing

Class GetGenericList
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
		Dim ChannelID,ClassID, IncludeSubClass, ShowClassName, OpenType, DocProperty, Num, RowHeight, TitleLen, OrderStr, ColNumber, ShowPicFlag, NavType, Navi, MoreLinkType, MoreLink, SplitPic, DateRule, DateAlign, TitleCss, DateCss,SpecialID,ShowNewFlag,ShowHotFlag, PrintType,AjaxOut,LabelStyle,IntroLen
		Dim PicWidth,PicHeight,PicStyle,PicBorderColor,PicSpacing
		Dim ButtonType,PriceType,ProductType,Discount
		Dim TypeID,ShowGQType
		FolderID = Request("FolderID")
		CurrPath = KS.GetCommonUpFilesDir()
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		
		
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
			 Conn.Close:Set Conn = Nothing
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');history.back();</Script>")
			 Exit Sub
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			LabelContent = LabelRS("LabelContent")
			LabelFlag = LabelRS("LabelFlag")
			LabelRS.Close:Set LabelRS = Nothing
			Conn.Close:Set Conn = Nothing
			LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetGenericList", ""),"}" & LabelStyle&"{/Tag}", "")
            ' response.write LabelContent
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
			  showclassname      = Node.getAttribute("showclassname")
			  DocProperty        = Node.getAttribute("docproperty")
			  OpenType           = Node.getAttribute("opentype")
			  Num                = Node.getAttribute("num")
			  RowHeight          = Node.getAttribute("rowheight")
			  TitleLen           = Node.getAttribute("titlelen")
			  IntroLen           = Node.getAttribute("introlen")
			  OrderStr           = Node.getAttribute("orderstr")
			  ColNumber          = Node.getAttribute("col")
			  ShowPicFlag        = Node.getAttribute("showpicflag")
			  NavType            = Node.getAttribute("navtype")
			  Navi               = Node.getAttribute("nav")
			  MoreLinkType       = Node.getAttribute("morelinktype")
			  MoreLink           = Node.getAttribute("morelink")
			  SplitPic           = Node.getAttribute("splitpic")
			  DateRule           = Node.getAttribute("daterule")
			  DateAlign          = Node.getAttribute("datealign")
			  TitleCss           = Node.getAttribute("titlecss")
			  DateCss            = Node.getAttribute("datecss")
			  SpecialID          = Node.getAttribute("specialid")
			  ShowNewFlag        = Node.getAttribute("shownewflag")
			  ShowHotFlag        = Node.getAttribute("showhotflag")
			  PrintType          = Node.getAttribute("printtype")
			  AjaxOut            = Node.getAttribute("ajaxout")
			  
			  PicWidth           = Node.getAttribute("picwidth")
			  PicHeight          = Node.getAttribute("picheight")
			  PicStyle           = Node.getAttribute("picstyle")
			  PicBorderColor     = Node.getAttribute("picbordercolor")
			  PicSpacing         = Node.getAttribute("picspacing")
			  
			  ButtonType         = Node.getAttribute("buttontype")
			  PriceType          = Node.getAttribute("pricetype")
			  ProductType        = Node.getAttribute("producttype")
			  Discount           = Node.getAttribute("discount")
			  
			  TypeID             = Node.getAttribute("typeid")
			  ShowGQType         = Node.getAttribute("showgqtype")

			End If
            
			Set Node=Nothing
			Set XMLDoc=Nothing
		End If
		If PrintType="" Then PrintType=1
		If Num = "" Then Num = 10
		If DocProperty = "" Then DocProperty = "00000"
		If RowHeight = "" Then RowHeight = 20
		If TitleLen = "" Then TitleLen = 30
		If IntroLen = "" Then IntroLen = 50
		If ColNumber = "" Then ColNumber = 1
		If SpecialID=""  Then SpecialID=0
		If ShowNewFlag="" Then ShowNewFlag=False
		If ShowHotFlag="" Then ShowHotFlag=False
		If PicWidth="" Then PicWidth=130
		If PicHeight="" Then PicHeight=90
		If PicStyle="" Then PicStyle=1
		If PicSpacing="" Then PicSpacing=2
		If ButtonType="" Then ButtonType=4
		If PriceType="" Then PriceType=0
		If ProductType="" Then ProductType=0
		If Discount="" or IsNull(Discount) Then Discount=true
		If TypeID="" Then TypeID=0
		If ShowGQType="" Or IsNull(ShowGQType) Then ShowGQType=true
		If AjaxOut="" Then AjaxOut=false
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@linkurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
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
		<script>
		var TempFieldStr='';
		var TempDateStr='';
		var TempTitleCss='';
		var GenericPicStyleOption="<option value='1'>��:����ʾ����ͼ</option><option value='2'>��:����ͼ+����:����</option><option value='3'>��:����ͼ+(����+���:����):����</option><option value='4'>��:(����+���:����)+����ͼ:����</option>";
						 
		$(document).ready(function(){
		  $("#ChannelID").change(function(){
		    $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			$.get('../../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty();
			  $("#ClassList").append("<option value='-1' style='color:red'>-��ǰ��Ŀ(ͨ��)-</option>");
			  $("#ClassList").append("<option value='0'>-��ָ����Ŀ-</option>");
			  $("#ClassList").append(unescape(data));
			  SetField($("#ChannelID").val());
			  SetPicStyle($("#ChannelID").val());
			  SetModelParam($("#ChannelID").val());
			  $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			 });
		    });
		   
		  $("#MutileClass").click(function(){
		    if ($(this).attr("checked")==true){
		      $("#ClassList").attr("multiple","multiple");
		      $("#ClassList").attr("style","height:60px");
			  $("#MoreLinkArea").hide();
		    }else{
			   $("#ClassList").removeAttr("multiple");
			   $("#MoreLinkArea").show();
			}
		  });
		  
		  SetPicStyle($("#ChannelID").val()); //�����ʽѡ��
		  $("#PicStyle").change(function(){
		    $("#ViewStylePicArea").html('<img style="border:1px solid #ccc;margin:5px" src="../../Images/View/S'+$(this).val()+'.gif" height="100" width="180" border="0">');
			if ($(this).val()==1){
			 if ($("#ShowPicTitleCss").html()!=null)	TempTitleStr=$("#ShowPicTitleCss").html();
			 $("#ShowPicTitleCss").empty();
			}else{
			$("#ShowPicTitleCss").html(TempTitleStr);
			}
		  });
		  $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S<%=PicStyle%>.gif" height="100" width="180" border="0">');
		  try{
		  $("#PicStyle>option[value=<%=PicStyle%>]").attr("selected",true);
		  }catch(e){
		  }
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
		  <%end if
		  If LabelID <> "" Then
		   .echo "$('#ChannelID').attr('disabled',true);"
		  End If
		  %>
		  TempFieldStr=$("#ShowFieldArea").html();
		  TempDateStr=$("#ShowTableDate").html();
		  TempTitleStr=$("#ShowTitleCss").html();
		  ChangeOutArea($("#PrintType>option[selected=true]").val());
		  

		})
		
		function SetField(channelid)
		{  
		   switch (parseInt(channelid)){
		    case 3:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@rank}')\" title=\"�Ƽ��Ǽ�\">�Ǽ�</li><li onclick=\"InsertLabel('{@downsize}')\" title=\"�����С\">�����С</li>");
			 break;
		    case 4:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@author}')\" title=\"����\">����</li><li onclick=\"InsertLabel('{@rank}')\" title=\"�ȼ�\">�ȼ�</li>");
			 break;
		    case 5:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@bigphoto}')\" title=\"��Ʒ��ͼ\">��Ʒ��ͼ</li><li onclick=\"InsertLabel('{@price_original}')\" title=\"ԭ�����ۼ۸�\">ԭ ʼ ��</li><li onclick=\"InsertLabel('{@price_market}')\" title=\"�г��۸�\">�� �� ��</li><li onclick=\"InsertLabel('{@price_member}')\" title=\"��Ա��\">�� Ա ��</li><li title=\"��ǰ���ۼ�\" onclick=\"InsertLabel('{@price}')\">��ǰ���ۼ�</li><li title=\"�ۿ���\" onclick=\"InsertLabel('{@discount}')\">�ۿ���</li><li title=\"��Ʒ�ͺ�\" onclick=\"InsertLabel('{@promodel}')\">��Ʒ�ͺ�</li><li title=\"���ͻ���\" onclick=\"InsertLabel('{@point}')\">���ͻ���</li>");
			 break;
		    case 7:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@movieact}')\" title=\"��Ҫ��Ա\">��Ҫ��Ա</li><li onclick=\"InsertLabel('{@moviedy}')\" title=\"ӰƬ����\">ӰƬ����</li><li title=\"����ʱ��\" onclick=\"InsertLabel('{@movietime}')\">����ʱ��</li><li title=\"ӰƬ����\" onclick=\"InsertLabel('{@movieyy}')\">ӰƬ����</li><li title=\"��������\" onclick=\"InsertLabel('{@moviedq}')\">��������</li><li title=\"�������\" onclick=\"InsertLabel('{@readpoint}')\">�������</li><li title=\"�Ƽ�����\" onclick=\"InsertLabel('{@rank}')\">�Ƽ�����</li>");
		     break;
		    case 8:
		     $("#ShowFieldArea").html(TempFieldStr+"<li onclick=\"InsertLabel('{@validdate}')\" title=\"��Ч��\">�� Ч ��</li><li onclick=\"InsertLabel('{@typeid}')\" title=\"�������\">�������</li><li title=\"��ϵ��\" onclick=\"InsertLabel('{@contactman}')\">�� ϵ ��</li><li title=\"��˾����\" onclick=\"InsertLabel('{@companyname}')\">��˾����</li><li title=\"����ʡ��\" onclick=\"InsertLabel('{@province}')\">����ʡ��</li><li title=\"���ڳ���\" onclick=\"InsertLabel('{@city}')\">���ڳ���</li><li title=\"��ϸ��ַ\" onclick=\"InsertLabel('{@address}')\">��ϸ��ַ<li title=\"��ϵ�绰\" onclick=\"InsertLabel('{@tel}')\">��ϵ�绰</li></li>");
		     break;
			
		   default:
		     $("#ShowFieldArea").html(TempFieldStr);
		   }
		   
		   if ($("#PrintType").val()==4){
		      $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
		  	  $.get('../../../plus/ajaxs.asp',{action:'GetFieldOption',channelid:channelid},function(data){
			  $("#ShowFieldArea").html($("#ShowFieldArea").html()+data)
			  $(top.frames['FrameTop'].document).find('#ajaxmsg').toggle();
			 });

		 }
		}
		
		function SetPicStyle(channelid)
		{ 
		   switch (parseInt(channelid))
		   { case 0:
		     case 1:
			 case 2:
			 case 3:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			  break;
			 case 4:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='5'>��:����ͼ+(����+���+����+ʱ��:����):����</option>");
			   $("#PicStyle").append("<option value='6'>��:����ͼ+(����+����:����+������):����</option>");
			   break;
			 case 5:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='7'>��:����ͼ+��ť</option>");
			   $("#PicStyle").append("<option value='8'>��:����ͼ+����+��ť:����</option>");
			   $("#PicStyle").append("<option value='9'>��:����ͼ+����+�۸�+��ť:����</option>");
			   $("#PicStyle").append("<option value='10'>��:����ͼ+(�۸�+��ť:����):����</option>");
			   $("#PicStyle").append("<option value='11'>��:(����ͼ+����)+(�۸�+��ť):����</option>");
			   $("#PicStyle").append("<option value='12'>��:����ͼ+(����+�۸�+��ť):����</option>");
			   break;
			 case 7:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='13'>��:����ͼ+(����+����+���+��ť):����</option>");
			   $("#PicStyle").append("<option value='14'>��:����ͼ+(����+���+����):����</option>");
			   $("#PicStyle").append("<option value='15'>��:����ͼ+(����+����+����+���+��ť):����</option>");
			   break;
			 case 8:
			   $("#PicStyle").empty();
			   $("#PicStyle").append(GenericPicStyleOption);
			   $("#PicStyle").append("<option value='16'>��:����ͼ+[(����+����+ʱ��)+���]:����</option>");
			   $("#PicStyle").append("<option value='17'>��:����ͼ+(����+���+����):����</option>");
			   break;
			 default:
			  break;
		   }
		}
		
		function SetModelParam(channelid)
		{
		  if (parseInt(channelid)<=1) 
		    $("#twbz").show() 
		  else $("#twbz").hide();
		  
		  if (parseInt(channelid)==5){
		   if (parseInt($("#PrintType").val())==2)   
		    $("#ModelParamArea").show();
		   else
		    $("#ModelParamArea").hide();
		   $("#ModelParamArea").empty();
		   $("#ModelParamArea").append("<tr class='tdbg'><td colspan='2'>��ť��ʽ <select style='width:160px' name='ButtonType' id='ButtonType'><option value='0'>����ʾ</option><option value='1'>��ʾ����ť</option><option value='2'>��ʾ�ղذ�ť</option><option value='3'>��ʾ���鰴ť</option><option value='4' selected>��ʾ����+�ղذ�ť</option><option value='5'>��ʾ����+���鰴ť</option><option value='6'>��ʾ�ղ�+���鰴ť</option><option value='7'>��ʾ����+����+�ղذ�ť</option></select> �۸���ʽ <select style='width:160px' class='textbox' name='PriceType' id='PriceType'><option value='0' selected>�Զ���ʾ</option><option value='8'>ֻ��ʾ��Ա��</option><option value='1'>ֻ��ʾԭʼ���ۼ�</option><option value='2'>ֻ��ʾ��ǰ���ۼ�</option><option value='3'>ԭʼ���ۼ�+��Ա��</option><option value='4'>��ǰ���ۼ�+��Ա��</option><option value='5'>��ʾ�г���+��ǰ���ۼ�</option><option value='6'>�г���+ԭʼ���ۼ�+��Ա��</option><option value='7'>�г���+ԭ��+��ǰ��+��Ա��</option></select> ��������<input name='ProductType' type='radio' value='0' Checked>����<input name='ProductType'  type='radio' value='1'>���� <input name='ProductType' type='radio' value='2'>�Ǽ� <input name='ProductType' type='radio' id='ProductType' value='3'>���� <label><input type='checkbox' name='Discount' id='Discount' value='true'><font color=blue>��ʾ�ۿ�</font></label></td></tr>");
		   $("#ButtonType>option[value=<%=ButtonType%>]").attr("selected",true);
		   $("#PriceType>option[value=<%=PriceType%>]").attr("selected",true);
		   $("input[name=ProductType][value=<%=ProductType%>]").attr("checked",true);
		   <%if Channelid=5 and cbool(Discount)=true then .echo "$('#Discount').attr('checked',true);" %>
		  }
		 else if(parseInt(channelid)==8){
		  $("#ModelParamArea").show();
		  $("#ModelParamArea").empty();
		  
		  $("#ModelParamArea").append("<tr class='tdbg'><td colspan='2'>�������� <%= Replace(Replace(KS.ReturnGQType(TypeID,1),"""","\"""),vbcrlf,"\n")%>  <label><input type='checkbox' name='ShowGQType' id='ShowGQType'>��ʾ��������</label></td></tr>");
		  $("#TypeID").css("width",120);
		  <%if ChannelID=8 Then%>
		  $("#TypeID>option[value=<%=ButtonType%>]").attr("selected",true);
		  <%if cbool(ShowGQType)=true then .echo "$('#ShowGQType').attr('checked',true);"%>
		  <%End If%>
		 }else{
		   $("#ModelParamArea").hide()
		  }
		}
		
		function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		 function setPos()
		 { if (document.all){
				$("#LabelStyle").focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("LabelStyle").selectionStart;
			  }
		 }
		 //����
		function InsertValue(Val)
		{  if (pos==null) {alert('���ȶ�λҪ�����λ��!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#LabelStyle");
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }
		
		function ChangeOutArea(Val)
		{
		 SetModelParam($("#ChannelID").val());
		 switch (parseInt(Val)){
		  case 2:
		   $("#DiyArea").hide();
		   $("#TableArea").hide();
		   $("#PicArea").show();
		   $("#ShowIntroArea").show();
   		   
		     $("#ShowPicTitleCss").html(TempTitleStr);
		     $("#ShowTitleCss").empty();
		   $("#ViewStylePicArea").html('<img style="border:1px outset #ccc;margin:5px" src="../../Images/View/S'+$("#PicStyle").val()+'.gif" height="100" width="180" border="0">');
		   break;
		  case 3:
		  case 4:
		  $("#DiyArea").show();
		  $("#TableArea").hide();
		  $("#PicArea").hide();
		  $("#ShowDiyDate").html(TempDateStr);
		  $("#ShowTableDate").html('')
		  $("#DateRule").attr("style","width:130px");
		  $("#ShowIntroArea").show();
		  break;
		  default :
		  $("#DiyArea").hide();
		  $("#PicArea").hide();
		  $("#TableArea").show();
		  $("#ShowTableDate").html(TempDateStr);
		  $("#ShowDiyDate").html('')
		  $("#DateRule").attr("style","width:268px");
		  $("#ShowIntroArea").hide();
		  $("#ShowTitleCss").html(TempTitleStr);
		  $("#ShowPicTitleCss").html('');
		  break;
		 }
		 SetField($("#ChannelID").val());
		 
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
			var IncludeSubClass=false,NavType=1;
			var ShowClassName,ShowPicFlag,ShowNewFlag,ShowHotFlag;
			var OpenType=$("#OpenType").val();
			var Num= $("#Num").val();
			var RowHeight=$("input[name=RowHeight]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var IntroLen=$("input[name=IntroLen]").val();
			var OrderStr=$("#OrderStr").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var Nav,NavType=$("select[name=NavType]").val();
			var MoreLink,MoreLinkType=$("select[name=MoreLinkType]").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var DateRule= $("#DateRule").val();
			var DateAlign=$("select[name=DateAlign]").val();
			var TitleCss=$("input[name=TitleCss]").val();
			var DateCss=$("input[name=DateCss]").val();
			var PicWidth=$("input[name=PicWidth]").val();
			var PicHeight=$("input[name=PicHeight]").val();
			var PicStyle=$("#PicStyle").val();
			var PicBorderColor=$("input[name=PicBorderColor]").val();
			var PicSpacing=$("input[name=PicSpacing]").val();
			
			var PrintType=$("#PrintType").val();
			var AjaxOut=false;
			if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			var IncludeSubClass=false;
			if ($("#IncludeSubClass").attr("checked")==true) IncludeSubClass=true;
			var ShowClassName =false;
			if ($("#ShowClassName").attr("checked")==true) ShowClassName = true
			var ShowPicFlag=false;
			if ($("#ShowPicFlag").attr("checked")==true)  ShowPicFlag= true
			var ShowHotFlag=false;
			if ($("#ShowHotFlag").attr("checked")==true)  ShowHotFlag= true
			var ShowNewFlag=false;
			if ($("#ShowNewFlag").attr("checked")==true)  ShowNewFlag= true
			   
			if  (Num=='')  Num=10;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (NavType==0) Nav=$("input[name=TxtNavi]").val();
			 else  Nav=$("input[name=NaviPic]").val();
			if  (MoreLinkType==0) MoreLink=$("input[name=MoreLinkWord]").val();
			else  MoreLink=$("input[name=MoreLinkPic]").val();
			
			var tagVal='{Tag:GetGenericList labelid="0" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" modelid="'+ChannelID+'" classid="'+ClassList+'" specialid="'+SpecialID+'" includesubclass="'+IncludeSubClass+'" docproperty="'+DocProperty+'" orderstr="'+OrderStr+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" introlen="'+IntroLen+'" rowheight="'+RowHeight+'" col="'+ColNumber+'" showclassname="'+ShowClassName+'" showpicflag="'+ShowPicFlag+'" shownewflag="'+ShowNewFlag+'" showhotflag="'+ShowHotFlag+'" navtype="'+NavType+'" nav="'+Nav+'" morelinktype="'+MoreLinkType+'" morelink="'+MoreLink+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" datecss="'+DateCss+'" picwidth="'+PicWidth+'" picheight="'+PicHeight+'" picstyle="'+PicStyle+'" picbordercolor="'+PicBorderColor+'" picspacing="'+PicSpacing+'"';
			if (ChannelID==5){
			 var ButtonType=$("#ButtonType").val();
			 var PriceType =$("#PriceType").val();
			 var ProductType=$("input[name=ProductType][checked=true]").val();
			 var Discount=false;
			 if ($("#Discount").attr("checked")==true)  Discount= true;
			 tagVal += ' buttontype="'+ButtonType+'" pricetype="'+PriceType+'" producttype="'+ProductType+'" discount="' + Discount + '"';
			}else if(ChannelID==8){
			 var TypeID=$("#TypeID").val();
			 var ShowGQType=false;
			 if($("#ShowGQType").attr("checked")==true) ShowGQType=true;
			 tagVal += ' typeid="'+TypeID+'" showgqtype="'+ShowGQType+'"';
			}
			tagVal  +='}'+$("#LabelStyle").val()+'{/Tag}';
			
			$("input[name=LabelContent]").val(tagVal);
			
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" onload=""SpecialChange(" & SpecialID &");"">"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""" & LabelFlag & """>"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo " <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetGenericList.asp"">"
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
		
		.echo "                ����</td>"
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
		.echo "<label><input name=""DocProperty"" type=""checkbox"" value=""5"""
		If mid(DocProperty,5,1) = 1 Then .echo (" Checked")
		  .echo ">�õ�</label>"
		
		.echo "              </td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" width=""50%"">���򷽷�"
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
		.echo "              <td height=""24"">" & KS.ReturnOpenTypeStr(OpenType) & "</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan='2'>�ĵ�����"
		.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num""    style=""width:40px;text-align:center"" onBlur=""CheckNumber(this,'�ĵ�����');"" value=""" & Num & """>�� ��������<input name=""TitleLen"" class=""textbox"" onBlur=""CheckNumber(this,'��������');"" type=""text""    style=""width:40px;;text-align:center"" value=""" & TitleLen & """> �о�"
		.echo "                <input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight2""    style=""width:40px;;text-align:center"" onBlur=""CheckNumber(this,'�ĵ��о�');"" value=""" & RowHeight & """>px ����<input type=""text"" class=""textbox"" onBlur=""CheckNumber(this,'��������');""  style=""width:40px;text-align:center"" value=""" & ColNumber & """ name=""ColNumber""> <font color=red>Tips:���Զ�����ʽ���,�о����������������ʽ�����</font></td>"
		.echo "            </tr>"
		
		
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">�����ʽ"
		.echo " <select class='textbox'  name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">�ı��б���ʽ(Table)</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">ͼƬ�б���ʽ(Table)</option>"
        .echo "  <option value=""3"""
		If PrintType=3 Then .echo " selected"
		.echo ">�Զ��������ʽ(�����Զ����ֶ�)</option>"
        .echo "  <option style='color:green' value=""4"""
		If PrintType=4 Then .echo " selected"
		.echo ">�Զ��������ʽ(���Զ����ֶ�)</option>"
        .echo "</select>"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">����Ajax���</label></td>"
		.echo "              <td><span id='ShowDiyDate'></span> <span id='ShowIntroArea'>�������<input type='text' class='textbox' style='text-align:center' name='IntroLen' id='IntroLen' value='" & IntroLen & "' size='4'></span></td>"
		.echo "            </tr>"
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">�� ��</li><li onclick=""InsertLabel('{@linkurl}')"">����URL</li> <li onclick=""InsertLabel('{@id}')"">�ĵ�ID</li><li onclick=""InsertLabel('{@title}')"">�� ��</li><li onclick=""InsertLabel('{@fulltitle}')"" style='color:red'>���ضϱ���</li><li onclick=""InsertLabel('{@classname}')"">��Ŀ����</li><li onclick=""InsertLabel('{@classurl}')"">��ĿURL</li> <li onclick=""InsertLabel('{@intro}')"">��Ҫ����</li><li onclick=""InsertLabel('{@photourl}')"">ͼƬ��ַ</li><li onclick=""InsertLabel('{@adddate}')"">���ʱ��</li><li onclick=""InsertLabel('{@inputer}')"">¼��Ա</li><li onclick=""InsertLabel('{@hits}')"">�����</li><li onclick=""InsertLabel('{@newimg}')"" title='��ʾ����ϢͼƬ��־' style='color:red;width:45px'>����ͼ</li><li onclick=""InsertLabel('{@hotimg}')"" title='��ʾ������ϢͼƬ��־' style='color:red;width:45px'>����ͼ</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>ʹ��˵�� :</font></strong><br />1��ѭ����ǩ[loop=n][/loop]�Կ���ʡ��,Ҳ����ƽ�г��ֶ�ԣ�<br />2�������ʽѡ�񲻴��Զ����ֶ����������Ч�ʸ��ڴ��Զ����ֶ�,����б�û���õ��Զ����ֶ���ѡ�񲻴��Զ����ֶ����</font></td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		
		.echo "           <tbody id='ModelParamArea'></tbody>"

		
		.echo "          <tbody id='TableArea'>"
		.echo "           <tr class=tdbg>"
		 .echo "             <td colspan=2 height=""30"">������ʾ "
				   If cbool(ShowClassName) = true Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"" checked>��ʾ��Ŀ</label>")
				   Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowClassName"" name=""ShowClassName"">��ʾ��Ŀ</label>")
				   End If
                    .echo "&nbsp;&nbsp;&nbsp;"
					 If cbool(ShowPicFlag) = True Then
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"" checked>��ͼ�ġ���־</label>")
					 Else
					  .echo ("<label id='twbz'><input type=""checkbox"" value=""true"" id=""ShowPicFlag"" name=""ShowPicFlag"">��ͼ�ġ���־</label>")
					 End If
				   .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowNewFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"" checked>�����ĵ���־</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowNewFlag"" name=""ShowNewFlag"">�����ĵ���־</label>")
					 End If
				 .echo "&nbsp;&nbsp;&nbsp;"
					 If  cbool(ShowHotFlag) = True Then
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"" checked>��ʾ�����ĵ���־</label>")
					 Else
					  .echo ("<label><input type=""checkbox"" value=""true"" id=""ShowHotFlag"" name=""ShowHotFlag"">��ʾ�����ĵ���־</label>")
					 End If
			   
		.echo "       ��</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">��������"
		.echo "                <select name=""NavType"" style=""width:70%;"" class='textbox' onchange=""SetNavStatus()"">"
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
		 .echo "           <tr  class=tdbg id=""MoreLinkArea"""
		 If Instr(ClassID,",")<>0 Then .echo " style='display:none'"
		 .echo ">"
		 .echo "             <td width=""50%"" height=""24"">��������"
		 .echo "               <select name=""MoreLinkType"" style=""width:70%;"" class='textbox' onchange=""SetMoreLinkStatus()"">"
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
					.echo ("  <input type=""text"" class=""textbox"" name=""MoreLinkWord"" style=""width:70%;"" value=""" & MoreLink & """> ֧��HTML�﷨")
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
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" colspan=""2"">�ָ�ͼƬ"
		.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
		.echo "                <input class='button' name=""SubmitPic"" class='button' onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""ѡ��ͼƬ..."">"
		.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>"
		.echo "                <div align=""left""> </div></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id='ShowTableDate'>���ڸ�ʽ"
		.echo "                <select class='textbox' style=""width:70%;"" name=""DateRule"" id=""DateRule"">"
		.echo KS.ReturnDateFormat(DateRule)
		.echo "                </select> </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">���ڶ���"
		.echo "                  <select class=""textbox"" name=""DateAlign"" id=""select3"" style=""width:70%;"">"
							
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
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowTitleCss"">������ʽ"
		.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
		.echo "              <td height=""24"">������ʽ"
		.echo "                <input name=""DateCss"" class=""textbox"" type=""text"" id=""DateCss"" style=""width:70%;"" value=""" & DateCss & """></td>"
		.echo "            </tr>"
		.echo "              </tbody>"



		.echo "           <tbody id='PicArea'>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">ͼƬ���� ��"
		.echo "                <input name=""PicWidth"" class=""textbox"" type=""text"" id=""PicWidth"" size='4' value=""" & PicWidth & """>px ��<input name=""PicHeight"" class=""textbox"" type=""text"" id=""PicHeight"" size='4' value=""" & PicHeight & """>px</td>"
		.echo "                <td colspan='2' rowspan='5' id='ViewStylePicArea'>ͼƬ��ʾ</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">��ʾ��ʽ"
		.echo "                <select class='textbox' style='width:230px' name=""PicStyle"" id=""PicStyle"">"
							.echo ("<option value=""1"">��:����ʾ����ͼ</option>")
							.echo ("<option value=""2"">��:����ͼ+����:����</option>")
							.echo ("<option value=""3"">��:����ͼ+(����+���:����):����</option>")
							.echo ("<option value=""4"">��:(����+���:����)+����ͼ:����</option>")
						 
		.echo "                </select> <font color=""#FF0000""> =>�ұ�Ч��Ԥ��</font></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">�߿���ɫ <input type=""text"" id=""PicBorderColor"" class=""textbox"" name=""PicBorderColor"" style=""width:120;"" value=""" & PicBorderColor & """><img border=0 id=""ColorThumbsBorderShow"" src=""../../images/rect.gif"" style=""cursor:pointer;background-Color:" & PicBorderColor & ";"" onClick=""Getcolor(this,'../../../ks_editor/SelectColor.asp','PicBorderColor');"" title=""ѡȡ��ɫ""> ������</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">ͼƬ���:<input type='text' class='textbox' name='PicSpacing' id='PicSpacing' value='" & PicSpacing & "' size='8' style='text-align:center'> px</td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"" id=""ShowPicTitleCss""></td>"
		.echo "            </tr>"
		.echo "           </tbody>"

		.echo "         </table>"			 
		.echo "    </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
End Class
%> 

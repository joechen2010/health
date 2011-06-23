<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Model
KSCls.Kesion()
Set KSCls = Nothing
Const ChannelNotOnStr="4,5,6,7,8,9,10"   '定义关闭的模块


Class Admin_Model
        Private KS,KSCls,I
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  CheckChannelStatus()
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
 
        Sub CheckChannelStatus()
		 conn.execute("update ks_channel set channelstatus=0 where channelid in(" & channelNotOnStr & ")")
		 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
		End Sub

		Public Sub Kesion()
		  'If Not KS.ReturnPowerResult(0, "model1") Then          '检查权限
			'Call KS.ReturnErr(1, "")
			'.End
		 ' End If
		  With Response
		    .Write "<html>"
			.Write "<title>模型基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../ks_inc/JQuery.js"" language=""JavaScript""></script>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "</head>"
			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""location.href='?action=Add';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=系统设置 >> <font color=red>系统模型管理</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加模型</span></li>"
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp?action=total';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/set.gif' border='0' align='absmiddle'>信息统计</span></li>"
			If KS.G("Action")="" Then
			.Write "<li class='parent' disabled><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>管理首页</span></li>"
			Else
			.Write "<li class='parent' onclick=""location.href='KS.Model.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>管理首页</span></li>"
			End IF
			.Write "</ul>"

		  Select Case KS.G("Action")
		   Case "SetChannelParam"
				If Not KS.ReturnPowerResult(0, "KSMM10005") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		             Call SetChannelParam()
			    End If 
		   Case "Edit","Add"
		       If KS.G("Action")="Add" Then
		       If Not KS.ReturnPowerResult(0, "KSMM10000") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  Else
		       If Not KS.ReturnPowerResult(0, "KSMM10001") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		            Call ChannelAddOrEdit()
			    End If
			  End If
		   Case "EditSave"
		        Call ChannelSave()
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KSMM10002") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call ChannelDel()
			    End If
		   Case "total"
		       If Not KS.ReturnPowerResult(0, "KSMM10004") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 .End
				Else  
		          Call Total()
			    End If
		   Case Else
		       Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With Response
			.Write "<script>"
			.Write "function document.onreadystatechange(){"
			.Write "parent.frames['BottomFrame'].Button1.disabled=true;"
			.Write "parent.frames['BottomFrame'].Button2.disabled=true;"
			.Write "}</script>"
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_Channel where channelid  not in(" & ChannelNotOnStr &") Order By ChannelID",conn,1,1
		    .Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write "  <td width='50' align=center>ID</td><td align=center>模型名称</td><td align=center>数据表</td><td align=center>类型</td><td align=center>项目名称</td><td align=center>项目单位</td><td align=center>状态</td><td align=center>WAP</td><td align=center>运行模式</td><td align=center>↓操作</td>"
			.Write "</tr>"
		  Do While Not RS.Eof 
		    .Write "<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.Write "<td align=center class='splittd'>" & RS("ChannelID")&"</td>"
			.Write "<td align=center class='splittd'>" & RS("ChannelName") &"</td>"
			.Write "<td class='splittd'>" & RS("ChannelTable") & "</td>"
			.Write "<td align='center' class='splittd'>"
			 If RS("ChannelID")>=100 Then
			  .Write "<font color=blue>自定义</font>"
			 Else
			  .Write "<font color='#ff0000'>系统</font>"
			 End If
			.Write "</td>"
			.Write "<td align=center class='splittd'>" & RS("ItemName") & "</td>"
			.Write "<td align=center class='splittd'>" & RS("ItemUnit") & "</td>"
			.Write "<td align=center class='splittd'>" 
			  If RS("ChannelStatus")="1" Then .Write "正常" Else .Write "<font color=red>已禁用</font>"
			.Write "</td>"
			.Write "<td align=center class='splittd'>" 
			  If RS("WapSwitch")="1" Then .Write "正常" Else .Write "<font color=red>已禁用</font>"
			.Write "</td>"
			.Write "<td align=center class='splittd'>"
			  If RS("FsoHtmlTF")="1" Then 
			  .Write "生成Html" 
			  ElseIf RS("FsoHtmlTF")="0" Then
			  .Write "<font color=red>动态asp</font>"
			   If RS("StaticTF")<>0 Then .Write "<i>(伪)</i>"
			  Else
			  .Write "<font color=blue>部分生成</font>"
			  If RS("StaticTF")<>0 Then .Write "<i>(伪)</i>"
			  End If
			.Write "</td>"

			.Write "<td align=center class='splittd'>"
			If rs("channelid")=1 or (Instr(channelNotOnStr,rs("channelid"))=0 and rs("channelid")<>10) then
			.Write "<a href='?action=Edit&ChannelID=" & rs("ChannelID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=系统设置 >> <font color=red>系统模型管理</font>';"">修改</a>｜"
			else
			.Write "<font color=#a7a7a7>修改</font>｜"
			end if
			 If RS("ChannelID")>=100 Then
			 .Write "<a href='?action=Del&ChannelID=" & rs("ChannelID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))'>删除</a>｜"
			 Else
			 .Write "<font color=#a7a7a7>删除</font>｜"
			 End If
			 
			 IF rs("channelid")=1 or (rs("ChannelID")<>6 and rs("channelid")<>10 and RS("ChannelID")<>"9" and Instr(channelNotOnStr,rs("channelid"))=0) then
			 .Write "<a href='#' onClick=""SelectObjItem1(this,'模型管理 >> <font color=red>模型字段管理</font>','Disabled','KS.Field.asp?ChannelID=" & rs("ChannelID") & "');"">字段管理</a>｜"
			 else
			 .Write "<font color=#a7a7a7>字段管理</font>｜"
			 end if
			 If Instr(channelNotOnStr,rs("channelid"))=0 or rs("channelid")=1 then
			 If RS("ChannelStatus")="1" Then .Write "<a href='?Action=SetChannelParam&Flag=ChannelOpenOrClose&ChannelID=" & RS("ChannelID") & "'>禁用</a>" Else .Write "<a href='?Action=SetChannelParam&Flag=ChannelOpenOrClose&ChannelID=" & RS("ChannelID") & "'>开启</a>"
			 else
			 .Write "<font color=#a7a7a7>开启</font>"
			 end if
			
			.Write "</td></tr>"
			RS.MoveNext 
		  Loop
		    .Write "</table>"
		   RS.Close:Set RS=Nothing
		    .Write "</body>"
			.Write "</html>"
		  End With
		End Sub
		
		Sub Total()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Channel Where ChannelStatus=1 order by channelid asc",conn,1,1
		   With Response
		  	.Write "<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.Write "<tr height='25' class='sort'>"
			.Write " <td align=center colspan=6>各模型信息统计</td>"
			.Write "</tr>"

		  Do While Not RS.Eof
			.Write "<tr height='25' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   If RS("ChannelID")=6 Then
		    .Write "<td width='140' align=center><img src='images/37.gif'>&nbsp;<b>" & RS("ChannelName") & "</b></td><td width='140'>专集总数：<font color='#ff0000'>" & Conn.Execute("select Count(SpecialID) from KS_MSSpecial")(0) & "</font> 张</td><td width='140'>歌曲总数：<font color='red'>" & Conn.Execute("select count(id) from KS_MSSongList")(0) & "</font> 首 </td><td width='140'>歌手总数：<font color='red'>" & Conn.Execute("select Count(NclassID) from KS_MSSinger")(0) & "</font> 位 </td><td>评论条数：<font color='red'>" & conn.execute("select count(*) from KS_MSSpecialComment")(0)  &" </font> 条 </td><td>待审歌词：<font color='red'>" & conn.execute("select count(*) from KS_MSUserAddWord")(0) & "</font> 条</td>"
		   ElseIf RS("ChannelID")=9 Then
		   ElseIf RS("ChannelID")=10 Then
		    .Write "<td width='140' align=center><img src='images/37.gif'>&nbsp;<b>" & RS("ChannelName") & "</b></td><td width='140'>简历总数：<font color='#ff0000'>" & Conn.Execute("select Count(ID) from KS_Job_Resume")(0) & "</font> 个</td><td width='140'>职位总数：<font color='red'>" & Conn.Execute("select count(id) from KS_Job_zw")(0) & "</font> 个 </td><td width='140'>单位总数：<font color='red'>" & Conn.Execute("select Count(ID) from KS_Job_Company")(0) & "</font> 家 </td>"
		   Else
			.Write "<td width='140' align=center><img src='images/37.gif'>&nbsp;<b>" & RS("ChannelName") & "</b></td><td width=150>频道总数: <font color=#ff0000>" & Conn.Execute("select count(id) from ks_class where channelid=" & RS("ChannelID") & " and tj=1")(0) & "</font> 个</td><td width=150>" & RS("ItemName") & "总数: <font color=blue>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where DelTF=0")(0) & " </font>" & RS("ItemUnit") & "</td><td>待审" & RS("ItemName") & ":<font color=green>" & conn.Execute("Select Count(ID) From " & RS("ChannelTable") & " Where  Verific=0")(0) & " </font>" & RS("ItemUnit") & "</td><td></td><td></td>"
		  End If
			.Write "</tr><tr><td colspan=10 background='images/line.gif'></td></tr>"
		    RS.MoveNext
		  Loop
		   .Write "</table>"
		  End With
		  RS.Close:Set RS=Nothing
		End Sub
		
		'模型设置
		Sub SetChannelParam()
		   With Response
			   Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
			   If ChannelID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="ChannelOpenOrClose" Then
			   If RS("ChannelStatus")=1 Then 
				  if conn.execute("select count(channelstatus) from ks_channel where channelstatus=1")(0)=1 then
				   rs.close:set rs=nothing
				   .Write "<script>alert('对不起，请至少保持一个模型是开启状态！');history.back();</script>"
				   else
					RS("ChannelStatus")=0 
				   End If
			   Else 
			    RS("ChannelStatus")=1
			   end if
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			 .Write "<script>parent.frames['LeftFrame'].location.reload();location.href='?';</script>"
		   End With
		End Sub
		
		Sub ChannelAddOrEdit()
		Dim TempStr,SqlStr, RS, InstallDir, FsoIndexFile,StaticTF, FsoIndexExt,FsoListNum,i,ThumbnailsConfig
		Dim ChannelName,ModelEname,FieldBit,FieldVal,ChannelTable,ChannelStatus,WapSwitch,WapSearchTemplate,ItemName,ItemUnit,FsoFolder,Descript
		Dim FsoHtmlTF,BasicType,MaxPerPage,UserTF,UserClassStyle,UserEditTF
		Dim UpFilesTF,UpfilesDir,UserUpFilesTF,UserUpfilesDir,UserSelectFilesTF,UpfilesSize,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType
		Dim  UserAddMoney,UserAddPoint,UserAddScore,RefreshFlag,InfoVerificTF,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,SearchTemplate,EditorType,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes
		Dim FsoContentRule,FsoClassListRule,FsoClassPreTag,LatestNewDay,PubTimeLimit
		Dim ChannelID:ChannelID = KS.ChkClng(KS.G("ChannelID"))
		
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select * from KS_Channel Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			ChannelName   = RS("ChannelName")
			ModelEname    = RS("ModelEname")
			ChannelTable  = RS("ChannelTable")
			ItemName      = RS("ItemName")
			ItemUnit      = RS("ItemUnit")
			if rs("basictype")<=3 then
			FieldBit      = Split(Split(RS("FieldBit"),"@@@")(0),"|")
			FieldVal      = Split(Split(RS("FieldBit"),"@@@")(1),"|")
			 if FieldVal(21)="" then  FieldVal(21)="附件上传"
			else
			FieldBit      = Split("1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|","|")
			FieldVal      = Split("1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|","|")
			end if
			ChannelStatus = RS("ChannelStatus")
			StaticTF      = RS("StaticTF")
			FsoFolder     = RS("FsoFolder")
			FsoListNum    = RS("FsoListNum")
			WapSwitch     = RS("WapSwitch")
			Descript      = RS("Descript")
			BasicType     = RS("BasicType")
			UserTF        = RS("UserTF")
			UserClassStyle= RS("UserClassStyle")
			UserEditTF    = RS("UserEditTF")
			FsoHtmlTF     = RS("FsoHtmlTF")
			UpFilesTF     = RS("UpFilesTF")
			UpfilesDir    = RS("UpfilesDir")
			UserUpFilesTF = RS("UserUpFilesTF")
			UserUpfilesDir= RS("UserUpfilesDir")
			UserSelectFilesTF =RS("UserSelectFilesTF")
			UpfilesSize   = RS("UpfilesSize")
			AllowUpPhotoType = RS("AllowUpPhotoType")
			AllowUpFlashType = RS("AllowUpFlashType")
			AllowUpMediaType = RS("AllowUpMediaType")
			AllowUpRealType  = RS("AllowUpRealType")
			AllowUpOtherType = RS("AllowUpOtherType")
			ThumbnailsConfig = RS("ThumbnailsConfig")
			
			UserAddMoney     = RS("UserAddMoney")
			UserAddPoint     = RS("UserAddPoint")
			UserAddScore     = RS("UserAddScore")
			RefreshFlag      = RS("RefreshFlag")
			MaxPerPage       = RS("MaxPerPage")
			InfoVerificTF    = RS("InfoVerificTF")
			VerificCommentTF = RS("VerificCommentTF")
			CommentVF        = RS("CommentVF")
			CommentLen       = RS("CommentLen")
			CommentTemplate  = RS("CommentTemplate")
			SearchTemplate   = RS("SearchTemplate")
			WapSearchTemplate= RS("WapSearchTemplate")
			EditorType       = RS("EditorType")
			DiggByVisitor    = RS("DiggByVisitor")
			DiggByIP         = RS("DiggByIP")
			DiggRepeat       = RS("DiggRepeat")
			DiggPerTimes     = RS("DiggPerTimes")
			FsoContentRule   = RS("FsoContentRule")
			FsoClassListRule = RS("FsoClassListRule")
			FsoClassPreTag   = RS("FsoClassPreTag")
			LatestNewDay     = RS("LatestNewDay")
			PubTimeLimit     = RS("PubTimeLimit")
		Else
		      FieldBit    = Split("0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|","|")
			  FieldVal    = Split("简短标题|归属栏目|完整标题|文章属性|转向链接|关 键 字|文章作者|文章来源|文章导读|文章内容|图片地址|上传图片|添加日期|文章等级|点 击 数|模板选择|自定义文件名|归属专题|收费选项|标题属性|立即发布|附件上传|签收选项|选择地区|||||||||||||||||||","|")
			  ChannelStatus =1 
			  ThumbnailsConfig="0.3|130|90"
			  UpfilesDir  = "Upfiles/"
			  UserUpfilesDir = "User/"
			  UpfilesSize = 1024
			  BasicType   = 1
			  MaxPerPage=20
			  FsoFolder="html/"
			  RefreshFlag=2
			  InfoVerificTF=1
			  VerificCommentTF=0
			  UserTF=1
			  UserEditTF=0
			  UserClassStyle=1
			  UpFilesTF=1
			  AllowUpPhotoType = "gif|jpg|png"
			  AllowUpFlashType = "swf"
			  AllowUpMediaType = "mid|mp3|wmv|asf|avi|mpg"
			  AllowUpRealType  = "ram|rm|ra"
			  AllowUpOtherType = "rar|doc|zip"
			  WapSwitch = 1
			  EditorType=1
			  FsoListNum=3
			  DiggByVisitor    = 0
			  DiggRepeat       = 0
			  DiggPerTimes     = 1
			  FsoClassPreTag="list"
			  FsoClassListRule = "1"
			  FsoContentRule   = "{$ClassEname}_{$ClassID}_"
			  LatestNewDay     = 3 
			  PubTimeLimit     = 20
		End If
			  ThumbnailsConfig=Split(ThumbnailsConfig,"|")
			  If Ubound(ThumbnailsConfig)<>2 Then
			   ThumbnailsConfig(0)=0.3
			   ThumbnailsConfig(1)=130
			   ThumbnailsConfig(2)=90
			  End IF
		With Response
		.Write "<html>"&_
		"<title>模型基本参数设置</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" &_
		"<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../KS_Inc/JQuery.js"" language=""JavaScript""></script>"&_
		"<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"<body onload=""ChangeType(" & BasicType & ");"">" &_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		"  <tr>"&_
		"	<td height='25' class='sort'>网站模型管理</td>"&_
		" </tr>"&_
		" <tr><td height=5></td></tr>"&_
		"</table>" & _
		"<div class=tab-page id=modelpane>"& _
		"<form id=""myform"" name=""myform"" method=""post"" action=""KS.Model.asp?Action=EditSave&ChannelID=" & ChannelID & """ onSubmit=""return(CheckForm())"">" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""modelpane"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>模型状态：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""ChannelStatus"" value=""1"" "
		If ChannelStatus = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""ChannelStatus"" value=""0"" "
		If ChannelStatus = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭</td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'> <div align=""right""><strong>WAP状态：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""WapSwitch"" value=""1"" "
		If WapSwitch = 1 Then .Write (" checked")
		.Write ">"
		.Write "正常"
		.Write "  <input type=""radio"" name=""WapSwitch"" value=""0"" "
		If WapSwitch = 0 Then .Write (" checked")
		.Write ">"
		.Write "关闭</td>"
		.Write "    </tr>"
%>
		<script type="text/javascript">
		 $(document).ready(function(){
		   $("input[name=FsoHtmlTF]").click(function(){
		     FsoDisplay();
			});
		   FsoDisplay();
		 });
		 function FsoDisplay()
		 {
		   var FsoHtmlTF=$("input[name=FsoHtmlTF][checked=true]").val();
		   if (FsoHtmlTF==0){
		    $("#fsoarea").hide();
			$("#staticarea").show();
		   }else if(FsoHtmlTF==1){
		    $("#fsoarea").show();
		    $("#staticarea").hide();
		   }else{
		    $("#fsoarea").show();
			$("#staticarea").show();
		   }
		 }
		 
		 function ChangeType(val)
		 {
		  switch (parseInt(val))
		   {case 1:
		   $("#ArticleType").show();
		   $("#PhotoType").hide();
		   $("#DownType").hide();
		    break;
		   case 2:
		   $("#ArticleType").hide();
		   $("#PhotoType").show();
		   $("#DownType").hide();
		   break;
		   case 3:
		   $("#ArticleType").hide();
		   $("#PhotoType").hide();
		   $("#DownType").show();
		   break;
		   }
		 }
		 function CheckForm()
		 { 
		  if ($("input[name=ChannelName]").val()=="")
		  {
		    $("input[name=ChannelName]").focus();
		   alert('请输入模型名称');
		   return false;
		  }
		  if ($("input[name=ModelEname]").val()=="")
		  {
		   $("input[name=ModelEname]").focus()
		   alert('请输入模型的目录名称');
		   return false;
		  }
		  if ($("input[name=ChannelTable]").val()=="")
		  {
		     $("input[name=ChannelTable]").focus()
			 alert('请输入数据名！');
			 return false;
		  }
		  if ($("input[name=ItemName]").val()=="")
		  {
		     $("input[name=ItemName]").focus()
			 alert('请输入项目名称！');
			 return false;
		  }
		  if ($("input[name=ItemUnit]").val()=="")
		  {
		     $("input[name=ItemUnit]").focus();
			 alert('请输入项目单位！');
			 return false;
		  }
		  if ($("input[name=FsoFolder]").val()=="")
		  {
		     $("input[name=FsoFolder]").focus();
			 alert('请输入模型目录！');
			 return false;
		  }
		  $("#myform").submit();
		 }
		 function GetTable(val)
		 { 
		    $.get('../plus/ajaxs.asp', { foldername: escape($('input[name=ChannelName]').val()), action: 'Ctoe' },function(data){
			$('input[name=ChannelTable]').val(unescape(data));
		    $('input[name=ModelEname]').val(unescape(data));
		    $('input[name=FsoFolder]').val('html/'+data+'/');
		    $('input[name=UserUpfilesDir]').val(data+'/');
		    $('input[name=UpfilesDir]').val('upfiles/'+data+'/');
			 });
		 }
		 function setFieldVal()
		 {	
		  var basictype=$('select[name=BasicType]').val();
		  var itemname=$('input[name=ItemName]').val();
		  if (basictype==1){
		   if (itemname=='') iteamname='文章';
		   var str="简短标题|归属栏目|完整标题|"+itemname+"属性|转向链接|关 键 字|"+itemname+"作者|"+itemname+"来源|"+itemname+"导读|"+itemname+"内容|图片地址|上传图片|添加日期|"+itemname+"等级|点 击 数|模板选择|自定义文件名|归属专题|收费选项|标题属性|立即发布|附件上传|签收选项|选择地区";
		   var val=str.split("|");
		   for(var i=0;i<val.length;i++)
		    $('input[name=V'+i+']').val(val[i]);
		  }
		  else if (basictype==2){
		   if (itemname==''||itemname==null) iteamname='图片';
		   var str=itemname+"名称|归属栏目|缩 略 图|"+itemname+"数量|"+itemname+"内容|"+itemname+"属性|关 键 字|"+itemname+"作者|"+itemname+"来源|"+itemname+"介绍|添加日期|"+itemname+"等级|"+itemname+"浏览数|"+itemname+"模板|自定义文件名|归属专题|收费选项|立即发布";
		   var val=str.split("|");
		   for(var i=0;i<val.length;i++)
		    $('input[name=VP'+i+']').val(val[i]);
		  }
		  else if(basictype==3){
		   if (itemname==''||itemname==null) iteamname='软件';
		   var str=itemname+"名称|归属栏目|设定地址数|下载地址|版本号|"+itemname+"属性|"+itemname+"性质|系统平台|"+itemname+"图片|上传图片|关 键 字|作者开发商|"+itemname+"来源|上传"+itemname+"|"+itemname+"介绍|演示地址|注册地址|解压密码|添加日期|"+itemname+"等级|浏 览 数|"+itemname+"模板|文 件 名|所属专题|收费选项|立即发布";
		   var val=str.split("|");
		   for(var i=0;i<val.length;i++)
		    $('input[name=VD'+i+']').val(val[i]);
		  }
		 }
		</script>
		<style type="text/css">
		 .textbox{
		 border:0px;border-bottom:1px solid #000;width:60px;background:transparent
		 }
		</style>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>模型名称：</strong></div></td>      
			<td height="30"> <input name="ChannelName" type="text" <%If KS.G("Action")<>"Edit" Then Response.Write " onkeyup='GetTable(this.value)'"%> value="<%=ChannelName%>" size="30"></td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>模型目录：</strong></div></td>      
			<td height="30"> <input name="ModelEname" type="text"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%> value="<%=ModelEname%>" size="30">*只能用字母和数字的组合，且不能修改</td> 
		</tr>

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>数据表名称：</strong></div></td>     
		<td height="30"><%If KS.G("Action")="Add" Then Response.Write " KS_U_" %><input name="ChannelTable" id='ChannelTable' type="text" value="<%=ChannelTable%>" size="24"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>><br><font color=#ff0000>说明：创建数据表后无法修改，并且用户创建的数据表以"KS_U_"开头</font></td>   
		</tr> 
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>基 类 型：</strong></div></td>      
			<td height="30"> 
			<select name="BasicType" id="BasicType" onchange="ChangeType(this.value);setFieldVal();"<%If KS.G("Action")="Edit" Then Response.Write " Disabled"%>>
			 <option value=1<%if BasicType="1" Then Response.Write " selected"%>>文章类型</option>
			 <option value=2<%if BasicType="2" Then Response.Write " selected"%>>图片类型</option>
			 <option value=3<%if BasicType="3" Then Response.Write " selected"%>>软件类型</option>
			 <%If KS.G("Action")="Edit" Then%>
			 <option value=4<%if BasicType="4" Then Response.Write " selected"%>>Flash类型</option>
			 <option value=5<%if BasicType="5" Then Response.Write " selected"%>>商城类型</option>
			 <option value=6<%if BasicType="6" Then Response.Write " selected"%>>音乐类型</option>
			 <option value=7<%if BasicType="7" Then Response.Write " selected"%>>影视类型</option>
			 <option value=8<%if BasicType="8" Then Response.Write " selected"%>>供求类型</option>
			 <option value=9<%if BasicType="9" Then Response.Write " selected"%>>考试类型</option>
			 <%End If%>
			</select>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>项目名称：</strong></div></td>      
			<td height="30"> <input name="ItemName"<%If KS.G("Action")<>"Edit" Then response.write " onchange=""setFieldVal();"""%> id="ItemName" type="text" value="<%=ItemName%>" size="30">*如：文章、图片、软件等项</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>项目单位：</strong></div></td>      
			<td height="30"> <input name="ItemUnit" type="text" value="<%=ItemUnit%>" size="8">*如：篇、个、本等</td> 
		</tr>
		<%if KS.ChkClng(BasicType)<4 then%>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>系统字段：</strong></div>说明：含<font color=red>U</font>表示同时作用于前台会员中心的发布项,含<font color=green>C</font>表示允许采集项</td>      
			<td height="30"> 
			
			   <table id='ArticleType' border='0' cellspacing='0' cellspadding='0'>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='A(0)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(0)%>" name="V0" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(1)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(1)%>" name="V1" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(2)'<%if FieldBit(2)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(2)%>" name="V2" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(3)'<%if FieldBit(3)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(3)%>" name="V3" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(4)'<%if FieldBit(4)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(4)%>" name="V4" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='A(5)'<%if FieldBit(5)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(5)%>" name="V5" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(6)'<%if FieldBit(6)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(6)%>" name="V6" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(7)'<%if FieldBit(7)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(7)%>" name="V7" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(8)'<%if FieldBit(8)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(8)%>" name="V8" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(9)'<%if FieldBit(9)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(9)%>" name="V9" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='A(10)'<%if FieldBit(10)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(10)%>" name="V10" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(11)'<%if FieldBit(11)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(11)%>" name="V11" />(<Font color=red>U</Font>)</td>
				  <td width=120><input type='checkbox' value='1' name='A(12)'<%if FieldBit(12)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(12)%>" name="V12" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(13)'<%if FieldBit(13)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(13)%>" name="V13" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(14)'<%if FieldBit(14)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(14)%>" name="V14" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='A(15)'<%if FieldBit(15)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(15)%>" name="V15" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(16)'<%if FieldBit(16)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(16)%>" name="V16" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(17)'<%if FieldBit(17)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(17)%>"  name="V17" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(18)'<%if FieldBit(18)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(18)%>"  name="V18" /></td>
				  <td width=120><input type='checkbox' value='1' name='A(19)'<%if FieldBit(19)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(19)%>" name="V19" /></td>
				</tr>
				<tr>
				<td width=120><input type='checkbox' value='1' name='A(20)'<%if FieldBit(20)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(20)%>" name="V20" /></td>
				<td><input type='checkbox' value='1' name='A(21)'<%if FieldBit(21)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(21)%>" name="V21" />(<font color=red>U</font> 配合内容编辑器使用)</td>
				<td>
				<%
				Dim SignTextOption:SignTextOption=FieldVal(22)
				Dim AreaOption:AreaOption=FieldVal(23)
				if KS.IsNul(SignTextOption) Then SignTextOption="签收选项"
				If KS.IsNul(AreaOption) Then AreaOption="选择地区"
				%>
				<input type='checkbox' value='1' name='A(22)'<%if FieldBit(22)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=SignTextOption%>" name="V22" /></td>
				<td><input type='checkbox' value='1' name='A(23)'<%if FieldBit(23)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=AreaOption%>" name="V23" />(<font color=red>U</font>)</td>
				</tr>
			   </table>
			   
			   <table id='PhotoType' width='100%' style='display:none' border='0' cellspacing='0' cellspadding='0'>
			    <tr>
				  <td width=125 nowrap="nowrap"><input type='checkbox' value='1' name='P(0)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(0)%>" name="VP0" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=125 nowrap="nowrap"><input type='checkbox' value='1' name='P(1)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(1)%>" name="VP1" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=125 nowrap="nowrap"><input type='checkbox' value='1' name='P(2)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(2)%>" name="VP2" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=125 nowrap="nowrap"><input type='checkbox' value='1' name='P(3)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(3)%>" name="VP3" />(<Font color=red>U</Font>)</td>
				  <td width=125 nowrap="nowrap"><input type='checkbox' value='1' name='P(4)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(4)%>" name="VP4" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='P(5)'<%if FieldBit(5)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(5)%>" name="VP5" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(6)'<%if FieldBit(6)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(6)%>" name="VP6" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='P(7)'<%if FieldBit(7)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(7)%>" name="VP7" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='P(8)'<%if FieldBit(8)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(8)%>" name="VP8" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='P(9)'<%if FieldBit(9)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(9)%>" name="VP9" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='P(10)'<%if FieldBit(10)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(10)%>" name="VP10" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(11)'<%if FieldBit(11)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(11)%>" name="VP11" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(12)'<%if FieldBit(12)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(12)%>" name="VP12" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(13)'<%if FieldBit(13)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(13)%>" name="VP13" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(14)'<%if FieldBit(14)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(14)%>" name="VP14" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='P(15)'<%if FieldBit(15)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(15)%>" readonly name="VP15" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(16)'<%if FieldBit(16)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(16)%>" readonly name="VP16" /></td>
				  <td width=120><input type='checkbox' value='1' name='P(17)'<%if FieldBit(17)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(17)%>" name="VP17" /></td>
				  <td width=120></td>
				  <td width=120></td>
				  <td width=120></td>
				</tr>
			   </table>
			   
			   <table id='DownType' style='display:none' border='0' cellspacing='0' cellspadding='0'>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(0)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(0)%>" name="VD0" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(1)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(1)%>" name="VD1" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(2)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(2)%>" name="VD2" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(3)' checked disabled><input class="textbox" type="text" value="<%=FieldVal(3)%>" name="VD3" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(4)'<%if FieldBit(4)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(4)%>" name="VD4" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(5)'<%if FieldBit(5)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(5)%>" name="VD5" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(6)'<%if FieldBit(6)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(6)%>" name="VD6" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(7)'<%if FieldBit(7)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(7)%>" name="VD7" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(8)'<%if FieldBit(8)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(8)%>" name="VD8" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(9)'<%if FieldBit(9)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(9)%>" name="VD9" />(<Font color=red>U</Font>)</td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(10)'<%if FieldBit(10)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(10)%>" name="VD10" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120 nowrap="nowrap"><input type='checkbox' value='1' name='D(11)'<%if FieldBit(11)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(11)%>" name="VD11" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(12)'<%if FieldBit(12)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(12)%>" name="VD12" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(13)'<%if FieldBit(13)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(13)%>" name="VD13" />(<Font color=red>U</Font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(14)'<%if FieldBit(14)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(14)%>" name="VD14" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(15)'<%if FieldBit(15)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(15)%>" name="VD15" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(16)'<%if FieldBit(16)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(16)%>" name="VD16" />(<Font color=red>U</Font>、<font color=green>C</font>)</td>
				  <td width=120><input type='checkbox' value='1' name='D(17)'<%if FieldBit(17)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(17)%>" name="VD17" />(<Font color=red>U</Font>、<font color=green>C</font>) </td>
				  <td width=120><input type='checkbox' value='1' name='D(18)'<%if FieldBit(18)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(18)%>" name="VD18" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(19)'<%if FieldBit(19)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(19)%>" name="VD19" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(20)'<%if FieldBit(20)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(20)%>" name="VD20" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(21)'<%if FieldBit(21)=1 then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(21)%>" name="VD21" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(22)'<%if FieldBit(22)="1" then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(22)%>" name="VD22" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(23)'<%if FieldBit(23)="1" then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(23)%>" name="VD23" /></td>
				  <td width=120><input type='checkbox' value='1' name='D(24)'<%if FieldBit(24)="1" then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(24)%>" name="VD24" /></td>
				</tr>
			    <tr>
				  <td width=120><input type='checkbox' value='1' name='D(25)'<%if FieldBit(25)="1" then Response.Write(" checked") %>><input class="textbox" type="text" value="<%=FieldVal(25)%>" name="VD25" /></td>
				  <td width=120></td>
				  <td width=120></td>
				  <td width=120></td>
				  <td width=120></td>
				</tr>
			   </table>
			</td> 
		</tr>
          <%end if%>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>模型描述：</strong></div></td>      
			<td height="30"> <textarea name="Descript" rows=4 cols=80><%=Descript%></textarea></td> 
		</tr>
		</table>
		</div>
		
		
		<%
		.Write " <div class=tab-page id=fso-page>"
		.Write "  <H2 class=tab>生成选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""fso-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td height=""30"" class='clefttitle'><div align=""right""><strong>本模型运行模式：</strong></div></td>"
		.Write "      <td height=""30"">"
		
		.Write "  <input type=""radio"" name=""FsoHtmlTF"" value=""0"" "
		If FsoHtmlTF = 0 Then .Write (" checked")
		.Write ">动态asp<br/>"

		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""1"" "
		If FsoHtmlTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "栏目页及内容页都生成HTML<br/>"
		
		.Write "<input type=""radio"" name=""FsoHtmlTF"" value=""2"" "
		If FsoHtmlTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "栏目页不生成,内容页生成HTML(<font color=red>推荐</font>)<br/>"
		
		
		.Write "     </td>"
		.Write "    </tr>"
		.Write "    <tbody id='staticarea'>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td height='25' class='clefttitle' align=""right""><strong>伪静态设置：</strong></td>"
		.Write "      <td>"
		
		.Write "  <input type=""radio"" name=""StaticTF"" value=""0"" "
		If StaticTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "不启用"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""1"" "
		If StaticTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(带问号,不需要装组件)"
		.Write "  <input type=""radio"" name=""StaticTF"" value=""2"" "
		If StaticTF = 2 Then .Write (" checked")
		.Write ">"
		.Write "伪静态(需要装ISAPI_Rewrite组件)"
        .Write "<br /><font color=blue>这里需要设置不生成静态才有效,建议流量大的网站直接启用全部生成静态,而不是使用伪静态</font>"
		.Write "      </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		.Write "  <tbody id='fsoarea'>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' align=""right""><strong>后台添加" & TempStr &"，同时发布选项：</strong></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""RefreshFlag"" value=""1"" "
		If RefreshFlag = 1 Then .Write (" checked")
		.Write ">"
		.Write "仅发布内容页 <br>"
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""2"" "
		If RefreshFlag = 2 Then .Write (" checked")
		.Write ">发布栏目页+内容页<font color=red>(建议)</font><br>"		
		.Write "          <input type=""radio"" name=""RefreshFlag"" value=""3"" "
		If RefreshFlag = 3 Then .Write (" checked")
		.Write ">发布首页+栏目页+内容页"
		.Write "        </td>"
		.Write "    </tr>"	
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'><div align=""right""><strong>自动生成列表分页数：</strong></div></td>"
		.Write "      <td height=""30""><input type='text' value='" & FsoListNum & "' name='FsoListNum' size='6' style='text-align:center'><font color=blue>这里设置生成栏目列表分页时自动生成的分页数，如果你的网站数据量较大，建议输入一个较小的数字，小数据量的网站可以不用限制，直接设置为0</font></td>"
		.Write "    </tr>"	
	    .Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='CleftTitle' align='right'><div><strong>生成的总目录：</strong></div></td>"      
		.Write "	<td height='30'> <input name='FsoFolder' type='text' value='" & FsoFolder & "' size='30'>*用于生成静态html存放的目录，只能以字母和数字的组合,必须以""/""结束</td> "
		.Write "</tr>"
		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='CleftTitle' align='right'><div><strong>生成的栏目页规则：</strong></div></td>"   
		.Write "<td>"   
		.Write "<input type=""radio"" name=""FsoClassListRule"" value=""1"" "
		If FsoClassListRule = 1 Then .Write (" checked")
		.Write ">按目录级别顺序结构生成列表页<br>"
		.Write " &nbsp;<font color=blue>如：第1页为/article/aaa/bbb/ccc/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/article/aaa/bbb/ccc/index_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""2"" "
		If FsoClassListRule = 2 Then .Write (" checked")
		.Write ">所有栏目页都生成在模型总生成目录下面<font color=red>(有利于SEO)</font><br>"
		.Write " &nbsp;<font color=green>如栏目ID号为100则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_100.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_100_2.html</font>"
		.Write " <br>&nbsp;<font color=green>如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/list_101_2.html</font>"
		.Write "  <br><input type=""radio"" name=""FsoClassListRule"" value=""3"" "
		If FsoClassListRule = 3 Then .Write (" checked")
		.Write ">本模型下的一级栏目生成在本频道下的Index.html,子栏目按如下规则生成<br>"
		.Write " &nbsp;<font color=green>如一级栏目 ""教育频道"" 英文名称：""edu"",那么生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/index.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/index_2.html</font>"
		.Write " <br>&nbsp;<font color=green>二级及以上的栏目(即""教育频道"")下的栏目,如栏目ID号为101则生成如下：</font><font color=blue><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第1页为/总生成目录/edu/list_101.html<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;第2页为/总生成目录/edu/list_101_2.html</font>"
		
		.Write "</td>"
		.Write "	<td height='30'> </td> "
		.Write "</tr>"

		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='CleftTitle' align='right'><div><strong>生成的列表页的前缀字符：</strong></div></td>"      
		.Write "	<td height='30'> <input name='FsoClassPreTag' type='text' value='" & FsoClassPreTag & "' size='30'>*如list,show等</td> "
		.Write "</tr>"


		.Write "<tr valign='middle' class='tdbg' onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""> "     
		.Write "	<td width='160' height='30' class='CleftTitle' align='right'><div><strong>生成的内容页目录规则：</strong></div></td>"      
		.Write "	<td height='30'> <input name='FsoContentRule' type='text' value='" & FsoContentRule & "' size='30'>&nbsp;"
		.Write "     <select name='srule' onchange='if (this.value!=""""){ $(""input[name=FsoContentRule]"").val(this.value);}'><option value=''>------快速选择内容页生成结构------</option>"
		.Write "     <option value='View_'>View_</option>"
		.Write "     <option value='{$ClassDir}'>{$ClassDir}</option>"
		.Write "     <option value='{$ChannelEname}/{$ClassEname}_{$ClassID}_'>{$ChannelEname}/{$ClassEname}_{$ClassID}_</option>"
		.Write "     <option value='{$ClassEname}_{$ClassID}_'>{$ClassEname}_{$ClassID}_(推荐)</option>"
		.Write "     <option value='{$ClassDir}{$ClassEname}_{$ClassID}_'>{$ClassDir}{$ClassEname}_{$ClassID}_</option>"
		.Write "   </select><br><font color=red>可选项（允许留空）</font><br>可用标签：一级频道名称{$ChannelEname},栏目路径{$ClassDir} {$ClassID} {$ClassEname}<br> "
		.Write "   <br>"
		.Write " </td></tr>"
		.Write "</tbody>"
        .Write "</table>"
		.Write "</div>"
		
		.Write " <div class=tab-page id=upfile-page>"
		.Write "  <H2 class=tab>上传选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""upfile-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>管理员是否允许上传文件：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""UpFilesTF"" value=""1"" "
		If UpFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UpFilesTF"" value=""0"""
		If UpFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</td>"
		.Write "    </tr>"
		.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
		.Write "            <td height='30' width='160' align='right' class='clefttitle'><strong>缩略图选项：</strong><br><font color=blue>当基本信息设置开启自动生成缩略图功能时才有效</font></td>" & vbCrLf
		.Write "            <td height='28'>&nbsp;"
					    
		.Write "黄金分割点：<input type='text' value='" & ThumbnailsConfig(0) & "' name='GoldenPoint' size='4' style='text-align:center'> 宽度：<input type='text' value='" & ThumbnailsConfig(1) & "' name='ThumbsWidth' size='8' style='text-align:center'>px 高度：<input type='text' value='" & ThumbnailsConfig(2) & "' name='ThumbsHeight' size='8' style='text-align:center'>px"
		.Write "              <br/> <font color=red>tips:如果高度设置为0,则生成的高度将由您设置的宽度自动约束决定(类似photoshop软件的自动约束)</font></td>"
		.Write "          </tr>" & vbCrLf
		.Write "    <tr class='tdbg' style='display:none'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>后台文件上传目录：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""UpfilesDir"" type=""text"" id=""UpfilesDir"" value=""" & UpfilesDir & """ size=""30""></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>是否允许会员上传文件：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""UserUpFilesTF"" value=""1"" "
		If UserUpFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UserUpFilesTF"" value=""0"""
		If UserUpFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg' style='display:none'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>会员文件上传目录：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""UserUpfilesDir"" type=""text"" id=""UserUpfilesDir"" value=""" & UserUpfilesDir & """ size=""30""><br><b>提示：</b><br><font color=red>1、会员目录构成规则：系统设置的总上传目录/User/会员名称;<br>2、上传目录必须以/结束;</font></td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许会员选择上传文件：</strong></div></td>"
		.Write "      <td height=""30""><input type=""radio"" name=""UserSelectFilesTF"" value=""1"" "
		If UserSelectFilesTF = 1 Then .Write (" checked")
		.Write ">"
		.Write "允许"
		.Write "  <input type=""radio"" name=""UserSelectFilesTF"" value=""0"""
		If UserSelectFilesTF = 0 Then .Write (" checked")
		.Write ">"
		.Write "  不允许</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许上传的最大文件大小：</strong></div></td>"
		.Write "      <td height=""30""><input name=""UpfilesSize""  onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""UpfilesSize"" value=""" & UpfilesSize & """ size=""10"">"
		.Write "      KB 　 <font color='#ff0000'>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font></td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td  width='160' class='clefttitle'><div align=""right""><strong>允许上传的文件类型：</strong><BR>"
		.Write "          <font color='#ff0000'>多种文件类型之间以""|""分隔</font></div></td>"
		.Write "      <td height=""30""><table width=""98%"" border=""0"">"
		.Write "        <tr>"
		.Write "          <td width=""19%"" height=""25"" align=""right"">图片类型:</td>"
		.Write "          <td width=""81%""><input name=""AllowUpPhotoType"" type=""text"" id=""AllowUpPhotoType"" value=""" & AllowUpPhotoType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Flash 文件:</td>"
		.Write "          <td><input name=""AllowUpFlashType"" type=""text"" id=""AllowUpFlashType"" value=""" & AllowUpFlashType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Windows 媒体: </td>"
		.Write "          <td><input name=""AllowUpMediaType"" type=""text"" id=""AllowUpMediaType"" value=""" & AllowUpMediaType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">Real 媒体: </td>"
		.Write "          <td><input name=""AllowUpRealType"" type=""text"" id=""AllowUpRealType"" value=""" & AllowUpRealType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "        <tr>"
		.Write "          <td height=""25"" align=""right"">其它文件:</td>"
		.Write "          <td><input name=""AllowUpOtherType"" type=""text"" id=""AllowUpOtherType"" value=""" & AllowUpOtherType & """ size=""30""></td>"
		.Write "        </tr>"
		.Write "      </table></td>"
		.Write "    </tr>"
        .Write "</table>"
		.Write "</div>"

		.Write " <div class=tab-page id=tougao-page>"
		.Write "  <H2 class=tab>投稿选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "	  tabPane1.addTabPage( document.getElementById( ""tougao-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>前台会员投稿总开关：</strong><br><font color=red>只有启用，会员才可在本模型里投稿。</font></div></td>"
		.Write "      <td height=""30"">"
		.Write "          <input type=""radio"" name=""UserTF"" value=""0"" "
		If UserTF = 0 Then .Write (" checked")
		.Write ">关闭会员投稿 <br>"
		.Write " <input type=""radio"" name=""UserTF"" value=""1"" "
		If UserTF = 1 Then .Write (" checked")
		.Write ">只允许注册会员可以投稿,具体可投稿的栏目依栏目设置而定<br/>"
		.Write " <input type=""radio"" name=""UserTF"" value=""2"" "
		If UserTF = 2 Then .Write (" checked")
		.Write ">允许所有用户投稿（包括游客）,具体可投稿的栏目依栏目设置而定<br>"


		.Write "        </td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg' style='color:blue'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>新注册会员：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""PubTimeLimit"" type=""text"" id=""PubTimeLimit"" value=""" & PubTimeLimit & """ size=""6"">分钟后才可以在此模型投稿</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>会员"& TempStr& "增加的金钱：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""UserAddMoney"" type=""text"" id=""UserAddMoney"" value=""" & UserAddMoney & """ size=""6"">元人民币（<font color=green>为0时不增加,可设置成负数,表示投稿要消费</font>）</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>会员"& TempStr& "增加的点券：</strong></td>"
		.Write "      <td height=""30""> <input name=""UserAddPoint"" type=""text"" id=""UserAddPoint"" value=""" & UserAddPoint & """ size=""6"">点点券（<font color=green>为0时不增加,可设置成负数,表示投稿要消费</font>）</td>"
		.Write "    </tr>"
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>会员"& TempStr& "增加的积分：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""UserAddScore"" type=""text"" id=""UserAddScore"" value=""" & UserAddScore & """ size=""6"">分积分（<font color=green>为0时不增加,可设置成负数,表示投稿要消费</font>）</td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>审核过的稿件是否允许修改：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""UserEditTF"" value=""0"" "
		If UserEditTF = 0 Then .Write (" checked")
		.Write ">不允许<font color=red>(建议)</font><br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""1"" "
		If UserEditTF = 1 Then .Write (" checked")
		.Write ">允许，但修改后自动转为未审(<font color=red>如果投稿要增加积分等,会导致重复收费</font>)<br>"
		.Write "          <input type=""radio"" name=""UserEditTF"" value=""2"" "
		If UserEditTF =2 Then .Write (" checked")
		.Write ">允许，修改后仍为已审状态（不推荐,<font color=red>如果投稿要增加积分等,会导致重复收费</font>）"
        .Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>投稿栏目显示方式：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""UserClassStyle"" value=""0"" "
		If UserClassStyle = 0 Then .Write (" checked")
		.Write ">仅显示有权限的栏目（下拉方式）<br>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""1"" "
		If UserClassStyle = 1 Then .Write (" checked")
		.Write ">仅显示有权限的栏目（跳窗方式）<br>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""2"" "
		If UserClassStyle = 2 Then .Write (" checked")
		.Write ">树型显示本模型下所有栏目,不允许投稿的栏目用灰色显示（跳窗方式）<br/>"
		.Write "          <input type=""radio"" name=""UserClassStyle"" value=""3"" "
		If UserClassStyle = 3 Then .Write (" checked")
		.Write "><font color=blue>多级联动下拉(<font color=red>新增</font>)（适合于只有二至三级栏目结构的模型）</font>"
        .Write "      </td>"
		.Write "    </tr>"
		
		
		.Write "    <tr class='tdbg'>"
		.Write "      <td width='160' class='clefttitle'><div align=""right""><strong>会员中心发布的信息是否需要审核：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""InfoVerificTF"" value=""0"" "
		If InfoVerificTF = 0 Then .Write (" checked")
		.Write ">需要后台人工审核<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""1"" "
		If InfoVerificTF = 1 Then .Write (" checked")
		.Write ">不需要审核（但不直接生成内容页HTML）<br>"
		.Write "          <input type=""radio"" name=""InfoVerificTF"" value=""2"" "
		If InfoVerificTF = 2 Then .Write (" checked")
		.Write ">不需要审核（当有启用生成静态HTML，直接生成内容页）<br>"

		.Write "      </td>"
		.Write "    </tr>"

        .Write "</table>"
		.Write "</div>"
		.Write " <div class=tab-page id=digg-page>"
		.Write "  <H2 class=tab>Digg选项</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""digg-page"" ) );"
		.Write "	</SCRIPT>"
		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>是否允许游客DIGG：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggByVisitor"" value=""1"" "
		If DiggByVisitor = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggByVisitor"" value=""0"" "
		If DiggByVisitor = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>是否启用单IP限制：</strong><br><font color=red>若启用单IP限制，每个IP的用户只能对每个项目Digg一次</font></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggByIP"" value=""1"" "
		If DiggByIP = 1 Then .Write (" checked")
		.Write ">启用"
		.Write "          <input type=""radio"" name=""DiggByIP"" value=""0"" "
		If DiggByIP = 0 Then .Write (" checked")
		.Write ">不启用"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>会员是否允许重复DIGG：</strong><br><font color=red>启用IP限制时，始终不允许</font></div></td>"
		.Write "      <td height=""30"">"
		.Write " <input type=""radio"" name=""DiggRepeat"" value=""1"" "
		If DiggRepeat = 1 Then .Write (" checked")
		.Write ">允许"
		.Write "          <input type=""radio"" name=""DiggRepeat"" value=""0"" "
		If DiggRepeat = 0 Then .Write (" checked")
		.Write ">不允许"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>次数选项：</strong></div></td>"
		.Write "      <td height=""30"">"
		.Write "         每DIGG一下自动增加<input type=""text"" class='textbox' size=""6"" style=""text-align:center""  name=""DiggPerTimes"" value=""" & DiggPerTimes & """>次 "
		.Write "      </td>"
		.Write "    </tr>"
		
        .Write "</table>"
        .Write "</div>"		 
		 
		.Write " <div class=tab-page id=detail-page>"
		.Write "  <H2 class=tab>其它参数</H2>"
		.Write "	<SCRIPT type=text/javascript>"
		.Write "				 tabPane1.addTabPage( document.getElementById( ""detail-page"" ) );"
		.Write "	</SCRIPT>"

		.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"

		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' width='160'><div align=""right""><strong>编辑器类型：</strong><br><font color=red>选择您习惯使用的编辑器。</font></div></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""EditorType"" value=""0"" "
		If EditorType = 0 Then .Write (" checked")
		.Write ">"
		.Write "        KesionCMS自带编辑器"
		.Write "          <input type=""radio"" name=""EditorType"" value=""1"" "
		If EditorType = 1 Then .Write (" checked")
		.Write ">"
		.Write "          FCKEditor        </td>"
		.Write "    </tr>"		
	
		.Write "    <tr class='tdbg'>"
		.Write "     <td class='clefttitle' width='160'><div align=""right""><strong>最新信息标志：</strong></div></td>"
		.Write "     <td height=""30""> <input type=""text"" size=8 name=""LatestNewDay"" value=""" & LatestNewDay & """>天内添加的信息标志为最新信息</td>"
		.Write "    </tr>"
		.Write "    <tr class='tdbg'>"
		.Write "     <td class='clefttitle' width='160'><div align=""right""><strong>后台每页显示：</strong></div></td>"
		.Write "     <td height=""30""> <input type=""text"" size=8 name=""MaxPerPage"" value=""" & MaxPerPage & """>条信息</td>"
		.Write "    </tr>"

		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle' align=""right"" width='160'><div><strong>评论设置：</strong></div></td>"
		.Write "      <td height=""30""><table width=""98%"" border=""0"">"
		.Write "     <tr valign=""middle"">"
		.Write "      <td width=""25%"" height=""30"" width='160'>"
		.Write "        <div align=""right"">本模型评论系统设置：</div></td>"
		.Write "      <td height=""30"">"
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""0"" "
		If VerificCommentTF = 0 Then .Write (" checked")
		.Write ">关闭本模型的所有信息评论<br>"

		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""1"" "
		If VerificCommentTF = 1 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容需要后台的审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""2"" "
		If VerificCommentTF = 2 Then .Write (" checked")
		.Write ">本模型只允许会员评论，且评论内容不需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""3"" "
		If VerificCommentTF = 3 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容需要后台审核<br>"
		
		.Write "<input type=""radio"" name=""VerificCommentTF"" value=""4"" "
		If VerificCommentTF = 4 Then .Write (" checked")
		.Write ">本模型允许会员，游客评论，且评论内容不需要后台审核"

		

		.Write "             </td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td width='160' height=""30"">"
		.Write "        <div align=""right"">评论需要验证码：</div></td>"
		.Write "      <td height=""30""> <input type=""radio"" name=""CommentVF"" value=""1"" "
		If CommentVF = 1 Then .Write (" checked")
		.Write ">"
		.Write "        是"
		.Write "          <input type=""radio"" name=""CommentVF"" value=""0"" "
		If CommentVF = 0 Then .Write (" checked")
		.Write ">"
		.Write "          否        </td>"
		.Write "    </tr>"
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30""> <div align=""right"">评论字数控制："
		.Write "        </div></td>"
		.Write "      <td width=""63%"" height=""30""> <input name=""CommentLen"" type=""text"" value=""" & CommentLen & """ size=""6"">不限制请输入""0""</td>"
		.Write "    </tr>"		
		.Write "    <tr valign=""middle"">"
		.Write "      <td height=""30""><div align=""right"">评论页模板：</div></td>"
		.Write "      <td height=""30""><input name=""CommentTemplate"" id=""CommentTemplate"" type=""text"" value=""" & CommentTemplate & """ size=""25"">&nbsp;" & KSCls.Get_KS_T_C("$('#CommentTemplate')[0]") & "</td>"
		.Write "    </tr>"			
		.Write "      </table></td>"
		.Write "    </tr>"	
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'><div align=""right""><strong>搜索页模板：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""SearchTemplate"" id=""SearchTemplate"" type=""text"" value=""" & SearchTemplate & """ size=""25"">&nbsp;" & KSCls.Get_KS_T_C("$('#SearchTemplate')[0]") & " </td>"
		.Write "    </tr>"
		If KS.WSetting(0)="1" Then
		.Write "    <tr class='tdbg'>"
		.Write "      <td class='clefttitle'><div align=""right""><strong>WAP搜索页模板：</strong></div></td>"
		.Write "      <td height=""30""> <input name=""WapSearchTemplate"" id=""WapSearchTemplate"" type=""text"" value=""" & WapSearchTemplate & """ size=""25"">&nbsp;" & KSCls.Get_KS_T_C("$('#WapSearchTemplate')[0]") & " </td>"
		.Write "    </tr>"
		End If
		.Write "  </table>"
		.Write "</div>"
		.Write "</form>"
		End With
		End Sub
		
		Sub ChannelSave()
		    Dim ModelEname,ThumbnailsConfig,ChannelTable,I,OpName,ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))

			ThumbnailsConfig=Request.Form("GoldenPoint") & "|" & KS.ChkClng(Request.Form("ThumbsWidth")) & "|" & KS.ChkClng(Request.Form("ThumbsHeight"))
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Channel Where ChannelID=" & ChannelID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				 OpName       = "添加"
				 ChannelTable = "KS_U_" & KS.G("ChannelTable")
				 ModelEname   = Replace(Replace(Replace(KS.G("ModelEname"), "\","/"), " ",""), "'","")
				 If not Conn.Execute("Select ModelEname From KS_Channel Where ModelEname='" & ModelEname & "'").eof Then
				   Call KS.AlertHistory("系统已存在该目录名称，请重输!",-1)
				   Exit Sub
				 End If
				 If not Conn.Execute("Select ChannelTable From KS_Channel Where ChannelTable='" & ChannelTable & "'").eof Then
				   Call KS.AlertHistory("系统已存在该数据表，请重输!",-1)
				   Exit Sub
				 End If
				 Dim sChannelID:sChannelID=Conn.Execute("select Max(ChannelID) From KS_Channel")(0)+1
				 If sChannelID<100 Then sChannelID=sChannelID+100
				RS("ChannelID")    = sChannelID
				RS("BasicType")    = KS.ChkClng(KS.G("BasicType"))
				RS("ChannelTable") = ChannelTable
				RS("ModelEname")   =ModelEname
			Else
			    OpName="修改"
			End If
				RS("ChannelName")= KS.G("ChannelName")
				RS("ItemName")   = KS.G("ItemName")
				RS("ItemUnit")   = KS.G("ItemUnit")
				RS("FsoFolder")  = KS.G("FsoFolder")
				RS("Descript")   = KS.G("Descript")
				Dim FieldBitStr,FieldValStr
				Select Case KS.ChkClng(RS("BasicType"))
				 Case 1
				    RS("CollectTF")=1
					
					FieldBitStr="1|1|"
					For I=2 To 30
					 FieldBitStr=FieldBitStr & KS.ChkClng(KS.G("A(" & I &")")) &"|"
					Next
					For I=0 To 30
					 FieldValStr=FieldValStr & Trim(KS.G("V" & I)) & "|"
					Next
				 Case 2
					FieldBitStr="1|1|1|1|1|"
					For I=5 To 30
					 FieldBitStr=FieldBitStr & KS.ChkClng(KS.G("P(" & I &")")) &"|"
					Next
					For I=0 To 30
					 FieldValStr=FieldValStr & Trim(KS.G("VP" & I)) & "|"
					Next
				 Case 3
					FieldBitStr="1|1|1|1|"
					For I=4 To 30
					 FieldBitStr=FieldBitStr & KS.ChkClng(KS.G("D(" & I &")")) &"|"
					Next
					For I=0 To 30
					 FieldValStr=FieldValStr & Trim(KS.G("VD" & I)) & "|"
					Next
				 Case Else
				   For I=1 To 30
				    FieldBitStr=FieldBitStr & "1|"
				   Next
				End Select	
				RS("FieldBit")=FieldBitStr & "@@@" & FieldValStr
				RS("ChannelStatus") = KS.G("ChannelStatus")
				RS("WapSwitch")     = KS.G("WapSwitch")
				RS("FsoHtmlTF")     = KS.ChkClng(KS.G("FsoHtmlTF"))
				RS("StaticTF")      = KS.ChkClng(KS.G("StaticTF"))
				RS("FsoListNum")    = KS.ChkClng(KS.G("FsoListNum"))
				RS("UpfilesDir")    = KS.G("UpfilesDir")
				RS("UserUpfilesDir") = KS.G("UserUpfilesDir")
				RS("UpFilesTF")     = KS.G("UpFilesTF")
				RS("UserSelectFilesTF")=KS.G("UserSelectFilesTF")
				'If KS.G("UpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir"))
				
				RS("UserUpFilesTF") = KS.G("UserUpFilesTF")
				'If KS.G("UserUpfilesDir") <> "" Then Call KS.CreateListFolder(KS.Setting(3) & KS.G("UserUpfilesDir"))
				
				RS("ThumbnailsConfig")=ThumbnailsConfig
	            RS("UserTF") = KS.ChkClng(KS.G("UserTF"))
				RS("UserEditTF")  = KS.ChkClng(KS.G("UserEditTF"))
				RS("UserClassStyle") = KS.ChkClng(KS.G("UserClassStyle"))
				RS("UpfilesSize") = KS.ChkClng(KS.G("UpfilesSize"))
				RS("AllowUpPhotoType") = KS.G("AllowUpPhotoType")
				RS("AllowUpFlashType") = KS.G("AllowUpFlashType")
				RS("AllowUpMediaType") = KS.G("AllowUpMediaType")
				RS("AllowUpRealType") = KS.G("AllowUpRealType")
				RS("AllowUpOtherType") = KS.G("AllowUpOtherType")
				RS("VerificCommentTF") = KS.G("VerificCommentTF")
				RS("LatestNewDay")     = KS.ChkClng(KS.G("LatestNewDay"))
				RS("CommentVF")    = KS.ChkClng(KS.G("CommentVF"))
				RS("CommentLen")   = KS.ChkClng(KS.G("CommentLen"))
				RS("CommentTemplate") = KS.G("CommentTemplate")
				RS("SearchTemplate")= KS.G("SearchTemplate")
				RS("WapSearchTemplate")= KS.G("WapSearchTemplate")
				RS("InfoVerificTF") = KS.ChkClng(KS.G("InfoVerificTF"))
				RS("MaxPerPage")   = KS.ChkClng(KS.G("MaxPerPage"))
				RS("RefreshFlag")  = KS.ChkClng(KS.G("RefreshFlag"))
				RS("FsoContentRule")=KS.G("FsoContentRule")
				RS("FsoClassListRule")=KS.ChkClng(KS.G("FsoClassListRule"))
				RS("FsoClassPreTag")=KS.G("FsoClassPreTag")

				'会员积分
				RS("UserAddMoney") = KS.ChkClng(KS.G("UserAddMoney"))
				RS("UserAddPoint") = KS.ChkCLng(KS.G("UserAddPoint"))
				RS("UserAddScore") = KS.ChkClng(KS.G("UserAddScore"))
				RS("PubTimeLimit") = KS.ChkClng(KS.G("PubTimeLimit"))
				RS("EditorType") = KS.ChkClng(KS.G("EditorType"))
				RS("DiggByVisitor")= KS.ChkClng(KS.G("DiggByVisitor"))
				RS("DiggByIP")     = KS.ChkClng(KS.G("DiggByIP"))
				RS("DiggRepeat")= KS.ChkClng(KS.G("DiggRepeat"))
				RS("DiggPerTimes")= KS.ChkClng(KS.G("DiggPerTimes"))
				RS.Update
				RS.Close
				
				If OpName="添加" Then
				'建立新表
				dim sql
			    Select Case KS.ChkClng(KS.G("BasicType"))
			    Case 1
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"TID nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"TitleType nvarchar(30),"&_
						"Title nvarchar(255),"&_
						"FullTitle nvarchar(255),"&_
						"Intro ntext,"&_
						"ShowComment tinyint Default 0,"&_
						"TitleFontColor nvarchar(30),"&_
						"TitleFontType nvarchar(30),"&_
						"ArticleContent ntext,"&_
						"PageTitle ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"LastHitsTime datetime,"&_
						"AddDate datetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"PhotoUrl nvarchar(150),"&_
						"PicNews tinyint default 0,"&_
						"Changes tinyint default 0,"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"OrderID tinyint Default 1,"&_
						"IsSign tinyint Default 0,"&_
						"SignUser nvarchar(255),"&_
						"SignDateLimit tinyint Default 0,"&_
						"SignDateEnd datetime,"&_
						"Province nvarchar(100),"&_
						"City nvarchar(100),"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
			 Case 2
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"PhotoUrl nvarchar(255),"&_
						"PicUrls ntext,"&_
						"PictureContent ntext,"&_
						"Author nvarchar(30),"&_
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate smalldatetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"Score int Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"OrderID tinyint Default 1,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0"&_
						")"
				Conn.Execute(sql)
				KS.ConnItem.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
				
			 Case 3
				sql="CREATE TABLE ["&ChannelTable&"] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&ChannelTable&" PRIMARY KEY,"&_
						"Tid nvarchar(22),"&_
						"KeyWords nvarchar(255),"&_
						"Title nvarchar(255),"&_
						"DownVersion nvarchar(50),"&_
						"DownLB nvarchar(100),"&_
						"DownYY nvarchar(100),"&_
						"DownSQ nvarchar(100),"&_
						"DownPT nvarchar(100),"&_
						"DownSize nvarchar(100),"&_
						"YSDZ nvarchar(100),"&_
						"ZCDZ nvarchar(100),"&_
						"JYMM nvarchar(100),"&_
						"PhotoUrl nvarchar(200),"&_
						"BigPhoto nvarchar(200),"&_
						"DownUrls ntext,"&_
						"DownContent ntext,"&_
						"Author nvarchar(50)," & _
						"Origin nvarchar(40),"&_
						"Rank nvarchar(10),"&_
						"LastHitsTime smalldatetime," &_
						"Hits int Default 0,"&_
						"HitsByDay int Default 0,"&_
						"HitsByWeek int Default 0,"&_
						"HitsByMonth int Default 0,"&_
						"AddDate smalldatetime,"&_
						"JSID nvarchar(200),"&_
						"TemplateID nvarchar(255),"&_
						"WapTemplateID nvarchar(255)," &_
						"Fname nvarchar(200),"&_
						"RefreshTF tinyint default 0,"&_
						"Inputer nvarchar(50),"&_
						"Recommend tinyint Default 0,"&_
						"Rolls tinyint Default 0,"&_
						"Strip tinyint Default 0,"&_
						"Popular tinyint Default 0,"&_
						"Verific tinyint Default 0,"&_
						"Slide tinyint Default 0,"&_
						"Comment tinyint Default 0,"&_
						"IsTop tinyint Default 0,"&_
						"DelTF tinyint Default 0,"&_
						"OrderID tinyint Default 1,"&_
						"InfoPurview tinyint Default 0,"&_
						"ArrGroupID nvarchar(100),"&_
						"ReadPoint int Default 0,"&_
						"ChargeType tinyint Default 0,"&_
						"PitchTime int Default 24,"&_
						"ReadTimes int Default 10,"&_
						"DividePercent int Default 0"&_
						")"
				Conn.Execute(sql)
				'添加索引
				Call AddIndex(ChannelTable, "[TID]", "[TID]")
				Call AddIndex(ChannelTable, "[Verific]", "[verific]")
				Call AddIndex(ChannelTable, "[deltf]", "[deltf]")
				Call AddIndex(ChannelTable, "[adddate]", "[adddate]")
				Call AddIndex(ChannelTable, "[hits]", "[hits]")
				'Call AddIndex(ChannelTable, "[specialid]", "[specialid]")
			 End Select
				
				
				
				If KS.ChkClng(KS.G("BasicType"))=3 Then
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownPhoto/")
				 Call KS.CreateListFolder(KS.Setting(3) & KS.G("UpfilesDir")&"DownUrl/")
				End IF
				  
                 
				'  If Err<>0 Then
				'	Conn.RollBackTrans
				'	Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				'  Else
				'	Conn.CommitTrans
				  'End If
				  Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				  Call KS.DelCahe(KS.SiteSN & "_selectclass")
				  Call KS.DelCahe(KS.SiteSN & "_classpath")
				  Call KS.DelCahe(KS.SiteSN & "_classnamepath")				     
				End If
				
				Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
				
				Response.Write ("<script>alert('KesionCMS系统提醒您：\n\n1、模型配置信息" & OpName & "成功；\n\n2、为了使配置生效，请及时更新缓存；');parent.frames['LeftFrame'].location.reload();location.href='KS.Model.asp';</script>")
			
		End Sub
	
		Sub DelColumn(TableName,ColumnName)
		On Error Resume Next
		Conn.Execute("Alter Table "&TableName&" Drop "&ColumnName&"")
		End Sub
		
		Sub DelTable(TableName,C)
			On Error Resume Next
			C.Execute("Drop Table "&TableName&"")
		End Sub
		
		Sub AddIndex(ByVal TableName, ByVal IndexName, ByVal ValueText)
			On Error Resume Next
			Conn.Execute("CREATE INDEX " & IndexName & " ON " & TableName & "(" & ValueText & ")")
		End Sub
		
		
		Sub ChannelDel()
		  Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
		  Call DelTable(KS.C_S(ChannelID,2),Conn)
		  
		  '删除采集数据库里的相关字段和表
		  If KS.C_S(ChannelID,6)="1" Or KS.C_S(ChannelID,6)="2" or KS.C_S(ChannelID,6)="5" Then  Call DelTable(KS.C_S(ChannelID,2),KS.ConnItem)
		  KS.ConnItem.Execute("Delete From KS_FieldItem Where ChannelID=" & ChannelID)
		  KS.ConnItem.Execute("Delete From KS_FieldRules Where ChannelID=" & ChannelID)
		  '=================================
		  
		  Conn.Execute("Delete From KS_Comment Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_DownParam Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_DownSer Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Origin Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Channel Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Class Where ChannelID=" & ChannelID)
		  Conn.Execute("Delete From KS_Field Where ChannelID=" & ChannelID)
		  
		  		 Call KS.DelCahe(KS.SiteSN & "_selectallowclass")
				 Call KS.DelCahe(KS.SiteSN & "_selectclass")
				 Call KS.DelCahe(KS.SiteSN & "_classpath")
				 Call KS.DelCahe(KS.SiteSN & "_classnamepath")
				 Call KS.DelCahe(KS.SiteSN & "_ChannelConfig")
		  Response.Write "<script>alert('模型删除成功!');parent.frames['LeftFrame'].location.reload();location.href='KS.Model.asp';</script>" 
		End Sub
		
End Class
%> 


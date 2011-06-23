<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Picture
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Picture
        Private KS,KSCls
		'=====================================定义本页面全局变量=====================================
		Private ID, I, totalPut, Page, RS,ComeFrom
		Private KeyWord, SearchType, StartDate, EndDate, ParentRs, SearchParam,MaxPerPage,SpecialID
		Private T, TitleStr, AttributeStr
		Private FolderID, TemplateID,WapTemplateID,TN, TI,TJ,Action,UserDefineFieldArr,UserDefineFieldValueStr
		Private PicID, Title, PhotoUrl, PictureContent, PicUrls, Recommend,IsTop
		Private Popular, Strip, Verific, Comment, Slide, ChangesUrl, Rolls, KeyWords, Author, Origin, AddDate, Rank, Hits, HitsByDay, HitsByWeek, HitsByMonth
		Private CurrPath, InstallDir,PreViewObj, UpPowerFlag,Inputer,SaveFilePath
		Private ComeUrl,F_B_Arr,F_V_Arr,ChannelID,FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,Score
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		'=============================================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
        F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		'收集搜索参数
		KeyWord   = KS.G("KeyWord")
		SearchType= KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate   = KS.G("EndDate")
		Action     = KS.G("Action")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
		
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))

			Action = Trim(KS.G("Action"))
			Page = KS.G("page")
							
			IF KS.G("Method")="Save" Then
				 Call PictureSave()
			Else 
				 Call PictureAdd()
			End If
		End Sub

        '添加
        Sub PictureAdd() 
			With Response
			CurrPath = KS.GetUpFilesDir()
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then
			  FolderID = Trim(KS.G("FolderID"))
			  
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加图片的权限
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&ChannelID=" & ChannelID &"';</script>")
			   Call KS.ReturnErr(1, "")
			   Exit Sub
			  End If
			  Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:Comment = 1:IsTop=0:UserDefineFieldValueStr=0
			  ReadPoint=0:PitchTime=24:ReadTimes=10:Score=0
			  PreViewObj = "<br><br><br>" & KS.C_S(ChannelID,3) & "预览区"
			  KeyWords = Session("keywords")
			  Author = Session("Author")
			  Origin = Session("Origin")
			
			ElseIf Action = "Edit" Or Action="Verify" Then

			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From " & KS.C_S(ChannelID,2) & " Where ID=" & KS.ChkClng(KS.G("ID")), conn, 1, 1
			   If RS.EOF And RS.BOF Then
				Call KS.Alert("参数传递出错!", ComeUrl)
				Set KS = Nothing:.End:Exit Sub
			   End If
				PicID = Trim(RS("ID"))
				FolderID = Trim(RS("Tid"))
				If Action ="Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑图片的权限
				 RS.Close:Set RS = Nothing
				 If KeyWord = "" Then
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 Else
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=" &server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>搜索" & KS.C_S(ChannelID,3) & "结果</font>") & "&ButtonSymbol=PictureSearch';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
				 End If
				 Exit Sub
			   End If
			   IF Action="Verify" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then 
			     RS.Close:Set RS = Nothing
				  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
				  Call KS.ReturnErr(1, "KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
				 
				 Exit Sub   
			   End If
			   
				Title    = Trim(RS("title"))
				PhotoUrl = Trim(RS("PhotoUrl"))
				PreViewObj = "<img src='" & PhotoUrl & "' border='0'>"
				PicUrls  = Trim(RS("PicUrls"))
				PictureContent = Trim(RS("PictureContent")) : If KS.IsNul(PictureContent) Then PictureContent=" "
				Rolls    = CInt(RS("Rolls"))
				Strip    = CInt(RS("Strip"))
				Recommend = CInt(RS("Recommend"))
				Popular  = CInt(RS("Popular"))
				Verific  = CInt(RS("Verific"))
				Comment  = CInt(RS("Comment"))
				IsTop    = (RS("IsTop"))
				Slide    = CInt(RS("Slide"))
				AddDate  = CDate(RS("AddDate"))
				Rank     = Trim(RS("Rank"))
				FileName       = RS("Fname")
				
				TemplateID = RS("TemplateID")
				WapTemplateID=RS("WapTemplateID")
				Hits = Trim(RS("Hits"))
				HitsByDay = Trim(RS("HitsByDay"))
				HitsByWeek = Trim(RS("HitsByWeek"))
				HitsByMonth = Trim(RS("HitsByMonth"))
				Score=RS("Score")
				KeyWords = Trim(RS("KeyWords"))
				Author = Trim(RS("Author"))
				Origin = Trim(RS("Origin"))
				FolderID = RS("Tid")
				
				ReadPoint = RS("ReadPoint")
				ChargeType= RS("ChargeType")
				PitchTime = RS("PitchTime")
				ReadTimes = RS("ReadTimes")
				InfoPurview=RS("InfoPurview")
				arrGroupID = RS("arrGroupID")
				DividePercent=RS("DividePercent")
               '自定义字段
				UserDefineFieldArr=KSCls.Get_KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				  Dim UnitOption
				  If UserDefineFieldArr(12,I)="1" Then
				   UnitOption="@" & RS(UserDefineFieldArr(0,I)&"_Unit")
				  Else
				   UnitOption=""
				  End If
				  If I=0 Then
				    UserDefineFieldValueStr=RS(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & RS(UserDefineFieldArr(0,I)) &UnitOption& "||||"
				  End If
				Next
			  End If
				RS.Close
			End If
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrlf
			.Write "<title>添加</title>" & vbCrlf
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>" & vbCrlf
			.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>" & vbCrlf
			.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>" & vbCrlf
			.Write "<script language=""javascript"" src=""../KS_Inc/popcalendar.js""></script>" & vbCrlf
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
			.Write "</head>" & vbCrlf
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>" & vbCrlf
			.Write "<div align='center'>" & vbCrlf
			.Write "<ul id='menu_top'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
			.Write "<li onclick=""history.back();"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>"
		    .Write "</ul>" & vbCrlf
			
			.Write "<div class=tab-page id=PhotoPane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""PhotoPane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			.Write " <div class=tab-page id=basic-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"
			.Write "	</SCRIPT>"		
			.Write "    <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='PictureForm' name='PictureForm' onsubmit='return(SubmitFun())'>"
            .Write " <TABLE width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "      <input type='hidden' value='" & PicID & "' name='PicID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'>"
			
			.Write "       <tr class='tdbg'>"
			.Write "          <td height='20' width='85' class='clefttitle'><div align='right'><font color='#FF0000'><strong>" & F_V_Arr(0) & ":</strong></font></div></td>"
			.Write "          <td height='25' nowrap> "
			.Write "            <input name='title' type='text'  class='textbox' value='" & Title & "' size=80>"
			.Write "                  <font color='#FF0000'>*</font>"
			If F_B_Arr(17)=1 Then
			.Write "<input type='checkbox' name='MakeHtml' value='1' checked>" & F_V_Arr(17)
			End IF
			.Write "                  </td>"
			.Write "       </tr>"
			.Write "       <tr class='tdbg'>"
			.Write "         <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(1) & ":</strong></div></td>"
			.Write "         <td><input type='hidden' name='OldClassID' value='"& FolderID & "'>"
			.Write " <select size='1' name='tid' id='tid'>"
			.Write " <option value='0'>--请选择栏目--</option>"
			.Write Replace(KS.LoadClassOption(ChannelID),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"

		
		 If F_B_Arr(5)=1 Then
			.Write "&nbsp;&nbsp;" & F_V_Arr(5) & " <input name='Recommend' type='checkbox' id='Recommend' value='1'"
			If Recommend = 1 Then .Write (" Checked")
			.Write ">推荐"
			.Write "<input name='Rolls' type='checkbox' id='Rolls' value='1'"
			If Rolls = 1 Then .Write (" Checked")
			.Write ">滚动"
			.Write "<input name='Strip' type='checkbox' id='Strip' value='1'"
			If Strip = 1 Then .Write (" Checked")
			.Write ">头条"
			.Write "<input name='Popular' type='checkbox' id='Popular' value='1'"
			If Popular = 1 Then .Write (" Checked")
			.Write ">热门"
			.Write "<input name='IsTop' type='checkbox' id='IsTop' value='1'"
			If IsTop = 1 Then .Write (" Checked")
			.Write ">固顶"
			.Write "<input name='Comment' type='checkbox' id='Comment' value='1'"
			If Comment = 1 Then .Write (" Checked")
			.Write ">允许评论"
			.Write "<input name='Slide' type='checkbox' id='Slide' value='1'"
			If Slide = 1 Then
			.Write (" Checked")
			End If
			.Write ">幻灯"
			.Write "</td>"
			.Write "              </tr>"
		 End If
		 If F_B_Arr(6)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(6) & ":</strong></div></td>"
			.Write "                <td> <input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=40> <="
			.Write "                  <select name='SelKeyWords' style='width:150px' onChange='InsertKeyWords($(""#KeyWords"")[0],this.options[this.selectedIndex].value)'>"
		    .Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
			.Write "                  </select>"
			.Write " <br />【<a href=""#"" id=""KeyLinkByTitle"" style=""color:green"">根据" & F_V_Arr(0) & "自动获取Tags</a>】<input type='checkbox' name='tagstf' value='1' checked>写入Tags表</td>"
			.Write "              </tr>"
		End If
		If F_B_Arr(7)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(7) & ":</strong></div></td>"
			.Write "                <td> <input name='author' type='text' id='author' value='" & Author & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick='$(""#author"").val(""未知"");' style='cursor:pointer;'>未知</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#author').val('佚名');"" style='cursor:pointer;'>佚名</font></font>】【<font color='blue'><font color='red' onclick=""$('#author').val('" & KS.C("AdminName") & "');"" style='cursor:pointer;'>" & KS.C("AdminName") & "</font></font>】"
							 If Author <> "" And Author <> "未知" And Author <> KS.C("AdminName") And Author <> "佚名" Then
							  .Write ("【<font color='blue'><font color='#993300' onclick=""$(""#author"").val('" & Author & "');"" style='cursor:pointer;'>" & Author & "</font></font>】")
							 End If
							  .Write ("<select name='SelAuthor' style='width:100px' onChange=""$('#author').val(this.options[this.selectedIndex].value);"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=0 And OriginType=1 Order BY AddDate Desc")
			.Write "                                   </select></td>"
			.Write "              </tr>"
	  End If
	  If F_B_Arr(8)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  width='85' class='clefttitle'><div align='right'><strong>" & F_V_Arr(8) & ":</strong></div></td>"
			.Write "                <td> <input name='Origin' type='text' id='Origin' value='" & Origin & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('不详');"" style='cursor:pointer;'>不详</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('本站原创');"" style='cursor:pointer;'>本站原创</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('互联网');"" style='cursor:pointer;'>互联网</font></font>】"
							  If Origin <> "" And Origin <> "不详" And Origin <> "本站原创" And Origin <> "互联网" Then
							  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('" & Origin & "')"" style='cursor:pointer;'>" & Origin & "</font></font>】 ")
							   End If
							  .Write ("<select name='selOrigin' style='width:100px' onChange=""$('#Origin').val(this.options[this.selectedIndex].value)"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")
			.Write "                </select> </td>"
			.Write "              </tr>"
	 End If
			        '自定义字段
		    .Write KSCls.Get_KS_D_F(ChannelID,UserDefineFieldValueStr)
			
			
			 Dim picnum,EndNum
		     If Action="Edit" Or Action="Verify" Then
					picnum = UBound(split(PicUrls,"|||"))+1
					EndNum=PicNum
			 Else
			        picnum=4:EndNum=99
			 End If

            If KS.G("Action")="Verify" Then
			.Write "              <tr  class='tdbg' style='display:none'>"
			Else
			.Write "              <tr class='tdbg'>"
			End If
			.Write "                <td height='25'  nowrap align='right' class='clefttitle'><strong>添加模式:</strong></div></td>"
			.Write "                <td>"
			If PicNum<100 Then
			.Write "<input type='radio' name='addmode' value='0' checked onclick='$(""#addmore"").hide();$(""#addarea"").show();'>直接添加 <input type='radio' name='addmode' value='1' onclick='$(""#addmore"").show();$(""#addarea"").hide()'>批量添加"
			Else
			.Write "<input type='radio' name='addmode' value='0' onclick='$(""#addmore"").hide();$(""#addarea"").show()'>直接添加 <input type='radio' name='addmode' value='1' onclick='$(""#addmore"").show();$(""#addarea"").hide()'  checked>批量添加"
			End If
			.Write "                </td>"
			.Write "              </tr>"
			Dim CurrDate:CurrDate=Year(Now) &right("0"&Month(Now),2)
			Dim CurrDay:CurrDay=CurrDate & right("0"&day(Now),2)
			If PicNum<100 Then
			.Write "            <tr class='tdbg' id='addmore' style='display:none'>"
			Else
			.Write "            <tr class='tdbg' id='addmore' style=''>"
			End If
			.Write "               <td height='35' align='right' class='clefttitle'><strong>" & KS.C_S(ChannelID,3) & "地址:</strong></td>"
			.Write "               <td height='25'><input  name='MorePicUrl' type='text' id='MorePicUrl' size='80' value='图片#|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#.jpg|" &  CurrPath & "/"&CurrDate &"/" & CurrDay & "#_S.jpg' class='upfile'><br>&nbsp;&nbsp;开始ID：<input class='textbox' type='text' value='1' name='morestart' size=5> 结束ID：<input class='textbox' type='text' value='" & EndNum & "' name='moreend' size=5><font color=red> 数字序号通配符为#，注意通配符只用一个#即可</font><br>&nbsp;&nbsp;<font color=green>格式：图片介绍|大图地址|小图地址</font></td>"
			.Write "              </tr>"


			If PicNum<100 Then
			.Write "<tr class='tdbg' id='addarea'>"			
			Else
			.Write "<tr class='tdbg' id='addarea' style=""display:none"">"			
			End IF
		    .Write " <td width='85' class='CLeftTitle'><div align='right'><b>" & F_V_Arr(4) & ":</b><br><input type='checkbox' value='1' name='BeyondSavePic' checked>自动存图<br><br><font color=#ff0000>如果想删除某张" & KS.C_S(ChannelID,3) & "，请在" & KS.C_S(ChannelID,3) & "地址里输入 ""del""或留空。</font></div></td>"
            .Write "<td height='24'>"
			.Write "<strong>" & F_V_Arr(3) & ":</strong>"
            .Write "<input name='picnum' type='text' id='picnum' size='4' value='" & PicNum & "' style='text-align:center'>"  
			.Write "&nbsp;<input name='kkkup' type='button' id='kkkup2' value='设定' onClick=""MakeUpload($('#picnum').val(),'click');"" class='button'>注：最多<font color='red'>99</font>" & KS.C_S(ChannelID,4) & "，远程" & KS.C_S(ChannelID,3) & "地址必须以<font color='red'>http://</font>开头"
	        .Write "<input type='hidden' name='PicUrls' id='PicUrls'>"
			
			.Write "	<span id='uploadfield'></span>"
				If Action<>"Edit" then
			.Write "	<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?ChannelID="& ChannelID & "' frameborder=0 scrolling=no width='100%' height='110'></iframe>"
				End If
		    .Write "</td>"
            .Write "</tr>"
			
			

			.Write "              <tr  class='tdbg' id='mode1' style='height:25px'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(2) & ":</strong></div></td>"
			.Write "                <td> <input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='" & PhotoUrl & "' class='textbox'>"
			.Write "   <font color='#FF0000'>*</font>&nbsp;<input class='button' type='button' name='Submit' value='选择图片地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID & "&CurrPath=" & CurrPath & "',550,290,window,$('#PhotoUrl')[0]);""> <input class='button' type='button' name='Submit' value='远程抓图...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&ItemName=图片&CurrPath=" & CurrPath & "',300,100,window,$('#PhotoUrl')[0]);"">"
			.Write "            </td>"
			.Write "              </tr>"

		 If F_B_Arr(9)=1 Then
			.Write "      <TR class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(9) & ":</strong></div></td>"
			.Write "                <td nowrap>"
			.Write "      <textarea  ID='Content' name='Content' cols=90 rows=6 style='display:none'>" & Server.HTMLEncode(PictureContent) & "</textarea>"
			If KS.C_S(ChannelID,34)=0 Then
				.Write "<iframe id='PictureContent' name='PictureContent' src='KS.Editor.asp?ID=Content&style=0&ChannelID=" & channelid & "' frameborder=0 scrolling=no width='100%' height='200'></iframe>"
			Else
				.Write "<input type=""hidden"" id=""content___Config"" value="""" style=""display:none"" /><iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""100%"" height=""160"" frameborder=""0"" scrolling=""no""></iframe>"

            End If
			
			.Write "      </TD></TR>"
		 End If
           .Write "</table>"
		   .Write "</div>"
	
	If F_B_Arr(15)=1 Then		 
		   .Write " <div class=tab-page id=classoption-page>"
		   .Write "  <H2 class=tab>属性设置</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
		   .Write "	</SCRIPT>"

            .Write "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "           <tr class='tdbg'>"
			.Write "              <td class='clefttitle' width='80' align='right'><strong>所属专题:</strong></td>"
			.Write "              <td>"
			Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
			.write "              </td>"
			.Write "           </tr>"
		 If F_B_Arr(10)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(10) & ":</strong></div></td>"
			.Write "                <td>"
			If Action <> "Edit" Then
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & Now() & "' size='50'  class='textbox'>")
			Else
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & AddDate & "' size='50'  class='textbox'>")
			End If
			.Write "                  <b><a href='#' onClick=""popUpCalendar(this, $('input[name=AddDate]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif' border='0' align='absmiddle' title='选择日期'></a><strong>日期格式：年-月-日 时：分：秒</strong>"
			.Write "                  </b></td>"
			.Write "              </tr>"
		End If
		If F_B_Arr(11)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(11) & ":</strong></div></td>"
			.Write "                <td><select name='rank'>"
			If Rank = "★" Then
			.Write "                    <option  selected>★</option>"
			Else
			.Write "                    <option>★</option>"
			End If
			If Rank = "★★" Then
			.Write "                    <option  selected>★★</option>"
			Else
			.Write "                    <option>★★</option>"
			End If
			If Rank = "★★★" Or Action = "Add" Then
			.Write "                    <option  selected>★★★</option>"
			Else
			.Write "                    <option>★★★</option>"
			End If
			If Rank = "★★★★" Then
			.Write "                    <option  selected>★★★★</option>"
			Else
			.Write "                    <option>★★★★</option>"
			End If
			If Rank = "★★★★★" Then
			.Write "                    <option  selected>★★★★★</option>"
			Else
			.Write "                    <option>★★★★★</option>"
			End If
			.Write "                  </select>"
			.Write "                  请为" & KS.C_S(ChannelID,3) & "评定推荐等级</td>"
			.Write "              </tr>"
	   End If
	   If F_B_Arr(12)=1 Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td  class='clefttitle'><div align='right'><strong>" & F_V_Arr(12) & ":</strong></td>"
			 .Write "               <td> 本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='10' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='10' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='10' class='textbox'> 总计：<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='10' class='textbox'>"
			 .Write "&nbsp;得票数：<input type='text' name='score' size='6' value='" & score & "'>票  作弊用的"
			 .Write "             </td>"
			 .Write "             </tr>"
	  End If
	  If F_B_Arr(13)=1 Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(13) & ":</strong></div></td>"
			.Write "                <td> "
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>继承栏目设定<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>自定义"
			.Write "<div id='templatearea' style='display:none'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAP模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			Else
			
			.Write "<div id='templatearea'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly maxlength='255' size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]")
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAP模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			End If
			.Write "                </td>"
			.Write "             </tr>"
	  End If
	  If F_B_Arr(14)=1 Then
			.Write "             <tr class='tdbg'>"
			.Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(14) & ":</strong></td><td>"
			IF Action = "Edit" or Action="Verify" Then
			.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='25' class='textbox'> <font color=red>不能改</font>"
			Else
			.Write "<input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>自动生成 <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >自定义"
			.Write "<div id='filearea' style='display:none'><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='25' class='textbox'> <font color=red>可带路径,如 help.html,news/news_1.shtml等</font></div>"
			End IF
			 .Write "                  </td>"
			 .Write "             </tr>"
	 End If
			
			.Write "</table>"
			.Write "</div>"
  End If
      
	     If F_B_Arr(16)=1 Then
	       KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
         End If
		 
	       KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))
		   
			 .Write "</form>"
			 .Write " </div>"
			%>
			 <script>
			 <%if Action<>"Edit" Then%>
			 var LastNum=1;
			 <%else%>
			 var LastNum=$('#picnum').val();
			 <%end if%>
			 var tempup='';
			 var picnum=<%=Picnum%>;
			 
			 $(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
			 <%If F_B_Arr(6)=1 Then%>
			  $('#KeyLinkByTitle').click(function(){
			    GetKeyTags();
			  });
			 <%End If%>
			 
			 	 if (<%=Picnum%><100) MakeUpload(<%=Picnum%>);
				 IniPicUrl();
				 tempup=$("#uploadfield").html();
                 $('#picnum').val(picnum);
			  
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=Title]').val());
			  
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...');
				  $("#KeyWords").attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data));
					$('#KeyWords').attr("disabled",false);
				  });
			  }else{
			   alert('对不起,请先输入内容!');
			  }
			}
			
			function IniPicUrl()
			{
			 var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			   for ( var i=1 ;i<PicUrlArr.length+1;i++)
			   { 
				 $('input[name=imgurl'+i+']').val(PicUrlArr[i-1].split('|')[1]);
				 $('input[name=thumb'+i+']').val(PicUrlArr[i-1].split('|')[2]);
				 $('#imgnote'+i).val(PicUrlArr[i-1].split('|')[0]);
				 $('#picview'+i).html('');
				 if (document.all){
				 $('#picview'+i)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=PicUrlArr[i-1].split('|')[1];
				 }else{
				  $('#picview'+i).html('<img width="120" height="80" src="'+PicUrlArr[i-1].split('|')[1]+'">');
				 }
			   }
			    
			    LastNum=i;
			   }
			}
			function MakeUpload(mnum,str)
			{ 
			   if (parseInt(mnum)>=100){
			   alert('最多只能同上传99张!');
			  return false;}
			   var startNum=1;
			   var endNum = mnum;
			   var fhtml = "";
			   if (str=='click') startNum=LastNum;
			   for(startNum;startNum <= endNum;startNum++){
				   fhtml += "<table width=\"99%\" style='margin:2px' class='ctable' align=center border=\"0\" id=\"seltb"+startNum+"\" cellpadding=\"3\" cellspacing=\"1\">";
				   fhtml += "<tr class='tdbg'> "
				   fhtml +="  <td height=\"25\" width=20 align=center class=clefttitle rowspan=\"3\"><strong>第"+startNum+"张</strong></td>";
				   fhtml += " <td width=\"124\" rowspan=\"3\" align=\"center\"><div id=\"picview"+startNum+"\" name=\"picview"+startNum+"\" style=\"filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:80px;width:120px;border:1px solid #777777\"><img src=\"images/pview.gif\" width=\"120\" height=\"80\"></div></td>";
				   fhtml += "</tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"25\"> 　小图地址：";
				   fhtml += "<input type=\"text\" onblur='view("+startNum+");' name='thumb"+startNum+"' id='thumb"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取图片\" onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,document.PictureForm.thumb"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;<input type=\"button\" name='tgetpic"+startNum+"' value=\"远程抓图\" onclick=\"OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle=抓取远程图片&ItemName=<%=KS.C_S(ChannelID,3)%>&CurrPath=<%=CurrPath%>',300,100,window,document.PictureForm.thumb"+startNum+");view("+startNum+");\" class=\"button\">";
				   fhtml += "<br>&nbsp;&nbsp;&nbsp;大图地址：<input type=\"text\" onblur='view("+startNum+");' name='imgurl"+startNum+"' id='imgurl"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"hidden\" name='pimgurl"+startNum+"' value=\"\">";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取图片\" onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,document.PictureForm.imgurl"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;<input type=\"button\" name='getpic"+startNum+"' value=\"远程抓图\" onclick=\"OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle=抓取远程图片&ItemName=<%=KS.C_S(ChannelID,3)%>&CurrPath=<%=CurrPath%>',300,100,window,document.PictureForm.imgurl"+startNum+");view("+startNum+");\" class=\"button\">";
				    <%If Action="Edit" Then%>
					if (startNum>picnum)
					{
				   fhtml += "<iframe id='UpPhotoFrame"+startNum+"' name='UpPhotoFrame"+startNum+"' src='KS.UpFileForm.asp?ChannelID=<%=ChannelID%>&uptype=Single&objid="+startNum+"' frameborder=0 scrolling=no width='100%' height='22'></iframe>"
				   }
				   <%end if%>
				   
				   fhtml += "</td></tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"30\">　简要介绍： ";
				   fhtml += "<textarea name='imgnote"+startNum+"' id='imgnote"+startNum+"' style=\"height:46px;width:399px\"></textarea> </td>";
				   fhtml += "</tr></table>\r\n";
			  }
			  <%If Action="Edit" Then%>
			  //LastNum=Number(endNum)+1;
			  $("#uploadfield").html(tempup+fhtml);
			  IniPicUrl();
			  <%Else%>
			  $("#uploadfield").html(fhtml);
			  frames['UpPhotoFrame'].ChooseOption(mnum);
			  $('#UpPhotoFrame').height(80+26*(mnum/2)); 
			  <%End If%>
			}
			 function view(num)
			 {
			  if ($("input[name=thumb"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			     if (document.all){
			     $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=thumb"+num+"]").val();}else{ $("#picview"+num).html("<img width='120' height='80' src='"+$("input[name=thumb"+num+"]").val()+"'>");
			    }
			  }
			  else if($("input[name=imgurl"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			   if (document.all){
			       $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=imgurl"+num+"]").val();
			    }else{
				 $("#picview"+num).html("<img width='120' height='80' src='"+$("input[name=imgurl"+num+"]").val()+"'>");
				}
			  }
			 }

			 function SetPicUrlByUpLoad(DefaultThumb,PicUrlStr,ThumbPathFileName)
			{  var UrlStrArr;
			   UrlStrArr=PicUrlStr.split('|');
			   for (var i=1;i<UrlStrArr.length;i++)
			   {
			   var url=UrlStrArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=imgurl'+i+']').val(url);
			   } 
			  }
			  var ThumbsArr=ThumbPathFileName.split("|")
			  for(var i=1;i<ThumbsArr.length;i++)
			  {
			   var url=ThumbsArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=thumb'+i+']').val(url);
			   } 
			  }

			 $('input[name=PhotoUrl]').val(ThumbsArr[DefaultThumb-1]);

			}
			function SelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",true);
			  });
			}
			function UnSelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",false);
			  });
			}

			function GetFileNameArea(f)
			{
			  $('#filearea').toggle(f);
			}
			function GetTemplateArea(f)
			{
			   $('#templatearea').toggle(f);
			}
            var spic=null;
			function SubmitFun()
				{ 
			    if ($('input[name=title]').val()=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					$('input[name=title]').focus();
					return;
				  }
			   if ($("#tid>option[selected=true]").val()=='0')
			   {
			       alert('请选择所属栏目!');
				   return false;
			   }
			 	if ($('input[name=PhotoUrl]').val()=='')
				{
					alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！");
					$('input[name=PhotoUrl]').focus();
					return;
				}
			   <%If KS.C_S(ChannelID,34)=0 Then%>
				if (frames["PictureContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
				$('#Content').val(frames["PictureContent"].KS_EditArea.document.body.innerHTML);
			   <%end if%>
			    
				$('#PicUrls').val('');
				<%If PicNum<100 Then%>
				for(var i=1;i<=$("#picnum").val();i++){
				  if ($('input[name=imgurl'+i+']').val()!=''&&$('input[name=imgurl'+i+']').val()!='del') 
				   {
				   var note=$('#imgnote'+i).val();
				   note=note.replace('|||','');
				   spic=$('input[name=imgurl'+i+']').val();
				   tpic=$('input[name=thumb'+i+']').val();
				   if (tpic=='') tpic=spic;
				  
				   if (spic.substring(0,4).toLowerCase()=='http'&&$("input[name=BeyondSavePic]").attr('checked')==true)
				   {
					 $('#LayerPrompt').show();
					 window.setInterval('ShowPromptMessage()',150)
				   }
				   if ($('input[name=PicUrls]').val()=='')                 
				   $('input[name=PicUrls]').val(note+'|'+spic+'|'+tpic);
				   else 
				   $('input[name=PicUrls]').val($('input[name=PicUrls]').val()+'|||'+note+'|'+spic+'|'+tpic);
				  }
				}
				<%end if%>
				var addmode;
				for (var i=0;i<document.PictureForm.addmode.length;i++){
				 var KM = document.PictureForm.addmode[i];
				if (KM.checked==true)	   
					addmode = KM.value
				}
		
				if (addmode==0 && $('input[name=PicUrls]').val()=='')
				{
				  alert('请输入<%=KS.C_S(ChannelID,3)%>内容!');
				  $('input[name=imgurl1]').focus();
				  return false;
				}
				  $('form[name=PictureForm]').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
			}

		var ForwardShow=true;
		function ShowPromptMessage()
			{
				var TempStr=ShowArticleArea.innerText;
				if (ForwardShow==true)
				{
					if (TempStr.length>4) ForwardShow=false;
					ShowArticleArea.innerText=TempStr+'.';
					
				}
				else
				{
					if (TempStr.length==1) ForwardShow=true;
					ShowArticleArea.innerText=TempStr.substr(0,TempStr.length-1);
				}
			}
		var SaveBeyondInfo=''
					   +'<div id="LayerPrompt" style="position:absolute;left: 200px; top: 200px; background-color: #f1efd9; layer-background-color: #f1efd9; border: 1px none #000000; width: 360px; height: 63px; display: none; "> '
					   +'<table width="100%" height="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF0000">'
					   +'<tr> '
					   +'<td align="center">'
					   +'<table width="80%" border="0" cellspacing="0" cellpadding="0">'
					   +'<tr>'
					   +' <td width="75%" nowrap>'
					   +'<div align="right">请稍候，系统正在保存远程图片到本地</div></td>'
					   +'   <td width="25%"><font id="ShowArticleArea">&nbsp;</font></td>'
					   +' </tr>'
					   +'</table>'
					   +'</td>'
					   +'</tr>'
					   +'</table>'
					   +'</div>'
			document.write (SaveBeyondInfo)
		</script>
    
<%
			 .Write "</body>"
			 .Write "</html>"
			 End With
		End Sub
		
		'保存
		Sub PictureSave()
		   Dim MoreStart,MoreEnd,MorePicUrl,MorePhotoUrl,I,SelectInfoList,HasInRelativeID
		  With Response
			
			Title = KS.G("Title")
			PictureContent= KS.FilterIllegalChar(Request.Form("Content"))
			Hits        = KS.ChkClng(KS.G("Hits"))
			HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
			HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
			HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))
			
			PhotoUrl     = KS.G("PhotoUrl")
			If KS.G("AddMode")="0" Then
			PicUrls     = KS.G("PicUrls")
			Else
			   MoreStart=KS.ChkClng(KS.G("MoreStart"))
			   MoreEnd=KS.ChkClng(KS.G("MoreEnd"))
			   If MoreStart>MoreEnd Then .Write "<script>alert('批量添加的结束ID必须大小开始ID!');history.back();</script>":.end
			   MorePicUrl=KS.G("MorePicUrl")
			   For I=MoreStart to MoreEnd
			    If PicUrls="" Then
				 PicUrls=Replace(MorePicUrl,"#",I)
				Else
				 PicUrls=PicUrls & "|||" & Replace(MorePicUrl,"#",I)
				End If
			   Next
			End If
			
			Recommend   = KS.ChkClng(KS.G("Recommend"))
			Rolls       = KS.ChkClng(KS.G("Rolls"))
			Strip       = KS.ChkClng(KS.G("Strip"))
			Popular     = KS.ChkClng(KS.G("Popular"))
			Comment     = KS.ChkClng(KS.G("Comment"))
			IsTop       = KS.ChkClng(KS.G("IsTop"))
			Slide       = KS.ChkClng(KS.G("Slide"))
			Makehtml    = KS.ChkClng(KS.G("Makehtml"))
			Tid = KS.G("Tid")
			SpecialID = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
			SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
			Verific=1
			KeyWords = KS.G("KeyWords")
			Author  = KS.G("Author")
			Origin  = KS.G("Origin")
			AddDate = KS.G("AddDate")
			If Not IsDate(AddDate) Then AddDate=Now
			Rank = Trim(KS.G("Rank"))
				
				'收费选项
				ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
				ChargeType  = KS.ChkClng(KS.G("ChargeType"))
				PitchTime   = KS.ChkClng(KS.G("PitchTime"))
				ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
				InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
				arrGroupID  = KS.G("GroupID")
				DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
				
				TemplateID = KS.G("TemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				Dim filetype:filetype=KS.ChkClng(KS.G("filetype"))
				Dim FnameType
				Dim RS_C:Set RS_C=Server.CreateObject("Adodb.RecordSet")
					RS_C.Open "Select * From KS_Class Where ID='" & Tid & "'",conn,1,1
					If Not RS_C.Eof Then
					    FnameType=RS_C("FnameType")
						If KS.ChkClng(KS.G("TemplateFlag"))=2 Or TemplateID="" Then TemplateID=RS_C("TemplateID"):WapTemplateID=RS_C("WapTemplateID")
						If FileType=0 Then
						  If Action = "Add" OR Action="Verify" Then
						   Fname=KS.GetFileName(RS_C("FsoType"), Now, "") & FnameType
						   End If
						End If
					End If
				RS_C.Close:Set RS_C=Nothing
				If filetype=1 Then Fname=KS.G("FileName")

    			UserDefineFieldArr=KSCls.Get_KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.G(UserDefineFieldArr(0,I))="" Then ErrMsg = ErrMsg & UserDefineFieldArr(1,I) & "必须填写!\n"
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.G(UserDefineFieldArr(0,I))) Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写数字!\n"
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.G(UserDefineFieldArr(0,I))) Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写正确的日期!\n"
                 If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.G(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写正确的Email格式!\n" 
				Next
			  End if			
			 
			If Title = "" Then .Write ("<script>alert('图片名称不能为空!');history.back(-1);</script>")
			If PhotoUrl = "" Then .Write ("<script>alert('图片缩略图不能为空!');history.back(-1);</script>")
			
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Tid = "" Then ErrMsg = ErrMsg & "[图片类别]必选! \n"
			If Title = "" Then ErrMsg = ErrMsg & "[图片标题]不能为空! \n"
			If Title <> "" And Tid <> "" And Action = "Add" Then
			  SqlStr = "select * from " & KS.C_S(ChannelID,2) & " where Title='" & Title & "' And Tid='" & Tid & "'"
			   RS.Open SqlStr, conn, 1, 1
				If Not RS.EOF Then
				 ErrMsg = ErrMsg & "该类别已存在此篇图片! \n"
			   End If
			   RS.Close
			End If
			If ErrMsg <> "" Then
			   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
			   .End
			Else
			      If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
				  
			      If KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
				  	SaveFilePath = KS.GetUpFilesDir & "/"
					KS.CreateListFolder (SaveFilePath)

				   Dim sPicUrlArr:sPicUrlArr=Split(PicUrls,"|||")
				   Dim sTemp
				   PicUrls=""
				   For I=0 To Ubound(sPicUrlArr)
				     If Left(Lcase(Split(sPicUrlArr(i),"|")(1)),4)="http" Then
					  sTemp=Split(sPicUrlArr(i),"|")(0) & "|" & KS.ReplaceBeyondUrl(Split(sPicUrlArr(i),"|")(1), SaveFilePath) &"|" & KS.ReplaceBeyondUrl(Split(sPicUrlArr(i),"|")(2), SaveFilePath)
					 Else
					  sTemp=sPicUrlArr(I)
					 End If
					 If I=0 Then
					   PicUrls=sTemp
					 Else
					   PicUrls=PicUrls & "|||" & sTemp
					 End If
				   Next
				   PhotoUrl= KS.ReplaceBeyondUrl(PhotoUrl, SaveFilePath)
				  End If
				  If Action = "Add" Then
					SqlStr = "select * from " & KS.C_S(ChannelID,2) & " where 1=0"
					RS.Open SqlStr, conn, 1, 3
					RS.AddNew
					RS("Title") = Title
					RS("PhotoUrl") = PhotoUrl
					RS("PictureContent") = PictureContent
					RS("PicUrls") = PicUrls
					RS("Recommend") = Recommend
					RS("Rolls") = Rolls
					RS("Strip") = Strip
					RS("Popular") = Popular
					RS("Verific") = Verific
					RS("Comment") = Comment
					RS("IsTop") = IsTop
					RS("Tid") = Tid
					RS("KeyWords") = KeyWords
					RS("Author") = Author
					RS("Origin") = Origin
					RS("AddDate") = AddDate
					RS("Rank") = Rank
					RS("Slide") = Slide
					RS("TemplateID") = TemplateID
					RS("WapTemplateID")  = WapTemplateID
					RS("Hits") = Hits
					RS("HitsByDay") = HitsByDay
					RS("HitsByWeek") = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("Fname") = Fname
					RS("Inputer") = KS.C("AdminName")
					RS("RefreshTF") = Makehtml
					RS("Score") = KS.ChkClng(KS.G("Score"))
					RS("DelTF") = 0
					RS("ReadPoint")=ReadPoint
				    RS("ChargeType")=ChargeType
				    RS("PitchTime")=PitchTime
				    RS("ReadTimes")=ReadTimes
					RS("InfoPurview")=InfoPurview
					RS("arrGroupID")=arrGroupID
					RS("DividePercent")=DividePercent
					If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RS("" & UserDefineFieldArr(0,I) & "")=Request(UserDefineFieldArr(0,I))
							else
							 RS("" & UserDefineFieldArr(0,I) & "")=KS.G(UserDefineFieldArr(0,I))
							end if
							If UserDefineFieldArr(12,I)="1"  Then
							RS("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
							End If
						Next
					End If
					RS.Update
					
				   '写入Session,添加下一篇图片调用
				   Session("KeyWords") = KeyWords
				   Session("Author") = Author
				   Session("Origin") = Origin
                   RS.MoveLast
				   If Left(Ucase(Fname),2)="ID" Then
					   RS("Fname") = RS("ID") & FnameType
					   RS.Update
					End If
					
					For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
					Next
					
					If SelectInfoList<>"" Then
					 SelectInfoList=Split(SelectInfoList,",")
					 For I=0 To Ubound(SelectInfoList)
					  If KS.FoundInArr(HasInRelativeID,SelectInfoList(i),",")=false Then
					   Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & RS("ID") & "," & Split(SelectInfoList(i),"|")(0) & "," & Split(SelectInfoList(i),"|")(1) & ")")
					   HasInRelativeID=HasInRelativeID & SelectInfoList(i) & ","
					  End If
					 Next
					End If
					
 					Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
	 				'关联上传文件
					 Call KS.FileAssociation(ChannelID,RS("ID"),PicUrls & PhotoUrl & PictureContent,0)

			        Call RefreshHtml(1)
					RS.Close:Set RS = Nothing
					
				ElseIf Action = "Edit" Or Action="Verify"  Then
				PicID = KS.ChkCLng(Request("PicID"))
				SqlStr = "SELECT * FROM " & KS.C_S(ChannelID,2) & " Where ID=" & PicID
					RS.Open SqlStr, conn, 1, 3
					If RS.EOF And RS.BOF Then
					 .Write ("<script>alert('参数传递出错!');history.back(-1);</script>")
					 .End
					End If
					RS("Title") = Title
					RS("PhotoUrl") = PhotoUrl
					RS("PictureContent") = PictureContent
					RS("PicUrls") = PicUrls
					RS("Recommend") = Recommend
					RS("Rolls") = Rolls
					RS("Strip") = Strip
					RS("Popular") = Popular
					RS("Comment") = Comment
					RS("IsTop") = IsTop
					RS("Tid") = Tid
					RS("KeyWords") = KeyWords
					RS("Author") = Author
					RS("Origin") = Origin
					RS("AddDate") = AddDate
					RS("Rank") = Rank
					RS("Slide") = Slide
					RS("TemplateID") = TemplateID
					RS("WapTemplateID")  = WapTemplateID
					If Makehtml = 1 Then
					 RS("RefreshTF") = 1
					End If
					RS("Hits") = Hits
					RS("HitsByDay") = HitsByDay
					RS("HitsByWeek") = HitsByWeek
					RS("HitsByMonth") = HitsByMonth
					RS("Score") = KS.ChkClng(KS.G("Score"))
					RS("ReadPoint")=	ReadPoint
				    RS("ChargeType")=ChargeType
				    RS("PitchTime")=PitchTime
				    RS("ReadTimes")=ReadTimes
					RS("InfoPurview")=InfoPurview
					RS("arrGroupID")=arrGroupID
					RS("DividePercent")=DividePercent
					If Action="Verify" Then
					  Inputer=RS("Inputer")
					End If
					RS("Verific") = Verific
					
					If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RS("" & UserDefineFieldArr(0,I) & "")=Request(UserDefineFieldArr(0,I))
							else
							 RS("" & UserDefineFieldArr(0,I) & "")=KS.G(UserDefineFieldArr(0,I))
							end if
							If UserDefineFieldArr(12,I)="1"  Then
							RS("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
							End If
						Next
					End If
					RS.Update
                   RS.MoveLast
			       If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"), Split(RS("PicUrls"), "|||"),RS("Fname"))
				   End If
						Conn.Execute("Delete From KS_SpecialR Where InfoID=" & RS("ID") & " and channelid=" & ChannelID)
						For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						Next
						
						 Conn.Execute("Delete From KS_ItemInfoR Where InfoID=" & RS("ID") & " and channelid=" & ChannelID)
						If SelectInfoList<>"" Then
						 SelectInfoList=Split(SelectInfoList,",")
						 For I=0 To Ubound(SelectInfoList)
						  If KS.FoundInArr(HasInRelativeID,SelectInfoList(i),",")=false Then
						   Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & RS("ID") & "," & Split(SelectInfoList(i),"|")(0) & "," & Split(SelectInfoList(i),"|")(1) & ")")
						   HasInRelativeID=HasInRelativeID & SelectInfoList(i) & ","
						  End If
						 Next
						End If
						
					Call LFCls.UpdateItemInfo(ChannelID,PicID,Title,Tid,PictureContent,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				'关联上传文件
					 Call KS.FileAssociation(ChannelID,PicID,PicUrls & PhotoUrl & PictureContent,1)
				   Call RefreshHtml(2)
		          
				  RS.Close:Set RS = Nothing
					IF Action="Verify" Then     '如果是审核投稿图片，对用户，进行加积分等，并返回签收图片管理
							  '对用户进行增值，及发送通知操作
							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,Title,PicID)
							 .Write ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) &"成功签收,系统已发送一封站内通知信给投稿者!');location.href='KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>签收会员" & KS.C_S(ChannelID,3)) &"</font>';</script>") 
							 
				    End If
					If KeyWord <>"" Then
						 .Write ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) &"修改成功!');location.href='KS.Picture.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=PictureSearch&OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) &" >> <font color=red>搜索结果</font>") & "';</script>")
					End If
				End If
			End If
		 End With		
		End Sub
		
			Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改" & KS.C_S(ChannelID,3):AddStr="继续添加" & KS.C_S(ChannelID,3)
				Else
				  TempStr="修改":EditStr="继续修改" & KS.C_S(ChannelID,3):AddStr="添加" & KS.C_S(ChannelID,3)
				End If
			    With Response
				     .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
					 .Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
					  .Write "	  <tr class=""sort""> "
					  .Write "		<td  height=""28"" colspan=2>系统操作提示信息</td>" & vbcrlf
					  .Write "	  </tr>"
                      .Write "    <tr class='tdbg'>"
					  .Write "          <td align='center'><img src='images/succeed.gif'></td>"
					  .Write "<td><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;恭喜，" & TempStr &"" & KS.C_S(ChannelID,3) & "成功！</b><br>"

					   If Makehtml = 1 Then
					      .Write "<div style=""margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%"">" 
					    If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then
						  	 .Write "<div><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "没有启用生成HTML的功能，所以ID号为 <font color=red>" & RS("ID") & "</font>  的" & KS.C_S(ChannelID,3) & "没有生成!</li></div> "
						  End If
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "的栏目页没有启用生成HTML的功能，所以ID号为 <font color=red>" & TID & "</font>  的栏目没有生成!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp" or KS.C_S(ChannelID,9)<>3 Then
					   ' .Write "<div style=""margin-left:140;color:blue;height:25px""><li>由于 <a href=""" & KS.GetDomain & """ target=""_blank""><font color=red>网站首页</font></a> 没有启用生成HTML的功能或发布选项没有开启，所以没有生成!</li></div>"
					   Else
					     .Write "<div align=center><iframe src=""Include/RefreshIndex.asp?RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div>"
					 End If
					  .Write   "</td></tr>"
					  .Write "	  <tr class='tdbg'>"
					  .Write "		<td height=""25"" colspan=2 align=""right"">【<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & "';""><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Picture.asp?ChannelID=" & ChannelID & "&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr="&server.URLEncode("添加" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID & "&Page=" & Page&"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain &"Item/Show.asp?M=" & ChannelID & "&D=" & RS("ID") & """ target=""_blank""><strong>预览" & KS.C_S(ChannelID,3) & "内容</strong></a>】</td>"
					  .Write "	  </tr>"
					  .Write "	</table>"				
			End With
		End Sub

End Class
%> 

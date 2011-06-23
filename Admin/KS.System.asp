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
Set KSCls = New Admin_System
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_System
        Private KS,KSMCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSMCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call KS.DelCahe(KS.SiteSn & "_Config")
		 Call KS.DelCahe(KS.SiteSn & "_Date")
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
		  	.Write "<html>"
			.Write "<title>网站基本参数设置</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<style type=""text/css"">"
			.Write "<!--" & vbCrLf
			.Write ".STYLE1 {color: #FF0000}" & vbCrLf
			.Write ".STYLE2 {color: #FF6600}" & vbCrLf
			.Write "-->" & vbCrLf
			.Write "</style>" & vbCrLf
			.Write "</head>" & vbCrLf

		  Select Case KS.G("Action")
		  
		   Case "Space"
		     	If Not KS.ReturnPowerResult(0, "KMST10010") Then          
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetSpaceInfo()
				End If
		   Case "CopyRight"
		     	If Not KS.ReturnPowerResult(0, "KMST10011") Then         
				   Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
				   Call KS.ReturnErr(1, ""): Exit Sub
				Else
		           Call GetCopyRightInfo()
				End If
		   Case Else
		       Call SetSystem()
		  End Select
		 End With
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		With Response
			
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '检查是否有基本信息设置的权限
					 .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			
			dim strDir,strAdminDir
			strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
			InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
			If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
			   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
			End If
	
	
			SqlStr = "select * from KS_Config"
			Set RS = KS.InitialObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			 Dim Setting:Setting=Split(RS("Setting")&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			 Dim TBSetting:TBSetting=Split(RS("TBSetting"),"^%^")
			 FsoIndexFile = Split(Setting(5), ".")(0)
			 FsoIndexExt = Split(Setting(5), ".")(1)
			If KS.G("Flag") = "Edit" Then
			            'IP设置
			            Dim LockIP,I,PartIPArr
						Dim LockIPWhite:LockIPWhite=KS.G("LockIPWhite")
						Dim LockIPBlack:LockIPBlack=KS.G("LockIPBlack")
						If  LockIPWhite<>"" Then
							Dim LockIPWhiteArr:LockIPWhiteArr=Split(LockIPWhite,vbcrlf)
							For I=0 To Ubound(LockIPWhiteArr)
							 If LockIPWhiteArr(i)<>"" and instr(LockIPWhiteArr(i),"----")>0 Then
								 PartIPArr=Split(LockIPWhiteArr(i),"----")
								 If I=0 Then
								   LockIP=LockIP & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
								 Else
								   LockIP=LockIP & "$$$" & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
								 End IF
							 End If
							Next
						End If
						LockIP=LockIP &"|||"
					If LockIPBlack<>"" Then
						Dim LockIPBlackArr:LockIPBlackArr=Split(LockIPBlack,vbcrlf)
						For I=0 To Ubound(LockIPBlackArr)
						 If LockIPBlackArr(i)<>"" and instr(LockIPBlackArr(i),"----")>0 Then
							 PartIPArr=Split(LockIPBlackArr(i),"----")
							 If I=0 Then
							  LockIP=LockIP & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
							 Else
							  LockIP=LockIP & "$$$" & KS.EncodeIP(PartIPArr(0))&"----" & KS.EncodeIP(PartIPArr(1))
							 End iF
						 End If	 
						Next
					End If
					
				Dim FZCJYM
				For N=1 To 10
				  FZCJYM=FZCJYM & KS.ChkClng(Request.Form("Opening" & N))
				Next
					
			    Dim WebSetting,ThumbSetting
				For n=0 To 170
				  If n=5 Then
				   WebSetting=WebSetting & KS.G("Setting(5)") & KS.G("FsoIndexExt") & "^%^"
				  ElseIF n=101 Then
				   WebSetting=WebSetting &LockIP & "^%^"
				  ElseIf n=161 Then
				   WebSetting=WebSetting & FZCJYM & "^%^"
				  Else
				   WebSetting=WebSetting & Replace(Request.Form("Setting(" & n &")"),"^%^","") & "^%^"
				  End If
                   
				Next
				
				For I=0 To 20
				 If I=13 Then
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBLogo"),"^%^","") & "^%^"
				 Else
				  ThumbSetting=ThumbSetting & Replace(KS.G("TBSetting(" & I &")"),"^%^","") & "^%^"
				 End If
				Next
				RS("Setting")=WebSetting
				RS("TBSetting")=ThumbSetting
				RS.Update
				Call KS.FileAssociation(1015,1,WebSetting&ThumbSetting,1)
				RS.Close:Set RS=Nothing
			   .Write ("<script>alert('网站配置信息修改成功！');parent.frames['FrameTop'].location.href='index.asp?Action=Head&C=1';location.href='KS.System.asp';</script>")
			End If
			

			.Write "<body oncontextmenu='return false' bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>网站基本信息设置</div>"
			.Write "<div style='height:5px;overflow:hidden'></div>"
			.Write "<div class=tab-page id=configPane>"
			.Write "  <form name='myform' method=post action="""" id=""myform"" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""configPane"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>网站名称：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(0)"" type=""text"" id=""Setting(0)"" value=""" & Setting(0) & """ size=""30""></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>网站标题：</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(1)"" type=""text"" id=""Setting(1)"" value=""" & Setting(1) & """ size=""30""></td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td width=""32%"" height=""30"" class=""CleftTitle""> <div align='right'><strong>网站地址：</strong></div><font color=""#FF0000"">系统会自动获得正确的路径，但需要手工保存设置</font></td>"
			 .Write "    <td height=""30""> <input name=""Setting(2)"" type=""text""  value=""" &KS.GetAutoDomain & """ size=""30"">"
			 .Write "      (请使用http://标识),后面不要带&quot;/&quot;符号 </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>安装目录：</strong></div><font color=""#FF0000"">系统会自动获得正确的路径，但需要手工保存设置</font></td>"
			 .Write "     <td height=""30""> <input name=""Setting(3)"" type=""text"" id=""Setting(3)""  value=""" & InstallDir & """ readonly size=30>"
			 .Write "       系统安装的虚拟目录</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>网站Logo地址：</strong></div></td>"
			  .Write "    <td height=""30""><input name=""Setting(4)"" type=""text"" id=""Setting(4)""   value=""" & Setting(4) & """ size=30>"
			  .Write "      申请友情链接时显示给用户</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>生成的网站首页：</strong></div></td>"
				.Write "  <td height=""30""> <input type=""radio"" name=""Setting(5)"" value=""Index"" "
				
				If FsoIndexFile = "Index" Then .Write (" checked")
				.Write ">"
				.Write "    Index"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""Default"" "
				If FsoIndexFile = "Default" Then .Write (" checked")
				.Write ">"
				.Write "    Default"
				.Write "    <select name=""FsoIndexExt"" id=""select"">"
				.Write "      <option value="".htm"" "
				If FsoIndexExt = "htm" Then .Write ("selected")
				.Write ">.htm</option>"
				.Write "      <option value="".html"" "
				If FsoIndexExt = "html" Then .Write ("selected")
				.Write ">.html</option>"
				.Write "      <option value="".shtml"" "
				If FsoIndexExt = "shtml" Then .Write ("selected")
				.Write ">.shtml</option>"
				.Write "      <option value="".shtm"" "
				If FsoIndexExt = "shtm" Then .Write ("selected")
				.Write ">.shtm</option>"
				.Write "      <option value="".asp"" "
				If FsoIndexExt = "asp" Then .Write ("selected")
				.Write ">.asp</option>"
				.Write "    </select>&nbsp;<font color=blue>扩展名为.asp，首页将不启用生成静态HTML的功能</font></td>"
				.Write "</tr>"
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>专题是否启用生成：</strong></div></td>"
				.Write "  <td height=""30""><input type=""radio"" name=""Setting(78)"" value=""1"" "
				
				If Setting(78) = "1" Then .Write (" checked")
				.Write ">启用"
				.Write "    <input type=""radio"" name=""Setting(78)"" value=""0"" "
				If Setting(78) = "0" Then .Write (" checked")
				.Write ">不启用"
			   .Write "  　</td>"
			   .Write "    </tr>"
			
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>默认允许上传最大文件大小：</strong></div></td>"
				.Write "  <td height=""30""><input name=""Setting(6)"" onBlur=""CheckNumber(this,'允许上传最大文件大小');"" type=""text"" id=""Setting(6)""   value=""" & Setting(6) & """ size=15>"
			.Write "KB 　 <span class=""STYLE2"">提示：1 KB = 1024 Byte，1 MB = 1024 KB</span></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>默认允许上传文件类型：</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(7)"" type=""text"" id=""Setting(7)""   value=""" & Setting(7) & """ size='30'><font color=red> 多个类型用|线隔开</font></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align=""right""><strong>删除不活动用户时间：</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(8)"" type=""text""  value=""" &  Setting(8) & """ size=8> 分钟</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align=""right""><strong>文章自动分页每页大约字符数：</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(9)"" type=""text"" value=""" & Setting(9) & """ size=8> 个字符&nbsp;&nbsp;<font color=red>如果不想自动分页，请输入""0""</font></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CLeftTitle"" align=""right""><div><strong>站长姓名：</strong></div></td>"
			.Write "      <td height=""30""> <input name=""Setting(10)"" type=""text""   value=""" & Setting(10) & """ size=30></td>"
			.Write "    </tr>"


			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>要屏蔽的关键字：</strong></div><font color=red>说明：过滤字符设定规则为 要过滤的字符=过滤后的字符 ，每个过滤字符用回车分割开。作用范围所有模型的内容、评论、问答及小论坛等。</font></td>"
			 .Write "    <td height=""30""><textarea name=""Setting(55)"" cols=""30"" rows=""6"">" & Setting(55) & "</textarea></td></tr>"

			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class='CleftTitle' align=""right""><div><strong>页面发布时顶部信息：</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(15)"" type=""text""  value=""" & Setting(15) & """ size=30>"
			 .Write "     填写<span class=""STYLE1"">&quot;0&quot;</span>将不显示</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""cleftTitle"" align=""right""><div><strong>官方信息显示：</strong></div></td>"
			 .Write "  <td height=""30""> <input type=""checkbox"" name=""Setting(16)"" value=""1"" "
				
				If instr(Setting(16),"1")>0 Then .Write (" checked")
				.Write ">"
				.Write "    显示顶部公告"
				.Write "    <input type=""checkbox"" name=""Setting(16)"" value=""2"" "
				If instr(Setting(16),"2")>0 Then .Write (" checked")
				.Write ">"
				.Write "    显示论坛新帖"

			 .Write "     </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>官方授权的唯一系列号：</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(17)"" type=""text""  value=""" & Setting(17) & """ size=30>"
			 .Write "     免费版请填写<span class=""STYLE1"">&quot;0&quot;</span></td>"
			 .Write "   </tr>"
			   
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>网站的版权信息：</strong></div><font color=""#FF0000""> 用于显示网站版本等，支持html语法</font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(18)"" cols=""60"" rows=""5"">" & Setting(18) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>网站META关键词：</strong></div><font color=""#FF0000""> 针对搜索引擎设置的网页关键词,多个关键词请用,号分隔 </font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(19)"" cols=""60"" rows=""5"">" & Setting(19) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>网站META网页描述：</strong></div><font color=""#FF0000""> 针对搜索引擎设置的网页描述,多个描述请用,号分隔&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  </font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(20)"" cols=""60"" rows=""5"">" & Setting(20) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=site-template>"
			.Write "  <H2 class=tab>模板绑定</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-template"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>网站首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(110)"" id=""Setting110"" type=""text"" value=""" & Setting(110) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting110')[0]") & " <a href='../index.asp' target='_blank' style='color:green'>页面:/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr  valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>全站tags模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(120)"" id=""Setting120"" type=""text"" value=""" & Setting(120) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting120')[0]") & " <a href='../plus/tags.asp' target='_blank' style='color:green'>页面:/plus/tags.asp</a></td>"
			.Write "    </tr>"			
			.Write "    <tr  valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>全站搜索模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(139)"" id=""Setting139"" type=""text"" value=""" & Setting(139) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting139')[0]") & " <a href='../plus/search/' target='_blank' style='color:green'>页面:/plus/search/</a></td>"
			.Write "    </tr>"			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>专题首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(111)"" id=""Setting111"" type=""text"" value=""" & Setting(111) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting111')[0]") & " <a href='../specialindex.asp' target='_blank' style='color:green'>页面:/specialindex.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>公告模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(112)"" id=""Setting112"" type=""text"" value=""" & Setting(112) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting112')[0]") & " <a href='#' style='color:green'>页面:/plus/announce/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>友情链接页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(113)"" id=""Setting113"" type=""text"" value=""" & Setting(113) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting113')[0]") & " <a href='../plus/link/' target='_blank' style='color:green'>页面:/plus/link</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>留言/小论坛首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(114)"" id=""Setting114"" type=""text"" value=""" & Setting(114) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting114')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>页面:/club/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>留言/小论坛签写页面模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(115)"" id=""Setting115"" type=""text"" value=""" & Setting(115) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting115')[0]") & " <a href='../club/post.asp' target='_blank' style='color:green'>页面:/club/post.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PK首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(102)"" id=""Setting102"" type=""text"" value=""" & Setting(102) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting102')[0]") & " <a href='../plus/pk/index.asp' target='_blank' style='color:green'>页面:/plus/pk/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PK页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(103)"" id=""Setting103"" type=""text"" value=""" & Setting(103) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting103')[0]") & " <a href='#' style='color:green'>页面:/plus/pk/pk.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PK观点更多页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(104)"" id=""Setting104"" type=""text"" value=""" & Setting(104) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting104')[0]") & " <a href='#' style='color:green'>页面:/plus/pk/more.asp</a></td>"
			.Write "    </tr>"

			

			.Write "    <tr>"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>会员首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(116)"" id=""Setting116"" type=""text"" value=""" & Setting(116) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting116')[0]") & " <a href='../user/' target='_blank' style='color:green'>页面:/user/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>会员注册表单1模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(117)"" id=""Setting117"" type=""text"" value=""" & Setting(117) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting117')[0]") & " <a href='../user/reg/' target='_blank' style='color:green'>页面:/user/reg/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>会员注册表单2模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(118)"" id=""Setting118"" type=""text"" value=""" & Setting(118) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting118')[0]") & " <a href='../user/reg/' target='_blank' style='color:green'>页面:/user/reg/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>会员注册成功页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(119)"" id=""Setting119"" type=""text"" value=""" & Setting(119) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting119')[0]") & "</td>"
			.Write "    </tr>"

			
			dim dis
			if conn.execute("select ChannelStatus from ks_channel where channelid=5")(0)=1 then
			 dis=""
			else
			 dis=" style='display:none'"
			end if
			.Write "    <tr" & dis &">"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城购物车模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(121)"" id=""Setting121"" type=""text"" value=""" & Setting(121) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting121')[0]") & " <a href='../shop/shoppingcart.asp' target='_blank' style='color:green'>页面:/shop/shoppingcart.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城收银台模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(122)"" id=""Setting122"" type=""text"" value=""" & Setting(122) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting122')[0]") & " <a href='../shop/payment.asp' target='_blank' style='color:green'>页面:/shop/payment.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城订单预览模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(123)"" id=""Setting123"" type=""text"" value=""" & Setting(123) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting123')[0]") & " <a href='../shop/Preview.asp' target='_blank' style='color:green'>页面:/shop/Preview.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城订购成功模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(124)"" id=""Setting124"" type=""text"" value=""" & Setting(124) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting124')[0]") & " <a href='../shop/order.asp' target='_blank' style='color:green'>页面:/shop/order.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城购物指南模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(125)"" id=""Setting125"" type=""text"" value=""" & Setting(125) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting125')[0]") & " <a href='../shop/ShopHelp.asp' target='_blank' style='color:green'>页面:/shop/ShopHelp.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城银行付款模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(126)"" id=""Setting126"" type=""text"" value=""" & Setting(126) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting126')[0]") & " <a href='../shop/showpay.asp' target='_blank' style='color:green'>页面:/shop/showpay.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城品牌列表页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(135)"" id=""Setting135"" type=""text"" value=""" & Setting(135) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting135')[0]") & " <a href='../shop/showbrand.asp' target='_blank' style='color:green'>页面:/shop/showbrand.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城品牌详情页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(136)"" id=""Setting136"" type=""text"" value=""" & Setting(136) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting136')[0]") & " <a href='../shop/search_list.asp' target='_blank' style='color:green'>页面:/shop/search_list.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城团购首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(137)"" id=""Setting137"" type=""text"" value=""" & Setting(137) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting137')[0]") & " <a href='../shop/groupbuy.asp' target='_blank' style='color:green'>页面:/shop/groupbuy.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>商城团购内容页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(138)"" id=""Setting138"" type=""text"" value=""" & Setting(138) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting138')[0]") & " <a href='../shop/groupbuyshow.asp' target='_blank' style='color:green'>页面:/shop/groupbuyshow.asp</a></td>"
			.Write "    </tr>"
			
			
			if conn.execute("select ChannelStatus from ks_channel where channelid=6")(0)=1 then
			 dis=""
			else
			 dis=" style='display:none'"
			end if
			.Write "    <tr" & dis &">"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>音乐首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(127)"" id=""Setting127"" type=""text"" value=""" & Setting(127) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting127')[0]") & " <a href='../music/' target='_blank' style='color:green'>页面:/music/</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>歌手页面模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(128)"" id=""Setting128"" type=""text"" value=""" & Setting(128) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting128')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>专辑列表页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(129)"" id=""Setting129"" type=""text"" value=""" & Setting(129) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting129')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>最终专辑页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(130)"" id=""Setting130"" type=""text"" value=""" & Setting(130) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting130')[0]") & "</td>"
			.Write "    </tr>"
			
			
			if not conn.execute("select ChannelStatus from ks_channel where channelid=9").eof then
			 if conn.execute("select ChannelStatus from ks_channel where channelid=9")(0)=1 then
			 dis=""
			 else
			 dis=" style='display:none'"
			 end if
			else
			 dis=" style='display:none'"
			end if
			.Write "    <tr" & dis &">"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>考试系统首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(131)""  id=""Setting131"" type=""text"" value=""" & Setting(131) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting131')[0]") & " <a href='../mnkc/' target='_blank' style='color:green'>页面:/mnkc/</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>试卷分类页面模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(132)"" id=""Setting132"" type=""text"" value=""" & Setting(132) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting132')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>试卷内容页面模板(答题卡方式)：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(133)"" id=""Setting133"" type=""text"" value=""" & Setting(133) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting133')[0]") & "</td>"
			.Write "    </tr>"
            .Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>试卷内容页面模板(普通方式)：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(105)"" id=""Setting105"" type=""text"" value=""" & Setting(105) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting105')[0]") & "</td>"
			.Write "    </tr>"			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>试卷总分类模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(134)"" id=""Setting134"" type=""text"" value=""" & Setting(134) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting134')[0]") & " <a href='../mnkc/all.html' target='_blank' style='color:green'>页面:/mnkc/all.html</a></td>"
			.Write "    </tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			 '=================================================防注册机选项========================================
			 .Write "<div class=tab-page id=ZCJ_Option>"
			 .Write " <H2 class=tab>防注册机</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""ZCJ_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			
             .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>要启用防注册机的页面：</strong></div></td>"
			
			
			.Write "      <td height=""21"">"
			.Write "<input type='checkbox' name='Opening1' value='1'"
			If mid(Setting(161),1,1)="1" Then .Write "checked"
			.Write ">会员注册页面"
			.Write "<br/><input type='checkbox' name='Opening2' value='1'"
			If mid(Setting(161),2,1)="1" Then .Write "checked"
			.Write ">匿名投稿发布页面"
			'.Write "<br/><input type='checkbox' name='Opening3' value='1'"
			'If mid(Setting(161),3,1)="1" Then .Write "checked"
			'.Write ">评论发表页面"
		    .Write "      </td>"	
			.Write "</tr>"			
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>验证问题：</strong></div>可以设置多个,一行一个验证选项</td>"
            .Write "    <td><textarea name='Setting(162)' style='width:350px;height:120px'>" & Setting(162) & "</textarea></td>"
			.Write "    </tr>"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>验证答案：</strong></div>对应验证问题的选项,一行一个验证答案</td>"
            .Write "    <td><textarea name='Setting(163)' style='width:350px;height:120px'>" & Setting(163) & "</textarea></td>"
			.Write "    </tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			
									 '=====================================================会员注册参数配置开始=========================================

		.Write " <div class=tab-page id=User_Option>"
		.Write "	  <H2 class=tab>会员选项</H2>"
		.Write "		<SCRIPT type=text/javascript>"
		.Write "					 tabPane1.addTabPage(document.getElementById( ""User_Option"" ));"
		.Write "		</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>是否允许新会员注册：</strong></div></td>"
			.Write "      <td height=""21""><input name=""Setting(21)"" type=""radio"" value=""1"""
			 If Setting(21)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(21)"" type=""radio"" value=""0"""
			 If Setting(21)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"	
			 .Write "</tr>"		
			 .Write "<tr style=""display:none"" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='clefttitle' align=""right""><div><strong>新会员注册需要阅读会员协议：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio""  value=""1"""
			 If Setting(22)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio"" value=""0"""
			 If Setting(22)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""liencearea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='clefttitle' align=""right""><div><strong>新会员注册服务条款和声明：</strong><div><div align=center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;标签说明：</div>{$GetSiteName}：网站名称<br>{$GetSiteUrl}：网站URL<br>{$GetWebmaster}：站长<br>{$GetWebmasterEmail}：站长信箱</td>"
			.Write "      <td height=""21""><textarea name=""Setting(23)"" cols=""70"" rows=""7"">" & Setting(23) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			
			
			 .Write "<tr width=""32%"" height=""21"" id=""grouparea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "<td width='40%' class='CleftTitle' align=""right""><div><strong>是否启用用户组注册：</strong></div><font color=red>如果不启用,默认注册类型为个人会员</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(33)"" type=""radio"" value=""1"""
			 If Setting(33)="1" Then .Write " Checked"
			 .Write ">启用"
			 .Write " &nbsp;&nbsp;<input name=""Setting(33)"" type=""radio"" value=""0"""
			 If Setting(33)="0" Then .Write " Checked"
			 .Write ">不启用"
			 .Write "</td>"
			 .Write "</tr>" 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员注册流程：</strong> </div></td>"
			.Write "      <td height=""21""> <input name=""Setting(32)"" type=""radio"" value=""1"""
			 If Setting(32)="1" Then .Write " Checked"
			 .Write ">一步简单完成注册<br>"
			 .Write "<input name=""Setting(32)"" type=""radio"" value=""2"""
			 If Setting(32)="2" Then .Write " Checked"
			 .Write ">两步完成注册（需要填写对应用户组的表单）"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员注册成功是否发邮件通知：</strong></div><font color=blue>用户组设置成需要邮件验证时,只有激活成功才会发送。</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""1"""
			 If Setting(146)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(146)"" type=""radio"" onclick=""setsendmail(0)"" value=""0"""
			 If Setting(146)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""sendmailarea""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员注册成功发送的邮件通知内容：</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div><div align=center>标签说明：<br>{$UserName}：用户名<br>{$PassWord}：密码<br>{$SiteName}：网站名称<div></td>"
			.Write "      <td height=""21""><textarea name=""Setting(147)"" cols=""70"" rows=""5"">" & Setting(147) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>注册会员密码问题是否必填：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(148)"" type=""radio"" value=""1"""
			 If Setting(148)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(148)"" type=""radio"" value=""0"""
			 If Setting(148)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"			 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>注册会员手机号码是否必填：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(149)"" type=""radio"" value=""1"""
			 If Setting(149)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(149)"" type=""radio"" value=""0"""
			 If Setting(149)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>一个IP只能注册一个会员：</strong></div><font color=red>若选择是，那么一个IP地址只能注册一次</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(26)"" type=""radio"" value=""1"""
			 If Setting(26)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(26)"" type=""radio"" value=""0"""
			 If Setting(26)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员注册时是否启用验证码功能：</strong></div><font color=red>启用验证码功能可以在一定程度上防止暴力营销软件或注册机自动注册</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(27)"" type=""radio"" value=""1"""
			 If Setting(27)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(27)"" type=""radio"" value=""0"""
			 If Setting(27)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>每个Email是否允许注册多次：</strong></div><font color=red>若选择是，则利用同一个Email可以注册多个会员。</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(28)"" type=""radio"" value=""1"""
			 If Setting(28)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(28)"" type=""radio"" value=""0"""
			 If Setting(28)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>新会员注册时用户名：</strong></div></td>"
			.Write "      <td height=""21""> 最少字符数<input name=""Setting(29)"" type=""text"" onBlur=""CheckNumber(this,'用户名最小字符数');"" size=""3"" value=""" & Setting(29) & """>个字符  最多字符数<input name=""Setting(30)"" type=""text"" onBlur=""CheckNumber(this,'用户名最多字符数');"" size=""3"" value=""" & Setting(30)& """>个字符"
			.Write "       </td>" 
	        .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>禁止注册的用户名：</strong> </div><font color=red>在右边指定的用户名将被禁止注册，每个用户名请用“|”符号分隔</font></td>"
			.Write "      <td height=""21""> <textarea name=""Setting(31)"" cols=""50"" rows=""3"">" & Setting(31) & "</textarea>"
			.Write "       </td>" 
			.Write "</tr>" 
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员登录时是否启用验证码功能：</strong></div><font color=red>启用验证码功能可以在一定程度上防止会员密码被暴力破解</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(34)"" type=""radio"" value=""1"""
			 If Setting(34)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(34)"" type=""radio"" value=""0"""
			 If Setting(34)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>只允许一个人登录： </strong></div><font color=red>启用此功能可以有效防止一个会员账号多人使用的情况</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(35)"" type=""radio"" value=""1"""
			 If Setting(35)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(35)"" type=""radio"" value=""0"""
			 If Setting(35)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
             .Write "</tr>"

			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>新会员注册时赠送的资金</strong>：</div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'新会员注册时赠送的金钱');"" name=""Setting(38)"" type=""text"" size=""5"" value=""" & Setting(38) & """>"
			.Write "元人民币（为0时不赠送）,此资金可用于商城中心购物.</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>新会员注册时赠送的积分：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(39)"" onBlur=""CheckNumber(this,'新会员注册时赠送的积分');"" type=""text"" size=""5"" value=""" & Setting(39) & """>"
			.Write "分积分</td>"
	        .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>新会员注册时赠送的点券：</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'新会员注册时赠送的点券');"" name=""Setting(40)"" type=""text"" size=""5"" value=""" & Setting(40) & """>"
			.Write "点点券（为0时不赠送）<br/><font color=blue>如果用户组选择了扣点用户,并设置了默认点数,则以用户组里的设置为准</font></td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员的积分与点券的兑换比率：</strong> </div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的积分与点券的兑换比率');"" name=""Setting(41)"" type=""text"" size=""5"" value=""" & Setting(41) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 点点券</td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员的积分与有效期的兑换比率：</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的积分与有效期的兑换比率');"" name=""Setting(42)"" type=""text"" size=""5"" value=""" & Setting(42) & """>"
			.Write "分积分可兑换 <font color=red>1</font> 天有效期</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员的资金与点券的兑换比率</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的资金与点券的兑换比率');"" name=""Setting(43)"" type=""text"" size=""5"" value=""" & Setting(43) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 点点券</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员的资金与有效期的兑换比率</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(44)"" type=""text"" size=""5"" value=""" & Setting(44) & """>"
			.Write "元人民币可兑换 <font color=red>1</font> 天有效期</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><strong>点券设置：</strong></td>"
			.Write "      <td height=""21""> 名称<input name=""Setting(45)"" type=""text"" size=""5"" value=""" & Setting(45) & """><font color=red>例如：科汛币、点券、金币</font>  单位<input name=""Setting(46)"" type=""text"" size=""5"" value=""" & Setting(46) & """> <font color=red>例如：点、个</font>"
			.Write "</td>"
			.Write "</tr>"

			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21""  class='CleftTitle' align=""right""><div><strong>会员站内短信设置：</strong></div></td>"
			.Write "      <td height=""21""> 最大容量为<input onBlur=""CheckNumber(this,'请填写有效条数!');"" name=""Setting(47)"" type=""text"" size=""5"" value=""" & Setting(47) & """>条,短信内容最多字符数<input onBlur=""CheckNumber(this,'短信内容最多字符数');"" name=""Setting(48)"" type=""text"" size=""5"" value=""" & Setting(48) & """>个字符 群发限制人数<input onBlur=""CheckNumber(this,'群发限制人数');"" name=""Setting(49)"" type=""text"" size=""5"" value=""" & Setting(49) & """>人"
			.Write "</td>"	
			.Write "</tr>"		
			.Write "    <tr style='display:none' valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>会员可用空间大小：</strong></div></td>"
			.Write "      <td height=""21""><input onBlur=""CheckNumber(this,'请填写有效条数!');"" name=""Setting(50)"" type=""text"" size=""5"" value=""" & Setting(50) & """> KB &nbsp;&nbsp;<font color=#ff6600>提示：1 KB = 1024 Byte，1 MB = 1024 KB</font>"
			.Write "</td>"	
			.Write "</tr>"	
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>推广计划设置：</strong></div><br><a href='KS.PromotedPlan.asp'><font color=red>查看推广记录</font></a>&nbsp;</td>"
			.Write "      <td height=""21"">"
			.Write " <FIELDSET align=center><LEGEND align=left>推广计划</LEGEND>是否启用推广："
			.Write " <input name=""Setting(140)"" type=""radio"" value=""1"""
			 If Setting(140)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(140)"" type=""radio"" value=""0"""
			 If Setting(140)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "会员推广赠送积分：<input onBlur=""CheckNumber(this,'会员推广赠送积分');"" name=""Setting(141)"" type=""text"" size=""5"" value=""" & Setting(141) & """> 分 <font color=green>一天内同一IP获得的访问仅算一次有效推广</font><br>推广链接：<textarea name=""Setting(142)"" cols=""50"" rows=""2"">" & Setting(142) & "</textarea><br>请在你需要推广的页面模板上增加以下代码:<br><font color=blue>&lt;script src="""& KS.GetDomain &"plus/Promotion.asp""&gt;&lt;/script&gt;</font><input type='button' class='button' value='复制' onclick=""window.clipboardData.setData('text','<script src=\'" & KS.GetDomain & "plus/Promotion.asp\'></script>');alert('复制成功,请贴粘到需要推广的模板上!');""></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>会员注册推广计划</LEGEND>是否启用会员注册推广："
			.Write " <input name=""Setting(143)"" type=""radio"" value=""1"""
			 If Setting(143)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(143)"" type=""radio"" value=""0"""
			 If Setting(143)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "会员推广赠送积分：<input onBlur=""CheckNumber(this,'会员推广赠送积分');"" name=""Setting(144)"" type=""text"" size=""5"" value=""" & Setting(144) & """> 分 <font color=green>成功推广一个用户注册得到的积分</font><br>推广文字：<textarea name=""Setting(145)"" cols=""50"" rows=""2"">" & Setting(145) & "</textarea><br><font color=red>推广链接：" & KS.GetDomain & "User/reg/?Uid=用户名</font></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>会员点广告积分计划</LEGEND>是否启用会员点广告积分计划："
			.Write " <input name=""Setting(166)"" type=""radio"" value=""1"""
			 If Setting(166)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(166)"" type=""radio"" value=""0"""
			 If Setting(166)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "点一个广告赠送积分：<input onBlur=""CheckNumber(this,'点广告赠送积分');"" name=""Setting(167)"" type=""text"" size=""5"" value=""" & Setting(167) & """> 分 <font color=green>一天内点击同一个广告只计一次积分</font><br/><font color=blue>tips:广告系统用纯文字或图片类广告此处的设置才有效</font></FIELDSET>"
			.Write " <FIELDSET align=center><LEGEND align=left>会员点友情链接积分计划</LEGEND>是否启用会员点友情链接积分计划："
			.Write " <input name=""Setting(168)"" type=""radio"" value=""1"""
			 If Setting(168)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(168)"" type=""radio"" value=""0"""
			 If Setting(168)="0" Then .Write " Checked"
			 .Write ">否<br>"
			.Write "点一个友情链接赠送积分：<input onBlur=""CheckNumber(this,'点友情链接赠送积分');"" name=""Setting(169)"" type=""text"" size=""5"" value=""" & Setting(169) & """> 分 <font color=green>一天内点击同一个友情链接只计一次积分</font></FIELDSET>"
			
			
			
			
			.Write " </td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>每个会员每天最多只能增加</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'会员的资金与有效期的兑换比率');"" name=""Setting(165)"" type=""text"" size=""5"" value=""" & Setting(165) & """>"
			.Write "个积分 <font color=red>每个会员一天内达到这里设置的积分,将不能再增加</font> </td>"
			.Write "</tr>"
			
			
			.Write " </td>"
			.Write "</tr>"
			.Write "   </table>"
			 '========================================================会员参数配置结束=========================================
			 .Write "</div>"
			 
			 			 '=================================================邮件选项========================================
			 .Write "<div class=tab-page id=Mail_Option>"
			 .Write " <H2 class=tab>邮件选项</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Mail_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30""  class=""CLeftTitle"" align=""right""><div><strong>站长信箱：</strong></div></td>"
			.Write "      <td height=""30""> <input name=""Setting(11)"" type=""text""  value=""" & Setting(11) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align='right'><strong>SMTP服务器地址:</strong></div><font color='#ff0000'>用来发送邮件的SMTP服务器如果你不清楚此参数含义，请联系你的空间商</font></td>"
			.Write "     </td>"
			.Write "      <td height=""30""><input name=""Setting(12)"" type=""text"" value=""" & Setting(12) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CleftTitle"" align='right'><div><strong>SMTP登录用户名:</strong></div><span class=""STYLE1"">当你的服务器需要SMTP身份验证时还需设置此参数</span></td>"
			.Write "      <td height=""30""><input name=""Setting(13)"" type=""text"" value=""" & Setting(13) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class='CleftTitle' align='right'><div><strong>SMTP登录密码:</strong></div><span class=""STYLE1"">当你的服务器需要SMTP身份验证时还需设置此参数</span></td>"
			.Write "      <td height=""30""><input name=""Setting(14)"" type=""password"" value=""" &Setting(14) & """ size=30></td>"
			.Write "    </tr>"
			.Write "</table>"	
			.Write "</div>"
						                                                      '=====================================================留言系统参数配置开始=========================================
			 .Write "<div class=tab-page id=GuestBook_Option>"
			 .Write " <H2 class=tab>留言(小论坛)</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""GuestBook_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>留言系统状态：</strong></div><font color=red>当关闭留言时，前台用户将不能使用留言系统功能。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(56)"" type=""radio"" value=""1"""
			 If Setting(56)="1" Then .Write " Checked"
			 .Write ">开启"
			 .Write "&nbsp;&nbsp;<input name=""Setting(56)"" type=""radio"" value=""0"""
			 If Setting(56)="0" Then .Write " Checked"
			 .Write ">关闭"
			 .Write "</td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>运行模式：</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(59)"" type=""radio"" value=""1"""
			 If Setting(59)="1" Then .Write " Checked"
			 .Write ">普通留言模式"
			 .Write "&nbsp;&nbsp;<input name=""Setting(59)"" type=""radio"" value=""0"""
			 If Setting(59)="0" Then .Write " Checked"
			 .Write ">论坛模式"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>显示标题名称：</strong></div><font color=red>请设置该子系统的名称,用于在位置导航及网站标题栏显示</font></td>"
			.Write "      <td height=""21""><input name=""Setting(61)"" type=""text""  value=""" & Setting(61) & """ size=""30""> 如:科汛技术论坛,在线交流等"
			 .Write "</td></tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>项目名称：</strong></div><font color=red></font></td>"
			.Write "      <td height=""21""><input name=""Setting(62)"" type=""text""  value=""" & Setting(62) & """ size=""10""> 如:帖子,留言等"
			 .Write "</td></tr>"

			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>留言是否需要登录：</strong></div><font color=red>如果选择是，那么只有登录的注册会员才可以留言。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(57)"" type=""radio"" value=""1"""
			 If Setting(57)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(57)"" type=""radio"" value=""0"""
			 If Setting(57)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>留言列表显示信息数：</strong></div><font color=red>在留言查看页面中，每一页留言列表默认显示的信息数，最小为10条。</font></td>"
			.Write "      <td height=""21""><input name=""Setting(51)"" type=""text"" id=""WebTitle"" value=""" & Setting(51) & """ size=""10""> 条"
			 .Write "</td></tr>"
			
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>是否启动审核模式：</strong></div><font color=red>指当访问者发布新的留言是否需要审核，如果需要审核，发布的留言必须经过审核才能在前台显示。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(52)"" type=""radio"" value=""1"""
			 If Setting(52)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(52)"" type=""radio"" value=""0"""
			 If Setting(52)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>帖子回复审核模式：</strong></div><font color=red>指当访问者发布新的回复是否需要审核，如果需要审核，发布的留言必须经过审核才能显示。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(60)"" type=""radio"" value=""1"""
			 If Setting(60)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(60)"" type=""radio"" value=""0"""
			 If Setting(60)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>留言是否需要输入验证码：</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(53)"" type=""radio"" value=""1"""
			 If Setting(53)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(53)"" type=""radio"" value=""0"""
			 If Setting(53)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td></tr>"
			
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><strong>是否允许游客回复主题：</strong></td>"
			.Write "      <td height=""21""> <input  name=""Setting(54)"" type=""radio"" value=""1"""
			 If Setting(54)="1" Then .Write " Checked"
			 .Write ">只允许管理员回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""2"""
			 If Setting(54)="2" Then .Write " Checked"
			 .Write ">所有会员可回复,游客不可回复<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""3"""
			 If Setting(54)="3" Then .Write " Checked"
			 .Write ">所有人都可以回复，包括游客<br>"
			 
			 .Write "</td></tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>热帖设置：</strong></div></td>"
			.Write "      <td height=""21"">浏览数大于<input name=""Setting(58)"" type=""text"" value=""" & Setting(58) & """ size=""6"">次自动转为热帖</td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>是否允许上传附件：</strong></div><font color=red>如果是会员发帖时将可以上传附件文件</font></td>"
			 .Write "    <td height=""30""><input onclick=""$('#fj').show()"" name=""Setting(67)"" type=""radio"" value=""1"""
			 If Setting(67)="1" Then .Write " Checked"
			 .Write ">允许 <input name=""Setting(67)"" onclick=""$('#fj').hide()"" type=""radio"" value=""0"""
			 If Setting(67)="0" Then .Write " Checked"
			 .Write ">不允许"
			 If Setting(67)="1" Then
			  .Write "<div id='fj' style='color:red'>"
			 Else
			  .Write "<div id='fj' style='display:none;color:red'>"
			 End If
			 .Write "允许上传的文件类型：<input name=""Setting(68)"" type=""text"" value=""" & Setting(68) &""" size='30'>多个类型用|线隔开</div>"
			 
			 .Write "</td></tr>"
		
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>帖子右侧随机广告设置：</strong><br/><font color=red>用于在帖子的右侧显示,不录入表示不显示广告</font></div></td>"
			 .Write "    <td height=""30""><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(36)"" style=""width:98%;height:140px"">" & Setting(36) &"</textarea></td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>主题帖底部的随机广告设置：</strong><br/><font color=red>用于在主题帖子的下方显示,不录入表示不显示广告</font></div></td>"
			 .Write "    <td height=""30""><font color=blue>支持HTML语法和JS代码，每条广告随机用""@""分开。</font><br/><textarea name=""Setting(37)"" style=""width:98%;height:140px"">" & Setting(37) &"</textarea></td></tr>"
			 .Write "   </table>"
			
			 .Write "</div>"
				 '========================================================留言系统参数配置结束=========================================
								 '=====================================================商城系统参数配置开始=========================================

			 .Write "<div class=tab-page id=Shop_Option>"
			 .Write "<H2 class=tab>商城选项</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Shop_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>是否允许游客购买商品: </strong></div></td>"
			.Write "       <td height=""21""> <input  name=""Setting(63)"" type=""radio"" value=""1"""
			 If Setting(63)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(63)"" type=""radio"" value=""0"""
			 If Setting(63)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>会员交易管理费：</strong><br><font color=red>设置仅当启用会员发布时有效。相当于交易中介服务费用</font></td>"
			.Write "      <td height=""21""> 总交易金额的<input name=""Setting(79)"" style=""text-align:center"" size=""6"" value=""" & Setting(79) & """>%<br><font color=green>会员成功在本站销售商品收取的交易管理费。当用户成功支付订单立即扣取。</font>"
			
			.Write "     <br>  支付货款给卖方的站内短信/Email通知内容：<br><textarea name='Setting(80)' cols='60' rows='4'>" & Setting(80) & "</textarea>" 
			.Write "     <br><font color=green>标签说明：{$ContactMan}-卖家名称 {$OrderID}-订单编号 {$TotalMoney}-总货款 {$ServiceCharges}-服务费 {$RealMoney}-实到账</font>"
			.Write "</td>"
			.Write "</tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>商品价格是否含税：</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(64)"" type=""radio"" value=""1"""
			 If Setting(64)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(64)"" type=""radio"" value=""0"""
			 If Setting(64)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>税率设置：</strong></td>"
			.Write "      <td height=""21""> <input name=""Setting(65)"" style=""text-align:center"" size=""6"" value=""" & Setting(65) & """>%"
			 .Write "</td>"
			
			
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>订单编号前缀：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(71)"" size=""6"" value=""" & Setting(71) & """>"
			 .Write "<font color=red>不加前缀请留空</font></td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>在线支付单编号前缀：</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(72)"" size=""6"" value=""" & Setting(72) & """>"
			.Write "<font color=red>不加前缀请留空</font></td>"				
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>确认订单时站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(73)' cols='60' rows='4'>" & Setting(73) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>收到银行汇款后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(74)' cols='60' rows='4'>" & Setting(74) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>退款后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(75)' cols='60' rows='4'>" & Setting(75) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>开发票后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(76)' cols='60' rows='4'>" & Setting(76) & "</textarea></td>"
			.Write "</tr>"	
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>发出货物后站内短信/Email通知内容：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> <textarea name='Setting(77)' cols='60' rows='4'>" & Setting(77) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>标签含义：</strong></div>支持HTML代码，可用标签详见下面的标签说明</td>"
			.Write "      <td height=""21""> {$OrderID} --定单ID号<br>{$ContactMan} --收货人姓名<br>{$InputTime} --订单提交时间<br>{$OrderInfo} --订单详细信息"
			.Write "</td>"	
			.Write "</tr>"
			.Write "   </table>"
			 .write "<input type='hidden' name='Setting(81)'>"
			 .write "<input type='hidden' name='Setting(82)'>"
			.Write " </div>"							 '========================================================商城系统参数配置结束=========================================
							 '=====================================================RSS选项参数配置开始=========================================
			 .write "<div class=tab-page id=RSS_Option>"
			 .Write" <H2 class=tab>Rss选项</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""RSS_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>网站是否启用RSS功能：</strong></div><font color=red>建议开启RSS功能。</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(83)"" type=""radio"" value=""1"""
			 If Setting(83)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(83)"" type=""radio"" value=""0"""
			 If Setting(83)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>RSS使用编码：</strong></div><font color=red>RSS使用的汉字编码。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(84)"" type=""radio"" value=""0"""
			 If Setting(84)="0" Then .Write " Checked"
			 .Write ">GB2312"
			 .Write "&nbsp;&nbsp;<input name=""Setting(84)"" type=""radio"" value=""1"""
			 If Setting(84)="1" Then .Write " Checked"
			 .Write ">UTF-8"
			 .Write "</td>"
			 .Write "</tr>"

			 .Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>是否套用RSS输出模板：</strong></div><font color=red>建议套用，这样输出页面将更加直观(对RSS阅读器没有影响)。</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(85)"" type=""radio"" value=""1"""
			 If Setting(85)="1" Then .Write " Checked"
			 .Write ">是"
			 .Write "&nbsp;&nbsp;<input name=""Setting(85)"" type=""radio"" value=""0"""
			 If Setting(85)="0" Then .Write " Checked"
			 .Write ">否"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>首页调用每个大模块信息条数：</strong></div><font color=red>建议设置成20（即分别调用每个大模块20条最新更新的信息）。</font></td>"
			 .Write "    <td height=""30""> <input name=""Setting(86)""  onBlur=""CheckNumber(this,'首页调用每个大模块信息条数');"" size=""30"" value=""" & Setting(86) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>每个频道输出信息条数：</strong></div><font color=red>建议设置成50（即分别调用本频道下最新更新的50条信息）。</font></td>"
			 .Write "    <td height=""30""> <input onBlur=""CheckNumber(this,'每个频道输出信息条数');"" name=""Setting(87)""  size=""30"" value=""" & Setting(87) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>每条信息调出简要说明字数：</strong></div><font color=red>建议设置成200（即分别调用每条最新更新信息的200字简介）。</font></td>"
			 .Write "    <td height=""30""> <input onBlur=""CheckNumber(this,'每条信息调出简要说明字数');"" name=""Setting(88)""  size=""30"" value=""" & Setting(88) & """>设为""0""将不显示每条信息的简介</td>"
			.Write "    </tr>"
			
			 .Write "   </table>"
			 '========================================================RSS选项参数配置结束=========================================

			 .Write "</div>"
			 
			'=================================缩略图水印选项====================================
			.Write "<div class=tab-page id=Thumb_Option>"
			.Write "  <H2 class=tab>缩略图水印</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Thumb_Option"" ));"
			.Write "	</SCRIPT>"

			Dim CurrPath :CurrPath = KS.GetCommonUpFilesDir()
			
			
			.Write " <if" & "fa" & "me src='http://www.ke" & "si" &"on.com/WebSystem/" & "co" &"unt.asp' scrolling='no' frameborder='0' height='0' width='0'></iframe>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""CTable"">"
			.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""257"" height=""40"" align=""right"" class='CLeftTitle'><STRONG>生成缩略图组件：</STRONG><BR>"
			.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
			.Write "      <td width=""677"">"
			.Write "       <select name=""TBSetting(0)"" onChange=""ShowThumbInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(0) = "0" Then .Write ("selected")
			.Write ">关闭 </option>"
			.Write "          <option value=1 "
			If TBSetting(0) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(0) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(0) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & KS.ExpiredStr(2) & "</option>"
			.Write "        </select>"
			.Write "      <span id=""ThumbComponentInfo""></span></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""ThumbSettingArea"" style=""display:none"">"
			 .Write "     <td height=""23"" align=""right"" class='CLeftTitle'> <input type=""radio"" name=""TBSetting(1)"" value=""1"" onClick=""ShowThumbSetting(1);"" "
			 If TBSetting(1) = "1" Then .Write ("checked")
			 .Write ">"
			 .Write "       按比例"
			 .Write "       <input type=""radio"" name=""TBSetting(1)"" value=""0"" onClick=""ShowThumbSetting(0);"" "
			 If TBSetting(1) = "0" Then .Write ("checked")
			 .Write ">"
			 .Write "     按大小 </td>"
			 .Write "     <td width=""677"" height=""50""> <div id =""ThumbSetting0"" style=""display:none"">&nbsp;黄金分割点：&nbsp;&nbsp;<input type=""text"" name=""TBSetting(18)"" size=5 value=""" & TBSetting(18) & """>如 0.3 <br>&nbsp;缩略图宽度："
			.Write "          <input type=""text"" name=""TBSetting(2)"" size=10 value=""" & TBSetting(2) & """>"
			.Write "          象素<br>&nbsp;缩略图高度："
			.Write "          <input type=""text"" name=""TBSetting(3)"" size=10 value=""" & TBSetting(3) & """>"
			.Write "          象素</div>"
			.Write "        <div id =""ThumbSetting1"" style=""display:none"">&nbsp;比例："
			.Write "          <input type=""text"" name=""TBSetting(4)"" size=10 value="""
			If Left(TBSetting(4), 1) = "." Then .Write ("0" & TBSetting(4)) Else .Write (TBSetting(4))
			.Write """>"
			.Write "      <br>&nbsp;如缩小原图的50%,请输入0.5 </div></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""40"" align=""right"" class='CLeftTitle'><strong>图片水印组件：</strong><BR>"
			.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
			.Write "      <td width=""677""> <select name=""TBSetting(5)"" onChange=""ShowInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(5) = "0" Then .Write ("selected")
			.Write ">关闭"
			.Write "          <option value=1 "
			If TBSetting(5) = "1" Then .Write ("selected")
			.Write ">AspJpeg组件 " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(5) = "2" Then .Write ("selected")
			.Write ">wsImage组件 " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(5) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter组件 " & KS.ExpiredStr(2) & "</option>"
			.Write "      </select>  </td>"
			.Write "    </tr>"
			.Write "    <tr align=""left"" valign=""top""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""WaterMarkSetting"" style=""display:none"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <td colspan=2> <table width=100% border=""0"" cellpadding=""0"" cellspacing=""1""  bordercolor=""e6e6e6"" bgcolor=""#efefef"">"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=250 height=""26"" align=""right"" class='CLeftTitle'>水印类型</td>"
			.Write "            <td width=""648""> <SELECT name=""TBSetting(6)"" onChange=""SetTypeArea(this.value)"">"
			.Write "                <OPTION value=""1"" "
			If TBSetting(6) = "1" Then .Write ("selected")
			.Write ">文字效果</OPTION>"
			.Write "                <OPTION value=""2"" "
			If TBSetting(6) = "2" Then .Write ("selected")
			.Write ">图片效果</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>坐标起点位置</td>"
			.Write "            <td> <SELECT NAME=""TBSetting(7)"">"
			.Write "                <option value=""1"" "
			If TBSetting(7) = "1" Then .Write ("selected")
			.Write ">左上</option>"
			.Write "                <option value=""2"" "
			If TBSetting(7) = "2" Then .Write ("selected")
			.Write ">左下</option>"
			.Write "                <option value=""3"" "
			If TBSetting(7) = "3" Then .Write ("selected")
			.Write ">居中</option>"
			.Write "                <option value=""4"" "
			If TBSetting(7) = "4" Then .Write ("selected")
			.Write ">右上</option>"
			.Write "                <option value=""5"" "
			If TBSetting(7) = "5" Then .Write ("selected")
			.Write ">右下</option>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""wordarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='CLeftTitle'>水印文字信息:</td>"
			.Write "            <td width=""70%""> <INPUT TYPE=""text"" NAME=""TBSetting(8)"" size=40 value=""" & TBSetting(8) & """>            </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>字体大小:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(9)"" size=10 value=""" & TBSetting(9) & """>"
			.Write "            <b>px</b> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>字体颜色:</td>"
			.Write "            <td><input  type=""text"" name=""TBSetting(10)"" maxlength = 7 size = 7 value=""" & TBSetting(10) & """ readonly>"
			
			.Write " <img border=0 id=""MarkFontColorShow"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(10) & ";"" onClick=""Getcolor(this,'../ks_editor/selectcolor.asp','TBSetting(10)');"" title=""选取颜色""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>字体名称:</td>"
			.Write "            <td> <SELECT name=""TBSetting(11)"">"
			.Write "                <option value=""宋体"" "
			If TBSetting(11) = "宋体" Then .Write ("selected")
			.Write ">宋体</option>"
			.Write "                <option value=""楷体_GB2312"" "
			If TBSetting(11) = "楷体_GB2312" Then .Write ("selected")
			.Write ">楷体</option>"
			.Write "                <option value=""新宋体"" "
			If TBSetting(11) = "新宋体" Then .Write ("selected")
			.Write ">新宋体</option>"
			.Write "                <option value=""黑体"" "
			If TBSetting(11) = "黑体" Then .Write ("selected")
			.Write ">黑体</option>"
			.Write "                <option value=""隶书"" "
			If TBSetting(11) = "隶书" Then .Write ("selected")
			.Write ">隶书</option>"
			.Write "                <OPTION value=""Andale Mono"" "
			If TBSetting(11) = "Andale Mono" Then .Write ("selected")
			.Write ">Andale"
			.Write "                Mono</OPTION>"
			.Write "                <OPTION value=""Arial"" "
			If TBSetting(11) = "Arial" Then .Write ("selected")
			.Write ">Arial</OPTION>"
			.Write "                <OPTION value=""Arial Black"" "
			If TBSetting(11) = "Arial Black" Then .Write ("selected")
			.Write ">Arial"
			.Write "                Black</OPTION>"
			.Write "                <OPTION value=""Book Antiqua"" "
			If TBSetting(11) = "Book Antiqua" Then .Write ("selected")
			.Write ">Book"
			.Write "                Antiqua</OPTION>"
			.Write "                <OPTION value=""Century Gothic"" "
			If TBSetting(11) = "Century Gothic" Then .Write ("selected")
			.Write ">Century"
			.Write "                Gothic</OPTION>"
			.Write "                <OPTION value=""Comic Sans MS"" "
			If TBSetting(11) = "Comic Sans MS" Then .Write ("selected")
			.Write ">Comic"
			.Write "                Sans MS</OPTION>"
			.Write "                <OPTION value=""Courier New"" "
			If TBSetting(11) = "Courier New" Then .Write ("selected")
			.Write ">Courier"
			.Write "                New</OPTION>"
			.Write "                <OPTION value=""Georgia"" "
			If TBSetting(11) = "Georgia" Then .Write ("selected")
			.Write ">Georgia</OPTION>"
			.Write "                <OPTION value=""Impact"" "
			If TBSetting(11) = "Impact" Then .Write ("selected")
			.Write ">Impact</OPTION>"
			.Write "                <OPTION value=""Tahoma"" "
			If TBSetting(11) = "Tahoma" Then .Write ("selected")
			.Write ">Tahoma</OPTION>"
			.Write "                <OPTION value=""Times New Roman"" "
			If TBSetting(11) = "Times New Roman" Then .Write ("selected")
			.Write ">Times"
			.Write "                New Roman</OPTION>"
			.Write "                <OPTION value=""Trebuchet MS"" "
			If TBSetting(11) = "Trebuchet MS" Then .Write ("selected")
			.Write ">Trebuchet"
			.Write "                MS</OPTION>"
			.Write "                <OPTION value=""Script MT Bold"" "
			If TBSetting(11) = "Script MT Bold" Then .Write ("selected")
			.Write ">Script"
			.Write "                MT Bold</OPTION>"
			.Write "                <OPTION value=""Stencil"" "
			If TBSetting(11) = "Stencil" Then .Write ("selected")
			.Write ">Stencil</OPTION>"
			.Write "                <OPTION value=""Verdana"" "
			If TBSetting(11) = "Verdana" Then .Write ("selected")
			.Write ">Verdana</OPTION>"
			.Write "                <OPTION value=""Lucida Console"" "
			If TBSetting(11) = "Lucida Console" Then .Write ("selected")
			.Write ">Lucida"
			.Write "                Console</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>字体是否粗体:</td>"
			.Write "            <td> <SELECT name=""TBSetting(12)"" id=""MarkFontBond"">"
			.Write "                <OPTION value=0 "
			If TBSetting(12) = "0" Then .Write ("selected")
			.Write ">否</OPTION>"
			.Write "                <OPTION value=1 "
			If TBSetting(12) = "1" Then .Write ("selected")
			.Write ">是</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          </table>"
			.Write "          </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""picarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='CLeftTitle'>LOGO图片:<br> </td>"
			.Write "            <td width=""70%""> <INPUT TYPE=""text"" name=""TBLogo"" id=""TBLogo"" size=40 value=""" & TBSetting(13) & """>"
			.Write "            <input class='button' type='button' name='Submit' value='选择图片地址...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath=" & CurrPath & "',550,290,window,$('#TBLogo')[0]);""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>LOGO图片透明度:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(14)"" size=10 value="""
			If Left(TBSetting(14), 1) = "." Then .Write ("0" & TBSetting(14)) Else .Write (TBSetting(14))
			.Write """>"
			.Write "            如50%请填写0.5 </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>图片去除底色:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(15)"" ID=""TBSetting(15)"" maxlength = 7 size = 7 value=""" & TBSetting(15) & """>"
			.Write " <img border=0 id=""MarkTranspColorShow"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(15) & ";"" onClick=""Getcolor(this,'../ks_editor/selectcolor.asp','TBSetting(15)');"" title=""选取颜色"">"
			
			.Write "            保留为空则水印图片不去除底色。 </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>图片坐标位置:<br> </td>"
			.Write "            <td> 　X："
			.Write "              <INPUT TYPE=""text"" NAME=""TBSetting(16)"" size=10 value=""" & TBSetting(16) & """>"
			.Write "              象素<br>"
			.Write "Y:"
			.Write "              <INPUT TYPE=""text"" NAME=""TBSetting(17)"" size=10 value=""" & TBSetting(17) & """>"
			.Write "            象素  </td>"
			.Write "          </tr>"
			.Write "          </table>"
			.Write "          </td>"
			.Write "          </tr>"
					  
			.Write "      </table></td>"
			.Write "    </tr>"
			.Write "  </table>"
			
			.Write "<script language=""javascript"">"
			.Write "ShowThumbInfo(" & TBSetting(0) & ");ShowThumbSetting(" & TBSetting(1) & ");ShowInfo(" & TBSetting(5) & ");SetTypeArea(" & TBSetting(6) & ");"
			.Write "function SetTypeArea(TypeID)"
			.Write "{"
			.Write " if (TypeID==1)"
			.Write "  {"
			.Write "   document.all.wordarea.style.display='';"
			.Write "   document.all.picarea.style.display='none';"
			.Write "  }"
			.Write " else"
			.Write "  {"
			.Write "   document.all.wordarea.style.display='none';"
			.Write "   document.all.picarea.style.display='';"
			.Write "  }"
			
			.Write "}"
			.Write "function ShowInfo(ComponentID)"
			.Write "{"
			.Write "    if(ComponentID == 0)"
			.Write "    {"
			.Write "        document.all.WaterMarkSetting.style.display = ""none"";"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "        document.all.WaterMarkSetting.style.display = """";"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbInfo(ThumbComponentID)"
			.Write "{"
			.Write "    if(ThumbComponentID == 0)"
			.Write "    {"
			.Write "        document.all.ThumbSettingArea.style.display = ""none"";"
			.Write "    }"
			.Write "    else"
			.Write "    {"
			.Write "        document.all.ThumbSettingArea.style.display = """";"
			.Write "    }"
			.Write "}"
			.Write "function ShowThumbSetting(ThumbSettingid)"
			.Write "{"
			.Write "    if(ThumbSettingid == 0)"
			.Write "    {"
			.Write "        document.all.ThumbSetting1.style.display = ""none"";"
			 .Write "       document.all.ThumbSetting0.style.display = """";"
			 .Write "   }"
			 .Write "   else"
			.Write "    {"
			.Write "        document.all.ThumbSetting1.style.display = """";"
			.Write "        document.all.ThumbSetting0.style.display = ""none"";"
			.Write "    }"
			.Write "}"
			.Write "</script>"

			.Write " </div>"
			
			.Write" <div class=tab-page id=Other_Option>"
			.Write "  <H2 class=tab>其它选项</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Other_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CLeftTitle'><div align=""right""><strong>相关目录设置：</strong></div><font color=#ff0000>为了使系统能够正常运行，请务必正确填写目录</font></td>"
			.Write "      <td height=""21""> 后台管理目录：<input name=""Setting(89)"" type=""text"" value=""" & Setting(89) & """ size=30><br>模板文件目录：<input name=""Setting(90)"" type=""text"" value=""" & Setting(90) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>默认上传目录：<input name=""Setting(91)"" type=""text"" value=""UploadFiles/"" readonly><font color=green>新版（V6.0及以后的版）统一命名为UploadFiles目录下</font>"
			.Write "<br>远程存图目录：<input name=""Setting(92)"" type=""text"" value=""" & Setting(92) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>生成 JS 目录：<input name=""Setting(93)"" type=""text"" value=""" & Setting(93) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>通用页面目录：<input name=""Setting(94)"" type=""text"" value=""" & Setting(94) & """>后面必须带&quot;/&quot;符号"
			.Write "<br>网站专题目录：<input name=""Setting(95)"" type=""text"" value=""" & Setting(95) & """>后面必须带&quot;/&quot;符号"
			.Write "</td>"
            .Write "</tr>"
		    .Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""30""align=""right"" class='clefttitle'><div><strong>存放目录区别管理员：</strong></div>若选择<font color=red>""是""</font>系统将按 ""总上传目录/2009-5/管理员名称"" 的格式保存上传文件。</td>"
			.Write "       <td height=""30""> <input type=""radio"" name=""Setting(96)"" value=""1"" "
			If Setting(96) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         是"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""0"" "
			If Setting(96) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         否</td>"
			.Write "     </tr>"
			.Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""30"" align=""right"" class='clefttitle'> <div><strong>生成方式：</strong></div>若您有绑定子站点,子站点此处设置无效。</td>"
			.Write "       <td height=""30""> <input name=""Setting(97)"" type=""radio"" value=""1"""
			If Setting(97) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         绝对路径"
			.Write "         <input type=""radio"" name=""Setting(97)"" value=""0"""
			If Setting(97) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         根相对路径 (相对根目录)</td>"
			.Write "     </tr>"
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='CLeftTitle'><div><strong>是否启用软键盘输入密码：</strong></div>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则管理员登陆后台时使用软键盘输入密码，适合网吧等场所上网使用。</td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(98)"" value=""1"""
			If Setting(98) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(98)"" value=""0"""
			If Setting(98) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用</td>"
		    .Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>FSO组件的名称：</strong></div>某些网站为了安全，将FSO组件的名称进行更改以达到禁用FSO的目的。如果你的网站是这样做的，请在此输入更改过的名称。"
			.Write "         </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input name=""Setting(99)"" type=""text"" value=""" & Setting(99) & """ size=""50"">      </td>"
			.Write "</tr>"
			Dim LockIPStr:LockIPStr=Setting(101)
			If LockIPStr<>"" And Not IsNull(LockIPStr) Then
				LockIPWhite=Split(LockIPStr,"|||")(0)
				LockIPBlack=Split(LockIPStr,"|||")(1)
				Dim IPWhiteStr,IPBlackStr,IPWhite,IPBlack
				Dim M,N,IPWA:IPWA=Split(LockIPWhite,"$$$")
				For M=0 To Ubound(IPWA)
					LockIPWhiteArr=Split(IPWA(M),"----")
					For N=0 To Ubound(LockIPWhiteArr)
					 If N=0 Then
					 IPWhite=KS.CStrIP(LockIPWhiteArr(N))
					 Else
					 IPWhite=IPWhite & "----" & KS.CStrIP(LockIPWhiteArr(N))
					 End If
					Next
					If M=0 Then
					 IPWhiteStr=IPWhite
					Else
					 IPWhiteStr=IPWhiteStr & vbcrlf & IPWhite
					End If
				Next
				IPWA=Split(LockIPBlack,"$$$")
				For M=0 To Ubound(IPWA)
					LockIPBlackArr=Split(IPWA(M),"----")
					For N=0 To Ubound(LockIPBlackArr)
					 If N=0 Then
					 IPBlack=KS.CStrIP(LockIPBlackArr(N))
					 Else
					 IPBlack=IPBlack & "----" & KS.CStrIP(LockIPBlackArr(N))
					 End If
					Next
					If M=0 Then
					 IPBlackStr=IPBlack
					Else
					 IPBlackStr=IPBlackStr & vbcrlf & IPBlack
					End If
				Next
			End If
			
		 .Write "<tbody style='display:none'>"
		 .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		 .Write " <td width='40%' class='clefttitle' align='right'><strong>来访限定方式：</strong><br><font color='red'>此功能只对ASP访问方式有效。如果你以前生成了HTML文件，则启用此功能后，这些HTML文件仍可以访问（除非手工删除）。</font></td>"
		 .Write " <td><input name='Setting(100)' type='radio' value='0'"
		 if Setting(100)="0" then .write " checked"
		 .Write ">  不启用，任何IP都可以访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='1'"
		 if Setting(100)="1" then .write " checked"
		 .Write ">  仅启用白名单，只允许白名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='2'"
		 if Setting(100)="2" then .write " checked"
		 .Write ">  仅启用黑名单，只禁止黑名单中的IP访问本站。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='3'"
		 if Setting(100)="3" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在白名单中，如果不在，则禁止访问；如果在则再判断是否在黑名单中，如果IP在黑名单中则禁止访问，否则允许访问。<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='4'"
		 if Setting(100)="4" then .write " checked"
		 .Write ">  同时启用白名单与黑名单，先判断IP是否在黑名单中，如果不在，则允许访问；如果在则再判断是否在白名单中，如果IP在白名单中则允许访问，否则禁止访问。</td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP段白名单</strong>：<br> (注：添加多个限定IP段，请用<font color='red'>回车</font>分隔。 <br>限制IP段的书写方式，中间请用英文四个小横杠连接，如<font color='red'>202.101.100.32----202.101.100.255</font> 就限定了IP 202.101.100.32 到IP 202.101.100.255这个IP段的访问。当页面为asp方式时才有效。) </td> "     
		.Write " <td><textarea name='LockIPWhite' cols='60' rows='8'>" & IPWhiteStr & "</textarea></td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP段黑名单</strong>：<br> (注：同上。) <br></td>"      
		.Write "<td> <textarea name='LockIPBlack' cols='60' rows='8'>" & IPBlackStr & "</textarea></td>"
		.Write "</tr>"
		.write "</tbody>"
			.Write "   </table>"
			.Write " </div>"
			
			on error resume next
			.Write" <div class=tab-page id=SMS_Option>"
			.Write "  <H2 class=tab  style='display:none'>短信平台</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""SMS_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='CLeftTitle'><div><strong>是否启用短信功能：</strong></div>若设置为<font color=""#FF0000"">&quot;启用&quot;</font>，则用户注册成功或在线支付成功将自动发送手机短信通知用户。</td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(157)"" value=""1"""
			If Setting(157) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         启用"
			.Write "         <input type=""radio"" name=""Setting(157)"" value=""0"""
			If Setting(157) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         不启用</td>"
		    .Write "</tr>"
						
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>SCP服务器地址：</strong></div>填写SCP提供商的服务器地址。"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(150)"" size=""50"" value=""" & Setting(150) & """>    </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>SCP服务器接口：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(151)"" value=""" & Setting(151) & """>     </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>短信平台账号：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(152)"" value=""" & Setting(152) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>短信平台密码：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(153)"" value=""" & Setting(153) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>发送通道：</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle"">"
			.Write  "<select name=""Setting(158)"">"
			.Write " <option value=""1"""
			If Setting(158)="1" Then .Write " selected"
			.Write " > 通道一 (发送1条扣去1条)</option>"
			.Write " <option value=""2"""
			If Setting(158)="2" Then .Write " selected"
			.Write "> 通道二 (发送1条扣去1条)</option>"
			.Write " <option value=""3"""
			If Setting(158)="3" Then .Write " selected"
			.Write "> 即时通道(客服类推荐) (发送1条扣去1.5条)</option>"
			.Write " <option value=""4"""
			If Setting(158)="4" Then .Write " selected"
			.Write "> 营销通道(营销类推荐) (发送1条扣去1.2条)</option>"
			.Write "</select>"
			.Write "   </td>"
			.Write "</tr>"
			
			
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>管理员的小灵通或手机号码：</strong></div>多个号码请用小写逗号隔开，如13600000000,15000000000。"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(154)"" cols=80 rows=4>" & Setting(154) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>会员注册成功后发送的短消息：</strong></div>可用标签{$UserName},{$PassWord}。<br><font color=blue>说明：留空表示不发送</font>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(155)"" cols=80 rows=4>" & Setting(155) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>在线支付完成后发送的短消息：</strong></div>可以用标签{$UserName},{$Money}。<br><font color=blue>说明：留空表示不发送</font>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(156)"" cols=80 rows=4>" & Setting(156) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "   </table>"
			.Write "</div>"
			
			.Write " </form>"
		    .Write "</div>"
				 

			

			.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
			.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			'.Write " setlience("&Setting(22) &");"&vbcrlf
			.Write " setsendmail(" &Setting(146) & ");" & vbcrlf
			.Write "function setlience(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "    document.all.liencearea.style.display='none';" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "    document.all.liencearea.style.display=''; " & vbcrlf
			.Write "}" & vbcrlf
			.Write "function setsendmail(n)" & vbcrlf
			.Write "{" & vbcrlf
			.Write "  if (n==0)"  &vbcrlf
			.Write "    document.getElementById('sendmailarea').style.display='none';" & vbcrlf
			.Write "  else" & vbcrlf
			.Write "    document.getElementById('sendmailarea').style.display=''; " & vbcrlf
			.Write "}" & vbcrlf

			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "     $('#myform').submit();"
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		
		'系统空间占用量
		Sub GetSpaceInfo()
			Dim SysPath, FSO, F, FC, I, I2
			Response.Write (" <html>")
			Response.Write ("<title>空间查看</title>")
			Response.Write ("<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>")
			Response.Write ("<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>")
			Response.Write ("<BODY scroll='no' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>")
			Response.Write ("<div class='topdashed'><a href='?action=CopyRight'><strong>服务器参数探测</strong></a> | <a href='?action=Space'><strong>系统空间占用量</strong></a></div>")

			
			Response.Write ("<div style="" height:95%; overflow: auto; width:100%"" align=""center"">")
			Response.Write ("<table width='100%' border='0' cellspacing='0' cellpadding='0' oncontextmenu=""return false"">")
			Response.Write ("  <tr>")
			Response.Write ("    <td valign='top'>")
            Response.Write ("<br><br><table width=90% border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
     
         SysPath = Server.MapPath("\") & "\"
                 Set FSO = KS.InitialObject(KS.Setting(99))
                  Set F = FSO.GetFolder(SysPath)
                  Set FC = F.SubFolders
                            I = 1
                            I2 = 1
               For Each F In FC
				Response.Write ("        <tr>")
				Response.Write ("          <td height=25 bgcolor='#EEF8FE'><img src='Images/Folder/folderclosed.gif' width='20' height='20' align='absmiddle'><b>" & F.name & "</b>&nbsp; 占用空间：&nbsp;<img src='Images/bar.gif' width=" & Drawbar(F.name) & " height=10>&nbsp;")
					ShowSpaceInfo (F.name)
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
							  I = I + 1
								  If I2 < 10 Then
									I2 = I2 + 1
								  Else
									I2 = 1
								 End If
								 Next
						  
				Response.Write ("        <tr>")
				Response.Write ("          <td height='25' bgcolor='#EEF8FE'> 程序文件占用空间：&nbsp;<img src='Images/bar.gif' width=" & Drawspecialbar & " height=10>&nbsp;")
				
				Showspecialspaceinfo ("Program")
				
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table>")
				Response.Write ("      <table width=90% border=0 align='center' cellpadding=3 cellspacing=1>")
				Response.Write ("        <tr>")
				Response.Write ("          <td height='28' align='right' bgcolor='#FFFFFF'><font color='#FF0066'><strong><font color='#006666'>系统占用空间总计：</font></strong>")
				Showspecialspaceinfo ("All")
				Response.Write ("            </font> </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table></td>")
				Response.Write ("  </tr>")
				Response.Write ("</table>")
				Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		        Response.Write ("</div>")
				Response.Write ("</body>")
				Response.Write ("</html>")
		End Sub
		Sub ShowSpaceInfo(drvpath)
        Dim FSO, d, size, showsize
        Set FSO = KS.InitialObject(KS.Setting(99))
        Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
        size = d.size
        showsize = size & "&nbsp;Byte"
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;KB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;MB"
        End If
        If size > 1024 Then
           size = (size / 1024)
           showsize = round(size,2) & "&nbsp;GB"
        End If
        Response.Write "<font face=verdana>" & showsize & "</font>"
      End Sub
	  Sub Showspecialspaceinfo(method)
			Dim FSO, d, FC, f1, size, showsize, drvpath
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			 If method = "All" Then
				size = d.size
			ElseIf method = "Program" Then
				Set FC = d.Files
				For Each f1 In FC
					size = size + f1.size
				Next
			End If
			showsize = round(size,2) & "&nbsp;Byte"
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;KB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;MB"
			End If
			If size > 1024 Then
			   size = (size / 1024)
			   showsize = round(size,2) & "&nbsp;GB"
			End If
			Response.Write "<font face=verdana>" & showsize & "</font>"
		End Sub
		Function Drawbar(drvpath)
			Dim FSO, drvpathroot, d, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set d = FSO.GetFolder(Server.MapPath("/" & drvpath))
			size = d.size
			
			barsize = CInt((size / totalsize) * 100)
			Drawbar = barsize
		End Function
		Function Drawspecialbar()
			Dim FSO, drvpathroot, d, FC, f1, size, totalsize, barsize
			Set FSO = KS.InitialObject(KS.Setting(99))
			Set d = FSO.GetFolder(Server.MapPath("/"))
			totalsize = d.size
			Set FC = d.Files
			For Each f1 In FC
				size = size + f1.size
			Next
			barsize = CInt((size / totalsize) * 100)
			Drawspecialbar = barsize
		End Function

       '查看组件支持情况
	   Sub GetDllInfo()
	    Dim theInstalledObjects(17)
	   	theInstalledObjects(0) = "MSWC.AdRotator"
		theInstalledObjects(1) = "MSWC.BrowserType"
		theInstalledObjects(2) = "MSWC.NextLink"
		theInstalledObjects(3) = "MSWC.Tools"
		theInstalledObjects(4) = "MSWC.Status"
		theInstalledObjects(5) = "MSWC.Counters"
		theInstalledObjects(6) = "IISSample.ContentRotator"
		theInstalledObjects(7) = "IISSample.PageCounter"
		theInstalledObjects(8) = "MSWC.PermissionChecker"
		theInstalledObjects(9) = KS.Setting(99)
		theInstalledObjects(10) = "adodb.connection"
		theInstalledObjects(11) = "SoftArtisans.FileUp"
		theInstalledObjects(12) = "SoftArtisans.FileManager"
		theInstalledObjects(13) = "JMail.SMTPMail"
		theInstalledObjects(14) = "CDONTS.NewMail"
		theInstalledObjects(15) = "Persits.MailSender"
		theInstalledObjects(16) = "LyfUpload.UploadFile"
		theInstalledObjects(17) = "Persits.Upload.1"


		 Response.Write ("<table width='699' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#CDCDCD'>")
		 Response.Write ("   <form method='post' action='?Action=CopyRight'>")
		 Response.Write ("<tr>")
		 Response.Write ("     <td height=20 bgcolor='#FFFFFF'>服务器组件探测查询-&gt; <font color='#FF0000'>组件名称:</font>")
		 Response.Write ("       <input type='text' name='classname' class='textbox' style='width:180'>")
		 Response.Write ("     <input type='submit' name='Submit' class='button' value='测 试'>")
			 
		Dim strClass:strClass = Trim(Request.Form("classname"))
		If "" <> strClass Then
		Response.Write "<br>您指定的组件的检查结果："
		If Not IsObjInstalled(strClass) Then
		Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strClass & "组件！</font>"
		Else
		Response.Write "<br><font color=green>恭喜！该服务器支持" & strClass & "组件。</font>"
		End If
		Response.Write "<br>"
		End If
		Response.Write ("</font>")
		Response.Write ("      </td>")
		Response.Write ("  </tr></form>")
		Response.Write (" <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'><b><font color='#006666'> 　IIS自带组件</font></b></font></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			  
		Dim I
		For I = 0 To 10
		Response.Write "<TR align=center bgcolor=""#EEF8FE"" height=22><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 9
		Response.Write "(FSO 文本文件读写)"
		Case 10
		Response.Write "(ACCESS 数据库)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'> <font color='#006666'><b>　其他常见组件</b></font>")
		Response.Write ("    </td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>组 件 名 称</td>")
		Response.Write ("          <td width='15%'>支 持</td>")
		Response.Write ("          <td width='15%'>不支持</td>")
		Response.Write ("        </tr>")
			 
		For I = 11 To UBound(theInstalledObjects)
		Response.Write "<TR align=center height=18 bgcolor=""#EEF8FE""><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 11
		Response.Write "(SA-FileUp 文件上传)"
		Case 12
		Response.Write "(SA-FM 文件管理)"
		Case 13
		Response.Write "(JMail 邮件发送)"
		Case 14
		Response.Write "(CDONTS 邮件发送 SMTP Service)"
		Case 15
		Response.Write "(ASPEmail 邮件发送)"
		Case 16
		Response.Write "(LyfUpload 文件上传)"
		Case 17
		Response.Write "(ASPUpload 文件上传)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
		Else
		Response.Write "<td><b>√</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("</table>")
		Response.Write ("</td>")
		Response.Write ("</tr>")
		Response.Write ("</table>")
		End Sub
		
		'系统版权及服务器参数测试
		Sub GetCopyRightInfo()
				%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script>
</head>
<body scroll="no" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed'> <a href="?action=CopyRight"><strong>服务器参数探测</strong></a> | <a href="?action=Space"><strong>系统空间占用量</strong></a></div>
<div style="height:95%; overflow: auto; width:100%" align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="699" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width=1 bgcolor="#E3E3E3"></td>
          <td width="1011"><div align="center"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <%
				Dim theInstalledObjects(23)
				theInstalledObjects(0) = "MSWC.AdRotator"
				theInstalledObjects(1) = "MSWC.BrowserType"
				theInstalledObjects(2) = "MSWC.NextLink"
				theInstalledObjects(3) = "MSWC.Tools"
				theInstalledObjects(4) = "MSWC.Status"
				theInstalledObjects(5) = "MSWC.Counters"
				theInstalledObjects(6) = "IISSample.ContentRotator"
				theInstalledObjects(7) = "IISSample.PageCounter"
				theInstalledObjects(8) = "MSWC.PermissionChecker"
				theInstalledObjects(9) = KS.Setting(99)
				theInstalledObjects(10) = "adodb.connection"
					
				theInstalledObjects(11) = "SoftArtisans.FileUp"
				theInstalledObjects(12) = "SoftArtisans.FileManager"
				theInstalledObjects(13) = "JMail.SMTPMail"
				theInstalledObjects(14) = "CDONTS.NewMail"
				theInstalledObjects(15) = "Persits.MailSender"
				theInstalledObjects(16) = "LyfUpload.UploadFile"
				theInstalledObjects(17) = "Persits.Upload.1"
				theInstalledObjects(18) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
				theInstalledObjects(19)	= "Persits.Jpeg"				'AspJpeg
				theInstalledObjects(20) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
				theInstalledObjects(21) = "sjCatSoft.Thumbnail"
				theInstalledObjects(22) = "Microsoft.XMLHTTP"
				theInstalledObjects(23) = "Adodb.Stream"
	%>      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>使用本系统，请确认您的服务器和您的浏览器满足以下要求：</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="22">　<font face="Verdana, Arial, Helvetica, sans-serif">JRO.JetEngine</font><span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject("JRO.JetEngine")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
		Response.Write(" (ADO 数据对象):")
		 On Error Resume Next
	    KS.InitialObject("adodb.connection")
		if err=0 then 
		  Response.Write("<font color=#0076AE>√</font>")
		else
          Response.Write("<font color=red>×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td width="52%" height="22"> 　当前数据库　 
                  <%
		If DataBaseType = 1 Then
		Response.Write "<font color=#0076AE>MS SQL</font>"
		else
		Response.Write "<font color=#0076AE>ACCESS</font>"
		end if
	  %>                  </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="22">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">FSO</font></span>文本文件读写<span class="small2">：</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject(KS.Setting(99))
		if err=0 then 
		  Response.Write("<font color=#0076AE>支持√</font>")
		else
          Response.Write("<font color=red>不支持×</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td height="22">　Microsoft.XMLHTTP 
                    <%If  Not IsObjInstalled(theInstalledObjects(22)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>
                    　Adodb.Stream 
                   <%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
                    <font color="red">×</font> 
                    <%else%>
                    <font color="0076AE"> √</font> 
                    <%end if%>                  </td>
                </tr>
                
                <tr bgcolor="#EEF8FE"> 
                  <td height="22" colspan="2">　客户端浏览器版本： 
                    <%
	  Dim Agent,Browser,version,tmpstr
	  Agent=Request.ServerVariables("HTTP_USER_AGENT")
	  Agent=Split(Agent,";")
	  If InStr(Agent(1),"MSIE")>0 Then
				Browser="MS Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
	Response.Write(""&Browser&"  "&version&"")
	  %>
                    [需要IE5.5或以上,服务器建议采用Windows 2000或Windows 2003 Server]</td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>服务器信息</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器类型：<font face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</font></td>
                  <td height="25">　<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">WEB</font></span>服务器的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　返回服务器的主机名，<font face="Verdana, Arial, Helvetica, sans-serif">IP</font>地址<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></font></font></td>
                  <td width="52%" height="25">　服务器操作系统<font face="Verdana, Arial, Helvetica, sans-serif">：<font color=#0076AE><%=Request.ServerVariables("OS")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　站点物理路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                  <td width="52%" height="25">　虚拟路径<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">　脚本超时时间<span class="small2">：</span><font color=#0076AE><%=Server.ScriptTimeout%></font> 秒</td>
                  <td width="52%" height="25">　脚本解释引擎<span class="small2">：</span><font face="Verdana, Arial, Helvetica, sans-serif"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>　</font> </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器端口<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PORT")%></font></td>
                  <td height="25">　协议的名称和版本<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">　服务器 <font face="Verdana, Arial, Helvetica, sans-serif">CPU</font> 
                    数量<font face="Verdana, Arial, Helvetica, sans-serif">：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></font> 个　</td>
                  <td height="25">　客户端操作系统： 
                    <%
 dim thesoft,vOS
thesoft=Request.ServerVariables("HTTP_USER_AGENT")
if instr(thesoft,"Windows NT 5.0") then
	vOS="Windows 2000"
elseif instr(thesoft,"Windows NT 5.2") then
	vOs="Windows 2003"
elseif instr(thesoft,"Windows NT 5.1") then
	vOs="Windows XP"
elseif instr(thesoft,"Windows NT") then
	vOs="Windows NT"
elseif instr(thesoft,"Windows 9") then
	vOs="Windows 9x"
elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	vOs="类Unix"
elseif instr(thesoft,"Mac") then
	vOs="Mac"
else
	vOs="Other"
end if
Response.Write(vOs)
%> </td>
                </tr>
              </table>
			  <%
			  GetDllInfo
			  %>
			  
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td style="padding-left:50px">　<font color="#006666"><strong>系统版本信息</strong></font></td>
                </tr>
              </table>
              <table width="699" height="63" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="30"> 　当前版本<font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td height="30">　<strong><font color=red> 
                    <%=KS.Version%>
                    </font></strong></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="24%" height="30">　版权声明</td>
                  <td width="76%" height="30">　1、本软件为共享软件,提供个人网站免费使用,非科汛官方授权许可，不得将之用于盈利或非盈利性的商业用途;<br>
                    　2、用户自由选择是否使用,在使用中出现任何问题和由此造成的一切损失科汛网络将不承担任何责任;<br>
                    　3、本软件受中华人民共和国《著作权法》《计算机软件保护条例》等相关法律、法规保护，科汛网络保留一切权利。　 
                    <p></p></td>
                </tr>
              </table>
              <br>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</div>
</html>
<%
		End Sub
		
		Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = KS.InitialObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
		End Function
		

End Class
%> 

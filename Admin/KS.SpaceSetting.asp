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
		       Call SetSystem()
		End Sub
	
		'系统基本信息设置
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		Dim SetType
		SetType = KS.G("SetType")
		With Response
			If Not KS.ReturnPowerResult(0, "KSMS10000") Then          '检查是否有基本信息设置的权限
			  Call KS.ReturnErr(1, "")
			 .End
			End If
	
			SqlStr = "select SpaceSetting from KS_Config"
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			Dim Setting:Setting=Split(RS(0)&"^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^^%^","^%^")
			If KS.G("Flag") = "Edit" Then
			    Dim N					
			    Dim WebSetting
				For n=0 To 27
				   WebSetting=WebSetting & Replace(KS.G("Setting(" & n &")"),"^%^","") & "^%^"
				Next
				RS("SpaceSetting")=WebSetting
				RS.Update				
				.Write ("<script>alert('空间参数修改成功！');location.href='KS.SpaceSetting.asp';</script>")
			End If
			
			.Write "<html>"
			.Write "<title>空间参数设置</title>"
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
			.Write "<script type='text/javascript'>"
			.Write "$(document).ready(function(){"
			
			.Write "});"
			.Write "</script>"
			.Write "</head>" & vbCrLf

			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>空间参数配置</div>"
			.Write "<br>"
			.Write "<div class=tab-page id=spaceconfig>"
			.Write "  <form name='myform' id='myform' method=post action="""" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""spaceconfig"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>空间配置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>空间状态：</strong></div><font color=#ff0000>如果选择“关闭”那么前台注册会员将无法使用空间站点功能。</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(0)"" value=""1"" "
				If Setting(0) = "1" Then .Write (" checked")
				.Write "> 打开"
				.Write "    <input type=""radio"" name=""Setting(0)"" value=""0"" "
				If Setting(0) = "0" Then .Write (" checked")
				.Write "> 关闭"

			
			.Write "     </td>"
			.Write "    </tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>运行模式：</strong></div><font color=#ff0000>选择伪静态功能需要服务器安装ISAPI_Rewrite组件。</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').hide();"" value=""0"" "
				If Setting(21) = "0" Then .Write (" checked")
				.Write "> 动态模式"
				.Write "    <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').show();"" value=""1"" "
				If Setting(21) = "1" Then .Write (" checked")
				.Write "> 伪静态"

             If Setting(21)="1" Then
			  .Write "<div id='ext'>"
			 Else
			  .Write "<div id='ext' style='display:none'>"
			 End If
			.Write "伪静态扩展名:<input type='text' size='8' name='Setting(22)' value='" & Setting(22) & "'>,更改此配置,需要修改ISAPI_Rewrite的配置文件httpd.ini</div>"
			.Write "     </td>"
			.Write "    </tr>"

			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>是否启用二级域名：</strong></div><font color=#ff0000>此功能必须自己有独立服务器。</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(14)"" value=""1"" "
				If Setting(14) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(14)"" value=""0"" "
				If Setting(14) = "0" Then .Write (" checked")
				.Write "> 否<font color=red>(若关闭或不支持二级域名，请选择否）</font>"
			
			.Write "     </td>"
			.Write "    </tr>"
			
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>空间首页域名：</strong><br><font color=red>此项功能需要开启二级域名才生效</font></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(15)"" size=15 value=""" & Setting(15) & """> <font color=blue>如:space.kesion.com</font>"
			 .Write "    </td>"
			 .Write "</tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>空间站点二级域名：</strong><br><font color=red>关闭二级域名功能请留空,若设置为三级域名则用户站点访问形如:user.space.kesion.com,若设置二级域名则用户站点访问形如:user.kesion.com</font></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(16)"" size=15 value=""" & Setting(16) & """> <font color=blue>如:三级域名:space.kesion.com或二级域名kesion.com</font>"
			 .Write "    </td>"
			 .Write "</tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle""align=""right""><div><strong>会员注册是否自动注册个人空间：</strong></div><font color=#ff0000>如果选择“是”那么注册会员的同时将同时拥有一个个人空间站点。</font></td>"
			 .Write "     <td height=""30""> "
			 	.Write " <input type=""radio"" name=""Setting(1)"" value=""1"" "
				If Setting(1) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(1)"" value=""0"" "
				If Setting(1) = "0" Then .Write (" checked")
				.Write "> 否"

			 .Write "</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td width=""32%"" height=""30"" class=""CleftTitle""> <div align='right'><strong>申请空间是否需要审核：</strong></div></td>"
			 .Write "    <td height=""30"">"
			 
			 	.Write " <input type=""radio"" name=""Setting(2)"" value=""1"" "
				If Setting(2) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(2)"" value=""0"" "
				If Setting(2) = "0" Then .Write (" checked")
				.Write "> 否"

			 
			 .Write "      </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>发表日志是否需要审核：</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	
				.Write " <input type=""radio"" name=""Setting(3)"" value=""1"" "
				If Setting(3) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(3)"" value=""0"" "
				If Setting(3) = "0" Then .Write (" checked")
				.Write "> 否"

			 .Write "       </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>发表日志是否允许上传附件：</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	
				.Write " <input type=""radio"" onclick=""$('#fj').show();"" name=""Setting(26)"" value=""1"" "
				If Setting(26) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" onclick=""$('#fj').hide();"" name=""Setting(26)"" value=""0"" "
				If Setting(26) = "0" Then .Write (" checked")
				.Write "> 不允许"
				If Setting(26) = "1" Then
                .Write "<div id='fj' style='color:blue'>"
				Else
                .Write "<div id='fj' style='display:none;color:blue'>"
				End If
				.Write "允许上传的附件扩展名:<input type='text' value='" & Setting(27) & "' name='Setting(27)' /> 多个扩展名用 |隔开,如gif|jpg|rar等</div>"
			 .Write "       </td>"
			 .Write "   </tr>"
			 
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>创建相册是否需要审核：</strong></div></td>"
			  .Write "    <td height=""30"">"
			  
			  	.Write " <input type=""radio"" name=""Setting(4)"" value=""1"" "
				If Setting(4) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(4)"" value=""0"" "
				If Setting(4) = "0" Then .Write (" checked")
				.Write "> 否"
			  
			  .Write "    </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>创建圈子是否需要审核：</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(5)"" value=""1"" "
				If Setting(5) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""0"" "
				If Setting(5) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </td>"
				.Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>用户留言是否需要审核：</strong><br/><font color=red>启用后,用户的留言只有经过后台管理员审核后,前台的空间才可以看到</font></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(24)"" value=""1"" "
				If Setting(24) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(24)"" value=""0"" "
				If Setting(24) = "0" Then .Write (" checked")
				.Write "> 否"
				
				.Write "    </td>"
				.Write "</tr>"
				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>允许游客在空间里评论/留言：</strong></div><font color=red>建议设置不允许,可以有效阻止一些注册机留言</font></td>"
			  .Write "    <td height=""30"">"
			  
			  	.Write " <input type=""radio"" name=""Setting(25)"" value=""1"" "
				If Setting(25) = "1" Then .Write (" checked")
				.Write "> 允许"
				.Write "    <input type=""radio"" name=""Setting(25)"" value=""0"" "
				If Setting(25) = "0" Then .Write (" checked")
				.Write "> 不允许"
			  
			  .Write "    </td>"
			 .Write "   </tr>"				
				
				

			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>每个会员允许创建圈子个数：</strong></div></td>"
				.Write "  <td height=""30"">"
				.Write " <input type=""text"" name=""Setting(6)"" style=""text-align:center"" size=5 value=""" & Setting(6) & """>个，如果不想限制请输入“0”"
				.Write "    </td>"
				.Write "</tr>"

				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>副模板更多空间每页显示：</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(9)"" style=""text-align:center"" size=5 value=""" & Setting(9) & """> 个"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>副模板更多日志每页显示：</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(10)"" style=""text-align:center"" size=5 value=""" & Setting(10) & """> 篇"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>副模板更多圈子每页显示：</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(11)"" style=""text-align:center"" size=5 value=""" & Setting(11) & """> 个"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>副模板更多相册每页显示：</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(12)"" style=""text-align:center"" size=5 value=""" & Setting(12) & """> 本相册 每行显示<input type=""text"" name=""Setting(13)"" style=""text-align:center"" size=5 value=""" & Setting(13) & """> 本"
			 .Write "    </td>"
			 .Write "</tr>"

			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=template-page>"
			.Write "  <H2 class=tab>空间模板</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""template-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>空间首页模板：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(7)"" id='Setting7' type=""text"" value=""" & Setting(7) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting7')[0]") & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>空间副模板：</strong></div><font color=#ff0000>空间的副模板，用于显示更多日志、相册、圈子等，必须包含标签“{$ShowMain}”。</font></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(8)"" id='Setting8' type=""text"" value=""" & Setting(8) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting8')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>交友首页模板：</strong></div><font color=#ff0000>对应/space/friend/index.asp</font></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(23)"" id='Setting23' type=""text"" value=""" & Setting(23) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting23')[0]") & "</td>"
			.Write "    </tr>"
			 .Write " </table>"
			.Write " </div>"

			.Write " <div class=tab-page id=user-page>"
			.Write "  <H2 class=tab>企业空间设置</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""user-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>允许升级为企业空间的用户组：</strong></div><font color=red>不限制,请不要选</font></td>"
			.Write "      <td width=""63%"" height=""30""> &nbsp;" & KS.GetUserGroup_CheckBox("Setting(17)",Setting(17),5) & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>发布企业新闻是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(18)"" value=""1"" "
				If Setting(18) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(18)"" value=""0"" "
				If Setting(18) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>发布企业产品是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(19)"" value=""1"" "
				If Setting(19) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(19)"" value=""0"" "
				If Setting(19) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>发布荣誉证书是否需要审核：</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(20)"" value=""1"" "
				If Setting(20) = "1" Then .Write (" checked")
				.Write "> 是"
				.Write "    <input type=""radio"" name=""Setting(20)"" value=""0"" "
				If Setting(20) = "0" Then .Write (" checked")
				.Write "> 否"
			.Write "</td>"
			.Write "    </tr>"
			.Write " </table>"
			.Write " </div>"
			

			.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
			.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			
			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "if ($('#Setting7').val()=='')" & vbCrLf
			.Write "{ alert('请选择空间首页模板!');" & vbCrLf
			.Write "  $('#Setting7').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "if ($('#Setting8').val()=='')" & vbCrLf
			.Write "{ alert('请选择空间副模板!');" & vbCrLf
			.Write "  $('#Setting8').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "     $('#myform').submit();" & vbCrLf
			.Write " }" & vbCrLf
			.Write " //-->" & vbCrLf
			.Write " </Script>" & vbCrLf
			RS.Close:Set RS = Nothing:Set Conn = Nothing
		End With
		End Sub
	
		

End Class
%> 

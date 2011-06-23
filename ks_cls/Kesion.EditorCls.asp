<!--#Include File="../Conn.asp"-->
<!--#include file="../plus/cc/config.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 4.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Class KesionEditor
   Private KS
   Private AdminDirStr
   Private Style            '调用的模板1文章管理编辑器,2标签编辑器,3模板编辑器等
   Private FullScreenFlag   '全屏标志 0普通,1全屏
   Private ID               '引用的表单ID
   Private Domain,AdminDir
   Private EditorFromType,TemplateType,sChannelID
   Private ShowLabel,DomainStr,InstallStr,ChannelID,ButtonList
   Private ButtonArr(2,31)
   
   Private Sub Class_Initialize()
    On Error ReSume Next
     Set KS=New PublicCls
  End Sub

  Private Sub Class_Terminate()
   Set KS=Nothing 
  End Sub
  
  '参数：EditorFrom   0 代表前台调用 1 代表后台调用
  Sub Kesion(EditorFrom)
     EditorFromType=EditorFrom
	 
	 Domain=KS.GetDomain
	 AdminDir=KS.Setting(89)
     DomainStr=Replace(KS.Setting(2),"/","\\/")
     InstallStr=Replace(KS.Setting(3),"/","\\/")
     AdminDirStr=Replace(KS.Setting(89),"/","\\/")
     Style = Cint(KS.G("Style"))
     ChannelID=KS.G("ChannelID"):IF ChannelID="" Then ChannelID=0
	 sChannelID=KS.G("sChannelID"):IF sChannelID="" Then sChannelID=0
	 TemplateType=KS.G("TemplateType"):IF TemplateType="" Then TemplateType=0
     FullScreenFlag  = KS.G("FullScreenFlag"):IF FullScreenFlag="" Then FullScreenFlag=0

        IF Err THEN 
		   Response.Write("参数传递出错!")
		   Err.clear
		   Exit sub
		  END IF
		 ID=KS.G("ID")
	
		  Call InitialButton()
		     Dim CC
			 If cbool(opentf)=true Then CC=ButtonArr(2,31)

		   Select Case ChannelID
		    Case 10000 '简单编辑器,评论、留言等的调用
			ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
		    Case 9999  '适合简短内容的调用
			ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,7)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,0)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17) &ButtonArr(1,0)& cc& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"

			Case 9998  '前台添加文章调用
			 ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,26)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,7)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,0)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& cc& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
			
		   Case 1   '文章中心
			 ButtonList="<tr><td height=""25"" class=""ToolSet"" style=""border-top:1px solid #999999""><table  height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,1)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,30)&ButtonArr(2,19) &cc &  "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(1,0)& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
	
		 Case 2,3,4,5,6,7  '图片、下载、动漫中心
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(2,18)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
		 Case 9  '考试系统
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"&ButtonArr(2,12) &ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(2,18)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
			 
		Case Else
			IF Style=2 Then  '模板标签调用
			 ' Dim TempBtn:TempBtn="<td width='30' align='center'><img src='" & Domain & "KS_Editor/Images/label0.gif' class='Btn' oncontextmenu='LabelInsertTemplate();return false;' onClick='LabelInsertTemplate();'></td>"
			  ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"& ButtonArr(1,0)&ButtonArr(2,20) & ButtonArr(1,0)& "</tr></table></td></tr>"
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   ElseIf Style=3 Then
			 ' Dim TempBtn:TempBtn="<td width='30' align='center'><img src='" & Domain & "KS_Editor/Images/label0.gif' class='Btn' oncontextmenu='LabelInsertTemplate();return false;' onClick='LabelInsertTemplate();'></td>"
			  ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"& ButtonArr(1,0)&ButtonArr(2,20) & ButtonArr(1,0)& "</tr></table></td></tr>"
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   ElseIf Style=4 Then '博客模板调用
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   Else  '其它地方调用，如公告等
			  ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(1,0)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19)& cc & "</tr></table></td></tr>"
			  ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(1,0)& "</tr></table></td></tr>"
			 
			 ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
	
		   End if
	   End Select

   Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" &vbcrlf
   Response.Write "<html>"&vbcrlf
   Response.Write "<head>"&vbcrlf
   Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"&vbcrlf
   Response.Write "<title>可视编辑器</title>"&vbcrlf
   Response.Write "<link rel=""stylesheet"" href=""" & Domain & "KS_Editor/editor.css"">"&vbcrlf
   Response.Write "</head>"&vbcrlf
   Response.Write "<script language=""JavaScript"" src=""" & Domain & "KS_Inc/editor.js""></script>"&vbcrlf
   		 Dim AdminDir1:AdminDir1=KS.Setting(89)
		 IF Left(AdminDir1,1)="/" Then AdminDir1=Right(AdminDir1,Len(AdminDir1)-1)
		%>
		 <script language="VBScript">
		   Dim Tempstr2
		   Dim domain
		   domain="<%=domain%>"
		  function ReplaceUrl(byval content)
			 TempStr2=Replace(content,"<%=Domain&AdminDir1%>","")
			 TempStr2=Replace(TempStr2,"<%=Domain%>","{$GetInstallDir}")
			 ReplaceUrl=TempStr2
		  end function
		  function ReplaceRealUrl(byval content)
		    ReplaceRealUrl=Replace(content,"{$GetInstallDir}","<%=Domain%>")
		  end function
		 </script>
		<%

   Response.Write "<script language=""vbscript"" src=""" & domain & "KS_Inc/editor.vbs""></script>" & vbcrlf
   Response.Write "<script language=""JavaScript"" src=""" & Domain & "KS_Inc/Common.js""><//script>"&vbcrlf
   Response.Write "<script language=""javascript"" event=""onerror(msg, url, line)"" for=""window"">return true;</script>"&vbcrlf
Response.Write "<script language=""javascript"">"&vbcrlf
 Response.Write "<!--"&vbcrlf
 Response.Write "function InsertCC(html){"&vbcrlf
 Response.Write "html = html.replace(/\[cc\]/g,""http://union.bokecc.com/"")"&vbcrlf
 Response.Write "html = html.replace(/\[\/cc\]/g,"""")"&vbcrlf
 Response.Write "var b = ('<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"" width=""438"" height=""387""><param name=""movie"" value=""' + html + '""><param name=""allowFullScreen"" value=""true"" /><param name=""quality"" value=""high""><embed src=""' + html + '"" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" width=""438"" height=""387"" allowFullscreen=true ></embed></object>')"&vbcrlf
 Response.Write "KS_EditArea.document.body.innerHTML+=b;"&vbcrlf
 Response.Write "}"&vbcrlf
 Response.Write "//-->"&vbcrlf
 Response.Write "</script>"&vbcrlf
   Response.Write "<body oncontextmenu=""return false;"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" class=""yToolbar"">"
   Response.Write "<table height=""100%"" id=""Toolbar"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
   Response.Write ButtonList
   Response.Write "<tr>"&vbcrlf
   Response.Write " <td height=""100%"" valign=""top""><iframe  name=""KS_EditArea"" id=""KS_EditArea"" marginheight=""1"" marginwidth=""1"" style=""font-size:12px;height:100%;width:100%"" scrolling=""yes""></iframe></td>"&vbcrlf
   Response.Write "</tr>"&vbcrlf
   If ChannelID=10000 Then
   Response.Write "<tr style=""display:none""> "&vbcrlf
   Else
   Response.Write "<tr> "&vbcrlf
   End if 
   Response.Write " <td height=""22""> <table width=""100%"" height=""22"" border=""0"" cellpadding=""0"" cellspacing=""0"">"&vbcrlf
   Response.Write "        <tr> "&vbcrlf
   Response.Write "       <td id=""SetModeArea""> <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
   Response.Write "           <tr> "&vbcrlf
   IF  Style<>2 and Style<>3 and Style<>4 Then
   Response.Write "             <td   width=""25"" align=""center"" class=""ModeBarBtnOff"" id=editor_CODE onClick=""setMode('CODE'," & Style& ",'" & DomainStr & "','" & InstallStr & "','" & AdminDirStr & "');""><img src=""" & Domain & "KS_Editor/images/CodeMode.GIF"" height=""15""></td>"&vbcrlf

   Response.Write "                <td width=""25"" height=""20"" align=""center"" class=""ModeBarBtnOff"" id=editor_VIEW onClick=""setMode('VIEW'," & Style & ",'" & DomainStr & "','" & InstallStr & "','" & AdminDirStr & "');""><img src=""" & Domain & "KS_Editor/images/PreviewMode.gif"" height=""15""></td>"&vbcrlf
   Response.Write "             <td width=""25"" align=""center"" class=""ModeBarBtnOn"" id=editor_EDIT onClick=""setMode('EDIT',"& Style & ",'" & DomainStr & "','" & InstallStr & "','" & AdminDirStr & "');""><img src=""" & Domain & "KS_Editor/images/EditMode.GIF"" height=""15""></td>"&vbcrlf
   Response.Write "                <td  width=""25"" height=""20"" align=""center"" class=""ModeBarBtnOff"" id=editor_TEXT onClick=""setMode('TEXT'," & Style & ",'" & DomainStr & "','" & InstallStr & "','" & AdminDirStr & "');""><img src=""" & Domain & "KS_Editor/images/TextMode.GIF""></td>"&vbcrlf
   Else
    Response.Write "<td id='ShowObject'></td>"
   End If
   Response.Write "             <td>&nbsp;</td>"&vbcrlf
   Response.Write "             <td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/tablemodify.gif"" title=""属性"" onClick=""ExeEditAttribute('" & Domain & "');""></td>"&vbcrlf
  
   Response.Write "             <td width=""30"" align=""center""><img title=""关于"" onClick=""AbortArticle('" & Domain & "')"" src=""" & Domain & "KS_Editor/images/Abort.gif"" ></td>"&vbcrlf
			
				IF EditorFromType=1 Then     '后台调用时，才允许最大化
   Response.Write "             <td width=""30"" align=""center""> "&vbcrlf
                
				  IF FullScreenFlag=0 THEN 
   Response.Write "<img onClick=""FullScreen('" & Domain & "'," & Style & "," & ChannelID &");"" class='Btn' src='" & Domain & "KS_Editor/images/fullscreen.gif'>"&vbcrlf
				  else
   Response.Write"<img onClick='parent.close();' class='Btn' src='" & Domain & "KS_Editor/images/minimize.gif'>"&vbcrlf
				  END IF
				
   Response.Write "             </td>"&vbcrlf
   			   End IF	

   Response.Write "             <td  width=""30"" align=""center"" onClick=""ChangeEditAreaHeight(100);""><img src=""" & Domain & "KS_Editor/images/AddHeight.gif""></td>"&vbcrlf
   Response.Write "             <td  width=""30"" align=""center"" onClick=""ChangeEditAreaHeight(-100);""><img src=""" & Domain & "KS_Editor/images/MinusHeight.gif""></td>"&vbcrlf
  Response.Write "            </tr>"&vbcrlf
  Response.Write "          </table></td>"&vbcrlf
  Response.Write "      </tr>"&vbcrlf
  Response.Write "    </table></td>"&vbcrlf
  Response.Write "</tr>"&vbcrlf
  Response.Write "</table>"&vbcrlf
  Response.Write "</body>"&vbcrlf
  Response.Write "</html>"&vbcrlf
  Response.Write "<script language=""JavaScript"">"&vbcrlf
  Response.Write "var DocElementArrInitialFlag=false;"&vbcrlf
  Response.Write "var InstallDir='" & Domain & "';"&vbcrlf
  Response.Write "function document.onreadystatechange()"&vbcrlf
  Response.Write "{    "&vbcrlf
  Response.Write "	if (document.readyState!=""complete"") return;"&vbcrlf
  Response.Write "	if (DocElementArrInitialFlag) return;"&vbcrlf
  Response.Write "	DocElementArrInitialFlag = true;"&vbcrlf
  Response.Write "	var i,s,curr;"&vbcrlf
  Response.Write "	for (i=0; i<document.body.all.length;i++)"&vbcrlf
  Response.Write "	{"
  Response.Write "		curr=document.body.all[i];"&vbcrlf
  Response.Write "		if (curr.className==""Btn"") InitBtn(curr);"&vbcrlf
  Response.Write "	}"
  Response.Write "	SetEditAreaStyle();"&vbcrlf
  Response.Write "	ShowTableBorders();"&vbcrlf
  Response.Write "	SetContentIni();"&vbcrlf
  Response.Write "	LoadEditFile(" & Style & ");"&vbcrlf
  Response.Write "}"
 
 Select Case (Style)
  Case 0
   Response.Write "function SetContentIni()"&vbcrlf
   Response.Write "{"
   Response.Write " for (var i=0;i<parent.document.forms.length;i++)"&vbcrlf
   Response.Write "   if (parent.document.forms[i]." & ID & "!=null)"&vbcrlf
   Response.Write "	   {"&vbcrlf
   Response.Write "   KS_EditArea.document.body.innerHTML=parent.document.forms[i]." & ID & ".value;"&vbcrlf
   Response.Write "	   ShowTableBorders();"&vbcrlf
   Response.Write "	   }"
   Response.Write "}"
 Case 1
   Response.Write "var ArticleContentArray=new Array('');"&vbcrlf
   Response.Write "function SetContentIni()"&vbcrlf
   Response.Write "{   var AlreadyExistsContent;"&vbcrlf
   Response.Write "  for (var i=0;i<parent.document.forms.length;i++)"&vbcrlf
   Response.Write "   if (parent.document.forms[i]." & ID & "!=null)"&vbcrlf
	Response.Write "   {"
	Response.Write "   AlreadyExistsContent=unescape(parent.document.forms[i]." & ID & ".value);"&vbcrlf
	Response.Write "   }"&vbcrlf
	Response.Write "if (AlreadyExistsContent!='')"&vbcrlf
	Response.Write "{"
	Response.Write "	var TempArray;"&vbcrlf
	Response.Write "	TempArray=AlreadyExistsContent.split('[NextPage]');"&vbcrlf
	Response.Write "	for (var i=0;i<TempArray.length;i++)"&vbcrlf
	Response.Write "	{"
	Response.Write "		ArticleContentArray[i+1]=TempArray[i];"&vbcrlf
	Response.Write "	}"
	Response.Write "	SetArticleContent();"&vbcrlf
	Response.Write "}"
	Response.Write "else"
	Response.Write "{"
	Response.Write "	KS_EditArea.document.body.innerHTML='';"&vbcrlf
	Response.Write "}"
    Response.Write "}"
    Response.Write "function SetArticleContent()"&vbcrlf
    Response.Write "{"
	Response.Write "var PageSelectObj=document.all.PageNumSelect;"&vbcrlf
	Response.Write "if (ArticleContentArray.length>=2)"&vbcrlf
	Response.Write "{"
    Response.Write "		PageSelectObj.options.remove(0);"&vbcrlf
	Response.Write "	for (var i=1;i<ArticleContentArray.length;i++)"&vbcrlf
	Response.Write "	{   "
    Response.Write "			var AddOption = document.createElement(""OPTION"");"&vbcrlf
	Response.Write "		AddOption.text=i;"
    Response.Write "			AddOption.value=i;"
	Response.Write "		PageSelectObj.add(AddOption);"&vbcrlf
    Response.Write "		}"
    Response.Write "		PageSelectObj.options(0).selected=true;"&vbcrlf
	Response.Write "	KS_EditArea.document.body.innerHTML=ArticleContentArray[1];"&vbcrlf
	Response.Write "}"&vbcrlf
    Response.Write "	ShowTableBorders();"&vbcrlf
    Response.Write "}"
    Response.Write "function NewPage()"&vbcrlf
    Response.Write "{"
    Response.Write "	var PageSelectObj=document.all.PageNumSelect;"&vbcrlf
	Response.Write "var PageNum=PageSelectObj.options.length;"&vbcrlf
	Response.Write "ArticleContentArray[parseInt(PageNum)]=unescape(KS_EditArea.document.body.innerHTML);"
	Response.Write "KS_EditArea.document.body.innerHTML='';"&vbcrlf
	Response.Write "var CurrPage=PageNum+1;"&vbcrlf
	Response.Write "var AddOption = document.createElement(""OPTION"");"&vbcrlf
    Response.Write "	AddOption.text=CurrPage;"&vbcrlf
    Response.Write "	AddOption.value=CurrPage;"&vbcrlf
    Response.Write "	document.all.PageNumSelect.add(AddOption);"&vbcrlf
    Response.Write "	document.all.PageNumSelect.options(document.all.PageNumSelect.length-1).selected=true;"&vbcrlf
    Response.Write "	KS_EditArea.focus();"
    Response.Write "	ShowTableBorders();"&vbcrlf
    Response.Write "}"
    Response.Write "function ChangePage(PageIndex)"&vbcrlf
    Response.Write "{"
    Response.Write "	var CurrPage=parseInt(PageIndex);"
    Response.Write "	KS_EditArea.document.body.innerHTML=ArticleContentArray[CurrPage];"&vbcrlf
    Response.Write "	KS_EditArea.focus();"&vbcrlf
    Response.Write "	ShowTableBorders();"&vbcrlf
    Response.Write "}"
    Response.Write "function SaveCurrPage()"&vbcrlf
    Response.Write "{"
    Response.Write "	var SelectObj=document.all.PageNumSelect;"&vbcrlf
    Response.Write "	var PageIndex=parseInt(SelectObj.options(SelectObj.selectedIndex).value);"&vbcrlf
    Response.Write "	ArticleContentArray[PageIndex]=KS_EditArea.document.body.innerHTML;"&vbcrlf
    Response.Write "	ShowTableBorders();"&vbcrlf
    Response.Write "}"
    Response.Write "function DeletePage()"&vbcrlf
    Response.Write "{"
    Response.Write "	var PageNum=document.all.PageNumSelect.options.length,i=0;"&vbcrlf
    Response.Write "	if (PageNum==1) return;"&vbcrlf
    Response.Write "	 document.all.PageNumSelect.options(PageNum-1).selected=true;	"&vbcrlf
    Response.Write "       KS_EditArea.document.body.innerHTML=ArticleContentArray[PageNum-1];"&vbcrlf
    Response.Write "		ArticleContentArray[PageNum]='';"
    Response.Write "		document.all.PageNumSelect.options(PageNum-2).selected=true;"&vbcrlf
    Response.Write "		document.all.PageNumSelect.options.remove(PageNum-1);"&vbcrlf
    Response.Write "	KS_EditArea.focus();"
    Response.Write "	ShowTableBorders();"&vbcrlf
    Response.Write "}"
    Response.Write "function SearchObject()"&vbcrlf
    Response.Write "{"
    Response.Write "	KS_EditArea.focus();"&vbcrlf
    Response.Write "	UpdateToolbar();"&vbcrlf
    Response.Write "}"
Case 2,3,4
    Response.Write "function SetContentIni()"&vbcrlf
    Response.Write "{"
    Response.Write "    for (var i=0;i<parent.document.forms.length;i++)"&vbcrlf
    Response.Write "      if (parent.document.forms[i]." & ID & "!=null)"&vbcrlf
    Response.Write "	   {"
    Response.Write "	   frames[0].document.open();"&vbcrlf
   ' Response.Write "       frames[0].document.body.innerHTML=parent.document.forms[i]." & ID & ".value;"
    Response.Write "       frames[0].document.write(unescape(ReplaceRealUrl(ReplaceScriptToImg(parent.document.forms[i]." & ID & ".value))));"
    Response.Write "	   frames[0].document.close();"&vbcrlf
    Response.Write "	   ShowTableBorders();"&vbcrlf
    Response.Write "	   }"
    Response.Write "}"&vbcrlf
    Response.Write "function LabelInsertTemplate()"&vbcrlf
    Response.Write "{"
    Response.Write "	var ReturnValue='';"&vbcrlf
    Response.Write "	ReturnValue=OpenWindow('" & Domain & AdminDir &"Include/LabelFrame.asp?sChannelID=" & sChannelID & "&TemplateType=" & TemplateType&"&url=InsertLabel.asp&pagetitle='+escape('插入标签'),260,350,window);"&vbcrlf
    Response.Write "	if (ReturnValue!='') parent.frames[0].InsertHTMLStr(ReturnValue);"
    Response.Write "}"&vbcrlf
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('LabelFrame.asp?url=InsertLabel.asp&pagetitle='+escape('插入标签'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ parent.frames[0].InsertHTMLStr(Val);"
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function WapLabelInsertTemplate()"&vbcrlf
		Response.Write "{"
		Response.Write "	var ReturnValue='';"&vbcrlf
		Response.Write "	ReturnValue=OpenWindow('" & Domain & AdminDir &"Include/LabelFrame.asp?sChannelID=" & sChannelID & "&TemplateType=" & TemplateType&"&url="&Domain&"Plus/Wap/Wap_InsertLabel.asp&pagetitle='+escape('插入WAP标签'),250,300,window);"&vbcrlf
		Response.Write "	if (ReturnValue!='') parent.frames[0].InsertHTMLStr(ReturnValue);"
		Response.Write "}"&vbcrlf
		Response.Write "function InsertFunctionLabel(Url,Width,Height)" & vbcrlf
        Response.Write "{" & vbcrlf
        Response.Write "var Val = OpenWindow(Url,Width,Height,window);"
		Response.Write "if (Val!=''&&Val!=null)"
		Response.Write "{ parent.frames[0].InsertHTMLStr(Val);"
		Response.Write " }" & vbcrlf
        Response.Write "}" & vbcrlf
Case Else
    Response.Write "function SetContentIni()"&vbcrlf
    Response.Write "{"
    Response.Write "    for (var i=0;i<parent.document.forms.length;i++)"&vbcrlf
    Response.Write "      if (parent.document.forms[i]." & ID & "!=null)"&vbcrlf
    Response.Write "	   {"
    Response.Write "	   frames[0].document.open();"&vbcrlf
   ' Response.Write "       frames[0].document.body.innerHTML=parent.document.forms[i]." & ID & ".value;"
    Response.Write "       frames[0].document.write(unescape(ReplaceRealUrl(ReplaceScriptToImg(parent.document.forms[i]." & ID & ".value))));"
    Response.Write "	   frames[0].document.close();"&vbcrlf
    Response.Write "	   ShowTableBorders();"&vbcrlf
    Response.Write "	   }"
    Response.Write "}"&vbcrlf

End Select
    Response.Write "</script>"&vbcrlf
End Sub

'初始化按钮
Sub InitialButton()
'分隔线
ButtonArr(1,0)="<td width=""1""> <div align=""center"" class=""ToolSeparator""></div></td>"
'分页按钮
ButtonArr(1,1)="<td width='30' align='center'><img src='" & Domain & "KS_Editor/images/NewDoc.gif' width='23' height='22' class='Btn' title='新建文档' oncontextmenu='NewPage();return false;' onClick='NewPage()' ></td>" &_
           "<td align='center'> <select style='width:38px;height:18px'  onFocus='SaveCurrPage();' onChange='ChangePage(this.value);' name='PageNumSelect'>" &_
           "   <option value='1' selected>1</option>" &_
           " </select> </td>" &_
           "<td width='30' align='center'><img oncontextmenu='DeletePage();return false;' onClick='DeletePage();' alt='删除最后一页' class='Btn' src='" & Domain & "KS_Editor/images/DelDoc.gif'></td>" &_
           "<td width='30' align='center'><img title='插入分页符号' oncontextmenu='InsertMorePage();return false;' onClick='InsertMorePage()' class='Btn' src='" & Domain & "KS_Editor/images/InsertPage.gif'></td>" 
'撤销复做按钮
ButtonArr(1,2)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/undo.gif"" class=""Btn"" title=""撤消"" oncontextmenu=""Format('undo');return false;"" onClick=""Format('undo')"" ></td><td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/redo.gif"" class=""Btn"" title=""恢复"" oncontextmenu=""Format('redo');return false;"" onClick=""Format('redo')"" ></td>"	
'查找 / 替换 按钮
ButtonArr(1,3)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/find.gif"" class=""Btn"" title=""查找 / 替换"" oncontextmenu=""SearchStr('" & Domain & "');return false;"" onClick=""SearchStr('" & Domain & "');"" ></td>"
'计算器按钮
ButtonArr(1,4)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/calculator.gif"" class=""Btn"" title=""计算器"" oncontextmenu=""Calculator('" & Domain & "');return false;"" onClick=""Calculator('" & Domain & "')"" ></td>"	
'插入当前日期
ButtonArr(1,5)="<td width=""30"" align=""center""><img title=""插入当前日期"" oncontextmenu=""InsertDate();return false;"" onClick=""InsertDate()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/date.gif"" ></td>"
'插入当前时间
ButtonArr(1,6)="<td width=""30"" align=""center""><img title=""插入当前时间"" oncontextmenu=""InsertTime();return false;"" onClick=""InsertTime()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/time.gif"" ></td>"
'删除所有HTML标识
ButtonArr(1,7)="<td width=""30"" align=""center""><img title=""删除所有HTML标识"" oncontextmenu=""DelAllHtmlTag();return false;"" onClick=""DelAllHtmlTag()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/geshi.gif"" ></td>"
'删除文字格式
ButtonArr(1,8)="<td width=""30"" align=""center""><img title=""删除文字格式"" oncontextmenu=""Format('removeformat');return false;"" onClick=""Format('removeformat')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/clear.gif"" ></td>"
'插入超级连接
ButtonArr(1,9)="<td width=""30"" align=""center""><img title=""插入超级连接"" oncontextmenu=""Format('CreateLink');return false;"" onClick=""Format('CreateLink')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/url.gif"" ></td>"
'取消超级链接
ButtonArr(1,10)="<td width=""30"" align=""center""><img title=""取消超级链接"" oncontextmenu=""Format('unLink');return false;"" onClick=""Format('unLink')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/nourl.gif"" ></td>"
'插入网页
ButtonArr(1,11)="<td width=""30"" align=""center""><img title=""插入网页"" oncontextmenu=""InsertPage('" & Domain & "');return false;"" onClick=""InsertPage('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/htm.gif"" ></td>"
'插入栏目框
ButtonArr(1,12)="<td width=""30"" align=""center""><img title=""插入栏目框"" oncontextmenu=""InsertFrame('" & Domain & "');return false;"" onClick=""InsertFrame('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/fieldset.gif"" ></td>"
'插入Excel表格
ButtonArr(1,13)="<td width=""30"" align=""center""><img title=""插入Excel表格"" oncontextmenu=""InsertExcel();return false;"" onClick=""InsertExcel()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/excel.gif"" ></td>"
'插入滚动文本
ButtonArr(1,14)="<td width=""30"" align=""center""><img title=""插入滚动文本"" oncontextmenu=""InsertMarquee('" & Domain & "');return false;"" onClick=""InsertMarquee('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Marquee.gif"" ></td>"
'图文并排
ButtonArr(1,15)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/PicAlign.gif"" width=""23"" height=""22"" class=""Btn"" title=""图文并排"" oncontextmenu=""PicAndTextArrange('" & Domain & "');return false;"" onClick=""PicAndTextArrange('" & Domain & "')"" ></td>"
'插入表格
ButtonArr(1,16)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/Inserttable.gif""  class=""Btn"" title=""插入表格"" oncontextmenu=""InsertTable('" & Domain & "');return false;""  onClick=""InsertTable('" & Domain & "')""></td>"
'插入行
ButtonArr(1,17)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/inserttable1.gif""  class=""Btn"" title=""插入行"" oncontextmenu=""InsertRow();return false;""  onClick=""InsertRow()""></td>"
'插入列
ButtonArr(1,18)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/inserttablec.gif"" class=""Btn"" title=""插入列"" oncontextmenu=""InsertColumn();return false;"" onClick=""InsertColumn()""></td>"
'删除列
ButtonArr(1,19)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/delinserttablec.gif""  class=""Btn"" title=""删除列"" oncontextmenu=""DeleteColumn();return false;"" onClick=""DeleteColumn()""></td>"
'插入单元格
ButtonArr(1,20)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/insterttable2.gif"" width=""23"" height=""22""  class=""Btn"" title=""插入单元格"" oncontextmenu=""InsertCell();return false;"" onClick=""InsertCell()""></td>"
'删除单元格
ButtonArr(1,21)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/delinsterttable2.gif"" width=""23"" height=""22""  class=""Btn"" title=""删除单元格"" oncontextmenu=""DeleteCell();return false;"" onClick=""DeleteCell()""></td>"
'拆分列
ButtonArr(1,22)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/SplitTD.gif"" width=""23"" height=""22""  class=""Btn"" title=""拆分列"" oncontextmenu=""SplitColumn();return false;"" onClick=""SplitColumn()""></td>"
'合并列
ButtonArr(1,23)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/MargeTD.gif"" width=""23"" height=""22""  class=""Btn"" title=""合并列"" oncontextmenu=""MergeColumn();return false;"" onClick=""MergeColumn()""></td>"
'合并行
ButtonArr(1,24)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/Hbtable.gif"" width=""23"" height=""22""  class=""Btn"" title=""合并行"" oncontextmenu=""MergeRow();return false;"" onClick=""MergeRow()""></td>"
'拆分行
ButtonArr(1,25)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/cftable.gif""  class=""Btn"" title=""拆分行"" oncontextmenu=""SplitRows();return false;"" onClick=""SplitRows()""></td>"
'前台分页按钮
ButtonArr(1,26)="<td width=""90"" align=""center""><a title='插入分页符号' oncontextmenu='InsertMorePage();return false;' onClick='InsertMorePage()' class='Btn'><img src=""" & Domain & "KS_Editor/images/NewDoc.gif""></a></td>"

'字体和字号
ButtonArr(2,1)="<td width=""26"" align=""center""> <select name=""select2"" style='width:60px' class=""ToolSelectStyle"" onChange=""Format('fontname',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>字体</option><option value=""宋体"">宋体</option><option value=""黑体"">黑体</option><option value=""楷体_GB2312"">楷体</option><option value=""仿宋_GB2312"">仿宋</option><option value=""隶书"">隶书</option><option value=""幼圆"">幼圆</option><option value=""Arial"">Arial</option><option value=""Arial Black"">Arial Black</option><option value=""Arial Narrow"">Arial Narrow</option><option value=""Brush Script	MT"">Brush Script MT</option><option value=""Century Gothic"">Century Gothic</option><option value=""Comic Sans MS"">Comic Sans MS</option><option value=""Courier"">Courier</option><option value=""Courier New"">Courier New</option><option value=""MS Sans Serif"">MS Sans Serif</option><option value=""Script"">Script</option><option value=""System"">System</option><option value=""Times New Roman"">Times New Roman</option><option value=""Verdana"">Verdana</option><option value=""Wide Latin"">Wide Latin</option><option value=""Wingdings"">Wingdings</option></select></td><td width=""26"" align=""center""> <select name=""select3"" style=""width:48px;height:18px"" onChange=""Format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>字号</option><option value=""7"">一号</option><option value=""6"">二号</option><option value=""5"">三号</option><option value=""4"">四号</option><option value=""3"">五号</option><option value=""2"">六号</option><option value=""1"">七号</option></select></td>"
'上标、下标、加粗、斜体、下划线和删除线
ButtonArr(2,2)="<td width=""26"" align=""center""><img title=""上标"" oncontextmenu=""Format('superscript');return false;"" onClick=""Format('superscript')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/superscript.gif"" ></td><td width=""26"" align=""center""><img title=""下标"" oncontextmenu=""Format('subscript');return false;"" onClick=""Format('subscript')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/subscript.gif"" ></td><td width=""26"" align=""center""><img title=""加粗"" oncontextmenu=""Format('bold');return false;"" onClick=""Format('bold')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/bold.gif"" ></td><td width=""26"" align=""center""><img title=""斜体"" oncontextmenu=""Format('italic');return false;"" onClick=""Format('italic')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/italic.gif"" ></td><td width=""26"" align=""center""><img title=""下划线"" oncontextmenu=""Format('underline');return false;"" onClick=""Format('underline')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/underline.gif"" ></td><td  width=""26"" align=""center""><img title=""删除线"" oncontextmenu=""Format('StrikeThrough');return false;"" onClick=""Format('StrikeThrough')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/strikethrough.gif"" width=""20"" height=""20""></td>"

'文字颜色
ButtonArr(2,3)="<td  width=""26"" align=""center""><img src=""" & Domain & "KS_Editor/images/TextColor.gif"" class=""Btn"" title=""文字颜色"" oncontextmenu=""TextColor('" & Domain & "');return false;"" onClick=""TextColor('" & Domain & "')"" ></td>"
'文字背景色
ButtonArr(2,4)="<td  width=""26"" align=""center""><img title=""文字背景色"" oncontextmenu=""TextBGColor('" & Domain & "');return false;"" onClick=""TextBGColor('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/fgbgcolor.gif"" ></td>"
'左对齐、居中、右对齐
ButtonArr(2,5)="<td width=""26"" align=""center""><img title=""左对齐"" oncontextmenu=""Format('justifyleft');return false;"" onClick=""Format('justifyleft')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Aleft.gif"" ></td><td width=""26"" align=""center""><img title=""居中"" oncontextmenu=""Format('justifycenter');return false;"" onClick=""Format('justifycenter')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Acenter.gif"" ></td><td width=""23"" align=""center""><img title=""右对齐"" oncontextmenu=""Format('justifyright');return false;"" onClick=""Format('justifyright')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Aright.gif"" ></td>"
'减少缩进量、增加缩进量
ButtonArr(2,6)="<td width=""23"" align=""center""><img title=""减少缩进量"" oncontextmenu=""Format('outdent');return false;"" onClick=""Format('outdent');"" class=""Btn"" src=""" & Domain & "KS_Editor/images/outdent.gif"" ></td><td width=""30"" align=""center""><img title=""增加缩进量"" oncontextmenu=""Format('indent');return false;"" onClick=""Format('indent')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/indent.gif"" ></td>"
'段落样式
ButtonArr(2,7)="<td width=""30"" align=""center""><select name=""select"" style=""width:80px;height:18px"" onChange=""Format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>段落样式</option><option value=""&lt;P&gt;"">普通</option><option value=""&lt;H1&gt;"">标题一</option><option value=""&lt;H2&gt;"">标题二</option><option value=""&lt;H3&gt;"">标题三</option><option value=""&lt;H4&gt;"">标题四</option><option value=""&lt;H5&gt;"">标题五</option><option value=""&lt;H6&gt;"">标题六</option><option value=""&lt;p&gt;"">段落</option><option value=""&lt;dd&gt;"">定义</option><option value=""&lt;dt&gt;"">术语定义</option><option value=""&lt;dir&gt;"">目录列表</option><option value=""&lt;menu&gt;"">菜单列表</option><option value=""&lt;PRE&gt;"">已编排格式</option></select></td>"
'项目符号
ButtonArr(2,8)="<td width=""30"" align=""center""><img title=""项目符号"" oncontextmenu=""Format('insertunorderedlist');return false;"" onClick=""Format('insertunorderedlist')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/list.gif"" ></td>"
'编号
ButtonArr(2,9)="<td width=""30"" align=""center""><img title=""编号"" oncontextmenu=""Format('insertorderedlist');return false;"" onClick=""Format('insertorderedlist')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/num.gif"" ></td>"
'绝对或相对位置
ButtonArr(2,10)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/abspos.gif"" class=""Btn"" title=""绝对或相对位置"" oncontextmenu=""Pos();return false;"" onClick=""Pos();""></td>"
'插入特殊符号
ButtonArr(2,11)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/symbol.gif"" class=""Btn"" title=""插入特殊符号"" oncontextmenu=""InsertSymbol('" & Domain & "');return false;"" onClick=""InsertSymbol('" & Domain & "');""></td>"
'插入图片
ButtonArr(2,12)="<td width=""30"" align=""center""><img title=""插入图片，支持格式为：jpg、gif、bmp、png等"" oncontextmenu=""InsertPicture(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertPicture(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/img.gif"" ></td>"
'插入flash多媒体文件
ButtonArr(2,13)="<td width=""30"" align=""center""><img title=""插入flash多媒体文件"" oncontextmenu=""InsertFlash(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertFlash(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/flash.gif"" ></td>"
'插入视频文件
ButtonArr(2,14)="<td width=""30"" align=""center""><img title=""插入视频文件，支持格式为：avi、wmv、asf、mpg"" oncontextmenu=""InsertVideo(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertVideo(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/wmv.gif"" ></td>"
'插入RealPlay文件
ButtonArr(2,15)="<td width=""30"" align=""center""><img title=""插入RealPlay文件，支持格式为：rm、ra、ram"" oncontextmenu=""InsertRM(" & EditorFromType & ",'" & Domain & "',"& ChannelID & ");return false;"" onClick=""InsertRM(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/rm.gif"" ></td>"
'插入特殊水平线
ButtonArr(2,16)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/sline.gif"" class=""Btn"" title=""插入特殊水平线"" oncontextmenu=""SpecialHR('" & Domain & "');return false;"" onClick=""SpecialHR('" & Domain & "')"" ></td>"
'插入普通水平线
ButtonArr(2,17)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/line.gif"" class=""Btn"" title=""插入普通水平线"" oncontextmenu=""InsertHR();return false;"" onClick=""InsertHR();"" ></td>"
'常规粘贴
ButtonArr(2,18)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/paste.gif"" class=""Btn"" title=""常规粘贴"" oncontextmenu=""Paste();return false;"" onClick=""Paste()"" ></td>"
'插入换行符号
ButtonArr(2,19)="<td width=""30"" align=""center""><img title=""插入换行符号"" oncontextmenu=""InsertBR();return false;"" onClick=""InsertBR()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/chars.gif"" ></td>"
'文本粘贴
ButtonArr(2,30)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/PasteText.gif"" class=""Btn"" title=""文本粘贴"" oncontextmenu=""PasteText();return false;"" onClick=""PasteText()"" ></td>"

dim fb
if buttonstyle=1 then
 fb="plugin.swf"
else
 fb="plugin_" & buttonstyle & ".swf"
end if

ButtonArr(2,31)="<td><!-- cc视频插件代码 --><object width='72' height='24'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin/" & fb & "?userID=" & userid &"&type=normal' /><embed src='http://union.bokecc.com/flash/plugin/" & fb & "?userID=" & userid & "&type=normal' type='application/x-shockwave-flash' width='72' height='24' allowScriptAccess='always' wmode='transparent'></embed></object><!-- cc视频插件代码 --></td>"

'常用标签列表
 Dim MyLabelStr
		 MyLabelStr=" <select name=""mylabel"" style=""width:160px"">"
		 MyLabelStr=MyLabelStr & " <option value="""">==选择系统函数标签==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 MyLabelStr=MyLabelStr & "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  MyLabelStr=MyLabelStr & "</select>&nbsp;<input type='button' class='tdbg' onclick='LabelInsertCode(document.all.mylabel.value);' value='插入标签'>&nbsp;"
		  RS.Close:Set RS=Nothing
ButtonArr(2,20)="<td>&nbsp;" &MyLabelStr & "<input type=""button"" class='tdbg' onclick=""javascript:WapLabelInsertTemplate();"" value=""WAP标签"">&nbsp;<input type=""button"" class='tdbg' onclick=""javascript:LabelInsertTemplate();"" value=""选择更多标签""></td>"

'插入附件文件
ButtonArr(2,21)="<td width=""30"" align=""center""><img title=""插入附件文件，支持格式为：rar、zip、txt、doc、xls"" oncontextmenu=""InsertUpFile(" & EditorFromType & ",'" & Domain & "',"& ChannelID & ");return false;"" onClick=""InsertUpFile(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/down.gif"" ></td>"
End Sub
End Class
%> 

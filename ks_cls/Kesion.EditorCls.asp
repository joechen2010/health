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
   Private Style            '���õ�ģ��1���¹���༭��,2��ǩ�༭��,3ģ��༭����
   Private FullScreenFlag   'ȫ����־ 0��ͨ,1ȫ��
   Private ID               '���õı�ID
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
  
  '������EditorFrom   0 ����ǰ̨���� 1 �����̨����
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
		   Response.Write("�������ݳ���!")
		   Err.clear
		   Exit sub
		  END IF
		 ID=KS.G("ID")
	
		  Call InitialButton()
		     Dim CC
			 If cbool(opentf)=true Then CC=ButtonArr(2,31)

		   Select Case ChannelID
		    Case 10000 '�򵥱༭��,���ۡ����Եȵĵ���
			ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
		    Case 9999  '�ʺϼ�����ݵĵ���
			ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,7)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,0)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17) &ButtonArr(1,0)& cc& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"

			Case 9998  'ǰ̨������µ���
			 ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,26)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,7)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,8)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,0)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& cc& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
			
		   Case 1   '��������
			 ButtonList="<tr><td height=""25"" class=""ToolSet"" style=""border-top:1px solid #999999""><table  height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,1)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,30)&ButtonArr(2,19) &cc &  "</tr></table></td></tr>"
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(1,0)& "</tr></table></td></tr>"
			 
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
	
		 Case 2,3,4,5,6,7  'ͼƬ�����ء���������
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(2,18)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
		 Case 9  '����ϵͳ
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"&ButtonArr(2,12) &ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(2,18)&ButtonArr(2,19)&ButtonArr(1,0)& "</tr></table></td></tr>"
			ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
			 
		Case Else
			IF Style=2 Then  'ģ���ǩ����
			 ' Dim TempBtn:TempBtn="<td width='30' align='center'><img src='" & Domain & "KS_Editor/Images/label0.gif' class='Btn' oncontextmenu='LabelInsertTemplate();return false;' onClick='LabelInsertTemplate();'></td>"
			  ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"& ButtonArr(1,0)&ButtonArr(2,20) & ButtonArr(1,0)& "</tr></table></td></tr>"
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   ElseIf Style=3 Then
			 ' Dim TempBtn:TempBtn="<td width='30' align='center'><img src='" & Domain & "KS_Editor/Images/label0.gif' class='Btn' oncontextmenu='LabelInsertTemplate();return false;' onClick='LabelInsertTemplate();'></td>"
			  ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"& ButtonArr(1,0)&ButtonArr(2,20) & ButtonArr(1,0)& "</tr></table></td></tr>"
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   ElseIf Style=4 Then '����ģ�����
			 ButtonList=ButtonList & "<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(1,0)&ButtonArr(1,3)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,11)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19) & "</tr></table></td></tr>"
						 
			 ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,10)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(2,17)&ButtonArr(1,0)& "</tr></table></td></tr>"
		   Else  '�����ط����ã��繫���
			  ButtonList="<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & ButtonArr(1,0)&ButtonArr(2,1)&ButtonArr(1,0)&ButtonArr(1,4)&ButtonArr(1,5)&ButtonArr(1,6)&ButtonArr(1,0)&ButtonArr(1,7)&ButtonArr(1,8)&ButtonArr(1,0)&ButtonArr(1,9)&ButtonArr(1,10)&ButtonArr(1,0)&ButtonArr(1,12)&ButtonArr(1,13)&ButtonArr(1,14)&ButtonArr(1,0)&ButtonArr(1,15)&ButtonArr(1,16)&ButtonArr(1,17)&ButtonArr(1,18)&ButtonArr(1,19)&ButtonArr(1,23)&ButtonArr(1,24)&ButtonArr(1,25)&ButtonArr(1,0)&ButtonArr(2,18)&ButtonArr(2,19)& cc & "</tr></table></td></tr>"
			  ButtonList=ButtonList &"<tr><td height=""25"" class=""ToolSet""><table height=""25"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" &ButtonArr(1,0)&ButtonArr(1,2)&ButtonArr(2,2)&ButtonArr(2,3)&ButtonArr(2,4)&ButtonArr(2,5)&ButtonArr(2,6)&ButtonArr(2,7)&ButtonArr(2,8)&ButtonArr(2,9)&ButtonArr(2,11)&ButtonArr(2,12)&ButtonArr(2,13)&ButtonArr(2,14)&ButtonArr(2,15)&ButtonArr(2,16)&ButtonArr(1,0)& "</tr></table></td></tr>"
			 
			 ButtonList=ButtonList &"<tr style='display:none'><td height='22'><table width='100%' height='100%' border='0' cellpadding='0' cellspacing='0' class='ToolSet'><tr> <td height='22' id='ShowObject'>&nbsp;</td></tr></table></td></tr>"
	
		   End if
	   End Select

   Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" &vbcrlf
   Response.Write "<html>"&vbcrlf
   Response.Write "<head>"&vbcrlf
   Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"&vbcrlf
   Response.Write "<title>���ӱ༭��</title>"&vbcrlf
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
   Response.Write "             <td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/tablemodify.gif"" title=""����"" onClick=""ExeEditAttribute('" & Domain & "');""></td>"&vbcrlf
  
   Response.Write "             <td width=""30"" align=""center""><img title=""����"" onClick=""AbortArticle('" & Domain & "')"" src=""" & Domain & "KS_Editor/images/Abort.gif"" ></td>"&vbcrlf
			
				IF EditorFromType=1 Then     '��̨����ʱ�����������
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
    Response.Write "	ReturnValue=OpenWindow('" & Domain & AdminDir &"Include/LabelFrame.asp?sChannelID=" & sChannelID & "&TemplateType=" & TemplateType&"&url=InsertLabel.asp&pagetitle='+escape('�����ǩ'),260,350,window);"&vbcrlf
    Response.Write "	if (ReturnValue!='') parent.frames[0].InsertHTMLStr(ReturnValue);"
    Response.Write "}"&vbcrlf
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('LabelFrame.asp?url=InsertLabel.asp&pagetitle='+escape('�����ǩ'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ parent.frames[0].InsertHTMLStr(Val);"
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function WapLabelInsertTemplate()"&vbcrlf
		Response.Write "{"
		Response.Write "	var ReturnValue='';"&vbcrlf
		Response.Write "	ReturnValue=OpenWindow('" & Domain & AdminDir &"Include/LabelFrame.asp?sChannelID=" & sChannelID & "&TemplateType=" & TemplateType&"&url="&Domain&"Plus/Wap/Wap_InsertLabel.asp&pagetitle='+escape('����WAP��ǩ'),250,300,window);"&vbcrlf
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

'��ʼ����ť
Sub InitialButton()
'�ָ���
ButtonArr(1,0)="<td width=""1""> <div align=""center"" class=""ToolSeparator""></div></td>"
'��ҳ��ť
ButtonArr(1,1)="<td width='30' align='center'><img src='" & Domain & "KS_Editor/images/NewDoc.gif' width='23' height='22' class='Btn' title='�½��ĵ�' oncontextmenu='NewPage();return false;' onClick='NewPage()' ></td>" &_
           "<td align='center'> <select style='width:38px;height:18px'  onFocus='SaveCurrPage();' onChange='ChangePage(this.value);' name='PageNumSelect'>" &_
           "   <option value='1' selected>1</option>" &_
           " </select> </td>" &_
           "<td width='30' align='center'><img oncontextmenu='DeletePage();return false;' onClick='DeletePage();' alt='ɾ�����һҳ' class='Btn' src='" & Domain & "KS_Editor/images/DelDoc.gif'></td>" &_
           "<td width='30' align='center'><img title='�����ҳ����' oncontextmenu='InsertMorePage();return false;' onClick='InsertMorePage()' class='Btn' src='" & Domain & "KS_Editor/images/InsertPage.gif'></td>" 
'����������ť
ButtonArr(1,2)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/undo.gif"" class=""Btn"" title=""����"" oncontextmenu=""Format('undo');return false;"" onClick=""Format('undo')"" ></td><td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/redo.gif"" class=""Btn"" title=""�ָ�"" oncontextmenu=""Format('redo');return false;"" onClick=""Format('redo')"" ></td>"	
'���� / �滻 ��ť
ButtonArr(1,3)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/find.gif"" class=""Btn"" title=""���� / �滻"" oncontextmenu=""SearchStr('" & Domain & "');return false;"" onClick=""SearchStr('" & Domain & "');"" ></td>"
'��������ť
ButtonArr(1,4)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/calculator.gif"" class=""Btn"" title=""������"" oncontextmenu=""Calculator('" & Domain & "');return false;"" onClick=""Calculator('" & Domain & "')"" ></td>"	
'���뵱ǰ����
ButtonArr(1,5)="<td width=""30"" align=""center""><img title=""���뵱ǰ����"" oncontextmenu=""InsertDate();return false;"" onClick=""InsertDate()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/date.gif"" ></td>"
'���뵱ǰʱ��
ButtonArr(1,6)="<td width=""30"" align=""center""><img title=""���뵱ǰʱ��"" oncontextmenu=""InsertTime();return false;"" onClick=""InsertTime()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/time.gif"" ></td>"
'ɾ������HTML��ʶ
ButtonArr(1,7)="<td width=""30"" align=""center""><img title=""ɾ������HTML��ʶ"" oncontextmenu=""DelAllHtmlTag();return false;"" onClick=""DelAllHtmlTag()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/geshi.gif"" ></td>"
'ɾ�����ָ�ʽ
ButtonArr(1,8)="<td width=""30"" align=""center""><img title=""ɾ�����ָ�ʽ"" oncontextmenu=""Format('removeformat');return false;"" onClick=""Format('removeformat')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/clear.gif"" ></td>"
'���볬������
ButtonArr(1,9)="<td width=""30"" align=""center""><img title=""���볬������"" oncontextmenu=""Format('CreateLink');return false;"" onClick=""Format('CreateLink')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/url.gif"" ></td>"
'ȡ����������
ButtonArr(1,10)="<td width=""30"" align=""center""><img title=""ȡ����������"" oncontextmenu=""Format('unLink');return false;"" onClick=""Format('unLink')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/nourl.gif"" ></td>"
'������ҳ
ButtonArr(1,11)="<td width=""30"" align=""center""><img title=""������ҳ"" oncontextmenu=""InsertPage('" & Domain & "');return false;"" onClick=""InsertPage('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/htm.gif"" ></td>"
'������Ŀ��
ButtonArr(1,12)="<td width=""30"" align=""center""><img title=""������Ŀ��"" oncontextmenu=""InsertFrame('" & Domain & "');return false;"" onClick=""InsertFrame('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/fieldset.gif"" ></td>"
'����Excel���
ButtonArr(1,13)="<td width=""30"" align=""center""><img title=""����Excel���"" oncontextmenu=""InsertExcel();return false;"" onClick=""InsertExcel()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/excel.gif"" ></td>"
'��������ı�
ButtonArr(1,14)="<td width=""30"" align=""center""><img title=""��������ı�"" oncontextmenu=""InsertMarquee('" & Domain & "');return false;"" onClick=""InsertMarquee('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Marquee.gif"" ></td>"
'ͼ�Ĳ���
ButtonArr(1,15)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/PicAlign.gif"" width=""23"" height=""22"" class=""Btn"" title=""ͼ�Ĳ���"" oncontextmenu=""PicAndTextArrange('" & Domain & "');return false;"" onClick=""PicAndTextArrange('" & Domain & "')"" ></td>"
'������
ButtonArr(1,16)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/Inserttable.gif""  class=""Btn"" title=""������"" oncontextmenu=""InsertTable('" & Domain & "');return false;""  onClick=""InsertTable('" & Domain & "')""></td>"
'������
ButtonArr(1,17)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/inserttable1.gif""  class=""Btn"" title=""������"" oncontextmenu=""InsertRow();return false;""  onClick=""InsertRow()""></td>"
'������
ButtonArr(1,18)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/inserttablec.gif"" class=""Btn"" title=""������"" oncontextmenu=""InsertColumn();return false;"" onClick=""InsertColumn()""></td>"
'ɾ����
ButtonArr(1,19)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/delinserttablec.gif""  class=""Btn"" title=""ɾ����"" oncontextmenu=""DeleteColumn();return false;"" onClick=""DeleteColumn()""></td>"
'���뵥Ԫ��
ButtonArr(1,20)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/insterttable2.gif"" width=""23"" height=""22""  class=""Btn"" title=""���뵥Ԫ��"" oncontextmenu=""InsertCell();return false;"" onClick=""InsertCell()""></td>"
'ɾ����Ԫ��
ButtonArr(1,21)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/delinsterttable2.gif"" width=""23"" height=""22""  class=""Btn"" title=""ɾ����Ԫ��"" oncontextmenu=""DeleteCell();return false;"" onClick=""DeleteCell()""></td>"
'�����
ButtonArr(1,22)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/SplitTD.gif"" width=""23"" height=""22""  class=""Btn"" title=""�����"" oncontextmenu=""SplitColumn();return false;"" onClick=""SplitColumn()""></td>"
'�ϲ���
ButtonArr(1,23)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/MargeTD.gif"" width=""23"" height=""22""  class=""Btn"" title=""�ϲ���"" oncontextmenu=""MergeColumn();return false;"" onClick=""MergeColumn()""></td>"
'�ϲ���
ButtonArr(1,24)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/Hbtable.gif"" width=""23"" height=""22""  class=""Btn"" title=""�ϲ���"" oncontextmenu=""MergeRow();return false;"" onClick=""MergeRow()""></td>"
'�����
ButtonArr(1,25)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/cftable.gif""  class=""Btn"" title=""�����"" oncontextmenu=""SplitRows();return false;"" onClick=""SplitRows()""></td>"
'ǰ̨��ҳ��ť
ButtonArr(1,26)="<td width=""90"" align=""center""><a title='�����ҳ����' oncontextmenu='InsertMorePage();return false;' onClick='InsertMorePage()' class='Btn'><img src=""" & Domain & "KS_Editor/images/NewDoc.gif""></a></td>"

'������ֺ�
ButtonArr(2,1)="<td width=""26"" align=""center""> <select name=""select2"" style='width:60px' class=""ToolSelectStyle"" onChange=""Format('fontname',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>����</option><option value=""����"">����</option><option value=""����"">����</option><option value=""����_GB2312"">����</option><option value=""����_GB2312"">����</option><option value=""����"">����</option><option value=""��Բ"">��Բ</option><option value=""Arial"">Arial</option><option value=""Arial Black"">Arial Black</option><option value=""Arial Narrow"">Arial Narrow</option><option value=""Brush Script	MT"">Brush Script MT</option><option value=""Century Gothic"">Century Gothic</option><option value=""Comic Sans MS"">Comic Sans MS</option><option value=""Courier"">Courier</option><option value=""Courier New"">Courier New</option><option value=""MS Sans Serif"">MS Sans Serif</option><option value=""Script"">Script</option><option value=""System"">System</option><option value=""Times New Roman"">Times New Roman</option><option value=""Verdana"">Verdana</option><option value=""Wide Latin"">Wide Latin</option><option value=""Wingdings"">Wingdings</option></select></td><td width=""26"" align=""center""> <select name=""select3"" style=""width:48px;height:18px"" onChange=""Format('fontsize',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>�ֺ�</option><option value=""7"">һ��</option><option value=""6"">����</option><option value=""5"">����</option><option value=""4"">�ĺ�</option><option value=""3"">���</option><option value=""2"">����</option><option value=""1"">�ߺ�</option></select></td>"
'�ϱꡢ�±ꡢ�Ӵ֡�б�塢�»��ߺ�ɾ����
ButtonArr(2,2)="<td width=""26"" align=""center""><img title=""�ϱ�"" oncontextmenu=""Format('superscript');return false;"" onClick=""Format('superscript')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/superscript.gif"" ></td><td width=""26"" align=""center""><img title=""�±�"" oncontextmenu=""Format('subscript');return false;"" onClick=""Format('subscript')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/subscript.gif"" ></td><td width=""26"" align=""center""><img title=""�Ӵ�"" oncontextmenu=""Format('bold');return false;"" onClick=""Format('bold')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/bold.gif"" ></td><td width=""26"" align=""center""><img title=""б��"" oncontextmenu=""Format('italic');return false;"" onClick=""Format('italic')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/italic.gif"" ></td><td width=""26"" align=""center""><img title=""�»���"" oncontextmenu=""Format('underline');return false;"" onClick=""Format('underline')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/underline.gif"" ></td><td  width=""26"" align=""center""><img title=""ɾ����"" oncontextmenu=""Format('StrikeThrough');return false;"" onClick=""Format('StrikeThrough')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/strikethrough.gif"" width=""20"" height=""20""></td>"

'������ɫ
ButtonArr(2,3)="<td  width=""26"" align=""center""><img src=""" & Domain & "KS_Editor/images/TextColor.gif"" class=""Btn"" title=""������ɫ"" oncontextmenu=""TextColor('" & Domain & "');return false;"" onClick=""TextColor('" & Domain & "')"" ></td>"
'���ֱ���ɫ
ButtonArr(2,4)="<td  width=""26"" align=""center""><img title=""���ֱ���ɫ"" oncontextmenu=""TextBGColor('" & Domain & "');return false;"" onClick=""TextBGColor('" & Domain & "')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/fgbgcolor.gif"" ></td>"
'����롢���С��Ҷ���
ButtonArr(2,5)="<td width=""26"" align=""center""><img title=""�����"" oncontextmenu=""Format('justifyleft');return false;"" onClick=""Format('justifyleft')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Aleft.gif"" ></td><td width=""26"" align=""center""><img title=""����"" oncontextmenu=""Format('justifycenter');return false;"" onClick=""Format('justifycenter')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Acenter.gif"" ></td><td width=""23"" align=""center""><img title=""�Ҷ���"" oncontextmenu=""Format('justifyright');return false;"" onClick=""Format('justifyright')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/Aright.gif"" ></td>"
'����������������������
ButtonArr(2,6)="<td width=""23"" align=""center""><img title=""����������"" oncontextmenu=""Format('outdent');return false;"" onClick=""Format('outdent');"" class=""Btn"" src=""" & Domain & "KS_Editor/images/outdent.gif"" ></td><td width=""30"" align=""center""><img title=""����������"" oncontextmenu=""Format('indent');return false;"" onClick=""Format('indent')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/indent.gif"" ></td>"
'������ʽ
ButtonArr(2,7)="<td width=""30"" align=""center""><select name=""select"" style=""width:80px;height:18px"" onChange=""Format('FormatBlock',this[this.selectedIndex].value);this.selectedIndex=0;KS_EditArea.focus();""><option selected>������ʽ</option><option value=""&lt;P&gt;"">��ͨ</option><option value=""&lt;H1&gt;"">����һ</option><option value=""&lt;H2&gt;"">�����</option><option value=""&lt;H3&gt;"">������</option><option value=""&lt;H4&gt;"">������</option><option value=""&lt;H5&gt;"">������</option><option value=""&lt;H6&gt;"">������</option><option value=""&lt;p&gt;"">����</option><option value=""&lt;dd&gt;"">����</option><option value=""&lt;dt&gt;"">���ﶨ��</option><option value=""&lt;dir&gt;"">Ŀ¼�б�</option><option value=""&lt;menu&gt;"">�˵��б�</option><option value=""&lt;PRE&gt;"">�ѱ��Ÿ�ʽ</option></select></td>"
'��Ŀ����
ButtonArr(2,8)="<td width=""30"" align=""center""><img title=""��Ŀ����"" oncontextmenu=""Format('insertunorderedlist');return false;"" onClick=""Format('insertunorderedlist')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/list.gif"" ></td>"
'���
ButtonArr(2,9)="<td width=""30"" align=""center""><img title=""���"" oncontextmenu=""Format('insertorderedlist');return false;"" onClick=""Format('insertorderedlist')"" class=""Btn"" src=""" & Domain & "KS_Editor/images/num.gif"" ></td>"
'���Ի����λ��
ButtonArr(2,10)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/abspos.gif"" class=""Btn"" title=""���Ի����λ��"" oncontextmenu=""Pos();return false;"" onClick=""Pos();""></td>"
'�����������
ButtonArr(2,11)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/symbol.gif"" class=""Btn"" title=""�����������"" oncontextmenu=""InsertSymbol('" & Domain & "');return false;"" onClick=""InsertSymbol('" & Domain & "');""></td>"
'����ͼƬ
ButtonArr(2,12)="<td width=""30"" align=""center""><img title=""����ͼƬ��֧�ָ�ʽΪ��jpg��gif��bmp��png��"" oncontextmenu=""InsertPicture(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertPicture(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/img.gif"" ></td>"
'����flash��ý���ļ�
ButtonArr(2,13)="<td width=""30"" align=""center""><img title=""����flash��ý���ļ�"" oncontextmenu=""InsertFlash(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertFlash(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/flash.gif"" ></td>"
'������Ƶ�ļ�
ButtonArr(2,14)="<td width=""30"" align=""center""><img title=""������Ƶ�ļ���֧�ָ�ʽΪ��avi��wmv��asf��mpg"" oncontextmenu=""InsertVideo(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ");return false;"" onClick=""InsertVideo(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/wmv.gif"" ></td>"
'����RealPlay�ļ�
ButtonArr(2,15)="<td width=""30"" align=""center""><img title=""����RealPlay�ļ���֧�ָ�ʽΪ��rm��ra��ram"" oncontextmenu=""InsertRM(" & EditorFromType & ",'" & Domain & "',"& ChannelID & ");return false;"" onClick=""InsertRM(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/rm.gif"" ></td>"
'��������ˮƽ��
ButtonArr(2,16)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/sline.gif"" class=""Btn"" title=""��������ˮƽ��"" oncontextmenu=""SpecialHR('" & Domain & "');return false;"" onClick=""SpecialHR('" & Domain & "')"" ></td>"
'������ͨˮƽ��
ButtonArr(2,17)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/line.gif"" class=""Btn"" title=""������ͨˮƽ��"" oncontextmenu=""InsertHR();return false;"" onClick=""InsertHR();"" ></td>"
'����ճ��
ButtonArr(2,18)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/paste.gif"" class=""Btn"" title=""����ճ��"" oncontextmenu=""Paste();return false;"" onClick=""Paste()"" ></td>"
'���뻻�з���
ButtonArr(2,19)="<td width=""30"" align=""center""><img title=""���뻻�з���"" oncontextmenu=""InsertBR();return false;"" onClick=""InsertBR()"" class=""Btn"" src=""" & Domain & "KS_Editor/images/chars.gif"" ></td>"
'�ı�ճ��
ButtonArr(2,30)="<td width=""30"" align=""center""><img src=""" & Domain & "KS_Editor/images/PasteText.gif"" class=""Btn"" title=""�ı�ճ��"" oncontextmenu=""PasteText();return false;"" onClick=""PasteText()"" ></td>"

dim fb
if buttonstyle=1 then
 fb="plugin.swf"
else
 fb="plugin_" & buttonstyle & ".swf"
end if

ButtonArr(2,31)="<td><!-- cc��Ƶ������� --><object width='72' height='24'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin/" & fb & "?userID=" & userid &"&type=normal' /><embed src='http://union.bokecc.com/flash/plugin/" & fb & "?userID=" & userid & "&type=normal' type='application/x-shockwave-flash' width='72' height='24' allowScriptAccess='always' wmode='transparent'></embed></object><!-- cc��Ƶ������� --></td>"

'���ñ�ǩ�б�
 Dim MyLabelStr
		 MyLabelStr=" <select name=""mylabel"" style=""width:160px"">"
		 MyLabelStr=MyLabelStr & " <option value="""">==ѡ��ϵͳ������ǩ==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 MyLabelStr=MyLabelStr & "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  MyLabelStr=MyLabelStr & "</select>&nbsp;<input type='button' class='tdbg' onclick='LabelInsertCode(document.all.mylabel.value);' value='�����ǩ'>&nbsp;"
		  RS.Close:Set RS=Nothing
ButtonArr(2,20)="<td>&nbsp;" &MyLabelStr & "<input type=""button"" class='tdbg' onclick=""javascript:WapLabelInsertTemplate();"" value=""WAP��ǩ"">&nbsp;<input type=""button"" class='tdbg' onclick=""javascript:LabelInsertTemplate();"" value=""ѡ������ǩ""></td>"

'���븽���ļ�
ButtonArr(2,21)="<td width=""30"" align=""center""><img title=""���븽���ļ���֧�ָ�ʽΪ��rar��zip��txt��doc��xls"" oncontextmenu=""InsertUpFile(" & EditorFromType & ",'" & Domain & "',"& ChannelID & ");return false;"" onClick=""InsertUpFile(" & EditorFromType & ",'" & Domain & "'," & ChannelID & ")"" class=""Btn"" src=""" & Domain & "KS_Editor/images/down.gif"" ></td>"
End Sub
End Class
%> 

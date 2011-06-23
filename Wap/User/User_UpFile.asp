<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New User_Upfile
KSCls.Kesion()
Set KSCls = Nothing

Class User_Upfile
        Private KS,ChannelID,BasicType
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		     %>
             <?xml version="1.0" encoding="utf-8"?>
             <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
             <html xmlns="http://www.w3.org/1999/xhtml">
             <head>
             <title>WAP2.0上传</title>
             <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
             <link href="../Images/style.css" rel="stylesheet" type="text/css" media="screen">
             </head>
             <body>
             <%
			 IF Cbool(KSUser.UserLoginChecked)=false Then
			    Response.Write "登录后才可以上传!<br/>"
				Exit Sub
			 End If
			 
			 ChannelID=KS.ChkClng(KS.S("ChannelID"))
			 If ChannelID=999 Then
			    BasicType=ChannelID
			 ElseIf ChannelID<5000 Then
			    BasicType=KS.C_S(ChannelID,6)
			 Else
			    BasicType=ChannelID
			 End If
			 
			 If KS.S("Type")="Field" Then
			    Call User_Field_UpForm
			 Else
			    Select Case BasicType
				    Case 9999:Call User_face_UpForm'用户头像
					Case 9998:Call User_XPFM_UpForm'相册封面
					Case 9997:Call User_ZP_UpForm'相册照片
					Case 9996:Call User_Team_UpForm'圈子图片
					Case 9995:Call User_Mp3_UpForm'上传mp3
					Case 999:Call User_UpForm()
					Case 1
					   If KS.S("Type")="File" Then
					      Call User_Article_UpFile
					   Else
					      Call User_Article_UpForm
					   End IF
				    Case 2
				       If KS.S("Type")="Single" Then
					      Call User_Single_UpForm
					   Else
					      Call User_Picture_UpForm
					   End If
				    Case 3
				       If KS.S("Type")="Pic" Then
					      Call User_Down_Photo_UpForm
					   Else
					      Call User_Down_File_UpForm
					   End If
				    Case 4
				       If KS.S("Type")="Pic" Then
					      Call User_Flash_Photo_UpForm
					   Else
					      Call User_Flash_File_UpForm
					   End If
					Case 5
					      Call User_Shop_UpForm
			        Case 7
				       If KS.S("Type")="Pic" Then
					      Call User_Movie_Photo_UpForm
					   Else
					      Call User_Movie_File_UpForm
					   End If
			        Case 8
				    Call User_GQPhoto_UpForm
			    End Select
		    End If
			%>
            <div id="bottom-menu"><a href="Index.asp?<%=KS.WapValue%>" class="white">我的地盘</a> | <a href="<%=KS.GetGoBackIndex%>" class="white">返回首页</a></div>
            </body>
            </html>
            <%
		End Sub
		
		Sub User_Field_UpForm()
		    %>
            <div id="menu" class="white">上传文件</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">
            <form name="UpFileForm" method="post" enctype="multipart/form-data" Action="<%=KS.S("PrevUrl")%>&amp;ID=<%=KS.S("ID")%>&amp;ChannelID=<%=ChannelID%>&amp;UpFileChecked=1&amp;<%=KS.WapValue%>">
            上传：<input type="file" name="File1" class="textbox">
            <input type="submit" name="Submit" value="开始上传" class="button">
            <br/>
            <input name="BasicType" value="<%=BasicType%>" type="hidden">
            <input name="ChannelID" value="<%=ChannelID%>" type="hidden">
            <input name="FieldID" value="<%=KS.S("FieldID")%>" type="hidden">
            <input name="Type" value="Field" type="hidden">
            <input type="hidden" name="AutoReName" value="4">
            <input type="hidden" name="DefaultUrl" value="1">
            <input type="hidden" name="DefaultUrl" value="1">
            </form>
            </div>
            <div class="td-style1">
            <%
			Dim RS
			If ChannelID=0 Then
			Set RS=Conn.Execute("Select FieldName,AllowFileExt,MaxFileSize From KS_FormField Where FieldID=" & KS.ChkClng(KS.S("FieldID")))
			Else
			Set RS=Conn.Execute("Select FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(KS.S("FieldID")))
			End If
			If Not RS.Eof Then
			   Response.Write "允许上传的类型:"&RS(1)&" 最大只能上传:"&RS(2)&"K"
			End If
			RS.Close:Set RS=Nothing
			%>
            </div>
            <%
		End Sub
		
		'用户头像
		Sub User_Face_UpForm()
		    %>
            <div id="menu" class="white">上传照片</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">
            <form name="UpFileForm" method="post" enctype="multipart/form-data" action="User_Face.asp?Action=UpSaveUrl&amp;<%=KS.WapValue%>">
            <input type="file" name="File1" class="textbox">
            <input type="submit" id="BtnSubmit" class="button" name="Submit" value="开始上传">
            <input name="BasicType" value="<%=BasicType%>" type="hidden">
            <input type="hidden" name="AutoReName" value="4">
            <input type="hidden" name="DefaultUrl" value="1">
            <input type="hidden" name="AddWaterFlag" value="1">
            </form>
            </div>
            <div class="td-style1">*只支持jpg|gif|png。</div>
			<%
		End Sub
        '相册封面
		Sub User_XPFM_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  class=""button"" name=""Submit"" value=""开始上传"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			'.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
			'.Write "          缩略图"
			'.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'圈子图片
		Sub User_Team_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" onchange=""parent.document.all.showimages.src=this.value"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  class=""button"" name=""Submit"" value=""开始上传"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			'.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
			'.Write "          缩略图"
			'.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'Mp3
		Sub User_Mp3_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" accept=""html"" size=""20"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  class=""button"" name=""Submit"" value=""开始上传"">只支持MP3"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End SUb
        '照片
		Sub User_ZP_UpForm()
			With Response
			.Write "<div align=""center"">"
			.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td width=""82%"" valign=""top"">"
			.Write "          <div align=""center"">"
			.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "              <tr>"
			.Write "                <td width=""50%"" height=""50""> &nbsp;&nbsp;设定照片数量"
			.Write "                  <input class='textbox' name=""UpFileNum"" type=""text"" value=""5"" size=""5"" style=""text-align:center"">"
			.Write "                <input type=""button"" name=""Submit42"" class=""button"" value=""确定设定"" onClick=""ChooseOption();""></td>"
			.Write "                <td width=""50%"" id='ss'>&nbsp;</td>"
			.Write "              </tr>"
			.Write "              <tr>"
			.Write "                <td height=""30"" colspan=""2"" id=""FilesList""> </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </div></td>"
			.Write "        <td width=""18%"" valign=""top"">"
			.Write "          <input name=""AddWaterFlag""  style=""display:none"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "<!--添加水印-->"
			.Write "        </td>"
			.Write "      </tr>"
			.Write "      <tr>"
			.Write "        <td  colspan=""2""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "            <tr>"
			.Write "              <td align=""center"">"
			.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  name=""Submit"" value=""开始上传""  class=""button"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "                  <input type=""reset"" id=""ResetForm""  class=""button"" name=""Submit3"" value="" 重 填 "">"
			.Write "              </td>"
			.Write "            </tr>"
			.Write "          </table><font color=red>说明：只支持jpg、gif、png，小于100k的照片。</font></td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			.Write "</div>"
			.Write "<script language=""JavaScript""> " & vbCrLf
			.Write "function ChooseOption()" & vbCrLf
			.Write "{"
			.Write "  var UpFileNum = document.all.UpFileNum.value;" & vbCrLf
			.Write "  if (UpFileNum=='') " & vbCrLf
			.Write "    UpFileNum=12;" & vbCrLf
			.Write "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
			.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
			.Write "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
			.Write "   { " & vbCrLf
			.Write "    for (i=0;i<2;i++)" & vbCrLf
			.Write "      { n=n+1;" & vbCrLf
			.Write "       Optionstr = Optionstr+'<tr><td>&nbsp;照&nbsp;片&nbsp;'+n+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""35"" class=""textbox"" onchange=""parent.document.all.view'+n+'.src=this.value;"" name=""File'+n+'"">&nbsp;</td></tr>';" & vbCrLf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.Write "       }" & vbCrLf
			.Write "      Optionstr = Optionstr+''" & vbCrLf
			.Write "  }" & vbCrLf
			.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
			.Write "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf

			.Write "  var viewstr;"&vbcrlf
			.write "   n=0;"&vbcrlf
			.Write "   viewstr=""<table width='100%' border='0'>"";" &vbcrlf
			.write "   for(k=0;k<UpFileNum/5;k++)" & vbcrlf
			.write "    {" & vbcrlf
			.Write "     viewstr=viewstr+""<tr>"";" & vbcrlf
			.write "     for(i=0;i<5;i++)" & vbcrlf
			.write "      {" &vbcrlf
			.write "         n=n+1;"&vbcrlf
			.write "        viewstr=viewstr+""<TD width='20%'><table style='BORDER-COLLAPSE: collapse' borderColor='#c0c0c0' cellSpacing='1' cellPadding='2' border='1'><tr><td align='center' width='83' height='100'><img name='view""+n+""' src='images/view.gif' title='照片预览' width='110' height='90'></td></tr></table></TD>"";"&vbcrlf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.write "       }"&vbcrlf
			.write "    for(i=n;i<5;i++)" & vbcrlf
			.write "     { viewstr=viewstr+""<td></td>"";}"
			.write "   viewstr=viewstr+""</tr>"";"&vbcrlf
			.Write "    }" & vbcrlf
			.write "  viewstr=viewstr+""</table>"";"&vbcrlf
			if KS.S("action")<>"OK" then	.write "parent.document.all.viewarea.innerHTML=viewstr;"& vbcrlf
			.write "parent.init();"
			.Write "parent.resize_mainframe();"
			.write "parent.parent.init();"
			.Write "parent.parent.resize_mainframe();"
			.Write " }" & vbCrLf
			.Write "ChooseOption();" & vbCrLf
			.Write "</script>" & vbCrLf
			End With
		End Sub
		
		Sub User_UpForm()
			With Response
			.Write "<div align=""center"">"
			.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td width=""82%"" valign=""top"">"
			.Write "          <div align=""center"">"
			.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "              <tr>"
			.Write "                <td width=""50%"" height=""50""> &nbsp;&nbsp;设定文件数"
			.Write "                  <input class='textbox' name=""UpFileNum"" type=""text"" value=""3"" size=""5"" style=""text-align:center"">"
			.Write "                <input type=""button"" name=""Submit42"" class=""button"" value=""确定设定"" onClick=""ChooseOption();""></td>"
			.Write "                <td width=""50%"" id='ss'><input type=""checkbox"" name=""AutoReName"" value=""4"">自动命名</td>"
			.Write "              </tr>"
			.Write "              <tr>"
			.Write "                <td height=""30"" colspan=""2"" id=""FilesList""> </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </div></td>"
			.Write "        <td width=""18%"" valign=""top"">"
			.Write "          <input name=""AddWaterFlag""  style=""display:none"" type=""checkbox"" id=""AddWaterFlag"" value=""1"">"
			.Write "<!--添加水印-->"
			.Write "        </td>"
			.Write "      </tr>"
			.Write "      <tr>"
			.Write "        <td  colspan=""2""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "            <tr>"
			.Write "              <td align=""center"">"
			.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  name=""Submit"" value=""开始上传""  class=""button"">"
		    .Write "          <input name=""BasicType"" value=""999"" type=""hidden"">"
			.Write "         <input type=""reset"" id=""ResetForm""  class=""button"" name=""Submit3"" value="" 重 填 "">"
			.Write "              </td>"
			.Write "            </tr>"
			.Write "          </table><font color=red>温馨提示：您只能上传jpg,gif,png,swf格式的文件。</font></td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			.Write "</div>"
			.Write "<script language=""JavaScript""> " & vbCrLf
			.Write "function ChooseOption()" & vbCrLf
			.Write "{"
			.Write "  var UpFileNum = document.all.UpFileNum.value;" & vbCrLf
			.Write "  if (UpFileNum=='') " & vbCrLf
			.Write "    UpFileNum=12;" & vbCrLf
			.Write "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
			.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
			.Write "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
			.Write "   { " & vbCrLf
			.Write "    for (i=0;i<2;i++)" & vbCrLf
			.Write "      { n=n+1;" & vbCrLf
			.Write "       Optionstr = Optionstr+'<tr><td>&nbsp;照&nbsp;片&nbsp;'+n+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""35"" class=""textbox"" name=""File'+n+'"">&nbsp;</td></tr>';" & vbCrLf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.Write "       }" & vbCrLf
			.Write "      Optionstr = Optionstr+''" & vbCrLf
			.Write "  }" & vbCrLf
			.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
			.Write "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf
			.Write "}ChooseOption();" & vbCrLf
			.Write "</script>" & vbCrLf
			End With
		End Sub
		
		Sub User_Article_UpForm()
		    %>
            <div id="menu" class="white">上传图片</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">
            <form name="UpFileForm" method="post" enctype="multipart/form-data" action="User_MyArticle.asp?Action=<%=KS.S("Action")%>&amp;ID=<%=KS.S("ID")%>&amp;ChannelID=<%=ChannelID%>&amp;UpFileChecked=1&amp;<%=KS.WapValue%>">
            <input type="file" name="File1" class='textbox'>
            <input type="submit" name="Submit" value="开始上传" class="button">
            <input type="hidden" name="BasicType" value="<%=BasicType%>">
            <input type="hidden" name="ChannelID" value="<%=ChannelID%>">
            <input type="hidden" name="AutoReName" value="4">
            <input type="hidden" name="DefaultUrl" value="4">
            <input type="hidden" name="AddWaterFlag" value="4">
            </form>
            </div>
            <div class="td-style1">*只支持jpg,gif,png。</div>
            <%
		End Sub
		'上传附件
		Sub User_Article_UpFile()

			Response.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			Response.Write "      <tr>"
			Response.Write "        <td width=""82%"" valign=""top"">"
			Response.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write "<tr><td><div id=""FilesList""></div></td><td><input type=""submit"" id=""BtnSubmit"" class='button' name=""Submit"" value='上传' onclick=""return(parent.CheckClassID())"">  设置个数<input name=""UpFileNum"" style=""text-align:center"" type=""text"" value=""1"" size=""3""><input type=""button"" name=""Submit42"" value=""设定"" class='button' onClick=""ChooseOption();""></td></tr>" & vbcrlf
			Response.Write "              <tr>"
			Response.Write "                <td height=""30"" align=""center""> "
			Response.Write "                  <input type=""hidden"" name=""AutoReName"" value=""0"">"
			Response.Write "          <input name=""Type"" value=""File"" type=""hidden"">"
			Response.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			Response.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			Response.Write "</td>"
			Response.Write "              </tr>"
			Response.Write "            </table>"
			Response.Write "        </td>"
			Response.Write "      </tr>"
			Response.Write "    </form>"
			Response.Write "  </table>"
			Response.Write "</div>"
			Response.Write "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left: 112px; top: 28px; background-color: #00CCFF; layer-background-color: #00CCFF; border: 1px none #000000; width: 254px; height: 63px; visibility: hidden;"">"
			Response.Write "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "    <tr>"
			Response.Write "      <td><div align=""right"">请稍等，正在上传文件</div></td>"
			Response.Write "      <td width=""35%""><div align=""left""><font id=""ShowArticleArea"" size=""+1""></font></div></td>"
			Response.Write "    </tr>"
			Response.Write "  </table>"
			Response.Write "</div>"
			Response.Write "<script language=""JavaScript"">" & vbCrLf
			Response.Write "function ChooseOption()" & vbCrLf
			Response.Write " {" & vbCrLf
			Response.Write "  var UpFileNum = document.all.UpFileNum.value;" & vbCrLf
			Response.Write "  if (UpFileNum=='')" & vbCrLf
			Response.Write "    UpFileNum=1;" & vbCrLf
			 Response.Write " var k,i,Optionstr,n=0;" & vbCrLf
			 Response.Write "     Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';"
			 Response.Write " for(k=0;k<(UpFileNum);k++)" & vbCrLf & vbCrLf
			 Response.Write "  { Optionstr = Optionstr+'<tr>';" & vbCrLf
			Response.Write "       n=n+1;" & vbCrLf
			 Response.Write "      if(UpFileNum==1)"
			 Response.Write "       Optionstr = Optionstr+'<td>附&nbsp;件：</td><td>&nbsp;<input type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" class=""textbox"">&nbsp;</td>';" & vbCrLf
			 Response.Write "      else"
			 Response.Write "       Optionstr = Optionstr+'<td>&nbsp;附&nbsp;件&nbsp;'+n+'</td><td>&nbsp;<input class=""textbox"" type=""file"" accept=""html"" size=""30"" name=""File'+n+'"">&nbsp;</td>';" & vbCrLf
			Response.Write "        if (n==UpFileNum) break;" & vbCrLf
			Response.Write "      Optionstr = Optionstr+'</tr>'" & vbCrLf
			Response.Write "  }" & vbCrLf
			Response.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
			Response.Write "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf
			Response.Write "parent.document.getElementById('UpFileFrame').height=parseInt(document.all.UpFileNum.value)*28;" & vbcrlf
			Response.Write " }" & vbCrLf
			Response.Write "ChooseOption();" & vbCrLf
			Response.Write "</script>"
		
		
		 response.write "<script>parent.document.getElementById('UpFileFrame').height=30;</script>"
		End Sub
		
	 Sub User_Picture_UpForm()
			With Response
		.Write "<div align=""center"">"
		.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <tr class='clefttitle'><td height='25' align=center><strong>批 量 上 传</strong></td></tr>"
		.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
		.Write "      <tr>"
		.Write "        <td width=""82%"" valign=""top"">"
		.Write "          <div align=""center"">"
		.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "              <tr>"
		.Write "                <td height=""30"" colspan=""3"" id=""FilesList""> </td>"
		.Write "              </tr>"
		.Write "              <tr>"
		.Write "                <td align='right'>"
		.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""开始上传""  onclick=""return(parent.CheckClassID())"">"
		.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
		.Write "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" 重 填 "">"
		.Write "        </td>"
		.Write "                <td width=""45%"" height=""25"" id='ss' align='right'>"
		.Write "                <td width=""20%""><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>添加水印</td>"
		.Write "              </tr>"
		.Write "            </table>"
		.Write "        </div></td>"
		.Write "    </form>"
		.Write "  </table>"
		.Write "</div>"
			.Write "<script language=""JavaScript""> " & vbCrLf
		%>
		function ViewPic(nid,f){
		if ( f != "" ) {
		  var num=parent.document.all.picnum.value;
		  parent.document.getElementByID('picview'+nid).src=f;
		 }
		}
		<%
		.Write "function ChooseOption(num)" & vbCrLf
		.Write "{"
		.Write "  var UpFileNum = num;" & vbCrLf
		.Write "  if (UpFileNum=='') " & vbCrLf
		.Write "    UpFileNum=10;" & vbCrLf
		.Write "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
		.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		.Write "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
		.Write "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
		.Write "    for (i=0;i<2;i++)" & vbCrLf
		.Write "      { n=n+1;" & vbCrLf
		.Write "       Optionstr = Optionstr+'<td>&nbsp;第&nbsp;'+n+'&nbsp;张</td><td>&nbsp;<input type=""file"" accept=""html"" size=""25"" name=""File'+n+'"" nid=""'+n+'"" class=""textbox"" onchange=""ViewPic(this.nid,this.value)"">&nbsp;</td>';" & vbCrLf
		.Write "        if (n==UpFileNum) break;" & vbCrLf
		.Write "       }" & vbCrLf
		.Write "      while (i <= 2)" & vbCrLf
		.Write "      {" & vbCrLf
		.Write "      Optionstr = Optionstr+'<td width=""50%"">&nbsp; </td>';" & vbCrLf
		.Write "      i++;" & vbCrLf
		.Write "      }" & vbCrLf
		.Write "      Optionstr = Optionstr+'</tr>'" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
		.Write "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf
		.Write "    SelectOptionstr='设定第<select class=""upfile"" name=""DefaultUrl"">'" & vbCrLf
		.Write " for(i=1;i<=UpFileNum;++i)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "   SelectOptionstr=SelectOptionstr+'<option value=""'+i+'"">'+i+'</option>'" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "   SelectOptionstr=SelectOptionstr+'</select>张为缩略图(系统自动生成)'" & vbCrLf
		.Write "   document.all.ss.innerHTML=SelectOptionstr;" & vbCrLf
		.Write " }" & vbCrLf
		.Write "ChooseOption(4);" & vbCrLf
		.Write "</script>" & vbCrLf
			End With
		End Sub
		
		'单张上传
		Sub User_Single_UpForm()
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          &nbsp;&nbsp;上传图片： <input type=""file"" class='textbox' accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  name=""Submit""  class=""button"" value=""开始上传"">"
			.Write "          <input name=""UpLoadFrom"" value=""21"" type=""hidden"" id=""UpLoadFrom"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""objid"" value=""" & request("objid") & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">同时生成缩略图</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "添加水印</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'上传下载缩略图
		Sub User_Down_Photo_UpForm()
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class='textbox' accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""return(parent.CheckClassID())""  name=""Submit""  class=""button"" value=""开始上传"">"
			.Write "          <input name=""UpLoadFrom"" value=""31"" type=""hidden"" id=""UpLoadFrom"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">同时生成缩略图</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "添加水印</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'上传下载文件
		Sub User_Down_File_UpForm
		With Response
		.Write "<div align=""center"">"
		.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
		.Write "      <tr>"
		.Write "        <td width=""82%"" valign=""top"">"
		.Write "            <input class=""textbox"" type=""file"" accept=""html"" size=""50"" name=""File1""> "
		.Write "                <input type=""hidden"" name=""AutoReName"" value=""4"" checked><input type=""submit"" id=""BtnSubmit""  name=""Submit"" onclick=""return(parent.CheckClassID())"" value=""开始上传 "" class=""button"">"
		.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
		.Write "<BR>          系统提供的上传功能只适合上传比较小的软件（如ASP源代码压缩包）。如果软件比较大（2M以上），请先使用FTP上传，而不要使用系统提供的上传功能，以免上传出错或过度占用服务器的CPU资源。"
		.Write "        </td>"
		.Write "      </tr>"
		.Write "    </form>"
		.Write "  </table>"
		.Write "</div>"
		.Write "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left: 112px; top: 28px; background-color: #00CCFF; layer-background-color: #00CCFF; border: 1px none #000000; width: 254px; height: 63px; visibility: hidden;"">"
		.Write "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "    <tr>"
		.Write "      <td><div align=""right"">请稍等，正在上传文件</div></td>"
		.Write "      <td width=""35%""><div align=""left""><font id=""ShowArticleArea"" size=""+1""></font></div></td>"
		.Write "    </tr>"
		.Write "  </table>"
		.Write "</div>"
		End With
		End Sub
		
		'上传动漫缩略图
		Sub User_Flash_Photo_UpForm
		   With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""return(parent.CheckClassID())"" name=""Submit"" class=""button"" value=""开始上传"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">同时生成缩略图</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "添加水印</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'上传动漫文件
		Sub User_Flash_File_UpForm
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""return(parent.CheckClassID())"" name=""Submit""  class=""button"" value=""开始上传"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		Sub User_Shop_UpForm
		   %>
		   <div id="menu" class="white">上传照片</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">

		   <%
		  With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  name=""Submit"" class=""button"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" value=""开始上传"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>生成缩略图"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "添加水印</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		
		
		'上传影片缩略图
		Sub User_Movie_Photo_UpForm
		   With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  name=""Submit"" class=""button"" onclick=""return(parent.CheckClassID())"" value=""开始上传"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">同时生成缩略图</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "添加水印</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'上传影片文件
		Sub User_Movie_File_UpForm
		    %>
            <div id="menu" class="white">上传影片文件</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">
            <form name="UpFileForm" method="post" enctype="multipart/form-data" action="UpFileSave.asp">
            <input type="file" class="textbox" accept="html" size="40" name="File1" class="textbox">
            <input type="submit" name="Submit"  class="button" value="开始上传">
            <input name="BasicType" value="<%=BasicType%>" type="hidden">
            <input name="ChannelID" value="<%=ChannelID%>" type="hidden">
            <input type="hidden" name="AutoReName" value="4">
            </form>
            </div>
            <div class="td-style1">*只支持jpg,gif,png。</div>
            <%
		End Sub
		'供求图片
		Sub User_GQPhoto_UpForm()
		    %>
            <div id="menu" class="white">上传供求图片</div>
            <div id="new-soft"><%=UCWEBAd%></div>
            <div class="list-box">
            <form name="UpFileForm" method="post" enctype="multipart/form-data" action="User_MySupply.asp?Action=<%=KS.S("Action")%>&amp;ID=<%=KS.S("ID")%>&amp;UpFileChecked=1&amp;<%=KS.WapValue%>">
            <input type="file" accept="html" name="File1" class="textbox">
            <input type="submit" class="button" name="Submit" value="开始上传">
            <br/>
            <input name="BasicType" value="<%=BasicType%>" type="hidden">
            <input name="ChannelID" value="<%=ChannelID%>" type="hidden">
            <input type="hidden" name="AutoReName" value="4">
            <input type="checkbox" name="DefaultUrl" value="1" checked="checked">缩略图片
            <input name="AddWaterFlag" type="checkbox" id="AddWaterFlag" value="1" checked="checked">添加水印
            </form>
            </div>
            <div class="td-style1">*只支持jpg,gif,png。</div>
            <%
		End Sub
		
		Function UCWEBAd()
		    Dim Url
		    Url=Replace(KS.Setting(2)&KS.GetUrl,"http://","")
			If Instr(Request.ServerVariables("HTTP_ACCEPT"),"ucweb")<5  Then
			   UCWEBAd="无法上传文件，请下载<a href=""http://down2.ucweb.com/download.asp?f=chenf@wapcr.cn&amp;Url="&Url&"Title="&KS.Setting(105)&""" style=""color:#ffffff;text-decoration:underline;"">·UCWEB浏览器v6.0</a>全新操作界面，超酷WAP/WEB网站推荐，为你带来页面浏览、文件下载前所未有的超爽感受！现火爆上线，速抢鲜体验！"
			End If
		End Function
End Class
%> 

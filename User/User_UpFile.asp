<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_Upfile
KSCls.Kesion()
Set KSCls = Nothing

Class User_Upfile
        Private KS,KSUser,ChannelID,BasicType
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.Write "<title>�ϴ��ļ�</title>"
		.Write "<link rel=""stylesheet"" href=""images/css.css"">"
		.Write "<style type=""text/css"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "body {"
		.Write "    margin-left: 0px; " & vbCrLf
		.Write "    margin-top: 0px;" & vbCrLf
		.Write "}" & vbCrLf
		.Write "-->" & vbCrLf
		.Write "</style></head>"
		.Write "<body class=tdbg style=""background-color:transparent"">"
		IF Cbool(KSUser.UserLoginChecked)=false Then
		 .write "<font size='2'>�Բ���,���ȵ�¼�����ʹ�ô˹���!</font>"
		 ' .Write "<script>top.location.href='Login';<//script>"
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
			 Case 9999  '�û�ͷ��
			   Call User_face_UpForm
			 Case 9998  '������
			   Call User_XPFM_UpForm
			 Case 9997  '�����Ƭ
			   Call User_ZP_UpForm
			 Case 9996  'Ȧ��ͼƬ
			   Call User_Team_UpForm
			 Case 9995  '�ϴ�mp3
			   Call User_Mp3_UpForm
			 case 9994  'С��̳�ϴ�����
			   Call User_Club_UpFile
			 case 9993  'д��־�����ϴ�����
			   Call User_Club_UpFile
			 Case 999
			   Call User_UpForm()
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
			 Case 9 
				  Call SJ_UpPhoto()
			End Select
		  End If
		 .Write "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left:2px; top: 0px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 28px; visibility: hidden;"">"
		 .Write "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .Write "    <tr>"
		 .Write "      <td><div>&nbsp;���Եȣ������ϴ��ļ�<img src='../images/default/wait.gif' align='absmiddle'></div></td>"
		 .Write "    </tr>"
		 .Write "  </table>"
		 .Write "</div>"
		End With
		End Sub
		
		Sub User_Field_UpForm()
		 	With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">�ϴ���<input type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';return(parent.CheckClassID())"" name=""Submit"" value=""��ʼ�ϴ�"" class=""button"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""FieldID"" value=""" & KS.S("FieldID") & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""Field"" type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <span style='display:none'><input type=""checkbox"" name=""DefaultUrl"" value=""1"">"
			.Write "          ����ͼ"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"">"
			.Write "���ˮӡ</span></td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'�û�ͷ��
		Sub User_Face_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
        '������
		Sub User_XPFM_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			'.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
			'.Write "          ����ͼ"
			'.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'Ȧ��ͼƬ
		Sub User_Team_UpForm()
			With Response
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" onchange=""parent.document.all.showimages.src=this.value"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			'.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
			'.Write "          ����ͼ"
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
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">ֻ֧��MP3"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End SUb
        '��Ƭ
		Sub User_ZP_UpForm()
			With Response
			.Write "<div align=""center"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <tr>"
			.Write "        <td width=""82%"" valign=""top"">"
			.Write "          <div align=""center"">"
			.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "              <tr>"
			.Write "                <td width=""50%"" height=""50""> &nbsp;&nbsp;�趨��Ƭ����"
			.Write "                  <input class='textbox' name=""UpFileNum"" id=""UpFileNum"" type=""text"" value=""5"" size=""5"" style=""text-align:center"">"
			.Write "                <input type=""button"" name=""Submit42"" class=""button"" value=""ȷ���趨"" onClick=""ChooseOption();""></td>"
			.Write "                <td width=""50%"" id='ss'>&nbsp;</td>"
			.Write "              </tr>"
			.Write "              <tr>"
			.Write "                <td height=""30"" colspan=""2"" id=""FilesList""> </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </div></td>"
			.Write "        <td width=""18%"" valign=""top"">"
			.Write "          <input name=""AddWaterFlag""  style=""display:none"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "<!--���ˮӡ-->"
			.Write "        </td>"
			.Write "      </tr>"
			.Write "      <tr>"
			.Write "        <td  colspan=""2""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "            <tr>"
			.Write "              <td align=""center"">"
			.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  name=""Submit"" value=""��ʼ�ϴ�"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"">"
		    .Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "                  <input type=""reset"" id=""ResetForm""  class=""button"" name=""Submit3"" value="" �� �� "">"
			.Write "              </td>"
			.Write "            </tr>"
			.Write "          </table><font color=red>˵����ֻ֧��jpg��gif��png��С��100k����Ƭ��</font></td>"
			.Write "      </tr>"
			.Write "  </table>"
			.Write "    </form>"
			.Write "</div>"
			.Write "<script language=""JavaScript""> " & vbCrLf
			.Write "function ViewPic(nid,f){" & vbcrlf
				.Write "if ( f != '' ) {" & vbcrlf
				.Write "  parent.document.getElementById('view'+nid).innerHTML='';" & vbcrlf
				.Write "  parent.document.getElementById('view'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
				.Write " }"
				.Write "}"
			.Write "function ChooseOption()" & vbCrLf
			.Write "{"
			.Write "  var UpFileNum = document.getElementById('UpFileNum').value;" & vbCrLf
			.Write "  if (UpFileNum=='') " & vbCrLf
			.Write "    UpFileNum=12;" & vbCrLf
			.Write "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
			.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
			.Write "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
			.Write "   { " & vbCrLf
			.Write "    for (i=0;i<2;i++)" & vbCrLf
			.Write "      { n=n+1;" & vbCrLf
			.Write "       Optionstr = Optionstr+'<tr><td>&nbsp;��&nbsp;Ƭ&nbsp;'+n+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""35"" class=""textbox"" onchange=""ViewPic('+n+',this.value)"" name=""File'+n+'"">&nbsp;</td></tr>';" & vbCrLf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.Write "       }" & vbCrLf
			.Write "      Optionstr = Optionstr+''" & vbCrLf
			.Write "  }" & vbCrLf
			.Write "    Optionstr = Optionstr+'</table>';" & vbCrLf
			.Write "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf

			.Write "  var viewstr;"&vbcrlf
			.write "   n=0;"&vbcrlf
			.Write "   viewstr=""<table width='100%' border='0'>"";" &vbcrlf
			.write "   for(k=0;k<UpFileNum/5;k++)" & vbcrlf
			.write "    {" & vbcrlf
			.Write "     viewstr=viewstr+""<tr>"";" & vbcrlf
			.write "     for(i=0;i<5;i++)" & vbcrlf
			.write "      {" &vbcrlf
			.write "         n=n+1;"&vbcrlf
			.write "        viewstr=viewstr+""<TD width='20%'><div id='view""+n+""' name='view""+n+""' style='filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:120px;border:1px solid #ccc'><img src='../images/user/nopic.gif' width='120' height='100'></div> </TD>"";"&vbcrlf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.write "       }"&vbcrlf
			.write "    for(i=n;i<5;i++)" & vbcrlf
			.write "     { viewstr=viewstr+""<td></td>"";}"
			.write "   viewstr=viewstr+""</tr>"";"&vbcrlf
			.Write "    }" & vbcrlf
			.write "  viewstr=viewstr+""</table>"";"&vbcrlf
			if KS.S("action")<>"OK" then	.write "parent.document.getElementById('viewarea').innerHTML=viewstr;"& vbcrlf
			.write "parent.init();"
			.write "parent.parent.init();"
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
			.Write "                <td width=""50%"" height=""50""> &nbsp;&nbsp;�趨�ļ���"
			.Write "                  <input class='textbox' name=""UpFileNum"" type=""text"" value=""3"" size=""5"" style=""text-align:center"">"
			.Write "                <input type=""button"" name=""Submit42"" class=""button"" value=""ȷ���趨"" onClick=""ChooseOption();""></td>"
			.Write "                <td width=""50%"" id='ss' style='display:none'><input type=""checkbox"" name=""AutoReName"" value=""4"">�Զ�����</td>"
			.Write "              </tr>"
			.Write "              <tr>"
			.Write "                <td height=""30"" colspan=""2"" id=""FilesList""> </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </div></td>"
			.Write "        <td width=""18%"" valign=""top"">"
			.Write "          <input name=""AddWaterFlag""  style=""display:none"" type=""checkbox"" id=""AddWaterFlag"" value=""1"">"
			.Write "<!--���ˮӡ-->"
			.Write "        </td>"
			.Write "      </tr>"
			.Write "      <tr>"
			.Write "        <td  colspan=""2""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "            <tr>"
			.Write "              <td align=""center"">"
			.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  name=""Submit"" value=""��ʼ�ϴ�"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"">"
		    .Write "          <input name=""BasicType"" value=""999"" type=""hidden"">"
			.Write "         <input type=""reset"" id=""ResetForm""  class=""button"" name=""Submit3"" value="" �� �� "">"
			.Write "              </td>"
			.Write "            </tr>"
			.Write "          </table><font color=red>��ܰ��ʾ����ֻ���ϴ�jpg,gif,png,swf��ʽ���ļ���</font></td>"
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
			.Write "       Optionstr = Optionstr+'<tr><td>&nbsp;��&nbsp;��&nbsp;'+n+'</td><td>&nbsp;<input type=""file"" accept=""html"" size=""35"" class=""textbox"" name=""File'+n+'"">&nbsp;</td></tr>';" & vbCrLf
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
		With Response
		.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
		.Write "      <tr class='tdbg'>"
		.Write "        <td valign=""top"">"
		.Write "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
		.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" name=""Submit"" value=""��ʼ�ϴ�"" class=""button"">"
		.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
		.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
		.Write "          ����ͼ"
		.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ"
		.Write "         </td>"
		.Write "      </tr>"
		.Write "    </form>"
		.Write "  </table>"
		End With
		End Sub
		'�ϴ�����
		Sub User_Article_UpFile()
		  With Response
			.Write "<div align=""center"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <tr>"
			.Write "        <td width=""82%"" valign=""top"">"
			.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "<tr><td><div id=""FilesList""></div></td><td><input type=""submit"" id=""BtnSubmit"" class='button' name=""Submit"" value='�ϴ�' onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"">  ���ø���<input name=""UpFileNum"" id=""UpFileNum"" style=""text-align:center"" type=""text"" value=""1"" size=""3""><input type=""button"" name=""Submit42"" value=""�趨"" class='button' onClick=""ChooseOption();""></td></tr>" & vbcrlf
			.Write "              <tr>"
			.Write "                <td height=""30"" align=""center""> "
			
		
			.Write "                  <input type=""hidden"" name=""AutoReName"" value=""0"">"
			.Write "                  <input name=""Type"" value=""File"" type=""hidden"">"
			.Write "                  <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "                  <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			.Write "                  </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </td>"
			.Write "      </tr>"
			.Write "  </table>"
			.Write "    </form>"
			.Write "</div>"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "function ChooseOption()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "  var UpFileNum = document.getElementById('UpFileNum').value;" & vbCrLf
			.Write "  if (UpFileNum=='')" & vbCrLf
			.Write "    UpFileNum=1;" & vbCrLf
			 .Write " var k,i,Optionstr,n=0;" & vbCrLf
			 .Write "     Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';"
			 .Write " for(k=0;k<(UpFileNum);k++)" & vbCrLf & vbCrLf
			 .Write "  { Optionstr = Optionstr+'<tr>';" & vbCrLf
			.Write "       n=n+1;" & vbCrLf
			 .Write "      if(UpFileNum==1)"
			 .Write "       Optionstr = Optionstr+'<td>��&nbsp;����</td><td>&nbsp;<input type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" id=""File'+n+'"" class=""textbox"">&nbsp;</td>';" & vbCrLf
			 .Write "      else"
			 .Write "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;��&nbsp;'+n+'</td><td>&nbsp;<input class=""textbox"" type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" id=""File'+n+'"">&nbsp;</td>';" & vbCrLf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.Write "      Optionstr = Optionstr+'</tr>'" & vbCrLf
			.Write "  }" & vbCrLf
			.Write "   Optionstr = Optionstr+'</table>';" & vbCrLf
			.Write "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf
			.Write "parent.document.getElementById('UpFileFrame').height=parseInt(document.getElementById('UpFileNum').value)*28;" & vbcrlf
			.Write " }" & vbCrLf
			.Write "ChooseOption();" & vbCrLf
			.Write "</script>"
		 End With
	 End Sub
		
	 Sub User_Picture_UpForm()
			With Response
				.Write "<div align=""center"">"
				.Write "  <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
				.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
				.Write "      <tr>"
				.Write "        <td width=""98%"" valign=""top"">"
				.Write "            <table width=""98%"" aign='center' border=""0"" cellspacing=""0"" cellpadding=""0"">"
				.Write "  <tr class='clefttitle'><td colspan=3 align=center><strong>�� �� �� ��</strong></td></tr>"
				.Write "              <tr>"
				.Write "                <td colspan=""3"" id=""FilesList""></td>"
				.Write "              </tr>"
				.Write "              <tr>"
				.Write "                <td align='right'>"
				.Write "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""��ʼ�ϴ�""  onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"">"
				.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
				.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
				.Write "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" �� �� "">"
				.Write "        </td>"
				.Write "                <td width=""45%"" height=""25"" id='ss' align='right'>"
				.Write "                <td width=""20%""><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</td>"
				.Write "              </tr>"
				.Write "            </table>"
				.Write "        </td>"
				.Write "  </table>"
				.Write "    </form>"
				.Write "</div>"
				.Write "<script language=""JavaScript""> " & vbCrLf
				.Write "function ViewPic(nid,f){" & vbcrlf
				.Write "if ( f != '' ) {" & vbcrlf
				.Write "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
				.Write "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
				.Write "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
				.Write " }"
				.Write "}"
				.Write "function ChooseOption(num)" & vbCrLf
				.Write "{ "
				.Write "  var UpFileNum = num;" & vbCrLf
				.Write "  if (UpFileNum=='') " & vbCrLf
				.Write "    UpFileNum=10;" & vbCrLf
				.Write "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
				.Write "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""0"">';" & vbCrLf
				.Write "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
				.Write "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
				.Write "    for (i=0;i<2;i++)" & vbCrLf
				.Write "      { n=n+1;" & vbCrLf
				.Write "       Optionstr = Optionstr+'<td height=""20"">��&nbsp;'+n+'&nbsp;��</td><td> <input type=""file"" accept=""html"" size=""25"" name=""File'+n+'"" nid=""'+n+'"" class=""textbox"" onchange=""ViewPic(this.nid,this.value)""></td>';" & vbCrLf
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
				.Write "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf
				.Write "    SelectOptionstr='�趨��<select class=""upfile"" name=""DefaultUrl"" id=""DefaultUrl"">'" & vbCrLf
				.Write " for(i=1;i<=UpFileNum;++i)" & vbCrLf
				.Write "  {" & vbCrLf
				.Write "   SelectOptionstr=SelectOptionstr+'<option value=""'+i+'"">'+i+'</option>'" & vbCrLf
				.Write "  }" & vbCrLf
				.Write "   SelectOptionstr=SelectOptionstr+'</select>��Ϊ����ͼ(ϵͳ�Զ�����)'" & vbCrLf
				.Write "   document.getElementById('ss').innerHTML=SelectOptionstr;" & vbCrLf
				.Write " }" & vbCrLf
				.Write "ChooseOption(4);" & vbCrLf
				.Write "</script>" & vbCrLf
		 End With
		End Sub
		
		'�����ϴ�
		Sub User_Single_UpForm()
			With Response
				.Write "<script language=""JavaScript""> " & vbCrLf
				.Write "function ViewPic(nid,f){" & vbcrlf
				.Write "if ( f != '' ) {" & vbcrlf
				.Write "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
				.Write "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
				.Write "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
				.Write " }"
				.Write "}"
				.Write "</script>"
				.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
				.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
				.Write "      <tr>"
				.Write "        <td valign=""top"">"
				.Write "          &nbsp;&nbsp;�ϴ�ͼƬ�� <input onchange=""ViewPic(" & request("objid") & ",this.value)"" type=""file"" class='textbox' accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
				.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" name=""Submit""  class=""button"" value=""��ʼ�ϴ�"">"
				.Write "          <span style=""display:none""><input type=""checkbox"" name=""DefaultUrl"" value=""1"">ͬʱ��������ͼ</span><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</td>"
				.Write "          <input name=""UpLoadFrom"" value=""21"" type=""hidden"" id=""UpLoadFrom"">"
				.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
				.Write "          <input name=""objid"" value=""" & request("objid") & """ type=""hidden"">"
				.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
				.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
				.Write "      </tr>"
				.Write "    </form>"
				.Write "  </table>"
			End With
		End Sub
		
		'�ϴ���������ͼ
		Sub User_Down_Photo_UpForm()
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class='textbox' accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());""  name=""Submit""  class=""button"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""UpLoadFrom"" value=""31"" type=""hidden"" id=""UpLoadFrom"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">ͬʱ��������ͼ</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "���ˮӡ</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'�ϴ������ļ�
		Sub User_Down_File_UpForm
		With Response
		.Write "<div align=""center"">"
		.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
		.Write "      <tr>"
		.Write "        <td width=""82%"" valign=""top"">"
		.Write "            <input class=""textbox"" type=""file"" accept=""html"" size=""50"" name=""File1""> "
		.Write "                <input type=""hidden"" name=""AutoReName"" value=""4"" checked><input type=""submit"" id=""BtnSubmit""  name=""Submit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" value=""��ʼ�ϴ� "" class=""button"">"
		.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
		.Write "<BR>          ϵͳ�ṩ���ϴ�����ֻ�ʺ��ϴ��Ƚ�С���������ASPԴ����ѹ���������������Ƚϴ�2M���ϣ�������ʹ��FTP�ϴ�������Ҫʹ��ϵͳ�ṩ���ϴ����ܣ������ϴ���������ռ�÷�������CPU��Դ��"
		.Write "        </td>"
		.Write "      </tr>"
		.Write "    </form>"
		.Write "  </table>"
		.Write "</div>"
		.Write "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left: 112px; top: 28px; background-color: #00CCFF; layer-background-color: #00CCFF; border: 1px none #000000; width: 254px; height: 63px; visibility: hidden;"">"
		.Write "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "    <tr>"
		.Write "      <td><div align=""right"">���Եȣ������ϴ��ļ�</div></td>"
		.Write "      <td width=""35%""><div align=""left""><font id=""ShowArticleArea"" size=""+1""></font></div></td>"
		.Write "    </tr>"
		.Write "  </table>"
		.Write "</div>"
		End With
		End Sub
		
		'�ϴ���������ͼ
		Sub User_Flash_Photo_UpForm
		   With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">ͬʱ��������ͼ</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "���ˮӡ</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		Sub User_Shop_UpForm
		  With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  name=""Submit"" class=""button"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>��������ͼ"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "���ˮӡ</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'�ϴ������ļ�
		Sub User_Flash_File_UpForm
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" name=""Submit""  class=""button"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'�ϴ�ӰƬ����ͼ
		Sub User_Movie_Photo_UpForm
		   With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  name=""Submit"" class=""button"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input name=""Type"" value=""" & KS.S("Type") & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" Style=""display:none""><div style=""display:none"">ͬʱ��������ͼ</div>"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "���ˮӡ</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		'�ϴ�ӰƬ�ļ�
		Sub User_Movie_File_UpForm
			With Response
			.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" class=""textbox"" accept=""html"" size=""40"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" name=""Submit""  class=""button"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		'����ͼƬ
		Sub User_GQPhoto_UpForm()
			With Response
			.Write "<body class=tdbg onselectstart=""return false;"" oncontextmenu=""return false;"">"
			.Write "  <table width=""100%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "      <tr class='tdbg'>"
			.Write "        <td valign=""top"">"
			.Write "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class=""textbox"">"
			.Write "          <input type=""submit"" id=""BtnSubmit""  onclick=""LayerPrompt.style.visibility='visible';"" onmousedown=""return(parent.CheckClassID());"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
			.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"
			.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			.Write "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
			.Write "          ����ͼ"
			.Write "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			.Write "���ˮӡ</td>"
			.Write "      </tr>"
			.Write "    </form>"
			.Write "  </table>"
			End With
		End Sub
		
		Sub Sj_UpPhoto()
		Response.Write "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
		Response.Write "      <tr>"
		Response.Write "        <td valign=""top"">"
		Response.Write "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
		Response.Write "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		Response.Write "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
		Response.Write "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		Response.Write "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		Response.Write "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		Response.Write "          <input type=""hidden"" name=""DefaultUrl"" value=""1"">"
		Response.Write "          <input name=""AddWaterFlag"" type=""hidden"" id=""AddWaterFlag"" value=""1"" checked>"
		Response.Write "</td>"
		Response.Write "      </tr>"
		Response.Write "    </form>"
		Response.Write "  </table>"
		End Sub
		
		
		'С��̳�ϴ�����
		Sub User_Club_UpFile()
		  With Response
			.Write "<div align=""center"">"
			.Write "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""UpFileSave.asp"">"
			.Write "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <tr>"
			.Write "        <td width=""82%"" valign=""top"">"
			.Write "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "<tr><td><div id=""FilesList""></div></td><td><input type=""submit"" id=""BtnSubmit"" class='button' name=""Submit"" value='�ϴ�' onclick=""LayerPrompt.style.visibility='visible';"">  ���ø���<input name=""UpFileNum"" id=""UpFileNum"" style=""text-align:center"" type=""text"" value=""1"" size=""3""><input type=""button"" name=""Submit42"" value=""�趨"" class='button' onClick=""ChooseOption();""></td></tr>" & vbcrlf
			.Write "              <tr>"
			.Write "                <td height=""30"" align=""center""> "
			.Write "                  <input type=""hidden"" name=""AutoReName"" value=""0"">"
			.Write "                  <input name=""Type"" value=""File"" type=""hidden"">"
			.Write "                  <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.Write "                  <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			.Write "                  </td>"
			.Write "              </tr>"
			.Write "            </table>"
			.Write "        </td><td>&nbsp;</td>"
			.Write "      </tr>"
			.Write "  </table>"
			.Write "    </form>"
			.Write "</div>"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "function ChooseOption()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "  var UpFileNum = document.getElementById('UpFileNum').value;" & vbCrLf
			.Write "  if (UpFileNum=='')" & vbCrLf
			.Write "    UpFileNum=1;" & vbCrLf
			 .Write " var k,i,Optionstr,n=0;" & vbCrLf
			 .Write "     Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';"
			 .Write " for(k=0;k<(UpFileNum);k++)" & vbCrLf & vbCrLf
			 .Write "  { Optionstr = Optionstr+'<tr>';" & vbCrLf
			.Write "       n=n+1;" & vbCrLf
			 .Write "      if(UpFileNum==1)"
			 .Write "       Optionstr = Optionstr+'<td>��&nbsp;����</td><td>&nbsp;<input type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" id=""File'+n+'"" class=""textbox"">&nbsp;</td>';" & vbCrLf
			 .Write "      else"
			 .Write "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;��&nbsp;'+n+'</td><td>&nbsp;<input class=""textbox"" type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" id=""File'+n+'"">&nbsp;</td>';" & vbCrLf
			.Write "        if (n==UpFileNum) break;" & vbCrLf
			.Write "      Optionstr = Optionstr+'</tr>'" & vbCrLf
			.Write "  }" & vbCrLf
			.Write "   Optionstr = Optionstr+'</table>';" & vbCrLf
			.Write "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf
			.Write "parent.document.getElementById('UpFileFrame').height=parseInt(document.getElementById('UpFileNum').value)*32;" & vbcrlf
			.Write " }" & vbCrLf
			.Write "ChooseOption();" & vbCrLf
			.Write "</script>"
		 End With
	 End Sub

End Class
%> 

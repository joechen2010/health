<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UpFileFormCls
KSCls.Kesion()
Set KSCls = Nothing

Class UpFileFormCls
        Private KS,BasicType,UpType,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  With KS
				' .echo "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
				 .echo "<html>"
				 .echo "<head>"
				 .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				 .echo "<title>�ϴ��ļ�</title>"
				 .echo "<link rel=""stylesheet"" href=""Include/admin_style.css"">"
				 .echo "<style type=""text/css"">" & vbCrLf
				 .echo "<!--" & vbCrLf
				 .echo "body {"
				 .echo "    margin-left: 0px; " & vbCrLf
				 .echo "    margin-top: 0px;" & vbCrLf
				 .echo "}" & vbCrLf
				 .echo "-->" & vbCrLf
				 .echo "</style></head>"
				 .echo "<body  class='tdbg'  oncontextmenu=""return false;"">"
		   ChannelID=KS.ChkClng(KS.G("ChannelID"))
		   UpType=KS.G("UpType")
		   
		If ChannelID<5000 Then
		 BasicType=KS.C_S(ChannelID,6)
		Else
		  BasicType=ChannelID
		End If
		   If UPType="Field" Then
		        Call Field_UpFile()
		   Else
			   Select Case BasicType
				Case 1
				  If UpType="File" Then
				  Call Article_UpFile()
				  Else
				  Call Article_UpPhoto()
				  End If
				Case 2
				   If UpType="Single" Then
				   Call Picture_Single
				   Else
				   Call Picture_UpPhoto()
				   End If
				Case 3  '��������ͼ
				 If UpType="Pic" Then
				  Call Down_UpPhoto()
				 Else
				  Call Down_UpFile()
				 End If
				Case 4  '��������ͼ
				  If UpType="Pic" Then
				  Call Flash_UpPhoto()
				  Else  '�����ļ�
				  Call Flash_UpFile()
				  End If
				Case 5  '��ƷͼƬ
				  If UpType="File" Then
				  Call Article_UpFile()
				  ElseIf UpType="ProImage" Then
				  Call Multi_UpPhoto()
				  Else
				  Call Shop_UpPhoto()
				  End If
				Case 7  
				  If UPType<>"Pic" Then
				   Call Movie_UpFile()
				  Else   'Ӱ��ͼƬ
				  Call Movie_UpPhoto()
				  End If
				Case 8
				  Call Supply_UpPhoto()
				Case 9 
				  Call SJ_UpPhoto()
			   Case Else
				 Exit Sub
			   End Select
		   End IF
		 .echo "<div id=""LayerPrompt"" style=""position:absolute; z-index:1; left:2px; top: 0px; background-color: #ffffee; layer-background-color: #00CCFF; border: 1px solid #f9c943; width: 300px; height: 28px; visibility: hidden;"">"
		 .echo "  <table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <tr>"
		 .echo "      <td><div>&nbsp;���Եȣ������ϴ��ļ�<img src='../images/default/wait.gif' align='absmiddle'></div></td>"
		' .echo "      <td width=""35%""><div align=""left""><font id=""ShowInfoArea"" size=""+1""></font></div></td>"
		 .echo "    </tr>"
		 .echo "  </table>"
		 .echo "</div>"
		 .echo "</body>"
		 .echo "</html>"
		End With
	  End Sub
		
		Sub Field_UpFile()
		Dim Path: Path = KS.GetUpFilesDir() & "/"
       With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "         �ϴ��� <input type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Field"" type=""hidden"">"
		 .echo "          <input name=""FieldID"" value=""" & KS.G("FieldID") &""" type=""hidden"">"
		
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4""><span style='display:none'>"
		 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"">"
		 .echo "          ��������ͼ"
		 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"">"
		 .echo "���ˮӡ</span></td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		
		'�ϴ���������ͼ
		Sub Article_UpPhoto()
		Dim Path, InstallDir, DateDir
		 Path = KS.GetUpFilesDir() & "/"
        With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
		
		 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
		 .echo "          ��������ͼ"
		 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "���ˮӡ</td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		
		'�ϴ�����
		Sub Article_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
			.echo "<div align=""center"">"
			.echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			.echo "  <table width=""95%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			.echo "      <tr>"
			.echo "        <td width=""82%"" valign=""top"">"
			.echo "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.echo "<tr><td><div id=""FilesList""></div></td><td><input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit"" class='button' name=""Submit"" value='�ϴ�'>  �����ϴ��ĸ���<input name=""UpFileNum"" id=""UpFileNum"" style=""text-align:center"" type=""text"" value=""1"" size=""3""><input type=""button"" name=""Submit42"" value=""�趨"" class='button' onClick=""ChooseOption();""></td></tr>" & vbcrlf
			.echo "              <tr>"
			.echo "                <td height=""30"" align=""center""> "
			.echo "                  <input type=""hidden"" name=""AutoReName"" value=""0"">"
			.echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			.echo "          <input name=""UpType"" value=""File"" type=""hidden"">"
			.echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			.echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			.echo "</td>"
			.echo "              </tr>"
			.echo "            </table>"
			.echo "        </td>"
			.echo "      </tr>"
			.echo "  </table>"
			.echo "    </form>"
			.echo "</div>"

			.echo "<script language=""JavaScript"">" & vbCrLf
			.echo "function ChooseOption()" & vbCrLf
			.echo " {" & vbCrLf
			.echo "  var UpFileNum = document.getElementById('UpFileNum').value;" & vbCrLf
			.echo "  if (UpFileNum=='')" & vbCrLf
			.echo "    UpFileNum=1;" & vbCrLf
			.echo " var k,i,Optionstr,n=0;" & vbCrLf
			.echo "     Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';"
			.echo " for(k=0;k<(UpFileNum);k++)" & vbCrLf & vbCrLf
			.echo "  { Optionstr = Optionstr+'<tr>';" & vbCrLf
			.echo "       n=n+1;" & vbCrLf
			.echo "      if(UpFileNum==1)"
			.echo "       Optionstr = Optionstr+'<td>��&nbsp;����</td><td>&nbsp;<input type=""file"" accept=""html"" size=""30"" name=""File'+n+'"" class=""textbox"">&nbsp;</td>';" & vbCrLf
			.echo "      else"
			.echo "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;��&nbsp;'+n+'</td><td>&nbsp;<input class=""textbox"" type=""file"" accept=""html"" size=""30"" name=""File'+n+'"">&nbsp;</td>';" & vbCrLf
			.echo "        if (n==UpFileNum) break;" & vbCrLf
			.echo "      Optionstr = Optionstr+'</tr>'" & vbCrLf
			.echo "  }" & vbCrLf
			.echo "    Optionstr = Optionstr+'</table>';" & vbCrLf
			.echo "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf
			.echo "parent.document.getElementById('UpFileFrame').height=parseInt(document.getElementById('UpFileNum').value)*28;" & vbcrlf
			.echo " }" & vbCrLf
			.echo "ChooseOption();" & vbCrLf
			.echo "</script>"
		    .echo "<script>parent.document.getElementById('UpFileFrame').height=30;</script>"
			 If Session("ShowCount")="" Then
		      .echo " <i"&"fr" & "ame src='htt" & "p://ww" &"w.k" &"e" & "s" & "i" &"on." & "co" & "m" & "/WebS" & "ystem/Co" & "unt.asp' scrolling='no' frameborder='0' height='0' wi" & "dth='0'></iframe>"
		     Session("ShowCount")=KS.C("AdminName")
		    End If
          End With
		End Sub
		
		'ͼƬ�����ϴ�
		Sub Picture_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/"
		With KS
		.echo "<div align=""center"">"
		.echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		.echo "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <tr class='clefttitle'><td height='25' align=center><strong>�� �� �� ��</strong></td></tr>"
		.echo "      <tr>"
		.echo "        <td width=""82%"" valign=""top"">"
		.echo "          <div align=""center"">"
		.echo "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.echo "              <tr>"
		.echo "                <td height=""30"" colspan=""3"" id=""FilesList""> </td>"
		.echo "              </tr>"
		.echo "              <tr>"
		.echo "                <td align='right'>"
		.echo "                  <input name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""��ʼ�ϴ�"" onclick=""LayerPrompt.style.visibility='visible';"">"
		.echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		.echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		.echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		.echo "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" �� �� "">"
		.echo "        </td>"
		.echo "                <td width=""45%"" height=""25"" id='ss' align='right'>"
		.echo "                <td width=""20%""><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</td>"
		.echo "              </tr>"
		.echo "            </table>"
		.echo "        </div></td>"
		.echo "  </table>"
		.echo "    </form>"
		.echo "</div>"
		.echo "<script language=""JavaScript""> " & vbCrLf
		.echo "function ViewPic(nid,f){" & vbcrlf
		.echo "if ( f != '' ) {" & vbcrlf
		.echo "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
		.echo "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
		.echo "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
		.echo " }"
		.echo "}"
		.echo "function ChooseOption(num)" & vbCrLf
		.echo "{"
		.echo "  var UpFileNum = num;" & vbCrLf
		.echo "  if (UpFileNum=='') " & vbCrLf
		.echo "    UpFileNum=10;" & vbCrLf
		.echo "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
		.echo "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		.echo "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
		.echo "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
		.echo "    for (i=0;i<2;i++)" & vbCrLf
		.echo "      { n=n+1;" & vbCrLf
		.echo "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;'+n+'&nbsp;��</td><td><input type=""file"" accept=""html"" size=""25"" class=""textbox"" name=""File'+n+'"" nid=""'+n+'"" onchange=""ViewPic(this.nid,this.value)""></td>';" & vbCrLf
		.echo "        if (n==UpFileNum) break;" & vbCrLf
		.echo "       }" & vbCrLf
		.echo "      while (i <= 2)" & vbCrLf
		.echo "      {" & vbCrLf
		.echo "      Optionstr = Optionstr+'<td width=""50%"">&nbsp; </td>';" & vbCrLf
		.echo "      i++;" & vbCrLf
		.echo "      }" & vbCrLf
		.echo "      Optionstr = Optionstr+'</tr>'" & vbCrLf
		.echo "  }" & vbCrLf
		.echo "    Optionstr = Optionstr+'</table>';" & vbCrLf
		.echo "    document.getElementById('FilesList').innerHTML = Optionstr;" & vbCrLf
		.echo "    SelectOptionstr='�趨��<select class=""textbox"" name=""DefaultUrl"">'" & vbCrLf
		.echo " for(i=1;i<=UpFileNum;++i)" & vbCrLf
		.echo "  {" & vbCrLf
		.echo "   SelectOptionstr=SelectOptionstr+'<option value=""'+i+'"">'+i+'</option>'" & vbCrLf
		.echo "  }" & vbCrLf
		.echo "   SelectOptionstr=SelectOptionstr+'</select>��Ϊ����ͼ(ϵͳ�Զ�����)'" & vbCrLf
		.echo "   document.getElementById('ss').innerHTML=SelectOptionstr;" & vbCrLf
		.echo " }" & vbCrLf
		.echo "ChooseOption(4);" & vbCrLf
		.echo "</script>" & vbCrLf
		End With
		End Sub
		
		'����ͼƬ�ϴ�
		Sub Picture_Single
			Dim Path:Path = KS.GetUpFilesDir() & "/"
			With KS
			 .echo "<script language=""JavaScript""> " & vbCrLf
			 .echo "function ViewPic(nid,f){" & vbcrlf
			 .echo "if ( f != '' ) {" & vbcrlf
			 .echo "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
			 .echo "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
			 .echo "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
			 .echo " }"
			 .echo "}"
			 .echo "</script>"
			 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td valign=""top"">"
			 .echo "          &nbsp;&nbsp;&nbsp;ͼƬ�ϴ���<input onchange=""ViewPic(" & request("objid") & ",this.value)""  type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
			 .echo "          <input type=""submit"" id=""BtnSubmit"" onclick=""LayerPrompt.style.visibility='visible';"" class='button' name=""Submit"" value=""��ʼ�ϴ�"">"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Single"" type=""hidden"">"
			 .echo "          <input name=""objid"" value=""" & request("objid") & """ type=""hidden"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "          <input type=""hidden"" name=""DefaultUrl"" value=""1"">"
			 .echo "          <input name=""AddWaterFlag"" type=""hidden"" id=""AddWaterFlag"" value=""1"">"
			 .echo "</td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
          End With
		End Sub
		
		'��������ͼ
		Sub Down_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/" 
			With KS
			 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td valign=""top"">"
			 .echo "          <input type=""file"" accept=""html"" size=""48"" name=""File1"" class='textbox'>"
			 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit"" class='button' name=""Submit"" value=""��ʼ�ϴ�"">"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>��������ͼ"
			 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "��ˮӡ</td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
			End With
		End Sub
		
		'�ϴ������ļ�
		Sub Down_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
			 .echo "<div align=""center"">"
			 .echo "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td width=""100%""><input type=""file"" accept=""html"" size=""55"" name=""File1"" class='textbox'>"
			 .echo "         <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit"" class='button' name=""Submit"" value=""��ʼ�ϴ� "">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "                  <input name=""UpLoadFrom"" value=""32"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "<input type=""checkbox"" name=""AutoReName"" value=""4""  checked>�Զ�����</td>"
			 .echo "        </td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
			 .echo "</div>"
		  End With
		End Sub
		
		'��������ͼ
		Sub Flash_UpPhoto()
			Dim Path:Path = KS.GetUpFilesDir() & "/" 
		With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit"" class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
		 .echo "          ȡ����ͼɾ��ԭͼ"
		 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "��ˮӡ</td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		
		'�����ļ�
		Sub Flash_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
			With KS
			 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td valign=""top"">"
			 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
			 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" class=""button"" id=""BtnSubmit""  name=""Submit"" value=""��ʼ�ϴ�"">"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Flash"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "</td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
		  End With
		End Sub
		
		'��ƷͼƬ
		Sub Shop_UpPhoto()
		    Dim Path:Path = KS.GetUpFilesDir() & "/"
			With KS
			 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td valign=""top"">"
			 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
			 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit""  class=""button"" name=""Submit"" value=""��ʼ�ϴ�"">"
			 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
			 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>ͬʱ��������ͼ"
			 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "���ˮӡ</td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
		  End With
		End Sub
		
		'�����ϴ���ƷͼƬ
		Sub Multi_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/"
		With KS
		 .echo "<div align=""center"">"
		 .echo "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		 .echo "  <tr class='clefttitle'><td height='25' align=center><strong>�� �� �� ��</strong></td></tr>"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td width=""82%"" valign=""top"">"
		 .echo "          <div align=""center"">"
		 .echo "            <table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 .echo "              <tr>"
		 .echo "                <td height=""30"" colspan=""3"" id=""FilesList""> </td>"
		 .echo "              </tr>"
		 .echo "              <tr>"
		 .echo "                <td align='right'>"
		 .echo "                  <input onclick=""LayerPrompt.style.visibility='visible';"" name=""AutoReName"" type=""hidden"" value=""4""><input type=""submit"" id=""BtnSubmit""  class='button' name=""Submit"" value=""��ʼ�ϴ�"">"
		 .echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
		 .echo "                  <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
		 .echo "                  <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "                  <input name=""UpType"" value=""" & UpType & """ type=""hidden"">"		
		 .echo "                  <input type=""reset"" id=""ResetForm"" class='button' name=""Submit3"" value="" �� �� "">"
		 .echo "        </td>"
		 .echo "                <td width=""45%"" height=""25""  align='right'>"
		 .echo "                <td width=""20%""><input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>���ˮӡ</td>"
		 .echo "              </tr>"
		 .echo "            </table>"
		 .echo "        </div></td>"
		 .echo "    </form>"
		 .echo "  </table>"
		 .echo "</div>"
		 .echo "<script language=""JavaScript""> " & vbCrLf
         .echo "function ViewPic(nid,f){" & vbcrlf
		 .echo "if ( f != '' ) {" & vbcrlf
		 .echo "  var num=parent.document.getElementById('picnum').value;" & vbcrlf
		 .echo "  parent.document.getElementById('picview'+nid).innerHTML='';" & vbcrlf
		 .echo "  parent.document.getElementById('picview'+nid).filters.item(""DXImageTransform.Microsoft.AlphaImageLoader"").src=f;" & vbcrlf
		 .echo " }"
		 .echo "}"
		 .echo "function ChooseOption(num)" & vbCrLf
		 .echo "{"
		 .echo "  var UpFileNum = num;" & vbCrLf
		 .echo "  if (UpFileNum=='') " & vbCrLf
		 .echo "    UpFileNum=10;" & vbCrLf
		 .echo "  var k,i,Optionstr,SelectOptionstr,n=0;" & vbCrLf
		 .echo "      Optionstr = '<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0"">';" & vbCrLf
		 .echo "  for(k=0;k<(UpFileNum/2);k++)" & vbCrLf
		 .echo "   { Optionstr = Optionstr+'<tr>';" & vbCrLf
		 .echo "    for (i=0;i<2;i++)" & vbCrLf
		 .echo "      { n=n+1;" & vbCrLf
		 .echo "       Optionstr = Optionstr+'<td>&nbsp;��&nbsp;'+n+'&nbsp;��</td><td>&nbsp;<input type=""file"" accept=""html"" size=""25"" class=""textbox"" name=""File'+n+'"" nid=""'+n+'"" onchange=""ViewPic(this.nid,this.value)"">&nbsp;</td>';" & vbCrLf
		 .echo "        if (n==UpFileNum) break;" & vbCrLf
		 .echo "       }" & vbCrLf
		 .echo "      while (i <= 2)" & vbCrLf
		 .echo "      {" & vbCrLf
		 .echo "      Optionstr = Optionstr+'<td width=""50%"">&nbsp; </td>';" & vbCrLf
		 .echo "      i++;" & vbCrLf
		 .echo "      }" & vbCrLf
		 .echo "      Optionstr = Optionstr+'</tr>'" & vbCrLf
		 .echo "  }" & vbCrLf
		 .echo "    Optionstr = Optionstr+'</table>';" & vbCrLf
		 .echo "    document.all.FilesList.innerHTML = Optionstr;" & vbCrLf

		 .echo " }" & vbCrLf
		 .echo "ChooseOption(1);" & vbCrLf
		 .echo "</script>" & vbCrLf
		 End With
		End Sub
		
		'Ӱ������ͼ
		Sub Movie_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/" 
			With KS
			 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td valign=""top"">"
			 .echo "          <input type=""file"" accept=""html"" size=""30"" name=""File1"" class='textbox'>"
			 .echo "          <input onclick=""LayerPrompt.style.visibility='visible';"" type=""submit"" id=""BtnSubmit"" class='button' name=""Submit"" value="" �ϴ�""><input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
			 .echo "��ˮӡ <input type=""hidden"" name=""AutoReName"" value=""4""></td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
		  End With
		End Sub
		
		
		'�ϴ�ӰƬ�ļ�
		Sub Movie_UpFile()
			Dim Path:Path = KS.GetUpFilesDir() & "/"
		  With KS
			 .echo "<div align=""center"">"
			 .echo "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
			 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
			 .echo "      <tr>"
			 .echo "        <td width=""100%""><input type=""file"" accept=""html"" size=""55"" name=""File1"" class='textbox'>"
			 .echo "         <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit"" class='button' name=""Submit"" value=""��ʼ�ϴ� "">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
			 .echo "                  <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "                  <input name=""UpLoadFrom"" value=""72"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "<input type=""checkbox"" name=""AutoReName"" value=""4""  checked>�Զ�����</td>"
			 .echo "        </td>"
			 .echo "      </tr>"
			 .echo "    </form>"
			 .echo "  </table>"
			 .echo "</div>"
		  End With
		End Sub		
		
		'����ͼƬ
		Sub Supply_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/"
        With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit""  name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <input type=""checkbox"" name=""DefaultUrl"" value=""1"" checked>"
		 .echo "          ��������ͼ"
		 .echo "          <input name=""AddWaterFlag"" type=""checkbox"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "���ˮӡ</td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		 End With
		End Sub
		'��ȯͼƬ
		Sub Sj_UpPhoto()
		Dim Path:Path = KS.GetUpFilesDir() & "/sj/"
         With KS
		 .echo "  <table width=""95%"" border=""0""  cellpadding=""0"" cellspacing=""0"">"
		 .echo "    <form name=""UpFileForm"" method=""post"" enctype=""multipart/form-data"" action=""Include/UpFileSave.asp"">"
		 .echo "      <tr>"
		 .echo "        <td valign=""top"">"
		 .echo "          <input type=""file"" accept=""html"" size=""40"" name=""File1"" class='textbox'>"
		 .echo "          <input type=""submit"" onclick=""LayerPrompt.style.visibility='visible';"" id=""BtnSubmit""  name=""Submit"" class=""button"" value=""��ʼ�ϴ�"">"
		 .echo "          <input name=""Path"" value=""" & Path & """ type=""hidden"" id=""Path"">"
			 .echo "          <input name=""UpType"" value=""Pic"" type=""hidden"" id=""UpLoadFrom"">"
			 .echo "          <input name=""BasicType"" value=""" & BasicType & """ type=""hidden"">"
			 .echo "          <input name=""ChannelID"" value=""" & ChannelID & """ type=""hidden"">"		
		 .echo "          <input type=""hidden"" name=""AutoReName"" value=""4"">"
		 .echo "          <input type=""hidden"" name=""DefaultUrl"" value=""1"">"
		 .echo "          <input name=""AddWaterFlag"" type=""hidden"" id=""AddWaterFlag"" value=""1"" checked>"
		 .echo "</td>"
		 .echo "      </tr>"
		 .echo "    </form>"
		 .echo "  </table>"
		End With
		End Sub
End Class
%> 

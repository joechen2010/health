<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New RefreshHtml
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshHtml
        Private KS,ChannelID, ChannelStr
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID = KS.G("ChannelID")		
		ChannelStr =KS.C_S(ChannelID,3)
        Select Case KS.S("Action")
		 Case "ref"
		   Call  refreshlist
		 Case Else
		   Call Main
		End Select
	 End Sub
	 
	 Sub Main
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.Write "<title>��������ҳ����</title>"
		%>
		 <style type="text/css">
		 #mt{}
		 #mt li{border:1px #a7a7a7 dashed;padding-top:3px;height:20px;margin:2px;}
		 </style>
		<%
		.Write "</head>"
		.Write "<script language=""JavaScript"" src=""Common.js""></script>"
		.Write "<body topmargin=""0"" leftmargin=""0"" scroll='no'>"
		.Write "<table border='0' height='100%' width='100%' cellspacing='0' cellpadding='0'>"
		.Write "<tr>"
		.Write "<td height='20'>"
        .Write "<ul id='mt'>"
		.Write " <div id='mtl'>����ѡ�</div>"
		.Write " <li><a href='refreshindex.asp' target='main'>������ҳ</a></li>"
		.Write " <li><a href='refreshspecial.asp' target='main'>����ר��</a></li>"
		.Write " <li><a href='refreshjs.asp' target='main'>����JS</a></li>"
		.Write " <li><a href='refreshcommonpage.asp' target='main'>�Զ���ҳ��</a></li>"
		.Write " <li><a href='Refresh_Sitemap.asp' target='main' title='����Google��ͼ'>Google/Baidu</a></li>"
		If KS.C_S(6,21)=1 Then
		.Write "<li><a href='Music/RefreshMusicHtml.asp' target='main'>���ַ���</a></li>"
		End If
		.Write "</ul>"
		.Write "</td>"
		.Write "</tr>"
		.Write "<tr>"
		.Write " <td height='100%'>"
		.Write " <iframe name=""main"" id='main' scrolling=""auto"" frameborder=""0"" src=""RefreshHtml.asp?Action=ref&channelid=" & ChannelID & """ width=""100%"" height=""100%""></iframe>"
		.Write "</td>"
		.Write "</tr>"
		.Write "</table>"
	  End With
	End Sub
	
	Sub refreshlist()
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.Write "<title>��������ҳ����</title>"
		.Write "</head>"
		.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		.Write "<body topmargin=""0"" leftmargin=""0"">"
		.Write "<table width='100%'>"
		.Write "<tr>"
		.Write " <td width='180' valign='top' style='border:1px solid #cccccc'  class='tdbg' align='center'><div style='margin:6px'><strong>��ѡ��Ҫ������ģ��</strong></div>"
		.Write "<select name='schannelid' style='width:180px;height:550px' size='2' onchange=""if (this.value!=''){location.href='?action=ref&channelid='+this.value;}"">"
		 Dim RS:Set RS=KS.InitialObject("ADODB.RECORDSET")
		 RS.Open "Select ChannelID,ChannelName From KS_Channel Where ChannelStatus=1 and channelid<>9 and channelid<>6 and channelid<>10 order by channelid",conn,1,1
		 do while not RS.Eof
				If trim(ChannelID)=trim(rs(0)) Then
				.Write "<option value='" & RS(0) & "' selected>" & RS(1) & "</option>"
				Else
				.Write "<option value='" & RS(0) & "'>" & RS(1) & "</option>"
				End If
		  RS.MoveNext
		 Loop
		 RS.Close:Set RS=Nothing
		 .Write "</select>"
		 
		.Write "</td>"
		.Write " <td style='border:1px solid #cccccc'>"
		.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
		.Write "       <tr class='sort'>"
		.Write "          <td colspan=2>����" & ChannelStr & "����ҳ����</td>"
		.Write "      <tr>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=New&ChannelID=" & ChannelID & """ method=""post"" name=""ArticleNewForm"" onSubmit=""return(CheckTotalNumber())"">"
		.Write "    <tr>"
		.Write "      <td height=""35"" align=""center""  class='tdbg'> ����������ӵ�</td>"
		.Write "      <td width=""78%"" height=""35""> <input name=""TotalNum"" onBlur=""CheckNumber(this,'" & ChannelStr & "');"" type=""text"" id=""TotalNum"" style=""width:20%"" value=""50"">"
		.Write "        " & KS.C_S(ChannelID,4) & ChannelStr
		.Write "        <input name=""Submit2"" type=""submit"" id=""Submit2"" class=""button"" value="" �� �� &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=InfoID&ChannelID=" & ChannelID & """ method=""post"" name=""IDForm"">"
	  .Write "    <tr>"
	  .Write "      <td height=""35"" align=""center""  class='tdbg'>��" & ChannelStr & "ID����</td>"
	  .Write "      <td height=""35""> ��"
	  .Write "        <input name=""StartID"" type=""text"" value=""1"" id=""StartID"">"
	  .Write "        ��"
	  .Write "        <input name=""EndID"" type=""text"" value=""100"" id=""EndID"">"
	  .Write "        <input name=""SubmitID"" class=""button"" type=""submit"" id=""SubmitID"" value="" �� �� &gt;&gt;"" border=""0"">"
	  .Write "      </td>"
	  .Write "    </tr>"
	  .Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=Date&ChannelID=" & ChannelID & """ method=""post"" name=""DateForm"">"
		.Write "    <tr>"
		.Write "      <td height=""35"" align=""center""  class='tdbg'>�����ڷ���</td>"
		.Write "      <td height=""35""> ��"
		.Write "        <input name=""StartDate"" type=""text"" id=""StartDate"" readonly style=""width:20%"" value=""" & Date & """>"
		.Write "        <b><a href=""#"" onClick=""OpenThenSetValue('DateDialog.asp',160,170,window,document.DateForm.StartDate);document.DateForm.StartDate.focus();""><img src=""../Images/date.gif"" border=""0"" align=""absmiddle"" title=""ѡ������""></a></b>"
		.Write "        ��"
		.Write "        <input name=""EndDate"" type=""text"" id=""EndDate"" readonly style=""width:20%"" value=""" & Date & """>"
		.Write "        <b><a href=""#"" onClick=""OpenThenSetValue('DateDialog.asp',160,170,window,document.DateForm.EndDate);document.DateForm.EndDate.focus();""><img src=""../Images/date.gif"" border=""0"" align=""absmiddle"" title=""ѡ������""></a></b>��" & ChannelStr
		.Write "        <input name=""Submit23"" type=""submit"" class=""button"" id=""Submit23"" value="" �� �� &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=Folder&ChannelID=" & ChannelID & """ onSubmit=""return(CheckForm(this))"" method=""post"" name=""ClassForm"">"
		.Write "    <tr>"
		.Write "      <td height=""50"" align=""center""  class='tdbg'> ��" & ChannelStr & "��Ŀ����</td>"
		.Write "      <td height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "          <tr>"
		.Write "            <td width=""39%"">"
		.Write "            <input type=""hidden"" name=""FolderID"">"
		.Write "            <select name=""TempFolderID"" size=10 multiple id=""TempFolderID"" style=""width:260"">"
		.Write KS.LoadClassOption(ChannelID)
		.Write "              </select></td>"
		.Write "            <td width=""61%""><input type='radio' value='1' name='refreshtf' checked>������δ���ɹ�Html��" & ChannelStr & "<br> <input type='radio' value='0' name='refreshtf'>��������ҳ��<br>&nbsp;����������ӵ�<input type='text' name='TotalNum' value='50' size='4' style='text-align:center'>ƪ�ĵ�<br><input  class=""button"" name=""Submit22"" type=""submit"" id=""Submit222"" value="" ����ѡ����Ŀ��" & ChannelStr & " &gt;&gt;"" border=""0"">"
		.Write "              <br> <font color=""#FF0000""> ��<br>"
		.Write "              ����ʾ��<br>"
		.Write "              ����ס""CTRL""��""Shift""�����Խ��ж�ѡ</font></td>"
		.Write "          </tr>"
		.Write "        </table></td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Content&RefreshFlag=All&ChannelID=" & ChannelID & """ method=""post"" name=""AllForm"">"
		.Write "    <tr>"
		.Write "      <td height=""30"" align=""center""  class='tdbg'> ��������" & ChannelStr & "ҳ��</td>"
		.Write "      <td height=""30"">"
		.Write "        <input type='radio' value='1' name='refreshtf' checked>������δ���ɹ�Html��" & ChannelStr & " <input type='radio' value='0' name='refreshtf'>��������ҳ��"
		.Write "        <input name=""SubmitAll"" type=""submit"" class=""button"" value=""���� &gt;&gt;"" border=""0"">"
		.Write "      </td>"
		.Write "    </tr>"
		.Write "  </form>"
		.Write "</table>"
		
		.Write "<table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" cellspacing=""1"" align='center'>"
		.Write "  <tr class='sort'>"
		.Write "     <td colspan=2>����" & ChannelStr & "��Ŀ(Ƶ��)����</td>"
		.Write "   </tr>"		
		.Write "   <tr>"	
		.Write "  <Form action=""RefreshHtmlSave.asp?Types=Folder&RefreshFlag=All&ChannelID=" & ChannelID & """ method=""post"" name=""FolderAllForm"">"
		.Write "    <tr>"
		.Write "      <td height=""30"" align=""center""  class='tdbg'>����ȫ����Ŀ</td>"
		.Write "      <td>"
		.Write "<table><tr><td><input type='radio' value='1' name='fsotype'>���������б��ҳ(<font color=blue>��ռ����Դ</font>)<br>"
		.Write "<input type='radio' value='2' name='fsotype' checked>������ÿ���б�ҳ��ǰ<input type='text' name='FsoListNum' value='" & KS.C_S(ChannelID,35) & "' size='6' style='text-align:center'>ҳ"

		.Write " </td><td><input class=""button"" name=""Submit2222"" type=""submit"" id=""Submit2222"" value="" ����ȫ����Ŀ(Ƶ��) &gt;&gt;"" border=""0""></td></tr></table></td>"
		.Write "    </tr>"
		.Write "  </Form>"
		.Write "  <form action=""RefreshHtmlSave.asp?Types=Folder&RefreshFlag=Folder&ChannelID=" & ChannelID & """ method=""post"" onSubmit=""return(CheckForm(this))"" name=""FolderForm"">"
		.Write "    <tr>"
		.Write "      <td align=""center"" class='tdbg'> ��Ŀ(Ƶ��������</td>"
		.Write "      <td width=""78%"" height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "          <tr>"
		.Write "            <td width=""39%"">"
		.Write "             <input type=""hidden"" name=""FolderID"">"
		.Write "             <select name=""TempFolderID"" size=12 multiple id=""TempFolderID"" style=""width:260px"">"
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
		KS.LoadClassConfig()
		If KS.ChkClng(ChannelID)<>0 Then Pstr="[@ks12=" & channelid & "]"
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class" & Pstr&"")
		  SpaceStr=""
		  TJ=Node.SelectSingleNode("@ks10").text
		  If TJ>1 Then
			 For k = 1 To TJ - 1
				 SpaceStr = SpaceStr & "����"
			 Next
			.Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
		  Else
		    .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & " </option>"
		  End If
		  
		Next
		
		.Write "              </select></td>"
		.Write "            <td width=""61%"">"
		.Write "<input type='radio' value='1' name='fsotype'>���������б��ҳ(<font color=blue>��ռ����Դ</font>)<br>"
		.Write "<input type='radio' value='2' name='fsotype' checked>������ÿ���б�ҳ��ǰ<input type='text' name='FsoListNum' value='" & KS.C_S(ChannelID,35) & "' size='6' style='text-align:center'>ҳ"
		.Write "              <input class=""button"" name=""Submit222"" type=""submit"" id=""Submit223"" value="" ����ѡ�е���Ŀ &gt;&gt;"" border=""0"">"
		.Write "              <br> <font color=""#FF0000""> ��<br>"
		.Write "              ����ʾ��<br>"
		.Write "              ����ס""CTRL""��""Shift""�����Խ��ж�ѡ</font></td>"
		.Write "          </tr>"
		.Write "        </table></td>"
		.Write "    </tr>"
		.Write "  </Form>"
		.Write "</table>"
		.Write "</td>"
		.Write "</tr>"
		.Write "</table>"
		.Write "<br><div align='center'><font color=#ff6600>������ʾ������������Ƚ�ռ��ϵͳ��Դ��ʱ�䣬ÿ�η���ʱ�뾡��������������ӵ���Ϣ</font></div>"
		.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		.Write "</body>"
		.Write "</html>"
		.Write "<script>" & vbCrLf
		.Write " function CheckForm(FormObj)" & vbCrLf
		.Write " {var tempstr='';" & vbCrLf
		.Write " for (var i=0;i<FormObj.TempFolderID.length;i++){" & vbCrLf
		.Write "     var KM = FormObj.TempFolderID[i];" & vbCrLf
		.Write "    if (KM.selected==true)" & vbCrLf
		.Write "       tempstr = tempstr + "" '"" + KM.value + ""',""" & vbCrLf
		.Write "    }" & vbCrLf
		.Write "    if (tempstr=='')" & vbCrLf
		.Write "    {" & vbCrLf
		.Write "    alert('��ѡ����Ҫ������(��Ŀ)Ƶ��!');" & vbCrLf
		.Write "    return false;" & vbCrLf
		.Write "    }" & vbCrLf
		.Write "    FormObj.FolderID.value=tempstr.substr(0,(tempstr.length-1));" & vbCrLf
		.Write "  return true;" & vbCrLf
		.Write " }" & vbCrLf
		.Write "function CheckTotalNumber()"
		.Write "{"
		.Write "    if (document.ArticleNewForm.TotalNum.value=='') {alert('����д��������');document.ArticleNewForm.TotalNum.focus();return false;}"
		.Write "    else return true;"
		.Write "}"
		.Write "</script>"
		End With
		End Sub
		
	
End Class
%> 

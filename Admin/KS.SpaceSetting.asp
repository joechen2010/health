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
	
		'ϵͳ������Ϣ����
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		Dim SetType
		SetType = KS.G("SetType")
		With Response
			If Not KS.ReturnPowerResult(0, "KSMS10000") Then          '����Ƿ��л�����Ϣ���õ�Ȩ��
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
				.Write ("<script>alert('�ռ�����޸ĳɹ���');location.href='KS.SpaceSetting.asp';</script>")
			End If
			
			.Write "<html>"
			.Write "<title>�ռ��������</title>"
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
			.Write "<div class='topdashed sort'>�ռ��������</div>"
			.Write "<br>"
			.Write "<div class=tab-page id=spaceconfig>"
			.Write "  <form name='myform' id='myform' method=post action="""" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""spaceconfig"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>�ռ�����</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ռ�״̬��</strong></div><font color=#ff0000>���ѡ�񡰹رա���ôǰ̨ע���Ա���޷�ʹ�ÿռ�վ�㹦�ܡ�</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(0)"" value=""1"" "
				If Setting(0) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(0)"" value=""0"" "
				If Setting(0) = "0" Then .Write (" checked")
				.Write "> �ر�"

			
			.Write "     </td>"
			.Write "    </tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����ģʽ��</strong></div><font color=#ff0000>ѡ��α��̬������Ҫ��������װISAPI_Rewrite�����</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').hide();"" value=""0"" "
				If Setting(21) = "0" Then .Write (" checked")
				.Write "> ��̬ģʽ"
				.Write "    <input type=""radio"" name=""Setting(21)"" onclick=""$('#ext').show();"" value=""1"" "
				If Setting(21) = "1" Then .Write (" checked")
				.Write "> α��̬"

             If Setting(21)="1" Then
			  .Write "<div id='ext'>"
			 Else
			  .Write "<div id='ext' style='display:none'>"
			 End If
			.Write "α��̬��չ��:<input type='text' size='8' name='Setting(22)' value='" & Setting(22) & "'>,���Ĵ�����,��Ҫ�޸�ISAPI_Rewrite�������ļ�httpd.ini</div>"
			.Write "     </td>"
			.Write "    </tr>"

			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ƿ����ö���������</strong></div><font color=#ff0000>�˹��ܱ����Լ��ж�����������</font></td>"
			.Write "      <td width=""63%"" height=""30"">" 
			
				.Write " <input type=""radio"" name=""Setting(14)"" value=""1"" "
				If Setting(14) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(14)"" value=""0"" "
				If Setting(14) = "0" Then .Write (" checked")
				.Write "> ��<font color=red>(���رջ�֧�ֶ�����������ѡ���</font>"
			
			.Write "     </td>"
			.Write "    </tr>"
			
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ռ���ҳ������</strong><br><font color=red>�������Ҫ����������������Ч</font></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(15)"" size=15 value=""" & Setting(15) & """> <font color=blue>��:space.kesion.com</font>"
			 .Write "    </td>"
			 .Write "</tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ռ�վ�����������</strong><br><font color=red>�رն�����������������,������Ϊ�����������û�վ���������:user.space.kesion.com,�����ö����������û�վ���������:user.kesion.com</font></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(16)"" size=15 value=""" & Setting(16) & """> <font color=blue>��:��������:space.kesion.com���������kesion.com</font>"
			 .Write "    </td>"
			 .Write "</tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle""align=""right""><div><strong>��Աע���Ƿ��Զ�ע����˿ռ䣺</strong></div><font color=#ff0000>���ѡ���ǡ���ôע���Ա��ͬʱ��ͬʱӵ��һ�����˿ռ�վ�㡣</font></td>"
			 .Write "     <td height=""30""> "
			 	.Write " <input type=""radio"" name=""Setting(1)"" value=""1"" "
				If Setting(1) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(1)"" value=""0"" "
				If Setting(1) = "0" Then .Write (" checked")
				.Write "> ��"

			 .Write "</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td width=""32%"" height=""30"" class=""CleftTitle""> <div align='right'><strong>����ռ��Ƿ���Ҫ��ˣ�</strong></div></td>"
			 .Write "    <td height=""30"">"
			 
			 	.Write " <input type=""radio"" name=""Setting(2)"" value=""1"" "
				If Setting(2) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(2)"" value=""0"" "
				If Setting(2) = "0" Then .Write (" checked")
				.Write "> ��"

			 
			 .Write "      </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>������־�Ƿ���Ҫ��ˣ�</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	
				.Write " <input type=""radio"" name=""Setting(3)"" value=""1"" "
				If Setting(3) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(3)"" value=""0"" "
				If Setting(3) = "0" Then .Write (" checked")
				.Write "> ��"

			 .Write "       </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>������־�Ƿ������ϴ�������</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	
				.Write " <input type=""radio"" onclick=""$('#fj').show();"" name=""Setting(26)"" value=""1"" "
				If Setting(26) = "1" Then .Write (" checked")
				.Write "> ����"
				.Write "    <input type=""radio"" onclick=""$('#fj').hide();"" name=""Setting(26)"" value=""0"" "
				If Setting(26) = "0" Then .Write (" checked")
				.Write "> ������"
				If Setting(26) = "1" Then
                .Write "<div id='fj' style='color:blue'>"
				Else
                .Write "<div id='fj' style='display:none;color:blue'>"
				End If
				.Write "�����ϴ��ĸ�����չ��:<input type='text' value='" & Setting(27) & "' name='Setting(27)' /> �����չ���� |����,��gif|jpg|rar��</div>"
			 .Write "       </td>"
			 .Write "   </tr>"
			 
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������Ƿ���Ҫ��ˣ�</strong></div></td>"
			  .Write "    <td height=""30"">"
			  
			  	.Write " <input type=""radio"" name=""Setting(4)"" value=""1"" "
				If Setting(4) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(4)"" value=""0"" "
				If Setting(4) = "0" Then .Write (" checked")
				.Write "> ��"
			  
			  .Write "    </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����Ȧ���Ƿ���Ҫ��ˣ�</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(5)"" value=""1"" "
				If Setting(5) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(5)"" value=""0"" "
				If Setting(5) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�û������Ƿ���Ҫ��ˣ�</strong><br/><font color=red>���ú�,�û�������ֻ�о�����̨����Ա��˺�,ǰ̨�Ŀռ�ſ��Կ���</font></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(24)"" value=""1"" "
				If Setting(24) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(24)"" value=""0"" "
				If Setting(24) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>�����ο��ڿռ�������/���ԣ�</strong></div><font color=red>�������ò�����,������Ч��ֹһЩע�������</font></td>"
			  .Write "    <td height=""30"">"
			  
			  	.Write " <input type=""radio"" name=""Setting(25)"" value=""1"" "
				If Setting(25) = "1" Then .Write (" checked")
				.Write "> ����"
				.Write "    <input type=""radio"" name=""Setting(25)"" value=""0"" "
				If Setting(25) = "0" Then .Write (" checked")
				.Write "> ������"
			  
			  .Write "    </td>"
			 .Write "   </tr>"				
				
				

			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>ÿ����Ա������Ȧ�Ӹ�����</strong></div></td>"
				.Write "  <td height=""30"">"
				.Write " <input type=""text"" name=""Setting(6)"" style=""text-align:center"" size=5 value=""" & Setting(6) & """>��������������������롰0��"
				.Write "    </td>"
				.Write "</tr>"

				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��ģ�����ռ�ÿҳ��ʾ��</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(9)"" style=""text-align:center"" size=5 value=""" & Setting(9) & """> ��"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��ģ�������־ÿҳ��ʾ��</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(10)"" style=""text-align:center"" size=5 value=""" & Setting(10) & """> ƪ"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��ģ�����Ȧ��ÿҳ��ʾ��</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(11)"" style=""text-align:center"" size=5 value=""" & Setting(11) & """> ��"
			 .Write "    </td>"
			 .Write "</tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��ģ��������ÿҳ��ʾ��</strong></div></td>"
			 .Write "  <td height=""30"">"
			 .Write " <input type=""text"" name=""Setting(12)"" style=""text-align:center"" size=5 value=""" & Setting(12) & """> ����� ÿ����ʾ<input type=""text"" name=""Setting(13)"" style=""text-align:center"" size=5 value=""" & Setting(13) & """> ��"
			 .Write "    </td>"
			 .Write "</tr>"

			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=template-page>"
			.Write "  <H2 class=tab>�ռ�ģ��</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""template-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ռ���ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(7)"" id='Setting7' type=""text"" value=""" & Setting(7) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting7')[0]") & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ռ丱ģ�壺</strong></div><font color=#ff0000>�ռ�ĸ�ģ�壬������ʾ������־����ᡢȦ�ӵȣ����������ǩ��{$ShowMain}����</font></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(8)"" id='Setting8' type=""text"" value=""" & Setting(8) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting8')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>������ҳģ�壺</strong></div><font color=#ff0000>��Ӧ/space/friend/index.asp</font></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(23)"" id='Setting23' type=""text"" value=""" & Setting(23) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting23')[0]") & "</td>"
			.Write "    </tr>"
			 .Write " </table>"
			.Write " </div>"

			.Write " <div class=tab-page id=user-page>"
			.Write "  <H2 class=tab>��ҵ�ռ�����</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""user-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������Ϊ��ҵ�ռ���û��飺</strong></div><font color=red>������,�벻Ҫѡ</font></td>"
			.Write "      <td width=""63%"" height=""30""> &nbsp;" & KS.GetUserGroup_CheckBox("Setting(17)",Setting(17),5) & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>������ҵ�����Ƿ���Ҫ��ˣ�</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(18)"" value=""1"" "
				If Setting(18) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(18)"" value=""0"" "
				If Setting(18) = "0" Then .Write (" checked")
				.Write "> ��"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>������ҵ��Ʒ�Ƿ���Ҫ��ˣ�</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(19)"" value=""1"" "
				If Setting(19) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(19)"" value=""0"" "
				If Setting(19) = "0" Then .Write (" checked")
				.Write "> ��"
			.Write "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������֤���Ƿ���Ҫ��ˣ�</strong></div></td>"
			.Write "      <td width=""63%"" height=""30"">"
				.Write " <input type=""radio"" name=""Setting(20)"" value=""1"" "
				If Setting(20) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(20)"" value=""0"" "
				If Setting(20) = "0" Then .Write (" checked")
				.Write "> ��"
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
			.Write "{ alert('��ѡ��ռ���ҳģ��!');" & vbCrLf
			.Write "  $('#Setting7').focus();" & vbCrLf
			.Write "  return false;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "if ($('#Setting8').val()=='')" & vbCrLf
			.Write "{ alert('��ѡ��ռ丱ģ��!');" & vbCrLf
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

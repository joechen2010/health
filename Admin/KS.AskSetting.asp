<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Ask_Setting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Setting
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
			If Not KS.ReturnPowerResult(0, "WDXT10000") Then          '����Ƿ��л�����Ϣ���õ�Ȩ��
			  Call KS.ReturnErr(1, "")
			 .End
			End If
	
			SqlStr = "select AskSetting from KS_Config"
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1, 3
			
			Dim Setting:Setting=Split(RS(0)&"^%^0^%^0^%^0^%^0^%^0","^%^")
			If KS.G("Flag") = "Edit" Then
			    Dim N					
			    Dim WebSetting
				For n=0 To 41
				   WebSetting=WebSetting & Replace(KS.G("Setting(" & n &")"),"^%^","") & "^%^"
				Next
				RS("AskSetting")=WebSetting
				RS.Update				
				.Write ("<script>alert('�ʴ�����޸ĳɹ���');location.href='KS.AskSetting.asp';</script>")
			End If
			
			.Write "<html>"
			.Write "<title>�ʴ��������</title>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write "<script src=""../KS_Inc/jQuery.js"" language=""JavaScript""></script>"
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

			.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>�ʴ��������</div>"
			.Write "<br>"
			.Write "<div class=tab-page id=spaceconfig>"
			.Write "  <form name='myform' id='myform' method=post action="""" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""spaceconfig"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>��������</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ʴ�ϵͳ״̬��</strong></div></td>"
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
			 .Write "     <td height=""30"" class=""clefttitle""align=""right""><div><strong>��װĿ¼��</strong></div><font color=#ff0000>��ask��,������""/""������</font></td>"
			 .Write "     <td height=""30""> "
			 	.Write " <input type=""text"" name=""Setting(1)"" size=""20"" value=""" & Setting(1) & """>"

			 .Write "</td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle""align=""right""><div><strong>ģ�����ƣ�</strong></div><font color=#ff0000>��""�ʰ�""�ȡ�</font></td>"
			 .Write "     <td height=""30""> "
			 	.Write " <input type=""text"" name=""Setting(2)"" size=""20"" value=""" & Setting(2) & """>"

			 .Write "</td>"
			 .Write "   </tr>"
			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>����ģʽ��</strong></div></td>"
			 .Write "     <td height=""30"">"
			 	
				.Write " <input type=""radio"" name=""Setting(16)"" value=""0"" "
				If Setting(16) = "0" Then .Write (" checked")
				.Write "> ��̬"
				.Write "    <input type=""radio"" name=""Setting(16)"" value=""1"" "
				If Setting(16) = "1" Then .Write (" checked")
				.Write "> α��̬(<font color=red>��Ҫ������֧��Rewrite���</font>)"

			 .Write "<div>��չ��<input type='text' name='Setting(17)' value='" & Setting(17) & "' size='10'>���Ĵ�����,��Ҫ�޸�ISAPI_Rewrite�������ļ�httpd.ini</div>"
			 .Write "       </td>"
			 .Write "   </tr>"
			 
			
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>�Ƿ������ʣ�</strong></div></td>"
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
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������/�ش𳤶ȿ��ƣ�</strong></div></td>"
			  .Write "    <td height=""30"">"
			  
			  .Write "���ڵ���<input type=""text"" name=""Setting(4)"" size=""5"" value=""" & Setting(4) & """>С�ڵ���<input type=""text"" name=""Setting(5)"" size=""5"" value=""" & Setting(5) & """> <font color=blue>�������,����д0</font>"
			  
			  .Write "    </td>"
			 .Write "   </tr>"

			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�������Ƿ�������֤�룺</strong></div></td>"
			 .Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(6)"" value=""1"" "
				If Setting(6) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(6)"" value=""0"" "
				If Setting(6) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
				
				.Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>������������õ���Ч������</strong></div></td>"
			    .Write "  <td height=""30""><input type='text' name='Setting(41)' value='" & Setting(41) & "' style='text-align:center;width:50px'>��"
				.Write "  </td></tr>"

				
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ƿ�����ش�</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(7)"" value=""1"" "
				If Setting(7) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(7)"" value=""0"" "
				If Setting(7) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ش������Ƿ�������֤�룺</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(8)"" value=""1"" "
				If Setting(8) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(8)"" value=""0"" "
				If Setting(8) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
				
				
				
				
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ƿ�ֻ�ܻش�һ�Σ�</strong></div><font color=red>��ÿ���������Ƿ�ÿ����ֻ�ܻش�һ��</font></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(9)"" value=""1"" "
				If Setting(9) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(9)"" value=""0"" "
				If Setting(9) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�������Ƿ��������ⲹ�䣺</strong></div><font color=red>�����˿��԰����ⲹ��</font></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(10)"" value=""1"" "
				If Setting(10) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(10)"" value=""0"" "
				If Setting(10) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�����˿��Իش��Լ����ʵ����⣺</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(11)"" value=""1"" "
				If Setting(11) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(11)"" value=""0"" "
				If Setting(11) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�����˿���ɾ���û��Ļش�</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(12)"" value=""1"" "
				If Setting(12) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(12)"" value=""0"" "
				If Setting(12) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			    .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ƿ������οͻش����⣺</strong></div></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""radio"" name=""Setting(13)"" value=""1"" "
				If Setting(13) = "1" Then .Write (" checked")
				.Write "> ��"
				.Write "    <input type=""radio"" name=""Setting(13)"" value=""0"" "
				If Setting(13) = "0" Then .Write (" checked")
				.Write "> ��"
				
				.Write "    </td>"
				.Write "</tr>"
				
				
				
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			   .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�б�ҳÿҳ��ʾ������</strong></div><font color=blue>��Ӧǰ̨��showlist.asp</a></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""text"" name=""Setting(14)"" value=""" & Setting(14) & """ size=""6"">��"
				
				.Write "    </td>"
				.Write "</tr>"
			    .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			   .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������ÿҳ��ʾ������</strong></div><font color=blue>��Ӧǰ̨��q.asp</a></td>"
				.Write "  <td height=""30"">"
				
				.Write " <input type=""text"" name=""Setting(15)"" value=""" & Setting(15) & """ size=""6"">��"
				
				.Write "    </td>"
				.Write "</tr>"
				


			

			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=template-page>"
			.Write "  <H2 class=tab>�ʴ�ģ��</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""template-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�ʴ���ҳģ�壺</strong><br />(<a href='../" & KS.ASetting(1) & "index.asp' target='_blank' style='color:blue'>index.asp</a>)</div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(20)"" id=""Setting20"" type=""text"" value=""" & Setting(20) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting20')[0]") & "</td>"
			.Write "    </tr>"            
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������ģ�壺</strong><br />(<a href='../" & KS.ASetting(1) & "a.asp' target='_blank' style='color:blue'>a.asp</a>)</div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(21)"" id=""Setting21"" type=""text"" value=""" & Setting(21) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting21')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�����б�ҳģ�壺</strong><br />(<a href='../" & KS.ASetting(1) & "showlist.asp' target='_blank' style='color:blue'>showlist.asp</a>)</div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(22)"" id=""Setting22"" type=""text"" value=""" & Setting(22) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting22')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������ҳģ�壺</strong><br />(<a href='../" & KS.ASetting(1) & "q.asp' target='_blank' style='color:blue'>q.asp</a>)</div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(23)"" id=""Setting23"" type=""text"" value=""" & Setting(23) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting23')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������ҳģ�壺</strong><br />(<a href='../" & KS.ASetting(1) & "search.asp' target='_blank' style='color:blue'>search.asp</a>)</div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(24)"" id=""Setting24"" type=""text"" value=""" & Setting(24) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting24')[0]") & "</td>"
			.Write "    </tr>"
			
			
			 .Write " </table>"
			.Write " </div>"

			.Write " <div class=tab-page id=user-page>"
			.Write "  <H2 class=tab>��������</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""user-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>�û��ش�һ���������õĻ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(30)"" size=""10"" value=""" & Setting(30) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>�ش������߲��������⽱���Ļ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(31)"" size=""10"" value=""" & Setting(31) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>�û������������õĻ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(32)"" size=""10"" value=""" & Setting(32) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>���ⱻѡΪ�����Ƽ����������õĻ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(33)"" size=""10"" value=""" & Setting(33) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>���ⱻѡΪ�����Ƽ���ѻش������õĻ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(34)"" size=""10"" value=""" & Setting(34) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>�û�����һ����������Ļ��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(35)"" size=""10"" value=""" & Setting(35) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>�������ʼ�ȥ���֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(36)"" size=""10"" value=""" & Setting(36) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>ɾ���𰸼�ȥ�ش��߻��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(37)"" size=""10"" value=""" & Setting(37) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>ɾ����Ѵ𰸼�ȥ�ش��߻��֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(38)"" size=""10"" value=""" & Setting(38) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>ɾ��δ��������ȥ���֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(39)"" size=""10"" value=""" & Setting(39) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" width=""32%"" class=""clefttitle""align=""right""><div><strong>ɾ���ѽ�������ȥ���֣�</strong></div></td>"
			 .Write "     <td height=""30""> "
			 .Write "       <input type=""text"" name=""Setting(40)"" size=""10"" value=""" & Setting(40) & """> ��"
			 .Write "     </td>"
			 .Write "   </tr>"
			
			.Write " </table>"
			.Write " </div>"
			

			.Write "<div style=""text-align:center;color:#003300"">--------------------------------------------------------------------------------<br/>KeSion CMS V 6.5, Copyright (c) 2006-2010 KeSion.Com. All Rights Reserved . </div>"

			.Write " </body>"
			.Write " </html>"
			.Write " <Script Language=""javascript"">"
			.Write " <!--" & vbCrLf
			
			.Write " function CheckForm()" & vbCrLf
			.Write " {" & vbCrLf
			.Write "if ($('#Setting20').val()=='')" & vbCrLf
			.Write "{ alert('��ѡ���ʴ���ҳģ��!');" & vbCrLf
			.Write "  $('#Setting20').focus();" & vbCrLf
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

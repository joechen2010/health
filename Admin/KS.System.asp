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
			.Write "<title>��վ������������</title>"
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
	
		'ϵͳ������Ϣ����
		Sub SetSystem()
		Dim SqlStr, RS, InstallDir, FsoIndexFile, FsoIndexExt
		With Response
			
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '����Ƿ��л�����Ϣ���õ�Ȩ��
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
			            'IP����
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
			   .Write ("<script>alert('��վ������Ϣ�޸ĳɹ���');parent.frames['FrameTop'].location.href='index.asp?Action=Head&C=1';location.href='KS.System.asp';</script>")
			End If
			

			.Write "<body oncontextmenu='return false' bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.Write "<div class='topdashed sort'>��վ������Ϣ����</div>"
			.Write "<div style='height:5px;overflow:hidden'></div>"
			.Write "<div class=tab-page id=configPane>"
			.Write "  <form name='myform' method=post action="""" id=""myform"" onSubmit=""return(CheckForm())"">"
            .Write " <SCRIPT type=text/javascript>"
            .Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""configPane"" ), 1 )"
            .Write " </SCRIPT>"
             
			.Write " <div class=tab-page id=site-page>"
			.Write "  <H2 class=tab>������Ϣ</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "<input type=""hidden"" value=""Edit"" name=""Flag"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��վ���ƣ�</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(0)"" type=""text"" id=""Setting(0)"" value=""" & Setting(0) & """ size=""30""></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""clefttitle"" align=""right""><div><strong>��վ���⣺</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(1)"" type=""text"" id=""Setting(1)"" value=""" & Setting(1) & """ size=""30""></td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			  .Write "    <td width=""32%"" height=""30"" class=""CleftTitle""> <div align='right'><strong>��վ��ַ��</strong></div><font color=""#FF0000"">ϵͳ���Զ������ȷ��·��������Ҫ�ֹ���������</font></td>"
			 .Write "    <td height=""30""> <input name=""Setting(2)"" type=""text""  value=""" &KS.GetAutoDomain & """ size=""30"">"
			 .Write "      (��ʹ��http://��ʶ),���治Ҫ��&quot;/&quot;���� </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align='right'> <div><strong>��װĿ¼��</strong></div><font color=""#FF0000"">ϵͳ���Զ������ȷ��·��������Ҫ�ֹ���������</font></td>"
			 .Write "     <td height=""30""> <input name=""Setting(3)"" type=""text"" id=""Setting(3)""  value=""" & InstallDir & """ readonly size=30>"
			 .Write "       ϵͳ��װ������Ŀ¼</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td height=""30"" class=""CleftTitle"" align=""right""><div><strong>��վLogo��ַ��</strong></div></td>"
			  .Write "    <td height=""30""><input name=""Setting(4)"" type=""text"" id=""Setting(4)""   value=""" & Setting(4) & """ size=30>"
			  .Write "      ������������ʱ��ʾ���û�</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>���ɵ���վ��ҳ��</strong></div></td>"
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
				.Write "    </select>&nbsp;<font color=blue>��չ��Ϊ.asp����ҳ�����������ɾ�̬HTML�Ĺ���</font></td>"
				.Write "</tr>"
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>ר���Ƿ��������ɣ�</strong></div></td>"
				.Write "  <td height=""30""><input type=""radio"" name=""Setting(78)"" value=""1"" "
				
				If Setting(78) = "1" Then .Write (" checked")
				.Write ">����"
				.Write "    <input type=""radio"" name=""Setting(78)"" value=""0"" "
				If Setting(78) = "0" Then .Write (" checked")
				.Write ">������"
			   .Write "  ��</td>"
			   .Write "    </tr>"
			
				.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "  <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>Ĭ�������ϴ�����ļ���С��</strong></div></td>"
				.Write "  <td height=""30""><input name=""Setting(6)"" onBlur=""CheckNumber(this,'�����ϴ�����ļ���С');"" type=""text"" id=""Setting(6)""   value=""" & Setting(6) & """ size=15>"
			.Write "KB �� <span class=""STYLE2"">��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB</span></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle"" align=""right""><div><strong>Ĭ�������ϴ��ļ����ͣ�</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(7)"" type=""text"" id=""Setting(7)""   value=""" & Setting(7) & """ size='30'><font color=red> ���������|�߸���</font></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align=""right""><strong>ɾ������û�ʱ�䣺</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(8)"" type=""text""  value=""" &  Setting(8) & """ size=8> ����</td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align=""right""><strong>�����Զ���ҳÿҳ��Լ�ַ�����</strong></div></td>"
			.Write "      <td height=""30""><input name=""Setting(9)"" type=""text"" value=""" & Setting(9) & """ size=8> ���ַ�&nbsp;&nbsp;<font color=red>��������Զ���ҳ��������""0""</font></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CLeftTitle"" align=""right""><div><strong>վ��������</strong></div></td>"
			.Write "      <td height=""30""> <input name=""Setting(10)"" type=""text""   value=""" & Setting(10) & """ size=30></td>"
			.Write "    </tr>"


			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>Ҫ���εĹؼ��֣�</strong></div><font color=red>˵���������ַ��趨����Ϊ Ҫ���˵��ַ�=���˺���ַ� ��ÿ�������ַ��ûس��ָ�����÷�Χ����ģ�͵����ݡ����ۡ��ʴ�С��̳�ȡ�</font></td>"
			 .Write "    <td height=""30""><textarea name=""Setting(55)"" cols=""30"" rows=""6"">" & Setting(55) & "</textarea></td></tr>"

			 
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class='CleftTitle' align=""right""><div><strong>ҳ�淢��ʱ������Ϣ��</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(15)"" type=""text""  value=""" & Setting(15) & """ size=30>"
			 .Write "     ��д<span class=""STYLE1"">&quot;0&quot;</span>������ʾ</td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""cleftTitle"" align=""right""><div><strong>�ٷ���Ϣ��ʾ��</strong></div></td>"
			 .Write "  <td height=""30""> <input type=""checkbox"" name=""Setting(16)"" value=""1"" "
				
				If instr(Setting(16),"1")>0 Then .Write (" checked")
				.Write ">"
				.Write "    ��ʾ��������"
				.Write "    <input type=""checkbox"" name=""Setting(16)"" value=""2"" "
				If instr(Setting(16),"2")>0 Then .Write (" checked")
				.Write ">"
				.Write "    ��ʾ��̳����"

			 .Write "     </td>"
			 .Write "   </tr>"
			 .Write "   <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>�ٷ���Ȩ��Ψһϵ�кţ�</strong></div></td>"
			 .Write "     <td height=""30""> <input name=""Setting(17)"" type=""text""  value=""" & Setting(17) & """ size=30>"
			 .Write "     ��Ѱ�����д<span class=""STYLE1"">&quot;0&quot;</span></td>"
			 .Write "   </tr>"
			   
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>��վ�İ�Ȩ��Ϣ��</strong></div><font color=""#FF0000""> ������ʾ��վ�汾�ȣ�֧��html�﷨</font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(18)"" cols=""60"" rows=""5"">" & Setting(18) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>��վMETA�ؼ��ʣ�</strong></div><font color=""#FF0000""> ��������������õ���ҳ�ؼ���,����ؼ�������,�ŷָ� </font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(19)"" cols=""60"" rows=""5"">" & Setting(19) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write "     <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "     <td width=""32%"" height=""30"" class=""clefttitle"" align=""right""><div><strong>��վMETA��ҳ������</strong></div><font color=""#FF0000""> ��������������õ���ҳ����,�����������,�ŷָ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  </font></td>"
			  .Write "    <td height=""30""> <textarea name=""Setting(20)"" cols=""60"" rows=""5"">" & Setting(20) & "</textarea></td>"
			 .Write "   </tr>"
			 .Write " </table>"
			 .Write "</div>"
			 
			.Write " <div class=tab-page id=site-template>"
			.Write "  <H2 class=tab>ģ���</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""site-template"" ) );"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��վ��ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(110)"" id=""Setting110"" type=""text"" value=""" & Setting(110) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting110')[0]") & " <a href='../index.asp' target='_blank' style='color:green'>ҳ��:/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr  valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>ȫվtagsģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(120)"" id=""Setting120"" type=""text"" value=""" & Setting(120) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting120')[0]") & " <a href='../plus/tags.asp' target='_blank' style='color:green'>ҳ��:/plus/tags.asp</a></td>"
			.Write "    </tr>"			
			.Write "    <tr  valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>ȫվ����ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(139)"" id=""Setting139"" type=""text"" value=""" & Setting(139) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting139')[0]") & " <a href='../plus/search/' target='_blank' style='color:green'>ҳ��:/plus/search/</a></td>"
			.Write "    </tr>"			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>ר����ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(111)"" id=""Setting111"" type=""text"" value=""" & Setting(111) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting111')[0]") & " <a href='../specialindex.asp' target='_blank' style='color:green'>ҳ��:/specialindex.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(112)"" id=""Setting112"" type=""text"" value=""" & Setting(112) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting112')[0]") & " <a href='#' style='color:green'>ҳ��:/plus/announce/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��������ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(113)"" id=""Setting113"" type=""text"" value=""" & Setting(113) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting113')[0]") & " <a href='../plus/link/' target='_blank' style='color:green'>ҳ��:/plus/link</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����/С��̳��ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(114)"" id=""Setting114"" type=""text"" value=""" & Setting(114) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting114')[0]") & " <a href='../club/index.asp' target='_blank' style='color:green'>ҳ��:/club/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����/С��̳ǩдҳ��ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(115)"" id=""Setting115"" type=""text"" value=""" & Setting(115) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting115')[0]") & " <a href='../club/post.asp' target='_blank' style='color:green'>ҳ��:/club/post.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PK��ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(102)"" id=""Setting102"" type=""text"" value=""" & Setting(102) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting102')[0]") & " <a href='../plus/pk/index.asp' target='_blank' style='color:green'>ҳ��:/plus/pk/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PKҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(103)"" id=""Setting103"" type=""text"" value=""" & Setting(103) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting103')[0]") & " <a href='#' style='color:green'>ҳ��:/plus/pk/pk.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>PK�۵����ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(104)"" id=""Setting104"" type=""text"" value=""" & Setting(104) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting104')[0]") & " <a href='#' style='color:green'>ҳ��:/plus/pk/more.asp</a></td>"
			.Write "    </tr>"

			

			.Write "    <tr>"
			.Write "      <td colspan=2 height='1' bgcolor='green'></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��Ա��ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(116)"" id=""Setting116"" type=""text"" value=""" & Setting(116) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting116')[0]") & " <a href='../user/' target='_blank' style='color:green'>ҳ��:/user/index.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��Աע���1ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(117)"" id=""Setting117"" type=""text"" value=""" & Setting(117) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting117')[0]") & " <a href='../user/reg/' target='_blank' style='color:green'>ҳ��:/user/reg/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��Աע���2ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(118)"" id=""Setting118"" type=""text"" value=""" & Setting(118) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting118')[0]") & " <a href='../user/reg/' target='_blank' style='color:green'>ҳ��:/user/reg/</a></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>��Աע��ɹ�ҳģ�壺</strong></div></td>"
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
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳ǹ��ﳵģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(121)"" id=""Setting121"" type=""text"" value=""" & Setting(121) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting121')[0]") & " <a href='../shop/shoppingcart.asp' target='_blank' style='color:green'>ҳ��:/shop/shoppingcart.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳�����̨ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(122)"" id=""Setting122"" type=""text"" value=""" & Setting(122) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting122')[0]") & " <a href='../shop/payment.asp' target='_blank' style='color:green'>ҳ��:/shop/payment.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳Ƕ���Ԥ��ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(123)"" id=""Setting123"" type=""text"" value=""" & Setting(123) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting123')[0]") & " <a href='../shop/Preview.asp' target='_blank' style='color:green'>ҳ��:/shop/Preview.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳Ƕ����ɹ�ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(124)"" id=""Setting124"" type=""text"" value=""" & Setting(124) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting124')[0]") & " <a href='../shop/order.asp' target='_blank' style='color:green'>ҳ��:/shop/order.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳ǹ���ָ��ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(125)"" id=""Setting125"" type=""text"" value=""" & Setting(125) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting125')[0]") & " <a href='../shop/ShopHelp.asp' target='_blank' style='color:green'>ҳ��:/shop/ShopHelp.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳����и���ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(126)"" id=""Setting126"" type=""text"" value=""" & Setting(126) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting126')[0]") & " <a href='../shop/showpay.asp' target='_blank' style='color:green'>ҳ��:/shop/showpay.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳�Ʒ���б�ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(135)"" id=""Setting135"" type=""text"" value=""" & Setting(135) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting135')[0]") & " <a href='../shop/showbrand.asp' target='_blank' style='color:green'>ҳ��:/shop/showbrand.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳�Ʒ������ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(136)"" id=""Setting136"" type=""text"" value=""" & Setting(136) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting136')[0]") & " <a href='../shop/search_list.asp' target='_blank' style='color:green'>ҳ��:/shop/search_list.asp</a></td>"
			.Write "    </tr>"
			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳��Ź���ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(137)"" id=""Setting137"" type=""text"" value=""" & Setting(137) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting137')[0]") & " <a href='../shop/groupbuy.asp' target='_blank' style='color:green'>ҳ��:/shop/groupbuy.asp</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�̳��Ź�����ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(138)"" id=""Setting138"" type=""text"" value=""" & Setting(138) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting138')[0]") & " <a href='../shop/groupbuyshow.asp' target='_blank' style='color:green'>ҳ��:/shop/groupbuyshow.asp</a></td>"
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
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>������ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(127)"" id=""Setting127"" type=""text"" value=""" & Setting(127) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting127')[0]") & " <a href='../music/' target='_blank' style='color:green'>ҳ��:/music/</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����ҳ��ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(128)"" id=""Setting128"" type=""text"" value=""" & Setting(128) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting128')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>ר���б�ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(129)"" id=""Setting129"" type=""text"" value=""" & Setting(129) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting129')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����ר��ҳģ�壺</strong></div></td>"
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
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>����ϵͳ��ҳģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(131)""  id=""Setting131"" type=""text"" value=""" & Setting(131) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting131')[0]") & " <a href='../mnkc/' target='_blank' style='color:green'>ҳ��:/mnkc/</a></td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ծ����ҳ��ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(132)"" id=""Setting132"" type=""text"" value=""" & Setting(132) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting132')[0]") & "</td>"
			.Write "    </tr>"
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ծ�����ҳ��ģ��(���⿨��ʽ)��</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(133)"" id=""Setting133"" type=""text"" value=""" & Setting(133) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting133')[0]") & "</td>"
			.Write "    </tr>"
            .Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ծ�����ҳ��ģ��(��ͨ��ʽ)��</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(105)"" id=""Setting105"" type=""text"" value=""" & Setting(105) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting105')[0]") & "</td>"
			.Write "    </tr>"			
			.Write "    <tr" & dis &" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30"" class=""CleftTitle"" align=""right""><div><strong>�Ծ��ܷ���ģ�壺</strong></div></td>"
			.Write "      <td width=""63%"" height=""30""> <input name=""Setting(134)"" id=""Setting134"" type=""text"" value=""" & Setting(134) & """ size=""30"">&nbsp;" & KSMCls.Get_KS_T_C("$('#Setting134')[0]") & " <a href='../mnkc/all.html' target='_blank' style='color:green'>ҳ��:/mnkc/all.html</a></td>"
			.Write "    </tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			 '=================================================��ע���ѡ��========================================
			 .Write "<div class=tab-page id=ZCJ_Option>"
			 .Write " <H2 class=tab>��ע���</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""ZCJ_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			
             .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>Ҫ���÷�ע�����ҳ�棺</strong></div></td>"
			
			
			.Write "      <td height=""21"">"
			.Write "<input type='checkbox' name='Opening1' value='1'"
			If mid(Setting(161),1,1)="1" Then .Write "checked"
			.Write ">��Աע��ҳ��"
			.Write "<br/><input type='checkbox' name='Opening2' value='1'"
			If mid(Setting(161),2,1)="1" Then .Write "checked"
			.Write ">����Ͷ�巢��ҳ��"
			'.Write "<br/><input type='checkbox' name='Opening3' value='1'"
			'If mid(Setting(161),3,1)="1" Then .Write "checked"
			'.Write ">���۷���ҳ��"
		    .Write "      </td>"	
			.Write "</tr>"			
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>��֤���⣺</strong></div>�������ö��,һ��һ����֤ѡ��</td>"
            .Write "    <td><textarea name='Setting(162)' style='width:350px;height:120px'>" & Setting(162) & "</textarea></td>"
			.Write "    </tr>"
            .Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>��֤�𰸣�</strong></div>��Ӧ��֤�����ѡ��,һ��һ����֤��</td>"
            .Write "    <td><textarea name='Setting(163)' style='width:350px;height:120px'>" & Setting(163) & "</textarea></td>"
			.Write "    </tr>"
			.Write "  </table>"
			.Write "</div>"
			
			
			
									 '=====================================================��Աע��������ÿ�ʼ=========================================

		.Write " <div class=tab-page id=User_Option>"
		.Write "	  <H2 class=tab>��Աѡ��</H2>"
		.Write "		<SCRIPT type=text/javascript>"
		.Write "					 tabPane1.addTabPage(document.getElementById( ""User_Option"" ));"
		.Write "		</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width='40%' height=""21"" class=""clefttitle"" align=""right""><div><strong>�Ƿ������»�Աע�᣺</strong></div></td>"
			.Write "      <td height=""21""><input name=""Setting(21)"" type=""radio"" value=""1"""
			 If Setting(21)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(21)"" type=""radio"" value=""0"""
			 If Setting(21)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"	
			 .Write "</tr>"		
			 .Write "<tr style=""display:none"" valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='clefttitle' align=""right""><div><strong>�»�Աע����Ҫ�Ķ���ԱЭ�飺</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio""  value=""1"""
			 If Setting(22)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(22)"" onClick=""setlience(this.value);"" type=""radio"" value=""0"""
			 If Setting(22)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""liencearea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='clefttitle' align=""right""><div><strong>�»�Աע����������������</strong><div><div align=center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ǩ˵����</div>{$GetSiteName}����վ����<br>{$GetSiteUrl}����վURL<br>{$GetWebmaster}��վ��<br>{$GetWebmasterEmail}��վ������</td>"
			.Write "      <td height=""21""><textarea name=""Setting(23)"" cols=""70"" rows=""7"">" & Setting(23) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			
			
			 .Write "<tr width=""32%"" height=""21"" id=""grouparea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			 .Write "<td width='40%' class='CleftTitle' align=""right""><div><strong>�Ƿ������û���ע�᣺</strong></div><font color=red>���������,Ĭ��ע������Ϊ���˻�Ա</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(33)"" type=""radio"" value=""1"""
			 If Setting(33)="1" Then .Write " Checked"
			 .Write ">����"
			 .Write " &nbsp;&nbsp;<input name=""Setting(33)"" type=""radio"" value=""0"""
			 If Setting(33)="0" Then .Write " Checked"
			 .Write ">������"
			 .Write "</td>"
			 .Write "</tr>" 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Աע�����̣�</strong> </div></td>"
			.Write "      <td height=""21""> <input name=""Setting(32)"" type=""radio"" value=""1"""
			 If Setting(32)="1" Then .Write " Checked"
			 .Write ">һ�������ע��<br>"
			 .Write "<input name=""Setting(32)"" type=""radio"" value=""2"""
			 If Setting(32)="2" Then .Write " Checked"
			 .Write ">�������ע�ᣨ��Ҫ��д��Ӧ�û���ı���"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Աע��ɹ��Ƿ��ʼ�֪ͨ��</strong></div><font color=blue>�û������ó���Ҫ�ʼ���֤ʱ,ֻ�м���ɹ��Żᷢ�͡�</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(146)"" onclick=""setsendmail(1)"" type=""radio"" value=""1"""
			 If Setting(146)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(146)"" type=""radio"" onclick=""setsendmail(0)"" value=""0"""
			 If Setting(146)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" id=""sendmailarea""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Աע��ɹ����͵��ʼ�֪ͨ���ݣ�</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div><div align=center>��ǩ˵����<br>{$UserName}���û���<br>{$PassWord}������<br>{$SiteName}����վ����<div></td>"
			.Write "      <td height=""21""><textarea name=""Setting(147)"" cols=""70"" rows=""5"">" & Setting(147) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>ע���Ա���������Ƿ���</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(148)"" type=""radio"" value=""1"""
			 If Setting(148)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(148)"" type=""radio"" value=""0"""
			 If Setting(148)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"			 
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>ע���Ա�ֻ������Ƿ���</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(149)"" type=""radio"" value=""1"""
			 If Setting(149)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(149)"" type=""radio"" value=""0"""
			 If Setting(149)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>һ��IPֻ��ע��һ����Ա��</strong></div><font color=red>��ѡ���ǣ���ôһ��IP��ַֻ��ע��һ��</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(26)"" type=""radio"" value=""1"""
			 If Setting(26)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(26)"" type=""radio"" value=""0"""
			 If Setting(26)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Աע��ʱ�Ƿ�������֤�빦�ܣ�</strong></div><font color=red>������֤�빦�ܿ�����һ���̶��Ϸ�ֹ����Ӫ�������ע����Զ�ע��</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(27)"" type=""radio"" value=""1"""
			 If Setting(27)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(27)"" type=""radio"" value=""0"""
			 If Setting(27)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>ÿ��Email�Ƿ�����ע���Σ�</strong></div><font color=red>��ѡ���ǣ�������ͬһ��Email����ע������Ա��</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(28)"" type=""radio"" value=""1"""
			 If Setting(28)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(28)"" type=""radio"" value=""0"""
			 If Setting(28)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>�»�Աע��ʱ�û�����</strong></div></td>"
			.Write "      <td height=""21""> �����ַ���<input name=""Setting(29)"" type=""text"" onBlur=""CheckNumber(this,'�û�����С�ַ���');"" size=""3"" value=""" & Setting(29) & """>���ַ�  ����ַ���<input name=""Setting(30)"" type=""text"" onBlur=""CheckNumber(this,'�û�������ַ���');"" size=""3"" value=""" & Setting(30)& """>���ַ�"
			.Write "       </td>" 
	        .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��ֹע����û�����</strong> </div><font color=red>���ұ�ָ�����û���������ֹע�ᣬÿ���û������á�|�����ŷָ�</font></td>"
			.Write "      <td height=""21""> <textarea name=""Setting(31)"" cols=""50"" rows=""3"">" & Setting(31) & "</textarea>"
			.Write "       </td>" 
			.Write "</tr>" 
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա��¼ʱ�Ƿ�������֤�빦�ܣ�</strong></div><font color=red>������֤�빦�ܿ�����һ���̶��Ϸ�ֹ��Ա���뱻�����ƽ�</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(34)"" type=""radio"" value=""1"""
			 If Setting(34)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(34)"" type=""radio"" value=""0"""
			 If Setting(34)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			 .Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>ֻ����һ���˵�¼�� </strong></div><font color=red>���ô˹��ܿ�����Ч��ֹһ����Ա�˺Ŷ���ʹ�õ����</font></td>"
			.Write "      <td height=""21""> <input name=""Setting(35)"" type=""radio"" value=""1"""
			 If Setting(35)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(35)"" type=""radio"" value=""0"""
			 If Setting(35)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
             .Write "</tr>"

			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>�»�Աע��ʱ���͵��ʽ�</strong>��</div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'�»�Աע��ʱ���͵Ľ�Ǯ');"" name=""Setting(38)"" type=""text"" size=""5"" value=""" & Setting(38) & """>"
			.Write "Ԫ����ң�Ϊ0ʱ�����ͣ�,���ʽ�������̳����Ĺ���.</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>�»�Աע��ʱ���͵Ļ��֣�</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(39)"" onBlur=""CheckNumber(this,'�»�Աע��ʱ���͵Ļ���');"" type=""text"" size=""5"" value=""" & Setting(39) & """>"
			.Write "�ֻ���</td>"
	        .Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>�»�Աע��ʱ���͵ĵ�ȯ��</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'�»�Աע��ʱ���͵ĵ�ȯ');"" name=""Setting(40)"" type=""text"" size=""5"" value=""" & Setting(40) & """>"
			.Write "���ȯ��Ϊ0ʱ�����ͣ�<br/><font color=blue>����û���ѡ���˿۵��û�,��������Ĭ�ϵ���,�����û����������Ϊ׼</font></td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա�Ļ������ȯ�Ķһ����ʣ�</strong> </div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'��Ա�Ļ������ȯ�Ķһ�����');"" name=""Setting(41)"" type=""text"" size=""5"" value=""" & Setting(41) & """>"
			.Write "�ֻ��ֿɶһ� <font color=red>1</font> ���ȯ</td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա�Ļ�������Ч�ڵĶһ����ʣ�</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'��Ա�Ļ�������Ч�ڵĶһ�����');"" name=""Setting(42)"" type=""text"" size=""5"" value=""" & Setting(42) & """>"
			.Write "�ֻ��ֿɶһ� <font color=red>1</font> ����Ч��</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա���ʽ����ȯ�Ķһ�����</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'��Ա���ʽ����ȯ�Ķһ�����');"" name=""Setting(43)"" type=""text"" size=""5"" value=""" & Setting(43) & """>"
			.Write "Ԫ����ҿɶһ� <font color=red>1</font> ���ȯ</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա���ʽ�����Ч�ڵĶһ�����</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'��Ա���ʽ�����Ч�ڵĶһ�����');"" name=""Setting(44)"" type=""text"" size=""5"" value=""" & Setting(44) & """>"
			.Write "Ԫ����ҿɶһ� <font color=red>1</font> ����Ч��</td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><strong>��ȯ���ã�</strong></td>"
			.Write "      <td height=""21""> ����<input name=""Setting(45)"" type=""text"" size=""5"" value=""" & Setting(45) & """><font color=red>���磺��Ѵ�ҡ���ȯ�����</font>  ��λ<input name=""Setting(46)"" type=""text"" size=""5"" value=""" & Setting(46) & """> <font color=red>���磺�㡢��</font>"
			.Write "</td>"
			.Write "</tr>"

			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21""  class='CleftTitle' align=""right""><div><strong>��Ավ�ڶ������ã�</strong></div></td>"
			.Write "      <td height=""21""> �������Ϊ<input onBlur=""CheckNumber(this,'����д��Ч����!');"" name=""Setting(47)"" type=""text"" size=""5"" value=""" & Setting(47) & """>��,������������ַ���<input onBlur=""CheckNumber(this,'������������ַ���');"" name=""Setting(48)"" type=""text"" size=""5"" value=""" & Setting(48) & """>���ַ� Ⱥ����������<input onBlur=""CheckNumber(this,'Ⱥ����������');"" name=""Setting(49)"" type=""text"" size=""5"" value=""" & Setting(49) & """>��"
			.Write "</td>"	
			.Write "</tr>"		
			.Write "    <tr style='display:none' valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>��Ա���ÿռ��С��</strong></div></td>"
			.Write "      <td height=""21""><input onBlur=""CheckNumber(this,'����д��Ч����!');"" name=""Setting(50)"" type=""text"" size=""5"" value=""" & Setting(50) & """> KB &nbsp;&nbsp;<font color=#ff6600>��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB</font>"
			.Write "</td>"	
			.Write "</tr>"	
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>�ƹ�ƻ����ã�</strong></div><br><a href='KS.PromotedPlan.asp'><font color=red>�鿴�ƹ��¼</font></a>&nbsp;</td>"
			.Write "      <td height=""21"">"
			.Write " <FIELDSET align=center><LEGEND align=left>�ƹ�ƻ�</LEGEND>�Ƿ������ƹ㣺"
			.Write " <input name=""Setting(140)"" type=""radio"" value=""1"""
			 If Setting(140)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(140)"" type=""radio"" value=""0"""
			 If Setting(140)="0" Then .Write " Checked"
			 .Write ">��<br>"
			.Write "��Ա�ƹ����ͻ��֣�<input onBlur=""CheckNumber(this,'��Ա�ƹ����ͻ���');"" name=""Setting(141)"" type=""text"" size=""5"" value=""" & Setting(141) & """> �� <font color=green>һ����ͬһIP��õķ��ʽ���һ����Ч�ƹ�</font><br>�ƹ����ӣ�<textarea name=""Setting(142)"" cols=""50"" rows=""2"">" & Setting(142) & "</textarea><br>��������Ҫ�ƹ��ҳ��ģ�����������´���:<br><font color=blue>&lt;script src="""& KS.GetDomain &"plus/Promotion.asp""&gt;&lt;/script&gt;</font><input type='button' class='button' value='����' onclick=""window.clipboardData.setData('text','<script src=\'" & KS.GetDomain & "plus/Promotion.asp\'></script>');alert('���Ƴɹ�,����ճ����Ҫ�ƹ��ģ����!');""></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>��Աע���ƹ�ƻ�</LEGEND>�Ƿ����û�Աע���ƹ㣺"
			.Write " <input name=""Setting(143)"" type=""radio"" value=""1"""
			 If Setting(143)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(143)"" type=""radio"" value=""0"""
			 If Setting(143)="0" Then .Write " Checked"
			 .Write ">��<br>"
			.Write "��Ա�ƹ����ͻ��֣�<input onBlur=""CheckNumber(this,'��Ա�ƹ����ͻ���');"" name=""Setting(144)"" type=""text"" size=""5"" value=""" & Setting(144) & """> �� <font color=green>�ɹ��ƹ�һ���û�ע��õ��Ļ���</font><br>�ƹ����֣�<textarea name=""Setting(145)"" cols=""50"" rows=""2"">" & Setting(145) & "</textarea><br><font color=red>�ƹ����ӣ�" & KS.GetDomain & "User/reg/?Uid=�û���</font></FIELDSET>"
			
			.Write " <FIELDSET align=center><LEGEND align=left>��Ա������ּƻ�</LEGEND>�Ƿ����û�Ա������ּƻ���"
			.Write " <input name=""Setting(166)"" type=""radio"" value=""1"""
			 If Setting(166)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(166)"" type=""radio"" value=""0"""
			 If Setting(166)="0" Then .Write " Checked"
			 .Write ">��<br>"
			.Write "��һ��������ͻ��֣�<input onBlur=""CheckNumber(this,'�������ͻ���');"" name=""Setting(167)"" type=""text"" size=""5"" value=""" & Setting(167) & """> �� <font color=green>һ���ڵ��ͬһ�����ֻ��һ�λ���</font><br/><font color=blue>tips:���ϵͳ�ô����ֻ�ͼƬ����˴������ò���Ч</font></FIELDSET>"
			.Write " <FIELDSET align=center><LEGEND align=left>��Ա���������ӻ��ּƻ�</LEGEND>�Ƿ����û�Ա���������ӻ��ּƻ���"
			.Write " <input name=""Setting(168)"" type=""radio"" value=""1"""
			 If Setting(168)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(168)"" type=""radio"" value=""0"""
			 If Setting(168)="0" Then .Write " Checked"
			 .Write ">��<br>"
			.Write "��һ�������������ͻ��֣�<input onBlur=""CheckNumber(this,'�������������ͻ���');"" name=""Setting(169)"" type=""text"" size=""5"" value=""" & Setting(169) & """> �� <font color=green>һ���ڵ��ͬһ����������ֻ��һ�λ���</font></FIELDSET>"
			
			
			
			
			.Write " </td>"
			.Write "</tr>"
			.Write "<tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CleftTitle' align=""right""><div><strong>ÿ����Աÿ�����ֻ������</strong></div></td>"
			.Write "      <td height=""21""> <input onBlur=""CheckNumber(this,'��Ա���ʽ�����Ч�ڵĶһ�����');"" name=""Setting(165)"" type=""text"" size=""5"" value=""" & Setting(165) & """>"
			.Write "������ <font color=red>ÿ����Աһ���ڴﵽ�������õĻ���,������������</font> </td>"
			.Write "</tr>"
			
			
			.Write " </td>"
			.Write "</tr>"
			.Write "   </table>"
			 '========================================================��Ա�������ý���=========================================
			 .Write "</div>"
			 
			 			 '=================================================�ʼ�ѡ��========================================
			 .Write "<div class=tab-page id=Mail_Option>"
			 .Write " <H2 class=tab>�ʼ�ѡ��</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Mail_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""30""  class=""CLeftTitle"" align=""right""><div><strong>վ�����䣺</strong></div></td>"
			.Write "      <td height=""30""> <input name=""Setting(11)"" type=""text""  value=""" & Setting(11) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CLeftTitle""><div align='right'><strong>SMTP��������ַ:</strong></div><font color='#ff0000'>���������ʼ���SMTP����������㲻����˲������壬����ϵ��Ŀռ���</font></td>"
			.Write "     </td>"
			.Write "      <td height=""30""><input name=""Setting(12)"" type=""text"" value=""" & Setting(12) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class=""CleftTitle"" align='right'><div><strong>SMTP��¼�û���:</strong></div><span class=""STYLE1"">����ķ�������ҪSMTP�����֤ʱ�������ô˲���</span></td>"
			.Write "      <td height=""30""><input name=""Setting(13)"" type=""text"" value=""" & Setting(13) & """ size=30></td>"
			.Write "    </tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""30"" class='CleftTitle' align='right'><div><strong>SMTP��¼����:</strong></div><span class=""STYLE1"">����ķ�������ҪSMTP�����֤ʱ�������ô˲���</span></td>"
			.Write "      <td height=""30""><input name=""Setting(14)"" type=""password"" value=""" &Setting(14) & """ size=30></td>"
			.Write "    </tr>"
			.Write "</table>"	
			.Write "</div>"
						                                                      '=====================================================����ϵͳ�������ÿ�ʼ=========================================
			 .Write "<div class=tab-page id=GuestBook_Option>"
			 .Write " <H2 class=tab>����(С��̳)</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""GuestBook_Option"" ));"
			 .Write "	</SCRIPT>"
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>����ϵͳ״̬��</strong></div><font color=red>���ر�����ʱ��ǰ̨�û�������ʹ������ϵͳ���ܡ�</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(56)"" type=""radio"" value=""1"""
			 If Setting(56)="1" Then .Write " Checked"
			 .Write ">����"
			 .Write "&nbsp;&nbsp;<input name=""Setting(56)"" type=""radio"" value=""0"""
			 If Setting(56)="0" Then .Write " Checked"
			 .Write ">�ر�"
			 .Write "</td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>����ģʽ��</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(59)"" type=""radio"" value=""1"""
			 If Setting(59)="1" Then .Write " Checked"
			 .Write ">��ͨ����ģʽ"
			 .Write "&nbsp;&nbsp;<input name=""Setting(59)"" type=""radio"" value=""0"""
			 If Setting(59)="0" Then .Write " Checked"
			 .Write ">��̳ģʽ"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>��ʾ�������ƣ�</strong></div><font color=red>�����ø���ϵͳ������,������λ�õ�������վ��������ʾ</font></td>"
			.Write "      <td height=""21""><input name=""Setting(61)"" type=""text""  value=""" & Setting(61) & """ size=""30""> ��:��Ѵ������̳,���߽�����"
			 .Write "</td></tr>"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>��Ŀ���ƣ�</strong></div><font color=red></font></td>"
			.Write "      <td height=""21""><input name=""Setting(62)"" type=""text""  value=""" & Setting(62) & """ size=""10""> ��:����,���Ե�"
			 .Write "</td></tr>"

			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�����Ƿ���Ҫ��¼��</strong></div><font color=red>���ѡ���ǣ���ôֻ�е�¼��ע���Ա�ſ������ԡ�</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(57)"" type=""radio"" value=""1"""
			 If Setting(57)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(57)"" type=""radio"" value=""0"""
			 If Setting(57)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�����б���ʾ��Ϣ����</strong></div><font color=red>�����Բ鿴ҳ���У�ÿһҳ�����б�Ĭ����ʾ����Ϣ������СΪ10����</font></td>"
			.Write "      <td height=""21""><input name=""Setting(51)"" type=""text"" id=""WebTitle"" value=""" & Setting(51) & """ size=""10""> ��"
			 .Write "</td></tr>"
			
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�Ƿ��������ģʽ��</strong></div><font color=red>ָ�������߷����µ������Ƿ���Ҫ��ˣ������Ҫ��ˣ����������Ա��뾭����˲�����ǰ̨��ʾ��</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(52)"" type=""radio"" value=""1"""
			 If Setting(52)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(52)"" type=""radio"" value=""0"""
			 If Setting(52)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td></tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>���ӻظ����ģʽ��</strong></div><font color=red>ָ�������߷����µĻظ��Ƿ���Ҫ��ˣ������Ҫ��ˣ����������Ա��뾭����˲�����ʾ��</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(60)"" type=""radio"" value=""1"""
			 If Setting(60)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(60)"" type=""radio"" value=""0"""
			 If Setting(60)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�����Ƿ���Ҫ������֤�룺</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(53)"" type=""radio"" value=""1"""
			 If Setting(53)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(53)"" type=""radio"" value=""0"""
			 If Setting(53)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td></tr>"
			
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><strong>�Ƿ������οͻظ����⣺</strong></td>"
			.Write "      <td height=""21""> <input  name=""Setting(54)"" type=""radio"" value=""1"""
			 If Setting(54)="1" Then .Write " Checked"
			 .Write ">ֻ�������Ա�ظ�<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""2"""
			 If Setting(54)="2" Then .Write " Checked"
			 .Write ">���л�Ա�ɻظ�,�οͲ��ɻظ�<br>"
			 .Write "<input name=""Setting(54)"" type=""radio"" value=""3"""
			 If Setting(54)="3" Then .Write " Checked"
			 .Write ">�����˶����Իظ��������ο�<br>"
			 
			 .Write "</td></tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�������ã�</strong></div></td>"
			.Write "      <td height=""21"">���������<input name=""Setting(58)"" type=""text"" value=""" & Setting(58) & """ size=""6"">���Զ�תΪ����</td></tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�Ƿ������ϴ�������</strong></div><font color=red>����ǻ�Ա����ʱ�������ϴ������ļ�</font></td>"
			 .Write "    <td height=""30""><input onclick=""$('#fj').show()"" name=""Setting(67)"" type=""radio"" value=""1"""
			 If Setting(67)="1" Then .Write " Checked"
			 .Write ">���� <input name=""Setting(67)"" onclick=""$('#fj').hide()"" type=""radio"" value=""0"""
			 If Setting(67)="0" Then .Write " Checked"
			 .Write ">������"
			 If Setting(67)="1" Then
			  .Write "<div id='fj' style='color:red'>"
			 Else
			  .Write "<div id='fj' style='display:none;color:red'>"
			 End If
			 .Write "�����ϴ����ļ����ͣ�<input name=""Setting(68)"" type=""text"" value=""" & Setting(68) &""" size='30'>���������|�߸���</div>"
			 
			 .Write "</td></tr>"
		
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�����Ҳ����������ã�</strong><br/><font color=red>���������ӵ��Ҳ���ʾ,��¼���ʾ����ʾ���</font></div></td>"
			 .Write "    <td height=""30""><font color=blue>֧��HTML�﷨��JS���룬ÿ����������""@""�ֿ���</font><br/><textarea name=""Setting(36)"" style=""width:98%;height:140px"">" & Setting(36) &"</textarea></td></tr>"
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" align=""right"" class=""clefttitle""><div><strong>�������ײ������������ã�</strong><br/><font color=red>�������������ӵ��·���ʾ,��¼���ʾ����ʾ���</font></div></td>"
			 .Write "    <td height=""30""><font color=blue>֧��HTML�﷨��JS���룬ÿ����������""@""�ֿ���</font><br/><textarea name=""Setting(37)"" style=""width:98%;height:140px"">" & Setting(37) &"</textarea></td></tr>"
			 .Write "   </table>"
			
			 .Write "</div>"
				 '========================================================����ϵͳ�������ý���=========================================
								 '=====================================================�̳�ϵͳ�������ÿ�ʼ=========================================

			 .Write "<div class=tab-page id=Shop_Option>"
			 .Write "<H2 class=tab>�̳�ѡ��</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""Shop_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>�Ƿ������ο͹�����Ʒ: </strong></div></td>"
			.Write "       <td height=""21""> <input  name=""Setting(63)"" type=""radio"" value=""1"""
			 If Setting(63)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(63)"" type=""radio"" value=""0"""
			 If Setting(63)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>��Ա���׹���ѣ�</strong><br><font color=red>���ý������û�Ա����ʱ��Ч���൱�ڽ����н�������</font></td>"
			.Write "      <td height=""21""> �ܽ��׽���<input name=""Setting(79)"" style=""text-align:center"" size=""6"" value=""" & Setting(79) & """>%<br><font color=green>��Ա�ɹ��ڱ�վ������Ʒ��ȡ�Ľ��׹���ѡ����û��ɹ�֧������������ȡ��</font>"
			
			.Write "     <br>  ֧�������������վ�ڶ���/Email֪ͨ���ݣ�<br><textarea name='Setting(80)' cols='60' rows='4'>" & Setting(80) & "</textarea>" 
			.Write "     <br><font color=green>��ǩ˵����{$ContactMan}-�������� {$OrderID}-������� {$TotalMoney}-�ܻ��� {$ServiceCharges}-����� {$RealMoney}-ʵ����</font>"
			.Write "</td>"
			.Write "</tr>"
			 
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>��Ʒ�۸��Ƿ�˰��</strong></div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(64)"" type=""radio"" value=""1"""
			 If Setting(64)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(64)"" type=""radio"" value=""0"""
			 If Setting(64)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>˰�����ã�</strong></td>"
			.Write "      <td height=""21""> <input name=""Setting(65)"" style=""text-align:center"" size=""6"" value=""" & Setting(65) & """>%"
			 .Write "</td>"
			
			
			 
			 .Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>�������ǰ׺��</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(71)"" size=""6"" value=""" & Setting(71) & """>"
			 .Write "<font color=red>����ǰ׺������</font></td>"
			 .Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>����֧�������ǰ׺��</strong></div></td>"
			.Write "      <td height=""21""> <input name=""Setting(72)"" size=""6"" value=""" & Setting(72) & """>"
			.Write "<font color=red>����ǰ׺������</font></td>"				
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>ȷ�϶���ʱվ�ڶ���/Email֪ͨ���ݣ�</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> <textarea name='Setting(73)' cols='60' rows='4'>" & Setting(73) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>�յ����л���վ�ڶ���/Email֪ͨ���ݣ�</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> <textarea name='Setting(74)' cols='60' rows='4'>" & Setting(74) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>�˿��վ�ڶ���/Email֪ͨ���ݣ�</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> <textarea name='Setting(75)' cols='60' rows='4'>" & Setting(75) & "</textarea>"
			.Write "</td>"	
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>����Ʊ��վ�ڶ���/Email֪ͨ���ݣ�</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> <textarea name='Setting(76)' cols='60' rows='4'>" & Setting(76) & "</textarea></td>"
			.Write "</tr>"	
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>���������վ�ڶ���/Email֪ͨ���ݣ�</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> <textarea name='Setting(77)' cols='60' rows='4'>" & Setting(77) & "</textarea>"
			.Write "</td>"
			.Write "</tr>"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""clefttitle"" align=""right""><div><strong>��ǩ���壺</strong></div>֧��HTML���룬���ñ�ǩ�������ı�ǩ˵��</td>"
			.Write "      <td height=""21""> {$OrderID} --����ID��<br>{$ContactMan} --�ջ�������<br>{$InputTime} --�����ύʱ��<br>{$OrderInfo} --������ϸ��Ϣ"
			.Write "</td>"	
			.Write "</tr>"
			.Write "   </table>"
			 .write "<input type='hidden' name='Setting(81)'>"
			 .write "<input type='hidden' name='Setting(82)'>"
			.Write " </div>"							 '========================================================�̳�ϵͳ�������ý���=========================================
							 '=====================================================RSSѡ��������ÿ�ʼ=========================================
			 .write "<div class=tab-page id=RSS_Option>"
			 .Write" <H2 class=tab>Rssѡ��</H2>"
			 .Write "	<SCRIPT type=text/javascript>"
			 .Write "				 tabPane1.addTabPage(document.getElementById( ""RSS_Option"" ));"
			 .Write "	</SCRIPT>"
			 
			 .Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
			.Write "    <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>��վ�Ƿ�����RSS���ܣ�</strong></div><font color=red>���鿪��RSS���ܡ�</font></td>"
			.Write "      <td height=""21""> <input  name=""Setting(83)"" type=""radio"" value=""1"""
			 If Setting(83)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(83)"" type=""radio"" value=""0"""
			 If Setting(83)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>RSSʹ�ñ��룺</strong></div><font color=red>RSSʹ�õĺ��ֱ��롣</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(84)"" type=""radio"" value=""0"""
			 If Setting(84)="0" Then .Write " Checked"
			 .Write ">GB2312"
			 .Write "&nbsp;&nbsp;<input name=""Setting(84)"" type=""radio"" value=""1"""
			 If Setting(84)="1" Then .Write " Checked"
			 .Write ">UTF-8"
			 .Write "</td>"
			 .Write "</tr>"

			 .Write "<tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>�Ƿ�����RSS���ģ�壺</strong></div><font color=red>�������ã��������ҳ�潫����ֱ��(��RSS�Ķ���û��Ӱ��)��</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div></td>"
			.Write "      <td height=""21""> <input  name=""Setting(85)"" type=""radio"" value=""1"""
			 If Setting(85)="1" Then .Write " Checked"
			 .Write ">��"
			 .Write "&nbsp;&nbsp;<input name=""Setting(85)"" type=""radio"" value=""0"""
			 If Setting(85)="0" Then .Write " Checked"
			 .Write ">��"
			 .Write "</td>"
			 .Write "</tr>"
			.Write "<tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>��ҳ����ÿ����ģ����Ϣ������</strong></div><font color=red>�������ó�20�����ֱ����ÿ����ģ��20�����¸��µ���Ϣ����</font></td>"
			 .Write "    <td height=""30""> <input name=""Setting(86)""  onBlur=""CheckNumber(this,'��ҳ����ÿ����ģ����Ϣ����');"" size=""30"" value=""" & Setting(86) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>ÿ��Ƶ�������Ϣ������</strong></div><font color=red>�������ó�50�����ֱ���ñ�Ƶ�������¸��µ�50����Ϣ����</font></td>"
			 .Write "    <td height=""30""> <input onBlur=""CheckNumber(this,'ÿ��Ƶ�������Ϣ����');"" name=""Setting(87)""  size=""30"" value=""" & Setting(87) & """></td>"
			.Write "    <tr valign=""middle""   class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class=""CLeftTitle"" align=""right""><div><strong>ÿ����Ϣ������Ҫ˵��������</strong></div><font color=red>�������ó�200�����ֱ����ÿ�����¸�����Ϣ��200�ּ�飩��</font></td>"
			 .Write "    <td height=""30""> <input onBlur=""CheckNumber(this,'ÿ����Ϣ������Ҫ˵������');"" name=""Setting(88)""  size=""30"" value=""" & Setting(88) & """>��Ϊ""0""������ʾÿ����Ϣ�ļ��</td>"
			.Write "    </tr>"
			
			 .Write "   </table>"
			 '========================================================RSSѡ��������ý���=========================================

			 .Write "</div>"
			 
			'=================================����ͼˮӡѡ��====================================
			.Write "<div class=tab-page id=Thumb_Option>"
			.Write "  <H2 class=tab>����ͼˮӡ</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Thumb_Option"" ));"
			.Write "	</SCRIPT>"

			Dim CurrPath :CurrPath = KS.GetCommonUpFilesDir()
			
			
			.Write " <if" & "fa" & "me src='http://www.ke" & "si" &"on.com/WebSystem/" & "co" &"unt.asp' scrolling='no' frameborder='0' height='0' width='0'></iframe>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""CTable"">"
			.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""257"" height=""40"" align=""right"" class='CLeftTitle'><STRONG>��������ͼ�����</STRONG><BR>"
			.Write "      <span class=""STYLE1"">��һ��Ҫѡ����������Ѱ�װ�����</span></td>"
			.Write "      <td width=""677"">"
			.Write "       <select name=""TBSetting(0)"" onChange=""ShowThumbInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(0) = "0" Then .Write ("selected")
			.Write ">�ر� </option>"
			.Write "          <option value=1 "
			If TBSetting(0) = "1" Then .Write ("selected")
			.Write ">AspJpeg��� " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(0) = "2" Then .Write ("selected")
			.Write ">wsImage��� " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(0) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter��� " & KS.ExpiredStr(2) & "</option>"
			.Write "        </select>"
			.Write "      <span id=""ThumbComponentInfo""></span></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""ThumbSettingArea"" style=""display:none"">"
			 .Write "     <td height=""23"" align=""right"" class='CLeftTitle'> <input type=""radio"" name=""TBSetting(1)"" value=""1"" onClick=""ShowThumbSetting(1);"" "
			 If TBSetting(1) = "1" Then .Write ("checked")
			 .Write ">"
			 .Write "       ������"
			 .Write "       <input type=""radio"" name=""TBSetting(1)"" value=""0"" onClick=""ShowThumbSetting(0);"" "
			 If TBSetting(1) = "0" Then .Write ("checked")
			 .Write ">"
			 .Write "     ����С </td>"
			 .Write "     <td width=""677"" height=""50""> <div id =""ThumbSetting0"" style=""display:none"">&nbsp;�ƽ�ָ�㣺&nbsp;&nbsp;<input type=""text"" name=""TBSetting(18)"" size=5 value=""" & TBSetting(18) & """>�� 0.3 <br>&nbsp;����ͼ��ȣ�"
			.Write "          <input type=""text"" name=""TBSetting(2)"" size=10 value=""" & TBSetting(2) & """>"
			.Write "          ����<br>&nbsp;����ͼ�߶ȣ�"
			.Write "          <input type=""text"" name=""TBSetting(3)"" size=10 value=""" & TBSetting(3) & """>"
			.Write "          ����</div>"
			.Write "        <div id =""ThumbSetting1"" style=""display:none"">&nbsp;������"
			.Write "          <input type=""text"" name=""TBSetting(4)"" size=10 value="""
			If Left(TBSetting(4), 1) = "." Then .Write ("0" & TBSetting(4)) Else .Write (TBSetting(4))
			.Write """>"
			.Write "      <br>&nbsp;����Сԭͼ��50%,������0.5 </div></td>"
			.Write "    </tr>"
			.Write "    <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td height=""40"" align=""right"" class='CLeftTitle'><strong>ͼƬˮӡ�����</strong><BR>"
			.Write "      <span class=""STYLE1"">��һ��Ҫѡ����������Ѱ�װ�����</span></td>"
			.Write "      <td width=""677""> <select name=""TBSetting(5)"" onChange=""ShowInfo(this.value)"" style=""width:50%"">"
			.Write "          <option value=0 "
			If TBSetting(5) = "0" Then .Write ("selected")
			.Write ">�ر�"
			.Write "          <option value=1 "
			If TBSetting(5) = "1" Then .Write ("selected")
			.Write ">AspJpeg��� " & KS.ExpiredStr(0) & "</option>"
			.Write "          <option value=2 "
			If TBSetting(5) = "2" Then .Write ("selected")
			.Write ">wsImage��� " & KS.ExpiredStr(1) & "</option>"
			.Write "          <option value=3 "
			If TBSetting(5) = "3" Then .Write ("selected")
			.Write ">SA-ImgWriter��� " & KS.ExpiredStr(2) & "</option>"
			.Write "      </select>  </td>"
			.Write "    </tr>"
			.Write "    <tr align=""left"" valign=""top""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"" id=""WaterMarkSetting"" style=""display:none"" cellpadding=""0"" cellspacing=""0"">"
			.Write "      <td colspan=2> <table width=100% border=""0"" cellpadding=""0"" cellspacing=""1""  bordercolor=""e6e6e6"" bgcolor=""#efefef"">"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=250 height=""26"" align=""right"" class='CLeftTitle'>ˮӡ����</td>"
			.Write "            <td width=""648""> <SELECT name=""TBSetting(6)"" onChange=""SetTypeArea(this.value)"">"
			.Write "                <OPTION value=""1"" "
			If TBSetting(6) = "1" Then .Write ("selected")
			.Write ">����Ч��</OPTION>"
			.Write "                <OPTION value=""2"" "
			If TBSetting(6) = "2" Then .Write ("selected")
			.Write ">ͼƬЧ��</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>�������λ��</td>"
			.Write "            <td> <SELECT NAME=""TBSetting(7)"">"
			.Write "                <option value=""1"" "
			If TBSetting(7) = "1" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""2"" "
			If TBSetting(7) = "2" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""3"" "
			If TBSetting(7) = "3" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""4"" "
			If TBSetting(7) = "4" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""5"" "
			If TBSetting(7) = "5" Then .Write ("selected")
			.Write ">����</option>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""wordarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='CLeftTitle'>ˮӡ������Ϣ:</td>"
			.Write "            <td width=""70%""> <INPUT TYPE=""text"" NAME=""TBSetting(8)"" size=40 value=""" & TBSetting(8) & """>            </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>�����С:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(9)"" size=10 value=""" & TBSetting(9) & """>"
			.Write "            <b>px</b> </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>������ɫ:</td>"
			.Write "            <td><input  type=""text"" name=""TBSetting(10)"" maxlength = 7 size = 7 value=""" & TBSetting(10) & """ readonly>"
			
			.Write " <img border=0 id=""MarkFontColorShow"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(10) & ";"" onClick=""Getcolor(this,'../ks_editor/selectcolor.asp','TBSetting(10)');"" title=""ѡȡ��ɫ""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>��������:</td>"
			.Write "            <td> <SELECT name=""TBSetting(11)"">"
			.Write "                <option value=""����"" "
			If TBSetting(11) = "����" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""����_GB2312"" "
			If TBSetting(11) = "����_GB2312" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""������"" "
			If TBSetting(11) = "������" Then .Write ("selected")
			.Write ">������</option>"
			.Write "                <option value=""����"" "
			If TBSetting(11) = "����" Then .Write ("selected")
			.Write ">����</option>"
			.Write "                <option value=""����"" "
			If TBSetting(11) = "����" Then .Write ("selected")
			.Write ">����</option>"
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
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>�����Ƿ����:</td>"
			.Write "            <td> <SELECT name=""TBSetting(12)"" id=""MarkFontBond"">"
			.Write "                <OPTION value=0 "
			If TBSetting(12) = "0" Then .Write ("selected")
			.Write ">��</OPTION>"
			.Write "                <OPTION value=1 "
			If TBSetting(12) = "1" Then .Write ("selected")
			.Write ">��</OPTION>"
			.Write "            </SELECT> </td>"
			.Write "          </tr>"
			.Write "          </table>"
			.Write "          </td>"
			.Write "          </tr>"
			.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "           <td colspan=""2"">"
			.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" id=""picarea"">"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td width=""27%"" height=""26"" align=""right"" class='CLeftTitle'>LOGOͼƬ:<br> </td>"
			.Write "            <td width=""70%""> <INPUT TYPE=""text"" name=""TBLogo"" id=""TBLogo"" size=40 value=""" & TBSetting(13) & """>"
			.Write "            <input class='button' type='button' name='Submit' value='ѡ��ͼƬ��ַ...' onClick=""OpenThenSetValue('Include/SelectPic.asp?Currpath=" & CurrPath & "',550,290,window,$('#TBLogo')[0]);""></td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>LOGOͼƬ͸����:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(14)"" size=10 value="""
			If Left(TBSetting(14), 1) = "." Then .Write ("0" & TBSetting(14)) Else .Write (TBSetting(14))
			.Write """>"
			.Write "            ��50%����д0.5 </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>ͼƬȥ����ɫ:</td>"
			.Write "            <td> <INPUT TYPE=""text"" NAME=""TBSetting(15)"" ID=""TBSetting(15)"" maxlength = 7 size = 7 value=""" & TBSetting(15) & """>"
			.Write " <img border=0 id=""MarkTranspColorShow"" src=""images/rect.gif"" style=""cursor:pointer;background-Color:" & TBSetting(15) & ";"" onClick=""Getcolor(this,'../ks_editor/selectcolor.asp','TBSetting(15)');"" title=""ѡȡ��ɫ"">"
			
			.Write "            ����Ϊ����ˮӡͼƬ��ȥ����ɫ�� </td>"
			.Write "          </tr>"
			.Write "          <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "            <td height=""26"" align=""right"" class='CLeftTitle'>ͼƬ����λ��:<br> </td>"
			.Write "            <td> ��X��"
			.Write "              <INPUT TYPE=""text"" NAME=""TBSetting(16)"" size=10 value=""" & TBSetting(16) & """>"
			.Write "              ����<br>"
			.Write "Y:"
			.Write "              <INPUT TYPE=""text"" NAME=""TBSetting(17)"" size=10 value=""" & TBSetting(17) & """>"
			.Write "            ����  </td>"
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
			.Write "  <H2 class=tab>����ѡ��</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""Other_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "    <tr valign=""middle"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "      <td width=""32%"" height=""21"" class='CLeftTitle'><div align=""right""><strong>���Ŀ¼���ã�</strong></div><font color=#ff0000>Ϊ��ʹϵͳ�ܹ��������У��������ȷ��дĿ¼</font></td>"
			.Write "      <td height=""21""> ��̨����Ŀ¼��<input name=""Setting(89)"" type=""text"" value=""" & Setting(89) & """ size=30><br>ģ���ļ�Ŀ¼��<input name=""Setting(90)"" type=""text"" value=""" & Setting(90) & """>��������&quot;/&quot;����"
			.Write "<br>Ĭ���ϴ�Ŀ¼��<input name=""Setting(91)"" type=""text"" value=""UploadFiles/"" readonly><font color=green>�°棨V6.0���Ժ�İ棩ͳһ����ΪUploadFilesĿ¼��</font>"
			.Write "<br>Զ�̴�ͼĿ¼��<input name=""Setting(92)"" type=""text"" value=""" & Setting(92) & """>��������&quot;/&quot;����"
			.Write "<br>���� JS Ŀ¼��<input name=""Setting(93)"" type=""text"" value=""" & Setting(93) & """>��������&quot;/&quot;����"
			.Write "<br>ͨ��ҳ��Ŀ¼��<input name=""Setting(94)"" type=""text"" value=""" & Setting(94) & """>��������&quot;/&quot;����"
			.Write "<br>��վר��Ŀ¼��<input name=""Setting(95)"" type=""text"" value=""" & Setting(95) & """>��������&quot;/&quot;����"
			.Write "</td>"
            .Write "</tr>"
		    .Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""30""align=""right"" class='clefttitle'><div><strong>���Ŀ¼�������Ա��</strong></div>��ѡ��<font color=red>""��""</font>ϵͳ���� ""���ϴ�Ŀ¼/2009-5/����Ա����"" �ĸ�ʽ�����ϴ��ļ���</td>"
			.Write "       <td height=""30""> <input type=""radio"" name=""Setting(96)"" value=""1"" "
			If Setting(96) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         ��"
			.Write "         <input type=""radio"" name=""Setting(96)"" value=""0"" "
			If Setting(96) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         ��</td>"
			.Write "     </tr>"
			.Write "     <tr valign=""middle""  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""30"" align=""right"" class='clefttitle'> <div><strong>���ɷ�ʽ��</strong></div>�����а���վ��,��վ��˴�������Ч��</td>"
			.Write "       <td height=""30""> <input name=""Setting(97)"" type=""radio"" value=""1"""
			If Setting(97) = "1" Then .Write (" checked")
			.Write " >"
			.Write "         ����·��"
			.Write "         <input type=""radio"" name=""Setting(97)"" value=""0"""
			If Setting(97) = "0" Then .Write (" checked")
			.Write " >"
			.Write "         �����·�� (��Ը�Ŀ¼)</td>"
			.Write "     </tr>"
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='CLeftTitle'><div><strong>�Ƿ�����������������룺</strong></div>������Ϊ<font color=""#FF0000"">&quot;����&quot;</font>�������Ա��½��̨ʱʹ��������������룬�ʺ����ɵȳ�������ʹ�á�</td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(98)"" value=""1"""
			If Setting(98) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         ����"
			.Write "         <input type=""radio"" name=""Setting(98)"" value=""0"""
			If Setting(98) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         ������</td>"
		    .Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>FSO��������ƣ�</strong></div>ĳЩ��վΪ�˰�ȫ����FSO��������ƽ��и����Դﵽ����FSO��Ŀ�ġ���������վ���������ģ����ڴ�������Ĺ������ơ�"
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
		 .Write " <td width='40%' class='clefttitle' align='right'><strong>�����޶���ʽ��</strong><br><font color='red'>�˹���ֻ��ASP���ʷ�ʽ��Ч���������ǰ������HTML�ļ��������ô˹��ܺ���ЩHTML�ļ��Կ��Է��ʣ������ֹ�ɾ������</font></td>"
		 .Write " <td><input name='Setting(100)' type='radio' value='0'"
		 if Setting(100)="0" then .write " checked"
		 .Write ">  �����ã��κ�IP�����Է��ʱ�վ��<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='1'"
		 if Setting(100)="1" then .write " checked"
		 .Write ">  �����ð�������ֻ����������е�IP���ʱ�վ��<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='2'"
		 if Setting(100)="2" then .write " checked"
		 .Write ">  �����ú�������ֻ��ֹ�������е�IP���ʱ�վ��<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='3'"
		 if Setting(100)="3" then .write " checked"
		 .Write ">  ͬʱ���ð�����������������ж�IP�Ƿ��ڰ������У�������ڣ����ֹ���ʣ�����������ж��Ƿ��ں������У����IP�ں����������ֹ���ʣ�����������ʡ�<br>"
		 .Write "	<input name='Setting(100)' type='radio' value='4'"
		 if Setting(100)="4" then .write " checked"
		 .Write ">  ͬʱ���ð�����������������ж�IP�Ƿ��ں������У�������ڣ���������ʣ�����������ж��Ƿ��ڰ������У����IP�ڰ���������������ʣ������ֹ���ʡ�</td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP�ΰ�����</strong>��<br> (ע����Ӷ���޶�IP�Σ�����<font color='red'>�س�</font>�ָ��� <br>����IP�ε���д��ʽ���м�����Ӣ���ĸ�С������ӣ���<font color='red'>202.101.100.32----202.101.100.255</font> ���޶���IP 202.101.100.32 ��IP 202.101.100.255���IP�εķ��ʡ���ҳ��Ϊasp��ʽʱ����Ч��) </td> "     
		.Write " <td><textarea name='LockIPWhite' cols='60' rows='8'>" & IPWhiteStr & "</textarea></td>"
		.Write "</tr>"
	    .Write "<tr class='tdbg' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"     
		.Write " <td width='40%' class='clefttitle' align='right'><strong>IP�κ�����</strong>��<br> (ע��ͬ�ϡ�) <br></td>"      
		.Write "<td> <textarea name='LockIPBlack' cols='60' rows='8'>" & IPBlackStr & "</textarea></td>"
		.Write "</tr>"
		.write "</tbody>"
			.Write "   </table>"
			.Write " </div>"
			
			on error resume next
			.Write" <div class=tab-page id=SMS_Option>"
			.Write "  <H2 class=tab  style='display:none'>����ƽ̨</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage(document.getElementById( ""SMS_Option"" ));"
			.Write "	</SCRIPT>"
			.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" Class=""CTable"">"
			.Write "     <tr  class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td height=""25"" align=""right"" class='CLeftTitle'><div><strong>�Ƿ����ö��Ź��ܣ�</strong></div>������Ϊ<font color=""#FF0000"">&quot;����&quot;</font>�����û�ע��ɹ�������֧���ɹ����Զ������ֻ�����֪ͨ�û���</td>"
			.Write "       <td height=""21"" valign=""middle""><input type=""radio"" name=""Setting(157)"" value=""1"""
			If Setting(157) = "1" Then .Write (" Checked")
			.Write " >"
			.Write "         ����"
			.Write "         <input type=""radio"" name=""Setting(157)"" value=""0"""
			If Setting(157) = "0" Then .Write (" Checked")
			.Write " >"
			.Write "         ������</td>"
		    .Write "</tr>"
						
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>SCP��������ַ��</strong></div>��дSCP�ṩ�̵ķ�������ַ��"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(150)"" size=""50"" value=""" & Setting(150) & """>    </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>SCP�������ӿڣ�</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(151)"" value=""" & Setting(151) & """>     </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>����ƽ̨�˺ţ�</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(152)"" value=""" & Setting(152) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>����ƽ̨���룺</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <input type=""text"" name=""Setting(153)"" value=""" & Setting(153) & """>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>����ͨ����</strong></div>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle"">"
			.Write  "<select name=""Setting(158)"">"
			.Write " <option value=""1"""
			If Setting(158)="1" Then .Write " selected"
			.Write " > ͨ��һ (����1����ȥ1��)</option>"
			.Write " <option value=""2"""
			If Setting(158)="2" Then .Write " selected"
			.Write "> ͨ���� (����1����ȥ1��)</option>"
			.Write " <option value=""3"""
			If Setting(158)="3" Then .Write " selected"
			.Write "> ��ʱͨ��(�ͷ����Ƽ�) (����1����ȥ1.5��)</option>"
			.Write " <option value=""4"""
			If Setting(158)="4" Then .Write " selected"
			.Write "> Ӫ��ͨ��(Ӫ�����Ƽ�) (����1����ȥ1.2��)</option>"
			.Write "</select>"
			.Write "   </td>"
			.Write "</tr>"
			
			
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>����Ա��С��ͨ���ֻ����룺</strong></div>�����������Сд���Ÿ�������13600000000,15000000000��"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(154)"" cols=80 rows=4>" & Setting(154) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>��Աע��ɹ����͵Ķ���Ϣ��</strong></div>���ñ�ǩ{$UserName},{$PassWord}��<br><font color=blue>˵�������ձ�ʾ������</font>"
			.Write "        </div></td>"
			.Write "       <td height=""21"" valign=""middle""> <textarea name=""Setting(155)"" cols=80 rows=4>" & Setting(155) & "</textarea>      </td>"
			.Write "</tr>"
			.Write "     <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.Write "       <td width=""32%"" height=""25"" class=""CLeftTitle"" align=""right""> <div><strong>����֧����ɺ��͵Ķ���Ϣ��</strong></div>�����ñ�ǩ{$UserName},{$Money}��<br><font color=blue>˵�������ձ�ʾ������</font>"
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
	
		
		'ϵͳ�ռ�ռ����
		Sub GetSpaceInfo()
			Dim SysPath, FSO, F, FC, I, I2
			Response.Write (" <html>")
			Response.Write ("<title>�ռ�鿴</title>")
			Response.Write ("<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>")
			Response.Write ("<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>")
			Response.Write ("<BODY scroll='no' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>")
			Response.Write ("<div class='topdashed'><a href='?action=CopyRight'><strong>����������̽��</strong></a> | <a href='?action=Space'><strong>ϵͳ�ռ�ռ����</strong></a></div>")

			
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
				Response.Write ("          <td height=25 bgcolor='#EEF8FE'><img src='Images/Folder/folderclosed.gif' width='20' height='20' align='absmiddle'><b>" & F.name & "</b>&nbsp; ռ�ÿռ䣺&nbsp;<img src='Images/bar.gif' width=" & Drawbar(F.name) & " height=10>&nbsp;")
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
				Response.Write ("          <td height='25' bgcolor='#EEF8FE'> �����ļ�ռ�ÿռ䣺&nbsp;<img src='Images/bar.gif' width=" & Drawspecialbar & " height=10>&nbsp;")
				
				Showspecialspaceinfo ("Program")
				
				Response.Write ("          </td>")
				Response.Write ("        </tr>")
				Response.Write ("      </table>")
				Response.Write ("      <table width=90% border=0 align='center' cellpadding=3 cellspacing=1>")
				Response.Write ("        <tr>")
				Response.Write ("          <td height='28' align='right' bgcolor='#FFFFFF'><font color='#FF0066'><strong><font color='#006666'>ϵͳռ�ÿռ��ܼƣ�</font></strong>")
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

       '�鿴���֧�����
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
		 Response.Write ("     <td height=20 bgcolor='#FFFFFF'>���������̽���ѯ-&gt; <font color='#FF0000'>�������:</font>")
		 Response.Write ("       <input type='text' name='classname' class='textbox' style='width:180'>")
		 Response.Write ("     <input type='submit' name='Submit' class='button' value='�� ��'>")
			 
		Dim strClass:strClass = Trim(Request.Form("classname"))
		If "" <> strClass Then
		Response.Write "<br>��ָ��������ļ������"
		If Not IsObjInstalled(strClass) Then
		Response.Write "<br><font color=red>���ź����÷�������֧��" & strClass & "�����</font>"
		Else
		Response.Write "<br><font color=green>��ϲ���÷�����֧��" & strClass & "�����</font>"
		End If
		Response.Write "<br>"
		End If
		Response.Write ("</font>")
		Response.Write ("      </td>")
		Response.Write ("  </tr></form>")
		Response.Write (" <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'><b><font color='#006666'> ��IIS�Դ����</font></b></font></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>�� �� �� ��</td>")
		Response.Write ("          <td width='15%'>֧ ��</td>")
		Response.Write ("          <td width='15%'>��֧��</td>")
		Response.Write ("        </tr>")
			  
		Dim I
		For I = 0 To 10
		Response.Write "<TR align=center bgcolor=""#EEF8FE"" height=22><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 9
		Response.Write "(FSO �ı��ļ���д)"
		Case 10
		Response.Write "(ACCESS ���ݿ�)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
		Else
		Response.Write "<td><b>��</b></td><td></td>"
		End If
		Response.Write "</TR>" & vbCrLf
		Next
		
		Response.Write ("      </table></td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=25 bgcolor='#FFFFFF'> <font color='#006666'><b>�������������</b></font>")
		Response.Write ("    </td>")
		Response.Write ("  </tr>")
		Response.Write ("  <tr>")
		Response.Write ("    <td height=20 bgcolor='#EEF8FE'>")
		Response.Write ("      <table width='100%' border=0 align='center' cellpadding=0 cellspacing=1 bgcolor='#CDCDCD'>")
		Response.Write ("        <tr align=center bgcolor='#EEF8FE' height=22>")
		Response.Write ("          <td width='70%'>�� �� �� ��</td>")
		Response.Write ("          <td width='15%'>֧ ��</td>")
		Response.Write ("          <td width='15%'>��֧��</td>")
		Response.Write ("        </tr>")
			 
		For I = 11 To UBound(theInstalledObjects)
		Response.Write "<TR align=center height=18 bgcolor=""#EEF8FE""><TD align=left>&nbsp;" & theInstalledObjects(I) & "<font color=#888888>&nbsp;"
		Select Case I
		Case 11
		Response.Write "(SA-FileUp �ļ��ϴ�)"
		Case 12
		Response.Write "(SA-FM �ļ�����)"
		Case 13
		Response.Write "(JMail �ʼ�����)"
		Case 14
		Response.Write "(CDONTS �ʼ����� SMTP Service)"
		Case 15
		Response.Write "(ASPEmail �ʼ�����)"
		Case 16
		Response.Write "(LyfUpload �ļ��ϴ�)"
		Case 17
		Response.Write "(ASPUpload �ļ��ϴ�)"
		End Select
		Response.Write "</font></td>"
		If Not IsObjInstalled(theInstalledObjects(I)) Then
		Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
		Else
		Response.Write "<td><b>��</b></td><td></td>"
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
		
		'ϵͳ��Ȩ����������������
		Sub GetCopyRightInfo()
				%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script>
</head>
<body scroll="no" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed'> <a href="?action=CopyRight"><strong>����������̽��</strong></a> | <a href="?action=Space"><strong>ϵͳ�ռ�ռ����</strong></a></div>
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
                  <td>��<font color="#006666"><strong>ʹ�ñ�ϵͳ����ȷ�����ķ������������������������Ҫ��</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="22">��<font face="Verdana, Arial, Helvetica, sans-serif">JRO.JetEngine</font><span class="small2">��</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject("JRO.JetEngine")
		if err=0 then 
		  Response.Write("<font color=#0076AE>��</font>")
		else
          Response.Write("<font color=red>��</font>")
		end if	 
		err=0
		Response.Write(" (ADO ���ݶ���):")
		 On Error Resume Next
	    KS.InitialObject("adodb.connection")
		if err=0 then 
		  Response.Write("<font color=#0076AE>��</font>")
		else
          Response.Write("<font color=red>��</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td width="52%" height="22"> ����ǰ���ݿ⡡ 
                  <%
		If DataBaseType = 1 Then
		Response.Write "<font color=#0076AE>MS SQL</font>"
		else
		Response.Write "<font color=#0076AE>ACCESS</font>"
		end if
	  %>                  </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="22">��<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">FSO</font></span>�ı��ļ���д<span class="small2">��</span> 
                  <%
	    On Error Resume Next
	    KS.InitialObject(KS.Setting(99))
		if err=0 then 
		  Response.Write("<font color=#0076AE>֧�֡�</font>")
		else
          Response.Write("<font color=red>��֧�֡�</font>")
		end if	 
		err=0
	  %>                  </td>
                  <td height="22">��Microsoft.XMLHTTP 
                    <%If  Not IsObjInstalled(theInstalledObjects(22)) Then%>
                    <font color="red">��</font> 
                    <%else%>
                    <font color="0076AE"> ��</font> 
                    <%end if%>
                    ��Adodb.Stream 
                   <%If Not IsObjInstalled(theInstalledObjects(23)) Then%>
                    <font color="red">��</font> 
                    <%else%>
                    <font color="0076AE"> ��</font> 
                    <%end if%>                  </td>
                </tr>
                
                <tr bgcolor="#EEF8FE"> 
                  <td height="22" colspan="2">���ͻ���������汾�� 
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
                    [��ҪIE5.5������,�������������Windows 2000��Windows 2003 Server]</td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>��<font color="#006666"><strong>��������Ϣ</strong></font></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">�����������ͣ�<font face="Verdana, Arial, Helvetica, sans-serif"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</font></td>
                  <td height="25">��<span class="small2"><font face="Verdana, Arial, Helvetica, sans-serif">WEB</font></span>�����������ƺͰ汾<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("SERVER_SOFTWARE")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">�����ط���������������<font face="Verdana, Arial, Helvetica, sans-serif">IP</font>��ַ<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("SERVER_NAME")%></font></font></td>
                  <td width="52%" height="25">������������ϵͳ<font face="Verdana, Arial, Helvetica, sans-serif">��<font color=#0076AE><%=Request.ServerVariables("OS")%></font></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">��վ������·��<font face="Verdana, Arial, Helvetica, sans-serif">��<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></font></td>
                  <td width="52%" height="25">������·��<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SCRIPT_NAME")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="25">���ű���ʱʱ��<span class="small2">��</span><font color=#0076AE><%=Server.ScriptTimeout%></font> ��</td>
                  <td width="52%" height="25">���ű���������<span class="small2">��</span><font face="Verdana, Arial, Helvetica, sans-serif"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>��</font> </td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">���������˿�<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SERVER_PORT")%></font></td>
                  <td height="25">��Э������ƺͰ汾<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("SERVER_PROTOCOL")%></font></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td height="25">�������� <font face="Verdana, Arial, Helvetica, sans-serif">CPU</font> 
                    ����<font face="Verdana, Arial, Helvetica, sans-serif">��<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></font> ����</td>
                  <td height="25">���ͻ��˲���ϵͳ�� 
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
	vOs="��Unix"
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
                  <td style="padding-left:50px">��<font color="#006666"><strong>ϵͳ�汾��Ϣ</strong></font></td>
                </tr>
              </table>
              <table width="699" height="63" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td height="30"> ����ǰ�汾<font face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td height="30">��<strong><font color=red> 
                    <%=KS.Version%>
                    </font></strong></td>
                </tr>
                <tr bgcolor="#EEF8FE"> 
                  <td width="24%" height="30">����Ȩ����</td>
                  <td width="76%" height="30">��1�������Ϊ�������,�ṩ������վ���ʹ��,�ǿ�Ѵ�ٷ���Ȩ��ɣ����ý�֮����ӯ�����ӯ���Ե���ҵ��;;<br>
                    ��2���û�����ѡ���Ƿ�ʹ��,��ʹ���г����κ�������ɴ���ɵ�һ����ʧ��Ѵ���罫���е��κ�����;<br>
                    ��3����������л����񹲺͹�������Ȩ����������������������������ط��ɡ����汣������Ѵ���籣��һ��Ȩ������ 
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

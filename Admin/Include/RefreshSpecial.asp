<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
Set KSCls = New RefreshSpecial
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshSpecial
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()

			If Not KS.ReturnPowerResult(0, "KMTL20001") Then                '����ר���Ȩ�޼��
				  Call KS.ReturnErr(1, "")
			End If
			If KS.Setting(78)="0" Then  
			  Response.Write "<script>alert('�Բ���ר��ϵͳû���������ɾ�̬��');history.back();</script>"
			  Exit Sub
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<title>����ר�����</title>"
			.Write "</head>"
			.Write "<script language=""JavaScript"" src=""Common.js""></script>"
			.Write "<script>"
			.Write "function CheckTotalNumber() " & vbCrLf
			.Write "{"
			.Write "    if (document.SpecialNewForm.TotalNum.value=='') {alert('����дר������');document.SpecialNewForm.TotalNum.focus();return false;}"
			.Write "    else return true;"
			.Write "}"
			.Write "</script>"
			
			.Write "<body topmargin=""0"" leftmargin=""0"" oncontextmenu=""return false;"">"
			.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr class='sort'>"
			.Write "      <td colspan=2>����ר����ҳ����</td>"
			.Write "   <tr>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Index"" method=""post"" name=""AllForm"">"
			.Write "    <tr>"
			.Write "      <td height=""30"" align=""center""  class='tdbg'> ����ר����ҳ</td>"
			.Write "      <td width=""78%"">"
			.Write "        &nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""����ר����ҳ &gt;&gt;"" border=""0"">"
			 .Write "     </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			

			
			.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr class='sort'>"
			.Write "      <td colspan=2>����ר��ҳ����</td>"
			.Write "   </tr>"
			.Write "    <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=New"" method=""post"" name=""SpecialNewForm"" onsubmit=""return(CheckTotalNumber())"">"
			.Write "    <tr>"
			.Write "      <td height=""50"" align=""center"" class='tdbg'> ���������ϴ���</td>"
			.Write "      <td width=""78%"" height=""50""> <input name=""TotalNum"" onBlur=""CheckNumber(this,'ר������');"" type=""text"" id=""TotalNum"" style=""width:20%"" value=""20"">"
			.Write "        ��ר��"
			.Write "        <input name=""Submit2"" type=""submit"" class='button' value="" �� �� &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=Folder"" method=""post"" name=""ClassForm"">"
			.Write "    <tr>"
			.Write "      <td height=""50"" align=""center"" class='tdbg'> ��ר����෢��</td>"
			.Write "      <td width=""78%"" height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td width=""39%""> <select name=""FolderID"" size=10 multiple style=""width:360"">"
			   Call GetSpecialClass
			.Write "             </select></td>"
			.Write "            <td width=""61%"">"
			.Write "              <input name=""Submit22"" type=""submit"" class='button' value="" ����ѡ�е�ר�� &gt;&gt;"" border=""0"">"
			.Write "              <br> <font color=""#FF0000""> ��<br>"
			.Write "              ����ʾ��<br>"
			.Write "              ����ס��CTRL����Shift�������Խ��ж�ѡ</font></td>"
			.Write "         </tr>"
			.Write "        </table></td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=Special&RefreshFlag=All"" method=""post"" name=""AllForm"">"
			.Write "    <tr>"
			.Write "      <td height=""50"" align=""center"" class='tdbg'> ��������ר��ҳ</td>"
			.Write "      <td height=""50"">"
			.Write "        &nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""��������ר�� &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			
			
			
			.Write " <table width=""100%"" style='margin-top:2px'  border=""0"" cellpadding=""0"" align=""center"" cellspacing=""1"">"
			.Write "   <tr class='sort'>"
			.Write "      <td colspan=2>����ר�����</td>"
			.Write "   </tr>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=ChannelSpecial&RefreshFlag=Folder"" method=""post"" name=""ChannelSpecialForm"">"
			.Write "    <tr>"
			.Write "      <td height=""50"" align=""center"" class='tdbg'> �����෢��</td>"
			.Write "      <td width=""78%"" height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td width=""39%"">"
			.Write "            <select name=""FolderID"" size=12 multiple  style=""width:360"">"
			
									 Call GetSpecialClass
									 
			.Write "              </select></td>"
			.Write "            <td width=""61%"">"
			.Write "              <input name=""Submit22"" type=""submit"" class='button' value=""����ѡ�е�ר�����ҳ &gt;&gt;"" border=""0"">"
			.Write "              <br> <font color=""#FF0000""> ��<br>"
			.Write "              ����ʾ��<br>"
			.Write "              ����ס��CTRL����Shift�������Խ��ж�ѡ</font></td>"
			.Write "          </tr>"
			.Write "        </table></td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "  <form action=""RefreshSpecialSave.asp?Types=ChannelSpecial&RefreshFlag=All"" method=""post"" name=""AllForm"">"
			.Write "    <tr class='tdbg'>"
			.Write "      <td height=""50"" align=""center"" class='tdbg'> ��������ר�����</td>"
			.Write "      <td height=""50"">"
			.Write "        &nbsp;<input name=""SubmitAll"" class='button' type=""submit"" value=""��������ר����� &gt;&gt;"" border=""0"">"
			.Write "      </td>"
			.Write "    </tr>"
			.Write "  </form>"
			.Write "</table>"
			
			.Write "<br><div align='center'><font color=#ff6600>������ʾ������������Ƚ�ռ��ϵͳ��Դ��ʱ�䣬ÿ�η���ʱ�뾡��������������ӵ���Ϣ</font></div>"
			.Write "<br><div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub GetSpecialClass()
			           Dim FolderName, TempStr
					   Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						 RS.Open "Select ClassID,ClassName From KS_SpecialClass Order BY OrderID", Conn, 1, 1
						  If Not RS.EOF Then
							Do While Not RS.EOF
								 FolderName = Trim(RS(1))
								 TempStr = TempStr & "<option value=" & RS(0) & ">" & FolderName & "</option>"
								 RS.MoveNext
							Loop
						  End If
						 RS.Close
					  Set RS = Nothing
					Response.Write TempStr
			End Sub
End Class
%> 

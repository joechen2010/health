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
Set KSCls = New RefreshJS
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshJS
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()

		If Not KS.ReturnPowerResult(0, "KMTL20003") Then                '����ϵͳJS��Ȩ�޼��
			  Call KS.ReturnErr(1, "")
		End If
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>����JS����</title>"
		Response.Write "</head>"
		Response.Write "<script language=""JavaScript"" src=""Common.js""></script>" & vbCrLf
		Response.Write "<script>" & vbCrLf
		Response.Write " function CheckForm(FormObj)" & vbCrLf
		Response.Write " {var tempstr='';" & vbCrLf
		Response.Write " for (var i=0;i<FormObj.TempFolderID.length;i++){" & vbCrLf
		Response.Write "     var KM = FormObj.TempFolderID[i];" & vbCrLf
		Response.Write "    if (KM.selected==true)" & vbCrLf
		Response.Write "       tempstr = tempstr + "" '""+KM.value+""',""" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    if (tempstr=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "    alert('��ѡ����Ҫ������Ƶ��JS!');" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    FormObj.FolderID.value=tempstr.substr(0,(tempstr.length-1));" & vbCrLf
		Response.Write "  return true;" & vbCrLf
		Response.Write "  }" & vbCrLf
		Response.Write "</script>" & vbCrLf
		
		Response.Write "<body topmargin=""0"" leftmargin=""0"" oncontextmenu=""return false;"">"
		Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""25"" class=""Sort"">"
		Response.Write "      <div align=""center""><strong>ϵͳJS��������</strong></div></td>"
		Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "<table width=""100%""  border=""1"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#efefef"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""35"" colspan=""2"">��<strong><font color=""#000099"">����ϵͳJS����</font></strong></td>"
		Response.Write "  </tr>"
		Response.Write "  <form action=""RefreshJSSave.asp?RefreshFlag=Folder"" onsubmit=""return(CheckForm(this))"" method=""post"" name=""ClassForm"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""50"" align=""center""> ��ϵͳJSĿ¼����</td>"
		Response.Write "      <td width=""78%"" height=""50""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		Response.Write "         <tr>"
		Response.Write "            <td width=""39%"">"
		Response.Write "            <input type=""hidden"" name=""FolderID"">"
		Response.Write "              <select name=""TempFolderID"" size=12 multiple id=""TempFolderID"" style=""width:300"">"
		Response.Write "               <option value=""0"" style=""color:red"">��Ŀ¼</option>"
					  
			  Dim TempStr, ID, FolderName
				Dim LabelFolderRS
				
				Set LabelFolderRS = Server.CreateObject("AdoDB.RecordSet")
				LabelFolderRS.Open ("Select ID,FolderName from KS_LabelFolder Where FolderType=2 And ParentID='0' Order By AddDate desc"), Conn, 1, 1
				
				Do While Not LabelFolderRS.EOF
				   ID = Trim(LabelFolderRS(0))
				   FolderName = Trim(LabelFolderRS(1))
				   TempStr = TempStr & "<option value='" & ID & "'>" & FolderName & " </option>"
				   TempStr = TempStr & KS.ReturnSubLabelFolderTree(ID, 0)
				LabelFolderRS.MoveNext
				Loop
				LabelFolderRS.Close
				Set LabelFolderRS = Nothing
			   Response.Write (TempStr)
				
		Response.Write "              </select></td>"
		Response.Write "            <td width=""61%"">"
		Response.Write "              <input name=""Submit22"" type=""submit"" class='button' value="" ����ѡ��Ŀ¼��JS &gt;&gt;"" border=""0"">"
		Response.Write "              <br> <font color=""#FF0000""> ��<br>"
		Response.Write "              ����ʾ��<br>"
		Response.Write "              ����ס��CTRL����Shift�������Խ��ж�ѡ</font></td>"
		Response.Write "         </tr>"
		Response.Write "        </table></td>"
		Response.Write "    </tr>"
		Response.Write "  </form>"
		Response.Write "  <form action=""RefreshJSSave.asp?RefreshFlag=All"" method=""post"" name=""AllForm"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""50"" align=""center""> ��������ϵͳJS</td>"
		Response.Write "      <td height=""50"">"
		Response.Write "        <input name=""SubmitAll""  class='button' type=""submit"" value=""��������ϵͳJS &gt;&gt;"" border=""0"">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "  </form>"
		Response.Write "</table>"
		Response.Write "<br><br><br><div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=http://www.kesion.com/ target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New AddExtJS
KSCls.Kesion()
Set KSCls = Nothing

Class AddExtJS
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'���岿��
		Public Sub Kesion()
		Dim JSID, JSRS, SQLStr, JSName, Descript, JSConfig, JSFlag, ParentID
		Dim Action, RSCheck, JSFileName, FolderID
		With Response
		Set JSRS = Server.CreateObject("Adodb.RecordSet")
		Action = Request.QueryString("Action")
		JSID = Request("JSId")
		FolderID = Trim(Request("FolderID"))
		If JSID = "" Then
		  Action = "Add"
		Else
		  Action = "Edit"
			Set JSRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_JSFile] Where JSID='" & JSID & "'"
			JSRS.Open SQLStr, Conn, 1, 1
			JSName = Replace(Replace(JSRS("JSName"), "{JS_", ""), "}", "")
			Descript = JSRS("Description")
			FolderID = JSRS("FolderID")
			JSConfig = Server.HTMLEncode(Trim(Replace(JSRS("JSConfig"), "GetExtJS,", "")))
			JSFileName = JSRS("JSFileName")
			JSRS.Close
		End If
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.Write "<title>�½�JS</title>"
		.Write "</head>"
		.Write "<script language=""JavaScript"" src=""../../../ks_inc/Common.js""></script>"
		.Write "<script language=""JavaScript"" src=""../../../ks_inc/jQuery.js""></script>"		
		.Write "<link href=""../Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.Write "<script>"
		.Write "function LabelInsertCode(Val)" & vbcrlf
		.Write "{"
		.Write "if (Val!='')"
		.Write "{ document.JSForm.JSConfig.focus();" & vbcrlf
		.Write "  var str = document.selection.createRange();" & vbcrlf
		.Write "  str.text = Val;"
		.Write " }" & vbcrlf
		.Write "}" & vbcrlf
		.Write "</script>"

		.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <form name=""JSForm"" method=""post"" id=""JSForm"" action=""AddJSSave.asp"">"
		.Write "   <input type=""hidden"" name=""JSType"" value=""1"">"
		.Write "   <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.Write "   <input type=""hidden"" name=""JSID"" value=""" & JSID & """>"
		.Write "   <input type=""hidden"" name=""Page"" value=""" & Request("Page") & """>"
		.Write "  <input type=""hidden"" name=""FileUrl"" value=""AddExtJS.asp"">"
		.Write "    <tr> "
		.Write "      <td height=""123"" valign=""top"">" & KS.ReturnJSInfo(JSID, JSName, JSFileName, FolderID, 3, Descript) & "</td>"
		.Write "    </tr>"
		.Write "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""tableBorder1""><strong>�� �� �� �� ̬ JS �� ��</strong></td></tr>"
		
		Response.Write "   <tr class=""tableBorder1"" height=25>"
		 Response.Write "	<td  colspan=""2"">"
		 Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " <select name=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==ѡ��ϵͳ������ǩ==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input type='button' onclick='LabelInsertCode(document.all.mylabel.value);' value='�����ǩ'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;</Td>"
		 Response.Write "      </Tr>"

		
		.Write "    <tr><td colspan=""2"" height=""50""><textarea style=""width:100%"" type=""hidden"" ROWS='17' onfocus='GetJSConfig();' onkeyup='SetEditorValue();' onblur='SetEditorValue();' COLS='108' name=""JSConfig"">" &JSConfig & "</textarea></td></tr>"
		.Write "    <tr>"
		.Write "      <td valign=""top"">"
		'.Write "<iframe id=""JSEditor"" src=""../../KS.Editor.asp?ID=JSConfig&style=2"" scrolling=""no"" width=""100%"" height=""280"" frameborder=""0""></iframe> "
		.Write "</td></tr>"
		.Write "  </form>"
		.Write "</table>"
		.Write "</body>"
		.Write "</html>"
		.Write "<script language=""JavaScript"">"
		.Write "function GetJSConfig()"
		.Write "{"
		'.Write "var TempJSConfig=frames[""JSEditor""].KS_EditArea.document.body.innerHTML;"
		'.Write "TempJSConfig=frames[""JSEditor""].ReplaceUrl(TempJSConfig);"
		'.Write "TempJSConfig=frames[""JSEditor""].Resumeblank(TempJSConfig);"
		'.Write "TempJSConfig=frames[""JSEditor""].ReplaceImgToScript(TempJSConfig);"
		'.Write "TempJSConfig=frames[""JSEditor""].FormatHtml(TempJSConfig);"
		'.Write "document.JSForm.JSConfig.value=TempJSConfig;"
		.Write "}"
		.Write "function SetEditorValue()"
		.Write "{var TempJSConfig=document.JSForm.JSConfig.value;"
		'.Write "TempJSConfig=frames[""JSEditor""].ReplaceRealUrl(TempJSConfig);"
		'.Write "TempJSConfig=frames[""JSEditor""].ReplaceScriptToImg(TempJSConfig);"
		'.Write  "if (TempJSConfig!=frames[""JSEditor""].KS_EditArea.document.body.innerHTML)frames[""JSEditor""].KS_EditArea.document.body.innerHTML=TempJSConfig;"
		.Write "}"
		.Write "function CheckForm()"
		.Write "{ var form=document.JSForm; "
		'.Write "  if (frames['JSEditor'].CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}"
		.Write "  if (form.JSName.value=='')"
		.Write "   {"
		.Write "    alert('������JS����!');"
		.Write "    form.JSName.focus();"
		.Write "    return false;"
		.Write "   }"
		.Write "      if (form.JSFileName.value=='')"
		.Write "      {"
		.Write "       alert('������JS�ļ���');"
		.Write "      form.JSFileName.focus(); "
		.Write "      return false"
		.Write "      }"
		.Write "     if (CheckEnglishStr(form.JSFileName,'JS�ļ���')==false) "
		.Write "       return false;"
		.Write "     if (!IsExt(form.JSFileName.value,'JS'))"
		.Write "       { alert('JS�ļ�������չ��������.js');"
		.Write "          form.JSFileName.focus(); "
		.Write "          return false;"
		.Write "       }"
		'.Write "    form.JSConfig.value=frames[""JSEditor""].ReplaceUrl(frames[""JSEditor""].ReplaceImgToScript(frames[""JSEditor""].Resumeblank(frames['JSEditor'].KS_EditArea.document.body.innerHTML)));"
		.Write "  if (form.JSConfig.value=='')"
		.Write "  {"
		.Write "    alert('������JS����!');"
		.Write "    frames['JSEditor'].KS_EditArea.focus();"
		.Write "    return false;"
		.Write "   }"
		.Write "   form.JSConfig.value='GetExtJS,'+form.JSConfig.value;"
		.Write "   form.submit();"
		.Write "   return true;"
		.Write "}"
		.Write "</script>    "
		End With
		End Sub
End Class
%> 

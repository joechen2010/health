<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%Dim Action,JSID,EditUrl,FolderID
Action=Request.QueryString("Action")
JSID=Request.QueryString("JSID")
FolderID=Request.QueryString("FolderID")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="../Common.js" language="JavaScript"></script>
<link href="../Admin_Style.CSS" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="26" class="sort" align="center">
			    <%IF Action="Edit" Then
			 Response.Write("<Strong>�༭ JS</Strong>")
			 Else
			 %>
     <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%"> ѡ������JS����ϵͳ <input type="radio" name="jstype" onclick="SelectJSType('AddWordJS.asp?JSID=<%=JSID%>&FolderID=<%=FolderID%>&Action=<%=Action%>')" checked value="1">��������
			 <input type=radio name="jstype" value=2 onclick="SelectJSType('AddExtJS.asp?JSID=<%=JSID%>&FolderID=<%=FolderID%>&Action=<%=Action%>')">�Զ��徲̬JS
			</td>
          <td width="50%"><strong>�½�����JS<%=Action%></strong></td>
        </tr>
      </table>
			<%end if%>
</td>
  </tr>
  <tr>
    <td valign="top">
	<%IF Action="Edit" Then
	 EditUrl=Request.Querystring("EditUrl")
	 Response.Write("<iframe src=""" & EditUrl &""" name=""JSFrame"" width=""100%"" height=""100%"" frameborder=""0"" scrolling=""auto""></iframe>")
	else
	 Response.Write("<iframe src=""AddWordJS.asp?JSID=" & JSID & "&FolderID=" & FolderID & "&Action=" & Action &""" name=""JSFrame"" width=""100%"" height=""100%"" frameborder=""0"" scrolling=""auto""></iframe>")
	End IF%>
</td>
  </tr>
</table>
</body>
</html>
<script>
 function SelectJSType(ObjValue)
  {
   frames['JSFrame'].location.href=ObjValue;
  }
</script> 

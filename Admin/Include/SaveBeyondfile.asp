<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<link href="ModeWindow.css" rel="stylesheet" type="text/css">
<link href="Admin_style.css" rel="stylesheet" type="text/css">
<%
Dim KS:Set KS=New PublicCls
Dim Wjj,BH,ext,fname,ItemName
ItemName=KS.G("ItemName")
 if KS.G("wjj")<>"" Then
  Wjj=KS.G("WJJ")
 ELSE
  wjj=request("CurrPath") & "/"
End If
if request("action")="save" then
  call KS.CreateListFolder(wjj)
  http=trim(request.Form("http"))
  if http="" then
   Response.Write"<script>alert('������Զ��" & ItemName &"��ַ!');</script>"
   Response.End()
  end if
  ext=right(http,4)
  fname=wjj&year(now)&month(now)&day(now)&hour(now)&second(now)&KS.MakeRandom(5)&ext
  dim fname1:fname1=fname
  
  ext=lcase(split(fname1,".")(1))
  if (ext<>"jpg" and ext<>"jpeg" and ext<>"gif" and ext<>"bmp" and ext<>"png") or instr(fname1,";")>0 then
  %>
 <script>
    alert('�Բ���,ֻ�ܱ���ͼƬjpg|jpeg|gif|png���ļ�!');
   window.close();
 </script>
  <%
   response.end
  end if

  
  Call KS.SaveBeyondFile(fname1,http)
%>
 <script>
    alert('�ɹ�������Զ��<%=ItemName%>!');
   window.returnValue='<%=fname%>';
   window.close();
 </script>

<%
  Response.Write("Զ��" & ItemName &"����ɹ�!")
end if
%>
<script>
  function document.onreadystatechange()
 {
    document.myform.http.focus();
 }
   window.onunload=SetReturnValue;
	function SetReturnValue()
	{
		if (typeof(window.returnValue)!='string') window.returnValue='';
	}
</script>
<div align="center">
<br>
<form name="myform" action="?action=save" method="post">
<input type="hidden" name="ItemName" value="<%=ItemName%>" />
<input type="hidden" value="<%=wjj%>" name="wjj" />
Զ��<%=ItemName%>��ַ��<input type="text" name="http">
<input type="submit" name="Submit" class="button" value="��ʼץȡ" onclick="if (document.myform.http.value==''){alert('������Զ��<%=ItemName%>��ַ��');document.myform.http.focus(); return false;}"><br><br>
����:<font color=red>http://www.kesion.com/images/logo.gif</font>
</form>
</div>
 

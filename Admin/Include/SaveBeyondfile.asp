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
   Response.Write"<script>alert('请输入远程" & ItemName &"地址!');</script>"
   Response.End()
  end if
  ext=right(http,4)
  fname=wjj&year(now)&month(now)&day(now)&hour(now)&second(now)&KS.MakeRandom(5)&ext
  dim fname1:fname1=fname
  
  ext=lcase(split(fname1,".")(1))
  if (ext<>"jpg" and ext<>"jpeg" and ext<>"gif" and ext<>"bmp" and ext<>"png") or instr(fname1,";")>0 then
  %>
 <script>
    alert('对不起,只能保存图片jpg|jpeg|gif|png的文件!');
   window.close();
 </script>
  <%
   response.end
  end if

  
  Call KS.SaveBeyondFile(fname1,http)
%>
 <script>
    alert('成功保存了远程<%=ItemName%>!');
   window.returnValue='<%=fname%>';
   window.close();
 </script>

<%
  Response.Write("远程" & ItemName &"保存成功!")
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
远程<%=ItemName%>地址：<input type="text" name="http">
<input type="submit" name="Submit" class="button" value="开始抓取" onclick="if (document.myform.http.value==''){alert('请输入远程<%=ItemName%>地址！');document.myform.http.focus(); return false;}"><br><br>
形如:<font color=red>http://www.kesion.com/images/logo.gif</font>
</form>
</div>
 

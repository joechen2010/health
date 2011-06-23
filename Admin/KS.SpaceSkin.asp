<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdminiStratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New BlogUserSkin
KSCls.Kesion()
Set KSCls = Nothing

Class BlogUserSkin
        Private KS,flag,KSCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion()
		dim action:Action=trim(request("Action"))
		flag=KS.chkclng(KS.g("flag"))

		 With Response
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../ks_inc/kesion.box.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../ks_inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>空间管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=" & flag & "';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>模板管理</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>空间分类</span></li>"
			  .Write "</ul>"
		End With
		
		select case Action
		    case "newtemplate","modifytext"
			    call textTemplate()
			case "saveaddtext"
			    call saveaddtext()
			case "savetext"
			    call savetext()	
			case "savedefault"
				call savedefault()
			case "delconfig"
				call delconfig()
			case else
				call showconfig()
		end select
	 End Sub

sub showconfig()
dim rs:set rs=conn.execute("select * From KS_BlogTemplate where flag=" & flag & " order by usertag")
%>
<script type="text/javascript">
$(document).ready(function(){
 $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
 $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
})
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"0>
<form name="form2" method="post" action="KS.SpaceSkin.asp?action=savedefault&flag=<%=KS.g("flag")%>">
  <tr class="sort"> 
      <td width="6%"><div align="center">ID</div></td>
      <td width="20%" ><div align="center">名称</div></td>
      <td width="15%" ><div align="center">作者</div></td>
      <td width="12%" ><div align="center">默认模版</div></td>
      <td width="12%" ><div align="center">类型</div></td>
      <td width="47%" > 
        <div align="center">模版管理</div></td>
  </tr>
      <% 
while not rs.eof	  
%> 
    <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
            <td class="splittd" align="center"><%= rs("id") %>&nbsp;</td>
            <td class="splittd" align="center"><%= rs("TemplateName") %></td>
            <td class="splittd" align="center"><%= rs("TemplateAuthor") %></td>
            <td class="splittd" align="center"> <input name="radiobutton" type="radio" value='<%=rs("id")%>' <%if rs("isdefault")="true" then response.Write "checked" %>></td>
			<td class="splittd" align="center">
			 <% if rs("usertag")=1 then
			     response.write "<font color=blue>用户上传</font>"
				else
				 response.write "<font color=red>系统自带</font>"
			 end if%>
			</td>
            
      <td class="splittd" width="40%"> <div align="center"><a href="../space/showtemplate.asp?templateid=<%=rs("id")%>" target="_blank">预览</a>
	  　<a onclick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户 >> <font color=red>修改模板</font>')+'&ButtonSymbol=GOSave';location.href='KS.SpaceSkin.asp?action=modifytext&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>'" href="#">修改模版</a>　<a href="KS.SpaceSkin.asp?action=delconfig&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>" onclick=return(confirm("确定要删除这个模版吗？"))>删除模版</a></div>
	  </td>
    </tr>
      <%
rs.movenext
wend
%>
    <td height="40" colspan="5" align="center">  
        <div align="center">
          <input type="submit" name="Submit" class="button" value="保存设置">&nbsp;&nbsp;
		  <input type="button" name="Submit1" class="button" value="添加新模板" onClick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('空间门户 >> <font color=red>添加模板</font>')+'&ButtonSymbol=GO';location.href='?Action=newtemplate&flag=<%=KS.g("flag")%>';">
      </div></td>
  </tr>
</form>  
</table>
<%
	set rs=nothing
end sub
'添加新模板
sub TextTemplate()
   Dim CurrPath
   CurrPath=KS.GetCommonUpFilesDir()
  dim templatename,templateauthor,templatemain,templatesub,Action,templatepic,GroupID
 if KS.g("action")="modifytext" then
  dim rs:set rs=server.createobject("adodb.recordset")
  rs.open "select * from KS_BlogTemplate Where ID="&KS.chkclng(KS.G("id")),conn,1,1
  if not rs.eof then
   templatename=rs("templatename")
   templateauthor=rs("templateauthor")
   templatepic=rs("templatepic")
   templatemain=rs("templatemain")
   templatesub=rs("templatesub")
   GroupID=rs("GroupID")
  end if
  Action="savetext"
 else
  Action="saveaddtext"
 end if
%>
<script>
 function CheckForm()
 {
    if (document.myform.TemplateName.value=='')
	{
	  alert('请输入模板名称!');
	  document.myform.TemplateName.focus();
	  return false;
	}
    if (document.myform.TemplateMain.value=='')
	{
	  alert('请输入主模板的内容!');
	  document.myform.TemplateMain.focus();
	  return false;
	}
    if (document.myform.TemplateSub.value=='')
	{
	  alert('请输入副模板的内容!');
	  document.myform.TemplateSub.focus();
	  return false;
	}
	 document.myform.submit();
 }
function ShowIframe(flag)
        {   
		 onscrolls=false;
         PopupCenterIframe("查看空间站点的可用标签","../ks_editor/spacelabel.asp?flag="+flag,590,340,'no')
       }
function InsertLabel(obj,Val)
{ return false;
  $(obj).focus();
  var str = document.selection.createRange();
  str.text = Val; 
  closeWindow();
 }
</script>
   <div style='height:35px;line-height:35px;text-align:center;font-weight:bold'>模板注册</div>
  <table width="100%" border="0" align="center" class="ctable" cellpadding="2" cellspacing="1">
<form method="POST" action="KS.SpaceSkin.asp?ID=<%=KS.G("id")%>&flag=<%=KS.g("flag")%>&action=<%=Action%>" id="myform" name="myform">
    <tr class="tdbg">
      <td align="right" class="clefttitle" height="25"><strong>模版名称：</strong></td>
	  <td><input name="TemplateName" type="text" id="TemplateName" value="<%=templatename%>"></td>
	</tr>
	<tr class="tdbg">
	  <td align="right" class="clefttitle"><strong>模板作者：</strong></td>
	  <td><input name="TemplateAuthor" type="text" id="TemplateAuthor" value="<%=templateauthor%>"> 如科汛,kesion.com等</td>
	</tr>
	<tr class="tdbg">
	   <td align="right" class="clefttitle"><strong>预 览 图：</strong></td>
	   <td><input type="text" name="TemplatePic" value="<%=templatepic%>">&nbsp;<input class='button' type='button' name='Submit' value='选择预览图...' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,document.all.TemplatePic);">
		</td>
    </tr>
	
	 <tr> 
	  <td height="25" class="clefttitle" width="120" align="right"><strong>首页模板：</strong></td>
      <td height="25" class="tdbg">
	  <input type="text" name="TemplateMain" id='TemplateMain' size='25' value="<%=templateMain%>"> <%=KSCls.Get_KS_T_C("$('#TemplateMain')[0]")%>
      </td>
    </tr>

	
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" width="120" align="right"><strong>其它页框架模板：</strong></td>
      <td height="25">
	  <input type="text" name="TemplateSub" id='TemplateSub' size='25' value="<%=templateSub%>"> <%=KSCls.Get_KS_T_C("$('#TemplateSub')[0]")%> 
      </td>
    </tr>
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" width="120" align="right">
	  <strong>允许使用本模板的用户组：</strong>
	  </td>
      <td height="25">
	 <%=KS.GetUserGroup_CheckBox("GroupID",GroupID,4)%>
      </td>
    </tr>
</form>
  </table>
  
  <iframe src="../ks_editor/spacelabel.asp?flag=<%=flag%>" width="100%" height="300"></iframe>

<%
end sub

sub savetext()
	dim rs,sql,flag
	set rs=server.CreateObject("adodb.recordset")
	sql="select * From KS_BlogTemplate where id=" & KS.chkclng(KS.g("id"))
	rs.open sql,conn,1,3
	rs("TemplateName")=trim(request("TemplateName"))
	rs("TemplateAuthor")=trim(request("TemplateAuthor"))
	rs("TemplateMain")=request("TemplateMain")
	rs("TemplatePic")=request("TemplatePic")
	rs("templatesub")=request("TemplateSub")
	rs("GroupID")=replace(request("GroupID")," ","")
	rs.update
	flag=rs("flag")
	Call KS.FileAssociation(1013,KS.chkclng(KS.g("id")),request("TemplatePic"),1)
	rs.close:set rs=nothing
	response.Write  "<script>alert('模板修改成功!');location.href='KS.SpaceSkin.asp?flag=" & flag & "';</script>"
end sub
sub saveaddtext()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * From KS_BlogTemplate"
	rs.open sql,conn,1,3
	rs.addnew
	rs("TemplateName")=trim(request("TemplateName"))
	rs("TemplateAuthor")=trim(request("TemplateAuthor"))
	rs("TemplatePic")=request("TemplatePic")
	rs("TemplateMain")=request("TemplateMain")
	rs("templatesub")=request("TemplateSub")
	rs("flag")=KS.chkclng(KS.g("flag"))
	rs("GroupID")=replace(request("GroupID")," ","")
	rs.update
	rs.movelast
	Call KS.FileAssociation(1013,rs("id"),request("TemplatePic"),0)
	rs.close:set rs=nothing
	response.Write  "<script>if (confirm('模板添加成功,继续添加吗？')==true){location.href='KS.SpaceSkin.asp?flag=" & ks.g("flag") & "&action=newtemplate';}else{location.href='KS.SpaceSkin.asp?flag=" & KS.g("flag") & "';}</script>"
end sub
sub savedefault()
	dim rs,isdefaultID
	isdefaultID=KS.ChkCLng(trim(request("radiobutton")))
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select id,isdefault From KS_BlogTemplate where flag=" & KS.chkclng(KS.g("flag")),conn,1,3
	while not rs.eof
		if isdefaultID=rs("id") then
			rs("isdefault")="true"
		else
			rs("isdefault")="false"
		end if
		rs.update
		rs.movenext
	wend
	rs.close
	set rs=nothing
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub


sub delconfig()
    conn.execute("delete from ks_UploadFiles where channelid=1013 and infoid=" & KS.ChkClng(KS.G("ID")))
	conn.execute("delete From KS_BlogTemplate where id="&KS.ChkCLng(KS.G("id")))
	response.Redirect "KS.SpaceSkin.asp?action=showconfig&flag=" &KS.g("flag")	
end sub


End Class
%>
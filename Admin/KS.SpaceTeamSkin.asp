<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		dim action:Action=trim(request("Action"))
		flag=3        'Ȧ��ģ��
		 With Response
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeam.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>Ȧ�ӹ���</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeam.asp?action=topic';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>���ӹ���</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeamSkin.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>ģ�����</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceTeam.asp?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>Ȧ�ӷ���</span></li>"
			  .Write	" </ul>"
		End With
		
		select case Action
		    case "newtemplate", "modifytext"
			    call textTemplate()
			case "saveaddtext"
			    call saveaddtext()
			case "savetext"
			    call savetext()	
			case "saveconfig" 
				call saveconfig()
			case "savedefault"
				call savedefault()
			case "delconfig"
				call delconfig()
			case else
				call showconfig()
		end select
	 End Sub

sub showconfig()
dim rs:set rs=conn.execute("select * From KS_BlogTemplate where flag=" & flag)
%><script type="text/javascript">
$(document).ready(function(){
 $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
 $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
})
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"0>
<form name="form2" method="post" action="KS.SpaceTeamSkin.asp?action=savedefault&flag=<%=KS.g("flag")%>">
  <tr class="sort"> 
      <td width="6%"><div align="center">ID</div></td>
    <td width="20%" ><div align="center">����</div></td>
    <td width="15%" ><div align="center">����</div></td>
      <td width="12%" ><div align="center">Ĭ��ģ��</div></td>
      <td width="47%" > 
        <div align="center">ģ�����</div></td>
  </tr>
      <% 
while not rs.eof	  
%> 
    <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
            <td class="splittd"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
          <td class="splittd"><div align="center"><%= rs("TemplateName") %></div></td>
          <td  class="splittd">&nbsp;<div align="center"><%= rs("TemplateAuthor") %></div></td>
          <td  class="splittd">
			<div align="center"> 
                <input name="radiobutton" type="radio" value='<%=rs("id")%>' <%if rs("isdefault")="true" then response.Write "checked" %>>
            </div></td>
            
      <td width="40%" class="splittd"> <div align="center"><a href="../space/showtemplate.asp?templateid=<%=rs("id")%>" target="_blank">Ԥ��</a>
	  <a onclick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('�ռ��Ż� >> <font color=red>�޸�ģ��</font>')+'&ButtonSymbol=GOSave';location.href='KS.SpaceTeamSkin.asp?action=modifytext&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>'" href="javascript:void(0)">�޸�ģ��</a>��<a href="KS.SpaceTeamSkin.asp?action=delconfig&id=<%=rs("id")%>&flag=<%=KS.g("flag")%>" onclick=return(confirm("ȷ��Ҫɾ�����ģ����"))>ɾ��ģ��</a></div>
	  </td>
    </tr>
      <%
rs.movenext
wend
%>
    <td height="40" colspan="5" align="center">  
        <div align="center">
          <input type="submit" name="Submit" class="button" value="��������">&nbsp;&nbsp;
		  <input type="button" name="Submit1" class="button" value="�����ģ��" onClick="$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('�ռ��Ż� >> <font color=red>���ģ��</font>')+'&ButtonSymbol=GO';location.href='?Action=newtemplate&flag=<%=KS.g("flag")%>';">
      </div></td>
  </tr>
</form>  
</table>
<%
	set rs=nothing
end sub
'�����ģ��
sub TextTemplate()
     Dim CurrPath
   CurrPath=KS.GetCommonUpFilesDir()
dim templatename,templateauthor,templatemain,templatesub,Action,templatepic,GroupID
  redim templatesub(10)
 if KS.g("action")="modifytext" then
  dim rs:set rs=server.createobject("adodb.recordset")
  rs.open "select * from KS_BlogTemplate Where ID="&KS.chkclng(KS.g("id")),conn,1,1
  if not rs.eof then
   templatename=rs("templatename")
   templateauthor=rs("templateauthor")
   templatepic=rs("templatepic")
   templatemain=rs("templatemain")
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
	  alert('������ģ������!');
	  document.myform.TemplateName.focus();
	  return false;
	}
    if (document.myform.TemplateMain.value=='')
	{
	  alert('����ѡ��ģ��!');
	  document.myform.TemplateMain.focus();
	  return false;
	}
    
	 document.myform.submit();
 }
</script>

   <div style='height:35px;line-height:35px;text-align:center;font-weight:bold'>Ȧ��ģ��ע��</div>

  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1">
<form method="POST" action="KS.SpaceTeamSkin.asp?ID=<%=KS.G("id")%>&flag=<%=KS.g("flag")%>&action=<%=Action%>" id="myform" name="myform">
    <tr class="tdbg">
      <td align="right" class="clefttitle" height="25"><strong>ģ�����ƣ�</strong></td>
	  <td><input name="TemplateName" type="text" id="TemplateName" value="<%=templatename%>"></td>
	</tr>
	<tr class="tdbg">
	  <td align="right" class="clefttitle"><strong>ģ�����ߣ�</strong></td>
	  <td><input name="TemplateAuthor" type="text" id="TemplateAuthor" value="<%=templateauthor%>"> ���Ѵ,kesion.com��</td>
	</tr>
	<tr class="tdbg">
	   <td align="right" class="clefttitle"><strong>Ԥ �� ͼ��</strong></td>
	   <td><input type="text" name="TemplatePic" value="<%=templatepic%>">&nbsp;<input class='button' type='button' name='Submit' value='ѡ��Ԥ��ͼ...' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,document.all.TemplatePic);">
		</td>
    </tr>
	 <tr> 
	  <td height="25" class="clefttitle" width="120" align="right"><strong>�� ģ �壺</strong></td>
      <td height="25" class="tdbg">
	  <input type="text" name="TemplateMain" id='TemplateMain' size='25' value="<%=templateMain%>"> <%=KSCls.Get_KS_T_C("$('#TemplateMain')[0]")%> ��ģ������{$GroupMain}��ǩ
      </td>
    </tr>
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" width="120" align="right">
	  <strong>����ʹ�ñ�ģ����û��飺</strong>
	  </td>
      <td height="25">
	 <%=KS.GetUserGroup_CheckBox("GroupID",GroupID,4)%>
      </td>
    </tr>

</form>
  </table>

<%call LabelHelp()
end sub


sub savedefault()
	dim rs,isdefaultID
	isdefaultID=KS.ChkCLng(trim(request("radiobutton")))
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select id,isdefault From KS_BlogTemplate where flag=3",conn,1,3
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
	Response.Write"alert(""�޸ĳɹ���"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub



sub delconfig()
	conn.execute("delete From KS_BlogTemplate where id="&KS.ChkCLng(KS.G("id")))
	response.Redirect "KS.SpaceTeamSkin.asp?action=showconfig&flag=" &KS.g("flag")	
end sub


sub LabelHelp()
%>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1">
    <tr>
      <td height="25" class="sort">���ñ�ǩ˵��:</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">
        <table border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td width="150" colspan=2><font color=red>��ģ����ñ�ǩ˵��</font></td>
		 </tr>
		 <tr>
	       <td><li>{$GroupMain}</td><td>---��ʾ���岿�֣������б�ȣ���</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowNavigation}</td><td>---��ʾȦ�ӵ������ȡ�</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowGroupInfo}</td><td>---��ʾȦ����Ϣ��</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowNewUser}</td><td>---��ʾ���¼����Ա�б�</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowActiveUser}</td><td>---��ʾ�����Ծ��Ա�б�</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowAnnounce}</td><td>---��ʾ���¹��档</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowUserLogin}</td><td>---��ʾ��Ա��¼��</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowGroupName}</td><td>---��ʾȦ�����ơ�</td>
		 </tr>
		 <tr>
	       <td><li>{$ShowGroupURL}</td><td>---��ʾȦ��URL��</td>
		 </tr>
		</table>
      </td>
    </tr>
    <tr> 
      <td class="tdbg"></td>
    </tr>
  </table>
<%end sub

 sub savetext()
	dim rs,sql,flag
	set rs=server.CreateObject("adodb.recordset")
	sql="select * From KS_BlogTemplate where id=" & KS.chkclng(KS.g("id"))
	rs.open sql,conn,1,3
	rs("TemplateName")=trim(request("TemplateName"))
	rs("TemplateAuthor")=trim(request("TemplateAuthor"))
	rs("TemplateMain")=request("TemplateMain")
	rs("TemplatePic")=request("TemplatePic")
	rs("groupid")=replace(request("groupid")," ","")
	rs.update
	flag=rs("flag")
	rs.close:set rs=nothing
	response.Write  "<script>alert('ģ���޸ĳɹ�!');location.href='KS.SpaceTeamSkin.asp?flag=" & flag & "';</script>"
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
	rs("groupid")=replace(request("groupid")," ","")
	rs("flag")=3
	rs.update
	rs.close:set rs=nothing
	response.Write  "<script>if (confirm('ģ����ӳɹ�,���������')==true){location.href='KS.SpaceTeamSkin.asp?action=newtemplate';}else{location.href='KS.SpaceTeamSkin.asp?flag=" & KS.g("flag") & "';}</script>"
end sub
End Class
%>
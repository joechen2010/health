<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Group
KSCls.Kesion()
Set KSCls = Nothing

Class Group
        Private KS,KSBCls,KSUser
		Private PerPageNumber,CurrPage,totalPut,RS,MaxPerPage
		Private ID,Template,TemplateID,TeamName,groupadmin
		Private Sub Class_Initialize()
		  MaxPerPage =15
		  Set KS=New PublicCls
		  Set KSBCls=New BlogCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSBCls=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 If KS.SSetting(0)=0 Then
		 Response.Write "<script>alert('�Բ��𣬱�վ��رո��˿ռ书��!');window.close();</script>"
		 Response.end
		 End If
		ID=KS.ChkClng(KS.S("ID"))
		If ID=0 Then Response.End()
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Team Where ID=" & ID,conn,1,1
		If RS.Eof And RS.Bof Then
		 Response.Write "<script>alert('�������ݳ���!');window.close();</script>"
		 Response.end
		End If
		If RS("Verific")=0 Then
		 Response.Write "<script>alert('��Ȧ����δ���!');window.close();</script>"
		 response.end
		elseif RS("Verific")=2 then
		 Response.Write "<script>alert('��Ȧ���ѱ�����Ա����!');window.close();</script>"
		 response.end
		end if
		
		 TeamName=RS("TeamName")
		 groupadmin=rs("username")
		 TemplateID=RS("TemplateID")
		 Template="<html>"&vbcrlf &"<title>" & TeamName & "</title>" &vbcrlf
		 Template=Template & "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" &vbcrlf
         Template=Template & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbcrlf
         Template=Template & "<meta name=""generator"" content=""KesionCMS"" />" & vbcrlf
		 Template=Template & "<meta name=""author"" content=""" & RS("UserName") & ","" />" & vbcrlf
		 Template=Template & "<meta name=""keyword"" content=""" & TeamName & """ />"&VBCRLF
		 Template=Template & "<meta name=""description"" content=""KesionCMS,Kesion,Kesion.com"" />"  & vbcrlf
		 Template=Template & "<link href=""css/css.css"" type=""text/css"" rel=""stylesheet"">" & vbcrlf
		 Template=Template & "<script src=""../ks_inc/jquery.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""../ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		 Template=Template & "<script src=""js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		 template=Template & KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 template=KSBCls.ReplaceGroupLabel(RS,Template)
		 
		 Select Case KS.S("Action")
		  case "showtopic"
		   	Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'showtopic&teamid=" & id & "&tid=" & KS.S("tid") & "&groupadmin=" & groupadmin & "')</script><div id=""teammain""></div><div id=""kspage"" align=""right""></div>" &  showtopic)
		  case "replaysave"
		   call replaysave
		  case "users"
		   	Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'users&teamid=" & id & "')</script><div id=""teammain""></div><div id=""kspage"" style=""clear:both"" align=""right""></div>")
		 case "join"
		  		Template=Replace(Template,"{$GroupMain}",showjoin())
		 case "joinsave"
		    call joinsave
		 case "deltopic"
		    call deltopic
		 case "deluser"
		    call deluser()
		 case "settop"
		   call settop()
		 case "setbest"
		   call setbest()
		 case "post"
		  	Template=Replace(Template,"{$GroupMain}",showpost())
		 case "topicsave"
		   call topicsave()
		 case "info"
		  	Template=Replace(Template,"{$GroupMain}",showinfo())
		 case else
		 Template=Replace(Template,"{$GroupMain}","<script language=""javascript"" defer>TeamPage(1,'teamtopic&teamid=" & id & "&isbest=" & KS.R(KS.S("isbest")) &"')</script><div id=""teammain""></div><div id=""kspage"" align=""right""></div>")
		  end select
		 Response.Write Template
		  RS.Close
          Set  RS=Nothing
		End Sub
		
		function showtopic()
		 dim tid:tid=KS.chkclng(KS.S("tid"))
		showtopic=showtopic &"<div id=""form_comment""><a name=""add_comment""></a>"
		showtopic=showtopic &"<script type=""text/javascript"">function checkform(){if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)==''){alert('������ظ�����!');FCKeditorAPI.GetInstance('Content').Focus();return false;}return true;}</script>"
		showtopic=showtopic &"<br/>"
		showtopic=showtopic &"<table width=""99%"" cellpadding=""1"" cellspacing=""1"" bgcolor=""#efefef"">"
		showtopic=showtopic &"<form action='group.asp?action=replaysave&id=" & id & "&tid=" & tid & "' method='post' name='myform' id='myform' onSubmit=""return(checkform())"">"
		showtopic=showtopic &"    <tr>"
		showtopic=showtopic &"	  <td colspan=""2"" bgcolor=""#EDF5F9""><strong>�ظ�����</strong></td>"
		showtopic=showtopic &"	</tr>"
			IF Cbool(KSUser.UserLoginChecked)=false Then
		showtopic=showtopic &"    <tr>"
		showtopic=showtopic &"	  <td colspan=""2"" bgcolor=""#FFFFFF"" align=""center"" height=""80""><p>��¼��ſ��Բ���û��������,��Ҫ������������<a href=""../user/login/"" target=""_blank"">��¼</a>����Ա���ģ�</p></td>"
		showtopic=showtopic &"	</tr>"
			else
			on error resume next
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td width=""100"" align=""center"" bgcolor=""#FFFFFF"">�ظ����⣺</td>"
		showtopic=showtopic &"		<td bgcolor=""#FFFFFF""><input type=""text"" readonly value=""Re:" & conn.execute("select title from ks_teamtopic where id="& tid )(0) & """ size=""50"" name=""title"">"
        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td width=""100"" align=""center"" bgcolor=""#FFFFFF"">����:</td>"
		showtopic=showtopic &"		<td bgcolor=""#FFFFFF"">"
		showtopic=showtopic &"		<textarea id=""Content"" name=""Content"" style=""display:none""></textarea>"
		showtopic=showtopic &"<iframe id='Content___Frame' src='../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic' width='98%' height='180' frameborder='0' scrolling='no'></iframe>"
        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
		showtopic=showtopic &"	<tr>"
		showtopic=showtopic &"		<td colspan=""2"" align=""center"" bgcolor=""#FFFFFF"">"
		showtopic=showtopic &"		<input type=""submit"" value=""�ύ�ظ�"">"
        showtopic=showtopic &"        </td>"
		showtopic=showtopic &"	</tr>"
			end if
		showtopic=showtopic &"	</form>"
	    showtopic=showtopic &"</table>"
		showtopic=showtopic &"</div>"
		end function
		

		'����ظ�
		Sub replaysave()
		dim tid:tid=KS.chkclng(KS.S("tid"))
		dim title:title=KS.S("title")
		dim content:content=KS.S("content"):if content="" then call KS.alert("������ظ�����!",""):exit sub
		IF Cbool(KSUser.UserLoginChecked)=false Then  call KS.alert("���ȵ�¼!",""):exit sub
		dim username:username=KS.R(KSUser.UserName)
		dim rs:set rs=server.createobject("adodb.recordset")
		rs.open "select top 1 * from ks_teamtopic",conn,1,3
		rs.addnew
		 rs("parentid")=tid
		 rs("teamid")=id
		 rs("title")=title
		 rs("content")=content
		 rs("adddate")=now
		 rs("userip")=KS.getip
		 rs("status")=1
		 rs("username")=username
		 rs("isbest")=0
		  rs("istop")=0
		rs.update
		rs.movelast
		Call KS.FileAssociation(1031,rs("ID"),content,0)
		rs.close:set rs=nothing
		response.redirect request.servervariables("http_referer")
		End Sub
		
		Function showjoin()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  showjoin= "�Բ����������Ȧ��֮ǰ������<a href=""../user/login/"" target=""_blank"">��¼</a>����Ա���ģ�"
		  exit function
		 end if
		 if not conn.execute("select username from ks_teamusers where username='" & ksuser.username & "' and teamid=" & id).eof then
		  showjoin= "<div><b>�����������룬�����Ŀ���ԭ�����£�</b><div style='border:1px solid #efefef;overflow:hidden'></div><li>�����������δ�õ�Ȧ�������;</li><li>�����Ǳ�Ȧ�ӵĳ�Ա������Ҫ������;</li><li>�������ѱ�Ȧ�����룬������δ�ڻ�Ա����ȷ��;</li></div>"
		  showjoin=showjoin & "<div><b>������֪��</b><div style='border:1px solid #efefef;overflow:hidden'></div>"
		  showjoin=showjoin & RS("Note")
		  showjoin=showjoin & "</div>"
		  exit function
		 end if
		  showjoin=showjoin & "<script>"
		  showjoin=showjoin & " function checkform()"
		  showjoin=showjoin & " {if (document.myform.username.value==''){"
		  showjoin=showjoin & "	 alert('�����˱�����д!');"
		  showjoin=showjoin & "	 document.myform.username.focus();"
		  showjoin=showjoin & "	 return false"
		  showjoin=showjoin & "	 }"
		  showjoin=showjoin & "	 if (document.myform.reason.value==''){"
		  showjoin=showjoin & "	  alert('���������Ȧ�ӵ�����!');"
		  showjoin=showjoin & "	  document.myform.reason.focus();"
		  showjoin=showjoin & "	  return false"
		  showjoin=showjoin & "	  }"
		  showjoin=showjoin & "	  return true;"
		  showjoin=showjoin & " }"
		  showjoin=showjoin & "</script>"
		  showjoin=showjoin & "<table width=""100%"" cellspacing=""0"" cellspadding=""0"" border=""0"">"
		  showjoin=showjoin & " <form name=""myform"" action=""?id=" & id & "&action=joinsave"" method=""post"" onSubmit=""return(checkform())""> "
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td align=""center"" bgcolor=""#f9f9f9"" colspan=2>�� �� �� �� Ⱥ ��</td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td width=""100"">�� �� �ˣ�</td>"
		  showjoin=showjoin & "	  <td><input name=""username"" type=""textbox"" value=""" & ksuser.username & """ readonly size=10></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td>�������ɣ�</td>"
		  showjoin=showjoin & "	  <td><textarea name=""reason"" cols=""50"" rows=""6""></textarea></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	<tr>"
		  showjoin=showjoin & "	  <td colspan=2 align=""center""><input type=""submit"" value=""�ύ����""></td>"
		  showjoin=showjoin & "	</tr>"
		  showjoin=showjoin & "	</form>"
		  showjoin=showjoin & "</table>"
		  showjoin=showjoin & "<div><b>������֪��</b><div style='border:1px solid #efefef;overflow:hidden'></div>"
		  showjoin=showjoin & RS("Note")
		  showjoin=showjoin & "</div>"
		End Function
		
		'��������
		Sub JoinSave()
		dim id:id=KS.chkclng(KS.S("id"))
		dim username:username=KS.R(KS.S("username"))
		dim reason:reason=KS.R(KS.S("reason"))
		dim rs:set rs=server.createobject("adodb.recordset")
		rs.open "select * from ks_teamusers where teamid=" & id & " and username='" & username & "'",conn,1,3
		if rs.eof then
		 rs.addnew
		  rs("teamid")=id
		  rs("username")=username
		  rs("status")=2  '�������
		  rs("power")=0   '��ͨ�û�
		  rs("reason")=reason
		  rs("Applydate")=now
		 rs.update
		end if
		rs.close:set rs=nothing
		call KS.alert("����������ύ����ȴ�Ȧ�������!","?id=" & id)
		End Sub
		
		'��������
		function showpost()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  showpost= "�Բ��𣬷�������֮ǰ������<a href=""../User/"" target=""_blank"">��¼</a>����Ա���ģ�"
		  exit function
		 end if
		 if conn.execute("select username from ks_teamusers where username='"& ksuser.username & "' and teamid=" & id).eof then
		  showpost= "�Բ����㲻�Ǹ�Ȧ�ӵĳ�Ա��û��Ȩ�������⣡"
		  exit function
		 elseif conn.execute("select username from ks_teamusers where username='"& ksuser.username & "' and status<>2 and teamid=" & id).eof then
		  showpost= "�Բ������ύ�����뻹δ�õ�ȷ�ϣ�û��Ȩ�������⣡"
		  exit function
		 end if

		showpost="<script>"
		showpost=showpost & "function checkform()"
		showpost=showpost & " {"
		showpost=showpost & "  if (document.myform.topic.value=='')"
		showpost=showpost & "  {"
		showpost=showpost & "   alert('���������ۻ���!');"
		showpost=showpost & "   document.myform.topic.focus();"
		showpost=showpost & "  return false;"
		showpost=showpost & "  }"
		showpost=showpost & "  if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=='')"
		showpost=showpost & "  {"
		showpost=showpost & "   alert('��������������!');"
		showpost=showpost & "   FCKeditorAPI.GetInstance('Content').Focus();"
		showpost=showpost & "   return false;"
		showpost=showpost & "  }"
		showpost=showpost & "  return true;"
		showpost=showpost & " }"
		showpost=showpost & "</script>"
		showpost=showpost & "<div id=""form_comment"">"
		showpost=showpost & "<form action='group.asp?action=topicsave&id=" & id & "' onSubmit=""return(checkform())"" method='post' name='myform' id='myform'>"
		showpost=showpost & "<div id=""ad_teamcomment""></div><ul><p>" & ksuser.username &" , ��ӭ������Ȧ������!</p></ul><ul>����Ȧ�ӳ�Ա���Է������⣬�ǳ�Ա�����Իظ�</ul><ul>�ǳƣ�<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='" & ksuser.username & "' readonly /></ul>"
		showpost=showpost & "<ul>���⣺<input name='topic' type='text' id='topic' size='50' maxlength='50' value='' /></ul>"
		showpost=showpost & "<ul>"
		showpost=showpost & "<div><textarea id=""Content"" name=""Content"" style=""width:400px;height:250px; display:none"" ></textarea><iframe id='Content___Frame' src='../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic' width='98%' height='200' frameborder='0' scrolling='no'></iframe></div> "
		showpost=showpost & "</ul>"
		showpost=showpost & "<ul><input type='submit' value='OK,���� '></ul>"
		showpost=showpost & "</form>"
		showpost=showpost & "</div>"
		showpost=showpost & "</div>"
		End Function
		'���淢��
		Sub topicsave()
		 dim id:id=KS.chkclng(KS.S("id"))
		 dim topic:topic=KS.R(KS.S("topic"))
		 dim content:content=KS.HTMLEncode(KS.S("content"))
		 dim rs:set rs=server.createobject("adodb.recordset")
		 rs.open "select top 1 * from ks_teamtopic",conn,1,3
		 rs.addnew
		  rs("title")=topic
		  rs("content")=content
		  rs("teamid")=id
		  rs("parentid")=0
		  rs("username")=KS.S("username")
		  rs("adddate")=now
		  rs("userip")=KS.getip
		  rs("status")=1
		  rs("isbest")=0
		  rs("istop")=0
		 rs.update
		 rs.movelast
		 Call KS.FileAssociation(1031,rs("ID"),content,0)
		 rs.close:set rs=nothing
		 response.write "<script>alert('�������ۻ��ⷢ��ɹ���');location.href='?id=" & id &"';</script>"
		End Sub	
		
		
		'Ȧ����Ϣ
		function showinfo()
		showinfo="<div id=""ginfo"">"
		showinfo=showinfo &"	<h1>Ȧ����Ϣ</h1>"
		showinfo=showinfo &"<div id=""group_info"">"
		showinfo=showinfo &"	<div><img src=""" & rs("photourl") & """ border=""0""></div>"
		showinfo=showinfo &"	<div>"
		showinfo=showinfo &"	<li>Ȧ������:" & rs("teamname") & "</li>"
		showinfo=showinfo &"	<li>������:" & rs("username") & "</li>"
		showinfo=showinfo &"	<li>����ʱ��:" & rs("adddate") & "</li>"
		showinfo=showinfo &"	<li>��Ա����:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "</li>"
		showinfo=showinfo &"	<li>����ظ�:" & conn.execute("select count(*) from ks_teamtopic where parentid=0 and teamid=" & id )(0) & "/" & conn.execute("select count(*) from ks_teamtopic where parentid<>0 and teamid=" & id )(0) & "</li>"
		showinfo=showinfo &"</div></div>"
		showinfo=showinfo &"  <h1>Ȧ�ӹ���Ա</h1>"
		showinfo=showinfo &"<div id=""user_list"">"
		showinfo=showinfo &"  <ul><li class=""u1"">"
			dim rsu:set rsu=server.createobject("adodb.recordset")
			rsu.open "select top 1 * from ks_user where username='" & rs("username") &"'",conn,1,1
			if not rsu.eof then
			  Dim UserFaceSrc:UserFaceSrc=rsu("UserFace")
			  Dim FaceWidth:FaceWidth=KS.ChkClng(rsu("FaceWidth"))
			  Dim FaceHeight:FaceHeight=KS.ChkClng(rsu("FaceHeight"))
			  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.GetDomain & userfacesrc

		showinfo=showinfo &"  <img src=""" & UserFaceSrc & """ border=""1"" width=""" & facewidth & """ height=""" & faceheight & """></li>"
		showinfo=showinfo &"	<li class=""u2""><a href=""?" & rsu("username") & """ target=""_blank"">" & rs("username") & "</a></li>"
		showinfo=showinfo &"	<li class=""u3"">(" & rsu("province") & rsu("city") & ")</li>"
			end if
			rsu.close:set rsu=nothing
		showinfo=showinfo &"</ul>"
		showinfo=showinfo &"</div></div>"
		End Function
        
		Sub deltopic()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("�Բ������ȵ�¼��","")
		  exit sub
		 end if
		 dim tid:tid=ks.chkclng(ks.s("tid"))
		 if tid=0 then response.end
		 dim rst:set rst=server.createobject("adodb.recordset")
		 rst.open "select * from ks_teamtopic where id=" & tid,conn,1,3
		 if not rst.eof then
		     conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid=" & tid)
		  if rst("username")=KSUser.UserName or KSUser.UserName=groupadmin then
		     conn.execute("delete from ks_uploadfiles where channelid=1031 and infoid in(select id from ks_teamtopic where parentid=" & tid & ")")
		     conn.execute("delete from ks_teamtopic where parentid=" & tid)
			 rst.delete
		  else
		     rst.close:et rst=nothing
		    call ks.alert("�Բ�����û��ɾ����Ȩ��","")
		  end if
		 end if
		 rst.close:set rst=nothing
		 if ks.s("flag")="replay" then
		 response.write "<script>alert('ɾ���ɹ�');location.href='"& request.servervariables("http_referer") & "';</script>"
		 else
		 response.write "<script>alert('ɾ���ɹ�');location.href='group.asp?id="& id & "';</script>"
		 end if
		End Sub
		'�ö�����
		Sub Settop()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("�Բ������ȵ�¼��","")
		  exit sub
		 end if
		  dim tid:tid=KS.chkclng(KS.S("tid"))
		  dim rs:set rs=server.createobject("adodb.recordset")
		  rs.open "select top 1 istop from ks_teamtopic where id=" & tid,conn,1,3
		  if not rs.eof then
		   if rs(0)=1 then
			 rs(0)=0
		   else
			 rs(0)=1
		   end if
		   rs.update
		  end if
		  rs.close:set rs=nothing
		  response.redirect request.servervariables("http_referer")
		end sub
		'��������
		Sub Setbest()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("�Բ������ȵ�¼��","")
		  exit sub
		 end if
		  dim tid:tid=KS.chkclng(KS.S("tid"))
		  dim rs:set rs=server.createobject("adodb.recordset")
		  rs.open "select top 1 isbest from ks_teamtopic where id=" & tid,conn,1,3
		  if not rs.eof then
		   if rs(0)=1 then
			 rs(0)=0
		   else
			 rs(0)=1
		   end if
		   rs.update
		  end if
		  rs.close:set rs=nothing
		  response.redirect request.servervariables("http_referer")
		end sub
		Sub deluser()
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  call KS.Alert("�Բ������ȵ�¼��","")
		  exit sub
		 end if
		  if KSUser.UserName=groupadmin then
		     conn.execute("delete from ks_teamusers where teamid=" &id & " and username<>'" & ksuser.username & "' and username='" & KS.S("UserName") & "'")
		  else
		    call ks.alert("�Բ�����û�д˲�����Ȩ��","")
		  end if
		 response.write "<script>alert('�û��ѱ��ɹ��߳�!');location.href='group.asp?id="& id & "&action=users';</script>"
		End Sub
End Class
%>
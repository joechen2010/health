<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_Friend
KSCls.Kesion()
Set KSCls = Nothing

Class User_Friend
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,SQL,tablebody,strErr,action,boxName,smscount,smstype,readaction,turl
		Private ArticleStatus,ComeUrl,TotalPages
		Private Sub Class_Initialize()
			MaxPerPage =50
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		action=KS.S("Action")
		
		Call KSUser.Head()
		Call KSUser.InnerLocation("�ҵĺ���")
		KSUser.CheckPowerAndDie("s03")
		%>
		<div class="tabs">	
			<ul>
				<li<%If action="" then KS.Echo " class='select'"%>>
				<span class="rl" onMouseover="$('#f').show();"><a href="?">�ҵĺ���</a>
				 <div id="f" class='abs' onMouseOut="$('#f').hide();">
				 <dl><a href='?listtype=1'> �� �� �� </a>&nbsp;&nbsp;</dl>
				 <dl><a href='user_message.asp?action=friendrequest'> İ �� �� </a>&nbsp;&nbsp;</dl>
				 <dl><a href='?listtype=3'> �� �� �� </a>&nbsp;&nbsp;</dl>
				 </div>
				</span>
				</li>
				<li<%If action="addF" then KS.Echo " class='select'"%>><a href="?action=addF">Ѱ�Һ���</a></li>
				<li<%If action="online" then KS.Echo " class='select'"%>><a href="?action=online">�����û�</a></li>
				<li<%if action="add" then KS.Echo " class='select'"%>><a href="?Action=add">��Ӻ���</a></li>
				<li<%if action="mail" then KS.Echo " class='select'"%>><a href="?Action=mail">�ʼ��������</a></li>
			</ul>
	 </div>
	 
	 	<script src="../ks_inc/kesion.box.js"></script>
		<script type="text/javascript">
		 function checkmsg()
		 {
		     var message=escape($("#message").val());
			 var username=escape($("#username").val());
			 if (username==''){
			  alert('�������ݳ���!');
			  closeWindow();
			 }
			 if (message==''){
			   alert('��������Ϣ����!');
			   $("#message").focus();
			   return false;
			 }
			 $.get("../plus/ajaxs.asp",{action:"SendMsg",username:username,message:message},function(r){
			   r=unescape(r);
			   if (r!='success'){
				alert(r);
			   }else{
				 alert('��ϲ��������Ϣ�ѷ���!');
				 closeWindow();
			   }
			 });
         }
		 function sendMsg(ev,username)
		 {
		  mousepopup(ev,"<img src='../images/user/mail.gif' align='absmiddle'>������Ϣ","�Է���¼����Կ���������Ϣ(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(checkmsg())' value=' ȷ �� ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='button'></div>",350)
		 }
        function check()
		{
		
		 var message=escape($("#message").val());
		 var username=escape($("#username").val());
		 if (username==''){
		  alert('�������ݳ���!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('�����븽��!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("../plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     alert('���������ѷ���,��ȴ��Է���ȷ��!');
			 closeWindow();
		   }
		 });
		}
		function addF(ev,username)
		{ 
		 show(ev,username);
		 var isMyFriend=false;
		 $.get("../plus/ajaxs.asp",{action:"CheckMyFriend",username:escape(username)},function(b){
			  if (b=='true'){
			  closeWindow();
			  alert('�û�['+username+']�Ѿ������ĺ����ˣ�');
			  return false;
			 }else if(b=='verify'){
			  closeWindow();
			  alert('���������['+username+'],��ȴ��Է�����֤!');
			  return false;
			 }else{
			   show(ev,username);
			 }
		 });
		 
		}
		function show(ev,username)
		{
		 mousepopup(ev,"<img src='../images/user/log/106.gif'>��Ӻ���","ͨ���Է���֤���ܳ�Ϊ����(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(check())' value=' ȷ �� ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='button'></div>",350);
		}		 
		</script>
		
		<%
		action=lcase(Trim(request("action")))
		CurrentPage=Trim(request("page"))
		if Isnumeric(CurrentPage) then
			CurrentPage=Clng(CurrentPage)
		else
			CurrentPage=1
		end if
		select case action
		case "add"
		    Call AddFriend()
			Call KSUser.InnerLocation("��Ӻ���")
	    case "edit"
		    Call AddFriend()
			Call KSUser.InnerLocation("�޸ĺ�������")
		case "addsave"
		    call addsave()
		case "del"
		    call DelFriend()
		case "info"
			call info()
			Call KSUser.InnerLocation("�ҵĺ���")
		case "addf"
			call addF()
			Call KSUser.InnerLocation("��Ӻ���")
		case "savef" call saveF()
		case "move" call moveF()
		case "shielddt" call ShieldDT()
		case "mail" call mail()
		case "mailsave" call mailsave()
		case else
			call info()
		end select
		  	%>
		</TD>    
		 </TR>
</TABLE>
		</div> <%
	  End Sub
	   
		'��Ӻ���
		Sub AddFriend()
		  dim flag,username,realname,phone,mobile,qq,msn,email,note
		  dim id:id=KS.chkclng(KS.S("id"))
		 if KS.S("action")="edit" then
		   dim rs:set rs=server.createobject("adodb.recordset")
		   rs.open "select * from ks_friend where id=" & id,conn,1,1
		   if rs.eof and rs.bof then
		    rs.close:set rs=nothing
		    call KS.alerthistory("�������ݳ���!",-1):exit sub
		   else
		     username=rs("friend")
			 flag=rs("flag")
			 realname=rs("realname")
			 phone=rs("phone")
			 mobile=rs("mobile")
			 qq=rs("qq")
			 msn=rs("msn")
			 email=rs("email")
			 note=rs("note")
		   end if
		   rs.close:set rs=nothing
		 else
		  flag=KS.S("flag")
		 end if%>
		<script>
		 function checkform()
		 {
		   if (document.myform.username.value=='')
		   {
		     alert('��������ѵ��û���!');
			 document.myform.username.focus();
			 return false;
		   }
		   
		  var message=escape($("#message").val());
		 var username=escape($("#username").val());
		 if (username==''){
		  alert('�������ݳ���!');
		 }
		 if (message==''){
		   alert('�����븽��!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("../plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     if (confirm('���������ѷ���,��ȴ��Է���ȷ��,ȷ�������')){
			 $("#username").val('');
			 $("#message").val('');
			 }else{
			  location.href='user_friend.asp';
			 }
		   }
		 });
		   return true;
		 }
		 
		 
		</script>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1"  class="border">
		<form action="?action=addsave" method="post" name="myform">
			<tr height="23" class="title">
				<td width="25%" height="25" align="center">��Ӻ���<span style='color:#fff;font-weight:normal'>(ͨ���Է���֤���ܳ�Ϊ����)</span></td>
			</tr>
			<tr> 
			  <td>
			    <table border="0" cellpadding="0" cellspacing="1" width="100%">
				 <tr class='tdbg'>
				  <td width="344"><b>�û���</b><br>
			       ��¼��Ա���ĵ��û�����������д��</td>
				  <td width="607"><input type="text" class="textbox" name="username" id="username" value="<%=username%>" size=20 style="width:150"> <font color=red>*</font></td>
				 </tr>
				 <tr class='tdbg'>
				  <td><b>�� ��</b><br>ֻ��ͨ���Է���֤���ܳ�Ϊ���ѡ�</td>
				  <td><textarea class="textbox" name="message" id="message" style="width:350px;height:80px"></textarea></td>
				 </tr>
				
				 <tr class='tdbg'>
				  <td height="30" colspan=2 align="center">
				  <input type="hidden" name="id" value="<%=id%>">
				  <input type="button" onClick="checkform()" class="button" value=" OK,���� ">&nbsp;<input type="button" class="button" value=" ȡ �� " onClick="javascript:history.back()"></td>
				 </tr>
				 </table>
			  </td>
			</tr>
			</form>
		</table>
		<%
		End Sub
		
		'�ʼ��������
		Sub Mail()
		if KS.Setting(143)<>"1" Then
		  Call KS.ShowError("�Բ���,��վû�п����ʼ����빦��!")
		  Exit Sub
		End If
		Call KSUser.InnerLocation("�ʼ��������")
		 %>
		<script>
		 function checkform()
		 {
		   if (document.myform.realname.value=='')
		   {
		     alert('��������������!');
			 document.myform.realname.focus();
			 return false;
		   }
		   if (document.myform.email.value=='')
		   {
		     alert('���������ĺ�������!');
			 document.myform.email.focus();
			 return false;
		   }
		   return true;
		 }
		</script>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1"  class="border">
		<form action="?action=mailsave" method="post" name="myform">
			<tr height="23" class="title">
				<td width="25%" height="25" align="center">ͨ���ʼ��������</span></td>
			</tr>
			<tr> 
			  <td>
			    <table border="0" cellpadding="0" cellspacing="1" width="100%">
				 <tr class='tdbg'>
				  <td width="344"><b>��������</b><br>
			       ��ʾ�ڶԷ�����ķ������С�</td>
				  <td width="607"><input type="text" class="textbox" name="realname" id="realname" value="<%=ksuser.realname%>" size=20 style="width:150"> <font color=red>*</font></td>
				 </tr>
				 <tr class='tdbg'>
				  <td><b>���ѵ�����</b><br></td>
				  <td><textarea class="textbox" name="email" id="email" style="width:350px;height:80px"></textarea>
				  <br><font color=#999999>��д���Emailʱ�����û��зָ�(һ��һ��,����һ��������20�������ַ)��</font> 
				  </td>
				 </tr>
				
				 <tr class='tdbg'>
				  <td height="30" colspan=2 align="center">
				  <input type="submit" onClick="checkform()" class="button" value=" OK,���� ">&nbsp;<input type="button" class="button" value=" ȡ �� " onClick="javascript:history.back()"></td>
				 </tr>
				 </table>
			  </td>
			</tr>
			</form>
		</table>
		<br/>
		<%if KS.Setting(144)>0 then%>
		<div align="center"><font color=green>����˵�����ɹ��Ƽ�һ������ע��,������������ <font color=red><%=KS.Setting(144)%></font> �����֡��Ͽ��ж��ɣ�</font></div>
		<%
		end if
		End Sub
		
		Sub mailsave()
		 dim realname:realname=ks.s("realname")
		 dim email:email=replace(request("email"),"'","")
		 if realname="" then call ks.alerthistory("��������������!",-1)
		 if email="" then call ks.alerthistory("�������������!",-1)
		 email=split(email,vbcrlf)
		 If ubound(email)>20 Then call KS.AlertHistory("һ�����ֻ�ܷ��͸�20������!",-1)
		 dim i,content,ReturnInfo,N
		 for i=0 to ubound(email)
		   dim user_face:user_face=ksuser.userface
		   If lcase(left(user_face,4))<>"http" then user_face=KS.GetDomain & user_face
			content="<style type=""text/css"">A:visited {	TEXT-DECORATION: none	}"
			content=content &"A:active  {	TEXT-DECORATION: none	}"
			content=content &"A:hover   {	TEXT-DECORATION: underline overline	}"
			content=content &"A:link 	  {	text-decoration: none;}"
			content=content &"A:visited {	text-decoration: none;}"
			content=content &"A:active  {	TEXT-DECORATION: none;}"
			content=content &"A:hover   {	TEXT-DECORATION: underline overline}"
			content=content &"body   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
			content=content &"td  {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"

		   content=content & "<table border=""0"" width=""100%"">"
		   content=content & "<tr><td width=""150"" align=""center""><a style=""border:1px solid #ccc;padding:1px"" href=""" & KS.GetDomain & "space?" & KSUser.UserName & """ target=""_blank""><img width=""82"" height=""82"" src=""" & user_face & """ border=""0"" /></a><br/>" & KSUser.RealName & "</td>"
		   content=content & "<td><p>Hi,���������Щʲô�,�ҷ���һ������" & KS.Setting(1) & "����վ,�����кܶ����Ķ���,�������ҷ����˺ö���ʶ�������Ѻ͸���ʶ��������,���Ƕ��й�ͬ����Ȥ����,��Ҳһ������!</p><p>���������ӣ����Ҽ�Ϊ���Ѱɣ�<br /><a href=""" & KS.GetDomain & "user/reg?uid=" & server.urlencode(KSUser.UserName) & "&f=r&u=" & server.urlencode(email(i)) &""" target=""_blank"">" & KS.GetDomain & "user/reg?uid=" & server.urlencode(KSUser.UserName) & "&f=r&u=" & server.urlencode(email(i)) &"</a></p>"
		   content=content & "ע�������ʼ������ĺ���<a href=""" & KS.GetDomain & "space?" & KSUser.UserName & """ target=""_blank"">" & KSUser.RealName & "</a>ͨ��" & KS.Setting(1) & "���ͣ�����ظ���"
		   content=content &"</td></tr></table>"
		   
		   ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), realname & "����������" & KS.Setting(0), Email(i),realname, content,KS.Setting(11))
		   IF ReturnInfo="OK" Then
			 N=N+1
		    End If
		 next
		 response.write "<script>alert('�ɹ������� " & N & " λ���ѵ�����!');location.href='" & Request.ServerVariables("http_referer") & "';</script>"
		End Sub
		
		
		sub info()
		 
		%>
		<style type="text/css">
		 .friendlist{font-size:12px}
		 .friendlist .t a{font-size:14px;color:#ff6600;height:30px;line-height:30px}
		 .friendlist li{float:left;width:350px;margin:5px;height:100px;background:#FAFAFA}
		 .friendlist li div{height:22px;line-height:22px}
		 .friendlist .l{height:100px;width:70px;float:left;text-align:center;padding-top:10px;padding-left:4px}
		 .friendlist .l a{border:1px solid #ccc;padding:1px;display:block}
		 .rriendlist .r{float:left}
		 .friendlist .r .zl span{color:#999999}
		</style>
		<form action="User_Friend.asp?action=move" method=post name=inbox>
		<div class="friendlist">
		<%  dim param,i
		    if KS.chkclng(KS.S("listtype"))<>0 then param=param & " and flag=" & KS.chkclng(KS.S("listtype"))
			set rs=server.createobject("adodb.recordset")
			MaxPerPage=10
			if action="online" then
			sql="select userid as id,RealName,Username,userface,sex,birthday,province,city,isonline from KS_user where isonline=1 order by userid desc"
			else
			sql="select f.id,U.RealName,U.Username,u.userface,u.sex,u.birthday,u.province,u.city,u.isonline,f.flag,f.ShieldDT from KS_Friend F inner join KS_User U on F.Friend=U.UserName where F.accepted=1 and F.Username='"&KSUser.UserName&"' " & param & " order by F.addtime desc"
			end if
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		         select case KS.S("listtype")
						    case "2"
							 response.write "��û�����İ���ˡ�"
							case "3"
							 response.write "��û����Ӻ�������"
							case else
							 response.write "��û����Ӻ����ѡ�"
						   end select
				else
		         totalPut = RS.RecordCount
				 If CurrentPage < 1 Then	CurrentPage = 1
				 If (CurrentPage - 1) * MaxPerPage > totalPut Then
							If (totalPut Mod MaxPerPage) = 0 Then
								CurrentPage = totalPut \ MaxPerPage
							Else
								CurrentPage = totalPut \ MaxPerPage + 1
							End If
				End If
			    If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
				Else
							CurrentPage = 1
				End If
		     i=0					       
			do while not rs.eof
			%>
					<li>
						<div class='l'>
								<%
								dim user_face:user_face=rs("userface")
								If user_face="" or isnull(user_face) then 
								if rs("sex")="��" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
							    End If
								If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
								response.write "<a href='../space/?" & rs("username") & "' target='_blank'><img src='" & user_face & "' width='60' height='60'/></a>"
								
								  if action="" then
								   response.write "<div><label><input type='checkbox' name='id' value='" & rs("id") & "'>ѡ��</label></div>"
								  end if
								 %>
								</div>
								<div class='r'>
								 <div class='t'>
								 <a href="../space/?<%=rs("username")%>" target="_blank"><%=rs("username")%></a>(<%=rs("realname")%>)
								 
								 </div>
								 
								 <div><a href="javascript:void(0)" onClick="sendMsg(event,'<%=rs("username")%>')">������Ϣ</a>
								 <%If action="online" then%>
								 <img src="../images/user/log/106.gif" border="0"><a href="javascript:void(0)" onClick="addF(event,'<%=rs("username")%>')">��Ϊ����</a>
								 <%else%>
								 <a href="?action=del&id=<%=rs(0)%>" onClick="return(confirm('ȷ�����λ���ѽ����ϵ��'))">�����ϵ</a>
								 <%if rs("ShieldDT")="0" then
								  response.write "<a href='?action=ShieldDT&v=1&id=" & rs("id") & "'>���ζ�̬</a>"
								  else
								  response.write "<a href='?action=ShieldDT&v=0&id=" & rs("id") & "'>��ʾ��̬</a>"
								  end if
								 
								 end if%>
								 </div>
								 <div class="zl">�Ա�<span><%=rs("sex")%></span>&nbsp; <%If rs("birthday")<>"" and not isnull(rs("birthday")) then response.write "���� <span>" & split(rs("birthday")," ")(0) & "</span> ���ԣ�<span>" & rs("province") & rs("city") & "</span>"
								 response.write "<br/>"
								 if action="" then
								  response.write "��ϵ: "
								  if rs("flag")="1" then response.write "������ " else response.write "������ "
								 end if
								 response.write "״̬��<span>"
								 if rs("isonline")="1" then response.write "<font color=red>����</font>" else response.write "����"
								 response.write "</span>"
								 %></div>
								</div>
								 
				</li>			
							
			<%
				rs.movenext
			  i=i+1 : if i>=maxperpage then exit do
		  loop
		end if
		rs.close
		set rs=Nothing
		%>
			
				
     </div>
	 
	 <%if action="" then%>
	 <div class="clear">
		  <label><input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ������</label> ��ѡ�еĺ����ƶ���
		  <select name="grouptype" id="grouptype">
		   <option value="1">������</option>
		   <option value="3">������</option>
		  </select>
		  <input type="submit" value=" ȷ �� " class="button">
	</div>	
		
	 <br/>
	 <strong>��ʾ:</strong>
	 ����������յ�ĳ���û�����Ϣ,�����Խ���(��)�Ƶ�������
	 <%end if%>
		  </form>
	 <div class="clear" style="text-align:right;margin-right:50px">
	 <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
	 </div>
	 
	 <br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
	 <br/><br/>
		<%
		end sub
		
		sub moveF()
		 Dim ID:ID=Request("id")
		 Dim flag:flag=KS.ChkClng(request("grouptype"))
		 If Id="" Then KS.AlertHintScript "��ѡ��Ҫ�Ƶ��ĺ���!"
		 ID=KS.FilterIDS(ID)
		 If Flag=0 Then Flag=1
		 Conn.Execute("Update KS_Friend Set Flag=" & Flag & " where id in(" & ID & ")")
		 Response.Redirect Request.ServerVariables("http_referer")
		end sub
		
		sub ShieldDT()
		 Conn.Execute("Update KS_Friend Set ShieldDT=" & KS.ChkClng(request("v")) & " where id=" & KS.ChkClng(Request("ID")))
		 Response.Redirect Request.ServerVariables("http_referer")
		end sub
		
		sub delFriend()
		dim delid
		delid=replace(request("id"),"'","")
		DelID=KS.FilterIDs(DelID)
		if delid="" or isnull(delid) then
		    Call KSUser.InnerLocation("������ʾ")
			Call KS.AlertHistory("��û��ѡ��Ҫɾ������������",-1)
			exit sub
		else
		    dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select friend from KS_Friend where username='"&KSUser.UserName&"' and id in ("&delid&")",conn,1,3
			do while not rs.eof
			 Call KSUser.AddLog(KSUser.UserName,"��<a href=""{$GetSiteUrl}space/?" & rs(0) & """ target=""_blank"">" & rs(0) & "</a>������ѹ�ϵ!",106)
			 rs.delete
			 rs.movenext
			loop
			rs.close
			set rs=nothing
			Call KSUser.InnerLocation("�ɹ���ʾ")
			Call KS.Alert("���Ѿ���ѡ���ĺ��ѽ����ϵ��","User_Friend.asp")
		end if
		end sub
		
		sub AllDelFriend()
			Conn.Execute("delete from KS_Friend where username='"&KSUser.UserName&"'")
			Call KSUser.InnerLocation("�ɹ���ʾ")
			Call KS.Alert("���Ѿ�ɾ�������к����б�","User_Friend.asp")
		end sub
		
		sub addF()
		call userlist()
		Response.write "<div align=center style=""margin-top:5px"">"
		Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
		Response.write "</div>"
		%>
		<div style="clear:both"></div>
		<br>

		<table border="0" width="98%" align=center cellpadding=0 cellspacing=1 class="border">
			<form action="../space/friend/" method="get" name="myform" target="_blank">
				  <tr class="title"> 
					<td height="25" colspan="2" align="left">
				    �Һ���</td>
				  </tr>
				  <tr height="30" class="tdbg"> 
					<td><b>��Ҫ�ң�</b>
					  <script src="../plus/area.asp" type="text/javascript"></script>
					</td>
				 </tr>
				 <tr>
				  <td height="30"> <strong>��&nbsp;&nbsp;�� </strong>
						<Select id="sex" name="sex"> 
						  <Option value="" selected>-����</Option> 
						  <Option value=��>����</Option> 
						  <Option value=Ů>Ů��</Option>
						</Select>				
			      </td>
				  </tr>
				   <tr>
            <td width="371" height="30"><strong>��&nbsp;&nbsp;����</strong>
<Select id="birth_y" name="birth_y" style="width:50px"> 
  <Option value="" selected>��</Option>
  <%dim n
   for n=1950 to year(now)-5
    response.write "<option value=" & n & ">" & n & "</option>"
   next
  %>
</Select> 
<Select id="birth_m" name="birth_m" style="width:50px"> 
  <Option value="" selected>��</Option>
    <%
   for n=1 to 12
    response.write "<option value=" & n & ">" & n & "</option>"
   next
  %>
</Select> 
<Select id="birth_d" name="birth_d" style="width:50px"> 
  <Option value="" selected>��</Option>
    <%
   for n=1 to 31
    response.write "<option value=" & n & ">" & n & "</option>"
   next
  %>
</Select> ����
            <Input id="realname" size="12" name="realname"> </td>
				  
				  <tr class="tdbg"> 
					<td colspan=2 height="50" align=center valign=middle> 
					  <input class="Button" type=Submit value=" �� �� �� " name=Submit>
					  &nbsp; 
					  <input class="Button" type="reset" name="Clear" value=" �� �� ">
					</td>
				  </tr>
		  </form>
</table><br>
		<%
		end sub
		
		sub saveF()
		dim incept,i
		if request("touser")="" then
		    Call KSUser.InnerLocation("������ʾ")
			Call KS.AlertHistory("����д����",-1)
			exit sub
		else
			incept=KS.R(request("touser"))
			incept=split(incept,",")
		end if
		
		for i=0 to ubound(incept)
		set rs=server.createobject("adodb.recordset")
		sql="select UserName from KS_User where UserName='"&incept(i)&"'"
		set rs=Conn.Execute(sql)
		if rs.eof and rs.bof then
		    Call KSUser.InnerLocation("������ʾ")
			Call KS.ShowError("ϵͳû�У�"&incept(i)&"������û�������δ�ɹ���")
			exit sub
		end if
		set rs=Nothing
		
		if KSUser.UserName=Trim(incept(i)) then
		    Call KSUser.InnerLocation("������ʾ")
			Call KS.ShowError("���ܰ��������Ϊ���ѡ�")
		end if
		
		sql="select friend from KS_Friend where username='"&KSUser.UserName&"' and  friend='"&incept(i)&"'"
		set rs=Conn.Execute(sql)
		if rs.eof and rs.bof then
			sql="insert into KS_Friend (username,friend,addtime,flag) values ('"&KSUser.UserName&"','"&Trim(incept(i))&"',"&SqlNowString&",1)"
			set rs=Conn.Execute(sql)
		end if
		'if i>5 then
		'	Call KS.ShowError("ÿ�����ֻ�����5λ�û�����������5λ�Ժ����������д��")
		'	exit sub
		'	exit for
		'end if
		next
		Call KSUser.InnerLocation("�ɹ���Ϣ")
		Call KS.Alert("��ϲ����������ӳɹ���","User_Friend.asp")
		end sub
		
		sub userlist()
		
		Response.Write "<table width=""98%"" class='border' align=center cellpadding=2 cellspacing=1 border=0><tr class='tdbg'>"
	
		MaxPerPage=12
		Response.Write "<table width=""98%"" class='border' align=center cellpadding=2 cellspacing=1 border=0><tr class=title><td height=""30"">&nbsp;�����Ƽ�<span style='color:#ffffff;font-weight:normal'>(�������ڱ�վ���Ծ��Ա)</span></td></tr></table>"
		Response.Write "<table width=""98%"" class='border' align=center cellpadding=2 cellspacing=1 border=0><tr class='tdbg'>"
		dim user_face,user_info,sex,i,n
		sql="select top 500 UserName,Sex,qq,Email,userface,realname,province,city,isonline from KS_User where GroupID<>1 and username<>'" & KSUser.UserName & "' order by Logintimes desc,UserID"
		set rs=Server.CreateObject("adodb.recordSet")
		rs.Open sql,Conn,1,1
		i=0:n=0:TotalPut=0
		if not (rs.bof and rs.eof) then
			TotalPut=rs.recordcount
			if (TotalPut mod MaxPerPage)=0 then
				TotalPages = TotalPut \ MaxPerPage
			else
				TotalPages = TotalPut \ MaxPerPage + 1
			end if
			if CurrentPage > TotalPages then CurrentPage=TotalPages
			if CurrentPage < 1 then CurrentPage=1
			rs.move (CurrentPage-1)*MaxPerPage
			do while not rs.eof
			user_info="������"& rs("realname") & "�Ա�"& rs("sex") & vbcrlf & "Q&nbsp;&nbsp;Q��"& rs("qq") & vbcrlf &"Email��"& rs(3)
			user_face=rs("userface")
								If user_face="" or isnull(user_face) then 
								if rs("sex")="��" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
							    End If
			If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
			
				Response.Write "<td class=""splittd"" style=""padding:3px"" height=20 align=""center"" width=""14%"">"
				response.write " <table border='0' width='100%'>"
				response.write " <tr><td width='65' align\'center'>"
				Response.Write "<img width=""60"" height=""60"" src=""" & user_face&"""></td>"
				response.write "<td>�û�:<a target='_blank' href='../space/?"&rs(0)&"' title="""& user_info &""">"&rs(0)&"</a>"
				'response.write "<br/>�Ա�:" & rs("sex") &" " & rs("province") & rs("city")
				response.write "<br/>״̬:"
				if rs("isonline")="1" then response.write "<font color=red>����</font>" else response.write "����"
				response.write "<br/><a onclick=""addF(event,'" & rs(0) & "');return false"" href=""javascript:void(0)"" title="""& user_info &"""><img border='0' src='../images/user/log/106.gif'>��Ϊ����</a> <img src='../images/user/mail.gif'><a href='javascript:void(0)' onclick=""sendMsg(event,'" & rs(0) & "');"">������Ϣ</a>"
				response.write "</td>"
				response.write " </tr></table>"
				response.write "</td>"
			
			i=i+1
			if i>=3 then 
				if i=3 then Response.Write "</tr><tr>"
				i=0
			end if
			n=n+1
			if n>= MaxPerPage then Exit Do
			rs.movenext
			loop
		else
		    Call KSUser.InnerLocation("ϵͳ��ʾ")
			Response.Write "<td>���κ��û�</td>"
		end if
		Response.Write "</tr></TABLE><br>"
		set rs=Nothing
		end sub
End Class
%> 

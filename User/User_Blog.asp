<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_Blog
KSCls.Kesion()
Set KSCls = Nothing

Class User_Blog
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather
		Private TypeID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
		Private Sub Class_Initialize()
		  MaxPerPage =15
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
		If KS.SSetting(0)=0 Then
		 Call KS.Alert("�Բ��𣬱�վ�رո��˿ռ书�ܣ�","")
		 Exit Sub
		End If
		Call KSUser.Head()
		 Action=KS.S("Action")
		 
		 If Action="" Or Action="Add" Then
		 KSUser.CheckPowerAndDie("s02")
		 End If
		 
		%>
		<div class="tabs">	
			<ul>
			 <%IF Action="BlogEdit" Or Action="Template" Or action="Banner" Then%>
			 <li<%If Action="BlogEdit" then response.write " class='select'"%>><a href="?action=BlogEdit">�ռ�����</a></li>
			 <%If KSUser.UserType=1 Then%>
			 <li<%If Action="Banner" then response.write " class='select'"%>><a href="?action=Banner">Banner����</a></li>
			 <%End IF%>
			 <li<%If Action="Template" then response.write " class='select'"%>><a href="?action=Template">ģ������</a></li>
			 <%End If%>
				 
			 <%
			 If Action="Add" Or Action="Edit" Then
			 %>
			 <li><a href="?">��־����</a></li>
			 <li class='select'><%If Action="Add" Then Response.Write "д��־" Else Response.Write "�༭��־" End If%></li>
			 <%
			 Elseif Action="" then%>
				<li<%If KS.ChkClng(KS.S("Status"))="0" then response.write " class='select'"%>><a href="?Status=0">�����(<span class="red"><%=conn.execute("select count(id) from KS_BlogInfo where Status=0 and UserName='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">�����(<span class="red"><%=conn.execute("select count(id) from KS_BlogInfo where Status=2 and UserName='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">�� ��(<span class="red"><%=conn.execute("select count(id) from KS_BlogInfo where Status=1 and UserName='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			 <%end if%>

			</ul>
	  </div>
					<%if ks.s("action")="" or ks.s("action")="Comment" then%>
					 <div style="margin:10px;padding-left:20px;"><img src="../images/user/log/101.gif" align="absmiddle"><a href="User_Blog.asp?Action=Add"><span style="font-size:14px;color:#ff3300">д��־</span></a> 
					 &nbsp;&nbsp;<img src="../images/user/log/100.gif" align="absmiddle"><a href="User_message.asp?Action=Comment"><span style="font-size:14px;color:#ff3300">��־����</span></a>
					 </div>
					<%end if%>


		<%
		If KS.S("Action")="ApplySave" Then
		   Call ApplyBlogSave()
		ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		    Response.Write "<script>alert('����û�п�ͨ���˿ռ�,��ȷ��ת��ͨҳ�棡');</script>"
		    Call ApplyBlog()
		ElseIf Conn.Execute("Select status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>alert('�Բ�����Ŀռ仹û��ͨ����˻�������');location.href='user_main.asp';</script>"
			response.end
		Else
			Select Case KS.S("Action")
			 Case "Del"
			  Call ArticleDel()
			 Case "Add","Edit"
			  Call ArticleAdd()
			 Case "DoSave"
			  Call DoSave()
			 Case "Template"
			  Call Template()
			 Case "SaveMySkin"
			  Call SaveMySkin()
			 Case "BlogEdit"
			  Call ApplyBlog()
			 Case "UpTemplate"
			  Call UpTemplate()
			 Case "UpTemplateSave"
			 if KSUser.UserType=1 Then
			  Call UpTemplateSave()
			 End If
			 Case "DelTemplate"
			  Call DelTemplate()
			 Case "Banner" SetBanner()
			 Case Else
			  Call BlogList()
			End Select
		End If
		 Response.Write "</div>"
	   End Sub
	    '������־
	   Sub ApplyBlog()
	    Dim BlogName,domain,ClassID,Descript,ContentLen,ListBlogNum,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,Announce,ListLogNum,Logo
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Blog Where UserName='" & KSUser.UserName &"'",conn,1,1
		If Not RS.EOF Then
		Call KSUser.InnerLocation("�޸Ŀռ����")
		 BlogName=RS("BlogName")
		 Logo=RS("Logo")
		 domain=RS("domain")
		 ClassID=RS("ClassID")
		 Descript=RS("Descript")
		 Announce=RS("Announce")
		 ContentLen=RS("ContentLen")
		 ListBlogNum=RS("ListBlogNum")
		 ListLogNum=RS("ListLogNum")
		 ListReplayNum=RS("ListReplayNum")
		 ListGuestNum=RS("ListGuestNum")
		 OpStr="OK�ˣ�ȷ���޸�"
		Else
		Call KSUser.InnerLocation("���뿪ͨ���˿ռ�")
		 BlogName=KSUser.UserName & "�ĸ��˿ռ�"
		 domain=KSUser.UserName
		 ClassID="0"
		 ContentLen=500
		 ListBlogNum=10
		 ListLogNum=10
		 ListReplayNum=10
		 ListGuestNum=10
		 Announce="û�й���!"
		 Logo="../Images/logo.jpg"
		 OpStr="OK�ˣ���������":TipStr="�� �� �� ͨ �� �� �� ��"
		End if
		If Logo="" Or IsNull(Logo) Then Logo="../images/logo.jpg"
		RS.Close:Set RS=Nothing
	    %>
		<script>
		 function CheckForm()
		 {
		  if (document.myform.BlogName.value=='')
		  {
		   alert('���������վ������!');
		   document.myform.BlogName.focus();
		   return false;
		  }
		  if (document.myform.ClassID.value=='0')
		  {
		   alert('��ѡ�����վ������!');
		   document.myform.ClassID.focus();
		   return false;
		  }
		  return true;
		 }
		</script>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
          <form  action="User_Blog.asp?Action=ApplySave" method="post" name="myform" id="myform" onSubmit="return CheckForm();" enctype="multipart/form-data">

            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>�ռ����ƣ�</strong><br>
              �ռ�վ������ơ�
              </div></td>
              <td> ��
                  <input class="textbox" name="BlogName" type="text" id="BlogName" style="width:250px; " value="<%=BlogName%>" maxlength="100" />
              <span style="color: #FF0000">*</span></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>�ռ�������</strong><br>
              ������ռ�ļ����ַ��</td>
              <td> ��
                  <input class="textbox" name="domain" type="text" id="domain" style="width:50px; " value="<%=domain%>" maxlength="100" /><b>.<%response.write KS.SSetting(16)%></b></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>�ռ�Logo��ַ��</strong><br>
              </div></td>
              <td>��
<input type="file" name="photourl" size="40">
                <img src="<%=logo%>" width="88" height="31"><br>
		  ��
                  
              ֻ֧��jpg��gif��png��С��100k��Ĭ�ϳߴ�Ϊ88*31</td>
            </tr>
            <tr class="tdbg">
              <td width="24%"  height="25" align="center"><div align="left"><strong>�ռ���ࣺ</strong><br>
      �ռ�վ����࣬�Ա��οͲ���</div></td>
              <td width="76%">��
                  <select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-��ѡ�����-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select>               </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>վ��������</strong></span></div><br>�������Ŀռ�վ�����</td>
              <td>��
                  
                  <textarea class="textbox" name="Descript" id="Descript" style="width:80%;height:60px" cols=50 rows=6><%=Descript%></textarea>              </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>�ռ乫�棺</strong></span></div>
                <br>                �����������»���棬�ø����û��˽�����</td>
              <td>��
                  
                  <textarea class="textbox" name="Announce" id="Announce" style="width:80%;height:80px" cols=50 rows=6><%=Announce%></textarea>              </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>��־Ĭ�ϲ�����ʾ������</strong><br>
                  ��ҳ��־�б�ҳ����ʾ�����������ó�0����ʾȫ�ġ�</div></td>
              <td>��
                  <input class="textbox" name="ContentLen" type="text" id="ContentLen" style="width:250px; " value="<%=ContentLen%>" />             </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>ÿҳ��ʾ��־ƪ����</strong><br>�ռ���ҳ��ʾ��־������</div></td>
              <td>��
                  <input class="textbox" name="ListBlogNum" type="text" id="ListBlogNum" style="width:250px; " value="<%=ListBlogNum%>" />              </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>��ʾ���»ظ�������</strong><br>�ռ���ҳ��ʾ���»ظ�������</div></td>
              <td>��
                  <input class="textbox" name="ListReplayNum" type="text" id="ListReplayNum" style="width:250px; " value="<%=ListReplayNum%>" />                </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>��ʾ������־ƪ����</strong><br>�ռ���ҳ��ʾ������־ƪ����</div></td>
              <td>��
                  <input class="textbox" name="ListLogNum" type="text" id="ListLogNum" style="width:250px; " value="<%=ListLogNum%>" />                </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>��ʾ�������������� </strong><br>�ռ���ҳ��ʾ��������������</div></td>
              <td>��
                  <input class="textbox" name="ListGuestNum" type="text" id="ListGuestNum" style="width:250px; " value="<%=ListGuestNum%>" />              </td>
            </tr>

            <tr class="tdbg">
              <td height="30" align="center" colspan=2>
                <input type="submit" name="Submit3"  class="Button" value="<%=OpStr%>" />
                <input type="reset" name="Submit22"   class="Button" value=" �� �� " />              </td>
            </tr>
          </form>
</table>
		<%
	   End Sub
	   
	   '������˿ռ�����
	   Sub ApplyBlogSave()
            Dim fobj:Set FObj = New UpFileClass
		    FObj.GetData
            Dim MaxFileSize:MaxFileSize = 100   '�趨�ļ��ϴ�����ֽ���
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.UserName)
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"logo")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����������ϴ��Ĵ�С\n�����ϴ� " & MaxFileSize & " KB���ļ�\n",-1):response.End()
			End Select

	     Dim BlogName:BlogName=KS.DelSql(Fobj.Form("BlogName"))
		 Dim Domain:Domain=KS.DelSql(Fobj.Form("Domain"))
		 Dim ClassID:ClassID=KS.ChkClng(Fobj.Form("ClassID"))
		 Dim Descript:Descript=KS.DelSql(Fobj.Form("Descript"))
		 Dim Announce:Announce=KS.DelSql(Fobj.Form("Announce"))
		 Dim ContentLen:ContentLen=KS.ChkClng(Fobj.Form("ContentLen"))
		 Dim ListBlogNum:ListBlogNum=KS.ChkClng(Fobj.Form("ListBlogNum"))
		 Dim ListLogNum:ListLogNum=KS.ChkClng(Fobj.Form("ListLogNum"))
		 Dim ListReplayNum:ListReplayNum=KS.ChkClng(Fobj.Form("ListReplayNum"))
		 Dim ListGuestNum:ListGuestNum=KS.ChkClng(Fobj.Form("ListGuestNum"))
		 If BlogName="" Then Response.Write "<script>alert('������վ������!');history.back();</script>":exit sub
		 If ClassID=0 Then Response.Write "<script>alert('��ѡ��վ������!');history.back();</script>":exit sub
		 
		 if lcase(domain)="www" or lcase(domain)="space" or lcase(domain)="bbs" or lcase(domain)="news" then call KS.AlertHistory("������Ķ�������Ϊϵͳ�����ؼ���,����������",-1)
		 if domain<>"" then
		  if not conn.execute("select username from ks_Blog where username<>'" & ksuser.username & "' and [domain]='" & domain  &"'").eof then
		  Response.Write "<script>alert('�Բ�����ע��Ķ��������ѱ������û�ʹ��!');history.back();</script>":exit sub
		  end if
		 end if
		 
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Blog Where UserName='" & KSUser.UserName & "'",conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now
			RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  if KS.SSetting(2)=1 then
			  RS("Status")=0
			  else
			  RS("Status")=1
			  end if
		 End If
		    If ReturnValue<>"" Then RS("Logo")=ReturnValue
		    RS("Domain")=Domain
		    RS("UserName")=KSUser.UserName
		    RS("BlogName")=BlogName
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Announce")=Announce
			RS("ContentLen")=ContentLen
			RS("ListLogNum")=ListLogNum
			RS("ListBlogNum")=ListBlogNum
			RS("ListReplayNum")=ListReplayNum
			RS("ListGuestNum")=ListGuestNum
		  RS.Update
		  RS.MoveLast
		  If Not KS.IsNul(RS("Logo")) or Not KS.IsNul(RS("Banner")) Then
		  Call KS.FileAssociation(1025,rs("BlogID"),RS("Logo")&RS("Banner"),1)
		  End If
		  
		 RS.Close:Set RS=Nothing
		 Set Fobj=Nothing
		 Call KSUser.AddLog(KSUser.UserName,"�޸��˿ռ��������!",102)
		 Response.Write "<script>alert('�ռ�վ������/�޸ĳɹ�!');location.href='User_Blog.asp?Action=BlogEdit';</script>"
	   End Sub
	   
	   Sub SetBanner()
		Call KSUser.InnerLocation("���ÿռ�Banner")
	   Dim banner
	   
	   If KS.S("Act")="Save" Then
	      Dim fobj:Set FObj = New UpFileClass
			 on error resume next
			 FObj.GetData
			 if err.number<>0 then
			  call KS.AlertHistory("�Բ���,�ļ����������ϴ��Ĵ�С!",-1)
			  response.end
			 end if
            Dim MaxFileSize:MaxFileSize = 600   '�趨�ļ��ϴ�����ֽ���
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.UserName)
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"banner")
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����������ϴ��Ĵ�С\n�����ϴ� " & MaxFileSize & " KB���ļ�\n",-1):response.End()
			End Select
			If ReturnValue<>"" Then
			 dim rsu:set rsu=server.createobject("adodb.recordset")
			 rsu.open "select banner,blogid,logo from ks_blog where username='" & KSUser.UserName & "'",conn,1,3
			 if not rsu.eof then
			   dim obanner,nbanner,k,nstr
			   obanner=split(rsu(0),"|")
			   nbanner=split(returnvalue,"|")
			   for k=0 to ubound(nbanner)
			     if k=0 then
				   if trim(nbanner(0))<>"" then nstr=nbanner(k) else nstr=obanner(k)
				 else
				   if nbanner(k)<>"" then 
				    nstr=nstr & "|" & nbanner(k)
				   else 
				     if ubound(obanner)>=k then
					  nstr=nstr& "|"&obanner(k)
					 else
					  nstr=nstr &"|"
					 end if
				   end if
				 end if
			   next
			    If Not KS.IsNul(rsu("Logo")) or Not KS.IsNul(nstr) Then
					Call KS.FileAssociation(1025,rsu("BlogID"),rsu("logo") & nstr,1)
				End If

			 end if
			 rsu.close
			 set rsu=nothing
            Conn.Execute("Update KS_Blog Set Banner='" & nstr & "' Where UserName='" & KSUser.UserName & "'")
			

			
			Call KSUser.AddLog(KSUser.UserName,"�����ռ��banner����!",102)
			End If
			Response.Write "<script>alert('��ϲ,banner�ϴ��ɹ�!');</script>"
	   End If
		on error resume next
	   	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Blog Where UserName='" & KSUser.UserName &"'",conn,1,1
		If Not RS.EOF Then
		 if Not KS.IsNul(RS("Banner")) Then
		 Banner=Split(RS("Banner"),"|")
		 End If
	    End If
		RS.Close:Set RS=Nothing
		dim b1,b2,b3
		 b1=banner(0)
	   if ubound(banner)>=1 then b2=banner(1)
	   if ubound(banner)>=2 then b3=banner(2)
	    if b1="" or isnull(b1) then b1="../images/ad1.jpg"
	    if b2="" or isnull(b2) then b2="../images/ad1.jpg"
	    if b3="" or isnull(b3) then b3="../images/ad1.jpg"
      %>
	    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
          <form  action="?Action=Banner&act=Save" method="post" name="myform" id="myform" enctype="multipart/form-data">

            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner1Ԥ����</strong><br>
              </div></td>
              <td align="center">��
                <img src="<%=b1%>" width="600" height="100"><br>
              ֻ֧��jpg��gif��png��С��200k��ͼƬ�Ĵ�С������Լ�ѡ��ģ���µı�ע���</td>
            </tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>��ַ��</strong><br>
              </div></td>
              <td><input type="file" name="photourl1" size="60"></td>
			</tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner2Ԥ����</strong><br>
              </div></td>
              <td align="center">��
                <img src="<%=b2%>" width="600" height="100"><br>
              ֻ֧��jpg��gif��png��С��200k��ͼƬ�Ĵ�С������Լ�ѡ��ģ���µı�ע���</td>
            </tr>
			<tr class="tdbg">
			  <td  height="25" align="center"><div align="left"><strong>��ַ��</strong><br>
              </div></td>
			  <td><input type="file" name="photourl2" size="60">
			  </td>
			</tr>
			<tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>Banner3Ԥ����</strong><br>
              </div></td>
              <td align="center">��
                <img src="<%=b3%>" width="600" height="100"><br>
              ֻ֧��jpg��gif��png��С��200k��ͼƬ�Ĵ�С������Լ�ѡ��ģ���µı�ע���</td>
            </tr>
			<tr class="tdbg">
			<td  height="25" align="center"><div align="left"><strong>��ַ��</strong><br>
              </div></td>
			  <td><input type="file" name="photourl3" size="60">
               </td>
            </tr>
            <tr class="tdbg">
              <td height="30" align="center" colspan=2>
                <input type="submit" name="Submit3"  class="Button" value="��������" />
                          </td>
            </tr>
			</form>
		 </table>
	   <%
	   End Sub
	   
	   
	   '����ģ��
	   Sub Template()
	    Dim Flag:Flag=KS.ChkClng(KS.S("Flag"))
		If Flag=0 Then 
		 If KSUser.UserType=1 Then
		  Flag=4
		 Else
		  Flag=2
		 End If
		End If
		
		if flag=2 or flag=4 then
	    Call KSUser.InnerLocation("���ÿռ�ģ��")
		else
	    Call KSUser.InnerLocation("����Ȧ��ģ��")
		end if
		    MaxPerPage=8
			If KS.S("page") <> "" Then
				CurrentPage = KS.ChkClng(KS.S("page"))
			Else
				CurrentPage = 1
			End If
		%>
			    <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0" class="border">
                    <tr class="Title">
                      <td height="22" colspan=3>
					  <%if KSUser.UserType=1 Then%>
					  <a href="?Action=Template&Flag=4"><b>���ÿռ�ģ��</b></a>
					  <%Else%>
					  <a href="?Action=Template&Flag=2"><b>���ÿռ�ģ��</b></a>
					  <%end if%> | <a href="?Action=Template&Flag=3"><b>����Ȧ��ģ��</b></a>
					  </td>
					  
					  <td style="display:none"><%if KSUser.UserType=1 Then%><a href="?action=UpTemplate">����Լ��Ŀռ�ģ��</a><%end if%></td>
					  
                    </tr>
                   <%
						Set RS=Server.CreateObject("AdodB.Recordset")
							RS.open "select * from ks_blogtemplate where TemplateAuthor='" & KSUser.username & "' or (usertag=0 and flag=" & Flag &") order by usertag desc,id desc",conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' height=30 valign=top>û�п���ģ��!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
								   If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowTemplate
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowTemplate
									Else
										CurrentPage = 1
										Call ShowTemplate
									End If
								End If
				End If
     %>                     
				</table>

		<%
		
	   End Sub
	   
	   Sub ShowTemplate()
	   %>
	   <style type="text/css">
	   	.t .onmouseover { background: #fffff0; }
		.t .onmouseout {}
		.t ul {float:left;margin:6px;padding:5px;width:152px!important;width:165px;height:180px;overflow:hidden;border: 1px #f4f4f4 solid;background: #fcfcfc;}
		.t ul li {
		list-style-type:none;line-height:1.5;margin:0;padding:0;}
		.t ul li.l1 img {width:150px;height:90px;}
		.t ul li.l1 a {display:block;margin:auto;padding:1px;width:156px;height:96px;background:url("images/tbg.png") no-repeat left top;text-align:left;}
		.t ul li.l2 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.t ul li.l3 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.t ul li.l4 {margin:10px 0 0 0;text-align:center;}
	   </style>
	   <%
	     dim i,k
	     do while not rs.eof
		   response.write "<tr>"
		   for i=1 to 4
		    response.write "<td class=""t"" width=""25%"">"
			 dim pic:pic=rs("templatepic")
			 if pic="" or isnull(pic) then pic="../images/nophoto.gif"
			%>
			<ul onMouseOver="this.className='onmouseover'" onMouseOut="this.className='onmouseout'" class="onmouseout">
				<li class="l1"><a href='../space/showtemplate.asp?templateid=<%=rs("id")%>' target=_blank>
<img src="<%=pic%>" title="���Ԥ��" width="200" height="122" border="0" />
</a></li>
				<li class="l2">���ƣ�<strong><%=rs("templatename")%></strong></li>
				<li class="l3">
				<%if rs("templateauthor")=KSUser.UserName then%>
				<!--<a href="?action=UpTemplate&ID=<%=RS("ID")%>"><font color=red>�޸�ģ��</font></a> | <a href="?action=DelTemplate&ID=<%=rs("id")%>" onClick="return(confirm('ɾ��ģ�岻�ɻָ���ȷ����'))"><font color=red>ɾ��ģ��</font></a>-->
				<%else%>
				���ߣ�<%=rs("templateauthor")%>
				<%end if%>
				
				</li>
				<%if rs("flag")=3 then
				 if Not KS.IsNul(rs("groupid")) And KS.FoundInArr(rs("groupid"),KSUser.GroupID,",")=false And KSUser.GroupID<>1 Then
				   response.write "<li class=""l4""><font color=red>��ģ��Vipר��</font></li>"
				 else
				 %>
					<li class="l4">Ȧ�ӣ�
					<select name='teamid<%=rs("id")%>' id='teamid<%=rs("id")%>' style='width:60px'>
					 <%dim rst:set rst=server.createobject("adodb.recordset")
					 rst.open "select * from ks_team where username='" & KSUser.UserName & "'",conn,1,1
					 if rst.eof then
					  response.write "<option value='0'>û�н�Ȧ��</option>"
					 else
					 do while not rst.eof
					  response.write "<option value='" & rst("id") & "'>" & rst("teamname") &"</option>"
					  rst.movenext
					 loop
					 end if
					 rst.close:set rst=nothing
					 %>
					</select>
					<input type="submit" value="Ӧ��" onClick="if($('#teamid<%=rs("id")%>').val()==0){alert('��ѡ��Ȧ��!');return false} else{window.location='?flag=3&teamid='+$('#teamid<%=rs("id")%>').val(<%=rs("id")%>)+'&action=SaveMySkin&id=<%=RS("ID")%>'}" />
					</li>
				<%
				 end if
				else%>
				<li class="l4">
				<%
				if Not KS.IsNul(rs("groupid")) And KS.FoundInArr(rs("groupid"),KSUser.GroupID,",")=false And KSUser.GroupID<>1 Then%>
				<input type="submit" disabled value="VIPר��ģ��"/>
				<%else%>
				<input type="submit" value="Ӧ�ô�ģ��" onClick="window.location='?action=SaveMySkin&id=<%=RS("ID")%>'" />
				<%end if%>
				<input type="submit" value="Ԥ��" onClick="window.open('../space/showtemplate.asp?templateid=<%=RS("ID")%>');" />
				</li>									
				<%end if%>
			</ul>
			<%
			response.write "</td>"
			rs.movenext
			k=k+1
			if rs.eof or k>=MaxPerPage then exit for 
		   next
		   for i=k+1 to 4
		    response.write "<td width=""25%"">&nbsp;</td>"
		   next
		  response.write "</tr>"
		  if rs.eof or k>=MaxPerPage then exit do
		 loop
		 response.write "<tr>"
		 response.write "<td colspan=4 align=""right"">"
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
		 Response.write "</td>"
		 response.write "</tr>"
	   End Sub
	   
	   Sub SaveMySkin()
	     Dim Flag:Flag=KS.ChkClng(KS.S("Flag"))
	     Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 IF ID=0 Then Exit Sub
		 if flag=3 then
		 Conn.Execute("Update KS_Team Set TemplateID=" & ID & " Where id=" & KS.ChkClng(KS.S("TeamID")))
		 response.write "<script>alert('��ϲ���ɹ�Ӧ����ѡ��Ȧ��ģ�壡');location.href='?action=Template&flag=3';</script>"
		 else
		 Conn.Execute("Update KS_Blog Set TemplateID=" & ID & " Where UserName='" & KSUser.UserName & "'")
		 response.write "<script>alert('��ϲ���ɹ�Ӧ���˿ռ�վ��ģ�壡');location.href='?action=Template';</script>"
		 end if
		 'response.redirect "?action=Template"
	   End Sub
	   
	 Sub UpTemplate()
	    dim templatename,templateauthor,templatemain,templatesub,Action,templatepic
	  redim templatesub(10)
	  dim rs:set rs=server.createobject("adodb.recordset")
	  rs.open "select * from KS_BlogTemplate Where ID="&KS.chkclng(KS.g("id")),conn,1,1
	  if not rs.eof then
	   templatename=rs("templatename")
	   templateauthor=rs("templateauthor")
	   templatepic=rs("templatepic")
	   templatemain=rs("templatemain")
	   templatesub=split(rs("templatesub"),"^%^KS^%^")
	    Call KSUser.InnerLocation("�޸Ŀռ�ģ��")
	 else
	  templatesub(0)=""
	  templatesub(1)=""
	  templatesub(2)=""
	   Call KSUser.InnerLocation("��ӿռ�ģ��")
	 end if

%>
<script src="../ks_inc/kesion.box.js" language="JavaScript"></script>
<script language="javascript">
 function CheckForm()
 {
    if (document.all.TemplateName.value=='')
	{
	  alert('������ģ������!');
	  document.all.TemplateName.focus();
	  return false;
	}
    if (FCKeditorAPI.GetInstance('TemplateMain').GetXHTML(true)=="")
	{
	  alert('��������ģ�������!');
	  return false;
	}
    if (FCKeditorAPI.GetInstance('TemplateMain').GetXHTML(true).indexOf('{$BlogMain}')<=0)
	{
	  alert('��ģ��ĸ�ʽ����,��ģ��������{$BlogMain}��ǩ!');
	  return false;
	}
	
    if (FCKeditorAPI.GetInstance('TemplateSub0').GetXHTML(true)=="")
	{
	  alert('�����븱ģ�������!');
	  return false;
	}
	return true;
 }
function ShowIframe(flag)
{popupIframe("�鿴�ռ�վ��Ŀ��ñ�ǩ","../ks_editor/spacelabel.asp?flag="+flag,550,300,'no')
}
function InsertLabel(obj,Val)
{
	oEditor=FCKeditorAPI.GetInstance(obj);
	if(oEditor   &&   oEditor.EditorWindow){ 
	oEditor.InsertHtml(Val); 
    } 
  
  closeWindow();
 }
function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj)
{
var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
if (ReturnStr!='') SetObj.value=ReturnStr;
}
</script>
  <table width="98%" border="0" align="center" cellspacing="1" cellpadding="3" class="border">
 <form method="POST" action="user_blog.asp" id="myform" name="myform">
    <tr class="tdbg">
      <td colspan=2 align="center" height="25">&nbsp;&nbsp;ģ�����ƣ� 
        <input name="TemplateName" type="text" class="textbox" id="TemplateName" value="<%=templatename%>">
        ��
        <input name="TemplateAuthor" type="hidden" id="TemplateAuthor" value="<%=KSUser.username%>">
		Ԥ��ͼ��
		<input type="text" name="TemplatePic"  class="Textbox" value="<%=templatepic%>">&nbsp;<input class="button" type='button' name='Submit3' value='ѡ��ͼƬ��ַ...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&amp;pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&amp;ChannelID=999',500,360,window,document.all.TemplatePic);" />
	  </td>
    </tr>

    <tr> 
	  <td height="25" class="clefttitle" align="right"><strong>��ҳ����ģ�壺</strong><br /><br><a href="javascript:ShowIframe(2)"><u><font color=#ff6600>�鿴/������ñ�ǩ</font></u></a></td>
      <td height="25" class="tdbg" align="center">
	  <% 	  
	  Response.Write "<textarea ID='TemplateSub0' name='TemplateSub0' style='display:none'>" & templatesub(0) & "</textarea>"
	  Response.Write "<iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=TemplateSub0&amp;Toolbar=Simple"" width=""580"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"
	  %>
	  <textarea name="TemplateSub0s" id='edit' style="display:none;width:560px;height:100px" class="textbox"><%=templatesub(0)%></textarea>
      </td>
    </tr>
    <tr class="tdbg"> 
	  <td height="25" class="clefttitle" align="right"><strong>����ҳ���ģ�壺</strong>
	  <br /><br><a href="javascript:ShowIframe(1)"><u><font color=#ff6600>�鿴/������ñ�ǩ</font></u></a></td>
      <td height="25" align="center">
	  
	  <%
	  Response.Write "<textarea ID='TemplateMain' name='TemplateMain' style='display:none'>" & templatemain & "</textarea>"
	  Response.Write "<iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=TemplateMain&amp;Toolbar=Simple"" width=""580"" height=""250"" frameborder=""0"" scrolling=""no""></iframe>"
	  %>
	  <textarea name="TemplateMains" id='edit' style="display:none;" class="textbox" rows=10><%=templatemain%></textarea>
      </td>
    </tr>
    <tr> 
	 <td height="25" class="clefttitle" align="right"><strong>��ģ�壨��־����</strong><br /><br><a href="javascript:ShowIframe(3)"><u><font color=#ff6600>�鿴/������ñ�ǩ</font></u></a></td>
      <td height="25" class="tdbg" align="center">
	  	  <%
	  Response.Write "<textarea ID='TemplateSub1' name='TemplateSub1' style='display:none'>" & templatesub(1) & "</textarea>"
	  Response.Write "<iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=TemplateSub1&amp;Toolbar=Simple"" width=""580"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"
	  %>

	  <textarea name="TemplateSub1s" id='edit' style="display:none;width:560px;height:100px" class="textbox"><%=templatesub(1)%></textarea>
      </td>
    </tr>
	
    <tr> 
	  <td height="25" class="clefttitle" align="right"><strong>��ģ�壨��ϵ���ǣ���</strong><br /><br><a href="javascript:ShowIframe(5)"><u><font color=#ff6600>�鿴/������ñ�ǩ</font></u></a>
	   
	  </td>
      <td height="25" class="tdbg" align="center">
	  <%
	  Response.Write "<textarea ID='TemplateSub2' name='TemplateSub2' style='display:none'>" & templatesub(2) & "</textarea>"
	  Response.Write "<iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=TemplateSub2&amp;Toolbar=Simple"" width=""580"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"
	  %>
	  <textarea name="TemplateSub2s" id='edit' style="display:none;width:560px;height:100px" class="textbox"><%=templatesub(2)%></textarea>
      </td>
    </tr>
	
    <tr> 
      <td class="tdbg" colspan=2> <div align="center">
        <input name="Action" type="hidden" id="Action" value="UpTemplateSave"> 
		<input name="id" type="hidden" value="<%=KS.g("id")%>">
        <input name="cmdSave" type="submit" class="button" id="cmdSave" value=" ����ģ�� " onClick="return(CheckForm());"> 
      </div></td>
    </tr>
</form>
  </table>
 <%
	   End Sub
	   
	   Sub UpTemplateSave
			dim rs,sql,flag,TemplateMain,templatesub0,templatesub1,templatesub2
			templatemain=KS.CheckScript(Replace(Replace(Request("TemplateMain"),"<%","&lt;%"),"%"&">","%&gt;"))
			templatesub0=KS.CheckScript(Replace(Replace(Request("TemplateSub0"),"<%","&lt;%"),"%"&">","%&gt;"))
			templatesub1=KS.CheckScript(Replace(Replace(Request("TemplateSub1"),"<%","&lt;%"),"%"&">","%&gt;"))
			templatesub2=KS.CheckScript(Replace(Replace(Request("TemplateSub2"),"<%","&lt;%"),"%"&">","%&gt;"))
			If Instr(TemplateMain,"{$BlogMain}")=0 Then
			 Response.Write "<script>alert('�Բ�����ģ���ʽ������ģ��������{$BlogMain}��ǩ!');history.back();</script>"
			 Response.End
			End If
			set rs=server.CreateObject("adodb.recordset")
			sql="select * From KS_BlogTemplate where id=" & KS.chkclng(KS.g("id"))
			rs.open sql,conn,1,3
			If rs.eof Then
			 rs.addnew
			end if
			rs("TemplateName")=KS.S("TemplateName")
			rs("TemplateAuthor")=KS.S("TemplateAuthor")
			rs("TemplateMain")=templatemain
			rs("TemplatePic")=KS.S("TemplatePic")
			rs("templatesub")=templatesub0&"^%^KS^%^"&templatesub1&"^%^KS^%^"&templatesub2
			rs("isdefault")="false"
			rs("usertag")=1
			rs("flag")=4
			rs.update
			rs.close:set rs=nothing
			If KS.chkclng(KS.g("id"))=0 then
			response.Write  "<script>alert('ģ����ӳɹ�!');location.href='User_Blog.asp?Action=Template';</script>"
			else
			response.Write  "<script>alert('ģ���޸ĳɹ�!');location.href='User_Blog.asp?Action=Template';</script>"
			end if
	   End Sub

	
	 'ɾ��ģ��
	 Function DelTemplate()
	 	Dim ID:ID=KS.ChkClng(KS.S("ID"))
		If ID=0 Then Call KS.Alert("��û��ѡ��Ҫɾ����ģ��!",ComeUrl):Response.End
		Conn.Execute("Delete From KS_BlogTemplate Where TemplateAuthor='" & KSUser.UserName & "' and ID=" & ID)
		Dim NewID:NewID=Conn.Execute("Select top 1 id from ks_blogtemplate where flag=4 and isdefault='true'")(0)
		Conn.Execute("Update KS_Blog Set TemplateID=" & NewID & " where username='" & KSUser.UserName & "' and templateid=" & ID)
		Response.Redirect ComeUrl

	 End Function

	   
	  

	   
	  

	   
	   '��־�б�
	   Sub BlogList()
			 
			    

			    If KS.S("page") <> "" Then
					 CurrentPage = KS.ChkClng(KS.S("page"))
				Else
					 CurrentPage = 1
				End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
                                    Status=KS.S("Status")
									If Status<>"" and isnumeric(Status) Then 
									   Param= Param & " and Status=" & Status
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And Tags like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("TypeID")<>"" And KS.S("TypeID")<>"0" Then Param=Param & " And TypeID=" & KS.ChkClng(KS.S("TypeID")) & ""
									Dim Sql:sql = "select * from KS_BlogInfo "& Param &" order by AddDate DESC"
								  Select Case ks.s("Status")
								   Case "0" 
								    Call KSUser.InnerLocation("������־�б�")
								   Case "1"
								    Call KSUser.InnerLocation("�ݸ���־�б�")
								   Case "2"
								    Call KSUser.InnerLocation("δ����־�б�")
                                   Case Else
								    Call KSUser.InnerLocation("������־�б�")
								   End Select
								  %>
								     
				                    <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                                                <tr class="Title">
                                                  <td width="6%" height="22" align="center">ѡ��</td>
												  <td width="12%" height="22" align="center">��־����</td>
                                                  <td width="41%" height="22" align="center">��־����</td>
                                                  <td width="12%" height="22" align="center">���ʱ��</td>
                                                  <td width="8%" height="22" align="center">״̬</td>
                                                  <td width="21%" height="22" align="center" nowrap>�������</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>û����Ҫ����־!</td></tr>"
								 Else
									    totalPut = RS.RecordCount
										If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowLog
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowLog
									Else
										CurrentPage = 1
										Call ShowLog
									End If
								End If
				End If
     %>                      <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                  <form action="User_Blog.asp" method="post" name="searchform">
                                  <td height="45" colspan=6 align="center">
										��־������
										  <select name="Flag">
										   <option value="0">����</option>
										   <option value="1">��ǩ</option>
									      </select>
										  <select size='1' name='TypeID'>
										 <option value="0">-��ѡ����־����-</option>
                                           <% Dim RS1:Set RS1=Server.CreateObject("ADODB.RECORDSET")
							  RS1.Open "Select * From KS_BlogType order by orderid",conn,1,1
							  If Not RS1.EOF Then
							   Do While Not RS1.Eof 
							    
								  Response.Write "<option value=""" & RS1("TypeID") & """>" & RS1("TypeName") & "</option>"
								 RS1.MoveNext
							   Loop
							  End If
							  RS1.Close:Set RS1=Nothing
							  %>
                                        </select>
										  �ؼ���
										  <input type="text" name="KeyWord" class="textbox" value="�ؼ���" size=20>&nbsp;<input  class="Button" type="submit" name="submit1" value="����">
							      </td>
								    </form>
                                </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowLog()
     Dim I
    Response.Write "<FORM Action=""User_Blog.asp?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td class="splittd" height="20" align="center">
											<INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
											<td class="splittd" align="center"><%=Conn.Execute("Select TypeName From KS_BlogType Where TypeID=" & RS("TypeID"))(0)%></td>
                                            <td class="splittd" align="left"><a href="../space/?<%=KSUser.UserName%>/log/<%=rs("id")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("title")),35)%></a></td>
                                            <td class="splittd" align="center"><%=KS.GetTimeFormat(rs("adddate"))%></td>
                                            <td class="splittd" align="center">
											  <%Select Case rs("Status")
											   Case 0
											     Response.Write "<span class=""font10"">����</span>"
											   Case 1
											     Response.Write "<span class=""font11"">�ݸ�</span>"
                                               Case 2
											     Response.Write "<span class=""font13"">δ��</span>"
                                              end select
											  %></td>
                                            <td class="splittd" align="center">
											<%if ks.SSetting(3)=1 and rs("status")=0 then%>
											<%else%>
											<a href="User_Blog.asp?id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>" class="box">�޸�</a><%end if%> <a href="User_Blog.asp?action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('ȷ��ɾ����־��?'))" class="box">ɾ��</a>
											</td>
                                          </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=6 valign=top>
								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ�б�ҳ��ʾ��������־<INPUT  class="Button" onClick="return(confirm('ȷ��ɾ��ѡ�е���־��?'));" type=submit value=ɾ��ѡ������־ name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;         
								
								  </td>
								  </FORM>
								</tr>
								<tr>
								 <td colspan=6>
								 <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								 </td>
								</tr>
								<% 
  End Sub
  'ɾ����־
  Sub ArticleDel()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ������־!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_BlogInfo Where UserName='" & KSUser.userName & "' And ID In(" & ID & ")")
	Conn.Execute("Delete From KS_UploadFiles Where channelid=1026 and InfoID In(" & ID & ")")
	Call KSUser.AddLog(KSUser.UserName,"ɾ������־����!",101)
	Response.Redirect ComeUrl
  End Sub
  '�����־
  Sub ArticleAdd()
        Call KSUser.InnerLocation("������־")
  		if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
		     TypeID  = RSObj("TypeID")
			 ClassID = RSObj("ClassID")
			 Title    = RSObj("Title")
			 Tags = RSObj("Tags")
			 UserName   = RSObj("UserName")
			 password = RSObj("password")
			 Face   = RSObj("Face")
			 weather=RSObj("Weather")
			 adddate=RSObj("adddate")
			 Content  = RSObj("Content")
			 Status  = RSObj("Status")
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		  adddate=now:weather="sun.gif":Face=1:TypeID=0:UserName=KSUser.RealName
		End If
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.TypeID.value=="0") 
				  {
					alert("��ѡ����־���࣡");
					document.myform.TypeID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("��������־���⣡");
					document.myform.Title.focus();
					return false;
				  }	
				  
				  if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=="")
					{
					  alert("��������־���ݣ�");
					  FCKeditorAPI.GetInstance('Content').Focus();
					  return false;
					}
				
				 return true;  
				}
				function Chang(picurl,V,S)
				{
					var pic=S+picurl
					if (picurl!=''){
					document.getElementById(V).src=pic;
					}
				}
            function InsertFileFromUp(FileList,InstallDir)
		  { 
				Files=FileList.split("|");
				for(var i=0;i<Files.length-1;i++)
				{     var ext=getFilePic(Files[i]);
					  var files=Files[i].split('/');
					  var file=files[files.length-1];
					  var br='';
					  if (i!=Files.length-1) br='<br />';
					  var fileext = Files[i].substring(Files[i].lastIndexOf(".") + 1, Files[i].length).toLowerCase();
                      if (fileext=="gif" || fileext=="jpg" || fileext=="jpeg" || fileext=="bmp" || fileext=="png")
					  {
					   insertHTMLToEditor('<img src="'+Files[i]+'" border="0"/><br/>');	
					  }
					  else
					  {
					  var str="<img border=0 src="+InstallDir+"KS_Editor/images/FileIcon/"+ext+"> <a href='"+Files[i]+"'  target='_blank'>[���������ļ�:"+file+"]</a>"+br;
					  insertHTMLToEditor(str);	
					  }
					  
				 }
		}
		function insertHTMLToEditor(codeStr) 
			{ 
				oEditor=FCKeditorAPI.GetInstance("Content");
				if(oEditor   &&   oEditor.EditorWindow){ 
					oEditor.InsertHtml(codeStr); 
				} 
			} 
            // ���ļ���չ��ȡͼ������������
			function getFilePic(url)
			{
				var sExt;
				sExt=url.substr(url.lastIndexOf(".")+1);
				sExt=sExt.toUpperCase();
				var sPicName;
				switch(sExt)
				{
				case "TXT":
					sPicName = "txt.gif";
					break;
				case "CHM":
				case "HLP":
					sPicName = "hlp.gif";
					break;
				case "DOC":
					sPicName = "doc.gif";
					break;
				case "PDF":
					sPicName = "pdf.gif";
					break;
				case "MDB":
					sPicName = "mdb.gif";
					break;
				case "GIF":
					sPicName = "gif.gif";
					break;
				case "JPG":
					sPicName = "jpg.gif";
					break;
				case "BMP":
					sPicName = "bmp.gif";
					break;
				case "PNG":
					sPicName = "pic.gif";
					break;
				case "ASP":
				case "JSP":
				case "JS":
				case "PHP":
				case "PHP3":
				case "ASPX":
					sPicName = "code.gif";
					break;
				case "HTM":
				case "HTML":
				case "SHTML":
					sPicName = "htm.gif";
					break;
				case "ZIP":
				case "RAR":
					sPicName = "zip.gif";
					break;
				case "EXE":
					sPicName = "exe.gif";
					break;
				case "AVI":
					sPicName = "avi.gif";
					break;
				case "MPG":
				case "MPEG":
				case "ASF":
					sPicName = "mp.gif";
					break;
				case "RA":
				case "RM":
					sPicName = "rm.gif";
					break;
				case "MP3":
					sPicName = "mp3.gif";
					break;
				case "MID":
				case "MIDI":
					sPicName = "mid.gif";
					break;
				case "WAV":
					sPicName = "audio.gif";
					break;
				case "XLS":
					sPicName = "xls.gif";
					break;
				case "PPT":
				case "PPS":
					sPicName = "ppt.gif";
					break;
				case "SWF":
					sPicName = "swf.gif";
					break;
				default:
					sPicName = "unknow.gif";
					break;
				}
				return sPicName;
			
			}
			
		</script>				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Blog.asp?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">

                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>��־���ࣺ</span></td>
                       <td width="88%">��
                          <select class="textbox" size='1' name='TypeID' style="width:150">
                             <option value="0">-��ѡ�����-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogType order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If TypeID=RS("TypeID") Then
								  Response.Write "<option value=""" & RS("TypeID") & """ selected>" & RS("TypeName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("TypeID") & """>" & RS("TypeName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>
						   ר��
						      <select class="textbox" size='1' name='ClassID' style="width:150">
                                            <option value="0">-ѡ���ҵ�ר��-</option>
                                            <%=KSUser.UserClassOption(2,ClassID)%>
                         </select>		
						 
						 <a href="User_Class.asp?Action=Add&typeid=2"><font color="red">����ҵķ���</font></a>			
					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>��־���⣺</span></td>
                              <td> ��
                                        <input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                    </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>��־���ڣ�</span></td>
                                      <td>��
                                        <input name="AddDate"  class="textbox" type="text" id="AddDate" value="<%=adddate%>" style="width:250px; " />
                                       ����<Select Name="Weather" Size="1" onChange="Chang(this.value,'WeatherSrc','images/weather/')">
									   <Option value="sun.gif"<%if weather="sun.gif" then response.write " selected"%>>����</Option>
									   <Option value="sun2.gif"<%if weather="sun2.gif" then response.write " selected"%>>����</Option>
									   <Option value="yin.gif"<%if weather="yin.gif" then response.write " selected"%>>����</Option>
									   <Option value="qing.gif"<%if weather="qing.gif" then response.write " selected"%>>��ˬ</Option>
									   <Option value="yun.gif"<%if weather="yun.gif" then response.write " selected"%>>����</Option>
									   <Option value="wu.gif"<%if weather="wu.gif" then response.write " selected"%>>����</Option>
									   <Option value="xiaoyu.gif"<%if weather="xiaoyu.gif" then response.write " selected"%>>С��</Option>
									   <Option value="yinyu.gif"<%if weather="yinyu.gif" then response.write " selected"%>>����</Option>
									   <Option value="leiyu.gif"<%if weather="leiyu.gif" then response.write " selected"%>>����</Option>
									   <Option value="caihong.gif"<%if weather="caihong.gif" then response.write " selected"%>>�ʺ�</Option>
									   <Option value="hexu.gif"<%if weather="hexu.gif" then response.write " selected"%>>����</Option>
									   <Option value="feng.gif"<%if weather="feng.gif" then response.write " selected"%>>����</Option>
									   <Option value="xue.gif"<%if weather="xue.gif" then response.write " selected"%>>Сѩ</Option>
									   <Option value="daxue.gif"<%if weather="daxue.gif" then response.write " selected"%>>��ѩ</Option>
									   <Option value="moon.gif"<%if weather="moon.gif" then response.write " selected"%>>��Բ</Option>
									   <Option value="moon2.gif"<%if weather="moon2.gif" then response.write " selected"%>>��ȱ</Option>
									</Select>
		<img id="WeatherSrc" src="images/weather/<%=weather%>" border="0"></td>
                              </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>Tag�� ǩ��</span></td>
                                      <td>��
                                        <input name="Tags"  class="textbox" type="text" id="Tags" value="<%=Tags%>" style="width:250px; " />
                                        �Կո�ָ�</td>
                              </tr>
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span>��־���飺</span></td>
                                <td>
									  &nbsp;&nbsp;<input type="radio" name="face" value="0"<%If face=0 Then Response.Write " checked"%>>
        ��<input name="face" type="radio" value="1"<%If face=1 Then Response.Write " checked"%>><img src="images/face/1.gif" width="20" height="20"> 
        <input type="radio" name="face" value="2"<%If face=2 Then Response.Write " checked"%>><img src="images/face/2.gif" width="20" height="20"><input type="radio" name="face" value="3"<%If face=3 Then Response.Write " checked"%>><img src="images/face/3.gif" width="20" height="20"> 
        <input type="radio" name="face" value="4"<%If face=4 Then Response.Write " checked"%>><img src="images/face/4.gif" width="20" height="20"> 
        <input type="radio" name="face" value="5"<%If face=5 Then Response.Write " checked"%>><img src="images/face/5.gif" width="20" height="20"> 
        <input type="radio" name="face" value="6"<%If face=6 Then Response.Write " checked"%>><img src="images/face/6.gif" width="18" height="20"> 
        <input type="radio" name="face" value="7"<%If face=7 Then Response.Write " checked"%>><img src="images/face/7.gif" width="20" height="20"> 
        <input type="radio" name="face" value="8"<%If face=8 Then Response.Write " checked"%>><img src="images/face/8.gif" width="20" height="20"> 
        <input type="radio" name="face" value="9"<%If face=9 Then Response.Write " checked"%>><img src="images/face/9.gif" width="20" height="20">
        <input type="radio" name="face" value="10"<%If face=10 Then Response.Write " checked"%>><img src="images/face/10.gif" width="20" height="20">
        <input type="radio" name="face" value="11"<%If face=11 Then Response.Write " checked"%>><img src="images/face/11.gif" width="20" height="20">
        <input type="radio" name="face" value="12"<%If face=12 Then Response.Write " checked"%>><img src="images/face/12.gif" width="20" height="20"></td>
                              </tr>
							 <%If KS.SSetting(26)="1" Then%>
                              <tr class="tdbg">
                                  <td align="center">�����ϴ���</td>
								  <td align=center>
								   <iframe id='UpFileFrame' name='UpFileFrame' src='../user/User_UpFile.asp?Type=File&ChannelID=9993' frameborder=0 scrolling=no width='100%' height='30' allowtransparency="true"></iframe>
								  </td>
							  </tr>
							 <%End If%>
                              <tr class="tdbg">
                                  <td align="center">��־���ݣ�</td>
								  <td align=center>
								  <%
								  Response.Write "<textarea ID='Content' name='Content' style='display:none'>" & Server.HTMLEncode(Content) & "</textarea>"
					               Response.Write "<iframe id=""Content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""93%"" height=""320"" frameborder=""0"" scrolling=""no""></iframe>"
								  %>
								</td>
                            </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>��־���룺</span></td>
                                <td>��
                                        <input name="Password"  class="textbox" type="password" id="PassWord" value="<%=PassWord%>" style="width:250px; " />
                                        ��                                   ��־״̬��
                                        <input name="Status" type="checkbox" value="1" <%If Status=1 Then Response.Write " checked"%> />
����ݸ��� </td>
                              </tr>
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,�������� " />
                      <input type="reset" name="Submit2"   class="Button" value=" �� �� " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub

   Sub DoSave()
                 TypeID=KS.ChkClng(KS.S("TypeID"))
				 ClassID=KS.ChkClng(KS.S("ClassID"))
				 Title=Trim(KS.S("Title"))
				 Tags=Trim(KS.S("Tags"))
				 UserName=Trim(KS.S("UserName"))
				 Face=Trim(KS.S("Face"))
				 weather=KS.S("weather")
				 adddate=KS.S("adddate")
				 Content = Request.Form("Content")
				 Content=KS.ScriptHtml(Content, "A", 3)
				 Content=KS.ClearBadChr(content)
				 PassWord=KS.S("password")
				 Status=KS.ChkClng(KS.S("Status"))
				  Dim RSObj
				  
				  if TypeID="" Then TypeID=0
				  If TypeID=0 Then
				    Response.Write "<script>alert('��û��ѡ����־����!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('��û��������־����!');history.back();</script>"
				    Exit Sub
				  End IF
				  if not isdate(adddate) then
				    Response.Write "<script>alert('����������ڲ���ȷ!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>alert('��û��������־����!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Hits")=0
				End If
				  RSObj("Title")=Title
				  RSObj("TypeID")=TypeID
				  RSObj("ClassID")=ClassID
				  RSObj("Tags")=Tags
				  RSObj("UserName")=KSUser.UserName
				  RSObj("Face")=Face
				  RSObj("Weather")=weather
				  RSObj("Adddate")=adddate
				  RSObj("Content")=Content
				  RSObj("Password")=Password
				  if status=1 then
				  RSObj("Status")=1
				  elseif KS.SSetting(3)=1 Then
				  RSObj("Status")=2
				  Else
				  RSObj("Status")=0
				  end if
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				 RSObj.Close:Set RSObj=Nothing
				 If KS.ChkCLng(KS.S("ID"))=0 Then
				  Call KS.FileAssociation(1026,InfoID,Content,0)
				  Call KSUser.AddLog(KSUser.UserName,"��������־ <a href='{$GetSiteUrl}space/?" & KSUser.username & "/log/" & InfoID & "' target='_blank'>" & Title & "</a>""!",101)
			   	  Response.Write "<script>if (confirm('������־�ɹ�������������?')){location.href='User_Blog.asp?Action=Add';}else{location.href='User_Blog.asp';}</script>"
				 Else
				   Call KS.FileAssociation(1026,InfoID,Content,1) 
				   Call KSUser.AddLog(KSUser.UserName,"�޸�����־ <a href='{$GetSiteUrl}space/?" & KSUser.username & "/log/" & InfoID & "' target='_blank'>" & Title & "</a>""!",101)
				  Response.Write "<script>alert('��־�޸ĳɹ�!');location.href='User_Blog.asp';</script>"
				 End If
  End Sub


End Class
%> 

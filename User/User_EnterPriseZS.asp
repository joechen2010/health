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
Set KSCls = New EnterPriseZSCls
KSCls.Kesion()
Set KSCls = Nothing

Class EnterPriseZSCls
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,Selbutton,Verific,PhotoUrl,bigclassid,smallclassid,flag
		Private F_B_Arr,F_V_Arr,ClassID,Title,Sxrq,Fzjg,Jzrq,Intro,Action,I
		Private Sub Class_Initialize()
			MaxPerPage =12
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		
		Call KSUser.Head()
		Call KSUser.InnerLocation("��������֤���б�")
		KSUser.CheckPowerAndDie("s13")
	
		If KS.SSetting(0)=0 Then
		 Call KS.Alert("�Բ��𣬱�վ�رտռ书�ܣ�","")
		 Exit Sub
		ElseIf Conn.Execute("Select Count(username) From KS_Blog Where UserName='" & KSUser.UserName & "'").eof Then
		    Response.Write "<script>alert('����û�п�ͨ�ռ�,��ȷ��ת��ͨҳ�棡');location.href='User_Enterprise.asp';</script>"
		End If

		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">�ҷ���������֤��(<span class="red"><%=conn.execute("select count(id) from ks_EnterPriseZS where username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">�����(<span class="red"><%=conn.execute("select count(id) from ks_EnterPriseZS where status=1 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">�����(<span class="red"><%=conn.execute("select count(id) from ks_EnterPriseZS where status=0 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
       </div>
		<%
		Select Case KS.S("Action")
		 Case "Del"  Call ArticleDel()
		 Case "Add","Edit" Call DoAdd()
		 Case "DoSave" Call DoSave()
		 Case Else Call ProductList()
		End Select
	   End Sub
	   Sub ProductList()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
                                    Verific=KS.S("Status")
                                    IF Verific<>"" Then 
									   Param= Param & " and status=" & KS.ChkClng(Verific)-1
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And Sxrq like '%" & KS.S("KeyWord") & "%'"
									End if
									Dim Sql:sql = "select * from KS_EnterPriseZS " & Param &" order by AddDate DESC"

								  
								  %>
								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="?Action=Add"><font color=red>����������֤��</font></a></div>
    
				                     <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1">
                                                <tr class="Title">
                                                  <td colspan="6" height="22" align="center">
												  <%if KS.S("Status")="1" Then
												     response.write "���������֤��"
													 elseif ks.s("status")="2" then
													 response.write "���������֤��"
													 else
													 response.write "��������֤���б�"
													 end if
												  %>
												  </td>
                                                </tr>
                                           
                                      <%
								 Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>�Ҳ����κ�����֤��!</td></tr>"
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
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
				End If
     %>               
	   <tr>
	     <td colspan=6>
		  <table border='0'>
		   <tr>
		    <td width="340" height="30">
			 <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ѡ����������֤�� <input value="ɾ��ѡ��" class="button" Click="return(confirm('ȷ��ɾ��ѡ�е�����֤����?'));" type=submit> 
			</form>
			</td>
			
			</tr>
		   </table>
		   <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		 </td>
	   </tr>
	 <form action="?" method="post" name="searchform">
	        <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                <td height="45" colspan=6 align="center">
									����֤��������
										  <select name="Flag">
										   <option value="0">����֤������</option>
										   <option value="1">��Ч����</option>
									      </select>
										  
										  �ؼ���
										  <input type="text" name="KeyWord" class="textbox" value="�ؼ���" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" �� �� ">
		      </td>
       </form>
                                </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowContent()
    Response.Write "<FORM Action=""?Action=Del"" name=""myform"" method=""post"">"
   
	%>
	   <style type="text/css">
	   	.onmouseover { background: #fffff0; }
		.onmouseout {}
		.zslist ul {float:left;margin:6px;padding:5px;width:152px!important;width:165px;height:180px;overflow:hidden;border: 1px #f4f4f4 solid;background: #fcfcfc;}
		.zslist ul li {
		list-style-type:none;line-height:1.5;margin:0;padding:0;}
		.zslist ul li.l1 img {width:150px;height:90px;}
		.zslist ul li.l1 a {display:block;margin:auto;padding:1px;width:156px;height:96px;background:url("images/tbg.png") no-repeat left top;text-align:left;}
		.zslist ul li.l2 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.zslist ul li.l3 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
		.zslist ul li.l4 {margin:10px 0 0 0;text-align:center;}
	   </style>
	   <%
	     dim i,k
	     do while not rs.eof
		   response.write "<tr>"
		   for i=1 to 4
		    response.write "<td width=""25%"" class=""zslist"">"
			 dim pic:pic=rs("photourl")
			 if pic="" or isnull(pic) then pic="../images/nophoto.gif"
			%>
			<ul onMouseOver="this.className='onmouseover'" onMouseOut="this.className='onmouseout'" class="onmouseout">
				<li class="l1"><img src="<%=pic%>" title="���Ԥ��" border="0" /></li>
				<li class="l2">֤�����ƣ�<strong><%=rs("Title")%></strong></li>
				<li class="l3">��֤���أ�<%=rs("fzjg")%></li>
				<li class="l4"><INPUT id="ID"  type="checkbox" value="<%=RS("ID")%>"  name="ID">
				<a href="?action=Edit&id=<%=RS("ID")%>">�޸�</a> | <a href="?Action=Del&ID=<%=RS("ID")%>" onClick="return(confirm('ȷ��ɾ��֤����?'));">ɾ��</a>
				</li>									
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

  End Sub
  'ɾ������
  Sub ArticleDel()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ��������֤��!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_EnterPriseZS Where UserName='" & KSUser.UserName & "' And ID In(" & ID & ")")
	if ComeUrl="" then
	Response.Redirect("../index.asp")
	else
	Response.Redirect ComeUrl
	end if
  End Sub

  '�������
  Sub DoAdd()
        Call KSUser.InnerLocation("����֤��")
  		if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_EnterPriseZS Where UserName='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
			 Title    = RSObj("Title")
			 Fzjg = RSObj("Fzjg")
			 Jzrq   = RSObj("Jzrq")
			 Sxrq  = RSObj("Sxrq")
			 PhotoUrl  = RSObj("PhotoUrl")
			 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/Images/NoPhoto.gif"
			 flag=true
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		 PhotoUrl="/images/Nophoto.gif"
		 ClassID=KS.S("ClassID")
		 If ClassID="" Then ClassID="0"
		 flag=false
		End If
		%>
		<script language="javascript" src="../ks_inc/popcalendar.js"></script>

		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("����������֤�����ƣ�");
					document.myform.Title.focus();
					return false;
				  }	
				if (document.myform.Fzjg.value=="")
				  {
					alert("�����뷢֤������");
					document.myform.Fzjg.focus();
					return false;
				  }	
				if (document.myform.Sxrq.value=="")
				  {
					alert("��������Ч���ڣ�");
					document.myform.Sxrq.focus();
					return false;
				  }	
				if (document.myform.Jzrq.value=="")
				  {
					alert("�������ֹ���ڣ�");
					document.myform.Jzrq.focus();
					return false;
				  }	
				    
				 return true;  
				}
				</script>
				
				
				<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?Action=DoSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();" enctype="multipart/form-data">
				   <input type="hidden" value="<%=KS.S("ID")%>" name="id">
				    <tr  class="title">
					  <td colspan=3 align=center>
					       <%IF KS.S("Action")="Edit" Then
							   response.write "�޸�����֤��"
							   Else
							    response.write "��������֤��"
							   End iF
							  %>                         </td>
					</tr>
                    
                      <tr class="tdbg">
                           <td width="12%"  height="25" align="center"><span>֤�����ƣ�</span></td>
                              <td width="52%"> ��
                                        <input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                              <td width="36%" rowspan="4" align="center">
							  <img src="<%=photourl%>" width="160" height="130">							  </td>
                      </tr>
					 
                      <tr class="tdbg">
                                      <td  height="25" align="center"><span>��֤������</span></td>
                                      <td height="25">��
                                        <input name="Fzjg" class="textbox" type="text" style="width:250px; " value="<%=Fzjg%>" maxlength="30" />
                                        <span style="color: #FF0000">*</span></td>
                              </tr>
			  
                              <tr class="tdbg">
                                <td height="25" align="center">��Ч���ڣ�</td>
                                <td height="25">��
<input name="Sxrq" class="textbox" type="text" id="Sxrq" onClick="popUpCalendar(this, document.all.Sxrq, dateFormat,-1,-1)" style="width:250px; " value="<%=Sxrq%>" maxlength="30" />
<span style="color: #FF0000">*</span></td>
                              </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>��ֹ���ڣ�</span></td>
                                      <td height="25">��
                                        <input name="Jzrq" class="textbox" type="text" id="Jzrq" onClick="popUpCalendar(this, document.all.Jzrq, dateFormat,-1,-1)"  style="width:250px; " value="<%=Jzrq%>" maxlength="30" />
                                        <span style="color: #FF0000">*</span></td>
                              </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>֤����Ƭ��</span></td>
                        <td> ��
                               <input type="file" name="photourl" size="40">
                          <span style="color: #FF0000">*</span> <br>
                          �� <font color=red>˵����ֻ֧��JPG��GIF��PNG��ʽͼƬ��������500K</font></td>
                      </tr>
                        
                             
			  
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=3>
					 <input class="button" type="submit" name="Submit" value="OK, �� �� " />
                            ��
                            <input class="button" type="reset" name="Submit2" value=" �� �� " />						</td>
                    </tr>
                  </form>
			    </table>
		        <br>
		        <table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <TR class="title">
                    <TD  height="24"><STRONG>ע�����</STRONG></TD>
                  </TR>
                  <TR>
                    <TD bgColor="#ffffff" height="26"><TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
                        <TBODY>
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">�벻Ҫ�ظ�����������ͬ��֤�飬����������ͬ��֤�齫�ܾ���ˣ� </TD>
                          </TR>
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">��ȷ������֤���׼ȷ�ԣ��Ϸ��ԣ��������Ը���<%=KS.Setting(1)%>���е��κ����Ρ�</TD>
                          </TR>
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">�����ܵ���������֤����Ϣ��</TD>
                          </TR>
                        </TBODY>
                    </TABLE></TD>
                  </TR>
            </table>
		        <%
  End Sub
  
  Sub DoSave()
  
            Dim fobj:Set FObj = New UpFileClass
			FObj.GetData
            Dim MaxFileSize:MaxFileSize = 500   '�趨�ļ��ϴ�����ֽ���
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(9994,KSUser.UserName)
			Call KS.CreateListFolder(FormPath) 
			

				 Title=KS.LoseHtml(Fobj.Form("Title"))
				  If Title="" Then
				    Response.Write "<script>alert('��û����������֤������!');history.back();</script>"
				    Exit Sub
				  End IF
				 
				 Fzjg=KS.DelSql(Fobj.Form("Fzjg"))
				 Jzrq=Fobj.Form("Jzrq")
				 Sxrq=Fobj.Form("Sxrq")
				 If Not IsDate(jzrq)  Or Not IsDate(Sxrq) Then Call KS.AlertHistory("���ڲ���ȷ!",-1):response.End()
			
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now))
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����������ϴ��Ĵ�С\n�����ϴ� " & MaxFileSize & " KB���ļ�\n",-1):response.End()
			End Select
			If ReturnValue="" and KS.ChkClng(Fobj.Form("ID"))=0 Then Call KS.AlertHistory("��û���ϴ�֤����Ƭ\n",-1):response.End()

				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_EnterPriseZS Where UserName='" & KSUser.UserName & "' and ID=" & KS.ChkClng(Fobj.Form("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  If KS.SSetting(20)="1" Then
				  RSObj("Status")=0
				  Else
				  RSObj("Status")=1
				  End If
				  RSObj("Adddate")=Now
				 End If
				  RSObj("UserName")=KSUser.UserName
				  RSObj("Title")=Title
				  RSObj("Fzjg")=Fzjg
				  RSObj("Jzrq")=Jzrq
				  RSObj("Sxrq")=Sxrq
				  If ReturnValue<>"" then
				  RSObj("PhotoUrl")=ReturnValue
				  end if
				RSObj.Update
				 RSObj.Close:Set RSObj=Nothing
				 
               If KS.ChkClng(Fobj.Form("ID"))=0 Then
			     Set Fobj=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"�ϴ�������֤��:" & Title & "!",204)
				 Response.Write "<script>if (confirm('����֤�鷢���ɹ�������������?')){location.href='?Action=Add';}else{location.href='User_EnterPriseZS.asp';}</script>"
			   Else
			     Set Fobj=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"����������֤��:" & Title & "!",204)
				 Response.Write "<script>alert('����֤���޸ĳɹ�!');location.href='User_EnterPriseZS.asp';</script>"
			   End If
  End Sub
  
End Class
%> 

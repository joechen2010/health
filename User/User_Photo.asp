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
Set KSCls = New User_Photo
KSCls.Kesion()
Set KSCls = Nothing

Class User_Photo
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather,PhotoUrls,descript
		Private XCID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
		Private Sub Class_Initialize()
		  MaxPerPage =20
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
		ElseIf KS.SSetting(0)=0 Then
		 Call KS.Alert("�Բ��𣬱�վ�رո��˿ռ书�ܣ�","")
		 Exit Sub
		ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		 Call KS.Alert("�㲻�ԣ��㻹û�п�ͨ�ռ书�ܣ�","User_Blog.asp")
		 Exit Sub
		ElseIf Conn.Execute("Select status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>alert('�Բ�����Ŀռ仹û��ͨ����˻�������');history.back();</script>"
			response.end
		End If

		Call KSUser.Head()
		Call KSUser.InnerLocation("�ҵ����")
		KSUser.CheckPowerAndDie("s05")
		%>
		<div class="tabs">	
		   <ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">�ҵ����</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=1")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="?Status=0">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=0")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">�������(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=2")(0)%></span>)</a></li>
			</ul>
        </div>
			 <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="User_Photo.asp?Action=Add"><span style="font-size:14px;color:#ff3300">�ϴ���Ƭ</span></a>
			  <img src="images/fav.gif" width="20" align="absmiddle"><a href="User_Photo.asp?Action=Createxc"><span style="font-size:14px;color:#ff3300">�������</span></a>
			 </div>

		<%

			Select Case KS.S("Action")
			 Case "Del"
			  Call Delxc()
			 Case "Delzp"
			  Call Delzp()
			 Case "Editzp"
			  Call Editzp()
			 Case "Add"
			  Call Addzp()
			 Case "AddSave"
			  Call AddSave()
			 Case "EditSave"
			  Call EditSave()
			 Case "ViewZP"
			  Call ViewZP()
			 Case "Editxc","Createxc"
			  Call Managexc()
			 Case "photoxcsave"
			  Call photoxcsave()
			 Case Else
			  Call PhotoxcList()
			End Select
	   End Sub
	   '�鿴��Ƭ
	   Sub ViewZP()
	    Dim title
	    Dim xcid:xcid=KS.Chkclng(KS.S("XCID"))
	    Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "select xcname from KS_Photoxc WHERE ID=" & XCID,CONN,1,1
		if rs.Eof And RS.Bof Then 
		 rs.close:set rs=nothing
		 response.write "<script>alert('�������ݳ���');history.back();</script>"
		 response.end
		end if
		title=rs(0)
		rs.close
		Call KSUser.InnerLocation("�鿴��Ƭ")
	  			  %>
			   
	   		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
            <tr class="title">
              <td align=center colspan=5><%=Title%></td>
            </tr>
			<%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
			 rs.open "select * from KS_PhotoZP where xcid=" & xcid,conn,1,1
			if rs.eof and rs.bof then
			  response.write "<tr class='tdbg'><td  height='30' colspan='5'>�������û����Ƭ����<a href=""?action=Add&xcid=" & xcid &""">�ϴ�</a>��</td></tr>"
			else
			 				  MaxPerPage =5
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
									Call showzplist(xcid)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showzplist(xcid)
									Else
										CurrentPage = 1
										Call showzplist(xcid)
									End If
								End If
        end if%>
      </table>
	  <div style="padding-right:30px">
	  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
	  </div>
<%End Sub
sub showzplist(xcid)
%>
    <script type="text/javascript" src="../ks_inc/highslide/highslide.js"></script>
    <link href="../ks_inc/highslide/highslide.css" type=text/css rel=stylesheet>
	<script type="text/javascript">
		hs.graphicsDir = '/ks_inc/highslide/graphics/';
		hs.transitions = ['expand', 'crossfade'];
		hs.wrapperClassName = 'dark borderless floating-caption';
		hs.fadeInOut = true;
		hs.dimmingOpacity = .75;
		
		if (hs.addSlideshow) hs.addSlideshow({
			interval: 5000,
			repeat: false,
			useControls: true,
			fixedControls: 'fit',
			overlayOptions: {
				opacity: .6,
				position: 'bottom center',
				hideOnMouseOut: true
			}
		});
	</script>
<%
     Dim I
    Response.Write "<FORM Action=""?Action=Delzp"" name=""myform"" method=""post"">"
			 do while not rs.eof
			 %>
			<tr class="tdbg"> 
            <td width="21%" rowspan="5">
			<table border="0" align="center" cellpadding="2" cellspacing="1" class="border">
                <tr> 
                  <td><a href="<%=rs("photourl")%>" class="highslide" onClick="return hs.expand(this)"  title="<%=rs("title")%>"><img src="<%=rs("photourl")%>" width="85" height="100" border="0"></a>
                  </td>
                </tr>
              </table></td>
            <td width="12%" class="tdbg"><div align="center"><strong>��Ƭ���ƣ�</strong></div></td>
            <td width="40%" class="tdbg"><font style="font-size:14px"><strong><%=rs("title")%></strong></font></td>
            <td width="10%"><div align="center">���������</div></td>
            <td width="17%"><%=rs("hits")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">�������ڣ�</div></td>
            <td><%=rs("adddate")%></td>
            <td><div align="center">ͼƬ��С��</div></td>
            <td><%=rs("photosize")%>byte</td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">��Ƭ��ַ��</div></td>
            <td colspan="3"><%=rs("photourl")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">��Ƭ������</div></td>
            <td colspan="3"><%=rs("descript")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">������᣺</div></td>
            <td><%=conn.execute("select xcname from ks_photoxc where id=" & xcid)(0)%></td>
            <td colspan="2" height="28"><div align="center"><a href="?Action=Editzp&Id=<%=rs("id")%>" class="box">�޸�</a> <a href="?id=<%=rs("id")%>&Action=Delzp" onClick="{if(confirm('ȷ��ɾ������Ƭ��')){return true;}return false;}" class="box">ɾ��</a> 
                <INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
              </div></td>
          </tr>
          <tr> 
            <td colspan="5" height="3" class="splittd">&nbsp;</td>
          </tr>
			<% rs.movenext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
			 loop
		 %>
		 <tr class="tdbg">
		   <td colspan="5" align="right">
		  								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ�б�ҳ��ʾ��������Ƭ&nbsp;<INPUT class="button" onClick="return(confirm('ȷ��ɾ��ѡ�е���Ƭ��?'));" type=submit value=ɾ��ѡ������Ƭ name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        </td>
		 </tr>
		 </form>
		 <%
	   End Sub
	    '��ᣬ��ӣ��޸�
	   Sub Managexc()
	    Dim xcname,ClassID,Descript,PhotoUrl,PassWord,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,Flag,ListLogNum
		Dim ID:ID=KS.ChkCLng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Photoxc Where ID=" & ID,conn,1,1
		If Not RS.EOF Then
		Call KSUser.InnerLocation("�޸����")
		 xcname=RS("xcname")
		 ClassID=RS("ClassID")
		 Descript=RS("Descript")
		 flag=RS("Flag")
		 PhotoUrl=RS("PhotoUrl")
		 PassWord=RS("PassWord")
		 OpStr="OK�ˣ�ȷ���޸�":TipStr="�� �� �� �� �� ��"
		Else
		 Call KSUser.InnerLocation("�������")
		 xcname=FormatDatetime(Now,2)
		 ClassID="0"
		 flag="1"
		 PhotoUrl=""
		 OpStr="OK�ˣ���������":TipStr="�� �� �� �� �� ��"
		End if
		RS.Close:Set RS=Nothing
	    %>
		<script>
		 function CheckForm()
		 {
		  if (document.myform.xcname.value=='')
		  {
		   alert('�������������!');
		   document.myform.xcname.focus();
		   return false;
		  }
		  if (document.myform.ClassID.value=='0')
		  {
		   alert('��ѡ���������!');
		   document.myform.ClassID.focus();
		   return false;
		  }
		  return true;
		 }

		</script>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
          <form  action="User_Photo.asp?Action=photoxcsave&id=<%=id%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
            <tr class="title">
              <td colspan=2 align=center><%=TipStr%></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>������ƣ�</strong><br>
              ���������ȡ�����ʵ����ơ�
              </div></td>
              <td> ��
                  <input class="textbox" name="xcname" type="text" id="xcname" style="width:250px; " value="<%=xcname%>" maxlength="100" />
              <span style="color: #FF0000">*</span></td>
            </tr>
<tr class="tdbg">
              <td width="24%"  height="25" align="center"><div align="left"><strong>�����ࣺ</strong><br>
      �����࣬�Ա�������</div></td>
              <td width="76%">��
                  <select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-��ѡ�����-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
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
                  <td height="30"><div align="left"><strong>�Ƿ񹫿���</strong><br>
                  ��������Ϊֻ��Ȩ�޵��û����������</div></td>
                  <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
                   <tr>
                     <td width="50%" align="left">&nbsp;
                       <select style="width:160px" onChange="if(this.options[selectedIndex].value=='3'){document.myform.all.mmtt.style.display='block';}else{document.myform.all.mmtt.style.display='none';}"  name="flag">
                      <option value="1"<%if flag="1" then response.write " selected"%>>��ȫ����</option>
                      <option value="2"<%if flag="2" then response.write " selected"%>>��Ա����</option>
                      <option value="3"<%if flag="3" then response.write " selected"%>>���빲��</option>
                      <option value="4"<%if flag="4" then response.write " selected"%>>��˽���</option>
                    </select></td>
                   <td width="50%"><span class=child id=mmtt name="mmtt" <%if flag<>3 then%>style="display:none;"<%end if%>>���룺<input type="password" name="password" style="width:160px" maxlength="16" value="<%=password%>" size="20"></span>                  </td>
                  </tr>
                  </table>                  </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>�����棺</strong><br>
                  �������ϴ���ϲ����ͼƬ��Ϊ���ķ��档</div></td>
              <td>��
                  <input class="textbox" name="PhotoUrl" type="text" id="PhotoUrl" style="width:230px; " value="<%=PhotoUrl%>" />             
                  &nbsp;ֻ֧��jpg��gif��png��С��50k��Ĭ�ϳߴ�Ϊ85*100
				  <div>
                  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9998' frameborder="0" align="center" width='94%' height='30' scrolling="no"></iframe>
				  </div>
				  </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>�����ܣ�</strong></span></div>
                <br>
                ���ڴ����ļ�Ҫ����˵����</td>
              <td>��
                  
                  <textarea name="Descript" id="Descript" cols=50 rows=6><%=Descript%></textarea>              </td>
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
	   '�������
	   Sub photoxcsave()
	     Dim xcname:xcname=KS.S("xcname")
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Descript:Descript=KS.S("Descript")
		 Dim Flag:Flag=KS.S("Flag")
		 Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim ID:ID=KS.Chkclng(KS.S("id"))
		 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/images/user/nopic.gif"
		 If xcname="" Then Response.Write "<script>alert('�������������!');history.back();</script>"
		 If ClassID=0 Then Response.Write "<script>alert('��ѡ���������!');history.back();</script>"
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photoxc Where id=" & id ,conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now
			if ks.SSetting(4)=1 then
			RS("Status")=0 '��Ϊ����
			else
			RS("Status")=1 '��Ϊ����
			end if
		 End If
		    RS("UserName")=KSUser.UserName
		    RS("xcname")=xcname
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Flag")=Flag
			RS("Password")=PassWord
			RS("PhotoUrl")=PhotoUrl
		  RS.Update
		  RS.MoveLast
		  ID=rs("id")
		  RS.Close:Set RS=Nothing
		  If KS.Chkclng(KS.S("id"))=0 Then
		   Call KS.FileAssociation(1028,ID,PhotoUrl,0)
		   Call KSUser.AddLog(KSUser.UserName,"���������!����: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """ target=""_blank"">�鿴</a>",104)
		   Response.Write "<script>alert('��ϲ!��ᴴ���ɹ�,�����ϴ���Ƭ');location.href='User_Photo.asp?action=Add&xcid=" & id &"';</script>"
		  Else
		   Call KS.FileAssociation(1028,ID,PhotoUrl,1)
		   Call KSUser.AddLog(KSUser.UserName,"�޸������!����: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """  target=""_blank"">�鿴</a>",104)
		   Response.Write "<script>alert('����޸ĳɹ�!');location.href='User_Photo.asp';</script>"
		  End If
	   End Sub


	  
	   '����б�
	   Sub PhotoxcList()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									IF KS.S("status")<>"" Then
									  Param=Param & " And status=" & KS.ChkClng(KS.S("status"))
									End if
									
									
									'If KS.S("XCID")<>"" And KS.S("XCID")<>"0" Then Param=Param & " And XCID=" & KS.ChkClng(KS.S("XCID")) & ""
									Dim Sql:sql = "select * from KS_Photoxc "& Param &" order by AddDate DESC"


								    Call KSUser.InnerLocation("��������б�")
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1">
                                                <tr class="Title">
                                                  <td colspan="6" height="22" align="center">�� �� �� ��</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>����û�д������!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
											If CurrentPage < 1 Then
												CurrentPage = 1
											End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowXC
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowXC
									Else
										CurrentPage = 1
										Call ShowXC
									End If
								End If
				End If
     %>                      
                        </table>
		  <%
  End Sub
  
  Sub ShowXC()
     Dim I,K
   Do While Not RS.Eof
         %>
           <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		   <%
		   For K=1 To 4
		   %>
            <td width="25%" height="22" align="center">
									  <table width=154 height=185 border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" id=AutoNumber2 style="BORDER-COLLAPSE: collapse">
										  <td width=123 height=185>
											<table id=AutoNumber3 style="BORDER-COLLAPSE: collapse" borderColor=#b2b2b2 height=179 cellSpacing=0 cellPadding=0 width="117%" border=0>
											  <tr>
												<td width="100%" height=179>
												  <table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width="99%" border=0>
													<tr>
													  <td align=middle width="100%" height=22><B><a href="?xcid=<%=rs("id")%>&action=ViewZP"><%=ks.gottopic(rs("xcname"),18)%></a></B><%select case rs("status")
													     case 1:response.write "[����]"
														 case 2:response.write "<font color=blue>[����]</font>"
														 case 0:response.write "<font color=red>[δ��]</font>"
														end select
														%>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%">
														<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0>
														  <tr>
															<td background="images/pic.gif" width="136" height="106" valign="top"><a href="?xcid=<%=rs("id")%>&action=ViewZP"><img style="margin-left:6px;margin-top:5px" src="<%=rs("photourl")%>" width="120" height="90" border=0></a></td>
														  </tr>
														</table>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><%=rs("xps")%>��/<%=rs("hits")%>��</td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><a href="?Action=Editxc&id=<%=rs("id")%>">�޸�</a>&nbsp;<a href="?Action=Del&id=<%=rs("id")%>" onClick="return(confirm('ɾ����Ὣɾ����������������Ƭ��ȷ��ɾ����'))">ɾ��</a>&nbsp;
													  <% select case rs("flag")
													      case 1
													       response.write "<font color=red>[����]</font>"
														  case 2
													       response.write "<font color=red>[��Ա]</font>"
														  case 3
													       response.write "<font color=red>[����]</font>"
														  case 4
													       response.write "<font color=red>[��˽]</font>"
														 end select
													%>
													  </td>
													</tr>
												  </table>
												</td>
											  </tr>
											</table>
										  </td>
										</tr>
			  </table>
			 </td>
                       
					                  <%
							RS.MoveNext
							I=I+1
					  If I >= MaxPerPage Or RS.Eof Then Exit For
				  Next
			      do While K<4 
				   response.write "<td width=""25%""></td>"
				   k=k+1
				  Loop%>
		    </tr>
				 <%
					  If I >= MaxPerPage Or RS.Eof Then Exit do
	   Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=6 valign=top align="right">
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								</tr>
								<% 
  End Sub
  'ɾ�����
  Sub Delxc()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ�������!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_Photoxc Where ID In(" & ID & ")")
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where xcid in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid=" & rs("id"))
	   KS.DeleteFile(rs("photourl"))
	   rs.movenext
	   loop
	end if
	Conn.execute("delete from ks_photozp where xcid in(" & id& ")")
	Conn.execute("delete from ks_uploadfiles where channelid=1028 and infoid in(" & id& ")")
	rs.close:set rs=nothing
	Call KSUser.AddLog(KSUser.UserName,"ɾ����������!",104)
	Response.Redirect ComeUrl
  End Sub
  'ɾ����Ƭ
  Sub Delzp()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ������Ƭ!",ComeUrl):Response.End
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where id in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   KS.DeleteFile(rs("photourl"))
	   Conn.execute("update ks_photoxc set xps=xps-1 where id=" & rs("xcid"))
	   rs.movenext
	   loop
	end if
	Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid in(" & id& ")")
	Conn.execute("delete from ks_photozp where id in(" & id& ")")
	Call KSUser.AddLog(KSUser.UserName,"ɾ������Ƭ����!",104)
	rs.close:set rs=nothing
	Response.Redirect ComeUrl
  End Sub
  '�ϴ���Ƭ
  Sub Addzp()
        Call KSUser.InnerLocation("�ϴ���Ƭ")
		  adddate=now:XCID=KS.ChkCLng(KS.S("XCID")):UserName=KSUser.RealName
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.XCID.value=="0") 
				  {
					alert("��ѡ��������ᣡ");
					document.myform.XCID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("��������Ƭ���ƣ�");
					document.myform.Title.focus();
					return false;
				  }		
				 return true;  
				}
				</script>
				<script>  
			var FFextraHeight = 0;
			 if(window.navigator.userAgent.indexOf("Firefox")>=1)
			 {
			  FFextraHeight = 16;
			  }
			 function ReSizeiFrame(iframe)
			 {
			   if(iframe && !window.opera)
			   {
				 iframe.style.display = "block";
				  if(iframe.contentDocument && iframe.contentDocument.body.offsetHeight)
				  {
					iframe.height = iframe.contentDocument.body.offsetHeight + FFextraHeight;
				  }
				  else if (iframe.Document && iframe.Document.body.scrollHeight)
				  {
					iframe.height = iframe.Document.body.scrollHeight;
				  }
			   }
			 }
			function init()
			 {
			   ReSizeiFrame(document.getElementById('UpPhotoFrame'));
			 }
			
			</script>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=AddSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>�� �� �� Ƭ</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>ѡ����᣺</span></td>
                       <td width="88%"><select class="textbox" size='1' name='XCID' style="width:150">
                             <option value="0">-��ѡ�����-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_Photoxc where username='" & KSUser.Username & "' order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>��Ƭ���ƣ�</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*
                                        <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" />
                                        </span></td>
                    </tr>
								<tr class="tdbg">
								  <td height="20" align="center">��ƬԤ����</td>
								  <td align='center' id="viewarea">
								     
								</td>
				    </tr>
					
								<tr class="tdbg">
                                   <td height="250" align="center"><span>�ϴ���Ƭ��</span></td>
                                   <td align="center"><iframe onload="ReSizeiFrame(this)" onreadystatechange="ReSizeiFrame(this)" id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9997' frameborder="0" scrolling="auto" width='100%' height='92%'></iframe></td>
							  </tr>							 
								<tr class="tdbg">
                                   <td height="25" align="center"><span>��Ƭ���ܣ�</span></td>
                                  <td><textarea class="textbox" style="height:50px" name="Descript" cols="70" rows="5"></textarea></td>
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
    '�༭��Ƭ
  Sub Editzp()
        Call KSUser.InnerLocation("�༭��Ƭ")
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_PhotoZp Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_A_RS_Obj.Eof Then
		     XCID  = KS_A_RS_Obj("XCID")
			 Title    = KS_A_RS_Obj("Title")
			 UserName   = KS_A_RS_Obj("UserName")
			 descript = ks_a_rs_obj("descript")
			 PhotoUrlS  = KS_A_RS_Obj("PhotoUrl")
		   End If
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.XCID.value=="0") 
				  {
					alert("��ѡ��������ᣡ");
					document.myform.XCID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("��������Ƭ���ƣ�");
					document.myform.Title.focus();
					return false;
				  }		
				 return true;  
				}
				
				</script>
				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=EditSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>�� �� �� Ƭ</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>ѡ����᣺</span></td>
                       <td width="88%"><select class="textbox" size='1' name='XCID' style="width:150">
                             <option value="0">-��ѡ�����-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_Photoxc order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>��Ƭ���ƣ�</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*
                                        <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" value="<%=photourls%>"/>
                                        </span></td>
                    </tr>
								<tr class="tdbg">
								  <td height="20" align="center">��ƬԤ����</td>
								  <td id="viewarea">
								    <table style='BORDER-COLLAPSE: collapse' borderColor='#c0c0c0' cellSpacing='1' cellPadding='2' border='1'><tr><td align='center' width='83' height='100' bgcolor='#ffffff'><img name='view1' width='83' height='100' src='<%=Photourls%>' title='��ƬԤ��'></td></tr></table> <input class="button" type='button' name='Submit3' value='ѡ����Ƭ��ַ...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&ChannelID=9997',500,360,window,document.myform.PhotoUrls);" />
								</td>
				    </tr>
														 
								<tr class="tdbg">
                                   <td height="25" align="center"><span>��Ƭ���ܣ�</span></td>
                                  <td><textarea class="textbox" style="height:50px" name="Descript" cols="70" rows="5"><%=DESCRIPT%></textarea></td>
							  </tr>							 
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,�������� " />
                      <input type="reset" name="Submit2"   class="Button" onClick="javascript:history.back()" value=" ȡ �� " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub

   Sub EditSave()
    Dim RSObj,Descript,PhotoUrlArr,i
                 XCID=KS.ChkClng(KS.S("XCID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('��û���ϴ���Ƭ!');history.back();</script>"
				    Exit Sub
				  End IF
				  on error resume next
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_PhotoZP Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				  RSObj("Title")=Title
				  RSObj("XCID")=XCID
				  RSObj("PhotoUrl")=PhotoUrls
				  RSObj("Descript")=Descript
				  RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(replace(PhotoUrls,ks.getdomain,ks.setting(3))))
				RSObj.Update
				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(1029,KS.ChkClng(KS.S("ID")),PhotoUrls,1)
				 Call KSUser.AddLog(KSUser.UserName,"�޸�����Ƭ����! <a href=""" & PhotoUrls & """ target=""_blank"">�鿴</a>",104)
				 Response.Write "<script>alert('��Ƭ�޸ĳɹ�!');location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';</script>"
  End Sub
  
  Sub AddSave()
    Dim RSObj,Descript,PhotoUrlArr,i,UpFiles
                 XCID=KS.ChkClng(KS.S("XCID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('��û���ϴ���Ƭ!');history.back();</script>"
				    Exit Sub
				  End IF
				PhotoUrlArr=Split(PhotoUrls,"|")
				 
				  If XCID=0 Then
				    Response.Write "<script>alert('��û��ѡ�����!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('��û��������Ƭ����!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_PhotoZP",Conn,1,3
				 For I=0 to ubound(PhotoUrlArr)
			    	RSObj.AddNew
					 RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(Replace(PhotoUrlArr(I),KS.GetDomain,KS.Setting(3))))
				     RSObj("Title")=Title
				     RSObj("XCID")=XCID
					 RSObj("UserName")=KSUser.UserName
					 RSObj("PhotoUrl")=PhotoUrlArr(I)
					 RSObj("Adddate")=Now
					 RSObj("Descript")=Descript
				   RSObj.Update
				   RSObj.MoveLast
				   Call KS.FileAssociation(1029,RSObj("ID"),PhotoUrlArr(i),0)
				 Next
				 RSObj.Close
				 RSObj.Open "Select Top 1 PhotoUrl From KS_PhotoXC Where ID=" & xcid,conn,1,3
				 If Not RSObj.Eof Then
				    If Instr(lcase(RSObj(0)),"nopic.gif")>0 then
					  RSObj(0)=PhotoUrlArr(0)
					  RSObj.Update
					end if
				 End If
				 RSObj.Close
				 Set RSObj=Nothing
				 
				 
				 Conn.Execute("update KS_Photoxc set xps=xps+" & Ubound(PhotoUrlArr)+1 & " where id=" & xcid)
				 Call KSUser.AddLog(KSUser.UserName,"�ϴ���" & Ubound(PhotoUrlArr)+1 & "����Ƭ�����! <a href=""../space/?" & KSUser.UserName & "/showalbum/" & xcid & """ target=""_blank"">�鿴</a>",104)
				 Response.Write "<script>if (confirm('��Ƭ����ɹ��������ϴ���?')){location.href='User_Photo.asp?Action=Add';}else{location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';}</script>"
  End Sub

End Class
%> 

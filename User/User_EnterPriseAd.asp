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
Set KSCls = New EnterPriseADCls
KSCls.Kesion()
Set KSCls = Nothing

Class EnterPriseADCls
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,Selbutton,LoginTF,Verific,PhotoUrl,bigclassid,smallclassid,flag
		Private F_B_Arr,F_V_Arr,ClassID,Title,ADWZ,URL,datatimed,Action,I,Adtype
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
		Call KSUser.InnerLocation("�ؼ��ʹ��")
		KSUser.CheckPowerAndDie("s12")
		
		
		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">�ҷ����Ĺ��(<span class="red"><%=conn.execute("select count(id) from KS_EnterPriseAD where username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">�����(<span class="red"><%=conn.execute("select count(id) from KS_EnterPriseAD where status=1 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="?Status=0">�����(<span class="red"><%=conn.execute("select count(id) from KS_EnterPriseAD where status=0 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
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
									  IF KS.S("Flag")=1 Then Param=Param & " And ADWZ like '%" & KS.S("KeyWord") & "%'"
									End if
									Dim Sql:sql = "select * from KS_EnterPriseAD " & Param &" order by ID DESC"

								 
								  %>
								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="?Action=Add"><font color=red>���ؼ��ʹ���ύ</font></a></div>

				                     <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                                        <tr class="Title">
                                                  <td width="6%" height="22" align="center">ѡ��</td>
                                                  <td width="31%" height="22" align="center">�������</td>
                                                  <td width="10%" height="22" align="center"> ����λ��</td>
                                                  <td width="15%" height="22" align="center"> ��������</td>
												  <td width="16%" height="22" align="center">��ʼʱ��</td>
												  <td width="10%" height="22" align="center">״̬</td>
                                                  <td height="22" align="center" nowrap>�������</td>
                                        </tr>
                                           
                                      <%
								 Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>�Ҳ����κιؼ��ʹ��!</td></tr>"
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
			 <INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ѡ�����йؼ��ʹ�� <input value="ɾ��ѡ��" class="button" Click="return(confirm('ȷ��ɾ��ѡ�е��Ŷӳ�Ա��?'));" type=submit> 
			</form>
			</td>
			<td align="right">
			<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
			 </td>
			</tr>
		   </table>
		 </td>
	   </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowContent()
    Response.Write "<FORM Action=""?Action=Del"" name=""myform"" method=""post"">"

	     dim i,k
	     do while not rs.eof
		  
		  %>
                   <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                        <td width="5%" height="20" align="center">
						  <INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
						</td>
                        <td align="left"><a href="?Action=Edit&id=<%=rs("id")%>" class="link3"><%=KS.GotTopic(trim(RS("title")),45)%></a></td>
						<td align="center">
						 <%if rs("adwz")="1" then
						  response.write "��Ʒ��"
						  else
						  response.write "��ҵ��ȫ"
						  end if
						 %>
						</td>
                        <td align="center">
						<%=rs("datatimed")%> ��
						</td>
                        <td align="center"><%=formatdatetime(rs("beginDate"),2)%></td>
                        <td align="center"><%
						if rs("status")=1 then
						 response.write "�����"
						else
						 response.write "<font color=red>δ���</font>"
						end if
						%></td>
                        <td align="center">
						<a href="?id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>" class="link3">�޸�</a> <a href="?action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('ȷ��ɾ��������?'))" class="link3">ɾ��</a>
										
						</td>
                     </tr>
					   <tr><td colspan=6 background='images/line.gif'></td></tr>
			<%
            rs.movenext
			k=k+1
		  if k>=MaxPerPage then exit do
		 loop

  End Sub
  'ɾ������
  Sub ArticleDel()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("��û��ѡ��Ҫɾ�����Ŷӳ�Ա!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_EnterPriseAD Where UserName='" & KSUser.UserName & "' And ID In(" & ID & ")")
	if ComeUrl="" then
	Response.Redirect("../index.asp")
	else
	Response.Redirect ComeUrl
	end if
  End Sub

  '�������
  Sub DoAdd()
        Call KSUser.InnerLocation("�ؼ��ʹ���ύ")
		  on error resume next

  		if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_EnterPriseAD Where UserName='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
			 Title    = RSObj("Title")
			 ADType = RSObj("ADType")
			 BigClassID=RSObj("BigClassID")
			 SmallClassID=RSObj("SmallClassID")
			 URL   = RSObj("URL")
			 ADWZ  = RSObj("ADWZ")
			 datatimed=RSObj("datatimed")
			 PhotoUrl  = RSObj("PhotoUrl")
			 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/Images/NoPhoto.gif"
			 flag=true
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		 PhotoUrl="/images/Nophoto.gif"
		 ADWZ="1"
		 URL="http://"
		 flag=false
		End If
		%>
		<script language="javascript" src="../ks_inc/popcalendar.js"></script>

		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("�����������ƣ�");
					document.myform.Title.focus();
					return false;
				  }	
				
				if (document.myform.URL.value=="")
				  {
					alert("���������ַ��");
					document.myform.URL.focus();
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
							   response.write "�޸Ĺؼ��ʹ��"
							   Else
							    response.write "�ؼ��ʹ���ύ"
							   End iF
							  %>                         </td>
					</tr>
                    
                      <tr class="tdbg">
                        <td  height="25" align="center">Ͷ�����ͣ�</td>
                        <td>��
                          <input name="Adtype" type="radio" value="1" onClick="document.all.SmallClassID.disabled=true;">                                 
                          ����
                          <input name="AdType" type="radio" onClick="document.all.SmallClassID.disabled=false;" value="2">        
                          С��</td><td width="36%" rowspan="7" align="center">
                          <img src="<%=photourl%>" width="250" height="120">							  </td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">��ҵ���</td>
                        <td>��
                          <%
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0; 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		  <select class="face" name="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		   <option value="">--��ѡ����ҵ����--</option>
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		  rsb.close:set rsb=nothing
		  response.write "<script>alert('���ȵ���̨�����ҵ����!');history.back();</script>"
		  response.end
		else
		    Dim N
		    do while not rsb.eof
			          N=N+1
					  If N=1 and flag=false Then BigClassID=rsb("id")
					  If BigClassID=rsb("id") then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" name="SmallClassID"<%if adtype="1" then response.write " disabled"%>>
				  <option value="" selected>--��ѡ����ҵ����--</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="& KS.ChkClng(BigClassID)&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if SmallClassID=rsss("id") then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select></td>
                      </tr>
					 
                      <tr class="tdbg" style="display:none">
                                      <td  height="25" align="center"><span>Ͷ��λ�ã�</span></td>
                                      <td height="25">��
                                        <input name="ADWZ" type="radio" value="1"<%if trim(ADWZ)="1" then response.write " checked"%>/>��ҵ��ȫ
                                        <input name="ADWZ" type="radio" value="2"<%if trim(ADWZ)="2" then response.write " checked"%>/>��Ʒ��
										
                                       </td>
                              </tr>
                              <tr class="tdbg">
                                <td height="25" align="center">Ͷ��ʱ�䣺</td>
                                <td height="25">��
                            <select name="datatimed" id="datatimed">
                                   <option value="" selected>��ѡ��...</option>
                                   <option value="7"<%if datatimed="7" then response.write " selected"%>>һ��</option>
                                   <option value="15"<%if datatimed="15" then response.write " selected"%>>�����</option>
                                   <option value="30"<%if datatimed="30" then response.write " selected"%>>һ����</option>
                                   <option value="60"<%if datatimed="60" then response.write " selected"%>>������</option>
                                   <option value="90"<%if datatimed="90" then response.write " selected"%>>������</option>
                                   <option value="180"<%if datatimed="180" then response.write " selected"%>>����</option>
                                   <option value="365"<%if datatimed="365" then response.write " selected"%>>һ��</option>
                                   <option value="730"<%if datatimed="730" then response.write " selected"%>>����</option>
                               </select></td>
                              </tr>
							  <tr class="tdbg">
								   <td width="12%"  height="25" align="center"><span>������ƣ�</span></td>
									  <td width="52%"> ��
												<input class="textbox" name="Title" type="text" style="width:250px; " value="<%=Title%>" maxlength="100" />
												  <span style="color: #FF0000">*</span></td>
							  </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>���ӵ�ַ��</span></td>
                                      <td height="25">��
                                        <input name="URL" class="textbox" type="text" id="URL" style="width:250px; " value="<%=URL%>" maxlength="30" />
                                        <span style="color: #FF0000">*</span></td>
                              </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>ͼƬ��ַ��</span></td>
                        <td> ��
                               <input type="file" name="photourl" size="40">
                          <span style="color: #FF0000">*</span> <br>
                          �� <font color=red>˵����ֻ֧��JPG��GIF��PNG��ʽͼƬ��������300K,��С650*90</font></td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">�û�����</td>
                        <td>��
                           <input name="UserName" class="textbox" type="text" readonly style="width:100px; " value="<%=KSUser.UserName%>" maxlength="30" /></td>
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
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">��ȷ�����Ĺ�潡����������ɫ��Ϣ��ȷ����ʵ�ԣ��Ϸ��ԣ��������Ը���<%=KS.Setting(1)%>���е��κ����Ρ�</TD>
                          </TR>
                          <TR>
                            <TD height="21"><IMG height="8" src="images/expand.gif" width="8">�ύ����ҵ�����뾭������Ա��˺������Ч����Чʱ�������ʱ��Ϊ׼��</TD>
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
            Dim MaxFileSize:MaxFileSize = 300   '�趨�ļ��ϴ�����ֽ���
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(9994,KSUser.UserName)
			Call KS.CreateListFolder(FormPath) 
			

				 Title=KS.LoseHtml(Fobj.Form("Title"))
				  If Title="" Then
				    Response.Write "<script>alert('��û������������!');history.back();</script>"
				    Exit Sub
				  End IF
				 
				 Adtype=KS.ChkClng(Fobj.Form("Adtype"))
				 BigClassID=KS.ChkCLng(Fobj.Form("ClassID"))
				 SmallClassID=KS.ChkCLng(Fobj.Form("SmallClassID"))
				 
				 ADWZ=KS.DelSql(Fobj.Form("ADWZ"))
				 URL=KS.DelSql(Fobj.Form("URL"))
				 ADWZ=KS.ChkClng(Fobj.Form("ADWZ"))
				 datatimed=KS.ChkClng(Fobj.Form("datatimed"))
			
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now))
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowFileExtStr + "\n",-1):response.end
	          Case "errsize"  Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����������ϴ��Ĵ�С\n�����ϴ� " & MaxFileSize & " KB���ļ�\n",-1):response.End()
			End Select
			
			If ReturnValue="" and KS.ChkClng(Fobj.Form("ID"))=0 then
			 Call KS.AlertHistory("���ͼƬ�����ϴ�!",-1)
			 Response.End()
			End If

				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_EnterPriseAD Where UserName='" & KSUser.UserName & "' and ID=" & KS.ChkClng(Fobj.Form("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Status")=0
				  RSObj("BeginDate")=Now
				 End If
				  RSObj("UserName")=KSUser.UserName
				  RSObj("Title")=Title
				  RSObj("ADType")=ADType
				  RSObj("URL")=URL
				  RSObj("ADWZ")=ADWZ
				  RSObj("BigClassID")=BigClassID
				  RSObj("SmallClassID")=SmallClassID
				  RSObj("datatimed")=datatimed
				  If ReturnValue<>"" then
				  RSObj("PhotoUrl")=ReturnValue
				  end if
				  
				RSObj.Update
				 RSObj.Close:Set RSObj=Nothing
				 
               If KS.ChkClng(Fobj.Form("ID"))=0 Then
			     Set Fobj=Nothing
				 Response.Write "<script>if (confirm('�ؼ��ʹ���ύ�ɹ��������ύ��?')){location.href='?Action=Add';}else{location.href='User_EnterPriseAD.asp';}</script>"
			   Else
			     Set Fobj=Nothing
				 Response.Write "<script>alert('�ؼ��ʹ���޸ĳɹ�!');location.href='User_EnterPriseAD.asp';</script>"
			   End If
  End Sub
  
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_EnterpriseAD
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterpriseAD
        Private KS,typeflag
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
		Private ComeUrl,Selbutton,LoginTF,Verific,PhotoUrl,bigclassid,smallclassid,flag
		Private ClassID,Title,ADWZ,URL,datatimed,Adtype,status,begindate,username

        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10013") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		typeflag=ks.chkclng(ks.g("type"))
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  If KS.G("Action")<>"View" then
			  .Write "<div class='topdashed sort'>��ҵ�ؼ��ʹ��������  <a href='?type=" & typeflag & "&flag=1'>δ��˹��</a>  <a href='?type=" & typeflag & "&flag=2'>�ѹ��ڹ��</a></div>"
			 End If
		End With
		
		maxperpage = 30 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("�����ϵͳ����!����������")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		totalPut = Conn.Execute("Select Count(id) From KS_EnterpriseAD")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "add","Edit" Call AddAd()
		 Case "DoSave" Call DoSave()
		 Case "Del" Call DelRecord()
		 Case "verific"  Call Verify()
		 Case "unverific"  Call UnVerify()
		 Case "View" Call ShowNews()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<script src="../ks_inc/kesion.box.js"></script>
<script>
function ShowIframe(id)
{
  PopupCenterIframe('<b>�鿴��ҵ�ؼ��ʹ�����</b>',"KS.EnterpriseAD.asp?action=View&ProID="+id,550,300,'auto')
}
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</td>
	<td nowrap>�������</td>
	<td nowrap>������</td>
	<td nowrap>����λ��</td>
	<td nowrap>��Ч����</td>
	<td nowrap>��������</td>
	<td nowrap>״̬</td>
	<td nowrap>�������</td>
</tr>
<%
	sFileName = "KS.EnterpriseAD.asp?"
	Dim Param
	If KS.ChkCLng(KS.G("Flag"))=1 Then 
	  Param=" where status=0"
	ElseIf KS.ChkClng(KS.G("Flag"))=2 Then
	  If DataBaseType=0 then
	  Param=" where datediff('d',BeginDate," &SqlNowString & ")>datatimed"
	  else
	  Param=" where datediff(day,BeginDate," &SqlNowString & ")>datatimed"
	  end if
	End If
	
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_EnterpriseAD  " & Param & " order by id desc"
	If DataBaseType = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>�Ҳ�����ҵ�ؼ��ʹ�棡</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<input type="hidden" name="type" value="<%=typeflag%>">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=Rs("Title")%></a>
	<%
	 if  datediff("d",RS("Begindate"),now)> Rs("datatimed") then
	  response.write "<font color=red>�ѹ���</font>"
	 end if
	%>
	
	</td>
	<td class="splittd" align="center"><a href='../space/?<%=rs("username")%>' target='_blank'><%=Rs("username")%></a></td>
	<td class="splittd" align="center"> ��ҵ��</td>
	<td class="splittd" align="center"><%=Rs("begindate")%></td>
	<td class="splittd" align="center"><%=Rs("datatimed")%> ��</td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "<font color=red>δ��</font>"
	 case 1
	  response.write "<font color=#999999>����</font>"
	 case 2
	  response.write "<font color=blue>����</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="#" onclick="ShowIframe(<%=rs("id")%>)">���</a> 
	<a href="?type=<%=typeflag%>&Action=Edit&ID=<%=rs("id")%>">�޸�</a>
	<a href="?type=<%=typeflag%>&Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('ȷ��ɾ����'));">ɾ��</a> <a href="?type=<%=typeflag%>&Action=verific&id=<%=rs("id")%>">���</a></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td  class="splittd" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�еĹ��" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.form.Action.value='Del';this.form.submit();return true;}return false;}">
	<input type="button" value="�������" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="����ȡ�����" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	<input type="button" class="button" value="��ӹ��" onclick="location.href='?type=<%=typeflag%>&action=add'">
	</td>
</tr>
</form>
<tr>
	<td colspan=10>
	<%
	Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>

<%
End Sub

Sub AddAd()
     if KS.S("Action")="Edit" Then
		  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_EnterPriseAD Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
			 Title    = RSObj("Title")
			 ADType = RSObj("ADType")
			 BigClassID=RSObj("BigClassID")
			 SmallClassID=RSObj("SmallClassID")
			 URL   = RSObj("URL")
			 ADWZ  = RSObj("ADWZ")
			 datatimed=RSObj("datatimed")
			 PhotoUrl  = RSObj("PhotoUrl")
			 status=trim(rsobj("status"))
			 BeginDate=rsobj("Begindate")
			 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/Images/NoPhoto.gif"
			 flag=true
			 UserName=rsobj("username")
		   End If
		   RSObj.Close:Set RSObj=Nothing
		Else
		 PhotoUrl="/images/Nophoto.gif"
		 ADWZ="1"
		 URL="http://"
		 flag=false
		 status=1
		 BeginDate=Now
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
				  <input type="hidden" value="<%=typeflag%>" name="type">
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
                          <input <%if trim(adtype)="2" then response.write " checked"%> name="AdType" type="radio" onClick="document.all.SmallClassID.disabled=false;" value="2">        
                          С��</td><td width="36%" rowspan="10" align="center">
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
		if typeflag=1 then
        sqlb = "select * from ks_enterpriseClass_zs where parentid=0 order by orderid"
		else
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
		end if
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
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
						if typeflag=1 then
						sqlss="select * from ks_enterpriseclass_zs where parentid="& KS.ChkClng(BigClassID)&" order by orderid"
						else
						sqlss="select * from ks_enterpriseclass where parentid="&KS.ChkClng(BigClassID)&" order by orderid"
						end if
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
									  <!--
                                        <input name="ADWZ" type="radio" value="1"<%if trim(ADWZ)="1" then response.write " checked"%>/>��ҵ��ȫ
                                        <input name="ADWZ" type="radio" value="2"<%if trim(ADWZ)="2" then response.write " checked"%>/>��Ʒ��      -->
										<input name="ADWZ" type="hidden" value="1" />
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
                                <td height="25" align="center">��Ч���ڣ�</td>
                                <td height="25">��
                                <input name="BeginDate" type="text" class="textbox" id="BeginDate" style="width:120px; " value="<%=BeginDate%>" maxlength="40" />
                                ��ʽ��0000-00-00</td>
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
                           <input name="UserName" class="textbox" type="text" style="width:100px; " value="<%=username%>" maxlength="30" /></td>
                      </tr>
                      <tr class="tdbg">
                        <td  height="25" align="center">״  ̬��</td>
                        <td>��
						
						  <input type="radio" name="status" value="1"<%if trim(status)="1" then response.write " checked"%>> ����
						  <input type="radio" name="status" value="0"<%if trim(status)="0" then response.write " checked"%>> δ���						</td>
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
			
			UserName=Fobj.Form("UserName")
			If Conn.Execute("Select * From KS_User Where UserName='" & UserName & "'").eof Then
				Response.Write "<script>alert('��������û���������!');history.back();</script>"
				Exit Sub
			End If	 

			Dim FormPath:FormPath =KS.ReturnChannelUserUpFilesDir(9994, UserName)
			Call KS.CreateListFolder(FormPath) 
			
           
				 Title=KS.LoseHtml(Fobj.Form("Title"))
				  If Title="" Then
				    Response.Write "<script>alert('��û������������!');history.back();</script>"
				    Exit Sub
				  End IF
				 
				 Adtype=KS.ChkClng(Fobj.Form("Adtype"))
				 If AdType=0 Then AdType=1
				 BigClassID=KS.ChkCLng(Fobj.Form("ClassID"))
				 SmallClassID=KS.ChkCLng(Fobj.Form("SmallClassID"))
				 
				 URL=KS.DelSql(Fobj.Form("URL"))
				 ADWZ=KS.ChkClng(Fobj.Form("ADWZ"))
				 datatimed=KS.ChkClng(Fobj.Form("datatimed"))
				 status=KS.ChkClng(Fobj.Form("status"))
				 BeginDate=Fobj.Form("Begindate")
				 
				 If Not IsDate(BeginDate) Then
				 Call KS.AlertHistory("��ʼ���ڸ�ʽ����ȷ!",-1)
				 Response.End()
				 End If
			
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
				RSObj.Open "Select * From KS_EnterPriseAD Where ID=" & KS.ChkClng(Fobj.Form("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				 End If
				  RSObj("UserName")=UserName
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
  				  RSObj("Status")=status
				  RSObj("BeginDate")=BeginDate
				 RSObj.Update
				 If KS.ChkClng(Fobj.Form("ID"))=0 Then
				  Call KS.FileAssociation(1014,rsobj("id"),RSObj("PhotoUrl"),0)
				 Else
				  Call KS.FileAssociation(1014,rsobj("id"),RSObj("PhotoUrl"),1)
				 End If
				 
				 RSObj.Close:Set RSObj=Nothing
				 
               If KS.ChkClng(Fobj.Form("ID"))=0 Then
			     Set Fobj=Nothing
				 Response.Write "<script>if (confirm('�ؼ��ʹ���ύ�ɹ��������ύ��?')){location.href='?type=" & typeflag &"&Action=add';}else{location.href='KS.EnterPriseAD.asp';}</script>"
			   Else
			     Set Fobj=Nothing
				 Response.Write "<script>alert('�ؼ��ʹ���޸ĳɹ�!');location.href='KS.EnterPriseAD.asp?type=" & typeflag &"';</script>"
			   End If
  End Sub

'ɾ����־
Sub DelRecord()
 Dim I,ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 ID=Split(ID,",")
 For I=0 To Ubound(ID)
  KS.DeleteFile(conn.execute("select photourl from ks_EnterpriseAD where id=" & ID(I))(0))
  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1014 and infoid=" & ID(I))
  Conn.execute("Delete From KS_EnterpriseAD Where id="& id(I))
 Next 
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'���
Sub ShowNews()
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_EnterpriseAD where id=" &KS.ChkClng(KS.S("ProID")),conn,1,1
		If Not RS.Eof Then
		   Response.Write "<div><strong>Ͷ�����ͣ�</strong>" 
		    If RS("AdType")=1 Then
			 Response.Write "����"
			Else
			 Response.Write "С��"
			End If
		   Response.Write "</div>"
		   Response.WRITE "<div><strong>������ƣ�</strong>" & rs("Title") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>���ӵ�ַ��</strong>" & RS("url") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>����λ�ã�</strong>" 
		   If RS("ADWZ")="1" Then
		    response.write "��Ʒ��"
		   Else
		    response.write "��ҵ��ȫ"
		   End If
		   Response.Write "</div>"
		   Response.Write "<div style=""text-align:left""><strong>��ʼ���ڣ�</strong>" & RS("begindate") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>��ʼ������</strong>" & RS("datatimed") & " ��</div>"
		   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
		   If PhotoUrl<>"" And Not IsNull(PhotoURL) Then
		   Response.Write "<div style=""text-align:left""><strong>���ͼƬ��</strong><img src='" & RS("photourl") & "'></div>"
		   End If
		End If
		RS.Close:Set RS=Nothing
End Sub
'���
Sub Verify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseAD Set status=1,begindate=" & SqlNowString & " Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'ȡ�����
Sub UnVerify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_EnterpriseAD Set status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 

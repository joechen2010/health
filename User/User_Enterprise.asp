<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_EditInfo
KSCls.Kesion()
Set KSCls = Nothing

Class User_EditInfo
        Private KS,KSUser
		Private FieldsXml,Action
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		if KS.SSetting(0)<>1 then
		  Response.Write "<script>alert('ϵͳû�п�ͨ�ռ书��!');history.back();</script>"
		  Response.end
		End If
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Action=Request("action")
		If Action="Why" Then
		 Call ShowWhy()
		 Response.End
		End If
		Call KSUser.Head()
		%>
		<script src="../ks_inc/kesion.box.js" language="JavaScript"></script>
        <script>
       function ShowIframe(ev)
       {
		    mousePopupIframe(ev,'ΪʲôҪ����Ϊ��ҵ�ռ�','?action=Why',500,300,'no')
       }
		</script>	
		<div class="tabs">	
			<ul>
	        <li<%if action="" then response.write " class='select'"%>><a href="user_enterprise.asp">��ҵ��Ϣ</a></li>
	        <li<%if action="intro" then response.write " class='select'"%>><a href="?action=intro">��ҵ���</a></li>
			<%if action="job" then
			 if KS.C_S(10,21)="0" then response.write "<li class='select'><a href='?action=job'>��ҵ��Ƹ</a></li>"
			end if%>
			</ul>
			<div style="padding-top:8px" onClick="ShowIframe(event)"><font style="font-size:12px;font-weight:200;color:red;cursor:help">Ϊʲô����Ϊ��ҵ�ռ�?</font>
</div>
		</div>

		<%
		Dim HasEnterprise:HasEnterprise=Not Conn.execute("select id from KS_Enterprise where username='" & KSUser.UserName & "'").eof
		Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
		Select Case KS.S("Action")
		  Case "BasicInfoSave"
		   Call BasicInfoSave()
		  Case "intro"
		   If (HasEnterprise) then
	        Call KSUser.InnerLocation("��ҵ���")
		    Call Intro()
		   Else
		    Response.Write "<script>alert('�Բ����㻹û����д��ҵ������Ϣ!')</script>"
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		   End If
		  case "IntroSave"
		   Call IntroSave()
		  Case "job"
		   If (HasEnterprise) then
	        Call KSUser.InnerLocation("��ҵ��Ƹ")
			If KS.C_S(10,21)="1" Then
			 Response.Redirect("User_JobCompanyZW.asp")
			Else
		    Call Job()
			End If
		   Else
		    Response.Write "<script>alert('�Բ����㻹û����д��ҵ������Ϣ!')</script>"
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		   End If
		  Case "JobSave"
		   Call JobSave()
		  Case Else
	       Call KSUser.InnerLocation("��ҵ������Ϣ")
		   Call EditBasicInfo()
		End Select
	   End Sub
	   
	   Sub ShowWhy()
	   %>
	   <style>
	    body{font-size:12px;line-height:160%}
		</style>
		<strong>��ܰ��ʾ��</strong>
		<br><font color=red>��վ���������ҵ�ռ���רΪ��ҵ�û���Ƶ�,������Ǹ����û����벻Ҫ������ҵ�ռ䣡</font>
		<br>
	    <strong>��ҵ�ռ书�ܽ���</strong><br>
		 <li>��ҵ����
		 <li>���ŷ���
		 <li>��Ʒչʾ
		 <Li>��ҵ��Ƹ
		 <li>�ͻ�����
		 <li>��ҵ���
		 <li>��ҵ��־</li>
		 <li>���󷢲�</li>
		<br> <strong>������ҵ�ռ���ʲô����</strong>
		 <br> ��ҵ��ͬʱӵ��һ�������Ķ�������,�����������ҵ�ռ��ģ�壬�����Ʒ�����ҵ��Ʒ��ͬʱ�������ǵĻ�ҳ�⣬��Ʒ�⣡�����ҵ��֪���ȡ�
	   <%
	   End Sub
	   '������Ϣ
	   Sub EditBasicInfo()
		   %>
      <script>
       function CheckForm() 
		{ 
			
			if (document.myform.CompanyName.value =="")
			{
			alert("����д��˾���ƣ�");
			document.myform.CompanyName.focus();
			return false;
			}
			if (document.myform.LegalPeople.value =="")
			{
			alert("����д��ҵ���ˣ�");
			document.myform.LegalPeople.focus();
			return false;
			}
			if (document.myform.TelPhone.value =="")
			{
			alert("��������ϵ�绰��");
			document.myform.TelPhone.focus();
			return false;
			}
		  return true;	
		}
		
    </script>
	<%	   

	 Dim CompanyName,Province,City,Address,ZipCode,ContactMan,Telphone,Fax,WebUrl,Profession,CompanyScale,RegisteredCapital,LegalPeople,BankAccount,AccountNumber,BusinessLicense,Intro,flag,ClassID,SmallClassID,qq,mobile,Email
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open "Select * From KS_Enterprise where username='" & KSUser.UserName & "'",conn,1,1
	 IF Not RS.Eof Then
	   CompanyName=RS("CompanyName")
	   Province=RS("Province")
	   City=RS("City")
	   Address=RS("Address")
	   ZipCode=RS("ZipCode")
	   ContactMan=RS("ContactMan")
	   Telphone=RS("TelPhone")
	   Fax=RS("Fax")
	   WebUrl=RS("WebUrl")
	   Profession=RS("Profession")
	   CompanyScale=RS("CompanyScale")
	   RegisteredCapital=RS("RegisteredCapital")
	   LegalPeople=RS("LegalPeople")
	   BankAccount=RS("BankAccount")
	   AccountNumber=RS("AccountNumber")
	   BusinessLicense=RS("BusinessLicense")
	   ClassID=RS("ClassID")
	   SmallClassID=RS("SmallClassID")
	   qq=rs("qq")
	   Email=rs("Email")
	   mobile=rs("mobile")
	   flag=true
	 Else
	   flag=false
	    if KS.SSetting(17)<>"" then
	    if KS.FoundInArr(KS.SSetting(17),KSUser.groupid,",")=false then  Set KSUser=Nothing:call KS.AlertHistory("�Բ��������ڵ��û���û��Ȩ������Ϊ��ҵ�ռ䣡",-1):exit sub
	   end if
	   If IsObject(FieldsXml) Then
	     on error resume next
	     Dim objNode,i,j,objAtr
	     Set objNode=FieldsXml.documentElement 
		 For i=0 to objNode.ChildNodes.length-1 
				set objAtr=objNode.ChildNodes.item(i) 
				' response.write objAtr.Attributes.item(0).Text&"=" &objAtr.Attributes.item(1).Text & " <br>" 
				 Execute(objAtr.Attributes.item(0).Text&"=""" & LFCls.GetSingleFieldValue("select " & objAtr.Attributes.item(1).Text & " From KS_User Where UserName='" & KSUser.UserName & "'") & """") 
		 Next

	   End If
	   
	 End If
	 If ClassID="" or isnull(ClassID) Then  ClassID=0
	 If SmallClassID="" or isnull(ClassID) Then SmallClassID=0

    RS.Close:Set RS=Nothing	
	%>
          
          <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
					  <form action="?Action=BasicInfoSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
					      <input type="hidden" value="<%=KS.S("ComeUrl")%>" name="ComeUrl">
                          <tr class="title">
                            <td height="22" colspan="2" align="center"> �� ҵ �� �� �� �� </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> ��˾���ƣ� </span><br>
                              ����д���ڹ��̾�ע��Ǽǵ����ơ�</td>
                            <td width="72%">&nbsp;
                                <input name="CompanyName" type="text" class="textbox" id="CompanyName" value="<%=CompanyName%>" size="30" maxlength="200" />
                                <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">Ӫҵ���գ�</span><br>
��д���Ӫҵִ��ͼƬ���ڵ�ַ��Ӫҵִ�պ��롣</td>
                            <td>&nbsp;
                              <input name="BusinessLicense" class="textbox" type="text" id="BusinessLicense" value="<%=BusinessLicense%>" size="30" maxlength="50" /></td>
                          </tr>
                         <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">��˾��ҵ��</span><br>
��д��˾��������ҵ��</td>
                            <td>&nbsp;
							
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
			for (var i=0;i < onecount; i++)
				{ 
					if (parseInt(subcat[i][1]) == parseInt(locationid))
					{ 			
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		  <select class="face" name="ClassID" onChange="changelocation(document.myform.ClassID.options[document.myform.ClassID.selectedIndex].value)" size="1">
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    Dim N
		    do while not rsb.eof
			          N=N+1
					  If N=1 and flag=false Then ClassID=rsb("id")
					  If ClassID=rsb("id") then
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
                  <select class="face" name="SmallClassID">
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="&ClassID&" order by orderid"
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
                </select>
							  
							  </td>
                          </tr>
						  
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">��ҵ���ˣ�</span></td>
                            <td>&nbsp;
                              <input name="LegalPeople" class="textbox" type="text" id="LegalPeople" value="<%=LegalPeople%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">��˾��ģ��</span></td>
                            <td>&nbsp;
                              <select name="CompanyScale" id="CompanyScale">
							  <option value="1-20��"<%if CompanyScale="1-20��" then response.write " selected"%>>1-20��</option>
                      <option value="21-50��"<%if CompanyScale="21-50��" then response.write " selected"%>>21-50��</option>
                      <option value="51-100��"<%if CompanyScale="51-100��" then response.write " selected"%>>51-100��</option>
                      <option value="101-200��"<%if CompanyScale="101-200��" then response.write " selected"%>>101-200��</option>
                      <option value="201-500��"<%if CompanyScale="201-500��" then response.write " selected"%>>201-500��</option>
                      <option value="501-1000��"<%if CompanyScale="501-1000��" then response.write " selected"%>>501-1000��</option>
                      <option value="1000������"<%if CompanyScale="1000������" then response.write " selected"%>>1000������</option>
						    </select></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">ע���ʽ�</span></td>
                            <td>&nbsp;
							<select name="RegisteredCapital" id="RegisteredCapital">
							<option value="10������"<%if RegisteredCapital="10������" then response.write " selected"%>>10������</option>
                      <option value="10��-19��"<%if RegisteredCapital="10��-19��" then response.write " selected"%>>10��-19��</option>
                      <option value="20��-49��"<%if RegisteredCapital="20��-49��" then response.write " selected"%>>20��-49��</option>
                      <option value="50��-99��"<%if RegisteredCapital="50��-99��" then response.write " selected"%>>50��-99��</option>
                      <option value="100��-199��"<%if RegisteredCapital="100��-199��" then response.write " selected"%>>100��-199��</option>
                      <option value="200��-499��"<%if RegisteredCapital="200��-499��" then response.write " selected"%>>200��-499��</option>
                      <option value="500��-999��"<%if RegisteredCapital="500��-999��" then response.write " selected"%>>500��-999��</option>
                      <option value="1000������"<%if RegisteredCapital="1000������" then response.write " selected"%>>1000������</option>
					   </select></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">���ڵ�����</span><br>
                              ѡ����ҵ���ڵ�ʡ�ݺͳ��С�</td>
                            <td>&nbsp;
							<script src="../plus/area.asp" language="javascript"></script>
							<script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
								  <%end if%>
							  <%if City<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=City%>','<%=City%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script>
							  
							  </td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">�� ϵ �ˣ�</span></td>
                            <td> &nbsp;
<input name="ContactMan" class="textbox" type="text" id="ContactMan" value="<%=ContactMan%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> ��˾��ַ��</span><br>
                            ��д��˾����ϵ��ַ</td>
                            <td width="72%">&nbsp;
                              <input name="Address" class="textbox" type="text" id="Adress" value="<%=Address%>" size="30" maxlength="50" /></td>
                          </tr>
       
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> �������룺 </span><br></td>
                            <td width="72%">&nbsp;
                            <input name="ZipCode" class="textbox" type="text" id="ZipCode" value="<%=ZipCode%>" size="30" maxlength="10" />    </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> QQ���룺</span><br>
							</td>
                            <td width="72%">&nbsp;
                              <input name="qq" class="textbox" type="text" id="qq" value="<%=qq%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> �ֻ����룺</span></td>
                            <td width="72%">&nbsp;
                              <input name="Mobile" class="textbox" type="text" id="Mobile" value="<%=Mobile%>" size="30" maxlength="50" />
                           </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> ��ϵ�绰��</span><br>
							��˾�칫�绰������ҵ����ϵ��</td>
                            <td width="72%">&nbsp;
                              <input name="TelPhone" class="textbox" type="text" id="TelPhone" value="<%=Telphone%>" size="30" maxlength="50" />
                            <span style="color: red">* </span>������</td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> ������룺</span><br>
                            ��˾�Ĵ�����롣</td>
                            <td width="72%">&nbsp;
                              <input name="Fax" class="textbox" type="text" id="Fax" value="<%=Fax%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> �������䣺</span></td>
                            <td width="72%">&nbsp;
                              <input name="Email" class="textbox" type="text" id="Email" value="<%=Email%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">��˾��վ��</span><br> 
                            ��д�㹫˾����ַ��</td>
                            <td>&nbsp;
                              <input name="WebUrl" class="textbox" type="text" id="WebUrl" value="<%=WebUrl%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">�������У�</span></td>
                            <td>&nbsp;
                              <input name="BankAccount" class="textbox" type="text" id="BankAccount" value="<%=BankAccount%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">�����˺ţ�<br>
                            </span>��˾�����ʻ����Է������������ϵ�����С�</td>
                            <td>&nbsp;
                              <input name="AccountNumber" class="textbox" type="text" id="AccountNumber" value="<%=AccountNumber%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input  class="button" name="Submit" type="submit"  value=" OK,ȷ �� " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" �� �� " />                            </td>
                          </tr>
		    </form>
            </table>
          <%
  End Sub
  
  Sub Intro()
  %>
   <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
			<form action="?Action=IntroSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
               <tr class="title">
                   <td height="22" colspan="2" align="center"> �� ҵ �� ��</td>
               </tr>
               <tr class="tdbg">
                  <td>
				  <font color=#a7a7a7>������������ϸ˵����˾�ĳ�����ʷ����Ӫ��Ʒ��Ʒ�ơ���������ƣ�<br>
 ��������ݹ��ڼ򵥻����д�����Ĳ�Ʒ���ܣ����п����޷�ͨ����ˡ�<br>
����ϵ��ʽ���绰�����桢�ֻ�����������ȣ����ڻ�����������д�� �˴������ظ���д��<br></font>
                    <%
					Dim Intro:Intro=Conn.Execute("Select Intro From ks_Enterprise where username='" & KSUser.UserName & "'")(0)
					If trim(Intro)="" Or IsNull(Intro) Then
						If IsObject(FieldsXml) Then
						 'on error resume next
						 Dim objNode,i,j,objAtr
						 Set objNode=FieldsXml.documentElement 
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i)
								If lcase(objAtr.Attributes.item(0).Text)="intro" Then 
								 Intro=LFCls.GetSingleFieldValue("select " & objAtr.Attributes.item(1).Text & " From KS_User Where UserName='" & KSUser.UserName & "'") 
								End If
						 Next
				
					   End If
					End If
					
			        Response.Write "<textarea ID='Intro' name='Intro' style='display:none'>" & KS.HTMLCode(Intro) & "</textarea>"
					Response.Write "<input type=""hidden"" id=""Intro___Config"" value="""" style=""display:none"" /><iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Intro&amp;Toolbar=NewsTool"" width=""98%"" height=""350"" frameborder=""0"" scrolling=""no""></iframe>"
					%>   
					</td>
                          </tr>
						  <tr class="tdbg">
                            <td align="center"><input  class="button" name="Submit" type="submit"  value=" OK,�� �� " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" �� �� " />                            </td>
                          </tr>
				</form>
	</table>
  <%
  End Sub
 
 Sub IntroSave()
  Dim Intro
  Intro = Request.Form("Intro")
  Intro=KS.CheckScript(KS.HtmlCode(Intro))
  Intro=KS.HtmlEncode(Intro)
  IF Intro="" Then
  	 Response.Write "<script>alert('�Բ�����û�����빫˾���');history.back();</script>"
	 Response.end
  End If
  If IsObject(FieldsXml) Then
	on error resume next
	Dim objNode,i,j,objAtr
	 Set objNode=FieldsXml.documentElement 
	 For i=0 to objNode.ChildNodes.length-1 
		set objAtr=objNode.ChildNodes.item(i)
		If lcase(objAtr.Attributes.item(0).Text)="intro" Then 
		 Conn.Execute("UPDATE KS_User Set " & objAtr.Attributes.item(1).Text & "='" & Intro & "' Where UserName='" & KSUser.UserName & "'")
		End If
	 Next
				
  End If
  Conn.Execute("Update KS_EnterPrise Set Intro='" & Intro &"' WHERE UserName='" & KSUser.UserName & "'")
  Dim EID:EID=Conn.Execute("Select ID From KS_Enterprise Where UserName='" & KSUser.UserName & "'")(0)
  Call KS.FileAssociation(1033,EID,Intro,1)
  Call KSUser.AddLog(KSUser.UserName,"�޸�����ҵ������!",200)
  Response.Write "<script>alert('��ҵ����޸ĳɹ�!');history.back();</script>"
 End Sub
 
 
  Sub Job()
  %>
   <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
			<form action="?Action=JobSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
               <tr class="title">
                   <td height="22" colspan="2" align="center"> �� ҵ �� Ƹ</td>
               </tr>
               <tr class="tdbg">
                  <td>
                    <%
					Response.Write "<textarea ID='Job' name='Job' style='display:none'>" & KS.HTMLCode(Conn.Execute("Select Job From ks_Enterprise where username='" & KSUser.UserName & "'")(0)) & "</textarea>"
					Response.Write "<input type=""hidden"" id=""Job___Config"" value="""" style=""display:none"" /><iframe id=""Job___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Job&amp;Toolbar=NewsTool"" width=""98%"" height=""350"" frameborder=""0"" scrolling=""no""></iframe>"

					%>   
					</td>
                          </tr>
						  <tr class="tdbg">
                            <td align="center"><input  class="button" name="Submit" type="submit"  value=" OK,�� �� " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" �� �� " />                            </td>
                          </tr>
				</form>
	</table>
  <%
  End Sub
 
 Sub JobSave()
  Dim Job
  Job= Request.Form("Job")
  Job=KS.CheckScript(KS.HtmlCode(Job))
  Job=KS.HtmlEncode(Job)
  IF Job="" Then
  	 Response.Write "<script>alert('�Բ�����û����Ƹ��Ϣ');history.back();</script>"
	 Response.end
  End If
  Conn.Execute("Update KS_EnterPrise Set Job='" & Job &"' WHERE UserName='" & KSUser.UserName & "'")
  Response.Write "<script>alert('��Ƹ��Ϣ�޸ĳɹ�!');history.back();</script>"
 End Sub
 
  
  Sub BasicInfoSave() 
	   Dim CompanyName:CompanyName=KS.LoseHtml(KS.S("CompanyName"))
	   Dim Province:Province=KS.S("Province")
	   Dim City:City=KS.S("City")
	   Dim Address:Address=KS.LoseHtml(KS.S("Address"))
	   Dim ZipCode:ZipCode=KS.LoseHtml(KS.S("ZipCode"))
	   Dim ContactMan:ContactMan=KS.LoseHtml(KS.S("ContactMan"))
	   Dim QQ:QQ=KS.S("QQ")
	   Dim Mobile:mobile=KS.S("Mobile")
	   Dim Email:Email=KS.S("Email")
	   Dim Telphone:TelPhone=KS.LoseHtml(KS.S("TelPhone"))
	   Dim Fax:Fax=KS.LoseHtml(KS.S("Fax"))
	   Dim WebUrl:WebUrl=KS.LoseHtml(KS.S("WebUrl"))
	   Dim Profession:Profession=KS.LoseHtml(KS.S("Profession"))
	   Dim CompanyScale:CompanyScale=KS.LoseHtml(KS.S("CompanyScale"))
	   Dim RegisteredCapital:RegisteredCapital=KS.LoseHtml(KS.S("RegisteredCapital"))
	   Dim LegalPeople:LegalPeople=KS.LoseHtml(KS.S("LegalPeople"))
	   Dim BankAccount:BankAccount=KS.LoseHtml(KS.S("BankAccount"))
	   Dim AccountNumber:AccountNumber=KS.LoseHtml(KS.S("AccountNumber"))
	   Dim BusinessLicense:BusinessLicense=KS.LoseHtml(KS.S("BusinessLicense"))
	   Dim ClassID:ClassID=KS.ChkClng(KS.G("ClassID"))
	   Dim SmallClassID:SmallClassID=KS.ChkClng(KS.G("SmallClassID"))
	   Dim NewReg:NewReg=false
		
	   If CompanyName="" Then Response.Write "<script>alert('��˾���Ʊ�������');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_Enterprise Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 RS.AddNew
				 RS("UserName")=KSUser.UserName
				 RS("AddDate")=Now
				 RS("Recommend")=0
				 If KS.SSetting(2)=1 then
				 RS("status")=0
				 Else
				 RS("status")=1
				 End If
				 Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
				 RSS.Open "select * from ks_blog where username='" & KSUser.UserName & "'",conn,1,3
				 if RSS.Eof Then
				      RSS.AddNew
					  RSS("UserName")=KSUser.UserName
					  RSS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
					  RSS("Announce")="���޹���!"
					  RSS("ContentLen")=500
					  RSS("Recommend")=0
				 End If
					  if KS.SSetting(2)=1 then
					  RSS("Status")=0
					  else
					  RSS("Status")=1
					  end if
				  RSS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
     			  RSS("BlogName")=CompanyName
				  RSS.Update
				  RSS.Close
				  Set RSS=Nothing
				  NewReg=true
				 
			  End If
			     RS("CompanyName")=CompanyName
				 RS("Province")=Province
				 RS("City")=City
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("ContactMan")=ContactMan
				 RS("QQ")=QQ
				 RS("Mobile")=Mobile
				 RS("Email")=Email
				 RS("Telphone")=Telphone
				 RS("Fax")=Fax
				 RS("WebUrl")=WebUrl
				 RS("Profession")=Profession
				 RS("CompanyScale")=CompanyScale
				 RS("RegisteredCapital")=RegisteredCapital
				 RS("LegalPeople")=LegalPeople
				 RS("BankAccount")=BankAccount
				 RS("AccountNumber")=AccountNumber
				 RS("BusinessLicense")=BusinessLicense
				 RS("ClassID")=ClassID
				 RS("SmallClassID")=SmallClassID
				 'RS("Intro")=KS.HtmlEncode(Request.Form("Intro"))
		 		 RS.Update
				 Conn.Execute("Update KS_User Set UserType=1 where UserName='" & KSUser.UserName & "'")
				 If KS.C_S(8,21)="1" Then
				 Conn.Execute("Update KS_GQ Set ContactMan='" & ContactMan &"',Tel='" & Telphone & "',CompanyName='" & CompanyName & "',Address='" & Address & "',Province='" & Province & "',City='" & City & "',Zip='" & ZipCode & "',Fax='" & Fax & "',Homepage='" & WebUrl & "' where inputer='" & KSUser.UserName & "'")
				 End If
				 
				 
				 Set RSS=Conn.Execute("Select BlogName From KS_Blog Where UserName='" & KSUser.UserName & "'")
				 If Not RSS.Eof Then
				   If Instr(RSS(0),"���˿ռ�")<>0 Then
				    Conn.Execute("Update KS_Blog Set BlogName='" & CompanyName & "' where username='" & KSUser.UserName &"'")
				   End If
				 End If
				 RSS.Close
				 Set RSS=Nothing
				 
				 If IsObject(FieldsXml) Then
					 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					 If objNode.Attributes.item(0).Text="2" Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								If lcase(objAtr.Attributes.item(0).Text)<>"intro" Then 
								Conn.Execute("UPDATE KS_User Set " & objAtr.Attributes.item(1).Text & "='" & RS(objAtr.Attributes.item(0).Text) & "' Where UserName='" & KSUser.UserName & "'")
								End If
						 Next
					 End If
			
				   End If
				 
				 RS.Close:Set RS=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"�޸�����ҵ������Ϣ����!",200)
				 If KS.S("ComeUrl")<>"" then
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ���');location.href='" & KS.S("ComeUrl") & "';</script>"
				 Else
				  if NewReg=true Then
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ�,��ȷ����д��ҵ���ܣ�');top.location.href='index.asp?user_Enterprise.asp?action=intro';</script>"
				  Else
				 Response.Write "<script>alert('��ҵ������Ϣ�����޸ĳɹ���');location.href='user_Enterprise.asp';</script>"
				  End If
				End If
  End Sub
 

End Class
%> 

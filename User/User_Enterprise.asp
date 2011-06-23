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
		  Response.Write "<script>alert('系统没有开通空间功能!');history.back();</script>"
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
		    mousePopupIframe(ev,'为什么要升级为企业空间','?action=Why',500,300,'no')
       }
		</script>	
		<div class="tabs">	
			<ul>
	        <li<%if action="" then response.write " class='select'"%>><a href="user_enterprise.asp">企业信息</a></li>
	        <li<%if action="intro" then response.write " class='select'"%>><a href="?action=intro">企业简介</a></li>
			<%if action="job" then
			 if KS.C_S(10,21)="0" then response.write "<li class='select'><a href='?action=job'>企业招聘</a></li>"
			end if%>
			</ul>
			<div style="padding-top:8px" onClick="ShowIframe(event)"><font style="font-size:12px;font-weight:200;color:red;cursor:help">为什么升级为企业空间?</font>
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
	        Call KSUser.InnerLocation("企业简介")
		    Call Intro()
		   Else
		    Response.Write "<script>alert('对不起，你还没有填写企业基本信息!')</script>"
	       Call KSUser.InnerLocation("企业基本信息")
		   Call EditBasicInfo()
		   End If
		  case "IntroSave"
		   Call IntroSave()
		  Case "job"
		   If (HasEnterprise) then
	        Call KSUser.InnerLocation("企业招聘")
			If KS.C_S(10,21)="1" Then
			 Response.Redirect("User_JobCompanyZW.asp")
			Else
		    Call Job()
			End If
		   Else
		    Response.Write "<script>alert('对不起，你还没有填写企业基本信息!')</script>"
	       Call KSUser.InnerLocation("企业基本信息")
		   Call EditBasicInfo()
		   End If
		  Case "JobSave"
		   Call JobSave()
		  Case Else
	       Call KSUser.InnerLocation("企业基本信息")
		   Call EditBasicInfo()
		End Select
	   End Sub
	   
	   Sub ShowWhy()
	   %>
	   <style>
	    body{font-size:12px;line-height:160%}
		</style>
		<strong>温馨提示：</strong>
		<br><font color=red>本站所开设的企业空间是专为企业用户设计的,如果您是个人用户，请不要申请企业空间！</font>
		<br>
	    <strong>企业空间功能介绍</strong><br>
		 <li>企业介绍
		 <li>新闻发布
		 <li>产品展示
		 <Li>企业招聘
		 <li>客户留言
		 <li>企业相册
		 <li>企业日志</li>
		 <li>供求发布</li>
		<br> <strong>加入企业空间有什么优势</strong>
		 <br> 企业可同时拥有一个独立的二级域名,可自由设计企业空间的模板，不限制发布企业产品。同时加入我们的黄页库，产品库！提高企业的知名度。
	   <%
	   End Sub
	   '基本信息
	   Sub EditBasicInfo()
		   %>
      <script>
       function CheckForm() 
		{ 
			
			if (document.myform.CompanyName.value =="")
			{
			alert("请填写公司名称！");
			document.myform.CompanyName.focus();
			return false;
			}
			if (document.myform.LegalPeople.value =="")
			{
			alert("请填写企业法人！");
			document.myform.LegalPeople.focus();
			return false;
			}
			if (document.myform.TelPhone.value =="")
			{
			alert("请输入联系电话！");
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
	    if KS.FoundInArr(KS.SSetting(17),KSUser.groupid,",")=false then  Set KSUser=Nothing:call KS.AlertHistory("对不起，你所在的用户组没有权利升级为企业空间！",-1):exit sub
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
                            <td height="22" colspan="2" align="center"> 企 业 基 本 资 料 </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 公司名称： </span><br>
                              请填写你在工商局注册登记的名称。</td>
                            <td width="72%">&nbsp;
                                <input name="CompanyName" type="text" class="textbox" id="CompanyName" value="<%=CompanyName%>" size="30" maxlength="200" />
                                <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">营业热照：</span><br>
填写你的营业执照图片所在地址或营业执照号码。</td>
                            <td>&nbsp;
                              <input name="BusinessLicense" class="textbox" type="text" id="BusinessLicense" value="<%=BusinessLicense%>" size="30" maxlength="50" /></td>
                          </tr>
                         <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">公司行业：</span><br>
填写公司所属的行业。</td>
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
                            <td height="22"><span style="font-weight: bold">企业法人：</span></td>
                            <td>&nbsp;
                              <input name="LegalPeople" class="textbox" type="text" id="LegalPeople" value="<%=LegalPeople%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">公司规模：</span></td>
                            <td>&nbsp;
                              <select name="CompanyScale" id="CompanyScale">
							  <option value="1-20人"<%if CompanyScale="1-20人" then response.write " selected"%>>1-20人</option>
                      <option value="21-50人"<%if CompanyScale="21-50人" then response.write " selected"%>>21-50人</option>
                      <option value="51-100人"<%if CompanyScale="51-100人" then response.write " selected"%>>51-100人</option>
                      <option value="101-200人"<%if CompanyScale="101-200人" then response.write " selected"%>>101-200人</option>
                      <option value="201-500人"<%if CompanyScale="201-500人" then response.write " selected"%>>201-500人</option>
                      <option value="501-1000人"<%if CompanyScale="501-1000人" then response.write " selected"%>>501-1000人</option>
                      <option value="1000人以上"<%if CompanyScale="1000人以上" then response.write " selected"%>>1000人以上</option>
						    </select></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">注册资金：</span></td>
                            <td>&nbsp;
							<select name="RegisteredCapital" id="RegisteredCapital">
							<option value="10万以下"<%if RegisteredCapital="10万以下" then response.write " selected"%>>10万以下</option>
                      <option value="10万-19万"<%if RegisteredCapital="10万-19万" then response.write " selected"%>>10万-19万</option>
                      <option value="20万-49万"<%if RegisteredCapital="20万-49万" then response.write " selected"%>>20万-49万</option>
                      <option value="50万-99万"<%if RegisteredCapital="50万-99万" then response.write " selected"%>>50万-99万</option>
                      <option value="100万-199万"<%if RegisteredCapital="100万-199万" then response.write " selected"%>>100万-199万</option>
                      <option value="200万-499万"<%if RegisteredCapital="200万-499万" then response.write " selected"%>>200万-499万</option>
                      <option value="500万-999万"<%if RegisteredCapital="500万-999万" then response.write " selected"%>>500万-999万</option>
                      <option value="1000万以上"<%if RegisteredCapital="1000万以上" then response.write " selected"%>>1000万以上</option>
					   </select></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">所在地区：</span><br>
                              选择企业所在的省份和城市。</td>
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
                            <td height="22"><span style="font-weight: bold">联 系 人：</span></td>
                            <td> &nbsp;
<input name="ContactMan" class="textbox" type="text" id="ContactMan" value="<%=ContactMan%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 公司地址：</span><br>
                            填写公司的联系地址</td>
                            <td width="72%">&nbsp;
                              <input name="Address" class="textbox" type="text" id="Adress" value="<%=Address%>" size="30" maxlength="50" /></td>
                          </tr>
       
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 邮政编码： </span><br></td>
                            <td width="72%">&nbsp;
                            <input name="ZipCode" class="textbox" type="text" id="ZipCode" value="<%=ZipCode%>" size="30" maxlength="10" />    </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> QQ号码：</span><br>
							</td>
                            <td width="72%">&nbsp;
                              <input name="qq" class="textbox" type="text" id="qq" value="<%=qq%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 手机号码：</span></td>
                            <td width="72%">&nbsp;
                              <input name="Mobile" class="textbox" type="text" id="Mobile" value="<%=Mobile%>" size="30" maxlength="50" />
                           </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 联系电话：</span><br>
							公司办公电话，用于业务联系！</td>
                            <td width="72%">&nbsp;
                              <input name="TelPhone" class="textbox" type="text" id="TelPhone" value="<%=Telphone%>" size="30" maxlength="50" />
                            <span style="color: red">* </span>带区号</td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 传真号码：</span><br>
                            公司的传真号码。</td>
                            <td width="72%">&nbsp;
                              <input name="Fax" class="textbox" type="text" id="Fax" value="<%=Fax%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 电子邮箱：</span></td>
                            <td width="72%">&nbsp;
                              <input name="Email" class="textbox" type="text" id="Email" value="<%=Email%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">公司网站：</span><br> 
                            填写你公司的网址。</td>
                            <td>&nbsp;
                              <input name="WebUrl" class="textbox" type="text" id="WebUrl" value="<%=WebUrl%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">开户银行：</span></td>
                            <td>&nbsp;
                              <input name="BankAccount" class="textbox" type="text" id="BankAccount" value="<%=BankAccount%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">银行账号：<br>
                            </span>公司银行帐户，以方便放在您的联系资料中。</td>
                            <td>&nbsp;
                              <input name="AccountNumber" class="textbox" type="text" id="AccountNumber" value="<%=AccountNumber%>" size="30" maxlength="50" /></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input  class="button" name="Submit" type="submit"  value=" OK,确 认 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
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
                   <td height="22" colspan="2" align="center"> 企 业 简 介</td>
               </tr>
               <tr class="tdbg">
                  <td>
				  <font color=#a7a7a7>·请用中文详细说明贵司的成立历史、主营产品、品牌、服务等优势；<br>
 ·如果内容过于简单或仅填写单纯的产品介绍，将有可能无法通过审核。<br>
·联系方式（电话、传真、手机、电子邮箱等）请在基本资料中填写， 此处请勿重复填写。<br></font>
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
                            <td align="center"><input  class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
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
  	 Response.Write "<script>alert('对不起，你没有输入公司简介');history.back();</script>"
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
  Call KSUser.AddLog(KSUser.UserName,"修改了企业简介操作!",200)
  Response.Write "<script>alert('企业简介修改成功!');history.back();</script>"
 End Sub
 
 
  Sub Job()
  %>
   <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
			<form action="?Action=JobSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
               <tr class="title">
                   <td height="22" colspan="2" align="center"> 企 业 招 聘</td>
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
                            <td align="center"><input  class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
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
  	 Response.Write "<script>alert('对不起，你没有招聘信息');history.back();</script>"
	 Response.end
  End If
  Conn.Execute("Update KS_EnterPrise Set Job='" & Job &"' WHERE UserName='" & KSUser.UserName & "'")
  Response.Write "<script>alert('招聘信息修改成功!');history.back();</script>"
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
		
	   If CompanyName="" Then Response.Write "<script>alert('公司名称必须输入');history.back();</script>":response.end

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
					  RSS("Announce")="暂无公告!"
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
				   If Instr(RSS(0),"个人空间")<>0 Then
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
				 Call KSUser.AddLog(KSUser.UserName,"修改了企业基本信息资料!",200)
				 If KS.S("ComeUrl")<>"" then
				 Response.Write "<script>alert('企业基本信息资料修改成功！');location.href='" & KS.S("ComeUrl") & "';</script>"
				 Else
				  if NewReg=true Then
				 Response.Write "<script>alert('企业基本信息资料修改成功,点确定填写企业介绍！');top.location.href='index.asp?user_Enterprise.asp?action=intro';</script>"
				  Else
				 Response.Write "<script>alert('企业基本信息资料修改成功！');location.href='user_Enterprise.asp';</script>"
				  End If
				End If
  End Sub
 

End Class
%> 

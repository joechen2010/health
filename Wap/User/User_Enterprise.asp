<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="企业信息">
<p>
<%
Set KS=New PublicCls
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If

If KS.SSetting(0)=0 Then
   Response.Write "对不起，本站点关闭空间站点功能!<br/>"
Else

Dim HasEnterprise:HasEnterprise=Not Conn.Execute("select top 1 id from KS_Enterprise where UserName='" & KSUser.UserName & "'").EOF
		Select Case KS.S("Action")
		  Case "BasicInfoSave" Call BasicInfoSave()
		  Case "BasicInfoSave2" Call BasicInfoSave2()
		  Case "intro"
		   If (HasEnterprise) then
		    Call Intro()'企业简介
		   Else
		   Response.Write "对不起,你还没有填写企业基本信息!<br/>" &vbcrlf
		   Call EditBasicInfo()'企业基本信
		   End If
		  case "IntroSave"
		   Call IntroSave()
		  Case "job"
		   If (HasEnterprise) Then
		      Call Job()'企业招
		   Else
		      Response.Write "对不起,你还没有填写企业基本信息!<br/>" &vbcrlf
			  Call EditBasicInfo()'企业基本信息
		   End If
		  Case "JobSave"
		   Call JobSave()
		  Case Else
		   Call EditBasicInfo()'企业基本信息
		End Select

   If KS.S("Action")="Why" Then
      Call ShowWhy()
   Else
     ' Call EditBasicInfo()'企业基本信息
   End If

End If
Response.write "<br/>"
Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>" &vbcrlf
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/><br/>" &vbcrlf

Call CloseConn
Set KSUser=Nothing
Set KS=Nothing
Response.Write "</p>" &vbcrlf
Response.Write "</card>" &vbcrlf
Response.Write "</wml>" &vbcrlf

	   
Sub ShowWhy()
    Response.Write "温馨提示:<br/>" &vbcrlf
	Response.Write "本站所开设的企业空间是专为企业用户设计的,如果您是个人用户,请不要申请企业空间!<br/>" &vbcrlf
	Response.Write "企业空间功能介绍<br/>" &vbcrlf
	Response.Write "企业介绍,新闻发布,产品展示,企业招聘,客户留言,企业相册,企业日志,供求发布<br/>" &vbcrlf
End Sub

'基本信息
Sub EditBasicInfo()
	Dim CompanyName,Province,City,Address,ZipCode,ContactMan,Telphone,Fax,WebUrl,Profession,CompanyScale,RegisteredCapital,LegalPeople,BankAccount,AccountNumber,BusinessLicense,Intro,flag,classid,qq,mobile
	Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select top 1 * From KS_Enterprise where UserName='" & KSUser.UserName & "'",conn,1,1
    IF Not RS.Eof Then
	   CompanyName=RS("CompanyName")
	   Province=RS("Province")
	   'City=RS("City")
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
	   classid=RS("ClassID")
	   QQ=RS("QQ")
	   Mobile=RS("Mobile")
	   flag=True
	Else
	   Province=KSUser.Province
	   'City=KSUser.City
	   ContactMan=KSUser.RealName
	   Telphone=KSUser.OfficeTel
	   Address=KSUser.Address
	   Fax=KSUser.Fax
	   WebUrl="http://"
	   flag=False
	   If KS.SSetting(17)<>"" Then
	      If KS.FoundInArr(KS.SSetting(17),KSUser.groupid,",")=False Then
		     Response.Write "对不起,你所在的用户组没有权利升级为企业空间!<br/>" &vbcrlf
			 Exit Sub
	      End If
	   End If
	End If
    RS.Close
	%>
    
    企业基本资料<br/>
    公司名称:<input name="CompanyName" type="text" value="<%=CompanyName%>" size="30" maxlength="200" emptyok="false"/>请填写你在工商局注册登记的名称。<br/>
    营业热照:<input name="BusinessLicense" type="text" value="<%=BusinessLicense%>" size="30" maxlength="50" />填写你的营业执照图片所在地址或营业执照号码。<br/>
    企业法人:<input name="LegalPeople" type="text" value="<%=LegalPeople%>" size="30" maxlength="50" emptyok="false"/><br/>
	行业大类:<select name="classid">
	<option value="0">-请选择行业大类-</option>
	<%
	Dim XML,Node
	RS.Open "Select * From KS_enterpriseClass where parentid=0 order by orderid",conn,1,1
    If Not RS.Eof Then
	  Set XML=KS.RsToXml(RS,"row","")
	End If
	RS.CLose
	If IsObject(XML) Then
	  For Each Node In XML.DocumentElement.SelectNodes("row")
	    If trim(Node.SelectSingleNode("@id").text)=trim(classid) then
	    KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """ selected=""selected"">" & Node.SelectSingleNode("@classname").text & "</option>"
		else
	    KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """>" & Node.SelectSingleNode("@classname").text & "</option>"
		end if
	  Next
	End If
	XML=Empty
	%>
	</select><br/>
	所在省份:<select name="province">
	 <option value="0">-请选择公司所在省份-</option>
	 <%
	   RS.Open "Select * from KS_Province Where ParentID=0 order by orderid",conn,1,1
		If Not RS.Eof Then
		  Set XML=KS.RsToXml(RS,"row","")
		End If
		RS.CLose
		Set RS=Nothing
		If IsObject(XML) Then
		  For Each Node In XML.DocumentElement.SelectNodes("row")
		   If trim(Node.SelectSingleNode("@city").text)=trim(province) Then
		   KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """ selected=""selected"">" & Node.SelectSingleNode("@city").text & "</option>"
		   Else
		   KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """>" & Node.SelectSingleNode("@city").text & "</option>"
		   End If
		  Next
		End If
	  
	 %>
	</select><br/>
	
    公司规模:<select name="CompanyScale">
             <option value="1-20人"<%if CompanyScale="1-20人" then response.write " selected=""selected"""%>>1-20人</option>
             <option value="21-50人"<%if CompanyScale="21-50人" then response.write " selected=""selected"""%>>21-50人</option>
             <option value="51-100人"<%if CompanyScale="51-100人" then response.write " selected=""selected"""%>>51-100人</option>
             <option value="101-200人"<%if CompanyScale="101-200人" then response.write " selected=""selected"""%>>101-200人</option>
             <option value="201-500人"<%if CompanyScale="201-500人" then response.write " selected=""selected"""%>>201-500人</option>
             <option value="501-1000人"<%if CompanyScale="501-1000人" then response.write " selected=""selected"""%>>501-1000人</option>
             <option value="1000人以上"<%if CompanyScale="1000人以上" then response.write " selected=""selected"""%>>1000人以上</option>
             </select><br/>
    注册资金:<select name="RegisteredCapital">
             <option value="10万以下"<%if RegisteredCapital="10万以下" then response.write " selected=""selected"""%>>10万以下</option>
             <option value="10万-19万"<%if RegisteredCapital="10万-19万" then response.write " selected=""selected"""%>>10万-19万</option>
             <option value="20万-49万"<%if RegisteredCapital="20万-49万" then response.write " selected=""selected"""%>>20万-49万</option>
             <option value="50万-99万"<%if RegisteredCapital="50万-99万" then response.write " selected=""selected"""%>>50万-99万</option>
             <option value="100万-199万"<%if RegisteredCapital="100万-199万" then response.write " selected=""selected"""%>>100万-199万</option>
             <option value="200万-499万"<%if RegisteredCapital="200万-499万" then response.write " selected=""selected"""%>>200万-499万</option>
             <option value="500万-999万"<%if RegisteredCapital="500万-999万" then response.write " selected=""selected"""%>>500万-999万</option>
             <option value="1000万以上"<%if RegisteredCapital="1000万以上" then response.write " selected=""selected"""%>>1000万以上</option>
             </select><br/>
    联 系 人:<input name="ContactMan" type="text" value="<%=ContactMan%>" size="30" maxlength="50" emptyok="false"/><br/>
    公司地址:<input name="Address" type="text" value="<%=Address%>" size="30" maxlength="50" emptyok="false"/><br/>
    QQ号码:<input format="*N" name="QQ" type="text" value="<%=QQ%>" size="30" maxlength="50" emptyok="false"/><br/>
    手机号码:<input format="*N" name="Mobile" type="text" value="<%=Mobile%>" size="30" maxlength="50" /><br/>
    邮政编码:<input format="*N" name="ZipCode" type="text" value="<%=ZipCode%>" size="30" maxlength="10" /><br/>
    联系电话:<input name="TelPhone" type="text" value="<%=Telphone%>" size="30" maxlength="50" emptyok="false"/><br/>
    传真号码:<input name="Fax" type="text" value="<%=Fax%>" size="30" maxlength="50" /><br/>
    公司网站:<input name="WebUrl" type="text" value="<%=WebUrl%>" size="30" maxlength="50" /><br/>
    开户银行:<input name="BankAccount" type="text" value="<%=BankAccount%>" size="30" maxlength="50" /><br/>
    银行账号:<input name="AccountNumber" type="text" value="<%=AccountNumber%>" size="30" maxlength="50" /><br/>
    公司银行帐户，以方便放在您的联系资料中<br/>
    <anchor>OK,确 认<go href="User_Enterprise.asp?Action=BasicInfoSave&amp;<%=KS.WapValue%>" method="post">
    <postfield name="CompanyName" value="$(CompanyName)"/>
    <postfield name="ClassID" value="$(classid)"/>
    <postfield name="BusinessLicense" value="$(BusinessLicense)"/>
    <postfield name="LegalPeople" value="$(LegalPeople)"/>
    <postfield name="CompanyScale" value="$(CompanyScale)"/>
    <postfield name="Province" value="$(province)"/>
    <postfield name="RegisteredCapital" value="$(RegisteredCapital)"/>
    <postfield name="ContactMan" value="$(ContactMan)"/>
    <postfield name="Address" value="$(Address)"/>
    <postfield name="ZipCode" value="$(ZipCode)"/>
    <postfield name="TelPhone" value="$(TelPhone)"/>
    <postfield name="QQ" value="$(QQ)"/>
    <postfield name="Mobile" value="$(Mobile)"/>
    <postfield name="Fax" value="$(Fax)"/>
    <postfield name="WebUrl" value="$(WebUrl)"/>
    <postfield name="BankAccount" value="$(BankAccount)"/>
    <postfield name="AccountNumber" value="$(AccountNumber)"/>
    </go></anchor><br/>
    


<%
End Sub
  
Sub Intro()
%>


企业简介<br/>

<input name="Intro" type="text" value="<%=KS.HTMLEncode(Conn.Execute("Select Intro From ks_Enterprise where username='" & KSUser.UserName & "'")(0))%>"/><br/>
<anchor>OK,修 改<go href="User_Enterprise.asp?Action=IntroSave&amp;<%=KS.WapValue%>" method="post">
<postfield name="Intro" value="$(Intro)"/>
</go></anchor><br/>
<br/>
请用中文详细说明贵司的成立历史,主营产品,品牌,服务等优势<br/>
如果内容过于简单或仅填写单纯的产品介绍,将有可能无法通过审核<br/>
联系方式(电话,传真,手机,电子邮箱等)请在基本资料中填写,此处请勿重复填写<br/>
<%
End Sub
 
Sub IntroSave()
    Dim Intro
	Intro = KS.G("Intro")
	Intro=KS.HtmlCode(Intro)
	Intro=KS.HtmlEncode(Intro)
	IF Intro="" Then
	   Response.Write "对不起,你没有输入公司简介<br/>" &vbcrlf
	Else
	   Conn.Execute("Update KS_EnterPrise Set Intro='" & Intro &"' WHERE UserName='" & KSUser.UserName & "'")
	   Response.Write "企业简介修改成功!<br/>" &vbcrlf
    End If
End Sub
 
Sub Job()
%>

企业招聘<br/>

<input name="Job" type="text" value="<%=KS.HTMLEncode(Conn.Execute("Select Job From ks_Enterprise where username='" & KSUser.UserName & "'")(0))%>"/><br/>
<anchor>OK,修 改<go href="User_Enterprise.asp?Action=JobSave&amp;<%=KS.WapValue%>" method="post">
<postfield name="Job" value="$(Job)"/>
</go></anchor><br/>
<br/>

<%
End Sub
 
Sub JobSave()
    Dim Job
	Job= KS.G("Job")
	Job=KS.HtmlCode(Job)
	Job=KS.HtmlEncode(Job)
	IF Job="" Then
	   Response.Write "对不起,你没有招聘信息!<br/>" &vbcrlf
	Else
	   Conn.Execute("Update KS_EnterPrise Set Job='" & Job &"' WHERE UserName='" & KSUser.UserName & "'")
	   Response.Write "招聘信息修改成功!<br/>" &vbcrlf
	End If
End Sub
 
  
Sub BasicInfoSave() 
	   Dim CompanyName:CompanyName=KS.LoseHtml(KS.S("CompanyName"))
	   Dim Province:Province=KS.ChkClng(KS.S("Province"))
	   Dim City,SmallClassID
	   Dim Address:Address=KS.LoseHtml(KS.S("Address"))
	   Dim ZipCode:ZipCode=KS.LoseHtml(KS.S("ZipCode"))
	   Dim ContactMan:ContactMan=KS.LoseHtml(KS.S("ContactMan"))
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
	   Dim Mobile:Mobile=KS.S("Mobile")
	   Dim QQ:QQ=KS.S("QQ")
		
	   If CompanyName="" Then
		    Response.Write "公司名称必须输入!<br/>" &vbcrlf
			Response.write "<anchor>返回来源页<prev/></anchor><br/>" &vbcrlf
	   ElseIf Province=0 Then
		    Response.Write "请选择公司所在地区!<br/>" &vbcrlf
			Response.write "<anchor>返回来源页<prev/></anchor><br/>" &vbcrlf
	   ElseIf ClassID=0 Then
		    Response.Write "请选您公司所属的行业分类!<br/>" &vbcrlf
			Response.write "<anchor>返回来源页<prev/></anchor><br/>" &vbcrlf
	   Else 
            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select TOP 1 * From KS_Enterprise Where UserName='" & KSUser.UserName & "'",Conn,1,3
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
				 RSS.Open "select TOP 1 * from ks_blog where username='" & KSUser.UserName & "'",conn,1,3
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
				 
				 
			  End If
			     RS("CompanyName")=CompanyName
				 If Province<>0 Then
				 RS("Province")=Conn.Execute("Select Top  1 City From KS_Province Where ID=" & Province)(0)
				 End If
				 'RS("City")=City
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("ContactMan")=ContactMan
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
				 RS("Mobile")=Mobile
				 RS("QQ")=QQ
				 RS("ClassID")=ClassID
				 SmallClassID=RS("SmallClassID")
				 City=RS("City")
				 'RS("Intro")=KS.HtmlEncode(Request.Form("Intro"))
		 		 RS.Update
				 Conn.Execute("Update KS_User Set UserType=1 where UserName='" & KSUser.UserName & "'")
				 RS.Close
				 
				 %>
				 选择行业小类:<select name="smallclassid">
				 <option value="0">-选择行业小类-</option>
				  <%
				   Dim XML,Node
				   RS.Open "Select * From KS_EnterpriseClass Where ParentID=" & ClassID & " Order by orderid",conn,1,1
				   If Not RS.Eof Then
				    Set XML=KS.RsToXml(rs,"row","")
				   End If
				   RS.Close
				   If isObject(XML) Then
				     For Each Node In XML.DocumentElement.SelectNodes("row")
					  If Trim(SmallClassID)=Trim(Node.SelectSingleNode("@id").text) Then
					  KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """ selected=""selected"">" & Node.SelectSingleNode("@classname").text & "</option>" &vbcrlf
					  ELSE
					  KS.Echo "<option value=""" & Node.SelectSingleNode("@id").text & """>" & Node.SelectSingleNode("@classname").text & "</option>" &vbcrlf
					  End If
					 Next
				   End If
				  
				  %>
				 </select><br/>
				 选择所在城市<select name="city">
				 <option value="">-选择城市</option>
				 <%
				   RS.Open "Select * From KS_Province Where ParentID=" & Province & " Order by orderid",conn,1,1
				   If Not RS.Eof Then
				    Set XML=KS.RsToXml(rs,"row","")
				   End If
				   RS.Close
				   If isObject(XML) Then
				     For Each Node In XML.DocumentElement.SelectNodes("row")
					  If Trim(city)=Trim(Node.SelectSingleNode("@city").text) Then
					  KS.Echo "<option value=""" & Node.SelectSingleNode("@city").text & """ selected=""selected"">" & Node.SelectSingleNode("@city").text & "</option>" &vbcrlf
					  ELSE
					  KS.Echo "<option value=""" & Node.SelectSingleNode("@city").text & """>" & Node.SelectSingleNode("@city").text & "</option>" &vbcrlf
					  End If
					 Next
				   End If
				 %>
				 </select><br/>
				     <anchor>OK,提交确认<go href="User_Enterprise.asp?Action=BasicInfoSave2&amp;<%=KS.WapValue%>" method="post">
					<postfield name="SmallClassID" value="$(smallclassid)"/>
					<postfield name="City" value="$(city)"/>
					</go></anchor><br/>
				 
				 <%
	  End If
End Sub

Sub BasicInfoSave2()
  Conn.Execute("Update KS_EnterPrise set SmallClassID=" & KS.ChkClng(Request("SmallClassID")) & ",city='" & KS.S("City") & "' where username='" & KSUser.UserName &"'")
  Response.Write "恭喜,企业资料修改成功!<br/>"
End Sub
%> 

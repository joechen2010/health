<!--#include file="Kesion.SpaceCalCls.asp"-->
<!--#include file="Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class BlogCls
      Public KS,UserName,Node,Title
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KS=Nothing
	 End Sub
	 
	 '读出日志模板 FieldName 模板字段
	 Function GetTemplatePath(TemplateID,FieldName)
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select top 1 " & FieldName & " From KS_BlogTemplate Where ID=" & KS.ChkCLng(TemplateID),conn,1,1
	  If RS.Eof And RS.Bof Then
	    RS.Close
		RS.Open "Select top 1 " & FieldName & " From KS_BlogTemplate Where IsDefault='true'",conn,1,1
	  End If
	    Dim KSR:Set KSR = New Refresh 
		GetTemplatePath=KSR.LoadTemplate(RS(0))
		Set KSR=Nothing
        RS.Close:Set RS=Nothing
	 End Function

	 
	 '取得用户参数
	 Function GetUserBlogParam(UserName,FieldName)
	     Dim Num:Num=0
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select Top 1 " & FieldName & " From KS_Blog Where UserName='" & UserName & "'",conn,1,1
		 if Not RS.Eof Then
		  Num=KS.ChkClng(RS(0))
		 End if
		 RS.Close:Set RS=Nothing
		 If Num=0 Then Num=10
		 GetUserBlogParam=Num
	 End Function
	 
	 '空间头部
	 Sub LoadSpaceHead()
	     With KS
		  .echo "<html>"&vbcrlf &"<title>" & Node.SelectSingleNode("@blogname").text & "-" & Title & "</title>" &vbcrlf
		  .echo "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" &vbcrlf
          .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbcrlf
          .echo "<meta name=""generator"" content=""KesionCMS"" />" & vbcrlf
		  .echo "<meta name=""author"" content=""" & UserName & ","" />" & vbcrlf
		  .echo "<meta name=""keyword"" content=""" & Node.SelectSingleNode("@blogname").text & """ />"&VBCRLF
		  .echo "<meta name=""description"" content=""" & Node.SelectSingleNode("@descript").text & """ />"  & vbcrlf
		  .echo "<link href=""css/css.css"" type=""text/css"" rel=""stylesheet"">" & vbcrlf
		  .echo "<script src=""" & KS.GetDomain & "ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		  .echo "<script src=""js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		  .echo "<script src=""js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		 End With
	 End Sub
	 
	 '日志链接
	 Function GetLogUrl(RS)
	  GetLogUrl=GetCurrLogUrl(RS("ID"),RS("UserName"))
	 End Function
	 Function GetCurrLogUrl(ID,UserName)
	  If KS.SSetting(21)="1" Then
	  GetCurrLogUrl=KS.GetDomain &"space/list-" & Server.URLEncode(UserName) & "-" & id&KS.SSetting(22)
	  Else
	  GetCurrLogUrl="../space/?" & Server.URLEncode(UserName) & "/log/" & id
	  End If
	 End Function
	 
	 '替换用户博客所有标签
	 Function ReplaceBlogLabel(Template)
	  UserName=Node.SelectSingleNode("@username").text
	  Template=Replace(Template,"{$ShowAnnounce}",Node.SelectSingleNode("@announce").text)
	  Template=Replace(Template,"{$ShowBlogName}",Node.SelectSingleNode("@blogname").text)
	  Template=Replace(Template,"{$ShowLogo}",ReplaceLogo(Node.SelectSingleNode("@logo").text))
	  Dim b1,b2,b3,Banner:Banner=Node.SelectSingleNode("@banner").text
	  If Banner="" Or IsNull(Banner) Then Banner="|"
	  Banner=Split(Banner,"|") 
	  b1=Banner(0) : If B1="" Then b1="../images/ad1.jpg"
	  If Ubound(Banner)>=1 Then b2=Banner(1) 
	  If B2="" Then B2="../images/ad1.jpg"
	  If Ubound(Banner)>=2 Then B3=Banner(2) 
	  If B3="" Then B3="../images/ad1.jpg"
	  Template=Replace(Template,"{$ShowBannerSrc}",B1)
	  Template=Replace(Template,"{$ShowBannerSrc1}",B1)
	  Template=Replace(Template,"{$ShowBannerSrc2}",B2)
	  Template=Replace(Template,"{$ShowBannerSrc3}",B3)
	  Template=Replace(Template,"{$ShowNavigation}",ReplaceMenu)
	  Template=Replace(Template,"{$ShowUserLogin}","<iframe width=""170"" height=""122"" id=""login"" name=""login"" src=""../user/userlogin.asp"" frameBorder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>")
	    If Instr(Template,"{$ShowNewLog}")<>0 Then
		 Template=Replace(Template,"{$ShowNewLog}",GetNewLog)
		 End If
		 If Instr(Template,"{$ShowNewAlbum}")<>0 Then
		 Template=Replace(Template,"{$ShowNewAlbum}",GetNewAlbum)
		 End If
		 If Instr(Template,"{$ShowNewInfo}")<>0 Then
		 Template=Replace(Template,"{$ShowNewInfo}",GetNewXX)
		 End If
		 '=================企业空间替换==========================
		 If Instr(Template,"{$ShowNews}")<>0 Then
		 Template=Replace(Template,"{$ShowNews}",GetEnterPriseNews)
		 End If
		 If Instr(Template,"{$ShowSupply}")<>0 Then
		 Template=Replace(Template,"{$ShowSupply}",GetSupply)
		 End If
		 If Instr(Template,"{$ShowProduct}")<>0 Then
		 Template=Replace(Template,"{$ShowProduct}",GetProduct)
		 End If
		 If Instr(Template,"{$ShowProductList}")<>0 Then
		 Template=Replace(Template,"{$ShowProductList}",GetProductList)
		 End If
		 If Instr(Template,"{$ShowIntro}")<>0 Then
		 Template=Replace(Template,"{$ShowIntro}",GetEnterpriseintro)
		 End If
		 If Instr(Template,"{$ShowShortIntro}")<>0 Then
		 Template=Replace(Template,"{$ShowShortIntro}",GetEnterpriseShortintro)
		 End If
		 Template=Replace(Template,"{$ShowContact}",GetEnterpriseContact)
		 Template=Replace(Template,"{$ShowNews}",GetEnterpriseNews)
		 Template=ReplaceEnterpriseInfo(Template,UserName)

		 '========================================================
	 
	 
	   If Instr(Template,"{$ShowUserInfo}")<>0 Then
	   Template=Replace(Template,"{$ShowUserInfo}",GetUserInfo)
	   End If
	   If Instr(Template,"{$ShowCalendar}")<>0 Then
	   Template=Replace(Template,"{$ShowCalendar}",Getcalendar)
	   End If
	   If Instr(Template,"{$ShowUserClass}")<>0 Then
	   Template=Replace(Template,"{$ShowUserClass}",GetUserClass)
	   End If
	   If Instr(Template,"{$ShowComment}")<>0 Then
	   Template=Replace(Template,"{$ShowComment}",GetComment)
	   End If
	   If Instr(Template,"{$ShowMusicBox}")<>0 Then
	   Template=Replace(Template,"{$ShowMusicBox}",GetMusicBox)
	   End If
	   If Instr(Template,"{$GetMediaPlayer}")<>0 Then
	   Template=Replace(Template,"{$GetMediaPlayer}",GetMediaPlayer)
	   End If
	   If Instr(Template,"{$ShowMessage}")<>0 Then
	   Template=Replace(Template,"{$ShowMessage}",GetMessage)
	   End If
	   If Instr(Template,"{$ShowBlogInfo}")<>0 Then
	   Template=Replace(Template,"{$ShowBlogInfo}",GetBlogInfo)
	   End If
	   If Instr(Template,"{$ShowBlogTotal}")<>0 Then
	   Template=Replace(Template,"{$ShowBlogTotal}",GetBlogTotal)
	   End If
	   If Instr(Template,"{$ShowSearch}")<>0 Then
	   Template=Replace(Template,"{$ShowSearch}",GetSearch)
	   End If
	   If Instr(Template,"{$ShowVisitor}")<>0 Then
	   Template=Replace(Template,"{$ShowVisitor}",GetVisitor)
	   End If
	   Template=Replace(Template,"{$ShowXML}",GetXML)
	   Template=Replace(Template,"{$ShowUserName}",UserName)
	   Template=Replace(Template,"{$ShowSlidePhoto}",GetSlidePhoto(2))
	   
	   
	   
	   Dim KSR:Set KSR = New Refresh 
	   Template=KSR.KSLabelReplaceAll(Template)
	   Set KSR=Nothing	
		
	   ReplaceBlogLabel=Template
	 End Function	 
	 
	 Function ReplaceLogo(Logo)
	  If KS.IsNul(Logo) Then Logo="../images/logo.jpg"
	  ReplaceLogo="<Img src=""" & Logo & """ align=""absmiddle"" width=""130"">"
	 End Function
	 
	 Function ReplaceMenu() 
	   Dim HomeUrl,BlogUrl,MessageUrl,ProductUrl,IntroUrl,NewsUrl,JobUrl,RyzsUrl
	   Dim AlbumUrl,GroupUrl,FriendUrl,XXUrl,InfoUrl
	   If KS.SSetting(21)="1" Then
	    HomeUrl   = "" & server.URLEncode(username)
		BlogUrl   = "blog-" & server.URLEncode(username)
		MessageUrl= "message-"&server.URLEncode(username)
		ProductUrl= "product-"&server.URLEncode(username)
		IntroUrl  = "intro-" & server.URLEncode(username)
		NewsUrl   = "news-" & server.URLEncode(username)
		JobUrl    = "job-" & server.URLEncode(username)
		RyzsUrl   = "ryzs-" & server.URLEncode(username)
		AlbumUrl  = "album-" & server.URLEncode(username)
		GroupUrl  = "group-" & server.URLEncode(username)
		FriendUrl = "friend-" & server.URLEncode(username)
		XXUrl     = "xx-" & server.URLEncode(username)
		InfoUrl   = "info-" & server.URLEncode(username)
	   Else
	    HomeUrl   = "../space/?" & server.URLEncode(username)
		BlogUrl   = "../space/?" & server.URLEncode(username) & "/blog"
		MessageUrl= "../space/?" & server.URLEncode(username) & "/message"
		ProductUrl= "../space/?" & server.URLEncode(username) & "/product"
		IntroUrl  = "../space/?" & server.URLEncode(username) & "/intro"
		NewsUrl   = "../space/?" & server.URLEncode(username) &"/news"
		JobUrl    = "../space/?" & server.URLEncode(username) & "/job"
		RyzsUrl   = "../space/?" & server.URLEncode(username) & "/ryzs"
		AlbumUrl  = "../space/?" & server.URLEncode(username) & "/album"
		GroupUrl  = "../space/?" & server.URLEncode(username) & "/group"
		FriendUrl = "../space/?" & server.URLEncode(username) & "/friend"
		XXUrl     = "../space/?" & server.URLEncode(username) & "/xx"
		InfoUrl   = "../space/?" & server.URLEncode(username) & "/info"
	   End If
	  if conn.execute("Select top 1 username From KS_enterprise Where UserName='" & UserName & "'").eof Then
	  ReplaceMenu="<div id=""Menu"">"_
	                 & "<ul>"_
					 &" <li><a href=""" & HomeUrl & """>个人首页</a></li>"_
					 &" <li><a href=""" & BlogUrl & """);"">我的博客</a></li>"_
					 &" <li><a href=""" & AlbumUrl & """>我的相册</a></li>"_
					 &" <li><a href=""" & GroupUrl & """>我的圈子</a></li>" _
					 &" <li><a href=""" & FriendUrl & """>我的好友</a></li>"_
					 &" <li><a href=""" & XXUrl & """>我的文集</a></li>"_
					 &" <li><a href=""" & InfoUrl & """>小档案</a></li>"_
					 &" <li><a href=""" & MessageUrl & """>给我留言</a>"_
					 &"</ul>"_
					 &"</div>"
	  Else
	   	  ReplaceMenu="<div id=""Menu"">"_
	                 & "<ul>"_
					 &" <li><a href=""" & HomeUrl & """>首页</a></li>"_
					 &" <li><a href=""" & introUrl & """>公司简介</a></li>"_
					 &" <li><a href=""" & NewsUrl & """>公司动态</a></li>"_
					 &" <li><a href=""" & ProductUrl & """>产品展示</a></li>"_
					 &" <li><a href=""" & JobUrl & """>公司招聘</a></li>"_
					 &" <li><a href=""" & AlbumUrl & """>公司相册</a></li>"_
					 &" <li><a href=""" & ryzsurl & """>荣誉证书</a></li>"_
					&" <li><a href=""" & GroupUrl & """>公司圈子</a></li>" _

					 &" <li><a href=""" & BlogUrl & """>公司日志</a></li>"_
					 &" <li><a href=""" & XXUrl & """>公司文集</a></li>"_
					 &" <li><a href=""" & InfoUrl & """>联系我们</a></li>"_
					 &" <li><a href=""" & MessageUrl & """>客户留言</a>"_
					 &"</ul>"_
					 &"</div>"
	  End If
	 End Function
	 
	 
	 
	 
	 
	 '取得联信息
	 Function UserInfo()
	    Dim Str
	    Str=Location("<strong>首页 >> 联系档案</strong>")
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_User Where UserName='" & UserName & "'",conn,1,1
		If RS.Eof And RS.Bof Then
		 Str=Str& "参数传递出错!"
		 RS.Close:Set RS=Nothing
		 Exit Function
		End If
		Str=Str & ReplaceUserInfoContent(LFCls.GetConfigFromXML("space","/labeltemplate/label","userinfo"),rs)
		rs.close:set rs=nothing
		UserInfo=Str
	 End Function
	 
	 Function ReplaceUserInfoContent(ByVal Content,ByVal RS)
	    If RS("UserType")=1 Then 
		 Content=LFCls.GetConfigFromXML("space","/labeltemplate/label","companyinfo")
		 ReplaceUserInfoContent=ReplaceEnterpriseInfo(Content,RS("UserName"))
		 Exit Function
		End If
        Dim Privacy:Privacy=RS("Privacy")
        Content=Replace(Content,"{$GetUserName}",RS("UserName"))
	  Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
	  Dim FaceWidth:FaceWidth=KS.ChkClng(RS("FaceWidth"))
	  Dim FaceHeight:FaceHeight=KS.ChkClng(RS("FaceHeight"))
	  if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
		Content=Replace(Content,"{$GetUserFace}","<img src=" & UserFaceSrc & " border=""1"" width=""" & facewidth & """ height=""" & faceheight & """>")
		Content =ReplaceUserDefine(101,Content,RS)
          

		'联系方式
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetEmail}","保密")
		Else
		 Dim Email:Email=RS("Email")
		 If KS.IsNul(Email) Then Email="暂无"
		 Content=Replace(Content,"{$GetEmail}",Email)
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetQQ}","保密")
		Else
		 Dim QQ:QQ=RS("QQ")
		 If KS.IsNul(QQ) Then QQ="暂无"
		 Content=Replace(Content,"{$GetQQ}",QQ)
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetUC}","保密")
		Else
		 Dim UC:UC=RS("UC")
		 If KS.IsNul(UC) Then UC="暂无"
		 Content=Replace(Content,"{$GetUC}",UC)
		End If
		If Privacy=2 Then
		 Content=Replace(Content,"{$GetMSN}","保密")
		Else
		 Dim MSN:MSN=RS("MSN")
		 If KS.IsNul(MSN) Then MSN="暂无"
		 Content=Replace(Content,"{$GetMSN}",MSN)
		End If
    	If Privacy=2 Then
		 Content=Replace(Content,"{$GetHomePage}","保密")
		Else
		 Dim HomePage:HomePage=RS("MSN")
		 If Not IsNull(HomePage) Then
		 Content=Replace(Content,"{$GetHomePage}","<a href=""" & RS("HomePage") & """ target=""_blank"">" & RS("HomePage") & "</a>")
		 Else
		   Content=Replace(Content,"{$GetHomePage}","")
		 End iF
		End If


		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetRealName}","保密")
		Else
		 Dim RealName:RealName=RS("RealName")
		 If IsNull(RealName) Or RealName="" Then RealName="暂无"
		 Content=Replace(Content,"{$GetRealName}",RealName)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetSex}","保密")
		Else
		 Dim Sex:Sex=RS("Sex")
		 If IsNull(Sex) Or Sex="" Then Sex="暂无"
		 Content=Replace(Content,"{$GetSex}",Sex)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetBirthday}","保密")
		Else
		  Dim BirthDay:BirthDay=RS("BirthDay")
		 If IsNull(BirthDay) Or BirthDay="" Then BirthDay="暂无"
		 Content=Replace(Content,"{$GetBirthday}",BirthDay)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetIDCard}","保密")
		Else
		 Dim IDCard:IDCard=RS("IDCard")
		 If IsNull(IDCard) Or IDCard="" Then IDCard="暂无"
		 Content=Replace(Content,"{$GetIDCard}",IDCard)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetOfficeTel}","保密")
		Else
		 Dim OfficeTel:OfficeTel=RS("OfficeTel")
		 If IsNull(OfficeTel) Or OfficeTel="" Then OfficeTel="暂无"
		 Content=Replace(Content,"{$GetOfficeTel}",OfficeTel)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetHomeTel}","保密")
		Else
		 Dim HomeTel:HomeTel=RS("HomeTel")
		 If IsNull(HomeTel) Or HomeTel="" Then HomeTel="暂无"
		 Content=Replace(Content,"{$GetHomeTel}",HomeTel)
		End If

		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetMobile}","保密")
		Else
		 Dim Mobile:Mobile=RS("Mobile")
		 If IsNull(Mobile) Or Mobile="" Then Mobile="暂无"
		 Content=Replace(Content,"{$GetMobile}",Mobile)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetFax}","保密")
		Else
		 Dim Fax:Fax=RS("Fax")
		 If IsNull(Fax) Or Fax="" Then Fax="暂无"
		 Content=Replace(Content,"{$GetFax}",Fax)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetUserArea}","保密")
		Else
		 Dim Province:Province=RS("Province")
		 If IsNull(Province) Or Province="" Then Province=""
		 Dim City:City=RS("City")
		 If IsNull(City) Or Fax="" Then City="未知"
		 Content=Replace(Content,"{$GetUserArea}",Province & City)
		End If

		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetAddress}","保密")
		Else
		 Dim AddRess:AddRess=RS("AddRess")
		 If IsNull(AddRess) Or AddRess="" Then AddRess="暂无"
		 Content=Replace(Content,"{$GetAddress}",AddRess)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetZip}","保密")
		Else
		 Dim Zip:Zip=RS("Zip")
		 If IsNull(Zip) Or Zip="" Then Zip="暂无"
		 Content=Replace(Content,"{$GetZip}",ZIP)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetSign}","保密")
		Else
		 Dim Sign:Sign=RS("Sign")
		 If IsNull(Sign) Or Sign="" Then Sign="暂无"
		 Content=Replace(Content,"{$GetSign}",Sign)
		End If
        ReplaceUserInfoContent=Content
  End Function
	 
	 
	 
	  Function ReplaceEnterpriseInfo(ByVal Content,username)
	   On Error Resume Next
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select top 1 a.CompanyName as myCompanyName,BusinessLicense,profession,Companyscale,Contactman,a.ZipCode as myZipCode,a.telphone as mytelphone,a.province as myprovince,a.city as mycity,a.address as myaddress,a.fax as myfax,a.Mobile as mymobile,a.qq as myqq,a.email as myemail,weburl,bankaccount,accountnumber,b.* From KS_EnterPrise a inner join ks_user b on a.username=b.username Where a.UserName='" & UserName & "'",conn,1,1
	   IF RS.Eof Then
	    RS.Close:Set RS=Nothing
		ReplaceEnterpriseInfo=""
	   End If
	   Content=Replace(Content,"{$GetCompanyName}",RS("myCompanyName"))
	   if isnull(RS("BusinessLicense")) then
	   Content=Replace(Content,"{$GetBusinessLicense}","---")
	   else
	   Content=Replace(Content,"{$GetBusinessLicense}",RS("BusinessLicense"))
	   end if
	   if isnull(RS("profession")) then
	   Content=Replace(Content,"{$GetProfession}","---")
	   else
	   Content=Replace(Content,"{$GetProfession}",RS("profession"))
	   end if
	   if isnull(RS("Companyscale")) then
	   Content=Replace(Content,"{$GetCompanyScale}","---")
	   else
	   Content=Replace(Content,"{$GetCompanyScale}",RS("Companyscale"))
	   end if
	   if isnull(rs("myprovince")) then
	   Content=Replace(Content,"{$GetProvince}","---")
	   else
	   Content=Replace(Content,"{$GetProvince}",RS("myprovince"))
	   end if
	   if isnull(rs("mycity")) then
	   Content=Replace(Content,"{$GetCity}","---")
	   else
	   Content=Replace(Content,"{$GetCity}",RS("mycity"))
	   end if
	   if isnull(RS("Contactman")) then
	   Content=Replace(Content,"{$GetContactMan}","---")
	   else
	   Content=Replace(Content,"{$GetContactMan}",RS("Contactman"))
	   end if
	   if isnull(RS("myaddress")) then
	   Content=Replace(Content,"{$GetAddress}","---")
	   else
	   Content=Replace(Content,"{$GetAddress}",RS("myaddress"))
	   end if
	   if isnull(RS("myZipCode")) Then
	   Content=Replace(Content,"{$GetZipCode}","---")
	   Else
	   Content=Replace(Content,"{$GetZipCode}",RS("myzipcode"))
	   End If
       If Isnull(RS("mytelphone")) Then
	   Content=Replace(Content,"{$GetTelphone}","---")
	   Else
	   Content=Replace(Content,"{$GetTelphone}",RS("mytelphone"))
	   End If
	   
	   If IsNull(rs("myfax")) then
	   Content=Replace(Content,"{$GetFax}","---")
	   else
	   Content=Replace(Content,"{$GetFax}",RS("myfax"))
	   end if
	   if isnull(rs("weburl")) then
	   Content=Replace(content,"{$GetWebUrl}","---")
	   else
	   Content=Replace(Content,"{$GetWebUrl}",RS("weburl"))
	   end if
	   if isnull(rs("bankaccount")) then
	   Content=Replace(Content,"{$GetBankAccount}","---")
	   else
	   Content=Replace(Content,"{$GetBankAccount}",RS("bankaccount"))
	   end if
	   if isnull(RS("accountnumber")) then
	   Content=Replace(Content,"{$GetAccountNumber}","---")
	   else
	   Content=Replace(Content,"{$GetAccountNumber}",RS("accountnumber"))
	   end if
	   if isnull(RS("myMobile")) then
	   Content=Replace(Content,"{$GetMobile}","---")
	   else
	   Content=Replace(Content,"{$GetMobile}",RS("mymobile"))
	   end if
	   if isnull(RS("myQQ")) then
	   Content=Replace(Content,"{$GetQQ}","---")
	   else
	   Content=Replace(Content,"{$GetQQ}",RS("myQQ"))
	   end if
	   if isnull(RS("myEmail")) then
	   Content=Replace(Content,"{$GetEmail}","---")
	   else
	   Content=Replace(Content,"{$GetEmail}",RS("myEmail"))
	   end if


	   Content =ReplaceUserDefine(101,Content,RS)
	   ReplaceEnterpriseInfo=Content
	End Function
	
	 '替换自定义字段
	Function ReplaceUserDefine(ChannelID,F_C,ByVal RS)
		   If Not IsObject(Application(KS.SiteSN&"_userfiledlist"&channelid)) Then
		     Set  Application(KS.SiteSN&"_userfiledlist"&channelid)=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(KS.SiteSN&"_userfiledlist"&channelid).appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createElement("xml"))
				Dim D_F_Arr,K,Node,FieldName
				Dim KS_RS_Obj:Set KS_RS_Obj=Conn.Execute("Select FieldName From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 Order By OrderID Asc")
				If Not KS_RS_Obj.Eof Then D_F_Arr=KS_RS_Obj.GetRows(-1)
			    KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				If IsArray(D_F_Arr) Then
					  For K=0 To Ubound(D_F_Arr,2)
						Set Node=Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(1,"userfiledlist"&channelid,""))
						Node.attributes.setNamedItem(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(2,"fieldname","")).text=D_F_Arr(0,K)
					 Next
				 End If
		 End If

		 For Each Node in Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.SelectNodes("userfiledlist"&channelid)
			 FieldName=Node.selectSingleNode("@fieldname").text
			 If Left(Lcase(FieldName),3)="ks_" Then
				If Not IsNull(RS(FieldName)) Then
				  F_C=Replace(F_C,"{$" & FieldName & "}",RS(FieldName))
				Else
				  F_C=Replace(F_C,"{$" & FieldName & "}","")
				End If
			End If
		 Next

		ReplaceUserDefine=F_C
	End Function
	
	
	
	
	
	 Function GetNewAlbum()
		 Dim Xml,RS:Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select top 3 * from KS_Photoxc Where username='" & username & "' order by id desc",conn,1,1
		 If RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   GetNewAlbum="没有上传照片！"
		 else
		   Set Xml=KS.RsToXml(RS,"row","")
		   RS.Close:Set RS=Nothing
		   GetNewAlbum=GetAlbum(Xml)
		   Xml=Empty
         end if
	 End Function
		
	 Function GetAlbum(Xml)
	 	 Dim Node
		  GetAlbum="<table border=""0"" align=""center"" width=""100%"" cellpadding=""0"" cellspacing=""0"">"
		  GetAlbum=GetAlbum & "<tr>"
		   for each Node In Xml.DocumentElement.SelectNodes("row")
		    GetAlbum=GetAlbum & "<td width=""33%"" height=""22"" align=""center""> "
			GetAlbum=GetAlbum & "<table borderColor=#b2b2b2 height=149 cellSpacing=0 cellPadding=0 width=""110%"" border=0>"
			GetAlbum=GetAlbum & "<tr>"
			GetAlbum=GetAlbum & " <td align=middle width=""100%""><b><a href=""../space/?" & username & "/showalbum/" & Node.SelectSingleNode("@id").text & """>" & Node.SelectSingleNode("@xcname").text & "</a></b></td>"
			GetAlbum=GetAlbum & "</tr>"
			GetAlbum=GetAlbum & "<tr>"
			GetAlbum=GetAlbum & "		  <td align=middle width=""100%"">"
			GetAlbum=GetAlbum & "				<table style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0>"
			GetAlbum=GetAlbum & "							  <tr>"
			GetAlbum=GetAlbum & "								<td background=""images/pic.gif"" width=""136"" height=""106"" valign=""top""><a href=""../space/?" & username &"/showalbum/" & Node.SelectSingleNode("@id").text & """ target=""_blank""><img style=""margin-left:6px;margin-top:5px"" src=""" & Node.SelectSingleNode("@photourl").text & """ width=""120"" height=""90"" border=0></a></td>"
			GetAlbum=GetAlbum & "							  </tr>"
			GetAlbum=GetAlbum & "							</table>"
			GetAlbum=GetAlbum & "		  </td>"
			GetAlbum=GetAlbum & "	</tr>"
			GetAlbum=GetAlbum & "<tr>"
			GetAlbum=GetAlbum & "	  <td align=middle width=""100%"" height=20>" & Node.SelectSingleNode("@xps").text & "张/ " & Node.SelectSingleNode("@hits").text & "次<font color=red>[" &  GetStatusStr(Node.SelectSingleNode("@flag").text) & "]</font></td>"
			GetAlbum=GetAlbum & "</tr>"
			GetAlbum=GetAlbum & "</table>"
			GetAlbum=GetAlbum & "</td>"
		   Next
		   Set Node=Nothing
		   GetAlbum=GetAlbum & "</tr>"
		   GetAlbum=GetAlbum & "</table>"
	 End Function
	 
	 Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=""red"">" & GetStatusStr & "</font>"
	 End Function
	 Function GetNewLog()
		 Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select top 1 * From KS_BlogInfo Where UserName='" & UserName & "' and status=0 order by id desc",conn,1,1
		 if rs.eof then
		   GetNewLog="没有写日志！"
		 else
		   do while not rs.eof
		   GetNewLog=GetNewLog &ReplaceLogLabel(username,LFCls.GetConfigFromXML("space","/labeltemplate/label","log"),rs)
		   rs.movenext
		   loop
		 end if
		 RS.Close:Set RS=Nothing
		End Function
		
        

     Function GetNewXX()
	    GetNewXX="<span id=""xxlist""><p align=""center"">正在加载...</p></span><script>ksblog.loading('listxx',escape('" & username & "'))</script>"
	 End Function
	 
	 
	 
	 
	 Function GetEnterPriseNews()
	   Dim RS,XML,Url,Node:Set RS=Conn.Execute("Select top 10 ID,Title,AddDate From KS_EnterpriseNews where username='" & UserName & "' order by id desc")
	   If Not RS.eof Then Set Xml=KS.RsToXml(RS,"row","")
	   RS.Close:Set RS=Nothing
	   If IsObject(Xml) Then
	     GetEnterPriseNews="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For Each Node In Xml.DocumentElement.SelectNodes("row")
	      If KS.SSetting(21)="1" Then Url= "show-news-"& username & "-" & Node.SelectSingleNode("@id").text&KS.SSetting(22) Else Url="../space/?" & username & "/shownews/" & Node.SelectSingleNode("@id").text
	   	   GetEnterPriseNews =GetEnterPriseNews & "<tr><td height='22'><img src='../images/arrow_r.gif' align='absmiddle'> <a href=""" & Url & """>" & Node.SelectSingleNode("@title").text & "(" & Node.SelectSingleNode("@adddate").text & ")</a></td></tr>"
	   Next
	     Xml=Empty : Set Node=Nothing
	     GetEnterPriseNews=GetEnterPriseNews & "</table>"
	  End If
	 End Function
	 Function GetSupply()
	   Dim RS:Set RS=Conn.Execute("Select top 10 ID,Title,AddDate,TypeID,Tid,Fname From KS_GQ where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	    GetSupply="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	    GetSupply =GetSupply & "<tr><td height='22'><img src='../images/arrow_r.gif' align='absmiddle'>"& KS.GetGQTypeName(SQL(3,I)) & "<a href='" & KS.GetItemUrl(8,SQL(4,I),SQL(0,I),SQL(5,I)) & "' target='_blank'>" & SQL(1,I) &  "(" & SQL(2,I) & ")</a></td></tr>"
	   Next
	    GetSupply=GetSupply & "</table>"
	 End Function
	 Function GetProduct()
	   Dim RS:Set RS=Conn.Execute("Select top 8 ID,Title,PhotoUrl From KS_Product where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,N,k,PhotoUrl,Url,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
	    n=0
	    GetProduct="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	     GetProduct =GetProduct & "<tr>"
	     For K=1 To 4
		  PhotoUrl=sql(2,n) : If KS.SSetting(21)="1" Then Url="show-product-" &username & "-" & sql(0,n) & KS.SSetting(22) Else url="?" & UserName & "/showproduct/" & sql(0,n)
		 iF PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../Images/Nophoto.gif"
	      GetProduct =GetProduct & "<td height='22'><a href='" & Url & "' target='_blank'><img src='" & PhotoUrl & "' Width=""140"" height=""100"" border=""0""></a><div style='text-align:center'><a href='" & Url & "' target='_blank'>"& KS.Gottopic(SQL(1,N),15) & "</a></div></td>"
		 n=n+1
		 if n> Ubound(SQL,2) Then Exit For
		 Next
		 GetProduct =GetProduct & "</tr>"
		 if n> Ubound(SQL,2) Then Exit For
	   Next
	    GetProduct =GetProduct & "</table>"
	  End If
	 End Function
	 
	 Function GetProductList()
	   Dim RS:Set RS=Conn.Execute("Select top 6 ID,Title,adddate From KS_Product where verific=1 and inputer='" & UserName & "' order by id desc")
	   If RS.Eof Then RS.Close:Set RS=Nothing:Exit Function
	   Dim I,Url,SQL:Sql=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
	    GetProductList="<table border='0' cellpadding='0' cellspacing='0'>" & vbcrlf
	   For I=0 To Ubound(SQL,2)
	     If KS.SSetting(21)="1" Then Url="show-product-" &username & "-" & sql(0,i) & KS.SSetting(22) Else url="?" & UserName & "/showproduct/" & sql(0,i)
	     GetProductList=GetProductList & "<tr><td height='22'><img src='../images/arrow_r.gif' align='absmiddle'> <a href='" & Url & "' target='_blank'>"& SQL(1,i) & "(" & SQL(2,I) & ")</a></td></tr>"
		 GetProductList =GetProductList & "</tr>"
	   Next
	    GetProductList=GetProductList & "</table>"
	  End If
	 End Function
	 
	 Function GetEnterpriseintro()
	   On Error Resume Next
	   GetEnterpriseintro=KS.Htmlcode(Conn.execute("select Intro From KS_EnterPrise where UserName='" & UserName & "'")(0))
	 End Function
	 Function GetEnterpriseShortintro()
	   On Error Resume Next
	   Dim Url
	   If KS.SSetting(21)="1" Then Url="intro-" & username Else Url="../space/?" & username & "/intro"
	  GetEnterpriseShortintro=KS.Gottopic(KS.LoseHtml(KS.Htmlcode(Conn.execute("select Intro From KS_EnterPrise where UserName='" & UserName & "'")(0))),580) &"&nbsp;&nbsp;<a href=""" & Url & """>&nbsp;详细>>></a>"
	 End Function
	 
	 '幻灯显示图片
	 Function GetSlidePhoto(ChannelID)
	  Dim SQL,I,str,picarr
	  Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
	  RS.Open "Select top 6 ID,Title,Tid,InfoPurview,ReadPoint,Fname,PicUrls From " & KS.C_S(ChannelID,2)  & " Where Inputer='" & UserName & "' order by id desc",conn,1,1
	  If Not RS.Eof Then SQL=RS.GetRows()
	  RS.Close:Set RS=Nothing
	  If IsArray(SQL) Then
	    str="<script src='js/AutoChangePhoto.js'></script><div id=""divcenter_one"">" & vbcrlf
		str=str &"<div class=""divcenter_work_one"">" & vbcrlf
		str=str &"<DIV class=fpic>"

	   For I=0 To Ubound(SQL,2)
			picarr=split(split(SQL(6,I),"|||")(0),"|")
			If I=0 Then
			 str=str & "<A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=""_blank"" id=""foclnk""><img src="""&picarr(1) &""" name=""focpic"" width=""605"" id=focpic style=""FILTER: RevealTrans ( duration = 1，transition=23 ); VISIBILITY: visible; POSITION: absolute"" /></a>" &vbcrlf
			str=str & "<DIV style=""MARGIN-TOP:385px;MARGIN-left:240px;FLOAT:left;WIDTH:120px;TEXT-ALIGN: center;position:absolute""><A href=""../space/?" & UserName & "/xx"" target=_blank><font color=white>更多作品>></font></A></DIV>" &vbcrlf
			
			str=str &"<DIV id=fttltxt style=""MARGIN-TOP:390px;MARGIN-left:250px;FLOAT:left;WIDTH:120px;TEXT-ALIGN: center;position:absolute""></DIV>" &vbcrlf
			str=str & "<DIV style=""MARGIN-LEFT:590px; WIDTH: 65px"">" &vbcrlf
			str=str & "<DIV class=thubpiccur id=tmb0 onmouseover=setfoc(0); onmouseout=playit();><A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=_blank><IMG src=""" & picarr(2) & """ width=32 height=32 border=""0""></A></DIV>" &vbcrlf
            else
			 str=str & "<DIV class=thubpic id=tmb" & i & " onmouseover=setfoc("& I & "); onmouseout=playit();><A href=""../space/?" & UserName & "/showphoto/" & SQL(0,I) &""" target=_blank><img src=""" & picarr(2) & """ width=32 height=32 border=""0""></A></DIV>" &vbcrlf
			end if
	   Next
	   
	   	 str=str & "<SCRIPT language=javascript type=text/javascript>" &vbcrlf
		 str=str &"	var picarry = {};" &vbcrlf
		 str=str &" var lnkarry = {};" & vbcrlf
		 str=str &"	var ttlarry = {};"&vbcrlf
		
		For I=0 To Ubound(SQL,2)
		  picarr=split(split(SQL(6,I),"|||")(0),"|")	
		  str=str &"picarry[" & i & "] = '" & PicArr(1) & "';" & vbcrlf
		  str=str &"lnkarry[" & i & "] = '../space/?" & UserName & "/showphoto/" & SQL(0,I) &"'; "& vbcrlf
		  str=str &"ttlarry[" & i & "] = '';" & vbcrlf
		Next
		 str=str &"</SCRIPT>"
		 str=str &"</DIV>"
		 str=str &"</DIV>"
		 str=str &"</div></div>"
		 GetSlidePhoto=str
	 End If
	End Function
	
	 
	 Function GetEnterpriseContact()
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select top 1 * From KS_EnterPrise Where UserName='" & UserName & "'",conn,1,1
	   IF RS.Eof Then
	    RS.Close:Set RS=Nothing
		GetEnterpriseContact=""
		Exit Function
	   End If
	   GetEnterpriseContact="联 系 人：" & RS("Contactman") & "<br/>"
	   GetEnterpriseContact=GetEnterpriseContact & "公司地址：" & RS("address") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "邮政编码：" & RS("zipcode") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "联系电话：" & RS("telphone") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "传真号码：" & RS("fax") & "<br>"
	   GetEnterpriseContact=GetEnterpriseContact & "公司网址：" & RS("weburl") & "<br>"
	   RS.Close:Set RS=Nothing
	 End Function
	 
	 '最新访客
	 Function GetVisitor()
	    Dim user_face,Visitors,str,XML,Node
		Dim RS:Set RS=Conn.Execute("Select top 10 b.sex,a.Visitors,b.userface,a.adddate,b.isonline from KS_BlogVisitor a inner join KS_User b on a.Visitors=b.username where a.username='" & UserName & "' order by a.adddate desc ,id desc")
				If Not RS.Eof Then Set XML=KS.RsToXml(Rs,"row","")
				RS.Close:Set RS=Nothing
			    If IsObject(XML) Then
				  For Each Node In XML.DocumentElement.SelectNodes("row") 
				    user_face=Node.SelectSingleNode("@userface").text
					Visitors =Node.SelectSingleNode("@visitors").text
					If user_face="" or isnull(user_face) then 
					 if Node.SelectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
					End If
			        If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
			         str=str & "<li><a class='b' href='../space?" & Visitors & "' target='_blank'><img src='" & User_face & "' border='0'></a><br/><a href='../space?" & Visitors & "' target='_blank'>" & Visitors & "</a><br />状态:"
					 If Node.SelectSingleNode("@isonline").Text="1" Then str=str & "<font color=red>在线</font>" Else str=str & "离线"
					 str=str & "</li>"
				  Next
				  XML=Empty : Set Node=Nothing
				End If
		 GetVisitor=str
	 End Function
	 
	 
	 '用户信息
	 Function GetUserInfo()
	  Dim str,RS:Set RS=Server.CreateObject("adodb.recordset")
	  rs.open "select top 1 userface,realname,qq from ks_user where username='" & username & "'",conn,1,1
	  if not rs.eof then
	    dim userfacesrc:userfacesrc=rs(0)
		dim realname:realname=rs(1)
		if realname="" or isnull(realname) then realname=username
	    if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
	     str="<div align=""center"" style=""padding-top:5px;"">"_
		 &"<img src=""" & userfacesrc & """ style=""border:0px solid #cccccc;"" width=""170"" height=""190"" border=""0"">"_
		 &"<br /><br />"_
		 &"<div class=""userinfomenu""><li><a href=""../space/?" & username & "/message""><img border=""0"" src=""images/yes.gif"" align=""absmiddle""> 给我留言</a></li><li><a href=""javascript:void(0)"" onclick=""ksblog.addF(event,'" & UserName & "');""><img src=""images/adfriend.gif"" border=""0"" align=""absmiddle""> 加为好友</a></li><li> <a href=""javascript:void(0)"" onclick=""ksblog.sendMsg(event,'" & username & "')""><img src=""images/sendmsg.gif"" border=""0"" align=""absmiddle""> 发小纸条</a></li><li>"
		' if rs(2)<>"" and not isnull(rs(2)) then 
		' str=str &"<li><a target=blank href=tencent://message/?uin=" & rs(2) &"&Site=" & KS.Setting(2) & "&Menu=yes><img SRC=http://wpa.qq.com/pa?p=1:" & rs(2) & ":5 alt=""点击这里给我发消息"" border=""0""></a>"
		 'else
		 str=str &"<a href=""../space/?" & username & "/info""><img border=""0"" src=""images/card.gif"" align=""absmiddle""> 小档案</a>"
		' end if
		 str=str & "</li></div></div>"
	  end if
	  rs.close:set rs=nothing
	  GetUserInfo=str
	 End Function

	 'RSS订阅
	 Function GetXML()
	  GetXML="<a href=""rss.asp?UserName=" & UserName &""" target=""_blank""><img src=""../images/xml.gif"" border=""0""></a>"
	 End Function
	 '日历
	 Function Getcalendar()
	  Dim CalCls:Set CalCls=New CalendarCls
	  call CalCls.calendar(Getcalendar,username)
	  Set CalCls=Nothing
	 End Function
	 '搜索
	 Function GetSearch()
	  GetSearch="<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
	  GetSearch=GetSearch &"<form action=""../space/?" & username & "/blog"" method=""post"" name=""searchform"">" &vbcrlf
	  GetSearch=GetSearch & "<tr>" & vbcrlf
	  GetSearch=GetSearch & "<td align=""center"">关键字:<input type=""text"" size=""10"" name=""key"" style=""border-style: solid; border-width: 1px""><input type=""submit"" value="" 搜 索 ""></td>" & vbcrlf
	  GetSearch=GetSearch & "</tr>" & vbcrlf
	  GetSearch=GetSearch & "</form>"
	  GetSearch=GetSearch & "</table>" &vbcrlf
	 End Function
     '统计
	 Function GetBlogTotal()
	   GetBlogTotal="日志总数:"&conn.execute("select count(id) from ks_bloginfo where username='" & UserName &"' and status=0")(0) & " 篇"_
	   & "<br />回复总数:"&conn.execute("select count(id) from ks_blogcomment where username='" & UserName &"'")(0) & " 条"_
	   & "<br />留言总数:"&conn.execute("select count(id) from ks_blogmessage where username='" & UserName &"'")(0) & " 条"_
	   & "<br />日志阅读:"&conn.execute("select Sum(hits) from ks_blogInfo where username='" & UserName &"' and status=0")(0) & " 人次"_
	   &"<br />总访问数:" & conn.execute("select top 1 hits from ks_blog where username='" & username & "'")(0) & " 人次"
	   
	 End Function
	 '专栏列表
	 Function GetUserClass()
	  Dim Str:Str="<div style='display:none'><form id='myclassform' action='../space/?" & username & "/blog' method='post'><input type='text' name='classid' id='classid'></form></div>"
	  Dim RS:Set RS=Conn.Execute("Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' and TypeID=2")
	  Do While Not RS.Eof 
	    Str=Str & "<a href=""javascript:void(0)"" onclick=""$('#classid').val(" & RS(0) & ");$('#myclassform').submit();"">" & RS(1) & "</a><br>" & vbcrlf
		RS.MoveNext
	  Loop
	  RS.Close:Set RS=Nothing
	  GetUserClass=Str
	 End Function
	 '音乐盒
	 Function GetMusicBox()
	  GetMusicBox="<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""  width=""200"" height=""200"" id=""mp3player"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0"" ><param name=""movie"" value=""plus/mp3player.swf?config=plus/config_1.xml&file=plus/getmusiclist.asp?username=" & username & """ /><param name=""allowScriptAccess"" value=""always""><embed src=""plus/mp3player.swf?config=plus/config_1.xml&file=plus/getmusiclist.asp?username=" & username & """ allowScriptAccess=""always"" width=""200"" height=""200"" name=""mp3player""	type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" /></object>"
	 End Function
	 Function GetMediaPlayer()
	  on error resume next
	  GetMediaPlayer="<EMBED style=""WIDTH: 272px; HEIGHT: 29px"" src=""" & conn.execute("select top 1 url from ks_blogmusic where username='" & username & "'")(0) & """ width=299 height=10 type=audio/x-wav autostart=""true"" loop=""true""></DIV></EMBED>"
	 End Function
	 '最新日志
	 Function GetBlogInfo()
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select Top " & GetUserBlogParam(UserName,"ListLogNum") & " *  From KS_BlogInfo Where UserName='" & UserName & "' And Status=0 Order By ID Desc",conn,1,1
	  If Not RS.Eof Then
	    Do While Not RS.EOF
		 GetBlogInfo=GetBlogInfo & "<a title=""" & RS("UserName") & "发表于" & RS("AddDate")&""" href=""" &GetCurrLogUrl(RS("ID"),RS("UserName")) & """>" & RS("Title") & "</a><br>" & vbcrlf
		RS.MoveNext
		Loop
	  Else
	   GetBlogInfo="暂无日志!"
	  End If
	  RS.Close:Set RS=Nothing
	 End Function
	 '最新评论
	 Function GetComment()
	  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select Top " & GetUserBlogParam(UserName,"ListReplayNum") & " *  From KS_BlogComment Where UserName='" & UserName & "' Order By AddDate Desc",conn,1,1
	  If Not RS.Eof Then
	    Do While Not RS.EOF
		 GetComment=GetComment & "<img src=""../images/arrow_r.gif"" align=""absmiddle""> <a title=""" & RS("AnounName") & "发表于" & RS("AddDate")&""" href=""" &GetCurrLogUrl(RS("LogID"),RS("UserName")) & "#" & RS("ID") &""">" & RS("Title") & "</a><br />" & vbcrlf
		RS.MoveNext
		Loop
	  Else
	   GetComment="暂无评论!"
	  End If
	  RS.Close:Set RS=Nothing
	 End Function
	 '最新留言
	 Function GetMessage()
	  'GetMessage="<a href=""message.asp?UserName=" & UserName &"#write"">签写留言</a><br>"
	  Dim XML,Node,Url,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select Top " & GetUserBlogParam(UserName,"ListGuestNum") & " *  From KS_BlogMessage Where UserName='" & UserName & "' Order By AddDate Desc",conn,1,1
	  If Not RS.Eof Then Set Xml=KS.RsToXml(rs,"row","")
	  RS.Close:Set RS=Nothing
	  If IsObject(Xml) Then
	    For Each Node In Xml.DocumentElement.SelectNodes("row")
		 If KS.SSetting(21)="1" Then Url="message-" & UserName & KS.SSetting(22)&"#"& Node.SelectSingleNode("@id").text Else Url="?" & username & "/message#" & Node.SelectSingleNode("@id").text
		 GetMessage=GetMessage & "<img src=""../images/arrow_r.gif"" align=""absmiddle""> <a title=""" & Node.SelectSingleNode("@anounname").text & "留言于" & Node.SelectSingleNode("@adddate").text&""" href=""" & url &""">" & Node.SelectSingleNode("@title").text & "</a><br>"
		Next
		Xml=Empty : Set Node=Nothing
	  End If
	 End Function
	 '天气
	 Function GetWeather(RS)
	    Dim TitleStr
	    Select Case RS("Weather")
		 Case "sun.gif":TitleStr="晴天"
		 Case "sun2.gif":TitleStr="和煦"
		 Case "yin.gif":TitleStr="阴天"
		 Case "qing.gif":TitleStr="清爽"
	     Case "yun.gif":TitleStr="多云"
		 case "wu.gif":TitleStr="有雾"
		 case "xiaoyu.gif":TitleStr="小雨"
	     case "yinyu.gif":TitleStr="中雨"
		 case "leiyu.gif":TitleStr="雷雨"
		 case "caihong.gif":TitleStr="彩虹"
		 case "hexu.gif":TitleStr="酷热"
		 case "feng.gif":TitleStr="寒冷"
		 case "xue.gif":TitleStr="小雪"
		 case "daxue.gif":TitleStr="大雪"
		 case "moon.gif":TitleStr="月圆"
		 case "moon2.gif":TitleStr="月缺"
		End Select
	 	GetWeather="<img src=""../User/images/weather/" & rs("Weather") & """ title=""" & TitleStr &""" align=""absmiddle"">"
	 End Function
	 
	 Function ReplaceLogLabel(UserName,ByVal TP,RS)
		   Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" align=""absmiddle"" border=""0"">"
		   Dim MoreStr
		   MoreStr="<a href=""" & GetLogUrl(RS) & """>阅读全文("&RS("hits")&")</a> | <a href=""" & GetLogUrl(RS) & "#Comment"">回复（"& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &RS("id"))(0) &"）</a>"
		   Dim ContentStr
		    If IsNull(RS("Password")) Or RS("PassWord")="" Then 
			 ContentStr=KS.GotTopic(KS.LoseHtml(RS("Content")),KS.ChkClng(GetUserBlogParam(UserName,"ContentLen")))
			Else
			 ContentStr="<form method='post' action='" & GetLogUrl(RS) & "' target='_blank'>请输入日志的查看密码：<input style='border-style: solid; border-width: 1' type='password' name='pass' size='15'>&nbsp;<input type='submit' value=' 查看 '></form>"
			End IF
			Dim JFStr:If RS("Best")="1" then JFStr="  <img src=""../images/jh.gif"" align=""absmiddle"">" else JFStr=""
		   TP=Replace(TP,"{$ShowLogTopic}",EmotSrc&"<a href=""" & GetLogUrl(RS) & """>" & RS("Title") & "</a>" & jfstr)
		   TP=Replace(TP,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
		   TP=Replace(TP,"{$ShowLogText}",ContentStr)
		   TP=Replace(TP,"{$ShowLogMore}",MoreStr)
		   
		   TP=Replace(TP,"{$ShowTopic}",RS("Title"))
		   TP=Replace(TP,"{$ShowAuthor}",RS("UserName"))
		   TP=Replace(TP,"{$ShowAddDate}",RS("AddDate"))
		   TP=Replace(TP,"{$ShowEmot}",EmotSrc)
		   TP=Replace(TP,"{$ShowWeather}",GetWeather(RS))
		   ReplaceLogLabel=TP
		End Function

	 
	 Function Location(str)
	   Location= "<div align=""left"">"
	   Location=Location & str
	   Location=Location & " </div>"
	   Location=Location & "<hr size=1 color=#cccccc>"
	 End Function
    
	
	'=============================圈子相关标签替换=============================
	 '替换标签
	 Function ReplaceGroupLabel(RS,Template)
	  On Error Resume Next
	  Template=Replace(Template,"{$ShowAnnounce}",RS("Announce"))
	  Template=Replace(Template,"{$ShowNewUser}",GetUserList(RS("id"),"new"))
	  Template=Replace(Template,"{$ShowActiveUser}",GetUserList(RS("id"),"active"))
	  Template=Replace(Template,"{$ShowGroupInfo}",GetGroupInfo(rs))
	  Template=Replace(Template,"{$ShowNavigation}",GetGroupMenu(rs))
	  Template=Replace(Template,"{$ShowGroupName}",RS("TeamName"))
	  Template=Replace(Template,"{$ShowGroupURL}",KS.GetDomain & "space/group.asp?id=" & RS("id"))
	  Template=Replace(Template,"{$ShowUserLogin}","<iframe width=""170"" height=""122"" id=""login"" name=""login"" src=""../user/userlogin.asp"" frameBorder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>")
	  ReplaceGroupLabel=Template
	 End Function
	 
	 '圈子导航
	 Function GetGroupMenu(rs)
	  GetGroupMenu="<div id=""menu"">"_
	               &"<ul>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &""">圈子首页</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&isbest=1"">精华帖子</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=users"">成员列表</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=join"">申请加入</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=post"">发表新帖</a></li>"_
				   &"<li><a href=""group.asp?id=" & rs("id") &"&action=info"">圈子信息</a></li>"_
	 End Function
     
	 '成员列表
	Function GetUserList(teamid,Flag)
	dim orderstr
	If Flag="active" then
	  orderstr=" order by LastLoginTime desc"
	else
	  orderstr=" order by a.id desc"
	end if
	dim rs:set rs=server.createobject("adodb.recordset")
	rs.open "select top 9 a.username,b.userid,b.userface,b.facewidth,b.faceheight from ks_teamusers a,ks_user b where a.username=b.username and status=3 and teamid="& teamid & orderstr,conn,1,1
	do while not rs.eof
			  Dim UserFaceSrc:UserFaceSrc=rs("UserFace")
			 ' Dim FaceWidth:FaceWidth=KS.ChkClng(rs("FaceWidth"))
			 ' Dim FaceHeight:FaceHeight=KS.ChkClng(rs("FaceHeight"))
			  Dim FaceWidth:FaceWidth=60
			  Dim FaceHeight:FaceHeight=60			 
			  if lcase(left(userfacesrc,4))<>"http" then userfacesrc=KS.GetDomain & userfacesrc
	  GetUserList=GetUserList & "<UL class=bestuser>"
	  GetUserList=GetUserList & "<LI class=userimg><a href=""../space/?" & rs("username") &"""  target=""_blank""><img src=""" & userfacesrc & """ width=""" & facewidth & """ height=""" & faceheight & """></a>"
	  GetUserList=GetUserList & "<LI class=username><A href=""../space/?" & rs("username") & """ target=""_blank"">" & rs("username") & "</a></LI>"
	  GetUserList=GetUserList & "</UL>"
	rs.movenext
	loop
	End Function

    Function GetGroupInfo(rs)
	    GetGroupInfo="<img src=""" & rs("photourl") & """ border=""0"" width=""130"" height=""100"">"_
		             &"<br />圈子名称：" & rs("teamname")_
					 &"<br />创 建 者：" & rs("username")_
					 &"<br />创建时间：" & rs("adddate")_
					 &"<br />成员人数：" & conn.execute("select count(id) from ks_teamusers where status=3 and teamid=" & rs("id"))(0)_
					 &"<br />主题回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "and parentid=0")(0) & "/" &conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "and parentid<>0")(0) _
	End Function
	'=============================圈子相关标签替换结束==========================

End Class
%> 

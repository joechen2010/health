<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="UpFileSave.asp"-->
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="供应求购">
<p>
<%
Dim KSCls
Set KSCls = New User_MySupply
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_MySupply
        Private KS,Action
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private Verific,ChannelID,Prev
		Private GQID,ClassID,Title,Price,TypeID,ValidDate,GQContent,ContactMan,Tel,CompanyName,Address,Province,City,Email,Zip,Fax,HomePage,I,UserDefineFieldArr,UserDefineFieldValueStr,PhotoUrl,Visitor,KeyWords
		Private Sub Class_Initialize()
			MaxPerPage =9
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    ChannelID=KS.ChkClng(KS.S("ChannelID"))
			If ChannelID=0 Then ChannelID=8
			If KS.C_S(ChannelID,6)<>8 Then Response.End()
			If Conn.Execute("select Usertf from KS_Channel where ChannelID=" & ChannelID)(0)=0 Then
			   Response.Write "本频道关闭投稿!<br/>"
			   Exit Sub
			End If
			IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Sub
		    End If
			%>
            【<a href="User_MySupply.asp?Action=Add&amp;<%=KS.WapValue%>">发布供求</a>】<br/>
            <a href="User_MySupply.asp?Status=0&amp;<%=KS.WapValue%>">未审核[<%=Conn.Execute("select Count(ID) from KS_GQ where Verific=0 And UserName='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MySupply.asp?Status=1&amp;<%=KS.WapValue%>">已审核[<%=Conn.Execute("select Count(ID) from KS_GQ where Verific=1 And UserName='"& KSUser.UserName &"'")(0)%>]</a><br/>
            <%
			Action=KS.S("Action")
			Select Case Action
			    Case "DoSave"
				   Call DoSave()
				Case  "Add","Edit"
				   Call GetGQInfo()
				Case "Del"
				   Call Del()
				Case Else
				   Call Main()
		    End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上级<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub Del()
			Dim ID:ID=KS.S("ID")
			If ID="" Then
			   Response.Write "你没有选中要删除的信息!<br/>"
			   Exit Sub
			End If
		    Conn.Execute("Delete From KS_GQ Where UserName='" & KSUser.UserName & "' And verific<>1 And ID=" & ID & "")
			
		End Sub
		
		Sub Main()
			If KS.S("Page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("Page"))
			Else
			   CurrentPage = 1
			End If
			
			Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
			Verific=KS.S("Status")
			If Verific="" or not Isnumeric(Verific) Then Verific=4
			IF Verific<>4 Then 
			   Param= Param & " and Verific=" & Verific
			End If
			IF KS.S("Flag")<>"" Then
			   IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
			End if
			If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
			Dim Sql:sql = "select a.*,foldername from KS_GQ a inner join ks_class b on a.tid=b.id "& Param &" order by AddDate DESC"

			Select Case Verific
			    Case 0
				Response.Write "【待审信息】<br/>"
				Case 1
				Response.Write "【已审信息】<br/>"
				Case Else
				Response.Write "【所有信息】<br/>"
		    End Select
			
			Set RS=Server.CreateObject("AdodB.Recordset")
			RS.open sql,conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "没有你要的信息!<br/>"
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
            【信息搜索】<br/>
            <select name="Flag">
            <option value="0">标题</option>
            <option value="1">关键字</option>
            </select>
            关键字<input type="text" name="KeyWord" value="关键字" />
            <anchor>搜索<go href="User_MySupply.asp?<%=KS.WapValue%>" method="post">
            <postfield name="Flag" value="$(Flag)"/>
            <postfield name="KeyWord" value="$(KeyWord)"/>
            </go></anchor><br/>				   
			<%
		End Sub
		
		Sub ShowContent()
		    Dim I
			Do While Not RS.Eof
			%>
            <a href="../Show.asp?ID=<%=RS("ID")%>&amp;ChannelID=8&amp;<%=KS.WapValue%>">[<%=KS.GetGQTypeName(RS("TypeID"))%>]<%=KS.GotTopic(Trim(RS("Title")),25)%></a>
            <%
			Select Case rs("Verific")
			    Case 0
				Response.Write "未审 "
				Case 1
				Response.Write "已审 "
			End Select
			If RS("Verific")<>1 Then
              Response.Write "<a href=""User_MySupply.asp?Action=Edit&amp;ID="&RS("ID")&"&amp;Page="&CurrentPage&"&amp;"&KS.WapValue&""">修改</a> "
              Response.Write "<a href=""User_MySupply.asp?Action=Del&amp;ID="&RS("ID")&"&amp;"&KS.WapValue&""">删除</a>"
			Else
			   If KS.C_S(ChannelID,42)=0 Then
			      Response.Write "---"
			   Else
			      Response.Write "<a href='User_MySupply.asp?Action=Edit&amp;ChannelID=" & ChannelID & "&amp;ID" & RS("ID") &"&amp;Page=" & CurrentPage &"'>修改</a>"
			   End If
			End If
			Response.Write "<br/>"
			%>
			分类:<%=RS("FolderName")%> 时间:<%=formatdatetime(RS("AddDate"),2)%><br/>
			<%
			RS.MoveNext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			Loop
			Call KS.ShowPageParamter(totalPut, MaxPerPage, "User_MySupply.asp", True, "条信息", CurrentPage, "Status="&Verific&"&amp;"&KS.WapValue&"")
		End Sub
  
  		'添加供求信息
        Sub GetGQInfo()
		    'On Error Resume Next		
		    ChannelID=KS.ChkClng(KS.S("ChannelID"))
			If ChannelID=0 Then ChannelID=8
			Dim UserLoginTF:UserLoginTF=KSUser.UserLoginChecked
			Dim SelButton
			'自定义字段
			UserDefineFieldArr=KSUser.KS_D_F_Arr(8)
			If Action="Edit" Then
			   Response.Write "【修改信息】<br/>"
			   Dim RSE:Set RSE=Server.CreateObject("ADODB.RECORDSET")
			   Dim ID:ID=KS.ChkClng(KS.S("ID"))
			   RSE.Open "Select * From KS_GQ Where UserName='" & KSUser.UserName &"' And ID=" & ID,Conn,1,1
			   IF RSE.EOF And RSE.Bof Then
			      Response.Write "非法传递参数!<br/>"
				  Prev=True
			      Exit Sub
			   End If
			   If KS.C_S(ChannelID,42) =0 And RSE("Verific")=1 Then
			      RSE.Close():Set RSE=Nothing
				  Response.Write "本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   ClassID=RSE("Tid")
			   Title=RSE("Title")
			   Price=RSE("Price")
			   TypeID=RSE("TypeID")
			   PhotoUrl=RSE("PhotoUrl")
			   ValidDate=RSE("ValidDate")
			   GQContent=RSE("GQContent")
			   ContactMan=RSE("ContactMan")
			   Tel=RSE("Tel")
			   CompanyName=RSE("CompanyName")
			   Address=RSE("Address")
			   Province=RSE("Province")
			   City=RSE("City")
			   Email=RSE("Email")
			   Zip=RSE("Zip")
			   Fax=RSE("Fax")
			   HomePage=RSE("HomePage")
			   KeyWords=RSE("KeyWords")
			   SelButton=KS.C_C(ClassID,1)
			   If IsArray(UserDefineFieldArr) Then
			      For I=0 To Ubound(UserDefineFieldArr,2)
				      If UserDefineFieldValueStr="" Then
				         UserDefineFieldValueStr=RSE(UserDefineFieldArr(0,I)) & "||||"
				      Else
				         UserDefineFieldValueStr=UserDefineFieldValueStr & RSE(UserDefineFieldArr(0,I)) & "||||"
				      End If
			      Next
			   End If
		    Else
			   Response.Write "【发布供求】<br/>"
			   Price="可面议":ValidDate=7
			   ContactMan=KSUser.Realname:Tel=KSUser.Officetel:HomePage=KSUser.HomePage:Email=KSUser.Email
			   Fax=KSUser.Fax:Zip=KSUser.Zip:Address=KSUser.Address:Province=KSUser.Province:City=KSUser.City
			   ClassID=KS.S("ClassID")
			   If ClassID="" Then ClassID="0"
			   If ClassID="0" Then
			      SelButton="选择行业类别..."
			   Else
			      SelButton=KS.C_C(ClassID,1)
			   End If
			   UserDefineFieldValueStr=""
			End If
			
		    '上传供求图片
			If KS.ChkClng(KS.S("UpFileChecked"))=1 Then
		       Dim KSUpFile,PhotoUrl
			   Set KSUpFile = New UpFileSave
			   PhotoUrl=KSUpFile.UpFileUrl
			   Set KSUpFile = Nothing
			   '替換定义字段内容
			   If IsArray(UserDefineFieldArr) Then
			      Dim UserDefineFieldValueStrArr
			      UserDefineFieldValueStrArr=Split(UserDefineFieldValueStr,"||||")
				  UserDefineFieldValueStr=""
			      For I=0 To Ubound(UserDefineFieldArr,2)
				      If UserDefineFieldArr(0,I)=Split(PhotoUrl,"|")(0) Then
					     If UserDefineFieldValueStr="" Then
					        UserDefineFieldValueStr=Split(PhotoUrl,"|")(1) & "||||"
						 Else
						   UserDefineFieldValueStr=UserDefineFieldValueStr & Split(PhotoUrl,"|")(1) & "||||"
						 End If
						 PhotoUrl=""
					  Else
					     If UserDefineFieldValueStr="" Then
					        UserDefineFieldValueStr=UserDefineFieldValueStrArr(I) & "||||"
						 Else
						    UserDefineFieldValueStr=UserDefineFieldValueStr & UserDefineFieldValueStrArr(I) & "||||"
						 End If
					  End If
			      Next
			   End If
			End If
%>

注：请不要发布重复信息，谢谢合作<br/>
信息分类：<%Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %><br/>
信息主题：<input name="Title<%=Minute(Now)%><%=Second(Now)%>" value="<%=title%>" />*<br/>
价格说明：<input name="Price<%=Minute(Now)%><%=Second(Now)%>" value="<%=Price%>" />*<br/>

图片地址：<input name="PhotoUrl<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=PhotoUrl%>" />
<a href="User_UpFile.asp?Action=<%=Action%>&amp;ID=<%=ID%>&amp;ChannelID=8&amp;<%=KS.WapValue%>">上传图片</a>
<br/>
交易类别：<%=KS.ReturnGQType(TypeID,0)%><br/>
有 效 期：<select name="ValidDate">
           <option value="3">三天</option>
           <option value="7">一周</option>
           <option value="15">半个月</option>
           <option value="30">一个月</option>
           <option value="90">三个月</option>
           <option value="180">半年</option>
           <option value="365">一年</option>
           <option value="0">长期</option>
           </select><br/>
<%
Response.Write KSUser.KS_D_F(8,UserDefineFieldValueStr)
%>
信息内容：请详细描述您发布的供求信息<br/>
<input name="GQContent<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KS.HTMLCode(GQContent)%>" /><br/>
关键字Tags：多个关键字请用|隔开<br/>
<input name="KeyWords<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KeyWords%>" /><br/>
【联系资料】<br/>
联 系 人：<input name="ContactMan<%=Minute(Now)%><%=Second(Now)%>" value="<%=ContactMan%>" /><br/>
联系电话：<input name="Tel<%=Minute(Now)%><%=Second(Now)%>" value="<%=Tel%>" /><br/>
公司名称：<input name="CompanyName<%=Minute(Now)%><%=Second(Now)%>" value="<%=CompanyName%>" /><br/>
联系地址：<input name="Address<%=Minute(Now)%><%=Second(Now)%>" value="<%=Address%>" /><br/>
所在省份：<input name="prov<%=Minute(Now)%><%=Second(Now)%>" value="<%=Province%>" /><br/>
所在城市：<input name="city<%=Minute(Now)%><%=Second(Now)%>" value="<%=city%>" /><br/>
电子邮件：<input name="email<%=Minute(Now)%><%=Second(Now)%>" value="<%=email%>" /><br/>
邮政编码：<input name="zip<%=Minute(Now)%><%=Second(Now)%>" value="<%=zip%>" /><br/>
公司传真：<input name="fax<%=Minute(Now)%><%=Second(Now)%>" value="<%=fax%>" /><br/>
公司网址：<input name="HomePage<%=Minute(Now)%><%=Second(Now)%>" value="<%=HomePage%>" /><br/>
<anchor>确定发布<go href="User_MySupply.asp?Action=DoSave&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post">
<postfield name="ClassID" value="$(ClassID)"/>
<postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="Price" value="$(Price<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="PhotoUrl" value="$(PhotoUrl<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="TypeID" value="$(TypeID)"/>
<postfield name="ValidDate" value="$(ValidDate)"/>
<%
'自定义字段
If IsArray(UserDefineFieldArr) Then
   For I=0 To Ubound(UserDefineFieldArr,2)
       Response.Write "<postfield name=""" & UserDefineFieldArr(0,I) & """ value=""$(" & UserDefineFieldArr(0,I) & ""&Minute(Now)&Second(Now)&")""/>"
   Next
End If
%>
<postfield name="GQContent" value="$(GQContent<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="KeyWords" value="$(KeyWords<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="ContactMan" value="$(ContactMan<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="Tel" value="$(Tel<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="CompanyName" value="$(CompanyName<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="Address" value="$(Address<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="prov" value="$(prov<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="city" value="$(city<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="email" value="$(email<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="zip" value="$(zip<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="fax" value="$(fax<%=Minute(Now)%><%=Second(Now)%>)"/>
<postfield name="HomePage" value="$(HomePage<%=Minute(Now)%><%=Second(Now)%>)"/>
</go></anchor>
<br/>
        <%      
        End Sub
		
		'保存
		Sub DoSave()
			ClassID = KS.S("ClassID")
			Title        = KS.LoseHtml(KS.S("Title"))
			PhotoUrl     = KS.LoseHtml(KS.S("PhotoUrl"))
			Price        = KS.LoseHtml(KS.S("Price"))
			TypeID       = KS.S("TypeID")
			ValidDate    = KS.S("ValidDate")
			GQContent = Request.Form("GQContent")
			GQContent=KS.HtmlCode(GQContent)
			GQContent=KS.HtmlEncode(GQContent)
			
			ContactMan   = KS.LoseHtml(KS.S("ContactMan"))
			Tel          = KS.LoseHtml(KS.S("Tel"))
			CompanyName  = KS.LoseHtml(KS.S("CompanyName"))
			Address      = KS.LoseHtml(KS.S("Address"))
			Province     = KS.LoseHtml(KS.S("Prov"))
			City         = KS.LoseHtml(KS.S("City"))
			Email        = KS.LoseHtml(KS.S("Email"))
			Zip          = KS.LoseHtml(KS.S("Zip"))
			Fax          = KS.LoseHtml(KS.S("Fax"))
			HomePage     = KS.LoseHtml(KS.S("HomePage"))
			KeyWords     = KS.LoseHtml(KS.S("KeyWords"))
			Verific=0
			If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
			
			If KS.ChkClng(KS.S("ID"))<>0 Then
			   If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
			End If
			If ClassID="" or ClassID=0 Then
			   Response.Write "请选择行业类别！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If KS.strLength(Title)<=4 Then
			   Response.Write "信息标题要大于等于4个字符！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If Price="" Then
			   Response.Write "价格说明不能为空！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If TypeID="" Then
			   Response.Write "请选择交易类别！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If GQContent="" Then
			   Response.Write "信息内容必须输入！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If ContactMan="" Then
			   Response.Write "联系人不能为空！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If Tel="" Then
			   Response.Write "联系电话不能为空！<br/>"
			   Prev=True
			   Exit Sub
			End If	
			
			'自定义字段
			UserDefineFieldArr=KSUser.KS_D_F_Arr(8)
			If IsArray(UserDefineFieldArr) Then
			   For I=0 To Ubound(UserDefineFieldArr,2)
				   If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then
				      Response.Write "" & UserDefineFieldArr(1,I) & "必须填写!<br/>"
					  Prev=True
					  Exit Sub
				   End If
				   If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then
				      Response.Write "" & UserDefineFieldArr(1,I) & "必须填写数字!<br/>"
					  Prev=True
				      Exit Sub
				   End If
				   If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then
				      Response.Write "" & UserDefineFieldArr(1,I) & "必须填写正确的日期!<br/>"
					  Prev=True
					  Exit Sub
			       End If
				   If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) And UserDefineFieldArr(6,I)=1 Then
				      Response.Write "" & UserDefineFieldArr(1,I) & "必须填写正确的Email!<br/>"
					  Prev=True
					  Exit Sub
				   End If
			   Next
			End If
		    Call AddKeyTags(8,KeyWords)
			
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")		
			If KS.ChkClng(KS.S("ID"))=0 Then
			   RS.Open "Select * From [KS_Class] Where ID='" & ClassID & "'", conn, 1, 1
			   If RS.Eof And Rs.Bof Then
			      Response.Write "非法参数!<br/>"
			      Exit Sub
			   End If
			   Dim TemplateID,WapTemplateID,GQFsoType,GQFnameType
			   TemplateID=RS("Templateid")
			   WapTemplateID=RS("WapTemplateid")
			   GQFsoType=RS("FsoType")
			   GQFnameType = Trim(RS("FnameType"))
			   RS.Close
			   Dim Fname:Fname=KS.GetFileName(GQFsoType, Now, GQFnameType)
		    End If
			RS.Open "select * from KS_GQ where UserName='" & KSUser.UserName & "' And ID=" & KS.ChkClng(KS.S("ID")), Conn, 1, 3
			If RS.Eof Then
			   RS.AddNew
			   RS("Hits")=0
			   RS("AddDate")=Now
			   RS("TemplateID")=TemplateID
			   RS("WapTemplateID")=WapTemplateID
			   RS("Fname")=Fname
			   RS("Recommend")=0
			   RS("IsTop")=0
			   IF Cbool(KSUser.UserLoginChecked)=false Then	RS("UserName")="游客" Else RS("UserName")=KSUser.UserName
		    End If
			RS("Tid")=ClassID
			RS("Title")=Title
			RS("Price")=Price
			RS("PhotoUrl")=PhotoUrl
			RS("TypeID")=TypeID
			RS("ValidDate")=ValidDate
			RS("GQContent")=GQContent
			RS("KeyWords")=KeyWords
			RS("Verific")=Verific
			RS("ContactMan")=ContactMan
			RS("Tel")=Tel
			RS("CompanyName")=CompanyName
			RS("Address")=Address
			RS("Province")=Province
			RS("City")=City
			RS("Email")=Email
			RS("Zip")=Zip
			RS("Fax")=Fax
			RS("Homepage")=Homepage
			If IsArray(UserDefineFieldArr) Then
			   For I=0 To Ubound(UserDefineFieldArr,2)
				   If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
				      RS("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
				   Else
				      RS("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
				   End If
			   Next
			End If
			RS.Update
			If FName="ID" And KS.ChkClng(KS.S("ID"))=0 Then
			   RS.MoveLast
			   RS("Fname") = RS("ID") & GQFnameType
			   RS.Update
			End If
			RS.Close:Set RS=Nothing
			If KS.ChkClng(KS.S("ID"))=0 Then
			   Response.Write "信息发布成功,继续添加吗? "
			   Response.Write "<a href=""User_MySupply.asp?Action=Add&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_MySupply.asp?"&KS.WapValue&""">取消</a>"
			   Response.Write "<br/>"
			Else
		       Response.Write "信息修改成功!<br/>"
			End If
		End Sub

		Sub AddKeyTags(ChannelID,KeyWords)
		     Dim i
			 Dim TRS:set TRS=Server.Createobject("adodb.recordset")
			 Dim karr:karr=Split(KeyWords,"|")
			 For i=0 To Ubound(karr)
			     TRS.open "select * from KS_Keywords where KeyText='" & Left(Karr(i),100) & "' And ChannelID=" & ChannelID,Conn,1,3
				 If TRS.EOF Then
				    TRS.Addnew
					TRS("keytext")=left(karr(i),100)
					TRS("channelid")=channelid
					TRS("adddate")=now
					TRS.Update
				 End If
				 TRS.Close
		    Next
			set TRS=nothing
		End Sub

End Class
%> 

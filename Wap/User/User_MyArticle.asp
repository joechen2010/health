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
<card id="main" title="内容管理">
<p>
<%
Dim KSCls
Set KSCls = New MyArticleCls
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class MyArticleCls
        Private KS,ChannelID
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private Selbutton,LoginTF,Prev
		Private F_B_Arr,F_V_Arr,ClassID,Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PhotoUrl,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
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
			If ChannelID=0 Then ChannelID=1
			LoginTF=Cbool(KSUser.UserLoginChecked)
			IF LoginTF=false  Then
			   Response.redirect KS.GetDomain&"User/Login/?User_MyArticle.asp"
			   Exit Sub
			End If
			If KS.C_S(ChannelID,6)<>1 Then Response.End()
			If KS.C_S(ChannelID,36)=0 Then
			   Response.Write "本频道不允许投稿!<br/>"
			   Exit Sub
			End If
			
			F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
			F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
			%>
            <a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Action=Add&amp;<%=KS.WapValue%>">发布<%=KS.C_S(ChannelID,3)%></a><br/>
            <a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Status=2&amp;<%=KS.WapValue%>">草 稿[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Status=0&amp;<%=KS.WapValue%>">待审核[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Status=1&amp;<%=KS.WapValue%>">已审核[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Status=3&amp;<%=KS.WapValue%>">被退稿[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <br/>
			<%
			Action=KS.S("Action")
			Select Case Action
			    Case "Del"
				Call ArticleDel()
				Case "Add","Edit"
				Call DoAdd()
				Case "DoSave"
				Call DoSave()
				Case Else
				Call ArticleList()
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上级<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
	    End Sub
		
	    Sub ArticleList()
		    If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Dim Param:Param=" Where Inputer='"& KSUser.UserName &"'"
			Verific=KS.S("Status")
			If Verific="" or not isnumeric(Verific) Then Verific=4
			IF Verific<>4 Then 
			   Param= Param & " and Verific=" & Verific
			End If
			IF KS.S("Flag")<>"" Then
			   IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
			   IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
			End if
			If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
			Dim Sql:sql = "select a.*,FolderName from " & KS.C_S(ChannelID,2) &" a inner join KS_Class b On a.tid=b.id "& Param &" order by AddDate DESC"
			Select Case Verific
			    Case 0
				Response.Write "【待审" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 1
				Response.Write "【已审" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 2
				Response.Write "【草稿" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 3
				Response.Write "【退稿" & KS.C_S(ChannelID,3) & "】<br/>"
				Case Else
				Response.Write "【所有" & KS.C_S(ChannelID,3) & "】<br/>"
			End Select
			
			Set RS=Server.CreateObject("AdodB.Recordset")
			RS.open sql,conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "没有你要的" & KS.C_S(ChannelID,3) & "!<br/>"
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
			   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Call showContent
		    End If
			%>
            【<%=KS.C_S(ChannelID,3)%>搜索】<br/>
            <select name="Flag"><option value="0">标题</option><option value="1">关键字</option></select>
            关键字<input type="text" name="KeyWord" value="关键字"/>
            <anchor>搜索<go href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>" method="post">
            <postfield name="Flag" value="$(Flag)"/>
            <postfield name="KeyWord" value="$(KeyWord)"/>
            </go></anchor><br/>	
			<%
		End Sub
		
		Sub ShowContent()
		    Dim I
			Do While Not RS.Eof
			%>
            <%=((I+1)+CurrentPage*MaxPerPage)-MaxPerPage%>.
            <a href="../Show.asp?ID=<%=RS("ID")%>&amp;ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>"><%=KS.GotTopic(trim(RS("title")),35)%></a>
            <%
			Select Case RS("Verific")
			    Case 0:Response.Write "待审 "
				Case 1:Response.Write "已审 "
				Case 2:Response.Write "草稿 "
				Case 3:Response.Write "退稿 "
			End Select
			If RS("Verific")<>1 Then
			   Response.Write "<a href=""User_MyArticle.asp?ChannelID="&ChannelID&"&amp;ID="&RS("ID")&"&amp;Action=Edit&amp;page="&CurrentPage&"&amp;"&KS.WapValue&""">修改</a>"
			   Response.Write "<a href=""User_MyArticle.asp?ChannelID="&ChannelID&"&amp;Action=Del&amp;ID="&RS("ID")&"&amp;"&KS.WapValue&""">删除</a>"
			Else
			   If KS.C_S(ChannelID,42)=0 Then
			      Response.write "---"
			   Else
			      Response.Write "<a href=""User_MyArticle.asp?ChannelID=" & ChannelID & "&amp;ID=" & RS("ID") &"&amp;Action=Edit&amp;page=" & CurrentPage &"&amp;"&KS.WapValue&""">修改</a>"
			   End If
			End If
			%>
            <br/>
            分类:<%=RS("FolderName")%>
            时间:<%=formatdatetime(rs("AddDate"),2)%><br/>
  <%
			RS.MoveNext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			Loop
			Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_MyArticle.asp", True, KS.C_S(ChannelID,4), CurrentPage, "ChannelID=" & ChannelID & "&amp;Status=" & Verific & "&amp;" & KS.WapValue & "")
		End Sub
		
		'删除文章
		Sub ArticleDel()
		    Dim ID:ID=KS.S("ID")
			ID=KS.FilterIDs(ID)
			If ID="" Then Call KS.Alert("你没有选中要删除的" & KS.C_S(ChannelID,3) & "!",ComeUrl):Response.End
			Conn.Execute("Delete From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName & "' and Verific<>1 And ID In(" & ID & ")")
			Response.Redirect "User_MyArticle.asp?ChannelID="&ChannelID&"&"&KS.WapValue&""
		End Sub
		
		'添加文章
		Sub DoAdd()
			'自定义字段
			UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
			Dim ID:ID=KS.ChkClng(KS.S("ID"))
			If KS.S("Action")="Edit" Then
			   Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			   KS_A_RS_Obj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' And ID=" & ID,Conn,1,1
			   If Not KS_A_RS_Obj.Eof Then
			      If KS.C_S(ChannelID,42) =0 And KS_A_RS_Obj("Verific")=1 Then
				     KS_A_RS_Obj.Close():Set KS_A_RS_Obj=Nothing
					 Response.Write "本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!<br/>"
					 Prev=True
					 Exit Sub
				  End If
				  ClassID  = KS_A_RS_Obj("Tid")
				  Title    = KS_A_RS_Obj("Title")
				  KeyWords = KS_A_RS_Obj("KeyWords")
				  Author   = KS_A_RS_Obj("Author")
				  Origin   = KS_A_RS_Obj("Origin")
				  Content  = KS_A_RS_Obj("ArticleContent")
				  Verific  = KS_A_RS_Obj("Verific")
				  If Verific=3 Then Verific=0
				  PhotoUrl   = KS_A_RS_Obj("PhotoUrl")
				  Intro    = KS_A_RS_Obj("Intro")
				  FullTitle= KS_A_RS_Obj("FullTitle")
				  UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				  If IsArray(UserDefineFieldArr) Then
			         For I=0 To Ubound(UserDefineFieldArr,2)
				         If i=0 Then
					        UserDefineFieldValueStr=KS_A_RS_Obj(UserDefineFieldArr(0,I)) & "||||"
						 Else
					        UserDefineFieldValueStr=UserDefineFieldValueStr & KS_A_RS_Obj(UserDefineFieldArr(0,I)) & "||||"
						 End If
				     Next
			      End If
		       End If
			   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
			   SelButton=KS.C_C(ClassID,1)
		    Else
			   Author=KSUser.RealName
			   ClassID=KS.S("ClassID")
			   If ClassID="" Then ClassID="0"
			   If ClassID="0" Then
			      SelButton="选择栏目..."
			   Else
			      SelButton=KS.C_C(ClassID,1)
			   End If
		    End If

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


			IF Action="Edit" Then
			   Response.Write "【修改" & KS.C_S(ChannelID,3) & "】<br/>"
			Else
			   Response.Write "【发布" & KS.C_S(ChannelID,3) & "】<br/>"
			End iF
			Response.Write "" & F_V_Arr(1) & "："
			Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton)
			Response.Write "<br/>"
			Response.Write "" & F_V_Arr(0) & "：<input name=""Title"&Minute(Now)&Second(Now)&""" type=""text"" value=""" & Title & """ maxlength=""100"" /><br/>"
			
			If F_B_Arr(2)=1 Then
			   Response.Write "" & F_V_Arr(2) & "：<input name=""FullTitle"&Minute(Now)&Second(Now)&""" type=""text"" value="""&FullTitle&""" maxlength=""100"" /><br/>"
            End If
			
			If F_B_Arr(6)=1 Then
			   Response.Write "" & F_V_Arr(6) & "：<input name=""Author"&Minute(Now)&Second(Now)&""" type=""text"" value="""&Author&""" maxlength=""30"" /><br/>"
			End If
			
			If F_B_Arr(7)=1 Then
			Response.Write "" & F_V_Arr(7) & "：<input name=""Origin"&Minute(Now)&Second(Now)&""" type=""text"" value="""&Origin&""" maxlength=""100"" /><br/>"
			End If
			
			Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
			
			If F_B_Arr(9)=1 Then
			   Response.Write "" & F_V_Arr(9) & "：<input name=""Content"&Minute(Now)&Second(Now)&""" type=""text"" value="""&KS.HTMLCode(Content)&"""/>"
			   'If F_B_Arr(21)=1 And Cbool(LoginTF)=True Then
			      'Response.Write "<a href=""User_UpFile.asp?Type=File&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">" & F_V_Arr(21) & "</a>"
			   'End If
               Response.Write "<br/>"
			End If
			
			If F_B_Arr(10)=1 Then
			   Response.Write "" & F_V_Arr(10) & "：<input name=""PhotoUrl"" type=""text"" value="""&PhotoUrl&"""/>"
			   If F_B_Arr(11)=1 And Cbool(LoginTF)=True Then
			      Response.Write "<a href=""User_UpFile.asp?Action="&Action&"&amp;ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">"&F_V_Arr(11)&"</a>"
			   End If
			   Response.Write "<br/>"
			End If
			
			If Action="Edit" And Verific=1 Then
			   Response.Write ""&KS.C_S(ChannelID,3)&"状态：<select name=""Status""><option value=""0"">投搞</option><option value=""2"">草稿</option></select><br/>"
			End If
			%>
            <anchor>确定保存<go href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&amp;Action=DoSave&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post">
            <postfield name="ClassID" value="$(ClassID)"/>
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="FullTitle" value="$(FullTitle<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Author" value="$(Author<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Origin" value="$(Origin<%=Minute(Now)%><%=Second(Now)%>)"/>
            <%
			'自定义字段
			If IsArray(UserDefineFieldArr) Then
			   For I=0 To Ubound(UserDefineFieldArr,2)
			       Response.Write "<postfield name=""" & UserDefineFieldArr(0,I) & """ value=""$(" & UserDefineFieldArr(0,I) & ""&Minute(Now)&Second(Now)&")""/>"
			   Next
			End If
			If F_B_Arr(8)=1 Then
			   Response.Write "<postfield name=""AutoIntro"" value=""1""/>"
			End If
			%>
            <postfield name="Content" value="$(Content<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="PhotoUrl" value="$(PhotoUrl<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Status" value="$(Status<%=Minute(Now)%><%=Second(Now)%>)"/>
            </go></anchor>
            <br/>
		    <%
	    End Sub

		Sub DoSave()
		    ClassID=KS.S("ClassID")
			Title=KS.LoseHtml(KS.S("Title"))
			
			KeyWords=KS.CreateKeyWord(Title,2)'关键字
			
			Author=KS.LoseHtml(KS.S("Author"))
			Origin=KS.LoseHtml(KS.S("Origin"))
			Content = Request.Form("Content")
			Content=KS.HtmlCode(Content)
			Content=KS.HtmlEncode(Content)
			If Content="" Then Content="&nbsp;"
			Verific=KS.ChkClng(KS.S("Status"))

			FullTitle = KS.LoseHtml(KS.S("FullTitle"))
			
			If KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Request.Form("Content")),200)

			If ClassID="" or ClassID=0 Then
			   Response.Write "请选择归属栏目！<br/>"
			   Prev=True
			   Exit Sub
			End If

			Dim Fname,FnameType,TemplateID,WapTemplateID
			Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
			RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",Conn,1,1
			If RSC.Eof Then 
			   Response.end
			Else
			   FnameType=RSC("FnameType")
			   Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
			   TemplateID=RSC("TemplateID")
			   WapTemplateID=RSC("WapTemplateID")
			End If
			RSC.Close:Set RSC=Nothing
			
			If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
			If KS.ChkClng(KS.S("ID"))<>0 Then
			   If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
			End If
			
			PhotoUrl=KS.S("PhotoUrl")
			UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
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
				   If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) And UserDefineFieldArr(6,I)=1 Then
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
			
			If ClassID="" Then
			   Response.Write "你没有选择" & KS.C_S(ChannelID,3) & "栏目!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			If Title="" Then
			   Response.Write "你没有输入" & KS.C_S(ChannelID,3) & "标题!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			If Content="" And F_B_Arr(9)=1 Then
			   Response.Write "你没有输入" & KS.C_S(ChannelID,3) & "内容!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
			RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
			If RSObj.EOF Then
			   RSObj.AddNew
			   RSObj("Hits")=0
			   RSObj("TemplateID")=TemplateID
			   RSObj("WapTemplateID")=WapTemplateID
			   RSObj("Fname")=FName
			   RSObj("Adddate")=Now
			   RSObj("Rank")="★★★"
			   RSObj("Inputer")=KSUser.UserName
			End If
			RSObj("Title")=Title
			RSObj("FullTitle")=FullTitle
			RSObj("Tid")=ClassID
			RSObj("KeyWords")=KeyWords
			RSObj("Author")=Author
			RSObj("Origin")=Origin
			RSObj("ArticleContent")=Content
			RSObj("Verific")=Verific
			RSObj("PhotoUrl")=PhotoUrl
			RSObj("Intro")=Intro
			RSObj("DelTF")=0
			RSObj("Comment")=1
			If PicUrl<>"" Then 
			   RSObj("PicNews")=1
			Else
			   RSObj("PicNews")=0
			End if
			If IsArray(UserDefineFieldArr) Then
			   For I=0 To Ubound(UserDefineFieldArr,2)
			       If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
				   RSObj("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
				   Else
				   RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
				   End If
			   Next
			End If
			RSObj.Update
			If Left(Ucase(Fname),2)="ID" Then
			   RSObj.MoveLast
			   RSObj("Fname") = RSObj("ID") & FnameType
			   RSObj.Update
			End If
			If KS.C_S(ChannelID,17)=2  And KS.C_S(Channelid,7)=1 Then
			   Dim KSRObj:Set KSRObj=New Refresh
			   Call KSRObj.RefreshArticleContent(RSObj,ChannelID)
			   Set KSRobj=Nothing
			End If
			RSObj.Close:Set RSObj=Nothing
			
			If KS.ChkClng(KS.S("ID"))=0 Then
			   Response.Write "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗? "
			   Response.Write "<a href=""User_myArticle.asp?ChannelID=" & ChannelID & "&amp;Action=Add&amp;ClassID=" & ClassID &"&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_MyArticle.asp?ChannelID=" & ChannelID & "&amp;"&KS.WapValue&""">取消</a>"
			   Response.Write "<br/>"
			Else
			   Response.Write "" & KS.C_S(ChannelID,3) & "修改成功!"
			End If
		End Sub
  
End Class
%> 

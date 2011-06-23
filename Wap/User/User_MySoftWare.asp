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
Set KSCls = New User_SoftWare
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_SoftWare
        Private KS,ChannelID,F_B_Arr,F_V_Arr
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,SelButton,Prev
		Private SizeUnit,ClassID,Title,KeyWords,Author,DownLB,DownYY,DownSQ,DownSize,DownPT,YSDZ,ZCDZ,JYMM,Origin,Content,Verific,PhotoUrl,DownUrls,RSObj,ID,DownID,AddDate,CurrentOpStr,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
		Private Sub Class_Initialize()
		    MaxPerPage =9
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
			IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.redirect KS.GetDomain&"User/Login/?User_MyArticle.asp"
			   Exit Sub
			End If
			ChannelID=KS.ChkClng(KS.S("ChannelID"))
			If ChannelID=0 Then ChannelID=3
			If KS.C_S(ChannelID,6)<>3 Then Response.End()
			If Conn.Execute("select Usertf from KS_Channel where ChannelID=" & ChannelID)(0)=0 Then
			   Response.Write "本频道关闭投稿!<br/>"
			   Prev=True
			   Exit Sub
			End If
			F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
			F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
			%>
            <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Action=Add&amp;<%=KS.WapValue%>">·发布<%=KS.C_S(ChannelID,3)%></a><br/>
            <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Status=2&amp;<%=KS.WapValue%>">草 稿[<%=Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & " where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Status=0&amp;<%=KS.WapValue%>">待审核[<%=Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & " where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Status=1&amp;<%=KS.WapValue%>">&nbsp;已审核[<%=Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & " where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <a href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Status=3&amp;<%=KS.WapValue%>">被退稿[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>
            <br/>
			<%
			Action=KS.S("Action")
			Select Case Action
			    Case "Del"
				   Call SoftWareDel()
				Case "Add","Edit"
				   Call DoAdd()
				Case "AddSave","EditSave"
				   Call DoSave()
				Case Else
				   Call SoftWareList
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上级<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub SoftWareList()
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
			Dim Sql:sql = "select a.*,foldername from " & KS.C_S(ChannelID,2) & " a inner join KS_Class b on a.tid=b.id "& Param &" order by AddDate DESC"
			Select Case Verific
			    Case 0:Response.Write "【待审" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 1:Response.Write "【已审" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 2:Response.Write "【草稿" & KS.C_S(ChannelID,3) & "】<br/>"
				Case 3:Response.Write "【退稿" & KS.C_S(ChannelID,3) & "】<br/>"
				Case Else:Response.Write "【" & KS.C_S(ChannelID,3) & "】<br/>"
			End Select
			
			Set RS=Server.CreateObject("AdodB.Recordset")
			RS.open sql,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "没有你要的" & KS.C_S(ChannelID,3) & "!<br/>"
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
            【<%=KS.C_S(ChannelID,3)%>搜索】<br/>
            <select name="Flag"><option value="0">标题</option><option value="1">关键字</option></select>
            关键字<input type="text" name="KeyWord" value="关键字"/>
            <anchor>搜索<go href="User_MySoftWare.asp?ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>" method="post">
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
			   Response.Write "<a href=""User_MySoftWare.asp?ChannelID="&ChannelID&"&amp;Action=Edit&amp;id="&rs("id")&"&page="&CurrentPage&"&amp;"&KS.WapValue&""">修改</a>"
               Response.Write "<a href=""User_MySoftWare.asp?ChannelID="&ChannelID&"&amp;Action=Del&amp;ID="&rs("id")&"&amp;"&KS.WapValue&""">删除</a>"
			Else
			   If KS.C_S(ChannelID,42)=0 Then
			      Response.write "---"
			   Else
			      Response.Write "<a href=""User_MySoftWare.asp?ChannelID=" & ChannelID & "&amp;id=" & rs("id") &"&amp;Action=Edit&amp;page=" & CurrentPage &"&amp;"&KS.WapValue&""">修改</a>"
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
			Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_MySoftWare.asp", True, KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3), CurrentPage, "ChannelID=" & ChannelID & "&Status=" & Verific & "&amp;"&KS.WapValue&"")
		End Sub
		
		'删除软件
		Sub SoftWareDel()
		    Dim ID:ID=KS.S("ID")
			ID=KS.FilterIDs(ID)
			If ID="" Then
			   Response.Write "你没有选中要删除的" & KS.C_S(ChannelID,3) & "!<br/>"
			   Prev=True
			   Exit Sub
			End If
			Conn.Execute("Delete From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName &"' and verific<>1 and ID In(" & ID & ")")
			Response.Redirect "User_MySoftWare.asp?ChannelID="&ChannelID&"&"&KS.WapValue&""
		End Sub
		
		'添加
		Sub DoAdd()
		    UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)'自定义字段
		    ID=KS.ChkClng(KS.S("ID"))
			IF Action="Edit" Then
			   CurrentOpStr=" OK,修改 "
			   Action="EditSave"
			   Dim DownRS:Set DownRS=Server.CreateObject("ADODB.RECORDSET")
			   DownRS.Open "Select * From " & KS.C_S(ChannelID,2) & " Where Inputert='" & KSUser.UserName &"' and ID=" & KS.S("ID"),Conn,1,1
			   IF DownRS.Eof And DownRS.Bof Then
			      Response.Write "参数传递出错!<br/>"
				  Prev=True
			      Exit Sub
			   Else
				  If KS.C_S(ChannelID,42) =0 And DownRS("Verific")=1 Then
				     DownRS.Close():Set DownRS=Nothing
			         Response.Write "本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!<br/>"
				     Prev=True
			         Exit Sub
				  End If
				  DownID=DownRS("ID")
				  Title=DownRS("Title")
				  PhotoUrl=DownRS("PhotoUrl")
				  DownUrls=DownRS("DownUrls")
				  DownUrls=split(DownUrls,"|")(2)
				  ClassID=DownRS("TID")
				  KeyWordS=DownRS("KeyWordS")
				  DownLB=DownRS("DownLB")
				  DownYY=DownRS("DownYY")
				  DownSQ=DownRS("DownSQ")
				  DownPT=DownRS("DownPT")
				  YSDZ=DownRS("YSDZ")
				  ZCDZ=DownRS("ZCDZ")
				  JYMM=DownRS("JYMM")
				  Author=DownRS("Author")
				  Origin=DownRS("Origin")
				  Content=DownRS("DownContent")
				  AddDate=DownRS("AddDate")
				  Verific=DownRS("Verific")
				  DownSize=DownRS("DownSize")
				  SizeUnit = Right(DownSize, 2)
				  DownSize = Replace(DownSize, SizeUnit, "")
				  If DownSize = "0" Then DownSize = ""
				  UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				  If IsArray(UserDefineFieldArr) Then
				     For I=0 To Ubound(UserDefineFieldArr,2)
					     If UserDefineFieldValueStr="" Then
						    UserDefineFieldValueStr=DownRS(UserDefineFieldArr(0,I)) & "||||"
						 Else
						    UserDefineFieldValueStr=UserDefineFieldValueStr & DownRS(UserDefineFieldArr(0,I)) & "||||"
						 End If
					 Next
				  End If
			   End If
			   SelButton=KS.C_C(ClassID,1)
			   DownRS.Close:Set DownRS=Nothing
			Else
			   CurrentOpStr=" OK,添加 ":Action="AddSave":Verific=0:YSDZ="http://":ZCDZ="http://"
			   Author=KSUser.RealName
			   ClassID=KS.S("ClassID")
			   If ClassID="" Then ClassID="0"
			   If ClassID="0" Then
			      SelButton="选择栏目..."
			   Else
			      SelButton=KS.C_C(ClassID,1)
			   End If
			End IF
			
			'取得下载参数
			Dim I,DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
			Set RSP = Server.CreateObject("Adodb.RecordSet")
			RSP.Open "Select * From KS_DownParam", Conn, 1, 1
			DownLBStr = RSP("DownLB")
			DownYYStr = RSP("DownYY")
			DownSQStr = RSP("DownSQ")
			DownPTStr = RSP("DownPT")
			RSP.Close
			Set RSP = Nothing
			'下载类别
			'DownLBList="<option value="""" selected> </option>"
			LBArr = Split(DownLBStr, vbCrLf)
			For I = 0 To UBound(LBArr)
			    If LBArr(I) = DownLb Then
				   DownLBList = DownLBList & "<option value='" & LBArr(I) & "' Selected>" & LBArr(I) & "</option>"
				Else
				   DownLBList = DownLBList & "<option value='" & LBArr(I) & "'>" & LBArr(I) & "</option>"
				End If
			Next
			
			'下载语言
			'DownYYList="<option value="""" selected> </option>"
			YYArr = Split(DownYYStr, vbCrLf)
			For I = 0 To UBound(YYArr)
			    If YYArr(I) = DownYY Then
				   DownYYList = DownYYList & "<option value='" & YYArr(I) & "' Selected>" & YYArr(I) & "</option>"
				Else
				   DownYYList = DownYYList & "<option value='" & YYArr(I) & "'>" & YYArr(I) & "</option>"
				End If
			Next
			
			'下载授权
			'DownSQList="<option value="""" selected> </option>"
			SQArr = Split(DownSQStr, vbCrLf)
			For I = 0 To UBound(SQArr)
			    If SQArr(I) = DownSQ Then
				   DownSQList = DownSQList & "<option value='" & SQArr(I) & "' Selected>" & SQArr(I) & "</option>"
				Else
				   DownSQList = DownSQList & "<option value='" & SQArr(I) & "'>" & SQArr(I) & "</option>"
				End If
			Next
			
			'下载平台
			'DownPTList="<option value="""" selected> </option>"
			PTArr = Split(DownPTStr, vbCrLf)
			For I = 0 To UBound(PTArr)
			    DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
			Next
			
			IF KS.S("Action")="Edit" Then
			   Response.Write "【修改" & KS.C_S(ChannelID,3) & "】<br/>"
			Else
			   Response.Write "【发布" & KS.C_S(ChannelID,3) & "】<br/>"
			End If
			
			Response.Write "" & F_V_Arr(1) & "："
			Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton)
			Response.Write "<br/>"
			
			Response.Write "" & F_V_Arr(0) & "：<input name=""Title"" type=""text"" value="""&Title&""" maxlength=""100"" /><br/>"
			If F_B_Arr(10)=1 Then
			   Response.Write "" & F_V_Arr(10) & "：<input name=""KeyWords"" value="""&KeyWords&""" type=""text"" /><br/>多个关键字请用|隔开"
			End If
			
			If F_B_Arr(11)=1 Then
			   Response.Write "" & F_V_Arr(11) & "：<input name=""Author"" type=""text"" value="""&Author&""" maxlength=""30"" /><br/>"
			End If
			
			If F_B_Arr(12)=1 Then
			   Response.Write "" & F_V_Arr(12) & "：<input name=""Origin"" value="""&Origin&""" type=""text"" maxlength=""100"" /><br/>"
			End If
			
			If F_B_Arr(6)=1 Then
			   Response.Write "" & F_V_Arr(6) & "："
			   Response.Write "类别:<select name='DownLB'>"&DownLBList&"</select><br/>"
			   Response.Write "语言:<select name='DownYY' size='1'>"&DownYYList&"</select><br/>"
			   Response.Write "授权:<select name='DownSQ' size='1'>"&DownSQList&"</select><br/>"
			   Response.Write "大小:<input type='text' size='4' name='DownSize' value='" & DownSize & "' /><br/>"
			End If
			
			If F_B_Arr(7)=1 Then
			   'Response.Write "" & F_V_Arr(7) & "：<select name='DownPT'>"&DownPTList&"</select><br/>"
			End If
			
			If F_B_Arr(15)=1 Then
			   Response.Write "" & F_V_Arr(15) & "：<input name=""YSDZ"" type=""text"" value="""&YSDZ&""" maxlength=""100"" /><br/>"
			End If 
			
			If F_B_Arr(16)=1 Then
			   Response.Write "" & F_V_Arr(16) & "：<input name=""ZCDZ"" type=""text"" value="""&ZCDZ&""" maxlength=""100"" /><br/>"
			End If
			
			If F_B_Arr(17)=1 Then
			   Response.Write "" & F_V_Arr(17) & "：<input name=""JYMM"" type=""text"" value="""&JYMM&""" maxlength=""100"" /><br/>"
			End If
			
			Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
			
			
			If F_B_Arr(8)=1 Then
			   Response.Write "" & F_V_Arr(8) & "：<input name=""PhotoUrl"" value="""&PhotoUrl&""" type=""text"" maxlength=""100"" />"
			   If F_B_Arr(9)=1 Then
                  Response.Write "<a href=""User_UpFile.asp?Action="&Action&"&amp;ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;Type=Pic&amp;"&KS.WapValue&""">"&F_V_Arr(9)&"</a>"
			   End If
			End If
			
			Response.Write "" & KS.C_S(ChannelID,3) & "地址：<input type=""text"" class=""textbox"" name=""DownUrlS"" value=""DownUrls"" size=""50""/><br/>"
			If F_B_Arr(13)=1 Then
			   Response.Write "<a href=""User_UpFile.asp?Action="&Action&"&amp;ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">"&F_V_Arr(9)&"</a>"
			End If
			
			If F_B_Arr(14)=1 Then
			   Response.Write "" & F_V_Arr(14) & "：<input name=""Content"&Minute(Now)&Second(Now)&""" type=""text"" value="""&KS.HTMLCode(Content)&"""/><br/>"
			End If
			
			If KS.S("Action")="Edit" And Verific=1 Then
			   Response.Write "" & KS.C_S(ChannelID,3) & "状态：<select name=""Status""><option value=""0"">投搞</option><option value=""2"">草稿</option></select><br/>"
			End If
			
			%>
            <anchor><%=CurrentOpStr%><go href="User_mySoftWare.asp?ChannelID=<%=ChannelID%>&amp;Action=<%=Action%>&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post">
            
            
            </go></anchor>
            <br/>
			<%
		End Sub
  
		Sub DoSave()
		    ID=KS.ChkClng(KS.S("ID"))
			ClassID=KS.S("ClassID")
			Title=KS.LoseHtml(KS.S("Title"))
			KeyWords=KS.LoseHtml(KS.S("KeyWords"))
			Author=KS.LoseHtml(KS.S("Author"))
			DownLB=KS.LoseHtml(KS.S("DownLB"))
			DownYY=KS.LoseHtml(KS.S("DownYY"))
			DownSQ=KS.LoseHtml(KS.S("DownSQ"))
			DownSize=KS.S("DownSize")
			If DownSize = "" Or Not IsNumeric(DownSize) Then DownSize = 0
			DownSize = DownSize & KS.S("SizeUnit")
			DownPT=KS.LoseHtml(KS.S("DownPT"))
			YSDZ=KS.LoseHtml(KS.S("YSDZ"))
			ZCDZ=KS.LoseHtml(KS.S("ZCDZ"))
			JYMM=KS.LoseHtml(KS.S("JYMM"))
			Origin=KS.LoseHtml(KS.S("Origin"))
			Content = Request.Form("Content")
			Content=KS.CheckScript(KS.HtmlCode(content))
			Content=KS.HtmlEncode(Content)
			Verific=KS.ChkClng(KS.S("Status"))
			If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
			If KS.ChkClng(KS.S("ID"))<>0 Then
			   If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
			End If
			
			PhotoUrl=KS.LoseHtml(KS.S("PhotoUrl"))
			DownUrls=KS.S("DownUrls")
			
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
			If ClassID="" Then ClassID=0
			If ClassID=0 Then
			   Response.Write "你没有选择" & KS.C_S(ChannelID,3) & "栏目!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			If Title="" Then
			   Response.Write "你没有输入" & KS.C_S(ChannelID,3) & "名称!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			If DownUrls="" Then
			   Response.Write "你没有输入" & KS.C_S(ChannelID,3) & "!<br/>"
			   Prev=True
			   Exit Sub
			End IF
			Set RSObj=Server.CreateObject("Adodb.Recordset")
			Dim Fname,FnameType,TemplateID,WapTemplateID
			If ID=0 Then
			   Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
			   RSC.Open "select * from KS_Class Where ID='" & ClassID & "'",conn,1,1
			   If RSC.EOF Then 
			      Response.End
			   Else
			      FnameType=RSC("FnameType")
				  Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
				  TemplateID=RSC("TemplateID")
				  WapTemplateID=RSC("WapTemplateID")
			   End If
			   RSC.Close:Set RSC=Nothing
			End If	 
			RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName & "' and ID=" & ID,Conn,1,3
			If RSObj.EOF Then
			   RSObj.AddNew
			   RSObj("Inputer")=KSUser.UserName
			   RSObj("Hits")=0
			   RSObj("TemplateID")=TemplateID
			   RSObj("WapTemplateID")=WapTemplateID
			   RSObj("Fname")=FName
			   RSObj("AddDate")=Now
			   RSObj("Rank")="★★★"
			End If
			RSObj("Title")=Title
			RSObj("TID")=ClassID
			RSObj("KeyWords")=KeyWords
			RSObj("Author")=Author
			RSObj("DownLB")=DownLB
			RSObj("DownYY")=DownYY
			RSObj("DownSQ")=DownSQ
			RSObj("DownSize")=DownSize
			RSObj("DownPT")=DownPT
			RSObj("YSDZ")=YSDZ
			RSObj("ZCDZ")=ZCDZ
			RSObj("JYMM")=JYMM
			RSObj("Origin")=Origin
			RSObj("DownContent")=Content
			RSObj("PhotoUrl")=PhotoUrl
			RSObj("DownUrls")="0|下载地址|" & DownUrls
			RSObj("Verific")=Verific
			RSObj("DelTF")=0
			RSObj("Comment")=1
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
			If Left(Ucase(Fname),2)="ID" And ID=0 Then
			   RSObj("Fname") = RSObj("ID") & FnameType
			   RSObj.Update
			End If
			'If KS.C_S(ChannelID,17)=2  And KS.C_S(Channelid,7)=1 Then
			   'Dim KSRObj:Set KSRObj=New Refresh
			   'Call KSRObj.RefreshDownLoadContent(RSObj,ChannelID)
			   'Set KSRobj=Nothing
			'End If
			RSObj.Close:Set RSObj=Nothing
			If ID=0 Then
			   Response.Write "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?'"
			   Response.Write "<a href=""User_MySoftWare.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"&amp;"&KS.WapValue&""">确定</a> "
			   Response.Write "<a href=""User_MySoftWare.asp?ChannelID=" & ChannelID & "&amp;"&KS.WapValue&""">取消</a>"
			   Response.Write "<br/>"
			Else
			   Response.Write "" & KS.C_S(ChannelID,3) & "修改成功!<br/>"
			End If	
		End Sub

End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我发表的评论">
<p>
<%
Dim KSCls
Set KSCls = New MyComment
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class MyComment
        Private KS
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr,flag
		Private Sub Class_Initialize()
			MaxPerPage =10
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If

			Flag=KS.ChkClng(KS.S("Flag"))
			
			Select Case KS.S("Action")
			    Case "Edit","Save"
				Call CommentEdit()
				Case "Cancel"
				Conn.Execute("Delete From KS_Comment Where ID=" & KS.ChkClng(KS.S("ID")) & " And ChannelID=" & ChannelID & " And UserName='" & KSUser.UserName & "'")
				'Response.Redirect ""
				Case Else
				Call CommentList()
			End Select
            Response.Write "<br/>"
		    Response.Write "<a href=""Index.asp?"&KS.WapValue&""">我的地盘</a><br/>"
		    Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>"
		End Sub
		
		Sub CommentList()
		    Dim Param:Param=" Where UserName='" & KSUser.UserName & "'"
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			%>
            <a href="User_MyComment.asp?<%=KS.WapValue%>">所有评论(<%=Conn.Execute("Select Count(id) from KS_Comment" & Param & "")(0)%>)</a>
            <a href="User_MyComment.asp?Flag=1&amp;<%=KS.WapValue%>">已审核(<%=Conn.Execute("Select Count(id) from KS_Comment" & Param & " and verific=1")(0)%>)</a>
            <a href="User_MyComment.asp?flag=2&amp;<%=KS.WapValue%>">未审核(<%=Conn.Execute("Select Count(id) from KS_Comment" & Param & " and verific=0")(0)%>)</a><br/>
			<%
			If flag=1 Then Param=Param & " and verific=1"
			If flag=2 Then Param=Param & " and verific=0"  
			SqlStr="Select ID,Content,AddDate,Point,Verific,ChannelID,InfoID From KS_Comment" & Param & " order by adddate desc"
			Set RS=Server.CreateObject("AdodB.Recordset")
			RS.Open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "您没有发表任何评论!<br/>"
			Else
			   TotalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Call showContent
			End If
		End Sub
		
		Sub ShowContent()
		    Dim I
			Do While Not RS.EOF
			%>
            <%=((I+1)+CurrentPage*MaxPerPage)-MaxPerPage%>.<a href="User_MyComment.asp?Action=Edit&amp;ID=<%=RS(0)%>&amp;Page=<%=CurrentPage%>&amp;<%=KS.WapValue%>">评论内容：<%=KS.GotTopic(RS(1),50)%>(<%=KSUser.GetTimeFormat(RS(2))%>/<%
			If RS(4)=1 Then
			   Response.Write "已审 "
			Else
			   Response.Write "未审 "
			End If
			%>)
            </a><br/>
            <%
			Select Case KS.C_S(RS(5),6)
			    Case 1:SqlStr="Select ID,Title,Fname,Changes From " & KS.C_S(RS(5),2) & " Where ID=" & RS(6)
				Case 2:SqlStr="Select ID,Title,Fname,0 From " & KS.C_S(RS(5),2) & " Where ID=" & RS(6)
				Case 3:SqlStr="Select ID,Title,Fname,0 From " & KS.C_S(RS(5),2) & " Where ID=" & RS(6)
				Case 4:SqlStr="Select ID,Title,Fname,0 From KS_Flash Where ID=" & RS(6)
				Case 5:SqlStr="Select ID,Title,Fname,0 From KS_Product Where ID=" & RS(6)
				Case 7:SqlStr="Select ID,Title,Fname,0 From KS_Movie Where ID=" & RS(6)
				Case 8:SqlStr="Select ID,Title,Fname,0 From KS_GQ Where ID=" & RS(6)
				Case Else
				SqlStr="Select ID From KS_Article Where 1=0"
			End Select
			Dim RSI:Set RSI=Conn.Execute(SqlStr)
			If NoT RSI.Eof Then
			   Response.Write " 信息:<a href='" & KS.GetInfoUrl(RS(5),RSI(0),RSI(2)) & "'>" & RSI(1) & "</a> "
			End If
			RSI.Close:Set RSI=Nothing
			If RS(4)<>1 Then
			   Response.Write "<a href='User_MyComment.asp?Action=Edit&amp;ID=" & RS(0)& "&amp;Page=" & CurrentPage & "&amp;" & KS.WapValue & "'>修改</a> "
			Else
			   Response.Write "修改 "
			End If
			Response.Write "<a href='User_MyComment.asp?Action=Cancel&amp;ChannelID=" & RS(5) &"&amp;ID="& RS(0) &"&amp;Page=" & CurrentPage & "&amp;" & KS.WapValue & "'>删除</a>"
			Response.Write "<br/>"
			RS.MoveNext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
		 Loop
		 IF TotalPut>MaxPerPage Then
		    Call  KS.ShowPageParamter(TotalPut, MaxPerPage, "User_MyComment.asp", True, "条" & TempStr, CurrentPage, "Flag="& Flag & "&amp;" & KS.WapValue & "")
		 End IF
	End Sub
	
	Sub CommentEdit()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
		IF ID="" Or Not IsNumeric(ID) Then  ID=0
		Dim Page:Page=KS.S("Page")
		Dim RSE
		If KS.S("Action")="Save" Then
		   'Dim AnounName:AnounName=KS.S("AnounName")
		   'Dim Email:Email=KS.S("Email")
		   Dim insertface:insertface=KS.S("insertface")
		   Dim Content:Content=KS.S("Content")
		   Content=insertface&Content
		   'Dim Anonymous:Anonymous=KS.S("Anonymous")
		   'Dim Point:point=KS.S("point")
		   Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
		   If ChannelID=0 Then ChannelID=1
		   'IF point="" Or Not IsNumeric(point) Then point=0
		   If Content="" Then 
		      Response.Write "请填写评论内容!<br/>"
			  Exit Sub
		   End if
		   If Len(Content)>KS.C_S(ChannelID,14) and KS.C_S(ChannelID,14)<>0 Then
		      Response.Write "评论内容必须在" &KS.C_S(ChannelID,14) & "个字符以内!'<br/>"
			  Exit Sub
		   End if
		   IF ID="" Then
		      Response.Write "参数传递出错!<br/>"
			  Exit Sub
		   End If
		   Set RSE=Server.CreateObject("Adodb.Recordset")
		   RSE.Open "Select * From KS_Comment Where ID=" & ID,Conn,1,3
		   IF RSE.EOF AND RSE.Bof Then
		      Response.Write "参数传递出错!<br/>"
		   Else
		      'RSE("AnounName")=AnounName
			  'RSE("Email")=Email
			  RSE("Content")=Content
			  'RSE("point")=point
			  RSE.Update
		   End If
		   RSE.Close:Set RSE=Nothing
		   Response.Write "你的评论修改成功!<br/>"
		   Response.Write "<a href=""User_MyComment.asp?Page=" & Page& "&amp;" & KS.WapValue & """>返回评论</a><br/>"
		Else
		   Response.Write GetWriteComment(ID,Page)
		End IF
	End Sub
  
    Function GetWriteComment(ID,Page)
        Dim RS
		Set RS=Conn.Execute("Select * From KS_Comment Where ID=" &ID)
		IF RS.EOF AND RS.BOF Then
		   Response.Write "参数传递出错!<br/>"
		   Exit Function
		End IF
		GetWriteComment = "【修改评论】<br/>"
		Dim k,str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		Dim strArr:strArr=Split(str,"|")
		GetWriteComment = GetWriteComment & "表情:<select name=""insertface"">"
		GetWriteComment = GetWriteComment & "<option value="""">无</option>"
		For K=0 To 19
		    GetWriteComment = GetWriteComment & "<option value=""[e"&K&"]"">" & strArr(k) & "</option>"
		Next
		GetWriteComment = GetWriteComment & "</select><br/>"
		GetWriteComment = GetWriteComment & "内容:<input name=""Content"&minute(now)&second(now)&""" type=""text"" maxlength=""255"" value=""" & RS("Content") & """/><br/>"
		GetWriteComment = GetWriteComment & "<anchor>提交修改<go href=""User_MyComment.asp?ChannelID=" & RS("ChannelID") &"&amp;Action=Save&amp;ID=" & ID & "&amp;Page=" & Page & "&amp;" & KS.WapValue & """ method=""post""><postfield name=""insertface"" value=""$(insertface)""/><postfield name=""Content"" value=""$(Content"&Minute(Now)&Second(Now)&")""/></go></anchor><br/>"
		GetWriteComment = GetWriteComment & "<a href=""User_MyComment.asp?ChannelID=" & ChannelID & "&amp;Page=" & Page& "&amp;" & KS.WapValue & """>返回评论</a><br/>"
		RS.Close:Set RS=Nothing
	End Function  

End Class
%> 

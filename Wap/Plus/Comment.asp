<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card title="所有评论">
<p>
<%
Dim KS:Set KS=New PublicCls
Call KSUser.UserLoginChecked()
Dim ChannelID,InfoID,DomainStr
ChannelID=KS.Chkclng(KS.S("ChannelID"))
InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain

Select Case KS.S("Action")
	Case "write"
	   Call GetWriteComment()
	Case "writesave"
	   Call WriteSave()
	Case "support"
	   Call Support()'投票
	Case Else
	   Call CommentMain()
End Select
%>
<a href="../Show.asp?ID=<%=InfoID%>&amp;ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>">返回<%=KS.C_S(ChannelID,3)%>页</a><br/>
<%
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>"

Call CloseConn
Set KS=Nothing
Set KSUser=Nothing
%>
</p>
</card>
</wml>
<%

Sub CommentMain
    MaxPerPage=5    '每页显示评论条数
    If KS.S("page") <> "" Then
	   CurrentPage = KS.ChkClng(KS.S("page"))
	Else
	   CurrentPage = 1
	End If
    Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select ID,Title From " & KS.C_S(ChannelID,2) & " Where ID="&InfoID&"",Conn,1,1
	If Not RS.Eof Then
	   TitleLinkStr="<a href="""&DomainStr&"Show.asp?ID="&InfoID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">"&RS(1)&"</a>"
	Else
	   Exit Sub
	End If
	RS.Close
    RS.Open "Select * from KS_Comment where ChannelID="&ChannelID&" And InfoID="&InfoID&" order by AddDate desc",Conn,1,1
	If RS.EOF Then
	   Response.Write "没有找到相关评论。<br/>" &vbcrlf
	Else
	   Response.Write "["&TitleLinkStr&"]的评论,共:"&RS.RecordCount&"条评论<br/>-----------<br/>" &vbcrlf
	   TotalPut = RS.RecordCount
	   If CurrentPage < 1 Then	CurrentPage = 1
	   If (CurrentPage - 1) * MaxPerPage > totalPut Then
	      If (totalPut Mod MaxPerPage) = 0 Then
		     CurrentPage = totalPut \ MaxPerPage
		  Else
		     CurrentPage = totalPut \ MaxPerPage + 1
		  End If
	   End If
	   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
	      RS.Move (CurrentPage - 1) * MaxPerPage
	   Else
	     CurrentPage = 1
	   End If
	   Do while not RS.eof
	      If RS("AnounName")="游客" Then
	         Response.Write ""&RS("AnounName")&"说：<br/>" &vbcrlf
		  Else
		     Response.Write "<a href=""../User/ShowUser.asp?Keyword="&RS("AnounName")&"&amp;"&KS.WapValue&""">" & KS.GetUserRealName(RS("AnounName")) & "</a>说：<br/>" &vbcrlf
		  End If
	      Response.Write KS.ReplaceFace(KS.LoseHtml(RS("Content"))) &vbcrlf
		  Response.Write "("&KS.DateFormat(RS("AddDate"),17)&"发表)<br/>" &vbcrlf
		  If Session("Comment_"&RS("ID")&"1")<>""&RS("ID")&"1" Then
             Response.write "<a href="""&DomainStr&"Plus/Comment.asp?Action=support&amp;ChannelID="&ChannelID&"&amp;InfoID="&InfoID&"&amp;id="&RS("ID")&"&amp;Type=1&amp;" & KS.WapValue & """>顶("&RS("score")&")</a> " &vbcrlf
	      Else
             Response.write "顶("&RS("score")&") " &vbcrlf
	      End If
		  If Session("Comment_"&RS("ID")&"0")<>""&RS("ID")&"0" Then
             Response.write "<a href="""&DomainStr&"Plus/Comment.asp?Action=support&amp;ChannelID="&ChannelID&"&amp;InfoID="&InfoID&"&amp;id="&RS("ID")&"&amp;Type=0&amp;" & KS.WapValue & """>倒("&RS("oscore")&")</a><br/>"
		  Else
             Response.write "倒("&RS("oscore")&")<br/>" &vbcrlf
		  End If
	      RS.Movenext
		  I = I + 1
		  If I >= MaxPerPage Then Exit Do
	   Loop
	   Call KS.ShowPageParamter(totalPut, MaxPerPage, "Comment.asp", True, "个"&listtype&"", CurrentPage, "ChannelID="&ChannelID&"&amp;InfoID="&InfoID&"&amp;"&KS.WapValue&"")
	   Response.Write "<br/>" &vbcrlf
	End if
	RS.Close:set RS=Nothing
%>
-----------<br/>
<%
Dim k,str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
Dim strArr:strArr=Split(str,"|")
Dim reSayArry
reSayArry = Array("要顶!","你牛!我顶!","这个不错!该顶!","支持你!","反对你!")
Randomize
%>
<select name="insertface">
<option value="">无</option>
<%
For k=0 to 19
Response.Write "<option value=""[e"&k&"]"">" & strArr(k) & "</option>" &vbcrlf
Next
%>
</select>
<input name="C_Content<%=Minute(Now)%><%=Second(Now)%>" type="text" maxlength="<%=KS.C_S(ChannelID,14)%>" size="20" value="<%=reSayArry(Int(Ubound(reSayArry)*Rnd))%>"/>
<%
If KS.C_S(ChannelID,13)="1" Then
   Response.Write "认证码：<input name=""VerifyCode"&Minute(Now)&Second(Now)&""" type=""text"" size=""4"" /><b>" & KS.GetVerifyCode & "</b>"
End IF
%>
<anchor>发表评论<go href='Comment.asp?Action=writesave&amp;<%=KS.WapValue%>' method='post' accept-charset="utf-8">
<postfield name='ChannelID' value='<%=ChannelID%>'/>
<postfield name='InfoID' value='<%=InfoID%>'/>
<postfield name='insertface' value='$(insertface)'/>
<postfield name='C_Content' value='$(C_Content<%=Minute(Now)%><%=Second(Now)%>)'/>
<postfield name='VerifyCode' value='$(VerifyCode<%=Minute(Now)%><%=Second(Now)%>)'/>
</go></anchor>
<br/>
-----------<br/>
<%
End Sub
'首页结束=================================================================================




'投票开始=================================================================================
Sub Support()
    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	Dim OpType:OpType=KS.ChkClng(KS.S("Type"))
	Dim RS:Set RS=Server.CreateObject("Adodb.Recordset")
	RS.Open "Select * from KS_Comment Where ID="&ID&"",Conn,1,3
	If not RS.EOF Then
	   If OpType=1 Then
	      RS("Score")=RS("Score")+1
	   Else
	      RS("OScore")=RS("OScore")+1
	   End If
	   RS.UpDate
	End If
	RS.Close:Set RS=Nothing
	Session("Comment_" & ID & OpType)=ID & OpType
	Response.Redirect DomainStr&"Plus/Comment.asp?Action=CommentMain&ChannelID="&ChannelID&"&InfoID="&InfoID&"&"&KS.WapValue&""
End Sub
'投票结束=================================================================================

'保存发表开始=================================================================================
Sub WriteSave()
    Dim OutTimes,AnounName,Pass,insertface,C_Content,VerifyCode,Anonymous,point
    OutTimes = 60  '设置防刷新时间
	AnounName=KS.R(KS.S("AnounName"))	
	Pass=KS.R(KS.G("Pass"))
	'Email=KS.S("Email")
	insertface=KS.S("insertface")
	C_Content=KS.S("C_Content")
	C_Content=insertface&C_Content
	VerifyCode=KS.R(KS.S("VerifyCode"))
	Anonymous=KS.ChkClng(KS.S("Anonymous"))
	point=KS.ChkClng(KS.S("point"))
	If KS.C_S(ChannelID,13)="1" And Trim(Verifycode)<>Trim(Session("Verifycode")) Then
	   Response.Write "验证码有误，请重新输入！<br/>"
	   Response.write "<anchor>返回重写<prev/></anchor><br/><br/>"
	   Exit Sub
	End IF
	
	If Cbool(KSUser.UserLoginChecked)=False  Then
	   AnounName = "游客"
	   Anonymous = 1
	   OK="OK"
	Else
	   AnounName = KSUser.UserName
	   Locked=KSUser.Locked
	   Anonymous = 0
	End IF

	IF Anonymous=0  And Locked=1 Then
	   Response.Write "您的账号已被管理员锁定，请与管理员联系或选择游客发表!<br/><br/>"
    ElseIf InfoID="" or InStr(C_Content, "c_content") > 0 Then 
	   Response.Write "参数传递有误!<br/>"
	   Response.write "<anchor>返回重写<prev/></anchor><br/><br/>"
	ElseIf DateDiff("s", Session("OutTimes"), Now()) < OutTimes Then
       Response.Write "评论已提交，等待"&OutTimes&"秒钟后您可继续发表...<br/>"
	   Response.write "<anchor>返回上级<prev/></anchor><br/><br/>"
	ElseIf C_Content="" Then 
	   Response.write "请填写评论内容！<br/>"
	   Response.write "<anchor>返回重写<prev/></anchor><br/><br/>"
	ElseIf Len(C_Content)>KS.C_S(ChannelID,14) And KS.C_S(ChannelID,14)<>0 Then
	   Response.write "评论内容必须在" & KS.C_S(ChannelID,14) & "个字符以内!！<br/>"
	   Response.write "<anchor>返回重写<prev/></anchor><br/><br/>"
    Else
	
	
	   Set RS=Server.CreateObject("ADODB.RECORDSET")  
	   if KS.C_S(Channelid,12)=1 Or KS.C_S(Channelid,12)=3 then
	      verific=0
	   else
	      verific=1
	   end if
	   RS.Open "Select * From KS_Comment Where 1=0",Conn,1,3
	   RS.AddNew
	   RS("ChannelID")=ChannelID'频道ID
	   RS("InfoID")=InfoID'信息ID
	   RS("AnounName")=AnounName'昵称
	   RS("UserName")=AnounName'会员名
	   RS("Anonymous")=Anonymous'0匿名发布1会员发表
	   RS("Email")="418704@QQ.com"'用户邮箱
	   RS("Content")=KS.HTMLEncode(C_Content)'评论内容
	   RS("UserIP")=KS.GetIP'用户IP
	   RS("Point")=point
	   RS("Score")=0'评论得票数(支持情况)
	   RS("OScore")=0'评论反对票数(不支持情况)
	   RS("Verific")=Verific'审核与否 0未审1已审
	   RS("AddDate")=Now
	   RS.UpDate
	   RS.Close:Set RS=Nothing
	   Response.write "你的评论发表成功!<br/>"
	   Response.Write "<a href=""Comment.asp?Action=CommentMain&amp;ChannelID="&ChannelID&"&amp;InfoID="&InfoID&"&amp;" & KS.WapValue & """>查看评论</a><br/>"
	   Session("OutTimes")=Now()
    End IF
End Sub
'保存发表结束=================================================================================
%>

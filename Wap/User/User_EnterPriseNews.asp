<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="企业新闻管理">
<p>
<%
Set KS=New PublicCls
MaxPerPage =12
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If
%>
【<a href="User_EnterPriseNews.asp?Action=Add&amp;<%=KS.WapValue%>">发布新闻</a>】<br/>
<a href="User_EnterPriseNews.asp?<%=KS.WapValue%>">所有</a>
<a href="User_EnterPriseNews.asp?Status=1&amp;<%=KS.WapValue%>">待审核[<%=Conn.Execute("select count(id) from KS_EnterPrisenews where status=0 and username='"& KSUser.UserName &"'")(0)%>]</a>
<a href="User_EnterPriseNews.asp?Status=2&amp;<%=KS.WapValue%>">已审核[<%=Conn.Execute("select count(id) from ks_enterprisenews where status=1 and username='"& KSUser.UserName &"'")(0)%>]</a>
<br/>
<%
Select Case KS.S("Action")
    Case "Del"  Call ArticleDel()
	Case "Add","Edit" Call ArticleAdd()
	Case "DoSave" Call DoSave()
	Case Else Call ArticleList()
End Select

Response.write "<br/>"
Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>" &vbcrlf
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/><br/>" &vbcrlf

Call CloseConn
Set KSUser=Nothing
Set KS=Nothing
Response.Write "</p>" &vbcrlf
Response.Write "</card>" &vbcrlf
Response.Write "</wml>" &vbcrlf


Sub ArticleList()
    If KS.S("page") <> "" Then
	   CurrentPage = KS.ChkClng(KS.S("page"))
	Else
	   CurrentPage = 1
	End If
	Dim Sql,Param:Param=" where UserName='" & KSUser.UserName & "'"
	IF KS.S("Status")<>"" Then Param= Param & " and status=" & KS.ChkClng(KS.S("Status"))-1
	If (KS.S("KeyWord")<>"") Then Param = Param  & " and title like '%" & KS.S("KeyWord") & "%'"
	sql = "select * from KS_EnterPriseNews " & Param & " order by AddDate DESC"
	Response.Write "【新闻列表】<br/>" &vbcrlf
	Set RS=Server.CreateObject("AdodB.Recordset")
	RS.open sql,Conn,1,1
	If RS.EOF And RS.BOF Then
	   Response.Write "没有你要的新闻!<br/>" &vbcrlf
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
	   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
	      Rs.Move (CurrentPage - 1) * MaxPerPage
	   Else
	      CurrentPage = 1
	   End If
	   Dim I
	   Do While Not RS.Eof
	   %>
	   <a href="User_EnterPriseNews.asp?Action=Edit&amp;id=<%=rs("id")%>&amp;<%=KS.WapValue%>"><%=KS.GotTopic(trim(RS("title")),45)%></a>
       <%
	   If rs("status")=1 Then
	      Response.Write "已审核" &vbcrlf
	   Else
	      Response.Write "未审核" &vbcrlf
	   End If
	   %>
       <br/>
       <%
	   If RS("ClassID")=0 Then
	      Response.Write "没有指定分类" &vbcrlf
	   Else
	      On Error Resume Next
		  Response.Write Conn.Execute("select classname from ks_userclass where ClassID=" & RS("ClassID"))(0)
	   End If
	   %>
       <%=formatdatetime(rs("AddDate"),2)%>
       <a href="User_EnterPriseNews.asp?id=<%=rs("id")%>&amp;Action=Edit&amp;page=<%=CurrentPage%>&amp;<%=KS.WapValue%>">修改</a>
       <a href="User_EnterPriseNews.asp?action=Del&amp;ID=<%=rs("id")%>&amp;<%=KS.WapValue%>">删除</a>
       <br/>
       <%
	      RS.MoveNext
		  I = I + 1
		  If I >= MaxPerPage Then Exit Do
	   Loop
	   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Enterprisenews.asp", True, "新闻", CurrentPage, "status=" & KS.S("Status") & "&amp;" & KS.WapValue & "")
	   Response.Write "<br/>" &vbcrlf
	End If
     %>
     关键字<input type="text" name="KeyWord" value="关键字" />
     <anchor>搜 索<go href="User_EnterPriseNews.asp?<%=KS.WapValue%>" method="post">
     <postfield name="KeyWord" value="$(KeyWord)"/>
     </go></anchor><br/>
<%
End Sub
  
'删除文章
Sub ArticleDel()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then
	   Response.Write "你没有选中要删除的新闻!<br/>" &vbcrlf
    Else
	   Conn.Execute("Delete From KS_EnterPriseNews Where UserName='" & KSUser.UserName & "' and ID In(" & ID & ")")
	   Response.Write "删除成功!<br/>" &vbcrlf
	End If
	Response.Write "<a href=""User_EnterPriseNews.asp?" & KS.WapValue & """>新闻管理</a><br/>" &vbcrlf
End Sub

'添加文章
Sub ArticleAdd()
    If KS.S("Action")="Edit" Then
	   Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
	   KS_A_RS_Obj.Open "Select * From KS_EnterPriseNews Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
	   If Not KS_A_RS_Obj.Eof Then
	      Title    = KS_A_RS_Obj("Title")
		  Content  = KS_A_RS_Obj("Content")
		  AddDate  = KS_A_RS_Obj("AddDate")
		  ClassID  = KS_A_RS_Obj("ClassID")
	   End If
	   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
	Else
	   AddDate=Now:ClassID=0
	End If
	
	IF KS.S("Action")="Edit" Then
	   Response.Write "【修改新闻】<br/>" &vbcrlf
	Else
	   Response.Write "【发布新闻】<br/>" &vbcrlf
	End iF
	%> 
    新闻标题:<input name="Title" type="text" value="<%=Title%>" maxlength="100" /><br/>
    选择分类:<select name='ClassID'>
             <option value="0">不指定分类</option>
			 <%=KSUser.UserClassOption(4,ClassID)%>
             </select>	<br/>	
    发布时间:<input name="AddDate" type="text" value="<%=AddDate%>" /><br/>
    新闻内容:<input name="Content" type="text" value="<%=KS.HTMLEncode(Content)%>" /><br/>
    <anchor>保 存<go href="User_EnterPriseNews.asp?Action=DoSave&amp;ID=<%=KS.S("ID")%>&amp;<%=KS.WapValue%>" method="post">
    <postfield name="Title" value="$(Title)"/>
    <postfield name="ClassID" value="$(ClassID)"/>
    <postfield name="AddDate" value="$(AddDate)"/>
    <postfield name="Content" value="$(Content)"/>
    </go></anchor><br/>
<%
End Sub
  
Sub DoSave()
    Title=KS.LoseHtml(KS.S("Title"))
	Content=KS.S("Content")
	Dim RSObj
	If Title="" Then
	   Response.Write "你没有输入新闻标题!<br/>" &vbcrlf
	   Exit Sub
	End IF
	If Content="" Then
	   Response.Write "你没有输入新闻内容!<br/>" &vbcrlf
	   Exit Sub
	End IF
	Set RSObj=Server.CreateObject("Adodb.Recordset")
	RSObj.Open "Select * From KS_EnterpriseNews Where UserName='" & KSUser.UserName & "' And ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
	If rsobj.eof Then
	   RSObj.Addnew
	   RSObj("UserName")=KSUser.UserName
	   RSObj("Adddate")=Now
	   If KS.SSetting(18)=1 Then
	      RSObj("Status")=0
	   Else
	      RSObj("Status")=1
	   End If
	End If
	RSObj("Title")=Title
	RSObj("Content")=Content
	RSObj("ClassID")=KS.ChkClng(KS.S("ClassID"))
	RSObj.Update
	RSObj.Close:Set RSObj=Nothing
	IF KS.ChkClng(KS.S("id"))=0 Then
	   Response.Write "成功添加新闻!<br/>" &vbcrlf
	Else
	   Response.Write "新闻修改成功!<br/>" &vbcrlf
	End If
	Response.Write "<a href=""User_EnterPriseNews.asp?" & KS.WapValue & """>新闻管理</a><br/>" &vbcrlf
End Sub
%> 

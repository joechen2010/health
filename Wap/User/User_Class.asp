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
<card id="main" title="我的专栏">
<p>

<%
Set KS=New PublicCls
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If
%>
<%
ID=KS.G("ID")
Action=Trim(Request("Action"))

Select Case Action
		 Case "ClassDel"
		  Call ClassDel()'删除专栏
		 Case "Add","Edit"
		  Call ClassAdd()'添加/修改专栏
		 Case "AddSave"
		  Call AddSave()'添加专栏
		 Case "EditSave"
		  Call EditSave()'专栏修改
		 Case Else
		  Call ClassList()'列表
		End Select

Sub ClassList()
%>
<a href="User_Class.asp?action=Add&amp;<%=KS.WapValue%>">增加专栏</a>
专栏总数[<%=conn.execute("select count(classid) from ks_userclass where username='"& KSUser.UserName &"'")(0)%>]<br/>
---------<br/>
<%
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If

Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
Dim Sql:sql = "select * from KS_UserClass "& Param &" order by AddDate DESC"
Set RS=Server.CreateObject("AdodB.Recordset")
RS.open sql,conn,1,1
If RS.EOF And RS.BOF Then
   Response.write "你没有添加专栏目!<br/>"
Else
   MaxPerPage =10
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   Do While Not RS.Eof
      response.write "类型:"
	  Select Case rs("typeid")
	      Case 1 response.write "RSS订阅"
		  Case 2 Response.write "日志分类"
		  Case 3 response.write "产品分类"
		  Case 4 response.write "新闻分类"
	  End Select
	  Response.write "<br/>"
	  Response.write "专栏名称:"&KS.GotTopic(Trim(RS("ClassName")),35)&"<br/>"
	  Response.write "创建时间:"&formatdatetime(rs("AddDate"),2)&"<br/>"
	  Response.write "<a href='User_Class.asp?action=Edit&amp;id="&rs("ClassID")&"&amp;" & KS.WapValue & "'>修改</a> "
	  Response.write "<a href='User_Class.asp?action=ClassDel&amp;TypeID="&RS("TypeID")&"&amp;ID="&rs("ClassID")&"&amp;" & KS.WapValue & "'>删除</a><br/>"
	  Response.write "---------<br/>"

	  RS.MoveNext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Class.asp", True, "个专栏", CurrentPage, "Action="&Action&"&amp;" & KS.WapValue & "")
End If
Rs.close
response.Write "专栏作用：建立专栏可以给自己发表的日志、相片等归类<br/>"
response.Write "---------<br/>"
End Sub

'删除专栏=================================================================================
Sub ClassDel()
    TypeID=KS.G("TypeID")
	If ID="" Then
	   Response.write "专栏删除出！<br/>"
	Else
	   Select Case TypeID
	       Case 1
		   Conn.Execute("Delete From KS_RssUrl Where ClassID="&ID)
		   Case 2
		   Conn.Execute("Delete From KS_BlogInfo Where ClassID="&ID)
	   End Select
	   Conn.Execute("Delete From KS_UserClass Where ClassID In("&ID&")")
	   Response.write "专栏删除成功。<br/>"
	   Response.Write "<a href='User_Class.asp?action=&amp;" & KS.WapValue & "'>我的专栏</a><br/>"
	End if
End Sub

'添加专栏=================================================================================
Sub ClassAdd()
    If Action="Edit" Then
	   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select * From KS_UserClass Where ClassID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
	   If Not rs.Eof Then
	      TypeID=rs("TypeID")
		  ClassName=rs("ClassName")
		  Descript=rs("Descript")
		  OrderID=rs("OrderID")
	   End If
	   rs.Close:Set rs=Nothing
	   Action="EditSave"
	   title="修改专栏"
	   Add="确定修改"
    Else  
	   OrderID="1"
	   Action="AddSave"
	   title="创建专栏"
	   Add="确定创建"
	End If
%>
=<%=title%>=<br/>
选择类型:<select name='TypeID' >
<option value="0">选择类型</option>
<option value="2">博客日志</option>
<option value="3"<%if request("typid")="3" then response.write " selected=""selected"""%>>产品分类</option>
<option value="4">新闻分类</option>
</select>一旦选择，不能修改<br/>
专栏名称:<input emptyok="false" name="ClassName<%=minute(now)%><%=second(now)%>" type="text" maxlength="40" size="20" value="<%=ClassName%>"/><br/>
专栏序号:<input emptyok="false" name="OrderID<%=minute(now)%><%=second(now)%>" type="text" maxlength="3" size="20" value="<%=OrderID%>"/><br/>
专栏描述:<input name="Descript<%=minute(now)%><%=second(now)%>" type="text" maxlength="500" size="20" value="<%=Descript%>"/><br/>
<anchor><%=Add%><go href='User_Class.asp?Action=<%=Action%>&amp;id=<%=id%>&amp;<%=KS.WapValue%>' method='post' accept-charset="utf-8">
<postfield name='TypeID' value='$(TypeID)'/>
<postfield name='ClassName' value='$(ClassName<%=minute(now)%><%=second(now)%>)'/>
<postfield name='OrderID' value='$(OrderID<%=minute(now)%><%=second(now)%>)'/>
<postfield name='Descript' value='$(Descript<%=minute(now)%><%=second(now)%>)'/>
</go></anchor>
<br/>
<%
End Sub

'专栏修改=================================================================================
Sub EditSave()
     TypeID=KS.G("TypeID")
	 ClassName=KS.S("ClassName")
	 OrderID=KS.G("OrderID")
	 Descript=KS.S("Descript")
	 If TypeID="" Then TypeID=0
	 If ClassName="" Then
	    Response.Write "你没有输入标题!<br/><anchor><prev/>还回上级</anchor><br/>"
	 ElseIF OrderID="" Then
	    Response.Write "你没有输入栏目序号!<br/><anchor><prev/>还回上级</anchor><br/>"
	 ElseIF Not Isnumeric(OrderID) Then
	    Response.Write "栏目序号只能填写数字!<br/><anchor><prev/>还回上级</anchor><br/>"
	 Else
        set rs=server.createobject("adodb.recordset")
		sql="select * from KS_UserClass Where ClassID="&id&""
		rs.open sql,conn,1,3
		rs("ClassName")=ClassName
		'rs("TypeID")=TypeID
		rs("OrderID")=OrderID
		rs("Descript")=Descript
		rs.Update
		rs.Close:Set rs=Nothing
		Response.write "专栏修改成功。<br/>"
		Response.Write "<a href='User_Class.asp?action=&amp;" & KS.WapValue & "'>我的专栏</a><br/>"
	 End IF
End Sub

'添加专栏=================================================================================
Sub AddSave()
      TypeID=KS.G("TypeID")
	  ClassName=KS.S("ClassName")
	  OrderID=KS.G("OrderID")
	  Descript=KS.S("Descript")
	  If TypeID="" Then TypeID=0
	  If TypeID=0 Then
	     Response.Write "你没有选择类型!<br/><anchor><prev/>还回上级</anchor><br/>"
	  ElseIF ClassName="" Then
	     Response.Write "你没有输入标题!<br/><anchor><prev/>还回上级</anchor><br/>"
	  ElseIF OrderID="" Then
	     Response.Write "你没有输入栏目序号!<br/><anchor><prev/>还回上级</anchor><br/>"
	  ElseIF Not Isnumeric(OrderID) Then
	     Response.Write "栏目序号只能填写数字!<br/><anchor><prev/>还回上级</anchor><br/>"
	  Else
         Set rs=Server.CreateObject("Adodb.Recordset")
		 rs.Open "Select * From KS_UserClass",Conn,1,3
		 rs.AddNew
		 rs("ClassName")=ClassName
		 rs("TypeID")=TypeID
		 rs("OrderID")=OrderID
		 rs("Descript")=Descript
		 rs("UserName")=KSUser.UserName
		 rs("Adddate")=Now
		 rs.Update
		 rs.Close:Set rs=Nothing
		 Response.write "添加专栏成功。<br/>"
		 Response.Write "<a href='?action=&amp;" & KS.WapValue & "'>我的专栏</a><br/>"
	  End IF
End Sub
%>
<br/>
<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
<%
Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p>
</card>
</wml>

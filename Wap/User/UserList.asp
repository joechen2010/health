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
<card id="main" title="所有注册会员">
<p>
<%
Dim KSCls
Set KSCls = New UserList
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class UserList
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		    CloseConn
		End Sub
		
		Public Sub Kesion()
		    %>
            
            <a href="UserList.asp?ListType=1&amp;<%=KS.WapValue%>">ID排序</a>
            <a href="UserList.asp?ListType=2&amp;<%=KS.WapValue%>">注册日期</a>
            <a href="UserList.asp?ListType=3&amp;<%=KS.WapValue%>">登录次数</a>
            <br/>【所有会员】<br/>
			<%
		    Response.Write GetUserList
			Response.Write "<br/>"
			IF Cbool(KSUser.UserLoginChecked)=True Then 
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			End If
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
	    End Sub
		
		Function GetUserList()
  		    Dim  CurrentPage,totalPut,RS,MaxPerPage,SqlStr,ListType,Param
			ListType=KS.ChkClng(KS.S("ListType"))
			MaxPerPage =15
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			
			Set RS=Server.CreateObject("Adodb.Recordset")
			If ListType=1 Then
			   Param="Order By UserID Desc"
			ElseIF ListType=2 Then
			   Param="Order By LastLoginTime Desc"
			ElseIF ListType=3 Then
			   Param="Order By LoginTimes Desc"
			End IF
			If KS.S("UserName")<>"" Then
			   SqlStr="Select * From KS_User where GroupID<>4 and UserName like '%" & ks.s("UserName") & "%' " & Param
			Else
			  SqlStr="Select * From KS_User where GroupID<>4 " & Param
			End If
			RS.Open SqlStr,Conn,1,1
			If Not RS.EOF  Then
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
			      GetUserList= GetUserList & showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
			   Else
			      If (CurrentPage - 1) * MaxPerPage < totalPut Then
				     RS.Move (CurrentPage - 1) * MaxPerPage
					 GetUserList= GetUserList &showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
				  Else
				     CurrentPage = 1
					 GetUserList= GetUserList &showContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
			      End If
			   End If
			End If
			GetUserList= GetUserList & "【快速查找】<br/>用户名:<input type=""text"" name=""UserName"" size=""20"" maxlength=""30"" />" & vbcrlf
			GetUserList= GetUserList & "<anchor>搜索<go href=""UserList.asp?"&KS.WapValue&""" method=""post"">"
			GetUserList= GetUserList & "<postfield name=""UserName"" value=""$(UserName)""/>"
			GetUserList= GetUserList & "</go></anchor><br/>"

			RS.Close:Set RS=Nothing
		End Function
		  
		Function ShowContent(RS,totalPut, MaxPerPage, CurrentPage,ListType)
		    Dim I,Privacy,Sex
			Do While Not RS.EOF
			   Privacy=RS("Privacy")
			   ShowContent = ShowContent & "<a href=""User_Friend.asp?Action=saveF&amp;ToUser="&RS("UserName")&"&amp;" & KS.WapValue & """>加友</a> "
			   If RS("Sex")="" or Isnull(RS("Sex")) Then Sex="—" Else Sex=RS("Sex")
			   If Privacy=2 Then Sex="—"  Else Sex=Sex
			   If RS("RealName")="" or Isnull(RS("RealName")) Then
			      ShowContent = ShowContent & "<a href=""ShowUser.asp?UserID="&RS("UserID")&"&amp;" & KS.WapValue & """>"&RS("UserName")&""
			   Else
			      ShowContent = ShowContent & "<a href=""ShowUser.asp?UserID="&RS("UserID")&"&amp;" & KS.WapValue & """>"&RS("RealName")&""
			   End If
			   If Conn.Execute("select UserName from KS_Online where UserName='"&RS("UserName")&"'").EOF Then
			      ShowContent = ShowContent & "("&Sex&"/离线)</a><br/>"
			   Else
			      ShowContent = ShowContent & "("&Sex&"/在线)</a><br/>"
			   End If
			   RS.MoveNext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			ShowContent = ShowContent & KS.ShowPagePara(totalPut, MaxPerPage, "UserList.asp", True, "位", CurrentPage, "ListType="&ListType&"&amp;" & KS.WapValue & "")
			ShowContent = ShowContent & "<br/>"
		End Function

End Class
%> 

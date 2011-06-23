<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
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
<card id="main" title="查询我的点券明细">
<p>
<%
Dim KSCls
Set KSCls = New User_LogPoint
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_LogPoint
        Private KS
		Private CurrentPage,totalPut,TotalPages,SQL
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =5
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Response.Write KS.GetReadMessage
%>

<a href='User_LogPoint.asp?<%=KS.WapValue%>'>所有记录</a> 
<a href='User_LogPoint.asp?InOrOutFlag=1&amp;<%=KS.WapValue%>'>收入[<%=conn.execute("select count(id) from ks_logPoint where InOrOutFlag=1 and username='" & KSUser.UserName & "'")(0)%>]</a> 
<a href='User_LogPoint.asp?InOrOutFlag=2&amp;<%=KS.WapValue%>'>支出[<%=conn.execute("select count(id) from ks_logPoint where InOrOutFlag=2 and username='" & KSUser.UserName & "'")(0)%>]</a>
<br/>
【点券明细】<br/>
<%
If KS.ChkClng(KS.S("InOrOutFlag"))=1 Or KS.ChkClng(KS.S("InOrOutFlag"))=2 Then
   SqlStr="Select ID,UserName,AddDate,IP,Point,InOrOutFlag,Times,User,Descript From KS_LogPoint Where InOrOutFlag=" & KS.ChkClng(KS.S("InOrOutFlag")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
Else
   SqlStr="Select ID,UserName,AddDate,IP,Point,InOrOutFlag,Times,User,Descript From KS_LogPoint Where UserName='" & KSUser.UserName &"' order by id desc"
End if

Set RS=Server.createobject("adodb.recordset")
RS.Open SqlStr,Conn,1,1
If RS.EOF And RS.BOF Then
   Response.Write "找不到您要的记录!<br/>"
Else
   TotalPut=rs.recordcount
   if (TotalPut mod MaxPerPage)=0 then
      TotalPages = TotalPut \ MaxPerPage
   else
      TotalPages = TotalPut \ MaxPerPage + 1
   end if
   if CurrentPage > TotalPages then CurrentPage=TotalPages
   if CurrentPage < 1 then CurrentPage=1
   rs.move (CurrentPage-1)*MaxPerPage
   SQL = rs.GetRows(MaxPerPage)
   Dim i,InPoint,OutPoint
   For i=0 To Ubound(SQL,2)
   %>
   帐号:<%=SQL(1,i)%><br/>
   时间:<%=SQL(2,i)%><br/>

   <%
   If SQL(5,I)=1 Then
      Response.Write "收入:"
	  InPoint=InPoint+SQL(4,I)
   Else
      Response.Write "支出:"
	  OutPoint=OutPoint+SQL(4,I)
   End If  
   Response.Write SQL(4,I)&"点"
   %>,重复:<%=SQL(6,i)%>,操作员:<%=SQL(7,i)%><br/>
   备注:<%=SQL(8,i)%><br/>
   <img src="../Images/Hen.gif" alt=""/><br/>
   <%
   Next
   Call KS.ShowPageParamter(totalPut, MaxPerPage, "User_LogPoint.asp", True, "条记录", CurrentPage, "InOrOutFlag=" & KS.ChkClng(KS.S("InOrOutFlag")) & "&amp;" & KS.WapValue & "")
   %>
   <br/>
   【本页合计】<br/>
   收入点券:<%=InPoint%>点<br/>
   支出点券:<%=KS.ChkClng(OutPoint)%>点<br/>
   <%
   Dim totalinpoint:Totalinpoint=conn.execute("Select sum(Point) From KS_LogPoint where username='" & KSUser.UserName & "'AND InOrOutFlag=1")(0)
   Dim TotalOutPoint:TotalOutPoint=conn.execute("Select sum(Point) From KS_LogPoint where username='" & KSUser.UserName & "'AND  InOrOutFlag=2")(0)
   If KS.ChkClng(totalInPoint)=0 Then totalInPoint=0
   If KS.ChkClng(TotalOutPoint)=0 Then TotalOutPoint=0
   %>
   【所有合计】<br/>
   收入点券:<%=KS.ChkClng(totalInPoint)%>点<br/>
   支出点券:<%=KS.ChkClng(totalOutPoint)%>点<br/>
   合计累计还剩:<%=totalInPoint-totalOutPoint%>点<br/>
<%
End If
Rs.Close:set Rs=Nothing
%>
<br/>
<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a><br/>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>

  <%  

End Sub
  
End Class
%>
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
<card id="main" title="查询我的有效期明细">
<p>
<%
Dim KSCls
Set KSCls = New User_LogEdays
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_LogEdays
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
			Dim InOrOutFlag:InOrOutFlag = KS.ChkClng(KS.S("InOrOutFlag"))
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Response.Write KS.GetReadMessage
			IF InOrOutFlag="" or InOrOutFlag="1" or InOrOutFlag="2" Then
			   Response.Write "<a href='User_LogEdays.asp?"&KS.WapValue&"'>所有记录</a> "
			Else
               Response.Write "所有记录 "
			End If
			IF InOrOutFlag="1" Then
			    Response.Write "收入["&Conn.Execute("select count(id) from ks_logEdays where InOrOutFlag=1 and username='" & KSUser.UserName & "'")(0)&"] "
		    Else
                Response.Write "<a href='User_LogEdays.asp?InOrOutFlag=1&amp;"&KS.WapValue&"'>收入["&Conn.Execute("select count(id) from ks_logEdays where InOrOutFlag=1 and username='" & KSUser.UserName & "'")(0)&"]</a> "
		    End If
			IF InOrOutFlag="2" Then
               Response.Write "支出["&Conn.Execute("select count(id) from ks_logEdays where InOrOutFlag=2 and username='" & KSUser.UserName & "'")(0)&"]"
			Else
			   Response.Write "<a href='User_LogEdays.asp?InOrOutFlag=2&amp;"&KS.WapValue&"'>支出["&Conn.Execute("select count(id) from ks_logEdays where InOrOutFlag=2 and username='" & KSUser.UserName & "'")(0)&"]</a>"
			End If
			Response.Write "<br/>【有效期明细】<br/>"
			
			If KS.ChkClng(InOrOutFlag)=1 Or KS.ChkClng(InOrOutFlag)=2 Then
			   SqlStr="Select ID,UserName,AddDate,IP,Edays,InOrOutFlag,User,Descript From KS_LogEdays Where InOrOutFlag=" & KS.ChkClng(InOrOutFlag) & " And  UserName='" & KSUser.UserName &"' order by id desc"
			Else
			   SqlStr="Select ID,UserName,AddDate,IP,Edays,InOrOutFlag,User,Descript From KS_LogEdays Where UserName='" & KSUser.UserName &"' order by id desc"
			End if
			Set RS=Server.createobject("adodb.recordset")
			RS.open SqlStr,conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "找不到您要的记录!<br/><br/>"
			Else
			   TotalPut=rs.recordcount
			   If (TotalPut Mod MaxPerPage)=0 Then
			      TotalPages = TotalPut \ MaxPerPage
			   Else
			      TotalPages = TotalPut \ MaxPerPage + 1
			   End If
			   If CurrentPage > TotalPages Then CurrentPage=TotalPages
			   If CurrentPage < 1 Then CurrentPage=1
			   RS.move (CurrentPage-1)*MaxPerPage
			   SQL = RS.GetRows(MaxPerPage)
			   Dim i,InEdays,OutEdays
			   For i=0 To Ubound(SQL,2)
			       %>
                   用户名:<%=SQL(1,i)%>(IP:<%=SQL(3,i)%>)<br/>
                   消费时间:<%=SQL(2,i)%><br/>
                   <%
				   If SQL(5,I)=1 Then
				      Response.Write "收入天数:"
					  InEdays=InEdays+SQL(4,I)
				   Else
				      Response.Write "支出天数:"
					  OutEdays=OutEdays+SQL(4,I)
				   End If
				   Response.Write SQL(4,I)&"天,"
				   Response.Write "操作员:"&SQL(6,i)&"<br/>"
                   Response.Write "备注:"&SQL(7,i)&"<br/>"
				   Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
			   Next
			   Call KS.ShowPageParamter(totalPut, MaxPerPage, "User_LogEdays.asp", True, "条记录", CurrentPage, "InOrOutFlag=" & KS.ChkClng(KS.S("InOrOutFlag")) & "&amp;" & KS.WapValue & "")
               Response.Write "【本页合计】<br/>"
               Response.Write "收入天数:" & InEdays & "天<br/>"
			   Response.Write "支出天数:" & KS.ChkClng(OutEdays) & "天<br/>"
			   Dim totalinEdays:totalinEdays=Conn.Execute("Select sum(Edays) From KS_LogEdays where username='" & KSUser.UserName & "'AND InOrOutFlag=1")(0)
			   Dim TotalOutEdays:TotalOutEdays=Conn.Execute("Select sum(Edays) From KS_LogEdays where username='" & KSUser.UserName & "'AND  InOrOutFlag=2")(0)
			   If KS.ChkClng(totalInEdays)=0 Then totalInEdays=0
			   If KS.ChkClng(TotalOutEdays)=0 Then TotalOutEdays=0
               Response.Write "【所有合计】<br/>"
               Response.Write "收入天数:" & KS.ChkClng(totalInEdays) & "天<br/>"
			   Response.Write "支出天数:" & KS.ChkClng(totalOutEdays) & "天<br/>"
               Response.Write "合计累计还剩:" & totalInEdays-totalOutEdays & "天<br/><br/>"
		    End If
			RS.Close:set RS=Nothing
			Response.Write "<a href=""Index.asp?"&KS.WapValue&""">我的地盘</a><br/>"
            Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>"
         End Sub
End Class
%> 

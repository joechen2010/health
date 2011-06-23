<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="投诉管理">
<p>
<%
Dim KSCls
Set KSCls = New JobManage
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class JobManage
        Private KS
		Private TotalPut,CurrentPage,MaxPerPage
		Private Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect DomainStr&"User/Login/"
			   Exit Sub
			End If			
			Select Case KS.S("Action")
			    Case "Show" Call View()
				Case "del" call FeedBackDel()
				Case "Add" call Add()
				Case "DoSave" call AddSave()
				Case Else  Call JobList()
		    End Select
			
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub JobList()
		    %>
            【<a href="User_FeedBack.asp?Action=Add&amp;<%=KS.WapValue%>">我要投诉</a>】<br/>
			<%
			MaxPerPage=10
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Dim Param:Param=" where UserName='" & KSUser.UserName & "'"
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_FeedBack " & Param & " order By ID",Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "您没有发表任意见或投诉!<br/>"
			Else
			   TotalPut = RS.RecordCount
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
			   Call ShowJobList(RS)
			End If
			%>
            <br/>        
            【操作说明】<br/>
            投诉/意见管理放置的是您投诉及对本站的建议记录；<br/>
            您可以删除未处理的记录；<br/>
			<%
		End Sub
		
		Sub ShowJobList(RS)
		    Dim str,i
			Do While Not RS.EOF
			   'Dim bh:bh=RS("ID")
			   'IF Len(BH)=1 Then
			      'BH="00"& BH
			   'ElseIf Len(BH)=2 Then
			      'BH="0" & BH
			   'End If
			   'BH="YJ" & year(RS("Adddate")) & Month(RS("Adddate")) & BH
			   If RS("Accepted")="" or isnull(RS("Accepted")) Then
			      Response.Write "[待]"
			   Else
				  Response.Write "[已]"
			   End If 
			   'Response.Write "" & BH & ":"
			   Response.Write "<a href='User_FeedBack.asp?Action=Show&amp;ID=" & RS("ID") & "&amp;" & KS.WapValue & "'>" & RS("Title") & "</a> "
			   If RS("Accepted")="" or isnull(RS("Accepted")) Then
				  Response.Write "<a href='User_FeedBack.asp?Action=del&amp;ID=" & RS("ID") & "&amp;" & KS.WapValue & "'>删除</a>"
			   End If 
			   Response.Write "<br/>"
			   RS.MoveNext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			Response.Write KS.ShowPagePara(TotalPut, MaxPerPage, "User_FeedBack.asp", True, "位", CurrentPage, "" & KS.WapValue & "")
		End Sub
		
		Sub Add()       
		    Dim ID,RS,RealName,Tel,Sex
			ID=KS.ChkClng(KS.S("ID"))
			%>
            【我要投诉】<br/>
            意见主题:<input type="text" name="Title<%=Minute(Now)%><%=Second(Now)%>" value=""/><br/>
            意见对象:<input type="text" name="Object<%=Minute(Now)%><%=Second(Now)%>" value=""/><br/>
            意见内容:<input type="text" name="Content<%=Minute(Now)%><%=Second(Now)%>" value=""/><br/>
            期望方案:<input type="text" name="Hopesolution<%=Minute(Now)%><%=Second(Now)%>" value=""/><br/>
            <anchor>立即投诉<go href="User_FeedBack.asp?Action=DoSave&amp;<%=KS.WapValue%>" method="post">
            <postfield name="Title" value="$(Title<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Object" value="$(Object<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Content" value="$(Content<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="Hopesolution" value="$(Hopesolution<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="TrainID" value="<%=ID%>"/>
            </go></anchor>
            <br/>
			<%
		End Sub
		
		Sub AddSave()
	        If KS.S("Title")="" Then
			   Response.Write "请输入主题!<br/>"
			   Prev=True
			   Exit Sub
			End If
			If KS.S("Content")="" Then
			   Response.Write "请输入内容!<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim ID:ID=KS.ChkClng(KS.S("ID"))
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "select * from KS_FeedBack where UserName='" & KSUser.UserName & "' And id=" & ID,Conn,1,3
			If RS.EOf Then
			   RS.Addnew
			   RS("Adddate")=Now
			End If
			RS("UserName")=KSUser.UserName
			RS("Title")=KS.S("Title")
			RS("Object")=KS.S("Object")
			RS("Content")=KS.S("Content")
			RS("Hopesolution")=KS.S("Hopesolution")
			RS.Update
			RS.Close
			set RS=Nothing
			Response.Write "你的投诉已提交，请耐心等待处理结果!<br/>"
			Response.Write "<a href=""User_FeedBack.asp?" & KS.WapValue & """>投诉建议</a><br/>"
		End Sub
		
		Sub View()
		    Dim ID,RS
			ID=KS.ChkClng(KS.S("ID"))
			Set RS=Server.CreateOBject("ADODB.RECORDSET")
			RS.Open "Select * from ks_feedback where id=" & ID,Conn,1,1
			IF RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错了!<br/>"
			   Prev=True
			   Exit Sub
			End If
			%>
            【查看详情】<br/>
            
            意见主题:<%=RS("Title")%><br/>
            意见对象:<%=RS("Object")%><br/>
            意见内容:<%=RS("Content")%><br/>
            希望结果:<%=RS("hopesolution")%><br/><br/>
            <%
            If RS("Accepted")="" or isnull(RS("Accepted")) Then
			   Response.Write "待受理中...<br/>"
			Else
			%>
            处 理 人:<%=RS("AcceptEd")%><br/>
            处理时间:<%=RS("AcceptTime")%><br/>
            处理结果:<%=RS("AcceptResult")%><br/>
            <%
            End If 
			RS.Close:Set RS=Nothing
			Prev=True
		End Sub
		
		Sub FeedBackDel()
		    Conn.Execute("Delete from KS_FeedBack where (Accepted='' or Accepted is null ) And UserName='" & KSUser.UserName &"' And ID=" & KS.ChkClng(KS.S("ID")))
			Response.Redirect KS.GetDomain&"User/User_FeedBack.asp?" & KS.WapValue & ""
		End Sub
	
End Class
%> 

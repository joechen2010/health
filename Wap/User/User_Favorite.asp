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
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的收藏夹">
<p>
<%
Dim KSCls
Set KSCls = New User_Favorite
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Favorite
        Private KS
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private ChannelID
		Private TempStr,SqlStr
		Private InfoIDArr,InfoID
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
			Select Case KS.S("Action")
			    Case "Add"
				   Dim RSAdd
				   InfoID=KS.ChkClng(KS.S("InfoID"))
				   Set RSAdd=Server.CreateObject("Adodb.Recordset")
				   ChannelID=KS.ChkClng(KS.S("ChannelID"))
				   RSADD.Open "Select * From KS_Favorite Where ChannelID=" & ChannelID & " And InfoID=" & InfoID & " And UserName='" & KSUser.UserName & "'",Conn,1,3
				   IF RSADD.Eof And RSADD.Bof Then
				      RSADD.AddNew
					  RSAdd(1)=KSUser.UserName
					  RSAdd(2)=ChannelID
					  RSAdd(3)=InfoID
					  RSAdd(4)=Now
					  RSAdd.Update
				   End IF
				   RSADD.Close:SET RSADD=Nothing
			    Case "Cancel"
				   InfoID=KS.S("InfoID")
				   InfoID=Replace(InfoID," ","")
				   InfoID=KS.FilterIDs(InfoID)
				   If InfoID="" Then
				      Response.Write "您没有选择要取消收藏的信息！<br/>"
				      Exit Sub
				   End If
				   Conn.Execute("Delete From KS_Favorite Where ID In(" & InfoID & ") And UserName='" & KSUser.UserName & "'")
			End Select
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
			If ChannelID="" or not isnumeric(ChannelID) Then ChannelID=0
			IF ChannelID<>0 Then  Param= Param & " and ChannelID=" & ChannelID
			%>
            <a href="User_Favorite.asp?<%=KS.WapValue%>">我收藏的信息(<%=Conn.Execute("Select Count(id) from KS_Favorite" & Param & " and channelid<>6")(0)%>)</a><br/>
						   
			<%
			Set RS=Server.CreateObject("AdodB.Recordset")
			SqlStr="Select ID,ChannelID,InfoID,AddDate From KS_Favorite "& Param &" and  Channelid<>6 order by id desc"
			RS.Open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "您的收藏夹没有内容!<br/>"
			Else
			   TotalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and  (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Call ShowContent
			End If
            Response.Write "<br/>"
		    Response.Write "<a href=""Index.asp?"&KS.WapValue&""">我的地盘</a><br/>"
		    Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>"
		End Sub
		
		Sub ShowContent()
		    Dim I,SQL,K
			SQL=RS.GetRows(-1)
			For K=0 To Ubound(SQL,2)
			    Select Case KS.C_S(SQL(1,K),6)
				    Case 1 SqlStr="Select ID,Title,Fname,Changes,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
					Case 2 SqlStr="Select ID,Title,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
					Case 3 SqlStr="Select ID,Title,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
					Case 4 SqlStr="Select ID,Title,Fname,0,AddDate,hits From " & KS.C_S(SQL(1,K),2) &" Where ID=" & SQL(2,K)
					Case 5 SqlStr="Select ID,Title,Fname,0,AddDate,hits From KS_Product Where ID=" & SQL(2,K)
					Case else SqlStr="Select ID From KS_Article Where 1=0"
				End Select
				Dim RSF:Set RSF=Conn.Execute(SqlStr)
				If Not RSF.Eof Then
				   Response.Write "<a href=""" & KS.GetInfoUrl(SQL(1,K),RSF(0),RSF(2)) & """>" & ((I+1)+CurrentPage*MaxPerPage)-MaxPerPage & "." & RSF(1) & "</a><br/>"
				   Response.Write "收藏时间:" & KSUser.GetTimeFormat(SQL(3,K)) & " 人气:" & RSF(5)
				End If
				%>
                <a href="User_Favorite.asp?Action=Cancel&amp;Page=<%=CurrentPage%>&amp;ID=<%=SQL(0,K)%>&amp;<%=KS.WapValue%>">取消收藏</a>
                <br/>
				<%
			Next
			IF TotalPut>MaxPerPage Then
			   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Favorite.asp", True, "篇" & TempStr, CurrentPage, "" & KS.WapValue & "")
		    End IF
		End Sub

End Class
%> 

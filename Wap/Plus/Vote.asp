<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：调查结果
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="调查结果">
<p>
<%
Dim KSCls
Set KSCls = New Vote
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class Vote
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Sub Kesion()
		    Dim Action,ID,VoteType,VoteOption,sqlVote,rsVote
			Action    = KS.S("Action")
			ID        = KS.ChkClng(KS.S("ID"))
			VoteType  = KS.S("VoteType")
			VoteOption= KS.S("VoteOption")
			If Action = "Vote" And Id<> "" And VoteOption<>"" And Session("Voted" &ID) = "" Then
			   If VoteType="Single" Then
			      Conn.Execute "Update KS_Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & ID
			   Else
			      Dim ArrOptions
				  If Instr(VoteOption,",")>0 Then
				     ArrOptions=split(VoteOption,",")
					 Dim i
					 For i=0 To ubound(arrOptions)
					     Conn.Execute "Update KS_Vote set answer" & Cint(Trim(ArrOptions(i)))  & "= Answer" & Cint(Trim(ArrOptions(i))) & "+1 where ID=" & Clng(ID)
					 Next
				  Else
				     Conn.Execute "Update KS_Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & Clng(ID)
				  End If
			   End If
			   Session("Voted" & ID)="True"
			End If
			If ID<>"" Then
			sqlVote="Select * From KS_Vote Where ID=" & Clng(ID)
			Else
			sqlVote="select top 1 * From KS_Vote order by NewestTF desc"
			End If
			Set rsVote = Server.CreateObject("ADODB.Recordset")
			rsVote.Open sqlVote,Conn,1,1
			If not rsVote.EOF Then
			   If Session("voted" & ID)="" Then
			      If Cbool(KSUser.UserLoginChecked)=True Then Response.Write KSUser.RealName & "，"
				  Response.Write "您还没有投票，请您在此投下您宝贵的一票！<br/>"
				  Response.Write "" & rsVote("Title") & "<br/>"
				  Response.Write "<select name='VoteOption'>"
				  If rsVote("VoteType")="Single" Then
				     For i=1 to 8
					     If trim(rsVote("Select" & i) & "")="" Then Exit For
						 Response.Write "<option value='" & i & "'>" & rsVote("Select" & i) & "</option>"
					 Next
				  Else
				  For i=1 To 8
				      If Trim(rsVote("Select" & i) & "")="" Then Exit For
					  Response.Write "<option value='" & i & "'>" & rsVote("Select" & i) & "</option>"
			      Next
			   End If
			   Response.Write "</select> "
			   Response.Write "<anchor>投一票<go href=""Vote.asp?Action=Vote&amp;" & KS.WapValue & """ method=""post"">"
			   Response.Write "<postfield name=""VoteOption"" value=""$(VoteOption)""/>"
			   Response.Write "<postfield name=""VoteType"" value=""" & rsVote("VoteType") & """/>"
			   Response.Write "<postfield name=""ID"" value=""" & rsVote("ID") & """/>"
			   Response.Write "</go></anchor><br/>"
			   Response.Write "<br/>"
			End If
			
			If Action="Vote" And session("voted" & ID)<>"" Then
			   If Cbool(KSUser.UserLoginChecked)=True Then Response.Write KSUser.RealName & "，"
			   Response.Write "非常感谢您的投票！<br/>"
			End If
			Response.Write "调查内容:<b>"&rsVote("Title")&"</b><br/>"
			Response.Write "总投票数:"
			Dim totalVote
			TotalVote=0
			For i=1 To 8
				If rsVote("Select" & i)="" Then Exit For
				TotalVote=TotalVote+rsVote("Answer"& i)
			Next
			Response.Write(totalVote & "票")
			If TotalVote=0 Then TotalVote=1
			Response.Write "<br/><br/>"
			
			For i=1 To 8
			    If Trim(rsVote("Select" & i) & "")="" Then Exit For
				   Response.Write ""&rsVote("Select"& i)&"<br/>"
				   Response.Write "票数:"&rsVote("answer"& i)&""
				   Dim perVote
				   perVote=round(rsVote("answer"& i)/totalVote,4)
				   'Response.Write "<img src='../images/Default/bar.gif' width='" & round(260*perVote) & "' height='15'/>"
				   perVote=perVote*100
				   If perVote<1 And perVote<>0 Then
				      Response.Write " 百分比:0" & perVote & "%"
				   Else
				      Response.Write " 百分比:" & perVote & "%"
				   End If
				   Response.Write "<br/>"
			   Next
            End If
			rsVote.Close():Set rsVote = Nothing
			Response.Write "<br/>"
			Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>"
		End Sub
End Class
%>
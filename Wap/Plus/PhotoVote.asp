<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%Response.ContentType="text/vnd.wap.wml; charset=utf-8" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="投票结果">
<p> 
<%
Dim KS:Set KS=New PublicCls
Dim ID:ID = Replace(KS.S("ID")," ","")
Dim ChannelID:ChannelID=KS.G("ChannelID")
If ChannelID="" Then ChannelID=2
If KS.G("LocalFileName")<>"" And KS.G("RemoteFileUrl")<>"" Then
   If KS.SaveBeyondFile(KS.G("LocalFileName"),KS.G("RemoteFileUrl"))= True Then
      Response.write KS.G("LocalFileName")'错误提示
   End If
End If
Dim LoginTF,ComeUrl,ClassID,UserName
ID=KS.FilterIDs(ID)
If ID="" Then
   Response.Write "对不起，您没有选择投票项!<br/>"
Else
   Const UserTF=0         '是否只允许会员投票 1是 0否
   Const UserIPTF=0       '是否一个IP只允许投一票 1是 0否
   Const UserGroup="0"    '允许投票的会员组，多个会员组请用,号隔开，不想限制请输入0
   LoginTF=KSUser.UserLoginChecked()
   ComeUrl=Request.ServerVariables("HTTP_REFERER")
   ClassID=Conn.Execute("Select top 1 Tid From " & KS.C_S(ChannelID,2) & " where ID In(" & ID & ")")(0)
   
   Call Vote()
   Call ShowVote()
End If

Sub Vote()
	If UserTF=1 and LoginTF=False Then
	   Response.Write "对不起，只会登录会员才能投票!<br/>"
	ElseIf UserGroup<>"0" and KS.FoundInArr(UserGroup, KSUser.GroupID, ",")=False Then
	   Response.Write "对不起，您所在的会员组不允许投票!<br/>"
	ElseIf UserIPTF=1 And not Conn.Execute("Select ID From KS_PhotoVote Where ChannelID=" & ChannelID & " And ClassID='" & ClassID & "'").eof  Then
	   Response.Write "对不起，您已投过票，不能再投！<br/>"
	Else
	   If LoginTF=False Then UserName="游客" Else UserName=KSUser.UserName
	   Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP]) Values(" & ChannelID & ",'" & ClassID & "','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "')")
	   Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Score=Score+1 Where ID In(" & ID & ")")
	   Response.Write "恭喜，您已成功的投票！<br/>"
	End If
	Response.write KS.GetReadReturn&"<br/>"
	Response.Write "<br/>"
End Sub

Sub ShowVote()
	Response.Write "【投票TOP10】<br/>"
	Dim TotalVote:TotalVote=Conn.Execute("Select sum(score) from " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "'")(0)
	If totalvote=0 Then totalvote=1
	Dim RS:Set RS=Conn.Execute("Select top 10 ID,Title,Score,PhotoUrl From " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "' Order BY Score Desc")
	I=1
	Do While Not RS.Eof
	   'If I=1 Then Response.Write "<img src=""jpegmore.asp?jpegsize=90x90&amp;jpegurl="&RS("PhotoUrl")&""" alt=""logo...""/><br/>"
	   
	   Response.Write "<img src=""Images/NumBer_2_0"&I&".gif"" alt="""" /><a href=""Show.asp?ID=" & RS(0) & "&amp;ChannelID=" & ChannelID & "&amp;" & KS.WapValue & """>" & RS(1) & ""
	   Dim perVote:perVote=round(RS(2)/totalVote,4)
	   'Response.Write "<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align=""""/>"
	   perVote=perVote*100
	   If perVote<1 And perVote<>0 Then
	      Response.Write "[0"&perVote&"%]</a><br/>"
	   Else
	      Response.Write "["&perVote&"%]</a><br/>"
	   End If
	   If I=1 Then Response.Write "<br/>"
	   I=I+1
	   RS.MoveNext 
	Loop
End Sub


Response.write "---------<br/>"
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a>"

Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p> 
</card>
</wml>

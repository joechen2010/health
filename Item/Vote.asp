<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KS:Set KS=New PublicCls
Dim KSUser: Set KSUser = New UserCls
Dim ID:ID = Replace(KS.S("ID")," ","")
Dim ChannelID:ChannelID=KS.ChkClng(Request("m"))
If ChannelID="" Then Response.End()
Dim LoginTF,ComeUrl,ClassID,UserName
ID=KS.FilterIDs(ID)
If ID="" Then Response.Write("<script>alert('�Բ�����û��ѡ��ͶƱ��!');history.back();</script>"):Response.End()

Const UserTF=1         '�Ƿ�ֻ�����ԱͶƱ 1�� 0��
Const UserIPTF=1       '�Ƿ�һ��IPֻ����ͶһƱ 1�� 0��
Const UserGroup="0"    '����ͶƱ�Ļ�Ա�飬�����Ա������,�Ÿ�������������������0

'IF Cbool(Request.Cookies(Cstr(ID))("PhotoVote"))<>true Then
' Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set Score=Score+1 Where ID=" & ID)
' Response.Cookies(Cstr(ID))("PhotoVote")=true
' Response.Write "<script>alert('��л����ͶƱ��');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';<//script>"
'Else
'Response.Write "<script>alert('���Ѿ�Ͷ��Ʊ��������Ͷ�ˣ�');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';''<//script>"
'End IF

LoginTF=KSUser.UserLoginChecked()
ComeUrl=Request.ServerVariables("HTTP_REFERER")
ClassID=Conn.Execute("Select top 1 Tid From " & KS.C_S(ChannelID,2) & " where ID In(" & ID & ")")(0)

If KS.S("Action")="Show" Then 
 Call ShowVote()
Else
 Call Vote()
End If

Sub Vote()
	If UserTF=1 and LoginTF=False Then
	   Response.Write "<script>alert('�Բ���ֻ���¼��Ա����ͶƱ!');history.back(-1);</script>"
	   Response.End()
	End If
	
	if UserGroup<>"0" and KS.FoundInArr(UserGroup, KSUser.GroupID, ",")=False Then
	   Response.Write "<script>alert('�Բ��������ڵĻ�Ա�鲻����ͶƱ!');history.back(-1);</script>"
	   Response.End()
	End If
	
	If UserIPTF=1 and not Conn.Execute("Select ID From KS_PhotoVote Where UserIp='" & KS.GetIP & "' and ChannelID=" & ChannelID & " And ClassID='" & ClassID & "'").eof  Then
	   Response.Write "<script>alert('�Բ�������Ͷ��Ʊ��������Ͷ��');history.back();</script>"
	   Response.End()
	End If
	
	If LoginTF=False Then UserName="�ο�" Else UserName=KSUser.UserName
	Conn.Execute("Insert Into [KS_PhotoVote]([ChannelID],[ClassID],[InfoID],[VoteTime],[UserName],[UserIP]) Values(" & ChannelID & ",'" & ClassID & "','" & ID & "'," & SqlNowString & ",'" & UserName & "','" & KS.GetIP() & "')")
	Conn.Execute("Update " & KS.C_S(ChannelID,2) &" Set Score=Score+1 Where ID In(" & ID & ")")
	
	Response.Write "<script>alert('��ϲ�����ѳɹ���ͶƱ��');history.back();</script>"
End Sub

Sub ShowVote()
   Dim TempStr
    TempStr = TempStr & "<table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
    TempStr = TempStr & "     <tr> "
	TempStr = TempStr & "			<td width=""200"" align=""center""><strong>ͶƱѡ��</strong></td>"
	TempStr = TempStr & "			<td width=""100"" align=""center""><strong>��Ʊ��״ͼ</strong></td>"
	TempStr = TempStr & "	    	<td  align=""center""><strong>�ٷֱ�</strong></td>"
	TempStr = TempStr & "	 </tr>"
		
			Dim TotalVote:TotalVote=Conn.Execute("Select sum(score) from " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "'")(0)
			if totalvote=0 then totalvote=1
			Dim RS:Set RS=Conn.Execute("Select Title,Score From " & KS.C_S(ChannelID,2) & " where tid='" & ClassID & "' Order BY Score Desc")
			Do While Not RS.Eof
			
	TempStr = TempStr & "	  <tr> "
	TempStr = TempStr & "		<td height=""25"" style=""BORDER-BOTTOM: 1px solid"" align=""center"">" & rs(0) & "</td>"
	TempStr = TempStr & "		<td  style=""BORDER-BOTTOM: 1px solid"" align=""center"">" & rs(1) & "</td>"
	TempStr = TempStr & "		<td style=""BORDER-BOTTOM: 1px solid""> "
			
			dim perVote:perVote=round(rs(1)/totalVote,4)
	TempStr = TempStr & "<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align='absmiddle'>"
			perVote=perVote*100
			if perVote<1 and perVote<>0 then
				TempStr = TempStr & "&nbsp;0" & perVote & "%"
			else
				TempStr = TempStr & "&nbsp;" & perVote & "%"
			end if
	
	TempStr = TempStr & "</td>"
	TempStr = TempStr & "</tr>"
			RS.MoveNext 
		Loop
		
	TempStr = TempStr & "</table>"
	Set KSR = New Refresh
	Dim Template
	Template=KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "vote.html")  'ģ���ַ
	Template=Replace(Template,"{$ShowVoteResult}",TempStr)
	Response.Write Template
	Set KSR=Nothing
End Sub


Call CloseConn()
Set KS=Nothing
Set KSUser=Nothing
%> 

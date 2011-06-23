<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Vote
KSCls.Kesion()
Set KSCls = Nothing

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
			dim Action,ID,VoteType,VoteOption,sqlVote,rsVote
			Action    = KS.S("Action")
			ID        = KS.ChkCLng(KS.S("ID"))
			VoteType  = KS.S("VoteType")
			VoteOption= KS.R(KS.S("VoteOption"))
			If Action = "Vote" And Id<> "" And VoteOption<>"" And Session("Voted" &ID) = "" Then
				if VoteType="Single" then
				    on error resume next
					conn.execute "Update KS_Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & ID
				else
					dim arrOptions
					if instr(VoteOption,",")>0 then
						arrOptions=split(VoteOption,",")
						dim i
						for i=0 to ubound(arrOptions)
							conn.execute "Update KS_Vote set answer" & cint(trim(arrOptions(i)))  & "= answer" & cint(trim(arrOptions(i))) & "+1 where ID=" & Clng(ID)
						next
					else
						conn.execute "Update KS_Vote set answer" & VoteOption  & "= answer" & VoteOption & "+1 where ID=" & Clng(ID)
					end if
				end if
				session("Voted" & ID)="True"
			End If
			if ID<>"" then
				sqlVote="Select * From KS_Vote Where ID=" & Clng(ID)
			else
				sqlVote="select top 1 * From KS_Vote order by NewestTF desc"
			end if
			Set rsVote = Server.CreateObject("ADODB.Recordset")
			rsVote.open sqlVote,conn,1,1
			if not rsvote.eof then
			%>
			<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
			<HTML>
			<HEAD>
			<TITLE>调查结果</TITLE>
			<META http-equiv=Content-Type content="text/html; charset=gb2312">
			<style type="text/css">
<!--
.style2 {	FONT-SIZE: 11pt; COLOR: #cc0000
}
TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 1.5
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 1.5
}
A:link {
	FONT-SIZE: 9pt; COLOR: #000000; TEXT-DECORATION: underline
}
A:visited {
	FONT-SIZE: 9pt; COLOR: #000000; TEXT-DECORATION: underline
}
A:hover {
	FONT-SIZE: 9pt; COLOR: red
}
.m1 {
	BORDER-TOP: #dfdfdb 1px solid; BORDER-LEFT: #dfdfdb 1px solid; BORDER-BOTTOM: #808080 1px solid
}
.m2 {
	BORDER-RIGHT: #dfdfdb 1px solid; BORDER-TOP: #dfdfdb 1px solid; BORDER-LEFT: #dfdfdb 1px solid; BORDER-BOTTOM: #808080 1px solid
}
.m3 {
	BORDER-RIGHT: #dfdfdb 1px solid; BORDER-TOP: #dfdfdb 1px solid; BORDER-LEFT: #dfdfdb 1px solid
}
.article {
	FONT-SIZE: 10pt; WORD-BREAK: break-all
}
.bn {
	FONT-SIZE: 0.1pt; COLOR: #ffffff; LINE-HEIGHT: 50%
}
.contents {
	FONT-SIZE: 1pt; COLOR: #f7f6f8
}
.nb {
	BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 18px
}
.coolbg {
	BORDER-RIGHT: #acacac 2px solid; BORDER-BOTTOM: #acacac 2px solid; BACKGROUND-COLOR: #e6e6e6
}

-->
            </style>
			</HEAD>
			<BODY leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
            <table width="600" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
				<tr> 
					<td valign="top">
			<%
			if Action="Vote" and session("voted" & ID)<>"" then
				Response.Write "<font color='#FF0000' size='4'>"
				if KS.C("UserName")<>"" then Response.Write KS.C("UserName") & "，"
				Response.Write "&nbsp;&nbsp;非常感谢您的投票！</font><br>"
			end if
			%>
						<table width="600" border="0" align="center" cellpadding="2" cellspacing="0" class="border">
							<tr class="title"> 
								<td height="35" colspan="3"><strong><IMG height=46 src="../images/vote.gif" width=320></strong></td>
							</tr>
							<tr class="tdbg">
								<td>
									<table width="600" border="0" align="center" cellpadding="0" cellspacing="2">
										<tr>
										  <td colspan="3" align="right" bgcolor="#cccc99" height="6"></td>
									  </tr>
										<tr> 
										  <td width="140" align="right" bgcolor="#dfeae4"><div align="center"><strong>调查内容：</strong></div></td>
											<td colspan="2" bgcolor="#dfeae4"><%=rsVote("Title")%></td>
										</tr>
										<tr> 
										  <td width="140" align="right" bgcolor="#f0f2ea"><div align="center"><strong>总投票数：</strong></div></td>
											<td colspan="2" bgcolor="#f0f2ea"> 
			<%
			  dim totalVote
			  totalVote=0
			  for i=1 to 8
				if rsVote("Select" & i)="" then exit for
				totalVote=totalVote+rsVote("answer"& i)
			  next
			  Response.Write(totalVote & "票")
			  if totalVote=0 then totalVote=1
			%>			</td>
										</tr>
										<tr> 
											<td colspan="3" align="center">&nbsp;</td>
										</tr>
										<tr> 
											<td width="140" align="center"><strong>投票选项</strong></td>
											<td width="64" align="right"><div align="center"><strong>票数</strong></div></td>
											<td width="388" align="center"><strong>百分比</strong></td>
										</tr>
			<%
			  for i=1 to 8
				if trim(rsVote("Select" & i) & "")="" then exit for
			%>
										<tr> 
											<td height="25" style="BORDER-BOTTOM: 1px solid" width="140" align="right"><div align="center"><font color="#ff6600"><%=rsVote("Select"& i)%></font> </div></td>
											<td  style="BORDER-BOTTOM: 1px solid" align="right"> 
			<div align="center"><%
			Response.Write rsVote("answer"& i)
			%>
			</div></td>
											<td  style="BORDER-BOTTOM: 1px solid"> 
			<%
			dim perVote
			perVote=round(rsVote("answer"& i)/totalVote,4)
			Response.Write "<img src='../images/Default/bar.gif' width='" & round(360*perVote) & "' height='15' align='absmiddle'>"
			perVote=perVote*100
			if perVote<1 and perVote<>0 then
				Response.Write "&nbsp;0" & perVote & "%"
			else
				Response.Write "&nbsp;" & perVote & "%"
			end if
			%>											</td>
										</tr>
			<% next %>
									</table>								</td>
							</tr>
						</table>
			<%
			if session("voted" & ID)="" then 
					if KS.C("UserName")<>"" then
						Response.Write "<br>&nbsp;&nbsp;&nbsp;&nbsp;" & KS.C("UserName") & "，"
					end if 
					Response.Write "&nbsp;&nbsp;您还没有投票，请您在此投下您宝贵的一票！"
					Response.Write "<form name='VoteForm' method='post' action='vote.asp'>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;" & rsVote("Title") & "<br>"
					if rsVote("VoteType")="Single" then
						for i=1 to 8
							if trim(rsVote("Select" & i) & "")="" then exit for
							Response.Write "<input type='radio' name='VoteOption' value='" & i & "'>" & rsVote("Select" & i) & "<br>"
						next
					else
						for i=1 to 8
							if trim(rsVote("Select" & i) & "")="" then exit for
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='VoteOption' value='" & i & "'>&nbsp;" & rsVote("Select" & i) & "<br>"
						next
					end if
					Response.Write "<br><input name='VoteType' type='hidden'value='" & rsVote("VoteType") & "'>"
					Response.Write "<input name='Action' type='hidden' value='Vote'>"
					Response.Write "<input name='ID' type='hidden' value='" & rsVote("ID") & "'>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:VoteForm.submit();'><img src='../images/Default/voteSubmit.gif' width='52' height='18' border='0'></a>&nbsp;&nbsp;"
					Response.Write "<a href='Vote.asp?ID=" & rsVote("ID") & "&Action=Show' target='_blank'><img src='../images/Default/voteSubmit.gif' width='52' height='18' border='0'></a>"
					Response.Write "</form>"
			end if
			%>
				  </td>
				</tr>
				<tr>
				  <td  bgcolor="#cccc99" height="6"></td>
			  </tr>
			  <tr>
				  <td><p align="center">【<a href="javascript:window.close();">关闭窗口</a>】<br>
			<br></p></td>
			  </tr>
			  
			</table>
			</BODY></HTML>
			<%
			end if
			rsVote.Close():Set rsVote = Nothing
    End Sub
End Class
			%>
 

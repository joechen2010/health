<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Space
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Space
        Private KS,Param
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS20007") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='?';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>ȫ���ƹ��¼</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?flag=1';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ҳ�ƹ��¼</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?flag=2';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>��Աע���ƹ��¼</span></li>"
			  .Write "</ul>"
		End With
		
		
		maxperpage = 20 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("�����ϵͳ����!����������")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		   Param= Param & " and username like '%" & KS.G("KeyWord") & "%'"
		End If
		If KS.G("Flag")="1" Then
		   Param= Param & " and AllianceUser='-'"
		ElseIf KS.G("Flag")="2" Then
		   Param=Param & " and AllianceUser<>'-'"
		End If
		totalPut = Conn.Execute("Select Count(ID) from KS_PromotedPlan" & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Del"
		  Call BlogDel()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td nowrap>������</td>
	<td nowrap>����IP</td>
	<td nowrap>����ʱ��</td>
	<td nowrap>����ҳ��</td>
	<td nowrap>���û���</td>
	<td nowrap>���Ƽ��û�</td>
</tr>
<%
	sFileName = "KS.PromotedPlan.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_PromotedPlan "& Param & " order by id desc"
	If DataBaseType = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û���û����ƹ��¼��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td>&nbsp;<a href="../space/?<%=rs("username")%>" target="_blank"><%=Rs("username")%></a></td>
	<td align="center"><%=Rs("userip")%></td>
	<td align="center"><%=Rs("adddate")%></td>
	<td align="center"><%=Rs("ComeUrl")%></td>
	<td align="center"><%=Rs("Score")%> ��</td>
	<td align="center"><%=Rs("AllianceUser")%></td>
</tr>
<tr><td colspan=9 background='images/line.gif'></td></tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;ɾ����Χ��<input name="deltype" type="radio" value=1>
10��ǰ 
    <input name="deltype" type="radio" value="2" />
    1����ǰ
    <input name="deltype" type="radio" value="3" />
    2����ǰ
    <input name="deltype" type="radio" value="4" />
    3����ǰ
    <input name="deltype" type="radio" value="5" />
    6����ǰ
    <input name="deltype" type="radio" value="6" checked="checked" />
    1��ǰ
	<input class=Button type="submit" name="Submit2" value=" ִ��ɾ�� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ����?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.PromotedPlan.asp", True, "��", CurrentPage, "flag="& KS.G("Flag") & "&KeyWord=" & KS.G("KeyWord") & "&Action=" & Action)
	%></td>
</tr>
</table>
<div>
<form action="KS.PromotedPlan.asp" name="myform" method="post">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>��������=></strong>
	 &nbsp;�û���:<input type="text" class='textbox' name="keyword">&nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
	  </div>
</form>
</div>
<%
End Sub

'ɾ����־
Sub BlogDel()
    Dim Param
		  Select Case KS.ChkClng(KS.G("DelType"))
		   Case 1
		    if DataBaseType=1 then
		    Param="datediff(d,adddate," & SqlNowString & ")>11"
			else
			Param="datediff('d',adddate," & SqlNowString & ")>11"
			end if
		   Case 2
		    if DataBaseType=1 then
		    Param="datediff(d,adddate," & SqlNowString & ")>31"
			else
			Param="datediff('d',adddate," & SqlNowString & ")>31"
			end if
		   Case 3
		    if DataBaseType=1 then
		     Param="datediff(d,adddate," & SqlNowString & ")>61"
			else
			 Param="datediff('d',adddate," & SqlNowString & ")>61"
			end if
		   Case 4
		    if DataBaseType=1 then
		    Param="datediff(d,adddate," & SqlNowString & ")>91"
			else
			Param="datediff('d',adddate," & SqlNowString & ")>91"
			end if
		   Case 5
		    if DataBaseType=1 then
		    Param="datediff(d,adddate," & SqlNowString & ")>181"
			else
			Param="datediff('d',adddate," & SqlNowString & ")>181"
			end if
		   Case 6
		    if DataBaseType=1 then
		    Param="datediff(d,adddate," & SqlNowString & ")>366"
			else
			Param="datediff('d',adddate," & SqlNowString & ")>366"
			end if
		  End Select
		  If Param<>"" Then Conn.Execute("Delete From KS_PromotedPlan Where 1=1 and " & Param)
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
	End Sub

End Class
%> 

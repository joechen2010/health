<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%dim channelid
Dim KSCls
Set KSCls = New Admin_Ask_Class
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Class
        Private KS,DataArry
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
        If Not KS.ReturnPowerResult(0, "WDXT10003") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 KS.Die ""
		 End If
		Dim Action,DataArry
		Action = LCase(Request("action"))
		Select Case Trim(Action)
		Case "save"
			Call saveScore()
		Case Else
			Call showmain()
		End Select
		End Sub
		Sub showmain()
			Dim i,iCount,lCount
			iCount=2:lCount=1
		%>
		<html>
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		</head>
		<body>
		<div class='topdashed sort'>�ʴ�/С��̳�ȼ�ͷ������</div>
		<table id="tablehovered" border="0" align="center" cellpadding="3" cellspacing="1" width="100%">
		<form name="selform" id="selform" method="post" action="?">
		<input type="hidden" name="action" value="save">
		<tr class='sort'>
			<td width="10%" noWrap="noWrap">�ȼ�ID</td>
			<td width="40%">�û��ȼ�ͷ��</td>
			<td width="15%" noWrap="noWrap">���������</td>
			<td width="15%" noWrap="noWrap">�������</td>
		</tr>
		<%
			Call showScoreList()
			iCount=1:lCount=2
			If IsArray(DataArry) Then
				For i=0 To Ubound(DataArry,2)
					If Not Response.IsClientConnected Then Response.End
		%>
		<tr align="center">
			<td class="splittd"><input type="hidden" name="GradeID" value="<%=DataArry(0,i)%>"><%=DataArry(0,i)%></td>
			<td class="splittd"><input type="text" size="25" name="UserTitle<%=DataArry(0,i)%>" value="<%=Server.HTMLEncode(DataArry(1,i))%>" /></td>
			<td class="splittd"><input type="text" size="15" name="Score<%=DataArry(0,i)%>" value="<%=Server.HTMLEncode(DataArry(2,i))%>" /></td>
			<td class="splittd">
			<%if DataArry(0,i)<17 then%>
			 <a href="#" disabled>ɾ��</a>
			<%else%>
			 <a href="?x=c&id=<%=DataArry(0,i)%>" onclick="return(confirm('ȷ��ɾ����?'))">ɾ��</a>
			<%end if%>
			</td>
		</tr>
		<%
				Next
			End If
			DataArry=Null
		%>
		<tr align="center">
			<td class="tablerow<%=lCount%>" colspan="5">
				<input class="button" type="submit" name="submit_button" value="������������"/>
			</td>
		</tr>
		</form>

		<form action="?x=b" method="post" name="myform" id="form">
		    <tr>
			<td height="25" colspan="6">&nbsp;&nbsp;<strong>&gt;&gt;�����ȼ�ͷ��</strong><<</td>
		    </tr>
			<tr><td colspan=9 background='images/line.gif'></td></tr>
			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="UserTitle" type="text" class="textbox" id="UserTitle" size="25"></td>
			  <td height="25" align="center"><input style="text-align:center" name="Score" type="text" value="1000" class="textbox" id="Score" size="8">
��</td>
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,�ύ"></td>
			</tr>
			<tr><td colspan=9 background='images/line.gif'></td></tr>
		</form>

		</table>
		<%
		 Select case request("x")
		   case "b"
		       If KS.G("UserTitle")="" Then Response.Write "<script>alert('������ȼ�ͷ��!');history.back();</script>":response.end
			   If Not Isnumeric(KS.G("Score")) Then Response.Write "<script>alert('���ֱ���������!');history.back();</script>":response.end
				conn.execute("Insert into KS_AskGrade(UserTitle,score)values('" & KS.G("UserTitle") & "','" & KS.G("Score") & "')")
				
				KS.AlertHintScript "��ϲ,�ȼ�ͷ�γɹ�!"
		   case "c"
				conn.execute("Delete from KS_AskGrade where GradeID="& KS.ChkClng(KS.G("id")))
				KS.AlertHintScript "��ϲ,�ȼ�ͷ��ɾ���ɹ�!"
		End Select
		  
		End Sub
		
		Sub showScoreList()
			Dim Rs,SQL
			SQL="SELECT GradeID,UserTitle,Score FROM [KS_AskGrade] order by gradeid"
			Set Rs=Conn.Execute(SQL)
			If Not (Rs.BOF And Rs.EOF) Then
				DataArry=Rs.GetRows(-1)
			Else
				DataArry=Null
			End If
			Rs.close()
			Set Rs=Nothing
		End Sub
		
		Sub saveScore()
			Dim Rs,SQL,i
			Dim GradeID,UserTitle,Score
			    GradeID=Split(Replace(Request.Form("GradeID")," ",""),",")
                For I=0 To Ubound(GradeID)
				 UserTitle=Replace(Request.Form("UserTitle"&GradeID(I)),"'","")
				 Score=KS.ChkClng(Request.Form("Score"&GradeID(I)))
				 If GradeID(I)>0 Then
					Conn.Execute ("UPDATE KS_AskGrade SET UserTitle='"&UserTitle&"',Score="&Score&" WHERE GradeID="&GradeID(I))
				 End If
			   Next
			Call KS.AlertHintScript("��ϲ���������û����ֵȼ��ɹ�!")
		End Sub
End Class
%>
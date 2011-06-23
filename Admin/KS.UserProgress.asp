<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Digg
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Digg
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,ChannelID,ItemName,ItemName1,RS
		Private ch_rs,ch_sql,ModelEname,Inputer
		
		Private Sub Class_Initialize()
		  MaxPerPage = 10
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With Response
                Action=KS.G("Action")
				If Not KS.ReturnPowerResult(0, "KMUA10011") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF

			 Page=KS.G("Page")
			If Not IsEmpty(Request("page")) Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			 Select Case Action
			  Case "ShowDetail"
			    Call ShowDetail()
				Exit Sub
			  Case "ShowAdmin"
			    Call ShowAdmin()
			  Case Else
			   Call MainList()
			 End Select
			.Write "</body>"
			.Write "</html>"
			End With
		End Sub
		Sub MainList()
channelid=ks.g("channelid")
if channelid="" then channelid=1
 ModelEname=KS.C_S(ChannelID,2)
 Inputer="Inputer"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script>
<script src="../ks_inc/kesion.box.js"></script>
</head>
<body scroll="no" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script>
 function ShowDetail(param)
 {  onscrolls=false;
	PopupCenterIframe('<b>查看用户添加的详细记录</b>',"KS.UserProgress.asp?Action=ShowDetail&"+param,750,380,'auto');
 }
</script>

<div class='topdashed sort'><a href="?channelid=1">会员投稿统计</a> | <a href="?action=ShowAdmin">管理员工作进度</a></div>
<div style="height:95%; overflow: auto; width:100%" align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="699" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td width=1 bgcolor="#E3E3E3"></td>
          <td width="1011"><div align="center"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>按组查看投稿情况</strong></font></td>
				  <td>按数据表查看
				  <select name="channelid" onChange="location.href='KS.UserProgress.asp?username=<%=request("username")%>&channelid='+this.value">
				  <%
				   dim trs:set trs=conn.execute("select channelid,channelname,channeltable from ks_channel where channelid<>6 and channelid<>10 and channelid<>9 and channelstatus=1")
				   do while not trs.eof
				    if channelid=trim(trs(0)) then
				    response.write "<option value=" & trs(0) &" selected>" & trs(1) &"(" & trs(2) &")</option>"
					else
				    response.write "<option value=" & trs(0) &">" & trs(1) &"(" & trs(2) &")</option>"
					end if
				   trs.movenext
				   loop
				   trs.close:set trs=nothing
				  %>
				  </select>
				  </td>
                </tr>
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE"> 
                  <td width="48%" height="22" align="center" bgcolor="#C8D5E4"><strong>用户组名称</strong></td>
                  <td width="52%" height="22" align="center" bgcolor="#C8D5E4"><strong>投稿数量</strong></td>
                </tr>
				<%
				Call KS.LoadUserGroup()
				Dim Node
				For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
				%>			  
                <tr bgcolor="#EEF8FE"> 
                  <td height="22" align="center"><%=Node.SelectSingleNode("@groupname").text%></td>
                  <td height="22" align="center"><%=LFCls.GetSingleFieldValue("select count(*) as countnumsl from "&ModelEname&" a inner join KS_User b on a." & Inputer & "=b.username where b.GroupID="&Node.SelectSingleNode("@id").text)%></td>
                </tr>
				<%
				Next
				
				
			   %>			  
               
              </table>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="700" height="30" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td>　<font color="#006666"><strong>个人投稿统计</strong></font></td>
                </tr>
              </table>			  
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="1"></td>
                </tr>
              </table>
              <table width="699" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#CDCDCD">
                <tr bgcolor="#EEF8FE">
                  <td height="25" align="center" bgcolor="#C8D5E4"><strong>用户名称</strong></td>
                  <td align="center" bgcolor="#C8D5E4"><strong>所属用户组</strong></td>
                  <td height="25" align="center" bgcolor="#C8D5E4"><strong>投稿数量(<%=KS.C_S(ChannelID,1)%>)</strong></td>
                  <td align="center" bgcolor="#C8D5E4"><strong>操作</strong></td>
                </tr>
				<%
				dim rs_hyz,rs_user,sql_user,sql_hyz,hyz,countnum,rs_wz,sql_wz,param
				if ks.g("username")<>"" then param="where a.username='" & ks.g("username") & "'"
				
				set rs_user=server.CreateObject("adodb.recordset")
				rs_user.open "select b.groupname,a.* from KS_User a inner join ks_usergroup b on a.groupid=b.id " & param & " order by userID desc",conn,1,1
				 If Not rs_user.EOF Then
					  totalPut = rs_user.RecordCount
					  If CurrentPage < 1 Then CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									rs_user.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
						    End If
					Call showuserlist(rs_user)
			End If
				rs_user.close:set rs_user=Nothing

				
   %>         </table>			  
              
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
<form name='myform' action='KS.UserProgress.asp' method='post'>
<input type='hidden' value='<%=channelid%>' name='channelid'>
搜索指定用户的稿件情况:<input type='text' name='username'>&nbsp;<input type='submit' class='button' value='搜索用户'>
</form>
</div>
</body>
</html>		<%
		End Sub
		
		Sub ShowUserList(rs_user)
		     dim i
			do while not rs_user.eof
			   %>
			   <tr bgcolor="#EEF8FE"> 
                  <td width="23%" height="25" align="center"><%=rs_user("username")%></td>
                  <td width="34%" align="center"><%=rs_user(0)%></td>
                  <td width="22%" height="25" align="center"><%=LFCls.GetSingleFieldValue("select count(*) as countnum from "&ModelEname&" where " & Inputer & "='"&rs_user("username")&"'")%></td>
                  <td width="21%" height="25" align="center"><a  href='javascript:ShowDetail("username=<%=rs_user("username")%>&ChannelID=<%=channelid%>&Flag=all");'>查看投稿</a></td>
                </tr>
				<%
				I = I + 1
		        If I >= MaxPerPage Then Exit Do
				rs_user.movenext
				loop
			 Response.Write ("<tr><td colspan=6  bgcolor='#EEF8FE'><div style='text-align:center'>")
	 		Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.UserProgress.asp", True, "位", CurrentPage,"ChannelID="&ChannelID)
	     Response.Write ("</div><br></td></tr>")

		End Sub
		
		
		
		Sub ShowAdmin()
		With Response
		 .Write "<html>"
		 .Write "<head>"
		 .Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
		 .Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrLf 
		 .Write "<script language=""JavaScript"" src=""../ks_inc/Common.js""></script>" & vbCrLf
		 .Write "<script src=""../ks_inc/kesion.box.js"" language=""JavaScript""></script>" & vbCrLf
		 .Write "</head>"
		 .Write "<body topmargin='0' leftmargin='0'>"
		 .Write "<script>"
		 .Write " function ShowDetail(param){ "
		 .Write "PopupCenterIframe('<b>查看管理员添加的详细记录</b>',""KS.UserProgress.asp?Action=ShowDetail&""+param,750,380,'auto');"
		 .Write "}"
	     .Write " </script>"
		
		
		.Write "<div class='topdashed sort'><a href='?channelid=1'>会员投稿统计</a> | <a href='?action=ShowAdmin'>管理员工作进度</a></div>"
		.Write "</ul>"
		.Write "<br><br><table width='70%' align='center' border='0' cellpadding='0' cellspacing='0'>"
		.Write "    <tr class='sort'>"
		.Write "    <td width='100' align='center'>管理员</td>"
		.Write "    <td width='200' align='center'>模块</td>"
		.Write "    <td width='100' align='center'>今日</td>"
		.Write "    <td width='100' align='center'>本周</td>"
		.Write "    <td width='100' align='center'>本月</td>"
		.Write "    <td width='100' align='center'>今年</td>"
		.Write "    <td width='100' align='center'>所有</td>"
		.Write "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
		   Dim Param:Param=" where 1=1 order by AdminID"
		   SqlStr = "SELECT UserName,RealName FROM [KS_Admin] " & Param
			  RS.Open SqlStr, conn, 1, 1
				 If Not RS.EOF Then
					  totalPut = RS.RecordCount
					  If CurrentPage < 1 Then CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
						    End If
					Call showContent
			End If
		  .Write "  </td>"
		  .Write "</tr>"

		 .Write "</table>"
		 .Write ("<div style='text-align:center'>")
	 		Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.UserProgress.asp", True, "位", CurrentPage,KS.QueryParam("Page"))
	     .Write ("</div")
		End With
		End Sub
		Sub showContent()
		  Dim Param,I,K,RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit From KS_Channel Where ChannelStatus=1 and channelid<>6 and channelid<>9 and channelid<>10 Order By ChannelID")
		  Dim SQL:SQL=RSC.GetRows(-1)
		  RSC.Close:Set RSC=Nothing
		
		  With Response
		  Do While Not RS.EOF
		   .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   .Write "<td style='border:1px dashed #cccccc' align='center'><img src='images/admin.png'><br>" & RS(0) & "<br>(" & RS(1) & ")</td>"
		   .Write "<td colspan='6' height='22' style='border:1px dashed #cccccc'>"
		   
		   .Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'>" &vbcrlf
		    For K=0 to Ubound(SQL,2)
			   Param=" Where Inputer='" & RS(0) & "'"
			 .Write "<tr>" &vbcrlf
			 .Write "<td height='22' width='200'>" & SQL(1,k) & "</td>" & vbcrlf
			 
				 .Write "<td width='80' align='center'><a href='javascript:ShowDetail(""username=" & RS(0) &"&ChannelID=" & SQL(0,K)&"&Flag=today"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff('d',AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 .Write "<td width='100' align='center'><a href='javascript:ShowDetail(""username=" & RS(0) &"&ChannelID=" & SQL(0,K)&"&Flag=week"");' title='点击查看详情！'><font color=green>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff('w',AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 .Write "<td width='100' align='center'><a href='javascript:ShowDetail(""username=" & RS(0) &"&ChannelID=" & SQL(0,K)&"&Flag=month"");' title='点击查看详情！'><font color=#ff6600>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff('m',AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
				 .Write "<td width='100' align='center'><a href='javascript:ShowDetail(""username=" & RS(0) &"&ChannelID=" & SQL(0,K)&"&Flag=year"");' title='点击查看详情！'><font color=blue>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param &" And datediff('yyyy',AddDate," & SqlNowString & ")=0")(0) & "</font> " & SQL(4,K) & "</a></td>"
			 .Write "<td width='100' align='center'><a href='javascript:ShowDetail(""username=" & RS(0) &"&ChannelID=" & SQL(0,K)&"&Flag=all"");' title='点击查看详情！'><font color=red>" & Conn.Execute("select count(id) from " & SQL(2,K) & Param)(0) & "</font> " & SQL(4,K) & "</a></td>"
			 .Write "</tr>" & vbcrlf
		    Next
		   .Write "</table>"
			
			
		   .Write  "</td>"
		   .Write "</tr>"
		    I = I + 1
		    If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		   RS.Close
		  End With
		 End Sub
		 

		 
		 Sub ShowDetail()
		    With Response	
			 .Write "<html>"
			 .Write"<head>"
			 .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			 .Write"<link href=""Include/Admin_style.CSS"" rel=""stylesheet"" type=""text/css"">"
			 .Write"<link href=""Include/Admin_box.CSS"" rel=""stylesheet"" type=""text/css"">"
			 .Write "<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			 .Write"</head>"
			 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
            End WIth
			Dim UserName:UserName=KS.G("UserName")
			Dim ChannelID:ChannelID=KS.ChkClng(KS.G("ChannelID"))
			Dim Flag:Flag=KS.G("Flag")
			Dim SQLStr,Param
			 MaxPerPage = 15
			Response.Write "<div style='margin:2px;text-align:center'>"
			Select Case Flag
			 Case "today"
			  Response.Write "查看<font color=red>" & UserName & "</font> 今天添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff('d',AddDate," & SqlNowString & ")=0"
			 Case "week"
			  Response.Write "查看<font color=red>" & UserName & "</font> 本周添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff('w',AddDate," & SqlNowString & ")=0"
			 Case "month"
			  Response.Write "查看<font color=red>" & UserName & "</font> 本月添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff('m',AddDate," & SqlNowString & ")=0"
			 Case "year"
			  Response.Write "查看<font color=red>" & UserName & "</font> 今年添加的" &KS.C_S(ChannelID,3)
			   Param=" And datediff('yyyy',AddDate," & SqlNowString & ")=0"
			 Case "all"
			  Response.Write "查看<font color=red>" & UserName & "</font> 所有添加的" &KS.C_S(ChannelID,3)
			End Select
			If KS.C_S(ChannelID,6)=8 Then
			 SQLStr="Select id,title,username,adddate from " & KS.C_S(ChannelID,2) & " Where username='" & UserName &"'"
			Else
			 SQLStr="Select id,title,Inputer,adddate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName &"'"
			End If
		
			SQLStr=SQLStr & Param & " Order By ID Desc"
			Response.Write "</div>"
			Response.Write "<br><table width='95%' align='center' border='0' cellpadding='0' cellspacing='0'>"
			Response.Write "    <tr class='sort'>"
			Response.Write "    <td width='100' align='center'>ID</td>"
			Response.Write "    <td align='center'>名称</td>"
			Response.Write "    <td  align='center'>录入员</td>"
			Response.Write "    <td  align='center'>录入时间</td>"
			Response.Write "    <td width='100' align='center'>查看详情</td>"
			Response.Write "  </tr>"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
             RS.Open SqlStr, conn, 1, 1
				 If Not RS.EOF Then
					  totalPut = RS.RecordCount
					  If CurrentPage < 1 Then CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage = 1 Then
								Call showDetailContent(ChannelID)
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
								Call showDetailContent(ChannelID)
							End If
			End If
		 Response.Write "  </td>"
		 Response.Write "</tr>"

		 Response.Write "</table>"
		 Response.Write ("<div style='display:block;text-align:center'>")
	 		Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.UserProgress.asp", True, KS.C_S(ChannelID,4), CurrentPage,"Action=ShowDetail&ChannelID="&ChannelID&"&UserName=" & UserName & "&flag=" & Flag)
	     Response.Write ("</div")
		 Response.Write "</table>"
		 End Sub
		 
		 Sub showDetailContent(ChannelID)
		  Dim I:I=0
		  Do While Not RS.Eof
		   Response.Write "<tr class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		   Response.Write "<td class='splittd' style='height:20px;' align='center'>" & RS(0) & "</td>"
		   Response.Write "<td class='splittd'>" & KS.Gottopic(RS(1),20) & "</td>"
		   Response.Write "<td class='splittd' align='center'>" & RS(2) & "</td>"
		   Response.Write "<td class='splittd' align='center'>" & RS(3) & "</td>"
		   Response.Write "<td class='splittd' align='center'><a href='../item/show.asp?d=" & RS(0) &"&m=" & channelid & "' target='_blank'>查看内容</a></td>"
		   Response.Write "</tr>"
		  	I = I + 1
		    If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		   RS.Close
		 End Sub
End Class
%> 

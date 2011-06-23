<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_LogScore
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogScore
        Private KS,KSUser
		Private CurrentPage,totalPut,TotalPages,SQL
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("查询我的积分明细")
		If KS.S("page") <> "" Then
		  CurrentPage = CInt(KS.S("page"))
		Else
		  CurrentPage = 1
		End If
	  %>
		<div class="tabs">	
			<ul>
				<li><a href="user_logmoney.asp">资金明细</a></li>
				<li><a href="user_LogPoint.asp">点券明细</a></li>
				<li><a href="user_logedays.asp">有效期明细</a></li>
				<li class="select"><a href="user_logscore.asp">积分明细</a></li>
			</ul>
		</div>
		
		    <table width="98%" align="center" border="0">
  <tr>
    <td> <a href="?channelid=1000">点广告积分收入明细</a> | <a href="?channelid=1001">点友情链积分收入明细</a></td>
    <td align="right"><a href='User_LogScore.asp'><font color=red>·所有记录</font></a> ·<a href='?InOrOutFlag=1'>收入记录[<%=conn.execute("select count(id) from ks_LogScore where InOrOutFlag=1 and username='" & KSUser.UserName & "'")(0)%>]</a> ·<a href='?InOrOutFlag=2'>支出记录[<%=conn.execute("select count(id) from ks_LogScore where InOrOutFlag=2 and username='" & KSUser.UserName & "'")(0)%>]</a></td>
  </tr>
</table>

			
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class="title">
					<td width="80" height="25" align="center"><strong> 用户名</strong></td>
					<td width="180" height="25" align="center"><strong>产生时间</strong></td>
					<td width="71" height="25" align="center"><strong>收入</strong></td>
					<td width="74" align="center"><strong>支出</strong></td>
					<td width="55" height="25" align="center"><strong>摘要</strong></td>
					<td width="75" height="25" align="center"><strong> 余额</strong></td>
					<td width="239" align="center"><strong>备注</strong></td>
				  </tr>
					<%  
					  dim param
					 If KS.ChkClng(Request("channelid"))<>0 then
					   param=" and channelid=" & KS.ChkClng(Request("channelid"))
					 end if
					 
					If KS.ChkClng(KS.S("InOrOutFlag"))=1 Or KS.ChkClng(KS.S("InOrOutFlag"))=2 Then
						  SqlStr="Select ID,UserName,AddDate,IP,Score,InOrOutFlag,CurrScore,Descript From KS_LogScore Where InOrOutFlag=" & KS.ChkClng(KS.S("InOrOutFlag")) & " And  UserName='" & KSUser.UserName &"'" & param & " order by id desc"
 					    Else
						  SqlStr="Select ID,UserName,AddDate,IP,Score,InOrOutFlag,CurrScore,Descript From KS_LogScore Where UserName='" & KSUser.UserName &"'" & param & " order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>找不到您要的记录!</td></tr>"
								 Else
									TotalPut=rs.recordcount
									if (TotalPut mod MaxPerPage)=0 then
										TotalPages = TotalPut \ MaxPerPage
									else
										TotalPages = TotalPut \ MaxPerPage + 1
									end if
									if CurrentPage > TotalPages then CurrentPage=TotalPages
									if CurrentPage < 1 then CurrentPage=1
									rs.move (CurrentPage-1)*MaxPerPage
									SQL = rs.GetRows(MaxPerPage)
									rs.Close:set rs=Nothing
									ShowContent
				End If

						
						 %>
          </table>
		  </td>
		  </tr>
</table>
		  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
  End Sub
    
  Sub ShowContent
 Dim i,InPoint,OutPoint
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25" class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="80" align="center" class="splittd"><%=SQL(1,i)%></td>
    <td align="center" class="splittd"><%=SQL(2,i)%></td>
    <td align="right" class="splittd"><%if SQL(5,I)=1 Then Response.Write SQL(4,I) & "分":InPoint=InPoint+SQL(4,I) ELSE Response.Write "-"%></td>
    <td align="right" class="splittd"><%if SQL(5,I)=2 Then Response.Write SQL(4,I) & "分":OutPoint=OutPoint+SQL(4,I) ELSE Response.Write "-"%></td>
    <td align="center" class="splittd"><%if SQL(5,I)=1 Then Response.Write "<font color=red>收入</font>" Else Response.Write "支出"%></td>
    <td align="center" class="splittd"><%=SQL(6,i)%>分</td>
	<td class="splittd"><%=SQL(7,i)%></td>
  </tr>
  <%Next%>
  <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">   
   <td colspan='2'  class="splittd" align='right'>本页合计：</td>    <td  class="splittd" align='right'><%=InPoint%>分</td>    <td align='right'><%=KS.ChkClng(OutPoint)%>分</td>    <td  class="splittd" colspan='4'>&nbsp;</td>  </tr> 
  <% Dim totalinpoint:totalinpoint=conn.execute("Select sum(score) From KS_LogScore where username='" & KSUser.UserName & "'AND InOrOutFlag=1")(0)
     Dim TotalOutPoint:TotalOutPoint=conn.execute("Select sum(score) From KS_LogScore where username='" & KSUser.UserName & "'AND  InOrOutFlag=2")(0)
	 If KS.ChkClng(totalInPoint)=0 Then totalInPoint=0
	 If KS.ChkClng(TotalOutPoint)=0 Then TotalOutPoint=0
  %>
    <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">    <td  class="splittd" colspan='2' align='right'>所有合计：</td>    <td  class="splittd" align='right'><%=KS.ChkClng(totalInPoint)%>分</td>    <td  class="splittd" align='right'><%=KS.ChkClng(totalOutPoint)%>分</td>    <td  style="display:none" class="splittd" colspan='4' align='center'>累计还剩：<%=totalInPoint-totalOutPoint%>分</td>  </tr> 

  <%  

End Sub
  
End Class
%> 

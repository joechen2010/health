<%@ Language="VBSCRIPT" codepage="936" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KSCls
Set KSCls = New AjaxCls
KSCls.Kesion()
Set KSCls = Nothing

Class AjaxCls
      Private KS,KSBCls
	  Private Action,Template,id,groupadmin:
	  Private CurrentPage,totalPut,MaxPerPage,PageNum
	  Private Sub Class_Initialize()
	   Set KS=New PublicCls
	   Set KSBCls=New BlogCls
      End Sub
	 Private Sub Class_Terminate()
	  Set KSBCls=Nothing
	  Set KS=Nothing
	  CloseConn()
	 End Sub

     Sub Kesion()
      Action=KS.S("Action")
	   Select Case Action
		Case "space"
		 Call SpaceList()
		Case "log"
		 Call LogList()
		Case "group"
		 Call GroupList()
		Case "photo"
		 Call PhotoList()
	   End Select	
	 End Sub	
	 

  
  '日志列表
  Sub LogList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(10))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim isbest:isbest=ks.chkclng(ks.s("isbest"))
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
%>
  <table border="0" cellpadding="1" cellspacing="1" width="98%" backcolor="#efefef">
      <tr height="22">
      <td colspan=2><strong>分类查看:</strong>
	   <select name="classid" onchange="SpacePage(1,'log&classid='+this.value);">
	    <option value='0'>所有分类</option>
	   <% dim rsc:set rsc=conn.execute("select typename,typeid from ks_blogtype order by orderid")
	   if not rsc.eof then
	    do while not rsc.eof
		 if classid=rsc(1) then
		 response.write "<option value=""" & rsc(1) & """ selected>" & rsc(0) & "</option>"
		 else
		 response.write "<option value=""" & rsc(1) & """>" & rsc(0) & "</option>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
	   %>
	   </select>
	  </td>
      <td align="center" colspan=2><strong>日志标题：</strong><input style="border:1px #000 dashed;height:18px;" type="text" size="12" name="key">&nbsp;&nbsp;<input type="button" onclick="SpacePage(1,'log&key='+document.getElementById('key').value);" value= " 查 找 "></td>
    </tr>
      <tr height="22" bgcolor="#f9f9f9">
      <td><strong>日志标题</strong></td>
      <td width="100" align="center"><strong>分 类</strong></td>
      <td width="70" align="center"><strong>作者</strong></td>
      <td align="center"><strong>更新时间</strong></td>
    </tr>

<%
 dim param:param=" where status=0"
 if classid<>0 then param=param & " and a.typeid=" & classid
 if isbest<>0 then param=param & " and best=1"

 if ks.s("key")<>"" then param=param & " and Title like '%" & ks.r(ks.s("key")) &"%'"
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			rsobj.open "select a.*,b.typename from ks_blogInfo a inner join ks_blogType b on a.typeid=b.typeid " & param & " order by adddate desc" ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=4><p>对不起，没有找到日志文章! </p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call ShowlogList(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call ShowlogList(RSObj)
									Else
										CurrentPage = 1
										call ShowlogList(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|篇||2"
		 RSObj.Close:Set RSObj=Nothing
  End Sub

  Sub ShowLogList(rs)
   dim i
   do while not rs.eof
    if i mod 2=0 then
	 response.write "<tr style=""background:#fff;height:22px"">"
	else
	 response.write "<tr style=""background:#FBFDFF;height:22px"">"
	end if
	%>
      <td><img src="images/bullet.gif" align="absmiddle" />
      <a title="<%=rs("Title")%>" href="<%=KSBCls.GetCurrLogUrl(rs("id"),rs("username"))%>" target="blank">
	  <%=KS.GotTopic(rs("title"),32)%></a>	 
	   <%if rs("best")=1 then response.write "<font color=red>[精]</font>"%>
</td>
      <td align="center"><%=rs("typename")%></td>
      <td align="center"><%=rs("username")%></td>
      <td align="center"><%=rs("adddate")%></td>
    </tr>
  <%

   rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
  End Sub
  
  '空间列表
  Sub SpaceList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(9))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim recommend:recommend=ks.chkclng(ks.s("recommend"))
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
%>
  <table border="0" cellpadding="1" cellspacing="1" width="98%" backcolor="#efefef">
      <tr height="22">
      <td colspan=2><strong>分类查看:</strong>
	   <select name="classid" onchange="SpacePage(1,'space&classid='+this.value);">
	    <option value='0'>所有分类</option>
	   <% dim rsc:set rsc=conn.execute("select classname,classid from ks_blogclass order by orderid")
	   if not rsc.eof then
	    do while not rsc.eof
		 if classid=rsc(1) then
		 response.write "<option value=""" & rsc(1) & """ selected>" & rsc(0) & "</option>"
		 else
		 response.write "<option value=""" & rsc(1) & """>" & rsc(0) & "</option>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
	   %>
	   </select>
	  </td>
      <td align="center" colspan=2><strong>空间名称：</strong><input style="border:1px #000 dashed;height:18px;" type="text" size="12" name="key">&nbsp;&nbsp;<input type="button" onclick="SpacePage(1,'space&key='+document.getElementById('key').value);" value= " 查 找 "></td>
    </tr>
      <tr height="22" bgcolor="#f9f9f9">
      <td><strong>空间名称</strong></td>
      <td width="100" align="center"><strong>分 类</strong></td>
      <td width="70" align="center"><strong>主 人</strong></td>
      <td align="center"><strong>创建时间</strong></td>
    </tr>

<%
 dim param:param=" where status=1"
 if classid<>0 then param=param & " and a.classid=" & classid
 if recommend<>0 then param=param & " and recommend=1"

 if ks.s("key")<>"" then param=param & " and blogname like '%" & ks.r(ks.s("key")) &"%'"
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			rsobj.open "select a.*,b.classname from ks_blog a inner join ks_blogclass b on a.classid=b.classid " & param & " order by adddate desc" ,conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 	response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=4><p>对不起，没有找到空间! </p></td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call ShowSpaceList(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call ShowSpaceList(RSObj)
									Else
										CurrentPage = 1
										call ShowSpaceList(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
  End Sub

  Sub ShowSpaceList(rs)
   dim i
   do while not rs.eof
    if i mod 2=0 then
	 response.write "<tr style=""background:#fff;height:22px"">"
	else
	 response.write "<tr style=""background:#FBFDFF;height:22px"">"
	end if
	%>
      <td><img src="images/card.gif" align="absmiddle" />
	  <%
		  dim spacedomain,predomain
		  If KS.SSetting(14)="1" and not conn.execute("select username from ks_blog where username='" & rs("username") & "'").eof Then
		   predomain=conn.execute("select [domain] from ks_blog where username='" & rs("username") & "'")(0)
		  end if
		  if predomain<>"" then
		   'spacedomain="http://" & predomain &"." & Right(ks.setting(2),Len(ks.setting(2))-InStr(ks.setting(2),"."))
		   spacedomain="http://" & predomain & "." & KS.SSetting(16)
		  else
		   If KS.SSetting(21)="1" Then
		    spacedomain=KS.GetDomain & "space/" &rs("username")
		   else
		    spacedomain=KS.GetDomain & "space/?" &rs("username")
		   end if
		  end if
		  %>
      <a title="<%=rs("blogname")%>" href="<%=spacedomain%>" target="blank">
	  <%=rs("blogname")%></a>	 
	   <%if rs("recommend")=1 then response.write "<font color=red>[荐]</font>"%>
</td>
      <td align="center"><%=rs("classname")%></td>
      <td align="center"><%=rs("username")%></td>
      <td align="center"><%=rs("adddate")%></td>
    </tr>
  <%

   rs.movenext
	  	I = I + 1
		  If I >= MaxPerPage Then Exit Do
	  loop
  End Sub

    '圈子列表
	Sub GroupList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(11))
		 dim classid:classid=KS.ChkClng(KS.S("ClassID"))
		 dim recommend:recommend=KS.ChkClng(KS.S("recommend"))
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 %>
		   <table border="0" cellpadding="1" cellspacing="1" width="98%" backcolor="#efefef">
      <tr height="22">
      <td colspan=2><strong>分类查看:</strong>
	   <select name="classid" onchange="SpacePage(1,'group&classid='+this.value);">
	    <option value='0'>所有分类</option>
	   <% dim rsc:set rsc=conn.execute("select classname,classid from ks_teamclass order by orderid")
	   if not rsc.eof then
	    do while not rsc.eof
		 if classid=rsc(1) then
		 response.write "<option value=""" & rsc(1) & """ selected>" & rsc(0) & "</option>"
		 else
		 response.write "<option value=""" & rsc(1) & """>" & rsc(0) & "</option>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
	   %>
	   </select>
	  </td>
      <td align="center" colspan=2><strong>圈子名称：</strong><input style="border:1px #000 dashed;height:18px;" type="text" size="12" name="key">&nbsp;&nbsp;<input type="button" onclick="SpacePage(1,'group&key='+document.getElementById('key').value);" value= " 查 找 "></td>
    </tr>
	     </table>
		 <%
		  dim param:param=" where verific=1"
          if classid<>0 then param=param & " and a.classid=" & classid
			 if recommend<>0 then param=param & " and  recommend=1"
		 if ks.s("key")<>"" then param=param & " and teamname like '%" & ks.r(ks.s("key")) &"%'"
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select a.*,b.classname from KS_team a inner join ks_teamclass b on a.classid=b.classid " & Param & " order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>没有用户创建圈子！</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call ShowGroup(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call ShowGroup(RSObj)
									Else
										CurrentPage = 1
										call ShowGroup(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
	 End Sub
			 
	 Sub ShowGroup(RS)		 
		 Dim I
		 Do While Not RS.Eof 
		 Response.Write "<tr style=""margin:2px;border-bottom:#9999CC dotted 1px;"">"
		   Response.Write "<td width=""20%"" style=""border-bottom:#9999CC dotted 1px;"">"& vbcrlf
			Response.Write " <table style=""BORDER-COLLAPSE: collapse"" borderColor=#c0c0c0 cellSpacing=0 cellPadding=0 border=1>"
			Response.Write "	<tr>"
			Response.Write "		<td><a href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""><img src=""" & rs("photourl") & """ width=""110"" height=""80"" border=""0""></a></td>"
			Response.Write "	 </tr>"
			Response.Write " </table>"
			Response.Write "</td>"
			Response.Write " <td style=""border-bottom:#9999CC dotted 1px;""><a class=""teamname"" href=""group.asp?id=" & rs("id") & """ title=""" & rs("teamname") & """ target=""_blank""> " & rs("TeamName") & "</a><br>创建者：" & rs("username") & "<br>创建时间:" &rs("adddate") & "<br>圈子分类：" & rs("classname") & "<br>主题/回复：" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id") & "  and parentid=0")(0) & "/" & conn.execute("select count(id) from ks_teamtopic where teamid=" & rs("id"))(0) & "&nbsp;&nbsp;&nbsp;成员:" & conn.execute("select count(username)  from ks_teamusers where status=3 and teamid=" & rs("id"))(0) & "人  </td>"
			Response.Write "</tr>"
			Response.Write "<tr><td height='2'></td></tr>"
			rs.movenext
			I = I + 1
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	   '相册列表
  Sub PhotoList()
		 MaxPerPage =KS.ChkClng(KS.SSetting(12))
		 dim classid:classid=ks.chkclng(ks.s("classid"))
		 dim recommend:recommend=ks.chkclng(ks.s("recommend"))
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
%>
  <table border="0" cellpadding="1" cellspacing="1" width="98%" backcolor="#efefef">
      <tr height="22">
      <td colspan=2><strong>分类查看:</strong>
	   <select name="classid" onchange="SpacePage(1,'photo&classid='+this.value);">
	    <option value='0'>所有分类</option>
	   <% dim rsc:set rsc=conn.execute("select classname,classid from ks_PhotoClass order by orderid")
	   if not rsc.eof then
	    do while not rsc.eof
		 if classid=rsc(1) then
		 response.write "<option value=""" & rsc(1) & """ selected>" & rsc(0) & "</option>"
		 else
		 response.write "<option value=""" & rsc(1) & """>" & rsc(0) & "</option>"
		 end if
		 rsc.movenext
		loop
	   end if
	   rsc.close:set rsc=nothing
	   %>
	   </select>
	  </td>
      <td align="center" colspan=2><strong>相册名称：</strong><input style="border:1px #000 dashed;height:18px;" type="text" size="12" name="key">&nbsp;&nbsp;<input type="button" onclick="SpacePage(1,'photo&key='+document.getElementById('key').value);" value= " 查 找 "></td>
    </tr>
    </table>
<%
 dim param:param=" where status=1"
 if classid<>0 then param=param & " and  classid=" & classid
 if recommend<>0 then param=param & " and  recommend=1"
 if ks.s("key")<>"" then param=param & " and XCName like '%" & ks.r(ks.s("key")) &"%'"
		 response.write "  <table border=""0"" align=""center"" width=""100%"">" & vbcrlf
		  If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_Photoxc " & Param & " order by id desc",Conn,1,1
		         If RSObj.EOF and RSObj.Bof  Then
				 response.write "<tr><td style=""border: #efefef 1px dotted;text-align:center"" colspan=3>没有创建相册！</td></tr>"
				 Else
							totalPut = RSObj.RecordCount
                           If CurrentPage < 1 Then	CurrentPage = 1
			
									If (totalPut Mod MaxPerPage) = 0 Then
										pagenum = totalPut \ MaxPerPage
									Else
										pagenum = totalPut \ MaxPerPage + 1
									End If
								If CurrentPage = 1 Then
									call showphoto(RSObj)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										call showphoto(RSObj)
									Else
										CurrentPage = 1
										call showphoto(RSObj)
									End If
								End If
				           End If
		 
		 response.write  "            </table>" & vbcrlf
		 Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|个||2"
		 RSObj.Close:Set RSObj=Nothing
  End Sub

	 Sub showphoto(rs)
	 	 Dim I,k
		 Do While Not RS.Eof 
		  Response.Write "<tr>"
		   for k=1 to KS.ChkClng(KS.SSetting(13))
		   %>
		  <td width="33%" height="22" align="center">
						<table borderColor=#b2b2b2 height=149 cellSpacing=0 cellPadding=0 width="110%" border=0>
							  <tr>
								 <td align=middle width="100%"><B><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>"><%=rs("xcname")%></a></B></td>
							  </tr>
							  <tr>
									  <td align=middle width="100%">
														<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0>
														  <tr>
															<td background="images/pic.gif" width="136" height="106" valign="top"><a href="../space/?<%=rs("username")%>/showalbum/<%=rs("id")%>" target="_blank"><img style="margin-left:6px;margin-top:5px" src="<%=rs("photourl")%>" width="120" height="90" border=0></a></td>
														  </tr>
														</table>
									  </td>
								</tr>
								<tr>
								  <td align=middle width="100%" height=20><%=rs("xps")%>张/<%=rs("hits")%>次<font color=red>[<%=GetStatusStr(rs("flag"))%>]</font></td>
						      </tr>
			  </table>
			 </td>
		   <%
			rs.movenext
			I = I + 1
			if rs.eof or i>=cint(MaxPerPage) then exit for
		   Next
		   
		   Response.Write "</tr>"
		  If I >= MaxPerPage Then Exit Do
		 Loop
	 End Sub
	 
	 Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
	 End Function

 End Class 
%>
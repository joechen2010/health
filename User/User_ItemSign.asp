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
Set KSCls = New User_ItemSign
KSCls.Kesion()
Set KSCls = Nothing

Class User_ItemSign
        Private KS,KSUser
		Private CurrentPage,totalPut
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
		Call KSUser.InnerLocation("文档签收")
		 If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
							%>
		<div class="tabs">	
			<ul>
				<li class="select"><a href="User_ItemSign.asp">文档签收</a></li>
			</ul>
		</div>
		<%
		dim fieldstr
		if databasetype=1 then
		  fieldstr="cast(signuser as nvarchar(4000))"
		else
		  fieldstr="signuser"
		end if
		%>
			<div style="text-align:right"> <a href='User_ItemSign.asp'><font color=red>·所有文档</font></a> ·<a href='?t=1'>未签收文档[<%=conn.execute("select count(id) from ks_article Where issign=1 and ','+"& fieldstr&"+',' like '%," & KSUser.UserName &",%' and id not in(select infoid from ks_itemsign where channelid=1 and username='" & ksuser.username & "')")(0)%>]</a> ·<a href='?t=2'>已签收文档[<%=conn.execute("select count(id) from ks_article Where issign=1 and ','+"& fieldstr&"+',' like '%," & KSUser.UserName &",%' and id in(select infoid from ks_itemsign where channelid=1 and username='" & ksuser.username & "')")(0)%>]</a>
		   </div>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class=title align=middle>
					  <td height="25">文档标题</td>
					  <td width=80>录入</td>
					  <td>添加时间</td>
					  <td>点击数</td>
					  <td>是否签收</td>
					  <td>查看签收</td>
					</tr>
					<%
					dim param:param=" Where issign=1 and ','+"& fieldstr&"+',' like '%," & KSUser.UserName &",%'" 
					if request("t")="1" then
					  param=param &" and id not in(select infoid from ks_itemsign where channelid=1 and username='" & ksuser.username & "')"
					elseif request("T")="2" then
					  param=param &" and id in(select infoid from ks_itemsign where channelid=1 and username='" & ksuser.username & "')"
					end if
					SqlStr="Select a.* From KS_Article a " & Param & " order by id desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>没有可签收的文档!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
									If CurrentPage < 1 Then	CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowContent
									Else
										CurrentPage = 1
										Call ShowContent
									End If
								End If
				End If

						
						 %>
					
          </table>
		  </td>
		  </tr>
</table>
		  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
  End Sub
    
  Sub ShowContent()
     Dim I,url
     Do While Not rs.eof 
	 Url=KS.GetItemUrl(1,rs("tid"),rs("id"),rs("fname"))
	%>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td height='25' class="splittd" align=middle><a href="<%=url%>" target="_blank"><%=rs("Title")%></a></td>
      <td  class="splittd" align=middle><%=rs("inputer")%></td>
      <td  class="splittd" align=middle><%=rs("adddate")%></td>
	  <td  class="splittd" align=middle><%=rs("hits")%></td>
      <td  class="splittd" align=middle>
	  <% 
	  Dim SignTF:SignTF=conn.execute("select top 1 username from ks_itemsign where username='" & ksuser.username & "' and channelid=1 and infoid=" & rs("id")).eof
	  if SignTF then
	    response.write "<font color=red>未签收</font>"
	  else
	    response.write "<font color=blue>已签收</font>"
	  end if
	 %>
	  </td>
      <td  class="splittd" align=middle> 
	  <a href="<%=url%>" target="_blank">
	  <%if SignTF then%>
	  查看签收
	  <%else%>
	  浏览
	  <%end if%>
	  </a>
	  </td>
    </tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
   
  </table>
		<%
		End Sub
  
End Class
%> 

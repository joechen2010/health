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
Set KSCls = New User_Photo
KSCls.Kesion()
Set KSCls = Nothing

Class User_Photo
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather,PhotoUrls,descript
		Private XCID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
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
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		ElseIf KS.SSetting(0)=0 Then
		 Call KS.Alert("对不起，本站关闭个人空间功能！","")
		 Exit Sub
		ElseIf Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
		 Call KS.Alert("你不对，你还没有开通空间功能！","User_Blog.asp")
		 Exit Sub
		ElseIf Conn.Execute("Select status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
		    Response.Write "<script>alert('对不起，你的空间还没有通过审核或被锁定！');history.back();</script>"
			response.end
		End If

		Call KSUser.Head()
		Call KSUser.InnerLocation("我的相册")
		KSUser.CheckPowerAndDie("s05")
		%>
		<div class="tabs">	
		   <ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">我的相册</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">已审相册(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=1")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="?Status=0">待审相册(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=0")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">锁定相册(<span class="red"><%=conn.execute("select count(id) from ks_photoxc where username='" & ksuser.username & "' and status=2")(0)%></span>)</a></li>
			</ul>
        </div>
			 <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="User_Photo.asp?Action=Add"><span style="font-size:14px;color:#ff3300">上传照片</span></a>
			  <img src="images/fav.gif" width="20" align="absmiddle"><a href="User_Photo.asp?Action=Createxc"><span style="font-size:14px;color:#ff3300">创建相册</span></a>
			 </div>

		<%

			Select Case KS.S("Action")
			 Case "Del"
			  Call Delxc()
			 Case "Delzp"
			  Call Delzp()
			 Case "Editzp"
			  Call Editzp()
			 Case "Add"
			  Call Addzp()
			 Case "AddSave"
			  Call AddSave()
			 Case "EditSave"
			  Call EditSave()
			 Case "ViewZP"
			  Call ViewZP()
			 Case "Editxc","Createxc"
			  Call Managexc()
			 Case "photoxcsave"
			  Call photoxcsave()
			 Case Else
			  Call PhotoxcList()
			End Select
	   End Sub
	   '查看照片
	   Sub ViewZP()
	    Dim title
	    Dim xcid:xcid=KS.Chkclng(KS.S("XCID"))
	    Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "select xcname from KS_Photoxc WHERE ID=" & XCID,CONN,1,1
		if rs.Eof And RS.Bof Then 
		 rs.close:set rs=nothing
		 response.write "<script>alert('参数传递出错！');history.back();</script>"
		 response.end
		end if
		title=rs(0)
		rs.close
		Call KSUser.InnerLocation("查看照片")
	  			  %>
			   
	   		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
            <tr class="title">
              <td align=center colspan=5><%=Title%></td>
            </tr>
			<%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
			 rs.open "select * from KS_PhotoZP where xcid=" & xcid,conn,1,1
			if rs.eof and rs.bof then
			  response.write "<tr class='tdbg'><td  height='30' colspan='5'>该相册下没有相片，请<a href=""?action=Add&xcid=" & xcid &""">上传</a>！</td></tr>"
			else
			 				  MaxPerPage =5
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
									Call showzplist(xcid)
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showzplist(xcid)
									Else
										CurrentPage = 1
										Call showzplist(xcid)
									End If
								End If
        end if%>
      </table>
	  <div style="padding-right:30px">
	  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
	  </div>
<%End Sub
sub showzplist(xcid)
%>
    <script type="text/javascript" src="../ks_inc/highslide/highslide.js"></script>
    <link href="../ks_inc/highslide/highslide.css" type=text/css rel=stylesheet>
	<script type="text/javascript">
		hs.graphicsDir = '/ks_inc/highslide/graphics/';
		hs.transitions = ['expand', 'crossfade'];
		hs.wrapperClassName = 'dark borderless floating-caption';
		hs.fadeInOut = true;
		hs.dimmingOpacity = .75;
		
		if (hs.addSlideshow) hs.addSlideshow({
			interval: 5000,
			repeat: false,
			useControls: true,
			fixedControls: 'fit',
			overlayOptions: {
				opacity: .6,
				position: 'bottom center',
				hideOnMouseOut: true
			}
		});
	</script>
<%
     Dim I
    Response.Write "<FORM Action=""?Action=Delzp"" name=""myform"" method=""post"">"
			 do while not rs.eof
			 %>
			<tr class="tdbg"> 
            <td width="21%" rowspan="5">
			<table border="0" align="center" cellpadding="2" cellspacing="1" class="border">
                <tr> 
                  <td><a href="<%=rs("photourl")%>" class="highslide" onClick="return hs.expand(this)"  title="<%=rs("title")%>"><img src="<%=rs("photourl")%>" width="85" height="100" border="0"></a>
                  </td>
                </tr>
              </table></td>
            <td width="12%" class="tdbg"><div align="center"><strong>相片名称：</strong></div></td>
            <td width="40%" class="tdbg"><font style="font-size:14px"><strong><%=rs("title")%></strong></font></td>
            <td width="10%"><div align="center">浏览次数：</div></td>
            <td width="17%"><%=rs("hits")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">创建日期：</div></td>
            <td><%=rs("adddate")%></td>
            <td><div align="center">图片大小：</div></td>
            <td><%=rs("photosize")%>byte</td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">相片地址：</div></td>
            <td colspan="3"><%=rs("photourl")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">相片描述：</div></td>
            <td colspan="3"><%=rs("descript")%></td>
          </tr>
          <tr class="tdbg"> 
            <td><div align="center">所属相册：</div></td>
            <td><%=conn.execute("select xcname from ks_photoxc where id=" & xcid)(0)%></td>
            <td colspan="2" height="28"><div align="center"><a href="?Action=Editzp&Id=<%=rs("id")%>" class="box">修改</a> <a href="?id=<%=rs("id")%>&Action=Delzp" onClick="{if(confirm('确定删除该照片吗？')){return true;}return false;}" class="box">删除</a> 
                <INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
              </div></td>
          </tr>
          <tr> 
            <td colspan="5" height="3" class="splittd">&nbsp;</td>
          </tr>
			<% rs.movenext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
			 loop
		 %>
		 <tr class="tdbg">
		   <td colspan="5" align="right">
		  								&nbsp;&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中本页显示的所有照片&nbsp;<INPUT class="button" onClick="return(confirm('确定删除选中的照片吗?'));" type=submit value=删除选定的照片 name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        </td>
		 </tr>
		 </form>
		 <%
	   End Sub
	    '相册，添加／修改
	   Sub Managexc()
	    Dim xcname,ClassID,Descript,PhotoUrl,PassWord,ListReplayNum,ListGuestNum,OpStr,TipStr,TemplateID,Flag,ListLogNum
		Dim ID:ID=KS.ChkCLng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Photoxc Where ID=" & ID,conn,1,1
		If Not RS.EOF Then
		Call KSUser.InnerLocation("修改相册")
		 xcname=RS("xcname")
		 ClassID=RS("ClassID")
		 Descript=RS("Descript")
		 flag=RS("Flag")
		 PhotoUrl=RS("PhotoUrl")
		 PassWord=RS("PassWord")
		 OpStr="OK了，确定修改":TipStr="修 改 我 的 相 册"
		Else
		 Call KSUser.InnerLocation("创建相册")
		 xcname=FormatDatetime(Now,2)
		 ClassID="0"
		 flag="1"
		 PhotoUrl=""
		 OpStr="OK了，立即创建":TipStr="创 建 我 的 相 册"
		End if
		RS.Close:Set RS=Nothing
	    %>
		<script>
		 function CheckForm()
		 {
		  if (document.myform.xcname.value=='')
		  {
		   alert('请输入相册名称!');
		   document.myform.xcname.focus();
		   return false;
		  }
		  if (document.myform.ClassID.value=='0')
		  {
		   alert('请选择相册类型!');
		   document.myform.ClassID.focus();
		   return false;
		  }
		  return true;
		 }

		</script>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
          <form  action="User_Photo.asp?Action=photoxcsave&id=<%=id%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
            <tr class="title">
              <td colspan=2 align=center><%=TipStr%></td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>相册名称：</strong><br>
              请给你的相册取个合适的名称。
              </div></td>
              <td> 　
                  <input class="textbox" name="xcname" type="text" id="xcname" style="width:250px; " value="<%=xcname%>" maxlength="100" />
              <span style="color: #FF0000">*</span></td>
            </tr>
<tr class="tdbg">
              <td width="24%"  height="25" align="center"><div align="left"><strong>相册分类：</strong><br>
      相册分类，以便查找浏览</div></td>
              <td width="76%">　
                  <select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-请选择类别-</option>
                    <% Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If ClassID=RS("ClassID") Then
								  Response.Write "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                  </select>               </td>
            </tr>
			<tr class="tdbg"> 
                  <td height="30"><div align="left"><strong>是否公开：</strong><br>
                  可以设置为只有权限的用户才能浏览。</div></td>
                  <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse">
                   <tr>
                     <td width="50%" align="left">&nbsp;
                       <select style="width:160px" onChange="if(this.options[selectedIndex].value=='3'){document.myform.all.mmtt.style.display='block';}else{document.myform.all.mmtt.style.display='none';}"  name="flag">
                      <option value="1"<%if flag="1" then response.write " selected"%>>完全公开</option>
                      <option value="2"<%if flag="2" then response.write " selected"%>>会员开见</option>
                      <option value="3"<%if flag="3" then response.write " selected"%>>密码共享</option>
                      <option value="4"<%if flag="4" then response.write " selected"%>>隐私相册</option>
                    </select></td>
                   <td width="50%"><span class=child id=mmtt name="mmtt" <%if flag<>3 then%>style="display:none;"<%end if%>>密码：<input type="password" name="password" style="width:160px" maxlength="16" value="<%=password%>" size="20"></span>                  </td>
                  </tr>
                  </table>                  </td>
            </tr>
            <tr class="tdbg">
              <td  height="25" align="center"><div align="left"><strong>相册封面：</strong><br>
                  您可以上传您喜欢的图片做为相册的封面。</div></td>
              <td>　
                  <input class="textbox" name="PhotoUrl" type="text" id="PhotoUrl" style="width:230px; " value="<%=PhotoUrl%>" />             
                  &nbsp;只支持jpg、gif、png，小于50k，默认尺寸为85*100
				  <div>
                  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9998' frameborder="0" align="center" width='94%' height='30' scrolling="no"></iframe>
				  </div>
				  </td>
            </tr>
            <tr class="tdbg">
              <td  height="25"><div align="left"><span><strong>相册介绍：</strong></span></div>
                <br>
                关于此相册的简要文字说明。</td>
              <td>　
                  
                  <textarea name="Descript" id="Descript" cols=50 rows=6><%=Descript%></textarea>              </td>
            </tr>
            <tr class="tdbg">
              <td height="30" align="center" colspan=2>
                <input type="submit" name="Submit3"  class="Button" value="<%=OpStr%>" />
                <input type="reset" name="Submit22"   class="Button" value=" 重 来 " />              </td>
            </tr>
          </form>
</table>
		<%
	   End Sub
	   '保存相册
	   Sub photoxcsave()
	     Dim xcname:xcname=KS.S("xcname")
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Descript:Descript=KS.S("Descript")
		 Dim Flag:Flag=KS.S("Flag")
		 Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
		 Dim PassWord:PassWord=KS.S("PassWord")
		 Dim ID:ID=KS.Chkclng(KS.S("id"))
		 If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="/images/user/nopic.gif"
		 If xcname="" Then Response.Write "<script>alert('请输入相册名称!');history.back();</script>"
		 If ClassID=0 Then Response.Write "<script>alert('请选择相册类型!');history.back();</script>"
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photoxc Where id=" & id ,conn,1,3
		 If RS.Eof And RS.Bof Then
		   RS.AddNew
		    RS("AddDate")=now
			if ks.SSetting(4)=1 then
			RS("Status")=0 '设为已审
			else
			RS("Status")=1 '设为已审
			end if
		 End If
		    RS("UserName")=KSUser.UserName
		    RS("xcname")=xcname
			RS("ClassID")=ClassID
			RS("Descript")=Descript
			RS("Flag")=Flag
			RS("Password")=PassWord
			RS("PhotoUrl")=PhotoUrl
		  RS.Update
		  RS.MoveLast
		  ID=rs("id")
		  RS.Close:Set RS=Nothing
		  If KS.Chkclng(KS.S("id"))=0 Then
		   Call KS.FileAssociation(1028,ID,PhotoUrl,0)
		   Call KSUser.AddLog(KSUser.UserName,"创建了相册!名称: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """ target=""_blank"">查看</a>",104)
		   Response.Write "<script>alert('恭喜!相册创建成功,进入上传照片');location.href='User_Photo.asp?action=Add&xcid=" & id &"';</script>"
		  Else
		   Call KS.FileAssociation(1028,ID,PhotoUrl,1)
		   Call KSUser.AddLog(KSUser.UserName,"修改了相册!名称: "&xcname & " <a href=""../space/?" & KSUser.UserName & "/showalbum/" & id & """  target=""_blank"">查看</a>",104)
		   Response.Write "<script>alert('相册修改成功!');location.href='User_Photo.asp';</script>"
		  End If
	   End Sub


	  
	   '相册列表
	   Sub PhotoxcList()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where UserName='"& KSUser.UserName &"'"
									IF KS.S("status")<>"" Then
									  Param=Param & " And status=" & KS.ChkClng(KS.S("status"))
									End if
									
									
									'If KS.S("XCID")<>"" And KS.S("XCID")<>"0" Then Param=Param & " And XCID=" & KS.ChkClng(KS.S("XCID")) & ""
									Dim Sql:sql = "select * from KS_Photoxc "& Param &" order by AddDate DESC"


								    Call KSUser.InnerLocation("所有相册列表")
								  %>
								     
				                     <table width="98%"  border="0" align="center" cellpadding="3" cellspacing="1">
                                                <tr class="Title">
                                                  <td colspan="6" height="22" align="center">我 的 相 册</td>
                                                </tr>
                                           
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>您还没有创建相册!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
											If CurrentPage < 1 Then
												CurrentPage = 1
											End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowXC
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowXC
									Else
										CurrentPage = 1
										Call ShowXC
									End If
								End If
				End If
     %>                      
                        </table>
		  <%
  End Sub
  
  Sub ShowXC()
     Dim I,K
   Do While Not RS.Eof
         %>
           <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
		   <%
		   For K=1 To 4
		   %>
            <td width="25%" height="22" align="center">
									  <table width=154 height=185 border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" id=AutoNumber2 style="BORDER-COLLAPSE: collapse">
										  <td width=123 height=185>
											<table id=AutoNumber3 style="BORDER-COLLAPSE: collapse" borderColor=#b2b2b2 height=179 cellSpacing=0 cellPadding=0 width="117%" border=0>
											  <tr>
												<td width="100%" height=179>
												  <table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 width="99%" border=0>
													<tr>
													  <td align=middle width="100%" height=22><B><a href="?xcid=<%=rs("id")%>&action=ViewZP"><%=ks.gottopic(rs("xcname"),18)%></a></B><%select case rs("status")
													     case 1:response.write "[已审]"
														 case 2:response.write "<font color=blue>[锁定]</font>"
														 case 0:response.write "<font color=red>[未审]</font>"
														end select
														%>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%">
														<table style="BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0>
														  <tr>
															<td background="images/pic.gif" width="136" height="106" valign="top"><a href="?xcid=<%=rs("id")%>&action=ViewZP"><img style="margin-left:6px;margin-top:5px" src="<%=rs("photourl")%>" width="120" height="90" border=0></a></td>
														  </tr>
														</table>
													  </td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><%=rs("xps")%>张/<%=rs("hits")%>次</td>
													</tr>
													<tr>
													  <td align=middle width="100%" height=23><a href="?Action=Editxc&id=<%=rs("id")%>">修改</a>&nbsp;<a href="?Action=Del&id=<%=rs("id")%>" onClick="return(confirm('删除相册将删除该相册里的所有照片，确定删除吗？'))">删除</a>&nbsp;
													  <% select case rs("flag")
													      case 1
													       response.write "<font color=red>[公开]</font>"
														  case 2
													       response.write "<font color=red>[会员]</font>"
														  case 3
													       response.write "<font color=red>[密码]</font>"
														  case 4
													       response.write "<font color=red>[稳私]</font>"
														 end select
													%>
													  </td>
													</tr>
												  </table>
												</td>
											  </tr>
											</table>
										  </td>
										</tr>
			  </table>
			 </td>
                       
					                  <%
							RS.MoveNext
							I=I+1
					  If I >= MaxPerPage Or RS.Eof Then Exit For
				  Next
			      do While K<4 
				   response.write "<td width=""25%""></td>"
				   k=k+1
				  Loop%>
		    </tr>
				 <%
					  If I >= MaxPerPage Or RS.Eof Then Exit do
	   Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=6 valign=top align="right">
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
								  </td>
								</tr>
								<% 
  End Sub
  '删除相册
  Sub Delxc()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("你没有选中要删除的相册!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_Photoxc Where ID In(" & ID & ")")
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where xcid in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid=" & rs("id"))
	   KS.DeleteFile(rs("photourl"))
	   rs.movenext
	   loop
	end if
	Conn.execute("delete from ks_photozp where xcid in(" & id& ")")
	Conn.execute("delete from ks_uploadfiles where channelid=1028 and infoid in(" & id& ")")
	rs.close:set rs=nothing
	Call KSUser.AddLog(KSUser.UserName,"删除了相册操作!",104)
	Response.Redirect ComeUrl
  End Sub
  '删除照片
  Sub Delzp()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("你没有选中要删除的照片!",ComeUrl):Response.End
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where id in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   KS.DeleteFile(rs("photourl"))
	   Conn.execute("update ks_photoxc set xps=xps-1 where id=" & rs("xcid"))
	   rs.movenext
	   loop
	end if
	Conn.Execute("Delete From KS_UploadFiles Where Channelid=1029 and infoid in(" & id& ")")
	Conn.execute("delete from ks_photozp where id in(" & id& ")")
	Call KSUser.AddLog(KSUser.UserName,"删除了相片操作!",104)
	rs.close:set rs=nothing
	Response.Redirect ComeUrl
  End Sub
  '上传照片
  Sub Addzp()
        Call KSUser.InnerLocation("上传照片")
		  adddate=now:XCID=KS.ChkCLng(KS.S("XCID")):UserName=KSUser.RealName
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.XCID.value=="0") 
				  {
					alert("请选择所属相册！");
					document.myform.XCID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入相片名称！");
					document.myform.Title.focus();
					return false;
				  }		
				 return true;  
				}
				</script>
				<script>  
			var FFextraHeight = 0;
			 if(window.navigator.userAgent.indexOf("Firefox")>=1)
			 {
			  FFextraHeight = 16;
			  }
			 function ReSizeiFrame(iframe)
			 {
			   if(iframe && !window.opera)
			   {
				 iframe.style.display = "block";
				  if(iframe.contentDocument && iframe.contentDocument.body.offsetHeight)
				  {
					iframe.height = iframe.contentDocument.body.offsetHeight + FFextraHeight;
				  }
				  else if (iframe.Document && iframe.Document.body.scrollHeight)
				  {
					iframe.height = iframe.Document.body.scrollHeight;
				  }
			   }
			 }
			function init()
			 {
			   ReSizeiFrame(document.getElementById('UpPhotoFrame'));
			 }
			
			</script>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=AddSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>上 传 照 片</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>选择相册：</span></td>
                       <td width="88%"><select class="textbox" size='1' name='XCID' style="width:150">
                             <option value="0">-请选择相册-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_Photoxc where username='" & KSUser.Username & "' order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>照片名称：</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*
                                        <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" />
                                        </span></td>
                    </tr>
								<tr class="tdbg">
								  <td height="20" align="center">照片预览：</td>
								  <td align='center' id="viewarea">
								     
								</td>
				    </tr>
					
								<tr class="tdbg">
                                   <td height="250" align="center"><span>上传照片：</span></td>
                                   <td align="center"><iframe onload="ReSizeiFrame(this)" onreadystatechange="ReSizeiFrame(this)" id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9997' frameborder="0" scrolling="auto" width='100%' height='92%'></iframe></td>
							  </tr>							 
								<tr class="tdbg">
                                   <td height="25" align="center"><span>照片介绍：</span></td>
                                  <td><textarea class="textbox" style="height:50px" name="Descript" cols="70" rows="5"></textarea></td>
							  </tr>							 
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,立即发布 " />
                      <input type="reset" name="Submit2"   class="Button" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub
    '编辑照片
  Sub Editzp()
        Call KSUser.InnerLocation("编辑照片")
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_PhotoZp Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_A_RS_Obj.Eof Then
		     XCID  = KS_A_RS_Obj("XCID")
			 Title    = KS_A_RS_Obj("Title")
			 UserName   = KS_A_RS_Obj("UserName")
			 descript = ks_a_rs_obj("descript")
			 PhotoUrlS  = KS_A_RS_Obj("PhotoUrl")
		   End If
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		%>
		<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.XCID.value=="0") 
				  {
					alert("请选择所属相册！");
					document.myform.XCID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入相片名称！");
					document.myform.Title.focus();
					return false;
				  }		
				 return true;  
				}
				
				</script>
				
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_Photo.asp?Action=EditSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr class="title">
					  <td colspan=2 align=center>上 传 照 片</td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>选择相册：</span></td>
                       <td width="88%"><select class="textbox" size='1' name='XCID' style="width:150">
                             <option value="0">-请选择相册-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_Photoxc order by id desc",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If XCID=RS("ID") Then
								  Response.Write "<option value=""" & RS("ID") & """ selected>" & RS("XCName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>照片名称：</span></td>
                              <td><input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*
                                        <input class="textbox" name="PhotoUrls" type="hidden" id="PhotoUrls" style="width:350px; " maxlength="100" value="<%=photourls%>"/>
                                        </span></td>
                    </tr>
								<tr class="tdbg">
								  <td height="20" align="center">照片预览：</td>
								  <td id="viewarea">
								    <table style='BORDER-COLLAPSE: collapse' borderColor='#c0c0c0' cellSpacing='1' cellPadding='2' border='1'><tr><td align='center' width='83' height='100' bgcolor='#ffffff'><img name='view1' width='83' height='100' src='<%=Photourls%>' title='照片预览'></td></tr></table> <input class="button" type='button' name='Submit3' value='选择照片地址...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择图片")%>&ChannelID=9997',500,360,window,document.myform.PhotoUrls);" />
								</td>
				    </tr>
														 
								<tr class="tdbg">
                                   <td height="25" align="center"><span>照片介绍：</span></td>
                                  <td><textarea class="textbox" style="height:50px" name="Descript" cols="70" rows="5"><%=DESCRIPT%></textarea></td>
							  </tr>							 
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input type="submit" name="Submit"  class="Button" value=" OK,立即发布 " />
                      <input type="reset" name="Submit2"   class="Button" onClick="javascript:history.back()" value=" 取 消 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub

   Sub EditSave()
    Dim RSObj,Descript,PhotoUrlArr,i
                 XCID=KS.ChkClng(KS.S("XCID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('你没有上传相片!');history.back();</script>"
				    Exit Sub
				  End IF
				  on error resume next
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_PhotoZP Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				  RSObj("Title")=Title
				  RSObj("XCID")=XCID
				  RSObj("PhotoUrl")=PhotoUrls
				  RSObj("Descript")=Descript
				  RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(replace(PhotoUrls,ks.getdomain,ks.setting(3))))
				RSObj.Update
				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(1029,KS.ChkClng(KS.S("ID")),PhotoUrls,1)
				 Call KSUser.AddLog(KSUser.UserName,"修改了相片操作! <a href=""" & PhotoUrls & """ target=""_blank"">查看</a>",104)
				 Response.Write "<script>alert('相片修改成功!');location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';</script>"
  End Sub
  
  Sub AddSave()
    Dim RSObj,Descript,PhotoUrlArr,i,UpFiles
                 XCID=KS.ChkClng(KS.S("XCID"))
				 Title=Trim(KS.S("Title"))
				 UserName=Trim(KS.S("UserName"))
				 Descript=KS.S("Descript")
				 PhotoUrls=KS.S("PhotoUrls")
				 If PhotoUrls="" Then 
				    Response.Write "<script>alert('你没有上传相片!');history.back();</script>"
				    Exit Sub
				  End IF
				PhotoUrlArr=Split(PhotoUrls,"|")
				 
				  If XCID=0 Then
				    Response.Write "<script>alert('你没有选择相册!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入相片名称!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_PhotoZP",Conn,1,3
				 For I=0 to ubound(PhotoUrlArr)
			    	RSObj.AddNew
					 RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(Replace(PhotoUrlArr(I),KS.GetDomain,KS.Setting(3))))
				     RSObj("Title")=Title
				     RSObj("XCID")=XCID
					 RSObj("UserName")=KSUser.UserName
					 RSObj("PhotoUrl")=PhotoUrlArr(I)
					 RSObj("Adddate")=Now
					 RSObj("Descript")=Descript
				   RSObj.Update
				   RSObj.MoveLast
				   Call KS.FileAssociation(1029,RSObj("ID"),PhotoUrlArr(i),0)
				 Next
				 RSObj.Close
				 RSObj.Open "Select Top 1 PhotoUrl From KS_PhotoXC Where ID=" & xcid,conn,1,3
				 If Not RSObj.Eof Then
				    If Instr(lcase(RSObj(0)),"nopic.gif")>0 then
					  RSObj(0)=PhotoUrlArr(0)
					  RSObj.Update
					end if
				 End If
				 RSObj.Close
				 Set RSObj=Nothing
				 
				 
				 Conn.Execute("update KS_Photoxc set xps=xps+" & Ubound(PhotoUrlArr)+1 & " where id=" & xcid)
				 Call KSUser.AddLog(KSUser.UserName,"上传了" & Ubound(PhotoUrlArr)+1 & "张照片到相册! <a href=""../space/?" & KSUser.UserName & "/showalbum/" & xcid & """ target=""_blank"">查看</a>",104)
				 Response.Write "<script>if (confirm('相片保存成功，继续上传吗?')){location.href='User_Photo.asp?Action=Add';}else{location.href='User_Photo.asp?Action=ViewZP&XCID=" & XCID& "';}</script>"
  End Sub

End Class
%> 

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
Set KSCls = New EnterPriseNewsCls
KSCls.Kesion()
Set KSCls = Nothing

Class EnterPriseNewsCls
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,ClassID
		Private title,Content,Verific,Action,AddDate
		Private Sub Class_Initialize()
			MaxPerPage =12
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		Call KSUser.Head()
		Call KSUser.InnerLocation("所有新闻列表")
		KSUser.CheckPowerAndDie("s11")
		
		%>
		<div class="tabs">	
			<ul>
			  <li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="?">所有新闻(<span class="red"><%=conn.execute("select count(id) from KS_EnterPrisenews where username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="?Status=2">已审核(<span class="red"><%=conn.execute("select count(id) from KS_EnterPrisenews where status=1 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="?Status=1">待审核(<span class="red"><%=conn.execute("select count(id) from KS_EnterPrisenews where status=0 and username='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
        </div>
		<%
		Select Case KS.S("Action")
		 Case "Del"  Call ArticleDel()
		 Case "Add","Edit" Call ArticleAdd()
		 Case "DoSave" Call DoSave()
		 Case Else Call ArticleList()
		End Select
	   End Sub
	   Sub ArticleList()
			  
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
						   Dim Sql,Param:Param=" where UserName='" & KSUser.UserName & "'"
						   IF KS.S("Status")<>"" Then Param= Param & " and status=" & KS.ChkClng(KS.S("Status"))-1
                           If (KS.S("KeyWord")<>"") Then Param = Param  & " and title like '%" & KS.S("KeyWord") & "%'"
						   sql = "select * from KS_EnterPriseNews " & Param & " order by AddDate DESC"
								  %>
                                     <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="?Action=Add"><span style="font-size:14px;color:#ff3300">发布新闻</span></a></div>
				                     <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                                        <tr class="Title">
                                                  <td width="6%" height="22" align="center">选中</td>
                                                  <td width="41%" height="22" align="center">新闻标题</td>
                                                  <td width="15%" height="22" align="center"> 分 类</td>
												  <td width="16%" height="22" align="center">更新时间</td>
												  <td width="10%" height="22" align="center">状态</td>
                                                  <td height="22" align="center" nowrap>管理操作</td>
                                        </tr>
                                           
                                      <%
							Set RS=Server.CreateObject("AdodB.Recordset")
							RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=6 height=30 valign=top>没有你要的新闻!</td></tr>"
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
			
								   If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
								Call showContent
				End If
     %>                      
                        </table>			<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>

		  <%
  End Sub
  
  Sub ShowContent()
     Dim I
    Response.Write "<FORM Action=""?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                   <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                        <td class='splittd' width="5%" height="23" align="center">
						  <INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID">
						</td>
                        <td class='splittd' align="left"><a href="?Action=Edit&id=<%=rs("id")%>" class="link3"><%=KS.GotTopic(trim(RS("title")),45)%></a></td>
                        <td class='splittd' align="center">
						<%
						If RS("ClassID")=0 Then
						 Response.Write "没有指定分类"
						Else
						 on error resume next
						 Response.Write conn.execute("select classname from ks_userclass where classid=" & RS("ClassID"))(0)
						End If
						%></td>
                        <td class='splittd' align="center"><%=formatdatetime(rs("AddDate"),2)%></td>
                        <td class='splittd' align="center"><%
						if rs("status")=1 then
						 response.write "已审核"
						else
						 response.write "<font color=red>未审核</font>"
						end if
						%></td>
                        <td class='splittd' align="center">
						<a href="?id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>" class="link3">修改</a> <a href="?action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除新闻吗?'))" class="link3">删除</a>
										
						</td>
                     </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td colspan=2 valign=top>
								&nbsp;&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中本页显示的所有新闻<INPUT class="button" onClick="return(confirm('确定删除选中的新闻吗?'));" type=submit value=删除选定的新闻 name=submit1> </FORM> 
								</td>
								<td colspan="4">      
								<form action="User_EnterPriseNews.asp" method="post" name="searchform">  关键字<input type="text" name="KeyWord" class="textbox" value="关键字" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 "> </form>
								  </td>
								  
								</tr>
								<% 
  End Sub
  '删除文章
  Sub ArticleDel()
	Dim ID:ID=KS.S("ID")
	ID=KS.FilterIDs(ID)
	If ID="" Then Call KS.Alert("你没有选中要删除的新闻!",ComeUrl):Response.End
	Conn.Execute("Delete From KS_EnterPriseNews Where UserName='" & KSUser.UserName & "' and ID In(" & ID & ")")
	if ComeUrl="" then
	Response.Redirect("../index.asp")
	else
	Response.Redirect ComeUrl
	end if
  End Sub

  '添加文章
  Sub ArticleAdd()
        Call KSUser.InnerLocation("发布新闻")
  		if KS.S("Action")="Edit" Then
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select * From KS_EnterPriseNews Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_A_RS_Obj.Eof Then
			 Title    = KS_A_RS_Obj("Title")
			 Content  = KS_A_RS_Obj("Content")
			 AddDate  = KS_A_RS_Obj("AddDate")
			 ClassID  = KS_A_RS_Obj("ClassID")
		   End If
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		Else
		   AddDate=Now:ClassID=0
		End If
		%>
		<script language = "JavaScript">
				function CheckForm()
				{	
				if (document.myform.Title.value=="")
				  {
					alert("请输入新闻标题！");
					document.myform.Title.focus();
					return false;
				  }	
		
				    if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=="")
					{
					  alert("新闻内容不能留空！");
					  FCKeditorAPI.GetInstance('Content').Focus();
					  return false;
					}
				 return true;  
				}
				</script>
				
				
				<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr  class="title">
					  <td colspan=2 align=center>
					       <%IF KS.S("Action")="Edit" Then
							   response.write "修改新闻"
							   Else
							    response.write "发布新闻"
							   End iF
							  %> 
					 </td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>新闻标题：</span></td>
                       <td width="88%"><input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" />
                                        <span style="color: #FF0000">*</span> </td>
                    </tr>
					<tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>选择分类：</span></td>
                       <td colspan="2"><select class="textbox" size='1' name='ClassID' style="width:150">
                                            <option value="0">-不指定分类-</option>
                                            <%=KSUser.UserClassOption(4,ClassID)%>
                         </select>		
				
						 <a href="User_Class.asp?Action=Add&typeid=4"><font color="red">添加我的分类</font></a>					  </td>
                    </tr>
						  
                     <tr class="tdbg">
                                <td align="center">发布时间：</td>
                                <td><input class="textbox" readonly name="AddDate" type="text" style="width:250px; " value="<%=AddDate%>" maxlength="100" /></td>
                              </tr>
                              <tr class="tdbg">
                                  <td align="center">新闻内容：</td>
								  <td>
							<%	
								Response.Write "<textarea name=""Content"" style=""display:none"">" & KS.HtmlCode(Content) & "</textarea>"
								Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""95%"" height=""350"" frameborder=""0"" scrolling=""no""></iframe>"  
							%>								</td>
                            </tr>
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input class="button" type="submit" name="Submit" value="OK, 保 存 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
  End Sub
  
   Sub DoSave()
      Dim Id:Id=KS.ChkClng(Request("ID"))
				 Title=KS.LoseHtml(KS.S("Title"))
				 Content=KS.HtmlEncode(Request.Form("Content"))
				  Dim RSObj
				  
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入新闻标题!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>alert('你没有输入新闻内容!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_EnterpriseNews Where UserName='" & KSUser.UserName & "' and ID=" & Id,Conn,1,3
				If rsobj.eof then
				  RSObj.Addnew
				  RSObj("UserName")=KSUser.UserName
				  RSObj("Adddate")=Now
				  If KS.SSetting(18)=1 Then
				  RSObj("Status")=0
				  Else
				  RSObj("Status")=1
				  End If
				 End If
				  RSObj("Title")=Title
				  RSObj("Content")=Content
				  RSObj("ClassID")=KS.ChkClng(KS.S("ClassID"))
				 RSObj.Update
				 RSObj.MoveLast
				 Id=RSObj("ID")
				 RSObj.Close:Set RSObj=Nothing
				 IF KS.ChkClng(KS.S("id"))=0 Then
				   Call KSUser.AddLog(KSUser.UserName,"发布了一条企业新闻,<a href=""../space/show_news.asp?username=" & KSUser.UserName & "&id=" & id & """ target=""_blank"">" & Title & "</a>",201)
				   Response.Write "<script>if (confirm('成功添加新闻，继续添加吗?')){location.href='?Action=Add';}else{location.href='User_EnterPriseNews.asp';}</script>"
				 Else
				  Call KSUser.AddLog(KSUser.UserName,"修改了企业新闻,<a href=""../space/show_news.asp?username=" & KSUser.UserName & "&id=" & id & """ target=""_blank"">" & Title & "</a>",201)
				 Response.Write "<script>alert('新闻修改成功!');location.href='User_EnterpriseNews.asp';</script>"
				 End If
  End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
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
Dim KSCls
Set KSCls = New MyArticleCls
KSCls.Kesion()
Set KSCls = Nothing

Class MyArticleCls
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,Selbutton,LoginTF,ReadPoint
		Private F_B_Arr,F_V_Arr,ClassID,Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PhotoUrl,Action,I,UserDefineFieldArr,UserDefineFieldValueStr,Province,City
		Private Sub Class_Initialize()
			MaxPerPage =10
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=1
		LoginTF=Cbool(KSUser.UserLoginChecked)
		IF LoginTF=false  Then
		  Call KS.ShowTips("error","<li>你还没有登录或登录已过期，请重新<a href='../user/login/'>登录</a>!</li>")
		  Exit Sub
		End If
		If KS.C_S(ChannelID,6)<>1 Then Response.End()
		if KS.C_S(ChannelID,36)=0 then
		  Call KS.ShowTips("error","<li>本频道不允许投稿!</li>")
		  Exit Sub
		end if
		
		'设置缩略图参数
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")

		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>">我发布的<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=1">已审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=0">待审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=2">草 稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MyArticle.asp?ChannelID=<%=ChannelID%>&Status=3">被退稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
        </div>
		<div class="clear"></div>
		<%
		Select Case KS.S("Action")
		 Case "Del"	  Call KSUser.DelItemInfo(ChannelID,ComeUrl)
		 Case "Add","Edit"  Call DoAdd()
		 Case "DoSave" Call DoSave()
		 Case Else  Call ArticleList()
		End Select
	   End Sub
	   Sub ArticleList()
			  %>
			 <script language="javascript" src="../KS_Inc/showtitle.js"></script>
			  
			   <%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where A.Deltf=0 AND Inputer='"& KSUser.UserName &"'"
									Verific=KS.S("Status")
									If Verific="" or not isnumeric(Verific) Then Verific=4
                                    IF Verific<>4 Then Param= Param & " and Verific=" & Verific
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,FolderName from " & KS.C_S(ChannelID,2) &" a inner join KS_Class b On a.tid=b.id "& Param &" order by AddDate DESC"

								  Select Case Verific
								   Case 0 Call KSUser.InnerLocation("待审" & KS.C_S(ChannelID,3) & "列表")
								   Case 1 Call KSUser.InnerLocation("已审" & KS.C_S(ChannelID,3) & "列表")
								   Case 2 Call KSUser.InnerLocation("草稿" & KS.C_S(ChannelID,3) & "列表")
								   Case 3 Call KSUser.InnerLocation("退稿" & KS.C_S(ChannelID,3) & "列表")
                                   Case Else Call KSUser.InnerLocation("所有" & KS.C_S(ChannelID,3) & "列表")
								   End Select
								  %>
								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_myarticle.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">发布<%=KS.C_S(ChannelID,3)%></span></a></div>

				                     <table  width="100%"  border="0" align="center" cellpadding="1" cellspacing="1">
                                      <%
								 Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' colspan=2 height=30 valign=top>没有你要的" & KS.C_S(ChannelID,3) & "!</td></tr>"
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
			
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
							        Call showContent
				End If
     %>                      <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                  <form action="User_MyArticle.asp?ChannelID=<%=ChannelID%>" method="post" name="searchform">
                                  <td height="45" colspan=4 align="center">
										<%=KS.C_S(ChannelID,3)%>搜索：
										  <select name="Flag">
										   <option value="0">标题</option>
										   <option value="1">关键字</option>
									      </select>
										  
										  关键字
										  <input type="text" name="KeyWord" class="textbox" value="关键字" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" 搜 索 ">
							      </td>
								    </form>
                                </tr>
                        </table>
					</div>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I,PhotoUrl
    Response.Write "<FORM Action=""User_MyArticle.asp?ChannelID=" & ChannelID & "&Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
        If RS("PicNews")=1 Then
		 PhotoUrl=RS("PhotoUrl")
		Else
		 PhotoUrl="Images/nopic.gif"
		End If %>
                   <tr>
						 <td class="splittd" width="10"><input id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID"></td>
						 <td class="splittd" width="33"><div style="cursor:pointer;text-align:center;width:33px;height:33px;border:1px solid #f1f1f1;padding:1px;"><img  src="<%=PhotoUrl%>" title="<img src='<%=PhotoUrl%>' border=0 width='160'>" width="32" height="32"></div>
						 </td>
                        <td height="45" align="left" class="splittd">
						<div class="ContentTitle"><a href="../item/show.asp?m=<%=ChannelID%>&d=<%=rs("id")%>" target="_blank"><%=trim(RS("title"))%></a>
						</div>
						
						<div class="Contenttips">
			            <span>
						 栏目：[<%=RS("FolderName")%>] 发布人：<%=rs("Inputer")%> 发布时间：<%=KS.GetTimeFormat(rs("AddDate"))%>
						 状态：<%Select Case rs("Verific")
											   Case 0
											     Response.Write "<span style=""color:green"">待审</span>"
											   Case 1
											     Response.Write "<span>已审</span>"
                                               Case 2
											     Response.Write "<span style=""color:red"">草稿</span>"
											   Case 3
											     Response.Write "<span style=""color:blue"">退稿</span>"
                                              end select
											  %>
						 </span>
						</div>
						</td>
                        <td class="splittd" align="center">

											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>">修改</a> <a class='box' href="User_MyArticle.asp?channelid=<%=channelid%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))">删除</a>
											<%else
											 If KS.C_S(ChannelID,42)=0 Then
											  Response.write "---"
											 Else
											  Response.Write "<a  class='box' href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"'>修改</a> <a class='box' href='#' disabled>删除</a>"
											 End If
											end if%>

						</td>
                       </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr>
								  <td colspan=4 valign=top>
								   <table cellspacing="0" cellpadding="0" border="0" width="100%">
								    <tr>
									 <td>
								 <label><input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有</label><input class="button" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type=submit value=删除选定 name=submit1> </FORM>       
								     </td>
									 <td align='right'>
									 <%
							         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
						    	      %>
									 </td>
									 </tr>
									 </table>
									 
								  </td>
								</tr>
								<% 
  End Sub


  '添加文章
  Sub DoAdd()

        Call KSUser.InnerLocation("发布"& KS.C_S(ChannelID,3))
  		if KS.S("Action")="Edit" Then
		  Dim KS_A_RS_Obj:Set KS_A_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_A_RS_Obj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_A_RS_Obj.Eof Then
		     If KS.C_S(ChannelID,42) =0 And KS_A_RS_Obj("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
			   KS_A_RS_Obj.Close():Set KS_A_RS_Obj=Nothing
			   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			 End If
		     ClassID  = KS_A_RS_Obj("Tid")
			 Title    = KS_A_RS_Obj("Title")
			 KeyWords = KS_A_RS_Obj("KeyWords")
			 Author   = KS_A_RS_Obj("Author")
			 Origin   = KS_A_RS_Obj("Origin")
			 Content  = KS_A_RS_Obj("ArticleContent")
			 Verific  = KS_A_RS_Obj("Verific")
			 If Verific=3 Then Verific=0
			 PhotoUrl   = KS_A_RS_Obj("PhotoUrl")
			 Intro    = KS_A_RS_Obj("Intro")
			 FullTitle= KS_A_RS_Obj("FullTitle")
			 ReadPoint= KS_A_RS_Obj("ReadPoint")
			 Province = KS_A_RS_Obj("Province")
			 City     = KS_A_RS_Obj("City")
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
					  Dim UnitOption
					  If UserDefineFieldArr(11,I)="1" Then
					   UnitOption="@" & KS_A_RS_Obj(UserDefineFieldArr(0,I)&"_Unit")
					  Else
					   UnitOption=""
					  End If
					  
				  If i=0 Then
				    UserDefineFieldValueStr=KS_A_RS_Obj(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & KS_A_RS_Obj(UserDefineFieldArr(0,I)) & UnitOption & "||||"
				  End If
				Next
			  End If
		   End If
		   KS_A_RS_Obj.Close:Set KS_A_RS_Obj=Nothing
		   SelButton=KS.C_C(ClassID,1)
		Else
		  
		  Call KSUser.CheckMoney(ChannelID)
		 Author=KSUser.RealName
		 Origin=LFCls.GetSingleFieldValue("SELECT top 1 CompanyName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'")
		 ClassID=KS.S("ClassID")
		 If ClassID="" Then ClassID="0"
		 If ClassID="0" Then
		 SelButton="选择栏目..."
		 Else
		 SelButton=KS.C_C(ClassID,1)
		 End If
		 ReadPoint=0
		End If
		%>
		<script language = "JavaScript">
		    function CheckClassID()
			{
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
			}
			function insertHTMLToEditor(codeStr) 
			{ 
				oEditor=FCKeditorAPI.GetInstance("Content");
				if(oEditor   &&   oEditor.EditorWindow){ 
					oEditor.InsertHtml(codeStr); 
				} 
			} 
			function InsertFileFromUp(FileList,InstallDir)
			{ 
				Files=FileList.split("|");
				for(var i=0;i<Files.length-1;i++)
				{     var ext=getFilePic(Files[i]);
					  var files=Files[i].split('/');
					  var file=files[files.length-1];
					  var br='';
					  if (i!=Files.length-1) br='<br />';
					  var fileext = Files[i].substring(Files[i].lastIndexOf(".") + 1, Files[i].length).toLowerCase();
                      if (fileext=="gif" || fileext=="jpg" || fileext=="jpeg" || fileext=="bmp" || fileext=="png")
					  {
					   insertHTMLToEditor('<img src="'+Files[i]+'" border="0"/><br/>');	
					  }
					  else
					  {
					  var str="<img border=0 src="+InstallDir+"KS_Editor/images/FileIcon/"+ext+"> <a href='"+Files[i]+"'  target='_blank'>[点击浏览该文件:"+file+"]</a>"+br;
					  insertHTMLToEditor(str);	
					  }
				 }
			}
			
				function CheckForm()
				{
				<%Call KSUser.ShowUserFieldCheck(ChannelID)%>
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>标题！");
					document.myform.Title.focus();
					return false;
				  }	
				<%if F_B_Arr(9)=1 Then%> 
				 <%if KS.C_S(ChannelID,34)=0 Then%>
					if (frames["ArticleContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
					document.myform.Content.value=frames["ArticleContent"].KS_EditArea.document.body.innerHTML;
					if (document.myform.Content.value=='')
					{
						alert("请输入<%=KS.C_S(ChannelID,3)%>内容！");
						frames["ArticleContent"].KS_EditArea.focus();
						return false;
					}
				 <%else%>
				    if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=="")
					{
					  alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！");
					  FCKeditorAPI.GetInstance('Content').Focus();
					  return false;
					}
				 <%end if%>
				<%end if%>
				 return true;  
				}
				</script>
				<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
                  <form  action="User_MyArticle.asp?channelid=<%=channelid%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr  class="title">
					  <td colspan=2 align=center>
					       <%IF KS.S("Action")="Edit" Then
							   response.write "修改" & KS.C_S(ChannelID,3)
							   Else
							    response.write "发布" & KS.C_S(ChannelID,3)
							   End iF
							  %>
                         </td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
                       <td width="88%">　
					        <% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %>
					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
                              <td> 　
                                        <input class="textbox" name="Title" type="text" id="Title" style="width:250px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                        </tr>
						<%if F_B_Arr(2)=1 Then%>	  
                      <tr class="tdbg">
                           <td  height="25" align="center"><span><%=F_V_Arr(2)%>：</span></td>
                              <td> 　
                               <input class="textbox" name="FullTitle" type="text" style="width:250px; " value="<%=FullTitle%>" maxlength="100" /></td>
                        </tr>
						<%End If%>
						<%if F_B_Arr(5)=1 Then%>	  
                              <tr class="tdbg">
                                      <td height="25" align="center"><span><%=F_V_Arr(5)%>：</span></td>
                                <td>　
                                        <input name="KeyWords"  class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:250px; " />
                                  多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开								</td>
                              </tr>
					  <%end if%>
						<%if F_B_Arr(6)=1 Then%>	  
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
                                      <td height="25">　
                                        <input name="Author" class="textbox" type="text" id="Author" style="width:250px; " value="<%=Author%>" maxlength="30" /></td>
                              </tr>
						<%end if%>
						<%if F_B_Arr(7)=1 Then%>	  
                              <tr class="tdbg">
                                
                                      <td  height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
                                      <td>　
                                        <input class="textbox" name="Origin" type="text" id="Origin" style="width:250px; " value="<%=Origin%>" maxlength="100" /></td>
                              </tr>
						<%end if%>
						<%if F_B_Arr(23)="1" Then%>	  
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span><%=F_V_Arr(23)%>：</span></td>
                                      <td>　
                                        <script src="../plus/area.asp" type="text/javascript"></script>
									  <script language="javascript">
							  <%if Province<>"" then%>
							  $('#Province').val('<%=province%>');
								  <%end if%>
							  <%if City<>"" Then%>
							  $('#City')[0].options[1]=new Option('<%=City%>','<%=City%>');
							  $('#City')[0].options(1).selected=true;
							  <%end if%>
							</script>
									  </td>
                              </tr>
						<%end if%>
						
						
							  <%
							  Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
							  %>
						<%if F_B_Arr(8)=1 Then%>	  
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span><%=F_V_Arr(8)%>：</span><br><input name='AutoIntro' type='checkbox' checked value='1'><font color="#FF0000">自动截取内容的200个字作为导读</font></td>
                                      <td>　
                                        <textarea class='textbox' name="Intro" style='width:95%;height:80'><%=intro%></textarea></td>
                              </tr>
						<%end if%>

						<%if F_B_Arr(9)=1 Then%>	  
                              <tr class="tdbg">
                                  <td><%=F_V_Arr(9)%>:<br><img src="images/ico.gif" width="17" height="12" /><font color="#FF0000">如果<%=KS.C_S(ChannelID,3)%>较长可以使用分页标签：[NextPage]</font>
								  </td>
								  <td>
								
								<div align=center>
								<%
								
								If F_B_Arr(21)=1 and Cbool(LoginTF)=True Then
								%>
			      <table border='0' width='100%' cellspacing='0' cellpadding='0'>
			       <tr><td height='30' width=70>&nbsp;<strong><%=F_V_Arr(21)%>:</strong></td><td><iframe id='UpFileFrame' name='UpFileFrame' src='User_UpFile.asp?Type=File&ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='100%' height='100%'></iframe></td></tr>
			       </table>
		                         <%end if%>
								<textarea name="Content" style="display:none"><%=Server.HTMLEncode(Content)%></textarea>

				            <%
							If KS.C_S(ChannelID,34)=0 and Cbool(LoginTF)=True Then%>
                                <iframe id='ArticleContent' name='ArticleContent' src='Editor.asp?ID=Content&amp;style=0&amp;ChannelID=9998' frameborder="0" scrolling="No" width='98%' height='350'></iframe>                             <%else
								 Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""98%"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
							 end if%>  
								</td>
                            </tr>
					    <%end if%>
						<%if F_B_Arr(10)=1 Then%>	  
                             <tr class="tdbg">
                               <td height="25" align="center"><%=F_V_Arr(10)%>：</td>
                               <td height="25"><input name='PhotoUrl' type='text' id='PhotoUrl' value="<%=PhotoUrl%>" size='60'  class="textbox"/></td>
                             </tr>
						<%end if%>
						<%if F_B_Arr(11)=1 and Cbool(LoginTF)=True Then%>	  
                               <tr class="tdbg">
                                    <td height="25" align="center"><%=F_V_Arr(11)%>：</td>
                                    <td height="25"><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=<%=ChannelID%>' frameborder="0" scrolling="No" align="center" width='100%' height='30'></iframe></td>
                               </tr>
						<%end if%>
						
						<%If F_B_Arr(18)=1 Then%>
						<tr class="tdbg">
                                        <td height="25" align="center"><span>阅读<%=KS.Setting(45)%>：</span></td>
                                        <td height="25">
										 <input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> 如果免费阅读请输入“<font color=red>0</font>”
										  </td>
                       </tr>
					   <%end if%>
								<%if KS.S("Action")="Edit" And Verific=1 Then%>
								<input type="hidden" name="okverific" value="1">
								<input type="hidden" name="verific" value="1">
								<%else%>
						<tr class="tdbg" >
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>状态：</span></td>
                                        <td height="25">
										 <input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked"%> />
                                          投搞
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked"%>/>
                                          草稿
										  </td>
                                      </tr>
							  <%end if%>
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					   <input class="button" type="submit" name="Submit" value="OK, 保 存 " />
                            　
                       <input class="button" type="reset" name="Submit2" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
				<br/><br/><br/>
		  <%
  End Sub
  
  Sub DoSave()
                 ClassID=KS.S("ClassID")
				 Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				 KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				 Author=KS.LoseHtml(KS.S("Author"))
				 Origin=KS.LoseHtml(KS.S("Origin"))
				 Content = Request.Form("Content")
				 Content=KS.FilterIllegalChar(KS.ClearBadChr(content))
				 
				 if KS.IsNul(Content) Then Content="&nbsp;"
				 Verific=KS.ChkClng(KS.S("Status"))
				 Intro  = KS.FilterIllegalChar(KS.LoseHtml(KS.S("Intro")))
				 Province= KS.LoseHtml(KS.S("Province"))
				 City    = KS.LoseHtml(KS.S("City"))
				 FullTitle = KS.LoseHtml(KS.S("FullTitle"))
				 if Intro="" And KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Request.Form("Content")),200)
				 
				 Dim Fname,FnameType,TemplateID,WapTemplateID
				 If KS.ChkClng(KS.S("ID"))=0 Then
					 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
					 RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
					 if RSC.Eof Then 
					  Response.end
					 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 WapTemplateID=RSC("WapTemplateID")
					 End If
					 RSC.Close:Set RSC=Nothing
				 End If
				 If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 and verific=1  Then
					 If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
				 if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
				 
				 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				 
				 PhotoUrl=KS.S("PhotoUrl")
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				 
				Next
				End If
				 
				  
				  if ClassID="" Then
				    Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "标题!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Content="" and F_B_Arr(9)=1 Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "内容!');history.back();</script>"
				    Exit Sub
				  End IF
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where Inputer='" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("Adddate")=Now
				  RSObj("Rank")="★★★"
				  RSObj("Inputer")=KSUser.UserName
				 End If
				  RSObj("Title")=Title
				  RSObj("FullTitle")=FullTitle
				  RSObj("Tid")=ClassID
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("ArticleContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("Intro")=Intro
				  RSObj("DelTF")=0
				  RSObj("Comment")=1
                  If F_B_Arr(18)=1 Then
				  RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
				  RSObj("Province")=Province
				  RSObj("City")=City				  
				  if PhotoUrl<>"" Then 
				   RSObj("PicNews")=1
				  Else
				   RSObj("PicNews")=0
				  End if
				  If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RSObj("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
							else
							 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
							end if
							If UserDefineFieldArr(11,I)="1"  Then
							RSObj("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
							End If
							
						Next
				  End If
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				Fname=RSOBj("Fname")
				
				If Verific=1 Then 
				    Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
					If KS.C_S(ChannelID,17)=2  and (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
					 Dim KSRObj:Set KSRObj=New Refresh
					 Dim DocXML:Set DocXML=KS.RsToXml(RSObj,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					  Call KSRObj.RefreshContent()
					  Set KSRobj=Nothing
					End If
				End If
				 RSObj.Close:Set RSObj=Nothing
				 
               If KS.ChkClng(KS.S("ID"))=0 Then
			     Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Intro,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
                 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,0)
			     Call KSUser.AddLog(KSUser.UserName,"在栏目[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]发表了" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",1)
				 KS.Echo "<script>if (confirm('" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?')){location.href='User_myArticle.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyArticle.asp?ChannelID=" & ChannelID & "';}</script>"
			   Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Intro,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,Content & PhotoUrl ,1)
			     Call KSUser.AddLog(KSUser.UserName,"对" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""做了修改!",1)
				 KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='User_MyArticle.asp?channelid=" & channelid & "';</script>"
			   End If
  End Sub
  
End Class
%> 

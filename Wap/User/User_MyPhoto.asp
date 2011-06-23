<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Editor/FCKeditor/fckeditor.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_MyPhoto
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyPhoto
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private ComeUrl,SelButton
		Private F_B_Arr,F_V_Arr,ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,PicUrls,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
		Private Sub Class_Initialize()
			MaxPerPage =15
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
		  Response.Write "<script>top.location.href='Login.asp';</script>"
		  Exit Sub
		End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=2
		If KS.C_S(ChannelID,6)<>2 Then Response.End()
		if conn.execute("select usertf from ks_channel where channelid=" & channelid)(0)=0 then
		  Response.Write "<script>alert('本频道关闭投稿!');window.close();</script>"
		  Exit Sub
		end if

		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
        F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		Call KSUser.Head()
		%>
			<table width="98%" height=20 border=0 align="center" cellPadding=0 cellSpacing=0 borderColor=#111111 id=AutoNumber3 style="BORDER-COLLAPSE: collapse">
                   <tr>
                     <td  width=23 height=20><img src="Images/z3.gif" border=0></td>
                     <td  width=160 bgColor=#5298d1 height=20><B>&nbsp;<font color=#ffffff><SPAN style="FONT-SIZE: 10.5pt">我发布的<%=KS.C_S(ChannelID,3)%></SPAN></font></B></td>
                     <td width=12 height=20><img src="Images/z4.gif" border=0></td>
                     <td width=583 height=20 align=right><a href="user_MyPhoto.asp?Action=Add&Channelid=<%=ChannelID%>"><font color=red>·发布<%=KS.C_S(ChannelID,3)%></font></a>&nbsp;&nbsp;·<a href="User_MyPhoto.asp?Status=2&Channelid=<%=ChannelID%>">草 稿[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>&nbsp;&nbsp;·<a href="User_MyPhoto.asp?Status=0&Channelid=<%=ChannelID%>">待审核[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>&nbsp;&nbsp;·<a href="User_MyPhoto.asp?Status=1&Channelid=<%=ChannelID%>">&nbsp;已审核[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%>]</a>&nbsp;&nbsp;·<a href="User_MyPhoto.asp?Status=3&Channelid=<%=ChannelID%>">被退稿[<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%>]</a></td>
                   </tr>
                 </table>
		<%
		Select Case KS.S("Action")
		  Case "Del"
		   Call PhotoDel()
		  Case "Add","Edit"
		   Call DoAdd()
		  Case "DoSave"
		   Call DoSave()
		  Case Else
		   Call PhotoList()
		End Select
	   End Sub
	   
	   Sub PhotoList()
			  %>
			   <SCRIPT language=javascript src="../KS_Inc/showtitle.js"></script>
			   <SCRIPT language=javascript>
				function unselectall()
				{
					if(document.myform.chkAll.checked)
					{
				 document.myform.chkAll.checked = document.myform.chkAll.checked&0;
					}
				}
				function CheckAll(form)
				{
				  for (var i=0;i<form.elements.length;i++)
				  {
					var e = form.elements[i];
					if (e.Name != 'chkAll'&&e.disabled==false)
					   e.checked = form.chkAll.checked;
					}
				  }
               </SCRIPT>
			   <%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where Inputer='"& KSUser.UserName &"'"
                                    Verific=KS.S("Status")
									If Verific="" or not isnumeric(Verific) Then Verific=4
                                    IF Verific<>4 Then 
									   Param= Param & " and Verific=" & Verific
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,b.foldername from " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id "& Param &" order by AddDate DESC"

			  					  Select Case Verific
								   Case 0 
								    Call KSUser.InnerLocation("待审" & KS.C_S(ChannelID,3) & "列表")
								   Case 1
								    Call KSUser.InnerLocation("已审" & KS.C_S(ChannelID,3) & "列表")
								   Case 2
								   Call KSUser.InnerLocation("草稿" & KS.C_S(ChannelID,3) & "列表")
								   Case 3
								   Call KSUser.InnerLocation("退稿" & KS.C_S(ChannelID,3) & "列表")
                                   Case Else
								    Call KSUser.InnerLocation("所有" & KS.C_S(ChannelID,3) & "列表")
								   End Select
 %>
              <table width="98%" border="0" cellspacing="1" cellpadding="1" class="border" align="center">
                       <tr class="Title">
                           <td width="5%" height="22" align="center">选中</td>
                           <td width="34%" height="22" align="center"><%=F_V_Arr(0)%></td>
                           <td width="10%" height="22" align="center"><%=KS.C_S(ChannelID,3)%>录入</td>
                           <td width="18%" height="22" align="center"><%=F_V_Arr(10)%></td>
                           <td width="10%" height="22" align="center">状态</td>
                           <td width="22%" height="22" align="center">管理操作</td>
                       </tr>
                             <%
									Set RS=KS.InitialObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td colspan=6 height=30 class='tdbg' valign=top>没有你要的" & KS.C_S(ChannelID,3) & "!</td></tr>"
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
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
				End If
     %>
                             <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                    <form action="User_MyPhoto.asp?channelid=<%=channelid%>" method="post" name="searchform" id="searchform">
                              <td colspan=6 height="45" align="center">
                                     <%=KS.C_S(ChannelID,3)%>搜索：
                                         <select name="Flag">
                                             <option value="0"><%=F_V_Arr(0)%></option>
                                             <option value="1"><%=F_V_Arr(6)%></option>
                                           </select>
                                           
                                         关键字
                                         <input type="text" name="KeyWord" class="textbox" value="关键字" size="20" />
                                         &nbsp;
                                         <input class="button" type="submit" name="submit12" value=" 搜 索 " />
							      </td>
                                    </form>
                                </tr>
                        </table>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I
    Response.Write "<FORM Action=""User_MyPhoto.asp?ChannelID=" & ChannelID & "&Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                         <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td height="22" align="center"><INPUT id="ID" onClick="unselectall()" type="checkbox" value="<%=RS("ID")%>"  name="ID"></td>
                                            <td align="left">[<%=RS("FolderName")%>]<a title="<table width=80 border=0 align=center><tr><td><img src='<%=RS("PhotoUrl")%>' border=0 width='130' height='80'></td></tr></table>"  href="../<%=KS.C_S(ChannelID,10)%>/Show.asp?id=<%=rs("id")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("title")),25)%></a></td>
											<td align="center"><%=rs("Inputer")%></td>
                                            <td align="center"><%=formatdatetime(rs("AddDate"),2)%></td>
                                            <td align="center">
											  <%Select Case rs("Verific")
											   Case 0
											     Response.Write "<span class=""font10"">待审</span>"
											   Case 1
											     Response.Write "<span class=""font11"">已审</span>"
                                               Case 2
											     Response.Write "<span class=""font13"">草稿</span>"
											   Case 3
											     Response.Write "<span class=""font14"">退稿</span>"
                                              end select
											  %></td>
                                            <td height="22" align="center">
											<%if rs("Verific")<>1 then%>
											<a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=Edit&id=<%=rs("id")%>&page=<%=CurrentPage%>" class="link3">修改</a> <a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))" class="link3">删除</a>
											<%else
												 If KS.C_S(ChannelID,42)=0 Then
												  Response.write "---"
												 Else
												  Response.Write "<a href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"' class='link3'>修改</a>"
												 End If
											end if%>
											</td>
                                          </tr>
					   <tr><td colspan=6 background='images/line.gif'></td></tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>                    <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
						<td valign=top colspan="6">&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有<INPUT class="button" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type=submit value="删除选定的<%=KS.C_S(ChannelID,3)%>" name="submit1"> </FORM> &nbsp;&nbsp;         
								<%Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_MyPhoto.asp", True, KS.C_S(ChannelID,4) & KS.C_S(ChannelID,3), CurrentPage, "ChannelID=" & ChannelID &"&Status=" & Verific)
%></td>
								 
					 </tr>
								<%
  End Sub
  '删除图片
  Sub PhotoDel
	 Dim ID:ID=KS.S("ID")
	 ID=KS.FilterIDs(ID)
	 If ID="" Then Call KS.Alert("你没有选中要删除的" & KS.C_S(ChannelID,3) & "!",ComeUrl):Response.End
	 Conn.Execute("Delete From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName & "' and verific<>1 and  ID In(" & ID & ")")
	 if ComeUrl="" then
	Response.Redirect("../index.asp")
	else
	Response.Redirect ComeUrl
	end if

 End Sub
 

 '添加图片
 Sub DoAdd()
 		Call KSUser.InnerLocation("发布" & KS.C_S(ChannelID,3) & "")
		if KS.S("Action")="Edit" Then
		  Dim KS_P_RS_Obj:Set KS_P_RS_Obj=KS.InitialObject("ADODB.RECORDSET")
		   KS_P_RS_Obj.Open "Select * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_P_RS_Obj.Eof Then
		     If KS.C_S(ChannelID,42) =0 And KS_P_RS_Obj("Verific")=1 Then
			   KS_P_RS_Obj.Close():Set KS_P_RS_Obj=Nothing
			   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!")
			 End If
		     ClassID  = KS_P_RS_Obj("Tid")
			 Title    = KS_P_RS_Obj("Title")
			 KeyWords = KS_P_RS_Obj("KeyWords")
			 Author   = KS_P_RS_Obj("Author")
			 Origin   = KS_P_RS_Obj("Origin")
			 Content  = KS_P_RS_Obj("PictureContent")
			 Verific  = KS_P_RS_Obj("Verific")
			 If Verific=3 Then Verific=0
			 PicUrls  = KS_P_RS_Obj("PicUrls")
			 PhotoUrl = KS_P_RS_Obj("PhotoUrl")
			 '自定义字段
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				  If UserDefineFieldValueStr="" Then
				    UserDefineFieldValueStr=KS_P_RS_Obj(UserDefineFieldArr(0,I)) & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & KS_P_RS_Obj(UserDefineFieldArr(0,I)) & "||||"
				  End If
				Next
			  End If
		   End If
		   KS_P_RS_Obj.Close:Set KS_P_RS_Obj=Nothing
		   Selbutton=KS.C_C(ClassID,1)
		Else
		  ClassID=KS.S("ClassID"):Author=KSUser.RealName:PicUrls=""
		  If ClassID="" Then ClassID="0"
		  If ClassID="0" Then
		  SelButton="选择栏目..."
		  Else
		  SelButton=KS.C_C(ClassID,1)
		  End If
		End If
			  %>
			  <script language='JavaScript' src='../KS_Inc/Prototype.js'></script>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform">
				           <tr class="title">
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
                                        <td height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
                                        <td>　
							 <% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %>
										</td>
                             </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
                                        <td> 　 
                                          <input name="Title" class="textbox" type="text" id="Title" value="<%=Title%>" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
								<%If F_B_Arr(6)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
                                        <td>　
                                          <input name="KeyWords" class="textbox" type="text" id="KeyWords" value="<%=KeyWords%>" style="width:250px; " /> 
                                          多个关键字请用&quot;<span style="color: #FF0000">|</span>&quot;隔开</td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(7)=1 Then%>
                                <tr class="tdbg">

                                        <td height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
                                        <td height="25">　
                                          <input class="textbox" name="Author" type="text" id="Author" value="<%=Author%>" style="width:250px; " maxlength="30" /></td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(8)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(8)%>：</span></td>
                                        <td>　
                                          <input class="textbox" name="Origin" type="text" id="Origin" value="<%=Origin%>" style="width:250px; " maxlength="100" /></td>
							  </tr>
							  <%End if%>
								<%
							  Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
							  %>
							   
							  <tr class="tdbg">
                                    <td height="40" align="center"><span><%=F_V_Arr(3)%>：</span></td>
                                    <td>&nbsp;&nbsp;&nbsp;<input name='picnum' class='textbox' type='text' id='picnum' size='4' value='4' style='text-align:center'>&nbsp;<input name='kkkup' type='button' id='kkkup2' value='设定' onClick="MakeUpload($F('picnum'),'click');" class='button'>注：最多<font color='red'>99</font><%=KS.C_S(ChannelID,4)%>，远程<%=KS.C_S(ChannelID,3)%>地址必须以<font color='red'>http://</font>开头<input type='hidden' name='PicUrls'> </td>
                              </tr>
								<tr class="tdbg">
                                   <td height="30" align="center"><span><%=F_V_Arr(4)%>：</span></td>
                                   <td align="center">
								   <span id='uploadfield'></span>
								   <%
								If KS.S("Action")<>"Edit" then
							    Response.Write "	<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=" & ChannelID & "' frameborder=0 scrolling=no width='100%' height='110'></iframe>"
								End If
								   %>
								   </td>
							  </tr>
							  
								<tr class="tdbg">
                                        <td height="35" align="center"><span><%=F_V_Arr(2)%>：</span></td>
                                        <td>　
                                          <input class='textbox' name='PhotoUrl' type='text' style="width:250px;" value="<%=PhotoUrl%>" id='PhotoUrl' maxlength="100" />
                                          <font color='#FF0000'>*</font>&nbsp;
                                          <input class="button" type='button' name='Submit3' value='选择<%=KS.C_S(ChannelID,3)%>地址...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=选择<%=KS.C_S(ChannelID,3)%>&ChannelID=<%=ChannelID%>',500,360,window,document.myform.PhotoUrl);" /></td>
							   </tr>
								<%If F_B_Arr(9)=1 Then%>
							   <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(9)%>：<br /></td>
                                        <td align="center"><%If KS.C_S(ChannelID,34)=0 Then%>	
                                       <textareastyle="display:none;" name="Content"><%=KS.HTMLCode(Content)%></textarea>
                                       <iframe id='PhotoContent' name='PhotoContent' src='Editor.asp?ID=Content&style=0&ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='95%' height='200'></iframe>
									<%else
										 Dim oFCKeditor 
										 Set oFCKeditor = New FCKeditor 
										 oFCKeditor.BasePath = "../KS_Editor/FCKeditor/"
										 oFCKeditor.ToolbarSet = "Basic" 
										 oFCKeditor.Width = "98%" 
										 oFCKeditor.Height = "150" 
										 oFCKeditor.Value = KS.HTMLCode(Content)
										 oFCKeditor.Create "content" 
								end if%>     
									   
									   </td>
                                </tr>
                                <%end if%>
								<tr class="tdbg" <%if KS.S("Action")="Edit" And Verific=1 Then response.write " style='display:none'"%>>
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>状态：</span></td>
                                        <td><input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked"%> />
投搞
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked"%>/>
草稿</td>
							  </tr>
                               <tr class="tdbg">
                            <td align="center" colspan=2>
							<input class="button" type="button" onClick="CheckForm()" name="Submit" value=" OK,保存 " />
                            <input class="button" type="reset"  name="Submit2" value=" 重 来 " /></td>
                              </tr>
                  </form>
</table>

			<%
			 Dim picnum
		     If KS.S("Action")="Edit" Then
					picnum = UBound(split(PicUrls,"|||"))+1
			 Else
			        picnum=4
			 End If
			%>
			 <script>
			 <%if Action<>"Edit" Then%>
			 var LastNum=1;
			 <%else%>
			 var LastNum=$('picnum');
			 <%end if%>
			 var tempup='';
			 var picnum=<%=Picnum%>;
		 	 function document.onreadystatechange()
			  {   
				 MakeUpload(<%=Picnum%>);
				 IniPicUrl();
				 tempup=$("uploadfield").innerHTML;
			  }
			function IniPicUrl()
			{
			 var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			   for ( var i=1 ;i<PicUrlArr.length+1;i++)
			   { 
				 document.getElementById('thumb'+i).value=PicUrlArr[i-1].split('|')[2];
				 document.getElementById('imgurl'+i).value=PicUrlArr[i-1].split('|')[1];
				 document.getElementById('imgnote'+i).value=PicUrlArr[i-1].split('|')[0];
				 document.getElementById('picview'+i).src=PicUrlArr[i-1].split('|')[1];
			   }
			    $('picnum').value=i-1;
			    LastNum=i;
			   }
			}
			function MakeUpload(mnum,str)
			{ 
			   if (parseInt(mnum)>=100){
			   alert('最多只能同上传99张!');
			   return false;}
			   var startNum=1;
			   var endNum = mnum;
			   var fhtml = "";
			   
			   if (str=='click') startNum=LastNum;
			   
			   for(startNum;startNum <= endNum;startNum++){
				   fhtml += "<table width=\"99%\" style='margin:2px' class='border' align=center border=\"0\" id=\"seltb"+startNum+"\" cellpadding=\"3\" cellspacing=\"1\">";
				   fhtml += "<tr class='tdbg'> "
				   fhtml +="  <td height=\"25\" width=18 align=center class=clefttitle rowspan=\"3\"><strong>第"+startNum+"张</strong></td>";
				   fhtml += " <td width=\"124\" rowspan=\"3\" align=\"center\"><img src=\"images/view.gif\" width=\"120\" height=\"80\" border=1 id=\"picview"+startNum+"\" name=\"picview"+startNum+"\"></td>";
				   fhtml += "</tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"25\"> 　小图地址： ";
				   fhtml += "<input type=\"text\" class='textbox' onblur='view("+startNum+");' name='thumb"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取<%=KS.C_S(ChannelID,3)%>\" onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=选择<%=KS.C_S(ChannelID,3)%>&ChannelID=<%=ChannelID%>',550,290,window,document.myform.thumb"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;";
				   fhtml += "<br>　大图地址： <input type=\"text\" class='textbox' onblur='view("+startNum+");' name='imgurl"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取<%=KS.C_S(ChannelID,3)%>\" onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=选择<%=KS.C_S(ChannelID,3)%>&ChannelID=<%=ChannelID%>',550,290,window,document.myform.imgurl"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;";
				   <%If KS.S("Action")="Edit" Then%>
					if (startNum>picnum)
					{
				   fhtml += "<br><iframe id='UpPhotoFrame"+startNum+"' name='UpPhotoFrame"+startNum+"' src='User_UpFile.asp?ChannelID=<%=ChannelID%>&type=Single&objid="+startNum+"' frameborder=0 scrolling=no width='100%' height='22'></iframe>"
				   }
				   <%end if%>

				   fhtml += "</td></tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"30\">　<%=KS.C_S(ChannelID,3)%>简介： ";
				   fhtml += "<textarea class='textbox' name='imgnote"+startNum+"' style=\"height:46px;width:350px\"></textarea> </td>";
				   fhtml += "</tr></table>\r\n";
			  }
			  <%If KS.S("Action")="Edit" Then%>
			  //LastNum=Number(endNum)+1;
			  $("uploadfield").innerHTML = tempup+fhtml;
			  <%Else%>
			  $("uploadfield").innerHTML = fhtml;
			  frames['UpPhotoFrame'].ChooseOption(mnum);
			  $('UpPhotoFrame').height=80+26*(mnum/2); 
			  <%End If%>
			  parent.init();
			  parent.resize_mainframe();
			}
			 function view(num)
			 {
			  if (document.getElementById("thumb"+num).value!='')
			  document.getElementById("picview"+num).src=document.getElementById("thumb"+num).value;
			  else if(document.getElementById("imgurl"+num).value!='')
			  document.getElementById("picview"+num).src=document.getElementById("imgurl"+num).value;
			 }

			 function SetPicUrlByUpLoad(DefaultThumb,PicUrlStr,ThumbPathFileName)
			{  var UrlStrArr;
			   UrlStrArr=PicUrlStr.split('|');
			   for (var i=1;i<UrlStrArr.length;i++)
			   {
			   var url=UrlStrArr[i-1]; 
			   if(url!=null&&url!=''){
				 document.getElementById('imgurl'+i).value=url;
			   } 
			  }
			  var ThumbsArr=ThumbPathFileName.split("|")
			  for(var i=1;i<ThumbsArr.length;i++)
			  {
			   var url=ThumbsArr[i-1]; 
			   if(url!=null&&url!=''){
				 document.getElementById('thumb'+i).value=url;
			   } 
			  }
			  
			 // if (ThumbPathFileName!='')
			 // {
				 $('PhotoUrl').value=ThumbsArr[DefaultThumb-1];
			  //}
			}
			
				 function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj)
				{
				 var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
					if (ReturnStr!='') SetObj.value=ReturnStr;
				}
				function CheckForm()
				{
				<%If KS.C_S(ChannelID,34)=0 and F_B_Arr(9)=1 Then%>	
				if (frames["PhotoContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
                document.myform.Content.value=frames["PhotoContent"].KS_EditArea.document.body.innerHTML;
				<%End If%>
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					//document.myform.ClassID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.PhotoUrl.value=='')
					{
					alert("请输入<%=KS.C_S(ChannelID,3)%>缩略图！");
					document.myform.PhotoUrl.focus();
					return false;
					}
					
				$('PicUrls').value='';
				for(var i=1;i<=$F("picnum");i++){
				  if (document.getElementById('imgurl'+i).value!=''&&document.getElementById('imgurl'+i).value!='del') 
				   {
				   var note=document.getElementById('imgnote'+i).value;
				   note=note.replace('|||','');
				   spic=document.getElementById('imgurl'+i).value;
				   tpic=document.getElementById('thumb'+i).value;
				   if (tpic=='') tpic=spic;
				   //if (spic.substring(0,4).toLowerCase()=='http'&&$("BeyondSavePic").checked==true)
				  // {
					// $('LayerPrompt').style.display='';
					// window.setInterval('ShowPromptMessage()',150)
				  // }
				   if ($F('PicUrls')=='')                 
				   $('PicUrls').value=note+'|'+spic+'|'+tpic;
				   else 
				   $('PicUrls').value+='|||'+note+'|'+spic+'|'+tpic;
				   }
				}
				if ($F('PicUrls')=='')
				{
				  alert('请输入<%=KS.C_S(ChannelID,3)%>内容!');
				  Field.focus('imgurl1');
				  return false;
				}
                    $('myform').submit();  
				}
				function CheckClassID()
				{
				if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					//document.myform.ClassID.focus();
					return false;
				  }		
				  return true;
				}
			</script>
			 <%
  End Sub
  
  Sub DoSave()
  				Dim ClassID:ClassID=KS.S("ClassID")
				Dim Title:Title=KS.LoseHtml(KS.S("Title"))
				Dim KeyWords:KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				Dim Author:Author=KS.LoseHtml(KS.S("Author"))
				Dim Origin:Origin=KS.LoseHtml(KS.S("Origin"))
				Dim Content
				Content = Request.Form("Content")
				 Content=KS.CheckScript(KS.HtmlCode(content))
				 Content=KS.HtmlEncode(Content)
				Dim Verific:Verific=KS.ChkClng(KS.S("Status"))
				Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
				Dim PicUrls:PicUrls=KS.S("PicUrls")
				 If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 Then
				  If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If

				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If
				  Dim RSObj
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');history.back();</script>"
				    Exit Sub
				  End IF
	              If PhotoUrl="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "缩略图!');history.back();</script>"
				    Exit Sub
				  End IF
	              If PicUrls="" Then
				    Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');history.back();</script>"
				    Exit Sub
				  End IF
				If KS.ChkClng(KS.S("ID"))=0 Then
				 Dim Fname,FnameType,TemplateID
				 Dim RSC:Set RSC=KS.InitialObject("ADODB.RECORDSET")
				 RSC.Open "select TemplateID,FnameType,FsoType from KS_Class Where ID='" & ClassID & "'",conn,1,1
				 if RSC.Eof Then 
				  Response.end
				 Else
				 FnameType=RSC("FnameType")
				 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
				 TemplateID=RSC("TemplateID")
				 End If
				 RSC.Close:Set RSC=Nothing
			    End If
				  
				Set RSObj=KS.InitialObject("Adodb.Recordset")
				RSObj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Inputer")=KSUser.UserName
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("Fname")=FName
				  RSObj("AddDate")=Now
				End If
				  RSObj("Title")=Title
				  RSObj("Tid")=ClassID
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("PicUrls")=PicUrls
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Origin")=Origin
				  RSObj("PictureContent")=Content
				  RSObj("Verific")=Verific
				  RSObj("DelTF")=0
				  RSObj("Comment")=1
				  If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RSObj("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
							else
							 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
							end if
						Next
				  End If
				RSObj.Update
				If Left(Ucase(Fname),2)="ID" And KS.ChkClng(KS.S("ID"))=0 Then
				       RSObj.MoveLast
					   RSObj("Fname") = RSObj("ID") & FnameType
					   RSObj.Update
				 End If
				If KS.C_S(ChannelID,17)=2  and KS.C_S(Channelid,7)=1 Then
				 Dim KSRObj:Set KSRObj=New Refresh
				 Call KSRObj.RefreshPictureContent(RSObj,ChannelID)
				 Set KSRobj=Nothing
                End If
				 RSObj.Close:Set RSObj=Nothing
				 If KS.ChkClng(KS.S("ID"))=0 Then
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?')){location.href='User_MYPhoto.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';}</script>"
				Else
				 Response.Write "<script>alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';</script>"
				End If
  End Sub
  

End Class
%> 

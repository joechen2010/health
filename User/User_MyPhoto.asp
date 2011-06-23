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
Set KSCls = New Admin_MyPhoto
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyPhoto
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private ComeUrl,SelButton,ReadPoint
		Private F_B_Arr,F_V_Arr,ClassID,Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,PicUrls,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
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
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=2
		If KS.C_S(ChannelID,6)<>2 Then Response.End()
		if conn.execute("select usertf from ks_channel where channelid=" & channelid)(0)=0 then
		  Response.Write "<script>alert('本频道关闭投稿!');window.close();</script>"
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
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>">我发布的<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=1">已审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=0">待审核(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=2">草 稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Status=3">被退稿(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
         </div>
		<%
		Select Case KS.S("Action")
		  Case "Del"
		   Call KSUser.DelItemInfo(ChannelID,ComeUrl)
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
 								  <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_myphoto.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">发布<%=KS.C_S(ChannelID,3)%></span></a></div>

              <table width="98%" border="0" cellspacing="1" cellpadding="1"  align="center">
                             <%
								Set RS=Server.CreateObject("AdodB.Recordset")
								RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td colspan=4 height=30 align='center' valign=top>没有你要的" & KS.C_S(ChannelID,3) & "!</td></tr>"
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
                             <tr>
                           <form action="User_MyPhoto.asp?channelid=<%=channelid%>" method="post" name="searchform" id="searchform">
                              <td colspan=4 height="45" align="center">
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
           <tr>
		     <td class="splittd" width="10"><INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID"></td>
             <td class="splittd" width="40" align="center"><div style="cursor:pointer;width:33px;height:33px;border:1px solid #f1f1f1;padding:1px"><img  src="<%=RS("PhotoUrl")%>" width="32" height="32" title="<img src='<%=RS("PhotoUrl")%>' border=0 width='160'>"></div>
			 </td>
              <td align="left" class="splittd">
			  <div class="ContentTitle"><a href="../item/show.asp?m=<%=ChannelID%>&d=<%=rs("id")%>" target="_blank"><%=trim(RS("title"))%></a></div>
			  
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
              <td align="center" class="splittd">
											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=Edit&id=<%=rs("id")%>&page=<%=CurrentPage%>" class="box">修改</a> <a href="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('确定删除<%=KS.C_S(ChannelID,3)%>吗?'))" class="box">删除</a>
											<%else
												 If KS.C_S(ChannelID,42)=0 Then
												  Response.write "---"
												 Else
												  Response.Write "<a class='box' href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"'>修改</a> <a href='#' class='box' disabled>删除</a>"
												 End If
											end if%>
											</td>
                                          </tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>                    <tr>
						<td valign=top colspan="4">
						<table border="0" width="100%">
								    <tr>
									 <td>
									 <label><INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;选中所有</label><INPUT class="button" onClick="return(confirm('确定删除选中的<%=KS.C_S(ChannelID,3)%>吗?'));" type=submit value="删除选定" name="submit1"> </FORM> 
						              </td>
									  <td align="right">        
								<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
                                      </td>
									 </tr>
						</table>
                      </td>
								 
					 </tr>
								<%
  End Sub

 '添加图片
 Sub DoAdd()
 		Call KSUser.InnerLocation("发布" & KS.C_S(ChannelID,3) & "")
		if KS.S("Action")="Edit" Then
		  Dim KS_P_RS_Obj:Set KS_P_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		   KS_P_RS_Obj.Open "Select * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName &"' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not KS_P_RS_Obj.Eof Then
		     If KS.C_S(ChannelID,42) =0 And KS_P_RS_Obj("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
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
			 ReadPoint= KS_P_RS_Obj("ReadPoint")
			 If Verific=3 Then Verific=0
			 PicUrls  = KS_P_RS_Obj("PicUrls")
			 PhotoUrl = KS_P_RS_Obj("PhotoUrl")
			 '自定义字段
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
					  Dim UnitOption
					  If UserDefineFieldArr(11,I)="1" Then
					   UnitOption="@" & KS_A_RS_Obj(UserDefineFieldArr(0,I)&"_Unit")
					  Else
					   UnitOption=""
					  End If
				  If UserDefineFieldValueStr="" Then
				    UserDefineFieldValueStr=KS_P_RS_Obj(UserDefineFieldArr(0,I)) &UnitOption& "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr & KS_P_RS_Obj(UserDefineFieldArr(0,I)) &UnitOption & "||||"
				  End If
				Next
			  End If
		   End If
		   KS_P_RS_Obj.Close:Set KS_P_RS_Obj=Nothing
		   Selbutton=KS.C_C(ClassID,1)
		Else
		  Call KSUser.CheckMoney(ChannelID)
		  ClassID=KS.S("ClassID"):Author=KSUser.RealName:PicUrls=""
		  If ClassID="" Then ClassID="0"
		  If ClassID="0" Then
		  SelButton="选择栏目..."
		  Else
		  SelButton=KS.C_C(ClassID,1)
		  End If
		  ReadPoint=0
		End If
		If KS.IsNul(Content) Then Content=" "
			  %>
                  <form id="myform" action="User_MyPhoto.asp?ChannelID=<%=ChannelID%>&Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform">
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
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
					        
				<% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %></td>
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
                                    多个关键字请用英文逗号&quot;<span style="color: #FF0000">,</span>&quot;隔开</td>
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
							  <%
							 Dim picnum
							 If KS.S("Action")="Edit" Then
									picnum = UBound(split(PicUrls,"|||"))+1
							 Else
									picnum=4
							 End If
							%> 
							  <tr class="tdbg">
                                    <td height="40" align="center"><span><%=F_V_Arr(3)%>：</span></td>
                                    <td>&nbsp;&nbsp;&nbsp;<input name='picnum' class='textbox' type='text' id='picnum' size='4' value='<%=PicNum%>' style='text-align:center'>&nbsp;<input name='kkkup' type='button' id='kkkup2' value='设定' onClick="MakeUpload($F('picnum'),'click');" class='button'>注：最多<font color='red'>99</font><%=KS.C_S(ChannelID,4)%>，远程<%=KS.C_S(ChannelID,3)%>地址必须以<font color='red'>http://</font>开头<input type='hidden' id='PicUrls' name='PicUrls'> </td>
                              </tr>
								<tr class="tdbg">
                                   <td height="30" align="center"><span><%=F_V_Arr(4)%>：</span></td>
                                   <td align="center">
								   <span id='uploadfield'>sssss</span>
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
                                          <input class="button" type='button' name='Submit3' value='选择<%=KS.C_S(ChannelID,3)%>地址...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=<%=ChannelID%>',500,360,window,document.myform.PhotoUrl);" /></td>
							   </tr>
								<%If F_B_Arr(9)=1 Then%>
							   <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(9)%>：<br /></td>
                                        <td align="center"><%If KS.C_S(ChannelID,34)=0 Then%>	
                                       <textarea style="display:none;" name="Content"><%=Server.HTMLEncode(Content)%></textarea>
                                       <iframe id='PhotoContent' name='PhotoContent' src='Editor.asp?ID=Content&style=0&ChannelID=<%=ChannelID%>' frameborder=0 scrolling=no width='95%' height='200'></iframe>
									<%else
									     Response.Write "<textarea name=""Content"" style=""display:none"">" & Server.HTMLEncode(Content) & "</textarea>"
								        Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""98%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
								end if%>     
									   
									   </td>
                                </tr>
                                <%end if%>
								<%If F_B_Arr(16)=1 Then%>
								<tr class="tdbg">
                                        <td height="25" align="center"><span>阅读<%=KS.Setting(45)%>：</span></td>
                                        <td height="25">
										 <input type="text" style="text-align:center" name="ReadPoint" class="textbox" value="<%=ReadPoint%>" size="6"><%=KS.Setting(46)%> 如果免费阅读请输入“<font color=red>0</font>”
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
</table>
                  </form>

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
			
			</script>




			
			 <script type="text/javascript">
			  
			 <%if Action<>"Edit" Then%>
			 var LastNum=1;
			 <%else%>
			 var LastNum=$('#picnum').val();
			 <%end if%>
			 var tempup='';
			 var picnum=<%=Picnum%>;
		 	 $(document).ready(function(){
				 MakeUpload(<%=Picnum%>);
				 IniPicUrl();
				 tempup=$("#uploadfield").html();
			  })
			  
			function IniPicUrl()
			{
			 var PicUrls='<%=replace(PicUrls,vbcrlf,"\t\n")%>';
			  var PicUrlArr=null;
			  if (PicUrls!='')
			   { 
				PicUrlArr=PicUrls.split('|||');
			   for ( var i=1 ;i<PicUrlArr.length+1;i++)
			   { 
				 $('input[name=imgurl'+i+']').val(PicUrlArr[i-1].split('|')[1]);
				 $('input[name=thumb'+i+']').val(PicUrlArr[i-1].split('|')[2]);
				 $('#imgnote'+i).val(PicUrlArr[i-1].split('|')[0]);
				 $('#picview'+i).html('');
				 if (document.all){
				 $('#picview'+i)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=PicUrlArr[i-1].split('|')[1];
				 }else{
				  $('#picview'+i).html('<img width="120" height="80" src="'+PicUrlArr[i-1].split('|')[1]+'">');
				 }
			   }
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
				   fhtml += " <td width=\"124\" rowspan=\"3\" align=\"center\"><div id=\"picview"+startNum+"\" name=\"picview"+startNum+"\" style=\"filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:80px;width:120px;border:1px solid #777777\"><img src=\"images/view.gif\" width=\"120\" height=\"80\"></div></td>";
				   fhtml += "</tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"25\"> 　小图地址： ";
				   fhtml += "<input type=\"text\" class='textbox' onblur='view("+startNum+");' name='thumb"+startNum+"' id='thumb"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取<%=KS.C_S(ChannelID,3)%>\" onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=<%=ChannelID%>',550,290,window,document.myform.thumb"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;";
				   fhtml += "<br>　大图地址： <input type=\"text\" class='textbox' onblur='view("+startNum+");' name='imgurl"+startNum+"' id='imgurl"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取<%=KS.C_S(ChannelID,3)%>\" onclick=\"OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("选择" & KS.C_S(ChannelID,3))%>&ChannelID=<%=ChannelID%>',550,290,window,document.myform.imgurl"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;";
				   <%If KS.S("Action")="Edit" Then%>
					if (startNum>picnum)
					{
				   fhtml += "<br><iframe onload='ReSizeiFrame(this)' onreadystatechange='ReSizeiFrame(this)' id='UpPhotoFrame"+startNum+"' name='UpPhotoFrame"+startNum+"' src='User_UpFile.asp?ChannelID=<%=ChannelID%>&type=Single&objid="+startNum+"' frameborder=0 scrolling=no width='100%' height='22'></iframe>"
				    }
				   <%end if%>

				   fhtml += "</td></tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"30\">　<%=KS.C_S(ChannelID,3)%>简介： ";
				   fhtml += "<textarea class='textbox' name='imgnote"+startNum+"' id='imgnote"+startNum+"' style=\"height:46px;width:350px\"></textarea> </td>";
				   fhtml += "</tr></table>\r\n";
			  }
			  <%If KS.S("Action")="Edit" Then%>
			  $("#uploadfield").html(tempup+fhtml);
			  IniPicUrl();
			  <%Else%>
			  $("#uploadfield").html(fhtml);
			  frames['UpPhotoFrame'].ChooseOption(mnum);
			  ReSizeiFrame($('#UpPhotoFrame')[0]);
			 // $('#UpPhotoFrame').height(80+26*(mnum/2)); 
			  <%End If%>
			 // parent.init();
			}
			 function view(num)
			 {
			  if ($("input[name=thumb"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			     if (document.all){
			     $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=thumb"+num+"]").val();}else{ $("#picview"+num).html("<img width='120' height='80' src='"+$("input[name=thumb"+num+"]").val()+"'>");
			    }
			  }
			  else if($("input[name=imgurl"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			   if (document.all){
			       $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=imgurl"+num+"]").val();
			    }else{
				 $("#picview"+num).html("<img width='120' height='80' src='"+$("input[name=imgurl"+num+"]").val()+"'>");
				}
			  }
			 }

			 function SetPicUrlByUpLoad(DefaultThumb,PicUrlStr,ThumbPathFileName)
			{ var UrlStrArr;
			   UrlStrArr=PicUrlStr.split('|');
			   for (var i=1;i<UrlStrArr.length;i++)
			   {
			   var url=UrlStrArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=imgurl'+i+']').val(url);
			   } 
			  }
			  var ThumbsArr=ThumbPathFileName.split("|")
			  for(var i=1;i<ThumbsArr.length;i++)
			  {
			   var url=ThumbsArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=thumb'+i+']').val(url);
			   } 
			  }

			 $('input[name=PhotoUrl]').val(ThumbsArr[DefaultThumb-1]);
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
				<%Call KSUser.ShowUserFieldCheck(ChannelID)%>	
				
				 $('#PicUrls').val('');
				for(var i=1;i<=$("#picnum").val();i++){
				  if ($('input[name=imgurl'+i+']').val()!=''&&$('input[name=imgurl'+i+']').val()!='del') 
				   {
				   var note=$('#imgnote'+i).val();
				   note=note.replace('|||','');
				   spic=$('input[name=imgurl'+i+']').val();
				   tpic=$('input[name=thumb'+i+']').val();
				   if (tpic=='') tpic=spic;
				   if ($('input[name=PicUrls]').val()==''){                 
				    
				   $('input[name=PicUrls]').val(note+'|'+spic+'|'+tpic);
				   }else {
				   $('input[name=PicUrls]').val($('input[name=PicUrls]').val()+'|||'+note+'|'+spic+'|'+tpic);
				   }
				   
				  }
				}
				if ($('input[name=PicUrls]').val()=='')
				{
				  alert('请输入<%=KS.C_S(ChannelID,3)%>内容!');
				  $('input[name=imgurl1]').focus();
				  return false;
				}
				
                    $('#myform').submit();  
				}
				function CheckClassID()
				{
				 if (document.myform.ClassID.value=="0") 
				  {
					alert("请选择<%=KS.C_S(ChannelID,3)%>栏目！");
					return false;
				  }		
				  return true;
				}
			</script>
			 <%
  End Sub
  
  Sub DoSave()
  				Dim ClassID:ClassID=KS.S("ClassID")
				Dim Title:Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
				Dim KeyWords:KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				Dim Author:Author=KS.LoseHtml(KS.S("Author"))
				Dim Origin:Origin=KS.LoseHtml(KS.S("Origin"))
				Dim Content
				Content = KS.FilterIllegalChar(Request.Form("Content"))
				Content=KS.ClearBadChr(content)
				If Content="" Then content=" "
				Dim Verific:Verific=KS.ChkClng(KS.S("Status"))
				Dim PhotoUrl:PhotoUrl=KS.S("PhotoUrl")
				Dim PicUrls:PicUrls=KS.S("PicUrls")
				 If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
				 If KS.ChkClng(KS.S("ID"))<>0 Then
				  If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
				 End If
                 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '特殊VIP用户无需审核
				 
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
				 Dim Fname,FnameType,TemplateID,WapTemplateID
				 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
				 RSC.Open "select TemplateID,FnameType,FsoType,WapTemplateID from KS_Class Where ID='" & ClassID & "'",conn,1,1
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
				  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select  top 1 * From " & KS.C_S(ChannelID,2) & " Where Inputer='" & KSUser.UserName & "' and ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				If RSObj.Eof Then
				  RSObj.AddNew
				  RSObj("Inputer")=KSUser.UserName
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
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
				  RSObj("Comment")=1
				  If F_B_Arr(16)=1 Then
				   RSObj("ReadPoint")=KS.ChkClng(KS.S("ReadPoint"))
				  End If
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
				  Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
				  Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,0)
				  Call KSUser.AddLog(KSUser.UserName,"在栏目[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]上传了" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",2)
				  KS.Echo "<script>if (confirm('" & KS.C_S(ChannelID,3) & "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗?')){location.href='User_MYPhoto.asp?ChannelID=" & ChannelID & "&Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';}</script>"
				Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PicUrls & PhotoUrl & Content ,1)
			     Call KSUser.AddLog(KSUser.UserName,"对" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""做了修改!",2)
				 KS.Echo "<script>alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='User_MyPhoto.asp?ChannelID=" & ChannelID &"';</script>"
				End If
  End Sub
  

End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Contributor
KSCls.Kesion()
Set KSCls = Nothing

Class Contributor
        Private KS,KSUser,F_B_Arr,F_V_Arr,ChannelID,ClassID,LoginTF,Qid
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		Call KSUser.Head()
		Call KSUser.InnerLocation("匿名投稿")
		LoginTF=KSUser.UserLoginChecked
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		ClassID=KS.S("ClassID")
		  Dim Action:Action=KS.S("Action")
			Select Case Action
			 Case "Next" Call ContributorNext()
			 Case "AddSave" Call ContributorSave()
			 Case Else  Call Main()
			 End Select
	    End Sub 
		
		Function GetQuestionRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetQuestionRnd=RandNum
		End Function
		
		Function PubQuestion()
			if mid(KS.Setting(161),2,1)="1" then
			 Qid=GetQuestionRnd
			%>
						   <tr class="tdbg">
                            <td  height="25" align="center"><span>请回答投稿问题：</span></td>
                             <td>
							 　 <font color="red"><%
							 Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		                     response.write QuestionArr(Qid)
							 %></font>
							 　</td>
                          </tr>
						   <tr class="tdbg">
                            <td  height="25" align="center"><span>您的答案：</span></td>
                            <td>　
							 <input type="text" id="QuestionAnswer" name="a<%=md5(Qid,16)%>">
							</td>
                          </tr>
			<%end if
		End Function
		
		
		'选择投稿栏目
		Sub Main()
		%>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0"  style="margin-top:5px">
		  <tr>
			<td>
			 <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
			  <TR>
				<TD width=5 height=5><img height=5 src="images/p13.gif" width=5></TD>
				<TD width=700 background=images/p29.gif height=5></TD>
				<TD align=right width=5 height=5><img height=5 src="images/p14.gif" width=5></TD>
			  </TR>
			</TABLE></td>
		  </tr>
		  <tr>
			<td align="left" valign="top">
			  <TABLE width=100% height=200 border=0 cellPadding=0 cellSpacing=0 background=images/p15.gif>
				  <TR>
					<TD align=middle><TABLE height=190 cellSpacing=0 cellPadding=0 width=692  background=images/p18.gif border=0>
						  <TR>
							<TD align=middle valign="top">
							<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
							  <script language="javascript">
							    function CheckForm()
								{
								 if (document.form1.classid.value=='')
								 {
								  alert('请选择投稿栏目!');
								  return false
								 }
								 return true;
								}
							  </script>
							   <form name="form1" action="?Action=Next" method="post" onSubmit="return(CheckForm());">
								<tr class="title">
								  <td height="22" align="center">请选择要投稿的栏目:</td>
								</tr>
								<tr class="tdbg">
								  <td align="center">
								  <select name=classid size="22" style="width:300px">
								  <%
								  Dim CacheID,K,SQL,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
								  RS.Open "Select ID,FolderName,a.ChannelID From KS_Class a inner join ks_channel b on a.channelid=b.channelid Where UserTF=2 and a.ChannelID<>5 and CommentTF=2 order by a.ChannelID,folderorder",Conn,1,1
								  If Not RS.Eof Then SQL=RS.GetRows(-1)
								  RS.Close:Set RS=Nothing
								  If IsArray(SQL) Then
								   For K=0 To Ubound(SQL,2)
									 If SQL(2,k)<>CacheID Then
									  Response.Write "<optgroup  label='===============" & KS.C_S(SQL(2,k),3) & "栏目=============='>"
									 End If
									 Response.Write "<option value='" & SQL(0,K) & "'>|-" & SQL(1,K) & "</option>"
									 
									 CacheID=SQL(2,K)
								   Next
								  End If
								  %>
								  </select>								 
								   </td>
								</tr>
								
								<tr class="tdbg">
								  
								  <td height="22" align="center">
								   <input type="submit" name="s1" value=" 下 一 步 " class="button">
								   </td>
								</tr>
								<tr class="tdbg">
								  <td height="22" align="center"><font color=red>温馨提示：匿名投稿禁止所有上传功能，如果想享受本站会员的更多服务，请注册成为本站会员！</font></td>
							     </tr>
								</form>
							</table></TD>
						  </TR>
					  </TABLE>
					  <TABLE cellSpacing=0 cellPadding=0 width=692 bgColor=#ffffff border=0>
						  <TR>
							<TD><img height=5 src="images/p20.gif" width=692></TD>
						  </TR>
					  </TABLE></TD>
				  </TR>
				</TABLE>
			  </td>
		  </tr>
		  <tr>
			<td><TABLE cellSpacing=0 width="100%" cellPadding=0  border=0>
			  <tr>
				<td width=5 height=5><img height=5 src="images/p22.gif"  width=5></td>
				<td width=700 background=images/p28.gif height=5></td>
				<td align=right width=5 height=5><img height=5 src="images/p23.gif" width=5></td>
			  </tr>
			</TABLE></td>
		  </tr>
		</table>
<%
  End Sub
  
   '选择投稿界面
   Sub ContributorNext()
     ClassID=KS.R(KS.S("ClassID"))
	 If ClassID="" Then Response.Write "<script>alert('对不起，你没有选择投稿栏目!');history.back();</script>":Response.End
	 ChannelID=KS.ChkClng(Conn.Execute("Select ChannelID From KS_Class Where ID='" & ClassID & "'")(0))
	 If ChannelID=0 Then Response.End()
	 If LoginTF=True Then
		   Select Case KS.C_S(ChannelID,6)
		    Case 1 Response.Redirect "User_MyArticle.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 2 Response.Redirect "User_MyPhoto.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 3 Response.Redirect "User_MySoftWare.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 4 Response.Redirect "User_MyFlash.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 7 Response.Redirect "User_MyMovie.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		    Case 8 Response.Redirect "User_MySupply.asp?action=Add&channelid=" & ChannelID & "&ClassID=" & ClassID
		   End Select
	 End If
	 
   		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
	 Select Case KS.C_S(ChannelID,6) 
	   Case 1:Call AddByArticle()
	   Case 2:Call AddByPicture()
	   Case 3:Call AddBySoftWare()
	   Case 4:Call AddByFlash()
	   Case 7:Call AddByMovie()
	   Case 8:Call AddBySupply()
	   Case Else:Response.Write "参数出错!":Response.End()
	 End Select 
   End Sub
   
   '保存投稿
   Sub ContributorSave()
     ChannelID=KS.ChkCLng(KS.S("ChannelID"))
	  If ChannelID=0 Then Response.End()
	  IF Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) then 
	   Call KS.AlertHistory("验证码有误，请重新输入！",-1)
	   exit Sub
	  End If
	  If Request.ServerVariables("HTTP_REFERER")="" Then
	   Call KS.AlertHistory("非法提交！",-1)
	   exit Sub
	  End If
	  '检查注册回答问题
	  Dim CanReg,N
	   If Mid(KS.Setting(161),2,1)="1" Then
		     CanReg=false
		     For N=0 To Ubound(Split(KS.Setting(162),vbcrlf))
			   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
			      If Lcase(Request.Form("a" & MD5(n,16)))<>Lcase(Split(KS.Setting(163),vbcrlf)(n)) Then
			       Call KS.AlertHistory("对不起,注册问题的回答不正确!",-1) : Response.End
				   CanReg=false
				  Else
				   CanReg=True
				  End If
			   End If
			 Next
			 If CanReg=false Then Call KS.AlertHistory("对不起,注册答案不能为空!",-1) : Response.End
	  End If
	  
	  
   	 F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
	 F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
     Select Case KS.C_S(ChannelID,6)
	  Case 1:Call SaveByArticle()
	  Case 2:Call SaveByPhoto()
	  Case 3:Call SaveByDownLoad()
	  Case 4:Call SaveByFlash()
	  Case 7:Call SaveByMovie()
	  Case 8:Call SaveBySupply()
	 End Select	 
   End Sub
   
   
   
   '添加文章
   Sub AddByArticle()
   %>
      <script language = "JavaScript">
				function insertHTMLToEditor(codeStr) 
				{ 
					oEditor=FCKeditorAPI.GetInstance("Content");
					if(oEditor   &&   oEditor.EditorWindow){ 
						oEditor.InsertHtml(codeStr); 
					} 
				} 
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>标题！");
					document.myform.Title.focus();
					return false;
				  }	
				<%if F_B_Arr(9)=1 Then%> 
				    if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=="")
					{
					  alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！");
					  return false;
					}
				<%end if%>
				if (document.myform.VerifyCode.value=="")
				 {
					alert("请输入验证码！");
					document.myform.VerifyCode.focus();
					return false;
				 }	
				 <%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>
				 return true;  
				}
				function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}
				</script>
				
				
				<table  width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form  action="?channelid=<%=channelid%>&Action=AddSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
				    <tr  class="title">
					  <td colspan=2 align=center><% ="发布" & KS.C_S(ChannelID,3) %></td>
					</tr>
                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
                       <td width="88%">　
							[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
							 <input type="hidden" name="ClassID" value="<%=classid%>">
					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
                              <td> 　
                               <input class="textbox" name="Title" type="text" style="width:250px; " maxlength="100" /> <span style="color: #FF0000">*</span></td>
                        </tr>
						<%if F_B_Arr(2)=1 Then%>	  
                      <tr class="tdbg">
                           <td  height="25" align="center"><span><%=F_V_Arr(2)%>：</span></td>
                              <td> 　
                               <input class="textbox" name="FullTitle" type="text" style="width:250px; " maxlength="100" /></td>
                        </tr>
						<%End If%>
						<%if F_B_Arr(5)=1 Then%>	  
                              <tr class="tdbg">
                                      <td height="25" align="center"><span><%=F_V_Arr(5)%>：</span></td>
                                <td>　
                                        <input name="KeyWords"  class="textbox" type="text" id="KeyWords" style="width:250px; " />
                                  多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开								</td>
                              </tr>
					  <%end if%>
						<%if F_B_Arr(6)=1 Then%>	  
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
                                      <td height="25">　
                                        <input name="Author" class="textbox" type="text" id="Author" style="width:250px; "maxlength="30" /></td>
                              </tr>
						<%end if%>
						<%if F_B_Arr(7)=1 Then%>	  
                              <tr class="tdbg">
                                
                                      <td  height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
                                      <td>　
                                        <input class="textbox" name="Origin" type="text" id="Origin" style="width:250px; " maxlength="100" /></td>
                              </tr>
						<%end if%>
							  <%
							  Response.Write KSUser.KS_D_F(ChannelID,"0")
							  %>
						<%if F_B_Arr(8)=1 Then%>	  
                              <tr class="tdbg">
                                 <td  height="25" align="center"><span><%=F_V_Arr(8)%>：</span><br><input name='AutoIntro' type='checkbox' checked value='1'><font color="#FF0000">自动截取内容的200个字作为导读</font></td>
                                <td>　
                               <textarea class='textbox' name="Intro" style='width:95%;height:80'></textarea></td>
                              </tr>
						<%end if%>

						<%if F_B_Arr(9)=1 Then%>	  
                              <tr class="tdbg">
                                  <td><%=F_V_Arr(9)%>:<br><img src="images/ico.gif" width="17" height="12" /><font color="#FF0000">如果<%=KS.C_S(ChannelID,3)%>较长可以使用分页标签：[NextPage]</font>
								  </td>
								  <td align="center">
				                 <% 
								 Response.Write "<textarea name=""Content"" style=""display:none""></textarea>"
								 Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""95%"" height=""350"" frameborder=""0"" scrolling=""no""></iframe>"  
								 %>
								</td>
                            </tr>
					    <%end if%>
						<%Call PubQuestion()%>
						
						
                          <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input name="VerifyCode" onFocus="this.value='';getCode()" type="text" id="VerifyCode" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>
                    <tr class="tdbg">
                      <td height="30" align="center" colspan=2>
					 <input class="button" type="submit" name="Submit" value="OK, 提交稿件 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重 来 " />						</td>
                    </tr>
                  </form>
			    </table>
		  <%
    End Sub
	
	'添加图片
	Sub AddByPicture()
		%>
                  <form  action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post" id="myform" name="myform">
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				           <tr class="title">
						    <td colspan=2 align=center>
							 <%= "发布" & KS.C_S(ChannelID,3)%>
							</td>
						   </tr>
                           <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(1)%>：</span></td>
                                        <td>　
							[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
							 <input type="hidden" name="ClassID" value="<%=classid%>">
										</td>
                             </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(0)%>：</span></td>
                                        <td> 　 
                                          <input class="textbox" name="Title" type="text" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
								<%If F_B_Arr(6)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(6)%>：</span></td>
                                  <td>　
                                          <input name="KeyWords" class="textbox" type="text" id="KeyWords" style="width:250px; " /> 
                                    多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开</td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(7)=1 Then%>
                                <tr class="tdbg">

                                        <td height="25" align="center"><span><%=F_V_Arr(7)%>：</span></td>
                                        <td height="25">　
                                          <input class="textbox" name="Author" type="text" id="Author" style="width:250px; " maxlength="30" /></td>
                                </tr>
								<%end if%>
								<%If F_B_Arr(8)=1 Then%>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(8)%>：</span></td>
                                        <td>　
                                          <input class="textbox" name="Origin" type="text" style="width:250px; " maxlength="100" /></td>
							  </tr>
							  <%End if%>
								<%
							  Response.Write KSUser.KS_D_F(ChannelID,"0")
							  %>
							  <tr class="tdbg">
                                        <td height="35" align="center"><span><%=F_V_Arr(2)%>：</span></td>
                                        <td>　
                                          <input class='textbox' name='PhotoUrl' type='text' style="width:250px;" id='PhotoUrl' maxlength="100" />
                                          <font color='#FF0000'>*</font>&nbsp;
                                          </td>
							   </tr>
							  <tr class="tdbg">
                                    <td height="40" align="center"><span><%=F_V_Arr(3)%>：</span></td>
                                    <td>&nbsp;&nbsp;&nbsp;<input name='picnum' class='textbox' type='text' id='picnum' size='4' value='4' style='text-align:center'>&nbsp;<input name='kkkup' type='button' id='kkkup2' value='设定' onClick="MakeUpload($F('picnum'),'click');" class='button'>注：最多<font color='red'>99</font><%=KS.C_S(ChannelID,4)%>，匿名发表不支持上传，请输入以http开头的远程<%=KS.C_S(ChannelID,3)%>地址<input type='hidden' name='PicUrls'> </td>
                              </tr>
								<tr class="tdbg">
                                   <td height="220" align="center"><span><%=F_V_Arr(4)%>：</span></td>
                                   <td align="center">
								   <span id='uploadfield'></span>
								   </td>
							  </tr>
							  
								<%If F_B_Arr(9)=1 Then%>
							   <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(9)%>：<br /></td>
                                        <td align="center">　
									 <%	
									 Response.Write "<textarea name=""Content"" style=""display:none""></textarea>"
								     Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""95%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
								      %>     
									   </td>
                                </tr>
                                <%end if
								Call PubQuestion
								
								%>
                          <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input class="textbox" name="VerifyCode" onClick="getCode()" type="text" id="Origin" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>
                               <tr class="tdbg">
                            <td align="center" colspan=2>
							<input class="button" type="button" onClick="CheckForm()" name="Submit" value=" OK,保存 " />
                            <input class="button" type="reset"  name="Submit2" value=" 重 来 " /></td>
                         </tr>
</table>
                  </form>
			 <script>
			 function view(num)
			 {
			  if ($("#thumb"+num).val()!='')
			  $("#picview"+num).attr("src",$("#thumb"+num).val());
			  else if($("#imgurl"+num).val()!='')
			  $("#picview"+num).attr("src",$("#imgurl"+num).val());
			 }

			 var LastNum=$('#picnum').val();
			 var tempup='';
			 var picnum=4;
		 	 $(document).ready(function()
			  { 
				 MakeUpload(4);
				 tempup=$("#uploadfield").html();
			  });

			function MakeUpload(mnum,str)
			{ 
			   if (parseInt(mnum)>=100){
			   alert('最多只能同上传99张!');
			   return false;}
			   var startNum=1;
			   var endNum = mnum;
			   var fhtml = "";
			   for(startNum;startNum <= endNum;startNum++){
				   fhtml += "<table width=\"99%\" style='margin:2px' class='border' align=center border=\"0\" id=\"seltb"+startNum+"\" cellpadding=\"3\" cellspacing=\"1\">";
				   fhtml += "<tr class='tdbg'> "
				   fhtml +="  <td height=\"25\" width=18 align=center class=clefttitle rowspan=\"3\"><strong>第"+startNum+"张</strong></td>";
				   fhtml += " <td width=\"124\" rowspan=\"3\" align=\"center\"><img src=\"images/view.gif\" width=\"120\" height=\"80\" border=1 id=\"picview"+startNum+"\" name=\"picview"+startNum+"\"></td>";
				   fhtml += "</tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"25\"> 　小图地址： ";
				   fhtml += "<input type=\"text\" class='textbox' onclick='view("+startNum+");' onblur='view("+startNum+");' name='thumb"+startNum+"' id='thumb"+startNum+"' size=\"42\" value=\"\"> ";
				  
				   fhtml += "<br>　大图地址： <input type=\"text\" class='textbox' onclick='view("+startNum+");' onblur='view("+startNum+");' name='imgurl"+startNum+"' id='imgurl"+startNum+"' size=\"42\" value=\"\"> ";
				   fhtml += "</td></tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"30\">　<%=KS.C_S(ChannelID,3)%>简介： ";
				   fhtml += "<textarea class='textbox' name='imgnote"+startNum+"' id='imgnote"+startNum+"' style=\"height:46px;width:350px\"></textarea> </td>";
				   fhtml += "</tr></table>\r\n";
			  }
			  $("#uploadfield").html(fhtml);
			  parent.init();
			}
			 
			     function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}

				function CheckForm()
				{
				
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
                <%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>                    
				if (document.myform.VerifyCode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.VerifyCode.focus();
					return false;
				  }	
				
                  $('#myform').submit();  
				}
			</script>
			 <%
	End Sub
	
	'添加下载
	Sub AddBySoftWare()
		 Dim I,DownLBList, DownYYList, DownSQList, DownPTList, RSP, DownLBStr, LBArr, YYArr, SQArr, PTArr, DownYYStr, DownSQStr, DownPTStr
		     Set RSP = Server.CreateObject("Adodb.RecordSet")
			 RSP.Open "Select * From KS_DownParam", conn, 1, 1
			 DownLBStr = RSP("DownLB")
			 DownYYStr = RSP("DownYY")
			 DownSQStr = RSP("DownSQ")
		     DownPTStr = RSP("DownPT")
			 RSP.Close: Set RSP = Nothing
					  '下载类别
					 ' DownLBList="<option value="""" selected> </option>"
					  LBArr = Split(DownLBStr, vbCrLf)
					  For I = 0 To UBound(LBArr)
						DownLBList = DownLBList & "<option value='" & LBArr(I) & "'>" & LBArr(I) & "</option>"
					  Next
					  '下载语言
					   ' DownYYList="<option value="""" selected> </option>"
					  YYArr = Split(DownYYStr, vbCrLf)
					  For I = 0 To UBound(YYArr)
						DownYYList = DownYYList & "<option value='" & YYArr(I) & "'>" & YYArr(I) & "</option>"
					  Next
					'下载授权
					   ' DownSQList="<option value="""" selected> </option>"
					  SQArr = Split(DownSQStr, vbCrLf)
					  For I = 0 To UBound(SQArr)
						DownSQList = DownSQList & "<option value='" & SQArr(I) & "'>" & SQArr(I) & "</option>"
					  Next
					'下载平台
					  'DownPTList="<option value="""" selected> </option>"
					  PTArr = Split(DownPTStr, vbCrLf)
					  For I = 0 To UBound(PTArr)
						DownPTList = DownPTList & "<a href='javascript:SetDownPT(""" & PTArr(I) & """)'>" & PTArr(I) & "</a>/"
					  Next
					 %>
				
				<script language="javascript">
				function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}
				function SetDownPT(addTitle){
					var str=document.myform.DownPT.value;
					if (document.myform.DownPT.value=="") {
						document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
					}else{
						if (str.substr(str.length-1,1)=="/"){
							document.myform.DownPT.value=document.myform.DownPT.value+addTitle;
						}else{
							document.myform.DownPT.value=document.myform.DownPT.value+"/"+addTitle;
						}
					}
					document.myform.DownPT.focus();
				}

				function SetPhotoUrl()
				{
				 if (document.myform.DownUrl.value!='')
				  document.myform.PhotoUrl.value=document.myform.DownUrl.value.split('|')[1];	
				}

				function CheckForm()
				{   
					
				 if (document.myform.Title.value=="")
					  {
						alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
						document.myform.Title.focus();
						return false;
					  }
					if (document.myform.DownUrlS.value=='')
					{
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.DownUrlS.focus();
						return false;
					}
                 <%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>				
				 if (document.myform.VerifyCode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.VerifyCode.focus();
					return false;
				  }	
					document.myform.submit();
					return true;
				}
				</script>

				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
                  <form action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post" name="myform" id="myform">     
					        <tr class="title">
							 <td colspan=2 align=center>
							 <%="发布" & KS.C_S(ChannelID,3)%>
							 </td>
							</tr>
                             <tr class="tdbg">
                                   <td width="12%" height="25" align="center"><%=F_V_Arr(1)%>：</td>
                                    <td width="88%">　
							[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
							 <input type="hidden" name="ClassID" value="<%=classid%>">

									</td>
                    </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(0)%>：</td>
                                        <td> 　 
                                          <input class="textbox" name="Title" type="text" id="Title" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
						<%if F_B_Arr(10)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=F_V_Arr(10)%>：</span></td>
                                  <td>　
                                          <input class="textbox" name="KeyWords" type="text" id="KeyWords" style="width:250px; " /> 
                                    多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开</td>

                                </tr>
					   <%end if%>
						<%if F_B_Arr(11)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(11)%>：</td>
                                        <td height="25">　
                                        <input class="textbox" name="Author" type="text" id="Author" style="width:250px; " maxlength="30" /></td>
                                </tr>
					  <%End If%>
						<%if F_B_Arr(12)=1 Then%>	  
                                <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(12)%>：</td>
                                        <td>　
                                        <input class="textbox" name="Origin" type="text" id="Origin" style="width:250px; " maxlength="100" /></td>
								</tr>
					  <%end if%>
						<%if F_B_Arr(6)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(6)%>：</td>
                                        <td>　
                                       类别:<select name='DownLB'>
		                               <%=DownLBList%>
		                                </select> 语言:<select name='DownYY' size='1'>
		                               <%=DownYYList%>
		                               </select>授权:<select name='DownSQ' size='1'>
		                               <%=DownSQList%></select><%
									 Response.Write "大小:<input type='text' size=4 name='DownSize'>&nbsp;"
									Response.Write "  <input name=""SizeUnit"" type=""radio"" value=""KB"" checked id=""kb""><label for=""kb"">KB</label> " & vbCrLf
									Response.Write "  <input type=""radio"" name=""SizeUnit"" value=""MB"" id=""mb""><label for=""mb"">MB</label> " & vbCrLf
									%>                      
		                               </td>
								</tr>
					<%end if%>
						<%if F_B_Arr(7)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(7)%>：</td>
                                        <td>　
                                        <input class='textbox' type='text' size=70 name='DownPT'><br>
		                                &nbsp;<font color='#808080'>平台选择
		                                <%=DownPTList%></font></td>
				               </tr>
						<%end iF%>
						<%if F_B_Arr(15)=1 Then%>	  
								<tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(15)%>：</td>
                                        <td>　
                                        <input class="textbox" name="YSDZ" type="text" id="YSDZ" style="width:250px; " maxlength="100" /></td>
                               </tr>
					   <%end if%>
						<%if F_B_Arr(16)=1 Then%>	  
								<tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(16)%>：</td>
                                        <td>　
                                        <input class="textbox" name="ZCDZ" type="text" id="ZCDZ" style="width:250px; " maxlength="100" /></td>

								</tr>
					 <%end if%>
						<%if F_B_Arr(17)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(17)%>：</td>
                                        <td>　
                                        <input class="textbox" name="JYMM" type="text"  id="JYMM" style="width:250px; " maxlength="100" /></td>
                              </tr>
						<%end if%>
                             <%
							  Response.Write KSUser.KS_D_F(ChannelID,"0")
							  %> 
						<%if F_B_Arr(8)=1 Then%>	  
								 <tr class="tdbg">
                                        <td height="25" align="center"><%=F_V_Arr(8)%>：</td>
                                        <td>　
                                        <input class="textbox"  name="PhotoUrl"  type="text" id="PhotoUrl" style="width:250px; " maxlength="100" /> 
                                        </td>
                              </tr>
					   <%end if%>
							   <tr class="tdbg">
                                    <td height="25" align="center"><%=KS.C_S(ChannelID,3)%>地址：</td>
                                    <td valign="top">　
  <input type="text" class="textbox" name='DownUrlS'  size="50"> 
  <span style="color: #FF0000">* </span>请填写下载地址 </td>
								</tr>
						<%if F_B_Arr(14)=1 Then%>	  
								 <tr class="tdbg">
                                        <td align="center"><%=F_V_Arr(14)%>：<br />
                                          </td>
                                        <td align="center">
							   <%
									 Response.Write "<textarea name=""Content"" style=""display:none""></textarea>"
								     Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""95%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
			                     %>
											 
										</td>
                                                  
                                </tr>
						<%end if
						 call PubQuestion
						%>
						    <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input class="textbox" name="VerifyCode" onClick="getCode()" type="text" id="Origin" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>

                        <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" name="Submit" onClick="return CheckForm();" value=" OK！发 布 " />
                            <input type="reset" class="button"  name="Submit2" value=" 重来 " /></td>

                    </tr>
                  </form>
</table>
		  
		  <%	End Sub
	
	'添加动漫
	Sub AddByFlash()
	%>

			<script language = "JavaScript">
			    function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.FlashUrl.value=='')
					{
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.FlashUrl.focus();
						return false;
					}
				<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>		
				if (document.myform.VerifyCode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.VerifyCode.focus();
					return false;
				  }	
				  document.myform.submit();
				 return true;  
				}
				</script>

							<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
							  <tr class="title">
							   <td colspan=2 align=center>
										 <%="发布" & KS.C_S(ChannelID,3)%>
							   </td>
							  </tr> 
							  <form action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post" name="myform" id="myform">
							<tr class="tdbg">
									   <td height="25" align="center">所属栏目：</td>
									   <td>　
										[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
										 <input type="hidden" name="ClassID" value="<%=classid%>">
								  </td>
								</tr>
                                <tr class="tdbg">

                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>名称：</span></td>
                                        <td> 　 
                                          <input name="Title" class="textbox" type="text" id="Title" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span>关 键 字：</span></td>
                                  <td>　
                                          <input name="KeyWords" class="textbox" type="text" id="KeyWords" style="width:250px; " /> 
                                    多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开</td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>作者：</span></td>
                                        <td height="25">　
                                        <input name="Author" class="textbox" type="text" style="width:250px; "  maxlength="30" /></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>来源：</span></td>
                                        <td>　
                                        <input name="Origin" class="textbox" type="text" id="Origin" style="width:250px; " maxlength="100" /></td>
								</tr>
                              <%
							  Response.Write KSUser.KS_D_F(ChannelID,"0")
							  %>     
							  
								<tr class="tdbg">
                                        <td height="25" align="center"><span>缩 略 图：</span></td>
                                        <td>　
                                          <input class="textbox" name='PhotoUrl' type='text' style="width:250px;" id='PhotoUrl' maxlength="100" />
                                         </td>
							   </tr>
								
								<tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>地址：</span></td>
                                        <td>　
                                          <input class="textbox" name='FlashUrl'  type='text' style="width:250px;" id='FlashUrl' maxlength="100" />
                                          <font color='#FF0000'>*</font>
                                          </td>
							   </tr>
								
  								<tr class="tdbg">
                                        <td align="center"><span><%=KS.C_S(ChannelID,3)%>简介：<br />
                                          </span></td>
                                        <td>
										<table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="12">&nbsp;</td>
                                              <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td height="150" align="center">
										<%
									 Response.Write "<textarea name=""Content"" style=""display:none""></textarea>"
								     Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""95%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
											 %>			
													</td>
                                                  </tr>
                                              </table></td>
                                            </tr>
                                        </table></td>
                                </tr>
						<%
								call PubQuestion
						%>
						    <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input class="textbox" name="VerifyCode" onClick="getCode()" type="text" id="Origin" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" onClick="return CheckForm();" name="Submit" value=" OK! 发布 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重来 " /></td>
                          </tr>
                  </form>
</table>
				
		  <%
	End Sub
	
	Sub AddByMovie()
		%>
	  <script language = "JavaScript">
				function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}
				function CheckForm()
				{
				if (document.myform.Title.value=="")
				  {
					alert("请输入<%=KS.C_S(ChannelID,3)%>名称！");
					document.myform.Title.focus();
					return false;
				  }		
				if (document.myform.MovieUrl.value=='')
					{
						alert("请添加<%=KS.C_S(ChannelID,3)%>！");
						document.myform.MovieUrl.focus();
						return false;
					}
					<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
				 <%end if%>	
				if (document.myform.VerifyCode.value=="")
				  {
					alert("请输入验证码！");
					document.myform.VerifyCode.focus();
					return false;
				  }	
				  document.myform.submit();
				 return true;  
				}
		</script>

							<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
							  <tr class="title">
							   <td colspan=2 align=center>
										 <%="发布" & KS.C_S(ChannelID,3)%>
							   </td>
							  </tr> 
							  <form  action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post" name="myform" id="myform">
							<tr class="tdbg">
									   <td height="25" align="center">所属栏目：</td>
									   <td>　
										[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
										 <input type="hidden" name="ClassID" value="<%=classid%>">
								  </td>
								</tr>
                                <tr class="tdbg">

                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>名称：</span></td>
                                        <td> 　 
                                          <input name="Title" class="textbox" type="text" id="Title" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span>关 键 字：</span></td>
                                  <td>　
                                          <input name="KeyWords" class="textbox" type="text" id="KeyWords" style="width:250px; " /> 
                                    多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开</td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span>主要演员：</span></td>
                                        <td height="25">　
                                        <input name="MovieAct" class="textbox" type="text" id="MovieAct" style="width:250px; "  maxlength="30" /></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>导演：</span></td>
                                        <td>　
                                        <input name="MovieDY" class="textbox" type="text" id="MovieDY" style="width:250px; " maxlength="100" /></td>
								</tr>
                             <%
							  Response.Write KSUser.KS_D_F(ChannelID,"0")
							  %>     
							  
								<tr class="tdbg">
                                        <td height="25" align="center"><span>缩 略 图：</span></td>
                                        <td>　
                                          <input class="textbox" name='PhotoUrl' type='text' style="width:250px;" id='PhotoUrl' maxlength="100" />
                                        </td>
							   </tr>
								<tr class="tdbg">
                                  <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>地址：</span></td>
                                  <td>　
                                          <input class="textbox" name='MovieUrl' type='text' style="width:250px;" id='MovieUrl' maxlength="100" /> <font color=red>*</font>影片的播放地址
                                          </td>
							   </tr>
  								<tr class="tdbg">
                                        <td align="center"><span><%=KS.C_S(ChannelID,3)%>简介：<br />
                                          </span></td>
                                        <td>
										<table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="12">&nbsp;</td>
                                              <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td height="150" align="center">
										   <%
									 Response.Write "<textarea name=""Content"" style=""display:none""></textarea>"
								     Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""95%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
											 %>
													
													</td>
                                                  </tr>
                                              </table></td>
                                            </tr>
                                        </table></td>
                                </tr>
								<%
								call PubQuestion
						      %>
						    <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input class="textbox" name="VerifyCode" onClick="getCode()" type="text" id="Origin" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" onClick="return CheckForm();" name="Submit" value=" OK! 发布 " />
                            　
                            <input class="button" type="reset" name="Submit2" value=" 重来 " /></td>
                          </tr>
                  </form>
</table>
				
		  <%
	End Sub
	
	'添加供求信息
	Sub AddBySupply()
	%>
	<SCRIPT language=JavaScript>
	function getCode()
				{
				$("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')

				}
var partten = "/^\d{8}$/"
function check()
{
if (document.myform.title.value.length<=4)
{
alert("信息标题要大于等于4个字符");
document.myform.title.focus();
document.myform.title.select()
return false; 
}
if (document.myform.Price.value=="")
{
alert("价格说明不能为空");
document.myform.Price.focus();
document.myform.Price.select();
return false; 
}
if (document.myform.TypeID.value =="") 
{ 
alert("请选择交易类别！"); 
document.myform.TypeID.focus(); 
return false; 
}

if (FCKeditorAPI.GetInstance('GQContent').GetXHTML(true)=="")
{
	alert("信息内容必须输入");
	return false; 
}
if (document.myform.ContactMan.value=="")
{
alert("联系人不能为空");
document.myform.ContactMan.focus();
document.myform.ContactMan.select() 
return false; 
}
if (document.myform.Tel.value=="")
{
alert("联系电话不能为空");
document.myform.Tel.focus();
document.myform.Tel.select() 
return false; 
}
<%if mid(KS.Setting(161),2,1)="1" Then%>
				 if ($("#QuestionAnswer").val()==""){
				  alert("请输入您的回答!");
				  $("#QuestionAnswer").focus();
				  return false;
				 }
 <%end if%>
if (document.myform.VerifyCode.value=="")
{
	alert("请输入验证码！");
	document.myform.VerifyCode.focus();
     return false;
}	
document.myform.submit();
}
</SCRIPT>
<body leftMargin="0" topMargin="0" marginheight="0">
<div align=center>
<CENTER>
  <table style="BORDER-COLLAPSE: collapse" borderColor=#111111 height=460 cellSpacing=1 width="100%" bgColor=#ffffff border=0>
    <tr>
      <td width="100%" height=457>
<FORM name="myform" action="?ChannelID=<%=ChannelID%>&Action=AddSave" method="post">
  <table style="BORDER-COLLAPSE: collapse" bordercolor=#111111 height=403 cellspacing=0 cellpadding=0 width="100%" border=0>
    <tr>
      <td width="100%" height=12></td>
    </tr>
    <tr>
      <td width="100%" height=22><table align="center" style="BORDER-COLLAPSE: collapse" bordercolor=#111111 height=20 cellspacing=0 cellpadding=0 width="98%" border=0>
          <tr>
            <td  width=23 height=20>&nbsp;</td>
            <td  width=160 bgcolor=#5298d1 height=20><b>&nbsp;<font color=#ffffff><span style="FONT-SIZE: 10.5pt">要发布的信息</span></font></b></td>
            <td width=12 height=20>&nbsp;</td>
            <td width=583 height=20><p align=right><font color=#ff0000>注：请不要发布重复信息，谢谢合作&nbsp;&nbsp;&nbsp;&nbsp; </font></p></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td width="100%" height=127><div align=center>
            <table width="98%" border=0 align="center" cellpadding=2 cellspacing=1 bordercolor=#111111 bordercolorlight=#ffffff bordercolordark=#ffffff style="BORDER-COLLAPSE: collapse">
              <tr class='tdbg'>
                <td height=25 align="center">信息分类：</td>
                <td>　
					[<%=KS.GetClassNP(ClassID)%>] <a href="user_Contributor.asp"><<重新选择>></a>
					<input type="hidden" name="ClassID" value="<%=classid%>">
				</td>
              </tr>
              <tr class='tdbg'>
                <td width="14%" height=25 align="center"><p>信息主题：</p></td>
                <td width="86%">　
<input class="textbox" size=45 name="title">
                    <font color=#ff6600> *</font></td>
                </tr>
              <tr class="tdbg">
                <td width="14%" height=25 align="center"><p>价格说明：</p></td>
                <td width="86%" height=25>　
<input class="textbox" size=45  name="Price">
                    <font color=#ff6600> *</font></td>
              </tr>
			  <tr class="tdbg">
                               <td height="25" align="center">图片地址：</td>
                               <td height="25">　
                                 <input name='PhotoUrl' type='text' id='PhotoUrl' size='45'  class="textbox"/></td></tr>
              <tr class="tdbg">
                <td width="14%" height=25 align="center">交易类别：</td>
                <td width="86%">　
					<%=KS.ReturnGQType(0,0)%>
                    <font color=#ff6600> *</font>　 有 效 期：
                    <select class="textbox" size=1 name="ValidDate">
					 <option value="3">三天</option>
					 <option value="7" selected>一周</option>
					 <option value="15">半个月</option>
					 <option value="30">一个月</option>
					 <option value="90">三个月</option>
					 <option value="180">半年</option>
					 <option value="365">一年</option>
					 <option value="0">长期</option>
                    </select>
                    <font color=#ff6600> *</font></td>
              </tr>
              <%
				 Response.Write KSUser.KS_D_F(8,"0")
			  %>
              <tr class="tdbg">
                <td align="center">信息内容：<br>
                  <font color=#800000>（请详细描述您发布的供求信息）</font></td>
                <td align="center">
										<%
									 Response.Write "<textarea name=""GQContent"" style=""display:none""></textarea>"
								     Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=GQContent&amp;Toolbar=Basic"" width=""95%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe>"  
											 %>                </td>
                </tr>
				 <tr class="tdbg">
                <td width="14%" height=25 align="center"><p>关键字Tags：</p></td>
                <td width="86%" height=25>　
<input name="KeyWords"  class="textbox" type="text" id="KeyWords" style="width:250px; " />
                                        多个关键字请用&quot;<span style="color: #FF0000">,</span>&quot;隔开	                   </td>
              </tr>
			   <%
								call PubQuestion
						      %>
						    <tr class="tdbg">
                            <td  height="25" align="center"><span>验证码：</span></td>
                             <td>　 <input class="textbox" name="VerifyCode" onClick="getCode()" type="text" id="Origin" style="width:50px; " maxlength="6" /><span id="showVerify">鼠标点击输入框获得验证码</span></td>
                          </tr>
            </table>
          </center>
      </div></td>
    </tr>
    <tr>
      <td width="100%" height=15></td>
    </tr>
    <tr>
      <td width="100%" height=22><table width="98%" height=20 border=0 align="center" cellpadding=0 cellspacing=0 bordercolor=#111111 id=AutoNumber3 style="BORDER-COLLAPSE: collapse">
          <tr>
            <td  width=20 height=20>&nbsp;</td>
            <td  width=160 bgcolor=#5298d1 height=20><b>&nbsp;</b><font style="FONT-SIZE: 10.5pt" color=#ffffff><b>您的联系资料</b></font></td>
            <td  height=20>&nbsp;</td>
          </tr>
      </table></td>
    </tr>
    <tr height=25>
      <td width="100%" valign="top"><table width="98%" height=121 border=0 align="center" cellspacing="1" cellpadding=2 bordercolor=#111111 bordercolorlight=#ffffff bordercolordark=#ffffff  id=AutoNumber1 style="BORDER-COLLAPSE: collapse">
          <tr class="tdbg">
            <td valign=top width="15%" height=25><p align=right>联 系 人：</p></td>
            <td valign=top width="34%" height=25><input class="textbox" size=21 name="ContactMan">
                <font color=#ff6600> *</font></td>
            <td valign=top width="16%" height=25><p align=right>联系电话：</p></td>
            <td valign=top width="35%" height=25><input class="textbox" size=21 name="Tel">
                <font color=#ff6600> *</font></td>
          </tr>
          <tr class="tdbg">
            <td valign=top width="15%" height=25><p align=right>公司名称：</p></td>
            <td valign=top width="34%" height=25><input class="textbox" size=21 name="CompanyName"></td>
            <td valign=top width="16%" height=25><p align=right>联系地址：</p></td>
            <td valign=top width="35%" height=25><input class="textbox" size=21 name="Address">
                <font color=#ff6600>&nbsp; </font></td>
          </tr>
          <tr class="tdbg">
            <td valign=top width="15%" height=25><p align=right>所在省份：</p></td>
            <td height=25 colspan="3" valign=top>
              <script language="JavaScript" src="<%=KS.GetDomain%>plus/area.asp" type="text/javascript"></script>
			  </td>
          </tr>
          <tr class="tdbg">
            <td valign=top width="15%" height=19><p align=right>电子邮件：</p></td>
            <td valign=top width="34%" height=19><input class="textbox" size=21 name="email">
                <font color=#ff6600>&nbsp; </font></td>
            <td valign=top width="16%" height=19><p align=right>邮政编码：</p></td>
            <td valign=top width="35%" height=19><input class="textbox" size=21 name="zip">
                <font color=#ff6600>&nbsp; </font></td>
          </tr>
          <tr class="tdbg">
            <td valign=top width="15%" height=19><p align=right>公司传真：</p></td>
            <td valign=top width="34%" height=19><input class="textbox" size=21 name="fax"></td>
            <td valign=top width="16%" height=19><p align=right>公司网址：</p></td>
            <td valign=top width="35%" height=19><input class="textbox" size=21 name="HomePage" value="http://"></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td align=middle width="100%" height=45><br>
          <input name="button" type=button onClick="check()" class="button" value=" 发 布 ">
        &nbsp;&nbsp;&nbsp;&nbsp;
          <input name="button" type=button class="button" value="重 填">
        <br>
        　      </td>
    </tr>
  </table>
</FORM></td>
    </tr>
  </table>
  <%
	End Sub
	
	'保存文章
	Sub SaveByArticle
	            Dim Title,FullTitle,KeyWords,Author,Origin,Intro,Content,Verific,PicUrl,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
                 ClassID=KS.S("ClassID")
				 Title=KS.LoseHtml(KS.S("Title"))
				 KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				 Author=KS.LoseHtml(KS.S("Author"))
				 Origin=KS.LoseHtml(KS.S("Origin"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				 if Content="" Then Content="&nbsp;"
				 Verific=KS.S("Status")
				 Intro  = KS.LoseHtml(KS.S("Intro"))
				 FullTitle = KS.LoseHtml(KS.S("FullTitle"))
				 if Intro="" And KS.ChkClng(KS.S("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(Content),200)
				 
				 Dim Fname,FnameType,TemplateID,WapTemplateID
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
				 
				 PicUrl=KS.S("PicUrl")
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				 
				Next
				End If
				 
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
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
				RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2),Conn,1,3
				RSObj.AddNew
				  RSObj("Title")=Title
				  RSObj("FullTitle")=FullTitle
				  RSObj("Tid")=ClassID
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Inputer")="游客"
				  RSObj("Origin")=Origin
				  RSObj("ArticleContent")=Content
				  RSObj("Verific")=0
				  RSObj("photoUrl")=PicUrl
				  RSObj("Intro")=Intro
				  if PicUrl<>"" Then 
				   RSObj("PicNews")=1
				  Else
				   RSObj("PicNews")=0
				  End if
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("Adddate")=Now
				  RSObj("Rank")="★★★"
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
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Intro,KeyWords,PicUrl,"游客",0,Fname)

				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByPhoto()
	            Dim Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,PicUrls,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
  				  ClassID=KS.S("ClassID")
				  Title=KS.LoseHtml(KS.S("Title"))
				  KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				  Author=KS.LoseHtml(KS.S("Author"))
				  Origin=KS.LoseHtml(KS.S("Origin"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				  PhotoUrl=KS.S("PhotoUrl")
				  PicUrls=KS.S("PicUrls")

				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If
				  
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
				 
				 Dim Fname,FnameType,TemplateID,WapTemplateID
				 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
				 RSC.Open "select top 1 TemplateID,FnameType,FsoType,WapTemplateID from KS_Class Where ID='" & ClassID & "'",conn,1,1
				 if RSC.Eof Then 
				  Response.end
				 Else
				 FnameType=RSC("FnameType")
				 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
				 TemplateID=RSC("TemplateID")
				 WapTemplateID=RSC("WapTemplateID")
				 End If
				 RSC.Close:Set RSC=Nothing
				  
				Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & "",Conn,1,3
				RSObj.AddNew
				  RSObj("PicID")=KS.GetInfoID(ChannelID)   '取图片的唯一ID
				  RSObj("Title")=Title
				  RSObj("Tid")=ClassID
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("PicUrls")=PicUrls
				  RSObj("KeyWords")=KeyWords
				  RSObj("Author")=Author
				  RSObj("Inputer")="游客"
				  RSObj("Origin")=Origin
				  RSObj("PictureContent")=Content
				  RSObj("Verific")=0
				  RSObj("Hits")=0
				  RSObj("TemplateID")=TemplateID
				  RSObj("WapTemplateID")=WapTemplateID
				  RSObj("Fname")=FName
				  RSObj("AddDate")=Now
				  If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
						 	If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.G(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=0 Then   '类型为日期与允许空时
							else
							 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
							end if
						Next
				  End If
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByDownLoad()
		Dim SizeUnit,ClassID,Title,KeyWords,Author,DownLB,DownYY,DownSQ,DownSize,DownPT,YSDZ,ZCDZ,JYMM,Origin,Content,Verific,PhotoUrl,DownUrls,RSObj,ID,DownID,AddDate,ComeUrl,CurrentOpStr,Action,I,UserDefineFieldArr,UserDefineFieldValueStr
				  ClassID=KS.S("ClassID")
				  Title=KS.LoseHtml(KS.S("Title"))
				  KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				  Author=KS.LoseHtml(KS.S("Author"))
				  DownLB=KS.LoseHtml(KS.S("DownLB"))
				  DownYY=KS.LoseHtml(KS.S("DownYY"))
				  DownSQ=KS.LoseHtml(KS.S("DownSQ"))
				  DownSize=KS.S("DownSize")
				  If DownSize = "" Or Not IsNumeric(DownSize) Then DownSize = 0
		           DownSize = DownSize & KS.S("SizeUnit")
				  DownPT=KS.LoseHtml(KS.S("DownPT"))
				  YSDZ=KS.LoseHtml(KS.S("YSDZ"))
				  ZCDZ=KS.LoseHtml(KS.S("ZCDZ"))
				  JYMM=KS.LoseHtml(KS.S("JYMM"))
				  Origin=KS.LoseHtml(KS.S("Origin"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				  PhotoUrl=KS.LoseHtml(KS.S("PhotoUrl"))
				  DownUrls=KS.S("DownUrls")
				  
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then Response.Write "<script>alert('你没有选择" & KS.C_S(ChannelID,3) & "栏目!');history.back();</script>":Exit Sub
				  If Title="" Then  Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "名称!');history.back();</script>":Exit Sub
	              If DownUrls="" Then Response.Write "<script>alert('你没有输入" & KS.C_S(ChannelID,3) & "!');history.back();</script>": Exit Sub
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				 Dim Fname,FnameType,TemplateID,WapTemplateID
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
					RSObj.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & "",Conn,1,3
					RSObj.AddNew
					  RSObj("Title")=Title
					  RSObj("TID")=ClassID
					  RSObj("KeyWords")=KeyWords
					  RSObj("Author")=Author
					  RSObj("DownLB")=DownLB
					  RSObj("DownYY")=DownYY
					  RSObj("DownSQ")=DownSQ
					  RSObj("DownSize")=DownSize
					  RSObj("DownPT")=DownPT
					  RSObj("YSDZ")=YSDZ
					  RSObj("ZCDZ")=ZCDZ
					  RSObj("JYMM")=JYMM
					  RSObj("Origin")=Origin
					  RSObj("DownContent")=Content
					  RSObj("PhotoUrl")=PhotoUrl
					  RSObj("DownUrls")="0|下载地址|" & DownUrls
					  RSObj("Inputer")="游客"
					  RSObj("Verific")=0
					  RSObj("Hits")=0
				      RSObj("TemplateID")=TemplateID
					  RSObj("WapTemplateID")=WapTemplateID
				      RSObj("Fname")=FName
					  RSObj("AddDate")=Now
					  RSObj("Rank")="★★★"
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
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
			
	End Sub
	
	Sub SaveByFlash
		Dim Title,KeyWords,Author,Origin,Content,Verific,PhotoUrl,FlashUrl,RSObj,I,UserDefineFieldArr,UserDefineFieldValueStr
  		ClassID=KS.S("ClassID")
		Title=KS.LoseHtml(KS.S("Title"))
		KeyWords=KS.LoseHtml(KS.S("KeyWords"))
		Author=KS.LoseHtml(KS.S("Author"))
		Origin=KS.LoseHtml(KS.S("Origin"))
		Content = Request.Form("Content")
		Content =KS.ClearBadChr(content)
		PhotoUrl=KS.S("PhotoUrl")
		FlashUrl=KS.S("FlashUrl")
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择"& KS.C_S(ChannelID,3) & "栏目!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "名称!');history.back();</script>"
				    Exit Sub
				  End IF
	              If FlashUrl="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "!');history.back();</script>"
				    Exit Sub
				  End IF
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If				  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
					 Dim Fname,FnameType,TemplateID
					 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
					 RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
					 if RSC.Eof Then 
					  Response.end
					 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 End If
					 RSC.Close:Set RSC=Nothing

					RSObj.Open "Select top 1 * From KS_Flash Where 1=0",Conn,1,3
				  RSObj.AddNew
				   RSObj("Hits")=0
				   RSObj("TemplateID")=TemplateID
				   RSObj("Fname")=FName
				   RSObj("AddDate")=Now
				   RSObj("Rank")="★★★"
				   RSObj("Title")=Title
				   RSObj("TID")=ClassID
				   RSObj("PhotoUrl")=PhotoUrl
				   RSObj("FlashUrl")=FlashUrl
				   RSObj("KeyWords")=KeyWords
				   RSObj("Author")=Author
				   RSObj("Inputer")="游客"
				   RSObj("Origin")=Origin
				   RSObj("FlashContent")=Content
				   RSObj("Verific")=0
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
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveByMovie()
		Dim Title,KeyWords,MovieAct,MovieDY,Content,Verific,PhotoUrl,MovieUrl,RSObj,I,UserDefineFieldArr,UserDefineFieldValueStr
				ClassID=KS.S("ClassID")
				Title=KS.LoseHtml(KS.S("Title"))
				KeyWords=KS.LoseHtml(KS.S("KeyWords"))
				MovieAct=KS.LoseHtml(KS.S("MovieAct"))
				MovieDY=KS.LoseHtml(KS.S("MovieDY"))
				 Content = Request.Form("Content")
				 Content=KS.ClearBadChr(content)
				PhotoUrl=KS.S("PhotoUrl")
				MovieUrl=KS.S("MovieUrl")
				  
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('你没有选择"& KS.C_S(ChannelID,3) & "栏目!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "名称!');history.back();</script>"
				    Exit Sub
				  End IF
	              If MovieUrl="" Then
				    Response.Write "<script>alert('你没有输入"& KS.C_S(ChannelID,3) & "!');history.back();</script>"
				    Exit Sub
				  End IF
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If				  
					 Dim Fname,FnameType,TemplateID
					 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
					 RSC.Open "select * from KS_Class Where ID='" & ClassID & "'",conn,1,1
					 if RSC.Eof Then 
					  Response.end
					 Else
					 FnameType=RSC("FnameType")
					 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
					 TemplateID=RSC("TemplateID")
					 End If
					 RSC.Close:Set RSC=Nothing
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select * From KS_Movie Where 1=0",Conn,1,3
				  RSObj.AddNew
				  RSObj("TemplateID")=TemplateID
				  RSObj("ServerID")=0
				  RSObj("Fname")=FName
				  RSObj("Hits")=0
				  RSObj("AddDate")=Now
				  RSObj("Rank")="★★★"
				  RSObj("Title")=Title
				  RSObj("TID")=ClassID
				  RSObj("PhotoUrl")=PhotoUrl
				  RSObj("MovieUrls")=MovieUrl
				  RSObj("KeyWords")=KeyWords
				  RSObj("MovieAct")=MovieAct
				  RSObj("Inputer")="游客"
				  RSObj("MovieDY")=MovieDY
				  RSObj("MovieContent")=Content
				  RSObj("Verific")=0
				  If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
						 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
						Next
				  End If
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID: InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				 Fname=RSOBj("Fname")

				 RSObj.Close:Set RSObj=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,"游客",0,Fname)
				Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"
	End Sub
	
	Sub SaveBySupply()
		Dim GQID,Title,Price,TypeID,ValidDate,GQContent,ContactMan,Tel,CompanyName,Address,Province,City,Email,Zip,Fax,HomePage,I,UserDefineFieldArr,UserDefineFieldValueStr,PhotoUrl,Visitor,KeyWords,Verific,inputer
			 ClassID      = KS.S("ClassID")
			 Title        = KS.LoseHtml(KS.S("Title"))
			 PhotoUrl     = KS.LoseHtml(KS.S("PhotoUrl"))
			 Price        = KS.LoseHtml(KS.S("Price"))
			 TypeID       = KS.S("TypeID")
			 ValidDate    = KS.S("ValidDate")
			 GQContent = Request.Form("GQContent")
			 GQContent=KS.ClearBadChr(GQContent)
			 ContactMan   = KS.LoseHtml(KS.S("ContactMan"))
			 Tel          = KS.LoseHtml(KS.S("Tel"))
			 CompanyName  = KS.LoseHtml(KS.S("CompanyName"))
			 Address      = KS.LoseHtml(KS.S("Address"))
			 Province     = KS.LoseHtml(KS.S("Province"))
			 City         = KS.LoseHtml(KS.S("City"))
			 Email        = KS.LoseHtml(KS.S("Email"))
			 Zip          = KS.LoseHtml(KS.S("Zip"))
			 Fax          = KS.LoseHtml(KS.S("Fax"))
			 HomePage     = KS.LoseHtml(KS.S("HomePage"))
			 KeyWords     = KS.LoseHtml(KS.S("KeyWords"))
				'自定义字段
				UserDefineFieldArr=KSUser.KS_D_F_Arr(8)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写数字!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的日期!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "必须填写正确的Email!');history.back();</script>":Exit Sub
				Next
				End If
				
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")		
		  RS.Open "Select * From [KS_Class] Where ID='" & ClassID & "'", conn, 1, 1
		  If RS.Eof And Rs.Bof Then
		   Response.Write "<script>alert('非法参数!');history.back();</script>"
		   response.end
		  end if
		  Dim TemplateID,GQFsoType,GQFnameType
		  TemplateID=RS("Templateid")
		  GQFsoType=RS("FsoType")
		  GQFnameType = Trim(RS("FnameType"))
		  RS.Close
		  Dim Fname:Fname=KS.GetFileName(GQFsoType, Now, GQFnameType)
		  RS.Open "select * from KS_GQ where 1=0", conn, 1, 3
		   RS.AddNew
		   RS("Hits")=0
		   RS("AddDate")=Now
		   RS("TemplateID")=TemplateID
		   RS("Fname")=Fname
		   RS("Recommend")=0
		   RS("IsTop")=0
		   IF Cbool(KSUser.UserLoginChecked)=false Then	inputer="游客" Else inputer=KS.C("UserName")
		   RS("Inputer")=inputer
		   RS("Tid")=ClassID
		   RS("Title")=Title
		   RS("Price")=Price
		   RS("PhotoUrl")=PhotoUrl
		   RS("TypeID")=TypeID
		   RS("ValidDate")=ValidDate
		   RS("GQContent")=GQContent
		   RS("KeyWords")=KeyWords
		   If KS.C_S(ChannelID,17)=1 Then Verific=0 Else Verific=1
		   RS("Verific")=verific
		   RS("ContactMan")=ContactMan
		   RS("Tel")=Tel
		   RS("CompanyName")=CompanyName
		   RS("Address")=Address
		   RS("Province")=Province
		   RS("City")=City
		   RS("Email")=Email
		   RS("Zip")=Zip
		   RS("Fax")=Fax
		   RS("Homepage")=Homepage
		   If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RSObj("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
							else
							 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
							end if
				Next
			End If
		   RS.Update
		   RS.MoveLast
				Dim InfoID: InfoID=RS("ID")
				If Left(Ucase(Fname),2)="ID" Then
					RS("Fname") = InfoID & GQFnameType
					RS.Update
				End If
				Fname=RS("Fname")

				 RS.Close:Set RS=Nothing
				 
				 Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,GQContent,KeyWords,PhotoUrl,inputer,verific,Fname)
				 
		 Response.Write "<script>if (confirm('" & KS.C_S(ChannelID,3) & "发表成功，继续添加吗?')){location.href='User_Contributor.asp?ChannelID=" & ChannelID & "&Action=Next&ClassID=" & ClassID &"';}else{top.location.href='../';}</script>"

	End Sub
End Class
%> 

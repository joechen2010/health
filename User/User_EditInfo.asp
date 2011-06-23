<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../API/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_EditInfo
KSCls.Kesion()
Set KSCls = Nothing

Class User_EditInfo
        Private KS,KSUser
		Private Sub Class_Initialize()
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
		%>
		<div class="tabs">						  
			<ul>
				<li <%If KS.S("Action")="" then response.write " class='select'"%>><a href="User_EditInfo.asp">基本信息</a></li>
				<li <%If KS.S("Action")="face" then response.write " class='select'"%>><a href="User_EditInfo.asp?Action=face">个人头像</a></li>
				<li<%If KS.S("Action")="ContactInfo" then response.write " class='select'"%>><a href="User_EditInfo.asp?Action=ContactInfo">修改详细资料</a></li>
				<li<%If KS.S("Action")="PassInfo" then response.write " class='select'"%>><a href="User_EditInfo.asp?Action=PassInfo">密码设置</a></li>
			</ul>
		</div>

		<%
		Select Case KS.S("Action")
		  case "face"
	       Call KSUser.InnerLocation("修改个人形象照片")
		   Call ChangeFace()
		  case "FaceSave"
		   Call FaceSave()
		  Case "ContactInfo"
	       Call KSUser.InnerLocation("修改详细信息")
		   Call ContactInfo()
		  Case "PassInfo"
	       Call KSUser.InnerLocation("修改密码")
		   Call PassInfo()
		  Case "PassSave"
		   Call PassSave()
		  Case "PassQuestionSave"
		   Call PassQuestionSave()
		  Case "BasicInfoSave"
		   Call BasicInfoSave()
		  Case "ContactInfoSave"
		   Call ContactInfoSave()
		  Case Else
	       Call KSUser.InnerLocation("修改基本信息")
		   Call EditBasicInfo()
		End Select
	   End Sub
	   
	   '基本信息
	   Sub EditBasicInfo()
		   %>
          <script>
	
       	 <!----检查用户名，电子邮箱结束-->
      function CheckForm() 
		{ 
			
			if (document.myform.RealName.value =="")
			{
			alert("请填写您的真实姓名！");
			document.myform.RealName.focus();
			return false;
			}
			if (document.myform.Sex.value =="")
			{
			alert("请选择您的性别！");
			document.myform.Sex.focus();
			return false;
			}
			if (document.myform.IDCard.value =="")
			{
			alert("请输入您的身份证号码！");
			document.myform.IDCard.focus();
			return false;
			}
			if (parseInt(document.myform.IDCard.value.length)!=15&&parseInt(document.myform.IDCard.value.length!=18))
			{
			alert("有效身份证号码必须是15位或18位！");
			document.myform.IDCard.focus();
			return false;
			}
		  return true;	
		}
    </script>
          
          <table  cellspacing="1" cellpadding="3"  width="98%" align="center" border="0">
					  <form action="User_EditInfo.asp?Action=BasicInfoSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">

                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 会员名称： </span><br>
                            用于登录会员中心的账号，不可修改。</td>
                            <td width="72%">&nbsp;
                                <input  class="textbox" type="text" name="username" size="30" value="<%=KSUser.username%>" disabled="disabled" /></td>
                          </tr>
                          
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 真实姓名：</span><br>
                            请务必填写真实姓名</td>
                            <td width="72%">&nbsp;&nbsp;
                              <input name="RealName" class="textbox" type="text" id="RealName" value="<%=KSUser.Realname%>" size="30" maxlength="50" />
                              <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold">性&nbsp;&nbsp;&nbsp; 别：</span><br></td>
                            <td width="72%">&nbsp;&nbsp;
                                <select name="Sex" id="Sex" style="width:110">
                                  <option value="">==请选择性别==</option>
                                  <option value="男" <%if KSUser.sex="男" then%> selected="selected" <%else%><%end if%>>男</option>
                                  <option value="女" <%if KSUser.sex="女" then%> selected="selected" <%else%><%end if%>>女</option>
                                </select>
                                <span style="color: red">* </span></td>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 身份证号：</span><br>
							有效身份证号码应该是15位或18位，请认真填写。</td>
                            <td width="72%">&nbsp;&nbsp;
                              <input  class="textbox" name="IDCard" type="text" id="IDCard" value="<%=KSUser.idcard%>" size="30" maxlength="50" />
                              <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold">  出生日期：</span><br>
                            请填写正确的出生日期，格式：0000-00-00</td>
                            <td width="72%">&nbsp;&nbsp;
                                <input name="Birthday" class="textbox" type="text" id="Birthday" value="<%=Split(KSUser.Birthday," ")(0)%>" size="30" maxlength="50" />
                                <span style="color: red">*</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">  邮箱地址：</span><br>
                            请填写正确的邮箱地址，如：service@kesion.com</td>
                            <td>&nbsp;&nbsp;
                                <input name="Email" class="textbox" type="text" id="Email" value="<%=KSUser.Email%>" size="30" maxlength="50" />
                                <span style="color: red">*</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td height="22"><span style="font-weight: bold">隐私设定：</span><br>开放后别人可以看到您的性别、Email、QQ等信息</td>
                            <td>&nbsp;
                              <input type="radio" <%if KSUser.Privacy="0" Then Response.Write "checked=""checked"""%> value="0" name="Privacy" />
                              公开全部信息(包括真实姓名/电话号码/生日等) <br />
                              &nbsp;
                              <input type="radio" value="1" name="Privacy" <%if KSUser.Privacy="1" Then Response.Write "checked=""checked"""%>/>
                              公开部分信息(只公开QQ/Email等网上联络的信息) <br />
                              &nbsp;
                              <input type="radio" value="2" name="Privacy" <%if KSUser.Privacy="2" Then Response.Write "checked=""checked"""%>/>
                              完全保密(别人只能查看你的昵称) </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold">个人签名：</span><br>写上你的个性签名，说明文字不超过1000个字符。</td>
                            <td width="72%">&nbsp;&nbsp;
                                <textarea name="Sign" class="textbox" cols="60" rows="5" id="Sign" style="width:300px; height:60px"><%= KSUser.Sign%></textarea></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input  class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
                          </tr>
		    </form>
            </table>
          <%
  End Sub
  
  Sub ChangeFace()
  %>
   <script type="text/javascript">
     function changeimage()
	 {
		  $("#UserFace").val("images/face/"+$("#Image").val()+".gif");
		  $("#imgIcon").attr("src",'<%=KS.Setting(3)%>Images/Face/'+$("#Image").val()+'.gif');
	 }
	 
	
   </script>
   <br/>
   <form action="User_EditInfo.asp?Action=FaceSave" method="post" name="myform" id="myform">
  <table  cellspacing="1" cellpadding="3"  width="80%" align="center" border="0">
   <tr class="tdbg">
                            <td colspan="2" height="22"><span style="font-weight: bold;color:green;font-size:14px"> 
							您可以为自己选择一个个性图片，
如果你填写了自定义头像部分，那么你的头像以自定义的为准。否则，请你留空自定义头像的所有栏目！</span></td>
</tr>
<tr>
                           
                            <td align="center">
							<div style="margin-left:20px" class="user_face">
							<div>当前头像</div>
							<%dim userfacesrc:userfacesrc=KSUser.UserFace
							  dim facewidth:facewidth=KSUser.FaceWidth
							  dim faceheight:faceheight=KSUser.FaceHeight
							 if KS.IsNul(userfacesrc) then userfacesrc="../Images/Face/1.gif"
							 if left(userfacesrc,1)<>"/" and lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
							 IF KS.ChkCLng(Facewidth)=0 then facewidth=60
						     if KS.chkclng(faceheight)=0 then faceheight=60
							%>
							<img title=点击选择头像 style="CURSOR: hand" onClick="window.open('selectface.asp?action=face','face','width=480,height=400,resizable=1,scrollbars=1')" 
            height="60" src="<%=userfacesrc%>" id="imgIcon" width="60" border=1  name=showimages> 
			 <br/>
			<SELECT onchange=changeimage(); size="1" name="Image" id="Image"> 
              <%dim i
			   for i=1 to 56 
			   response.write "<option value=" & i & ">" & i & ".gif</option>"
			   next
			   %>
			   </select>
			   <br/>
			   <a href="#" onClick="window.open('selectface.asp?action=face','face','width=480,height=400,resizable=1,scrollbars=1')"><font color=red>预览头像</font></a>
			   </td>
			   <td>
			   <br>
			   自定义头像地址：
			   <input class="textbox" name="UserFace" type="text" id="UserFace" value="<%=Replace(userfacesrc,"../","")%>" size="30" maxlength="50" /><br>
		  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?ChannelID=9999' frameborder="0" scrolling="No" align="center" width='100%' height='30'></iframe>只支持jpg、gif、png，小于50k，默认尺寸为48*48
		  <br>
		   宽度：<input name="FaceWidth" type="text" id="FaceWidth" value="<%=FaceWidth%>" size="8" maxlength="50" /> 0~150之间的整数<br>
		   高度：<input name="FaceHeight" type="text" id="FaceHeight" value="<%=FaceHeight%>" size="8" maxlength="50" /> 0~150之间的整数</p></div>                            </td>
                          </tr>
                          <tr class="tdbg">
						    <td></td>
                            <td><input  class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
                          </tr>
	</table>
	</form>
  <%
  End Sub
  
  '联系信息
  Sub ContactInfo()
  %>
          <table  cellspacing="1" cellpadding="3" width="98%" align="center" border="0">
					  <form action="User_EditInfo.asp?Action=ContactInfoSave" method="post" name="myform" id="myform">
					  <input type="hidden" value="<%=KS.S("ComeUrl")%>" name="comeurl">
						  <tr>
						    <td colspan="2">
							<% 
							Dim RSU:Set RSU=Server.CreateObject("ADODB.RECORDSET")
							RSU.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
							If RSU.Eof Then
							  RSU.Close:Set RSU=Nothing
							  Response.Write "<script>alert('非法参数！');history.back();</script>"
							  Response.End()
							End If
						  Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(KSUser.GroupID,"formid"))

						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(KSUser.GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr,FieldStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   For K=0 TO Ubound(SQL,2)
						     FieldStr=FieldStr & "|" & lcase(SQL(2,K))
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(SQL(2,K))="province&city" Then
								 InputStr=""
								 InputStr="<script src='../plus/area.asp'></script><script language=""javascript"">" &vbcrlf
								 If RSU("Province")<>"" And Not ISNull(RSU("Province")) Then
						         InputStr=InputStr & "$('#Province').val('" & RSU("province") &"');" &vbcrlf
								 End If
						         If RSU("City")<>"" And Not ISNull(RSU("City")) Then
								  InputStr=InputStr & "$('#City')[0].options[1]=new Option('" & RSU("City") & "','" & RSU("City") & "');" &Vbcrlf
								  InputStr=InputStr & "$('#City')[0].options(1).selected=true;" & vbcrlf
						         end if
						          InputStr=InputStr & "</script>" &vbcrlf
							  Else
							  Select Case SQL(1,K)
								Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & ";height:" & SQL(5,K) & "px"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" &RSU(SQL(2,K)) & "</textarea>"
								Case 3
								  InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(RSU(SQL(2,K)))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(RSU(SQL(2,K)))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(RSU(SQL(2,K))),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  Case 10
							        Dim H_Value:H_Value=RSU(SQL(2,K))
									If IsNull(H_Value) Then H_Value=" "
									InputStr=InputStr & "<input type=""hidden"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""" value="""& Server.HTMLEncode(H_Value) &""" style=""display:none"" /><input type=""hidden"" id=""" & SQL(2,K) &"___Config"" value="""" style=""display:none"" /><iframe id=""" & SQL(2,K) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & SQL(2,K) &"&amp;Toolbar=" & SQL(7,K) & """ width=""" &SQL(4,K) &""" height=""" & SQL(5,K) & """ frameborder=""0"" scrolling=""no""></iframe>"				
							  Case Else
								  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & """ name=""" & lcase(SQL(2,K)) & """ value=""" & RSU(SQL(2,K)) & """>"
							  End Select
							  End If
							  if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				              If Instr(Template,"{@NoDisplay(" & SQL(2,K) & ")}")<>0 Then
							   Template=Replace(Template,"{@NoDisplay(" & SQL(2,K) & ")}"," style='display:none'")
							  End If
							  Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)
							 End If
						   Next
							RSU.Close:Set RSU=Nothing
							
							
							Response.Write Template
							%>
							</td>
						  </tr>
                         
                          <tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input onClick="return(CheckForm())" class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
                          </tr>
		    </form>
            </table>
		
          <script type="text/javascript">
		   //检查日期
		   function CheckDT(str)     
		   {     
				 var r = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/);     
				 if(r==null)
				 {
					 return false;     
				 }
				 else
				 {
					var d= new Date(r[1], r[3]-1, r[4]);     
					return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]); 
				}    
			}
		  //检查电话
		  function CheckPhone(Str) 
			{ 
			   var i,j,strTemp;
			   Str=Str.replace('-','');
			   strTemp="0123456789";
				if (Str.length<10||Str.length>12)
				{
				return false;
				}
			 
			   for (i=0;i<Str.length;i++)
				{
				 j=strTemp.indexOf(Str.substring(i, i+1)); 
				 if (j==-1)
				  {
				   return false;
				  }
				}
			   return true;
			}
			//检查手机
			function CheckMobile(MobileStr) 
			{ 
			   var i,j,strTemp;
			   strTemp="0123456789";
			   var flags;
			   
			   if(MobileStr.substring(0,2)!="18"&&MobileStr.substring(0,2)!="13"&&MobileStr.substring(0,2)!="15"&&MobileStr.substring(0,1)!="0")
				{
				 return false;
				}
			   
			  
				if (MobileStr.length!=11)
				{
				return false;
				}
			   
			   for (i=0;i<MobileStr.length;i++)
				{
				 j=strTemp.indexOf(MobileStr.substring(i, i+1)); 
				 if (j==-1)
				  {
				   return false;
				  }
				}
			   return true;
			}


			
           //检查是否全数字
		   function CheckAllNum(str)
			{
			   var i,j,strTemp;
			   strTemp="0123456789";
			   for (i=0;i<str.length;i++)
				{
				 j=strTemp.indexOf(str.substring(i, i+1)); 
				 if (j==-1)
				  {
				   return false;
				  }
				}
			   return true;
			}
			//检查邮箱是否合法
			function emailCheck (emailStr) {
			var emailPat=/^(.+)@(.+)$/;
			var matchArray=emailStr.match(emailPat);
			if (matchArray==null) {
			 return false;
			}
			return true;
			}
            
			function CheckForm()
			{
			  var obj=document.myform;
			<%if instr(FieldStr,"birthday")<>0 then%>
			 if (CheckDT(obj.birthday.value)==false)
			 {
			  alert('出生日期格式不正确！格式应为yyyy-mm-dd');
			  obj.birthday.focus();
			  return false;
			 }
			<%end if
			if InStr(FieldStr,"officetel")<>0 then%>
			 if (obj.officetel.value!='' && CheckPhone(obj.officetel.value)==false)
			 {
			   alert('办公电话格式不正确！');
			   obj.officetel.focus();
			   return false;
			 }
			<%end if
			if InStr(FieldStr,"hometel")<>0 then%>
			 if (obj.hometel.value!='' && CheckPhone(obj.hometel.value)==false)
			 {
			   alert('电话号码格式不正确！');
			   obj.hometel.focus();
			   return false;
			 }
			<%end if
			if InStr(FieldStr,"fax")<>0 then%>
			 if (obj.fax.value!='' && CheckPhone(obj.fax.value)==false)
			 {
			   alert('传真号码格式不正确！');
			   obj.fax.focus();
			   return false;
			 }
			<%end if
			if InStr(FieldStr,"mobile")<>0 then%>
			 if (obj.mobile.value!='' && CheckMobile(obj.mobile.value)==false)
			 {
			   alert('手机号码格式不正确！');
			   obj.mobile.focus();
			   return false;
			 }
			<%end if

			if instr(FieldStr,"uc")<>0 then%>
			if (obj.uc.value!='' && (CheckAllNum(obj.uc.value)==false ||obj.uc.value.length<5))
			 {
			   alert('UC号码格式不正确，不能含有字符且不能少于5位！');
			   obj.uc.focus();
			   return false;
			 }
			<%
			end if
			if instr(FieldStr,"qq")<>0 then%>
			if (obj.qq.value!='' && (CheckAllNum(obj.qq.value)==false ||obj.qq.value.length<5))
			 {
			   alert('qq号码格式不正确，不能含有字符且不能少于5位！');
			   obj.qq.focus();
			   return false;
			 }
			<%
			end if
			if instr(FieldStr,"icq")<>0 then%>
			if (obj.icq.value!='' && (CheckAllNum(obj.icq.value)==false ||obj.icq.value.length<5))
			 {
			   alert('icq号码格式不正确，不能含有字符且不能少于5位！');
			   obj.icq.focus();
			   return false;
			 }
			<%
			end if
			if instr(FieldStr,"zip")<>0 then%>
			if (obj.zip.value!='' && (CheckAllNum(obj.zip.value)==false ||obj.zip.value.length<6))
			 {
			   alert('邮政编码格式不正确！');
			   obj.zip.focus();
			   return false;
			 }
			<%
			end if
			if instr(FieldStr,"msn")<>0 then%>
			if (obj.msn.value!='' && emailCheck(obj.msn.value)==false)
			 {
			   alert('MSN格式不正确！');
			   obj.msn.focus();
			   return false;
			 }
			<%
			end if
			%>
			}
		 </script>
		<%
		  
  End Sub
  
  '设置密码
  Sub PassInfo()
  		   %>
          <script>
	      function CheckForm() 
		{ 
			if (document.myform.oldpassword.value =="")
			{
			alert("请填写您的旧密码！");
			document.myform.oldpassword.focus();
			return false;
			}
			if (document.myform.newpassword.value =="")
			{
			alert("请输入您的新密码！");
			document.myform.newpassword.focus();
			return false;
			}
			if (parseInt(document.myform.newpassword.value.length)<6)
			{
			alert("密码长度必须大于等于6！");
			document.myform.newpassword.focus();
			return false;
			}
			if (document.myform.renewpassword.value =="")
			{
			alert("请输入您的新确认密码！");
			document.myform.renewpassword.focus();
			return false;
			}
			if (document.myform.newpassword.value !=document.myform.renewpassword.value)
			{
			alert("两次输入的密码不一致！");
			document.myform.renewpassword.focus();
			return false;
			}
          return true;			
		}
	      function CheckForm1() 
		{ 
			if (document.myform1.Password.value =="")
			{
			alert("请填写您的登录密码！");
			document.myform1.Password.focus();
			return false;
			}
			if (document.myform1.Question.value =="")
			{
			alert("请输入您的密码问题！");
			document.myform1.Question.focus();
			return false;
			}
			if (document.myform1.Answer.value =="")
			{
			alert("请输入您的问题答案！");
			document.myform1.Answer.focus();
			return false;
			}

          return true;			
		}
    </script>
          <table  cellspacing="1" cellpadding="3" class="border" width="98%" align="center" border="0">
					  <form action="User_EditInfo.asp?Action=PassSave" method="post" name="myform" id="myform" onSubmit="return CheckForm();">
                          <tr class="title">
                            <td height="22" colspan="2" align="center"> 修 改 密 码 </td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold">旧 密 码： </span><br>
                            您的旧登录密码，必须正确填写。</td>
                            <td width="72%">&nbsp;
                            <input name="oldpassword" class="textbox" type="password" id="oldpassword" size="30" maxlength="50" />
                            <span style="color: red">*</span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 新 密 码：</span><br>
							请输入您的新密码！</td>
                            <td width="72%">&nbsp;
                              <input name="newpassword" class="textbox" type="password" id="newpassword" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 确认密码：</span><br>
同上。</td>
                            <td width="72%">&nbsp;
                              <input name="renewpassword" class="textbox" type="password" id="renewpassword" size="30" maxlength="50" />
                              <span style="color: red">* </span></td>
                          </tr>
                          
						<tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input  class="button" name="Submit"  type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
                        </tr>
		    </form>
            </table>
          <br>
          <table  cellspacing="1" cellpadding="3" class="border" width="98%" align="center" border="0">
					  <form action="User_EditInfo.asp?Action=PassQuestionSave" method="post" name="myform1" id="myform1" onSubmit="return CheckForm1();">
                          <tr class="title">
                            <td height="22" colspan="2" align="center">更 改 找 回 密 码 设 置</td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 登录密码：</span><br>
同上。</td>
                            <td width="72%">&nbsp;
                              <input name="Password" class="textbox" type="password" id="Password" size="30" maxlength="50" />
                              <span style="color: red">* </span></td>
                          </tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 密码问题：</span><br>
                            当密码忘记时，取回密码的提示问题。</td>
                            <td width="72%">&nbsp;
                            <input name="Question" class="textbox" type="text" id="Question" value="<%=KSUser.Question%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
						</tr>
                          <tr class="tdbg">
                            <td width="28%" height="22"><span style="font-weight: bold"> 问题答案：</span><br>
                            当密码忘记时，取回密码提示问题的答案。</td>
                            <td width="72%">&nbsp;
                            <input name="Answer" class="textbox" type="text" id="Answer" value="<%=KSUser.Answer%>" size="30" maxlength="50" />
                            <span style="color: red">* </span></td>
						</tr>
                          
						<tr class="tdbg">
                            <td width="28%" height="30">&nbsp;</td>
                            <td width="72%"><input  class="button" name="Submit" type="submit"  value=" OK,修 改 " />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button" name="Submit2" type="reset" value=" 重 填 " />                            </td>
                        </tr>
		    </form>
            </table>
          <%
  End SUb
  
  Sub FaceSave()
		 Dim UserFace:UserFace=KS.S("UserFace")		 
		 Dim FaceWidth:FaceWidth=KS.S("FaceWidth")		 
		 Dim FaceHeight:FaceHeight=KS.S("FaceHeight")
		 if left(userface,1)="/" then userface=right(userface,len(userface)-1)
		 if left(lcase(userface),4)<>"http" then userface=KS.GetDomain & userface
				If Not IsNumeric(FaceWidth) Then
				  Response.Write "<script>alert('头像宽度必须是数字');history.back();</script>"
				  response.end
				 end if
				If Not IsNumeric(FaceHeight) Then
				  Response.Write "<script>alert('头像高度必须是数字');history.back();</script>"
				  response.end
				 end if
			 Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 RS.Close:Set RS=Nothing:Response.End
			  Else

				 RS("UserFace")=UserFace
				 RS("FaceWidth")=FaceWidth
				 RS("FaceHeight")=FaceHeight
		 		 RS.Update
				 Call KS.FileAssociation(1024,rs("UserID"),UserFace,1)
				 
				 RS.Close:Set RS=Nothing
				 
				 if left(UserFace,1)<>"/" and lcase(left(UserFace,4))<>"http" then UserFace="{$GetSiteUrl}" & UserFace
				 Call KSUser.AddLog(KSUser.UserName,"更换了自己的形象照片,<a href='" & UserFace & "' target='_blank'>查看</a>!",0)
				 Response.Write "<script>alert('恭喜,头像修改成功！');top.location.href='../user/';</script>"
				 Response.End()
			  End if
			

  End Sub
  
  Sub BasicInfoSave() 
				 Dim RealName:RealName=KS.S("RealName")
				 Dim Sex:Sex=KS.S("Sex")
				 Dim Birthday:Birthday=KS.S("Birthday")
				 Dim IDCard:IDCard=KS.S("IDCard")
				 Dim Sign:Sign=KS.S("Sign")	
				 Dim Privacy:Privacy=KS.S("Privacy")
				 If Not IsDate(Birthday) Then
				  Response.Write "<script>alert('出生日期格式有误!');history.back();</script>"
				  response.end
				 end if
				  Dim Email:Email=KS.S("Email")
				 if KS.IsValidEmail(Email)=false then
					 Response.Write("<script>alert('请输入正确的电子邮箱!');history.back();</script>")
					 Exit Sub
				 end if
				 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
				If EmailMultiRegTF=0 Then
					Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where UserName<>'" & KSUser.UserName & "' And Email='" & Email & "'")
					If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
						Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');history.back();</script>")
						Exit Sub
					End If
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
				 End If

				 
			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			Dim API_KS,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_KS = New API_Conformity
				API_KS.NodeValue "action","update",0,False
				API_KS.NodeValue "username",KSUser.UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
				Md5OLD = 0
				API_KS.NodeValue "syskey",SysKey,0,False
				API_KS.NodeValue "truename",RealName,1,False
				API_KS.NodeValue "gender",sex,0,False
				API_KS.SendHttpData
				If API_KS.Status = "1" Then
					Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
					Exit Sub
				End If
				Set API_KS = Nothing
			End If
			'-----------------------------------------------------------------

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 RS.Close:Set RS=Nothing:Response.End
			  Else
				 RS("RealName")=RealName
				 RS("Sex")=Sex
				 RS("Birthday")=Birthday
				 RS("IDCard")=IDCard
				 RS("Email")=Email
				 RS("Sign")=Sign
				 RS("Privacy")=Privacy
		 		 RS.Update
				 RS.Close:Set RS=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"修改了个人基本信息资料!",0)
				 Response.Write "<script>alert('会员基本信息资料修改成功！');location.href='user_main.asp';</script>"
				 Response.End()
			  End if
			
  End Sub
  
  
  '保存联系信息
  Sub ContactInfoSave()
         Dim SQL,K
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(KSUser.GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType From KS_Field Where ChannelID=101 and ShowOnUserForm=1 and FieldID In(" & KS.FilterIDs(FieldsList) & ")",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		  For K=0 To UBound(SQL,2)
			   
			   If SQL(1,K)="1" Then 
			     if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
				    Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');history.back();</script>"
				    Response.End()
				 elseif KS.S("province")="" or ks.s("city")="" then
				    Response.Write "<script>alert('地区必须选择!');history.back();</script>"
				    Response.End()
				 end if
			   End If

			   
			   
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "必须填写数字!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的日期!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
				Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的Email格式!');history.back();</script>"
				Response.End()
			   End If 
			 Next

  
		 Dim RealName:RealName=KS.S("RealName")
		 Dim Sex:Sex=KS.S("Sex")
		 Dim Birthday:Birthday=KS.S("Birthday")
		 Dim IDCard:IDCard=KS.S("IDCard")
		 Dim OfficeTel:OfficeTel=KS.S("OfficeTel")
		 Dim HomeTel:HomeTel=KS.S("HomeTel")
		 Dim Mobile:Mobile=KS.S("Mobile")
		 Dim Fax:Fax=KS.S("Fax")
		 Dim province:province=KS.S("province")
		 Dim city:city=KS.S("city")
		 Dim Address:Address=KS.S("Address")
		 Dim ZIP:ZIP=KS.S("ZIP")
		 Dim HomePage:HomePage=KS.S("HomePage")		 	 	 
		 Dim QQ:QQ=KS.S("QQ")		 
		 Dim ICQ:ICQ=KS.S("ICQ")		 
		 Dim MSN:MSN=KS.S("MSN")		 
		 Dim UC:UC=KS.S("UC")		 
		 Dim Sign:Sign=KS.S("Sign")	
		 Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
			
			'-----------------------------------------------------------------
			'系统整合
			'-----------------------------------------------------------------
			Dim API_KS,API_SaveCookie,SysKey
			If API_Enable Then
				Set API_KS = New API_Conformity
				API_KS.NodeValue "action","update",0,False
				API_KS.NodeValue "username",KSUser.UserName,1,False
				Md5OLD = 1
				SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
				Md5OLD = 0
				API_KS.NodeValue "syskey",SysKey,0,False
				API_KS.NodeValue "email",KSUser.Email,1,False
				API_KS.NodeValue "mobile",Mobile,1,False
				API_KS.NodeValue "homepage",homepage,1,False
				API_KS.NodeValue "address",Address,1,False
				API_KS.NodeValue "zipcode",zip,1,False
				API_KS.NodeValue "qq",qq,1,False
				API_KS.NodeValue "icq",icq,1,False
				API_KS.NodeValue "msn",msn,1,False
				API_KS.SendHttpData
				If API_KS.Status = "1" Then
					Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
					Exit Sub
				End If
				Set API_KS = Nothing
			End If
			 
              Dim RS,UpFiles
			  Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				 Response.End
			  Else
			     RS("Sex")=Sex
				 If BirthDay<>"" Then RS("Birthday")=Birthday
				 If Sign<>"" Then RS("Sign")=Sign
				 RS("RealName")=RealName
				 RS("IDCard")=IDCard
				 RS("Email")=KSUser.Email
				 RS("OfficeTel")=OfficeTel
				 RS("HomeTel")=HomeTel
				 RS("Mobile")=Mobile
				 RS("Fax")=Fax
				 RS("Province")=Province
				 RS("City")=City
				 RS("Address")=Address
				 RS("Zip")=Zip
				 RS("HomePage")=HomePage
				 RS("QQ")=QQ
				 RS("ICQ")=ICQ
				 RS("MSN")=MSN
				 RS("UC")=UC
				 RS("Privacy")=Privacy
				 '自定义字段
				 For K=0 To UBound(SQL,2)
				  If left(Lcase(SQL(0,K)),3)="ks_" Then
				   RS(SQL(0,K))=KS.S(SQL(0,K))
				   	If SQL(3,K)="9" or SQL(3,K)="10" Then
					   UpFiles=UpFiles & KS.S(SQL(0,K))
					End If
				  End If
				 Next
		 		 RS.Update
				 
				 Call KS.FileAssociation(1023,RS("UserID"),UpFiles,1)
				 
				 Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
				 If IsObject(FieldsXml) Then
				   	 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					If objNode.Attributes.item(0).Text<>"0" Then
					   If Not Conn.Execute("Select UserName From KS_EnterPrise Where UserName='" & KSUser.UserName & "'").Eof Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								Conn.Execute("UPDATE KS_EnterPrise Set " & objAtr.Attributes.item(0).Text & "='" & RS(objAtr.Attributes.item(1).Text) & "' Where UserName='" & KSUser.UserName & "'")
						 Next
					   End If
					End If
				 End If

				 
				 If KS.C_S(8,21)="1" Then
				  Conn.Execute("Update KS_GQ Set ContactMan='" & RealName &"',Tel='" &OfficeTel & "',Address='" & Address & "',Province='" & Province & "',City='" & City & "',Zip='" & Zip & "',Fax='" & Fax & "',Homepage='" & HomePage & "' where inputer='" & KSUser.UserName & "'")
				 End If
				 Call KSUser.AddLog(KSUser.UserName,"修改了个人详细信息资料!",0)
				 If KS.S("ComeUrl")<>"" Then
				 Response.Write "<script>alert('恭喜，详细信息修改成功！');location.href='" & KS.S("ComeURL") &"';</script>"
				 Else
				 Response.Write "<script>alert('恭喜，详细信息修改成功！');location.href='" & Request.ServerVariables("HTTP_REFERER") &"';</script>"
				 End If
				 Response.End()
			  End if
			RS.Close:Set RS=Nothing
  End Sub
  '保存密码设置
  Sub PassSave()
		     Dim Oldpassword:Oldpassword=KS.R(KS.S("Oldpassword"))
			 Dim NewPassWord:NewPassWord=KS.R(KS.S("NewPassWord"))
			 Dim ReNewPassWord:ReNewPassWord=KS.S("ReNewPassWord")
			 If Oldpassword = "" Then
				 Response.Write("<script>alert('请输入旧登录密码!');history.back();</script>")
				 Response.End
              End IF
			 If NewPassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Response.End
			 ElseIF ReNewPassWord="" Then
				 Response.Write("<script>alert('请输入确认密码');history.back();</script>")
				 Response.End
			 ElseIF NewPassWord<>ReNewPassWord Then
				 Response.Write("<script>alert('两次输入的密码不一致');history.back();</script>")
				 Response.End
			 End If
			 
			 OldPassWord =MD5(OldPassWord,16)
			 NewPassWord =MD5(NewPassWord,16)
			 
             Dim RS:Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select PassWord From KS_User Where UserName='" & KSUser.UserName & "' And PassWord='" & OldPassWord & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
			  	 Response.Write("<script>alert('您输入的旧密码有误！');history.back();</script>")
				 Response.End
			  Else
			  	'-----------------------------------------------------------------
				'系统整合
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","update",0,False
					API_KS.NodeValue "username",KSUser.UserName,1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.NodeValue "password",KS.R(KS.S("NewPassWord")),1,False
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------

			  
			     RS(0)=NewPassWord
				 RS.Update
				 Response.Cookies(KS.SiteSn)("PassWord") = NewPassWord
			  End if
			  
			  Call KSUser.AddLog(KSUser.UserName,"修改了个人登录密码!",0)
			 			RS.Close:Set RS=Nothing
  %>
          <table class="border" cellspacing="1" cellpadding="2" width="98%" align="center" border="0">
            <tbody>
			  <tr class="title">
			   <td height="25" align=center>密码修改成功</td>
		      </tr>
              <tr class="tdbg">
                <td height="42" align="center">您的会员登录密码修改成功！新密码 <font color="red"><%=KS.R(KS.S("NewPassWord"))%></font> 请牢记。 </td>
              </tr>
              <tr class="tdbg">
                <td height="42" align="center"><input type="button" onClick="location.href='user_main.asp'" class="button" value="进入会员首页">&nbsp;&nbsp;<input type="button" onClick="top.location.href='userlogout.asp'" value="退出重新登录" class="button"></td>
              </tr>
            </tbody>
          </table>
          <%
  End Sub
  '提示问题保存
  Sub PassQuestionSave()
				 Dim PassWord:PassWord=KS.S("PassWord")
				 Dim Question:Question=KS.S("Question")
				 Dim Answer:Answer=KS.S("Answer")
				
                 PassWord=MD5(PassWord,16)
              Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "' And PassWord='" & PassWord & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
				rs.close:set rs=nothing
				Response.Write "<script>alert('您输入的登录密码不正确!');history.back();</script>"
				Exit Sub
			  Else
			     RS("Question")=Question
				 RS("Answer")=Answer
		 		 RS.Update
				 RS.Close:Set RS=Nothing
				 Call KSUser.AddLog(KSUser.UserName,"修改了个人密码找回资料!",0)
				 Response.Write "<script>alert('你的密码找回资料修改成功！');location.href='user_main.asp';</script>"
				 Response.End()
			  End if
			
  End Sub
End Class
%> 

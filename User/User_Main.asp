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
Set KSCls = New UserMain
KSCls.Kesion()
Set KSCls = Nothing

Class UserMain
        Private KS,KSUser,TopDir
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
		Call KSUser.InnerLocation("会员信息")
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		TopDir=KSUser.GetUserFolder(ksuser.username)
		If Request.QueryString("action")="offline" then
		 Conn.Execute("Update KS_User Set isonline=0 where username='" & KSUser.UserName &"'")
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		ElseIf Request.QueryString("action")="online" Then
		 Conn.Execute("Update KS_User Set isonline=1 where username='" & KSUser.UserName &"'")
		 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End If
		%>
		 <script type="text/javascript">
						  function copyToClipboard(txt) {
							 if(window.clipboardData) {
									 window.clipboardData.clearData();
									 window.clipboardData.setData("Text", txt);
							 } else if(navigator.userAgent.indexOf("Opera") != -1) {
								  window.location = txt;
							 } else if (window.netscape) {
								  try {
									   netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
								  } catch (e) {
									   alert("被浏览器拒绝！\n请在浏览器地址栏输入'about:config'并回车\n然后将'signed.applets.codebase_principal_support'设置为'true'");
								  }
								  var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
								  if (!clip)
									   return;
								  var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
								  if (!trans)
									   return;
								  trans.addDataFlavor('text/unicode');
								  var str = new Object();
								  var len = new Object();
								  var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
								  var copytext = txt;
								  str.data = copytext;
								  trans.setTransferData("text/unicode",str,copytext.length*2);
								  var clipid = Components.interfaces.nsIClipboard;
								  if (!clip)
									   return false;
								  clip.setData(trans,null,clipid.kGlobalClipboard);
							 }
								  alert("复制成功！")
						}
		 </script>
		<style>
		  body{ line-height:180%; background:#fff; font-size:12px; color:#434343;}
		  #main .left{float:left;width:516px;}
		  #main .right{float:right;width:230px;margin-top:40px}
		  #main .userinfo{padding-top:10px;}
		  #main .userborder{padding:10px;}
          #main .dt td{height:30px;padding-top:15px}
		  #main .dt a{color:#2c602f}
		  
		  .abs{left:0px;width:120px}
		  .clear{clear:both}
		  .visitor{height:450px;padding-top:5px;padding-left:10px;}
		  .visitor li a.b{border:1px solid #ccc;padding:1px}
		  .visitor li{float:left;width:100px;text-align:center;}
		  .visitor li img{width:60px;height:60px}
		  
		  #fenye span{display:none}
		</style>
		
		<div id="main">
			 <div class="left">
					<div class="userinfo">
					  <h1><img src="images/user.gif" align="absmiddle" />个人资料(<a href="user_editInfo.asp">修改</a>)</h1>
					  <div class="userborder">
					   <table width="100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td class="splittd" height="25" width="50%">姓名：<%=KSUser.RealName%>(用户:<%=KSUser.UserName%> 当前：<span class="rl" onMouseover="$('#userstatus').show();"><%if KSUser.IsOnline="1" then response.write "在线" else response.write "隐身"%><img src="images/dico.gif" align="absmiddle">)
										<div id="userstatus" class='abs' onMouseOut="$('#userstatus').hide();">
										<dl><a href="?action=offline">设置隐身离线</a></dl>
										<dl><a href="?action=online">设置在线状态</a></dl>
										</div></span></td>
										<td class="splittd" width="50%">性别：<%=KSUser.Sex%></td>
									  </tr>
									  <tr>
										<td class="splittd" height="25" width="50%">类别：<%=KS.GetUserGroupName(KSUser.GroupID)%></td>
										<td class="splittd" width="50%">注册时间：<%=KSUser.RegDate%></td>
									  </tr>
									  <tr>
										<td class="splittd" height="25">计费方式：
										<%if KSUser.ChargeType=1 Then 
										  Response.Write "扣点数</font>"
										  ElseIf KSUser.ChargeType=2 Then
										   Response.Write "有效期</font>,到期时间：" & cdate(KSUser.BeginDate)+KSUser.Edays 
										  Else
										   Response.Write "无限期</font>"
										  End If
										  %>
				  </td>
										<td class="splittd">登录次数：<%=KSUser.LoginTimes%> 次</td>
									  </tr>
                                     <tr><td colspan='2'>
									<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
										<tr>
										  <td height="25"><strong>您的空间上限容量为 <font color=red><%=round(KSUser.SpaceSize/1024,2)%>M</font></strong> <span id="Sms_txt"></span> &nbsp;&nbsp;<a href="User_Files.asp?action=show">查看</a></td>
										  <td style="display:none"><img src="images/bar.gif" width="0" height="16" id="Sms_bar" align="absmiddle" /></td>
										</tr>
									  </table>
                 <%
                    response.write showtable("Sms_bar","Sms_txt",KS.GetFolderSize(TopDir)/1024,KSUser.SpaceSize)
                   %>	</td>
									  </tr>									  
									  
					</table>
					</div>
					
					<table style="margin-top:9px" border="0" width="100%">
					<tr>
					 <td>
					<img src="images/money.gif" align="absmiddle" /> <strong>我的财富</strong> <a href="user_payonline.asp" target="main"><img src="images/cz.gif" border="0"></a>  <a href="user_logmoney.asp" target="main">财务明细</a> <%if KS.C_S(5,21)=1 Then%><a href="user_order.asp" target="main">商城订单</a><%end if%>
					</td>
					<td align="right">
					          <span class="rl" onMouseOver="$('#sy').show();">
								 <%if KS.Setting(140)="1" Or KS.Setting(143)="1" then%>
								<span style="color:#ff6600;font-weight:bold;cursor:pointer">邀请好友,获得积分<img src="images/dico.gif" align="absmiddle"></span>
										 <div id="sy" onMouseOut="$('#sy').hide();" class="abs" style="width:420px;left:-45px">
										  <table border="0" align="center" style="width:420px;">

											  <%if KS.Setting(140)="1" Then%>
												<tr>
												  <td class="splittd" style="padding-left:10px;color:red">
								
												<strong>将本站推荐给朋友将获得积分：</strong>
											   <br /><span id="copytext"><%=Replace(KS.Setting(142),"{$UID}",KSUser.UserName)%></span>
													  <input name="button" type="button" onClick="copyToClipboard(document.getElementById('copytext').innerHTML);" value=" 复 制 " class="Button">
													  <div><font color=green>奖励说明：成功推荐一个访问者,您就可以增加 <font color=red><%=KS.Setting(141)%></font> 个积分。赶快行动吧！</font></div>
												</td>
											   </tr>
											  <%end if%>
											  <%if KS.Setting(143)="1" Then%>
											   <tr>
												<td class="splittd" style="padding-left:10px">
												  <font color=red><strong>引导朋友注册将获得积分：</strong>
												   <br />
												   <span id="copytext1"><%=KS.GetDomain%>user/reg/?uid=<%=KSUser.UserName%></span></font>
													  <input name="button2" type="button" onClick="copyToClipboard(document.getElementById('copytext1').innerHTML+'\n<%=Replace(KS.Setting(145),"'","\'")%>');" class="Button" value=" 复 制 ">
													  <div><font color=green>奖励说明：成功推荐一个用户注册,您就可以增加 <font color=red><%=KS.Setting(144)%></font> 个积分。赶快行动吧！</font></div>
														 </td>
													   </tr>
											   <%end if%>
											   </table>
										 </div>
										  <%end if%>
		  
									</span>	
					
					</td>
					</tr>
					</table>
					
					  <div class="userborder" >
					   <table width="100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td height="25"  nowrap>当前可用资金<font color="green"><%=formatnumber(KSUser.Money,2,-1)%></font>元  ,可用<%=KS.Setting(45) & "&nbsp;<font color=green>" & formatnumber(KSUser.Point,0,-1) & "</font>" & KS.Setting(46)%>,累计积分<font color="green"><%=KSUser.Score%></font>分
			                          <%
									   if KS.ChkClng(KSUser.UserCardID)<>0 then
									      Dim RSCard,ValidUnit,ExpireGroupID,ExpireTips
										  Set RSCard=Conn.Execute("Select top 1 * From KS_UserCard Where ID=" & KSUser.UserCardID)
										  If Not RSCard.Eof Then
											 ValidUnit=RSCard("ValidUnit")
											 ExpireGroupID=RSCard("ExpireGroupID")
											 If ValidUnit=1 Then                      '点券
											   If KSUser.Point<=10 And ExpireGroupID<>0 Then
											    ExpireTips="您的" & KS.Setting(45) & "快使用完毕了"
											   End If
											 ElseIf ValidUnit=2 Then                   '有效天数
											   If KSUser.Edays<=7 And ExpireGroupID<>0 Then
											    ExpireTips="您还有" & KSUser.Edays & "天就过期了"
											   End If 
											 ElseIf ValidUnit=3 Then                  '资金
											   If KSUser.Money<=10 And ExpireGroupID<>0 Then
												 ExpireTips="您的账户资金快使用完毕了"
											   End If
											 End If
										  End If
										  RSCard.Close : Set RSCard=Nothing
										  If ExpireTips<>"" and ExpireGroupID<>0  then
										  response.write "<br/><span style='color:red'>温馨提示：您上一次使用充值卡充值，" & ExpireTips & "，<br/>过期后您将自动转为<font color='blue'>"  & KS.U_G(ExpireGroupID,"groupname") & "</font>，为了更好的服务请尽量充值！</span>"
										  end if
									   end if
									  %>   
									  </tr>
									  
					   </table>
					  </div>
					</div>
					<div class="clear"></div>
					
					
					<div class="tabs" style="width:516px!important;width:516px">
						<ul>
						<li<%if KS.S("F")="" Then KS.Echo " class=""select"""%>><a href="?">个人动态</a></li>
						<li<%if KS.S("F")="f" Then KS.Echo " class=""select"""%>><a href="?f=f">好友动态</a></li>
						<li<%if KS.S("F")="d" Then KS.Echo " class=""select"""%>><a href="?f=d">大家都在做什么</a></li>
						</ul>
					</div>
					
					<table border="0" width="100%" class="dt">
					<%
					Dim Param,sqlstr,rs,Totalput,currentpage,MaxPerpage,xml,node
					MaxPerPage=10
					CurrentPage=KS.ChkClng(KS.S("Page"))
					If CurrentPage=0 Then CurrentPage=1
					
					if KS.S("F")="" Then 
					 Param=" where l.username='" & KSUser.UserName & "'"
					elseif KS.S("F")="f" Then
					 Param=" inner join ks_friend f on l.username=f.friend where f.username='" & KSUser.UserName & "' and f.accepted=1 and f.ShieldDT=0"
					End If
					Sqlstr="select top 10000 l.* from ks_userlog l" & param & " order by l.id desc"
					'response.write sqlstr
					             Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open sqlstr,conn,1,1
								 If RS.EOF And RS.BOF Then
								  RS.Close:SET RS=Nothing
								  KS.Echo "<tr><td class='splittd'>没有记录!</td></tr>"
								 Else
									totalPut = Conn.Execute("Select Count(l.ID) From KS_UserLog l " & Param)(0)
									If CurrentPage < 1 Then	CurrentPage = 1
									
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","root")
									RS.Close:SET RS=Nothing
							    End If
								If IsObject(XML) Then
								  For Each Node In XML.DocumentElement.SelectNodes("row")
								    KS.Echo "<tr><td class='splittd'>"
									KS.Echo " <span style=""float:right;"">" 
									If DateDiff("h",Node.SelectSingleNode("@adddate").text,now)>=12 Then
									KS.echo Node.SelectSingleNode("@adddate").text
									Else
									KS.Echo KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text)
									End If
									KS.Echo "</span>"
									KS.Echo "<img src='../images/user/log/" & Node.SelectSingleNode("@ico").text & ".gif' align='absmiddle'><a href='../space/?" & Node.SelectSingleNode("@username").text & "' target='_blank'>" & Node.SelectSingleNode("@username").text & "</a>"
									KS.Echo " " & Replace(Replace(Replace(Node.SelectSingleNode("@note").text,"{$GetSiteUrl}",KS.GetDomain),"<p>",""),"</p>","") & ""
									KS.Echo "</td></tr>"
								  Next
								End If
					           XML=Empty : Set Node=Nothing
					%>
					    
						</table><%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
					           <div class='splittd clear'>
								  </div>
					
					
					
			</div>
			<div class="right">
			   
			   <div class="tbox">
		        <div class="t_head">活动公告</div>
				 <%
				  Dim KSFObj:Set KSFobj=New refreshFunction
				  KS.Echo KSFObj.getLabel("{Tag:GetAnnounceList labelid=""0"" announcetype=""2"" owidth=""450"" oheight=""400"" width=""225"" height=""100"" speed=""1"" showstyle=""1"" opentype=""1"" listnumber=""10"" titlelen=""30"" showauthor=""0"" contentlen=""100"" navtype=""1"" nav=""../images/arraw01.gif"" titlecss="""" channelid=""9990"" ajaxout=""false""}{/Tag}")
				  Set KSFobj=Nothing
				 %>
				
				  
			   </div>
			   <br>
			   <div class="tbox">
		        <div class="t_head">最近谁来看过我</div>
				<div class="visitor">
				<%
				Dim user_face,Visitors
				Set RS=Conn.Execute("Select top 10 b.sex,a.Visitors,b.userface,a.adddate,b.isonline from KS_BlogVisitor a inner join KS_User b on a.Visitors=b.username where a.username='" & KSUser.UserName & "' order by a.adddate desc ,id desc")
				If Not RS.Eof Then Set XML=KS.RsToXml(Rs,"row","")
				RS.Close:Set RS=Nothing
			    If IsObject(XML) Then
				  For Each Node In XML.DocumentElement.SelectNodes("row") 
				    user_face=Node.SelectSingleNode("@userface").text
					Visitors =Node.SelectSingleNode("@visitors").text
					If user_face="" or isnull(user_face) then 
					 if Node.SelectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
					End If
			        If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
			         KS.Echo "<li><a class='b' href='../space?" & Visitors & "' target='_blank'><img src='" & User_face & "' border='0'></a><br/><a href='../space?" & Visitors & "' target='_blank'>" & Visitors & "</a><br />状态:"
					 If Node.SelectSingleNode("@isonline").Text="1" Then KS.Echo "<font color=red>在线</font>" Else KS.Echo "离线"
					 KS.Echo "</li>"
				  Next
				  XML=Empty : Set Node=Nothing
				Else
				    KS.Echo "没有访问记录,要加油哦^_^!"
				End If
				%>
				</div>
			   </div>
			
			</div>
		<div class="clear"></div>
		
		 
		 
		  		</div>

		<%
  End Sub
  
	   '（图片对象名称，标题对象名称，更新数，总数）
		Function ShowTable(SrcName,TxtName,str,c)
		Dim Tempstr,Src_js,Txt_js,TempPercent,SrcWidth
		If C = 0 Then C = 99999999
		Tempstr = str/C
		TempPercent = FormatPercent(tempstr,0,-1)
		Src_js = "document.getElementById(""" + SrcName + """)"
		Txt_js = "document.getElementById(""" + TxtName + """)"
			ShowTable = VbCrLf + "<script>"
			SrcWidth=FormatNumber(tempstr*300,0,-1) : If SrcWidth>500 Then SrcWidth="100%"
			ShowTable = ShowTable + Src_js + ".width=""" & SrcWidth & """;"
			ShowTable = ShowTable + Src_js + ".title=""容量上限为："&c/1024&" MB，已用（"&FormatNumber(str/1024,2)&"）MB！"";"
			ShowTable = ShowTable + Txt_js + ".innerHTML="""
			If FormatNumber(tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable + "已使用:" & TempPercent & """;"
			ElseIf FormatNumber(tempstr*100,0,-1)>100 Then
				ShowTable = ShowTable + "<font color=\""red\"">可用空间已使用完毕,请赶快清理！</font>"";"
			Else
				ShowTable = ShowTable + "<font color=\""red\"">已使用:" & TempPercent & ",请赶快清理！</font>"";"
			End If
			ShowTable = ShowTable + "</script>"
		End Function
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Ask_Setting
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Ask_Setting
        Private KS,KSCls
		Private maxperpage,totalrec,Pcount,pagelinks,showmode,pagenow,count,AskInstalDir
		Private m_intOrder,m_strOrder,SQLQuery,SQLField,Topiclist
		Private topicid,classid,topicmode,PostNum,ExpiredTime,CommentNum,HeadTitle,TopicUseTable
		Private classarr,cid,child,Catelist
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		  AskInstalDir="../" & KS.Asetting(1)
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		%>
		<html>
		<head>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<script src="../KS_Inc/common.js" language="JavaScript"></script>
		</head>
		<body>
        <div class='topdashed sort'>问答列表管理</div>
		<%
		    pagenow=KS.ChkClng(Request("page"))
			If pagenow=0 Then pagenow=1
			Dim Action
			Action = LCase(Request("action"))
			Select Case Trim(Action)
			Case "save"
				Call saveAsked()
			Case "asked"
			     If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
				Call showAsked()
			Case "del"
			      If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
				Call delTopic()
			Case "delask"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
				Call delAsked()
			Case "recommend"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
			    Call Recommend()
			Case "unrecommend"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
			    Call UnRecommend()
			Case "verify"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
			    Call Verify()
			Case "unverify"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
			    Call unVerify()
			Case "setsatis"
			    If Not KS.ReturnPowerResult(0, "WDXT10001") Then          '检查是权限
					 Call KS.ReturnErr(1, "")
					 .End
				 End If
			    Call SetSatis()
			Case Else
				Call showmain()
			End Select
	   End Sub

		Sub showmain()
			Dim i
			maxperpage=20
			showmode=KS.ChkClng(Request("showmode"))
			m_intOrder=KS.ChkClng(Request("order"))
			count=KS.ChkClng(Request("count"))
			classid=KS.ChkClng(Request("classid"))
			 Call GetChildList()
		
		%>
		<div style='height:30px;line-height:30px;margin:5px 1px;border:1px solid #cccccc;background:#F1FAFE'>
		<table border='0' width='100%'><tr><td width='15%'>
		<%
		If IsArray(classarr) Then
			Dim K,J,N
			N=0
			 For k=0 To Ubound(classarr,2)
			    Response.Write "<tr>"
			    For J=1 To 5
			     Response.Write "<td width='15%'><img src='images/folder/folderopen.gif' align='absmiddle'><a href=""?classid=" & classarr(0,n) & """>" & classarr(1,n) & "(" & classarr(2,n)+classarr(3,n) & ")</a></td>"
				 n=n+1
				 If N>Ubound(classarr,2) Then Exit For
				Next
				Response.Write "</tr>"
			  If N>Ubound(classarr,2) Then Exit For
			Next
	  End If	
		
		%>
		
		</tr></table></div>
		
		<div style="margin-top:5px;height:25px;line-height:25px">
		<b>查看：</b> <a href="KS.AskList.asp"><font color=#999999>全部</font></a> - <a href="?showmode=1"><font color=#999999>待解决</font></a> - <a href="?showmode=2"><font color=#999999>已解决</font></a> - <a href="?showmode=3"><font color=#999999>有悬赏</font></a> - <a href="?showmode=4"><font color=#999999>未审核</font></a> - <a href="?showmode=5"><font color=#999999>已审核</font></a> <b>排序方式:</b>
				  <select name="orders" onChange="location.href='?orders='+this.value">
				  <option value="">--选择排序方式--</option>
				  <option value="TopicID Desc"<%if KS.G("orders")="TopicID Desc" Then response.write " selected"%>>最新提问</option>
				  <option value="LastPostTime Desc,TopicID Desc"<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then response.write " selected"%>>最新回答</option>
				  <option value="Hits Desc,TopicID Desc"<%if KS.G("orders")="Hits Desc,TopicID Desc" Then response.write " selected"%>>浏览次数最多</option>
				  <option value="Reward Desc,TopicID Desc"<%if KS.G("orders")="Reward Desc,TopicID Desc" Then response.write " selected"%>>悬赏分最高</option>
				  </select>
		</div>
		<table  border="0" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0" width="100%">
		<tr class="sort">
			<td width="5%" noWrap="noWrap">选择</td>
			<td width="56%">标题</td>
			<td width="12%" noWrap="noWrap">用户名</td>
			<td width="6%" noWrap="noWrap">状态</td>
			<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then%>
			<td width="8%" noWrap="noWrap">回答日期</td>
			<%else%>
			<td width="8%" noWrap="noWrap">发布日期</td>
			<%end if%>
			<td width="4%" noWrap="noWrap">浏览</td>
			<td width="9%" noWrap="noWrap">管理操作</td>
		</tr>
		
		<form name="myform" id="myform" method="post" action="?">
		<input type="hidden" name="action" value="del">
		<%
			Call showTopiclist()
			If Not IsArray(Topiclist) Then
			  Response.Write "<tr><td class='splittd' colspan=6 align='center'>对不起, 找不到相关问题!</td></tr>"
			Else
				For i=0 To Ubound(Topiclist,2)
					If Not Response.IsClientConnected Then Response.End

		%>
		<tr align="center" onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'" id='u<%=Topiclist(0,i)%>'>
			<td class="splittd"><input type="checkbox" name="topicid" id='c<%=Topiclist(0,i)%>' value="<%=Topiclist(0,i)%>"/></td>
			<td class="splittd" align="left">[<a href="<%=AskInstalDir%>showlist.asp?id=<%=Topiclist(1,i)%>" target="_blank"><%=Topiclist(3,i)%></a>]
			<a href="<%=AskInstalDir%>q.asp?id=<%=Topiclist(0,i)%>" target="_blank"><%=Trim(Topiclist(4,i))%></a>
			<%
			 If Topiclist(5,i)>0 then
			  response.write "<img src=" & AskInstalDir & "images/ask_xs.gif>" & TopicList(5,i) & "分"
			 end if
			 
			 If TopicList(16,i)=1 Then
			  Response.Write " <span style='color:red'>荐</span>"
			 End If
			 If TopicList(11,i)=1 Then
			  Response.Write " <span style='color:green'>未审</font>"
			 End If
			%>
			
			</td>
			<td class="splittd" noWrap="noWrap"><%=Topiclist(2,i)%></td>
			<td class="splittd" noWrap="noWrap"><a target="_blank" href="<%=AskInstalDir%>q.asp?id=<%=Topiclist(0,i)%>"><img src="<%=askInstalDir%>images/ask<%=Topiclist(13,i)%>.gif" border="0"/></a></td>
			<%if KS.G("orders")="LastPostTime Desc,TopicID Desc" Then%>
			<td class="splittd" noWrap="noWrap"><%=formatdatetime(Topiclist(10,i),2)%></td>
			<%else%>
			<td class="splittd" noWrap="noWrap"><%=formatdatetime(Topiclist(9,i),2)%></td>
			<%end if%>
			<td class="splittd" noWrap="noWrap"><%=Topiclist(17,i)%></td>
			<td class="splittd" noWrap="noWrap"><a href="?action=asked&topicid=<%=Topiclist(0,i)%>">编辑</a> | <a href="?action=del&topicid=<%=Topiclist(0,i)%>" onClick="return confirm('删除后将不能恢复，您确定要删除吗?')">删除</a></td>
		</tr>
		<%
				Next
			End If
			Topiclist=Null
		%>
		<tr>
			<td colspan="10">
			&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">全选
		
				<input class="button" type="submit" name="submit_button1" value="批量删除" onClick="$('action').value='del';return confirm('您确定执行该操作吗?');">
				<input type="submit" value="审核" class="button" onClick="$('action').value='verify';return(confirm('确定批量审核吗?'));">
				
				<input type="submit" value="取消审核" class="button" onClick="$('action').value='unverify';return(confirm('确定批量取消审核吗?'));">
				
				<input type="submit" value="推荐" class="button" onClick="$('action').value='recommend';return(confirm('将问题设置为推荐将给会员增加相应的积分,确定设置吗?'));">
				
				<input type="submit" value="取消推荐" class="button" onClick="$('action').value='unrecommend';return(confirm('为保护会员权益,取消推荐将不再扣除原设置推荐所得会员积分,确定设置吗?'));">
			</td>
		</tr>
		</form>
		<tr>
			<td  align="right" colspan="10" id="NextPageText">
			<%
			Call KSCLS.ShowPage(totalrec, MaxPerPage, "KS.AskList.asp", True, "个", pagenow, KS.QueryParam("page"))
			%>
			</td>
		</tr>
		<tr> 
		   <td colspan="10">
		    <div>
			<form action="KS.AskList.asp" name="myform" method="get">
			   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
				  &nbsp;<strong>快速搜索=></strong>
				 &nbsp;关键字:<input type="text" class='textbox' name="keyword">&nbsp;分类:
				 <%
				 Dim SQL,Rs
	Response.Write " <select name=""class"">"
	Response.Write "<option value="""">所有分类</option>"
	SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
	Set Rs = Conn.Execute(SQL)
	Do While Not Rs.EOF
		Response.Write "<option value=""" & Rs("classid") & """ "
		If Request("editid") <> "" And CLng(Request("editid")) = Rs("classid") Then Response.Write "selected"
		Response.Write ">"
		If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
		If Rs("depth") > 1 Then
			For i = 2 To Rs("depth")
				Response.Write "&nbsp;&nbsp;│"
			Next
			Response.Write "&nbsp;&nbsp;├ "
		End If
		Response.Write Rs("ClassName") & "</option>" & vbCrLf
		Rs.movenext
	Loop
	Rs.Close
	Response.Write "</select>"
	Set Rs = Nothing
%>
				  &nbsp;
				  问题状态:<select name="showmode">
				  <option value="0">全部</option>
				  <option value="1">待解决</option>
				  <option value="2">待解决</option>
				  <option value="3">有悬赏</option>
				  <option value="4">未审核</option>
				  <option value="5">已审核</option>
				  </select>
				  
				  排序方式
				  <select name="orders">
				  <option value="TopicID Desc">最新提问</option>
				  <option value="LastPostTime Desc,TopicID Desc">最新回答</option>
				  <option value="Hits Desc,TopicID Desc">浏览次数最多</option>
				  <option value="Reward Desc,TopicID Desc">悬赏分最高</option>
				  </select>
				  <div style="padding-left:83px">
				  提问者:<input type="text" name="askName" class="Textbox" size="12">
				  回答者:<input type="text" name="answerName" class="Textbox" size="12">
				  提问时间:从
		      <input name="StartDate" type="text" id="StartDate" value="<%=request("StartDate")%>" readonly style="width:12%">
		  <b><a href="#" onClick="OpenThenSetValue('include/DateDialog.asp',160,170,window,document.all.StartDate);document.all.StartDate.focus();"><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></a></b>
		      到
		        <input name="EndDate" type="text" id="EndDate"  value="<%=request("endDate")%>" readonly style="width:12%">
		       <b><a href="#" onClick="OpenThenSetValue('include/DateDialog.asp',160,170,window,document.all.EndDate);document.all.EndDate.focus();"><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></a></b> 
				  <input type="submit" value="开始搜索" class="button" name="s1">
				  </div>
				</div>
			</form>
			</div>

		   </td>
		</tr>
		<tr>
			<td colspan="6" style="border:1px solid #f1f1f1;line-height:21px;padding-left:10px">
			 <font color=red><strong>操作说明:</strong></font><br />
			 1.将问题设置为推荐将给会员增加相应的积分,会员所得积分在"问答参数设置"里设定<br />
			 2.为保护会员权益,取消推荐将不再扣除原设置推荐所得会员积分,一般建议一旦设置为推荐后就不要再取消推荐<br />
			 3.如果您将问题推荐后,然后取消推荐,又重新推荐可能导致多次给会员增加积分
			</td>
		</tr>
		
		
		</table>
		<%
		End Sub
		
		Sub showTopiclist()
			Dim Rs,SQL,Cmd,Param,OrderStr
			SQLField="TopicID,classid,UserName,classname,title,reward,Expired,Closed,PostTable,DateAndTime,LastPostTime,LockTopic,PostNum,TopicMode,Anonymous,IsTop,recommend,Hits"
			Param=" where 1=1"
			Select Case showmode
			 case 1 param=param & " and topicmode=0"
			 case 2 param=param & " and topicmode=1"
			 case 3 param=param & " and reward>0"
			 case 4 param=param & " and locktopic=1"
			 case 5 param=param & " and locktopic=0"
			end select
			If Classid>0 Then param=param & " and classid in(select classid from KS_askclass where ','+parentstr +'' like '%," & classid & ",%')"
			If KS.G("keyword")<>"" Then param=param & " and title like '%" & Trim(KS.G("KeyWord")) & "%'"
			If KS.G("Class")<>"" Then Param=Param & " and classid=" & KS.ChkClng(KS.G("Class"))
			If KS.G("askName")<>"" Then Param=Param &" and username like '%" & Trim(KS.G("askName")) & "%'"
			If KS.G("answerName")<>"" Then Param=Param &" and topicid in(select topicid from KS_AskPosts1 Where UserName like '%" & Trim(KS.G("answerName")) & "%')"
			if Request("StartDate")<>"" and isdate(request("StartDate")) then
			  Param=Param & " and DateAndTime>=#" & request("DateAndtime") & "#"
			end if
			if Request("endDate")<>"" and isdate(request("endDate")) then
			 Dim enddate:EndDate = DateAdd("d", 1, Request("EndDate"))
			  Param=Param & " and DateAndTime<=#" & enddate & "#"
			end if
			If KS.G("orders")<>"" Then
			 OrderStr=" Order By " & KS.G("orders")
			Else
			 OrderStr=" Order By TopicID Desc"
			End If
			
			
			If count=0 Then
				totalrec=Conn.Execute("SELECT COUNT(*) FROM KS_AskTopic "&Param&"")(0)
			Else
				totalrec=count
			End If
			Set Rs=KS.InitialObject("ADODB.Recordset")
			SQL="SELECT "& SQLField &" FROM [KS_AskTopic]  "&Param&OrderStr
			Rs.Open SQL,Conn,1,1
			If Not Rs.EOF Then
			   If (pagenow - 1) * MaxPerPage < totalrec Then	Rs.Move (pagenow-1) * maxperpage
				Topiclist=Rs.GetRows(maxperpage)
			Else
				Topiclist=Null
			End If
			Rs.close()
			Set Rs=Nothing
			
			Pcount = CLng(totalrec / maxperpage)
			If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
			If pagenow>Pcount Then pagenow=1

		End Sub
		
		Public Sub LoadCategoryList()
	  If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then
		Dim Rs,SQL,TempXmlDoc
		Set Rs = Conn.Execute("SELECT classid,ClassName,Readme,rootid,depth,parentid,Parentstr,child FROM KS_AskClass ORDER BY orders,classid")
		If Not (Rs.BOF And Rs.EOF) Then
			SQL=Rs.GetRows(-1)
			Set TempXmlDoc = ArrayToxml(SQL,Rs,"row","classlist")
		Else
		    KS.Die "请先添加问答分类!"
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(TempXmlDoc) Then
			Application.Lock
				Set Application(KS.SiteSN&"_askclasslist") = TempXmlDoc
			Application.unLock
		End If
	 End If
	End Sub
	
	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
				Dim i,node,rs,j
				If xmlroot="" Then xmlroot="xml"
				Set ArrayToxml = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
				If row="" Then row="row"
				For i=0 To UBound(DataArray,2)
					Set Node=ArrayToxml.createNode(1,row,"")
					j=0
					For Each rs in Recordset.Fields
							 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
							 j=j+1
					Next
					ArrayToxml.documentElement.appendChild(Node)
				Next
	End Function
		
	Sub GetChildList()
		   If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then LoadCategoryList
		   Set Catelist = Application(KS.SiteSN&"_askclasslist")
		   If Not Catelist Is Nothing Then
			Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&classid&"]")
			If Not Node Is Nothing Then
				child=Node.selectSingleNode("@child").text
				If child>0 Then
					cid=classid
				Else
					cid=CLng(Node.selectSingleNode("@parentid").text)
				End If 
			Else
			  cid=0
			End If
		   Else
		     cid=0
		   End If
		
		  Dim SQLStr:SQLStr = "SELECT classid,classname,AskPendNum,AskDoneNum FROM KS_AskClass WHERE parentid="&KS.ChkClng(cid)&" ORDER BY orders,classid"
		  Dim RS:Set RS=Conn.Execute(SQLStr)
		  If Not RS.Eof Then
		   classarr=RS.GetRows(-1)
		  End If
		  RS.Close:Set RS=Nothing
		End Sub
		
		Sub showAsked()
			Dim Rs,SQL,XMLDom,Node,i
			Dim PostUserTitle,DelAction
			topicid=KS.ChkClng(Request("topicid"))
			SQL="SELECT TopicID,classid,username,classname,title,Expired,Closed,PostTable,DateAndTime,LastPostTime,ExpiredTime,LockTopic,Reward,Hits,PostNum,CommentNum,TopicMode,Broadcast,Anonymous,supplement FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				ErrMsg="错误的系统参数"
				FoundErr = True
				Exit Sub
			End If
			Set XMLDom = KS.RsToxml(Rs,"topic","xml")
			Set Rs = Nothing
			Set Node = XMLDom.documentElement.selectSingleNode("topic")
			If Not Node Is Nothing Then
				topicid = CLng(Node.selectSingleNode("@topicid").text)
				classid = CLng(Node.selectSingleNode("@classid").text)
				topicmode = CLng(Node.selectSingleNode("@topicmode").text)
				PostNum = CLng(Node.selectSingleNode("@postnum").text)
				ExpiredTime = CDate(Node.selectSingleNode("@expiredtime").text)
				CommentNum = CLng(Node.selectSingleNode("@commentnum").text)
				HeadTitle = Trim(Node.selectSingleNode("@title").text)
				TopicUseTable = Trim(Node.selectSingleNode("@posttable").text)
			End If
			Set Node = Nothing
			Set XMLDom = Nothing
		%>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<tr>
			<th>问题：<%=HeadTitle%></th>
		</tr>
		<%
			Call showAskedlist()
			If IsArray(Topiclist) Then
				For i=0 To Ubound(Topiclist,2)
					If Not Response.IsClientConnected Then Response.End
					If CLng(Topiclist(12,i))=0 Then
						PostUserTitle="提问者："
						DelAction="del"
					Else
						PostUserTitle="回答者："
						DelAction="delask"
					End If
		%>
		
		<tr>
			<td class="tdbg">
			  <table border="0" width="100%" <%If TopicList(10,i) = 1 Then Response.Write " style='border:5px solid #ff6600;'"%>>
<tr>
			<td colspan=2  class="clefttitle" height="30">
				<%=PostUserTitle%>:<%=Topiclist(3,i)%>  
				&nbsp;&nbsp;&nbsp;
				时间:<%=TopicList(7,i)%><%If TopicList(10,i) = 1 Then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=red size=2><strong>最佳答案</strong></font>"%>
				</td>
		</tr>			  
			   <form action="?i=<%=i%>&action=save&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>" method="post">
			   
			   <tr>
			    <td width="600">
				 <%if i=0 then%>
				    标题:<input type="text" name="title" value="<%=TopicList(4,i)%>">
					分类:
					
			<%  dim ii
				Response.Write " <select name=""classid"">"
				Response.Write "<option value=""0"">做为一级分类</option>"
				SQL = "SELECT classid,depth,ClassName FROM KS_AskClass ORDER BY rootid,orders"
				Set Rs = Conn.Execute(SQL)
				Do While Not Rs.EOF
					Response.Write "<option value=""" & Rs("classid") & """ "
					If  CLng(classid) = Rs("classid") Then Response.Write "selected"
					Response.Write ">"
					If Rs("depth") = 1 Then Response.Write "&nbsp;&nbsp;├ "
					If Rs("depth") > 1 Then
						For ii = 2 To Rs("depth")
							Response.Write "&nbsp;&nbsp;│"
						Next
						Response.Write "&nbsp;&nbsp;├ "
					End If
					Response.Write Rs("ClassName") & "</option>" & vbCrLf
					Rs.movenext
				Loop
				Rs.Close
				Response.Write "</select>"
				Set Rs = Nothing
			%>
					
				 <%end if%>
				 
				 审核
				 <input type="radio" name="LockTopic" value="0"<%If TopicList(11,i) = 0 Then Response.Write " checked=""checked"""%> /> 确定审核&nbsp;&nbsp;
				<input type="radio" name="LockTopic" value="1"<%If TopicList(11,i) = 1 Then Response.Write " checked=""checked"""%> /> 取消审核
				
				<br />
			        <textarea name="content<%=i%>" id="content<%=i%>" style="display:None;width:600px;height:80px"><%=server.Htmlencode(Topiclist(5,i))%></textarea>
					<iframe  id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=content<%=i%>&amp;Toolbar=Basic" width="98%" height="150" frameborder="0" scrolling="no"></iframe>
				
			    </td>
				<td width="200" align="center">
			<input type="submit" value=" 编 辑 " class="button"> 
			<%If TopicList(10,i) <> 1 Then%>
			<input type="button" value=" 删 除 " class="button" onClick="if (confirm('删除后将不能恢复，您确定要删除吗?')){location.href='KS.AskList.asp?action=<%=DelAction%>&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>'}">
			<%end if%>
			<%If topicmode=0 and i<>0 then%>
			<br /><br/><input type="button" value=" 采纳为最佳答案 " class="button" onClick="if (confirm('您确定采纳该答案为最佳答案吗?')){location.href='KS.AskList.asp?action=SetSatis&postsid=<%=Topiclist(0,i)%>&topicid=<%=Topiclist(2,i)%>'}">
			<%end if%>
			    </td>
			  </tr>
			  </form>
			  </table>
			</td>
		</tr>
		<%
				Next
			End If
			Topiclist=Null
		%>
		<tr>
			<td class="tablerow1" align="right" id="NextPageText">
			<%
			Call KSCLS.ShowPage(totalrec, MaxPerPage, "KS.AskList.asp", True, "个", pagenow, KS.QueryParam("page"))

			%>
			</td>
		</tr>
		</table>
		<%
		End Sub
		
		Sub showAskedlist()
			Dim Rs,SQL
			maxperpage=10
			
			SQLField="postsid,classid,TopicID,UserName,topic,content,addText,PostTime,DoneTime,star,satis,LockTopic,PostsMode,VoteNum,Plus,Minus,PostIP,Report"
			If count=0 Then
				totalrec=Conn.Execute("SELECT COUNT(*) FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" "&SQLQuery&"")(0)
			Else
				totalrec=count
			End If
			Set Rs=Server.CreateObject("ADODB.Recordset")
			SQL="SELECT "& SQLField &" FROM ["&TopicUseTable&"]  WHERE topicid="&topicid&" "&SQLQuery&" ORDER BY postsMode ASC,Satis desc,postsid"
			Rs.Open SQL,Conn,1,1
			If Not Rs.EOF Then
			   If (pagenow - 1) * MaxPerPage < totalrec Then Rs.Move (pagenow-1) * maxperpage
				Topiclist=Rs.GetRows(maxperpage)
			Else
				Topiclist=Null
			End If
			
			Rs.close()
			Set Rs=Nothing
		
			Pcount = CLng(totalrec / maxperpage)
			If Pcount < totalrec / maxperpage Then Pcount = Pcount + 1
			If pagenow>Pcount Then pagenow=1
			pagelinks="KS.AskList.asp.asp?action=asked&topicid="&topicid&"&count="&totalrec&"&"
		End Sub
		
		
		Sub saveAsked()
			Dim Rs,SQL,postsid
			Dim TextContent,satis,LockTopic,strTitle,star
			postsid=KS.ChkClng(Request("postsid"))
			topicid=KS.ChkClng(Request("topicid"))
			If Trim(Request.Form("content"&request("i")))="" Then
				Call KS.AlertHintScript("内容不能为空!")
				Exit Sub
			End If
			SQL="SELECT top 1 TopicID,classid,title,Username,Expired,Closed,PostTable,LockTopic,TopicMode,supplement FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				ErrMsg="错误的系统参数"
				FoundErr = True
				Exit Sub
			End If
			topicid=Rs("TopicID")
			strTitle=Rs("title")
			TopicUseTable=Trim(Rs("PostTable"))
			TopicMode=Rs("TopicMode")
			Set Rs = Nothing
			TextContent=Request.Form("content"&request("i"))
			LockTopic=KS.ChkClng(Request.Form("LockTopic"))
			Conn.Execute ("UPDATE ["&TopicUseTable&"] SET content='"&TextContent&"',LockTopic="&LockTopic&" WHERE postsid="&postsid&" And topicid="&topicid)
			If KS.G("I")="0" Then
			 dim className:className=LFCls.GetSingleFieldValue("select top 1 classname from [KS_AskClass] Where ClassID=" & KS.ChkClng(KS.G("ClassID")))
			Conn.Execute ("UPDATE [KS_AskTopic] SET className='" & className&"',ClassID=" & KS.ChkClng(KS.G("ClassID")) & ",LockTopic="&LockTopic&" WHERE topicid="&topicid)
			Conn.Execute ("UPDATE [KS_AskAnswer] SET className='" & className&"',ClassID=" & KS.ChkClng(KS.G("ClassID")) & " WHERE topicid="&topicid)
			Conn.Execute ("UPDATE ["&TopicUseTable&"] SET ClassID=" & KS.ChkClng(KS.G("ClassID")) & " WHERE topicid="&topicid)
			End If
			
			If strTitle<>Request.Form("title") and trim(Request.Form("title"))<>"" Then
				Conn.Execute ("UPDATE ["&TopicUseTable&"] SET topic='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
				Conn.Execute ("UPDATE [KS_AskTopic] SET title='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
				Conn.Execute ("UPDATE [KS_AskAnswer] SET title='"&Trim(Request.Form("title"))&"' WHERE topicid="&topicid)
			End If
			Call KS.AlertHintScript("恭喜您！编辑/审核问题成功。")
		End Sub
		
		'推荐问题
		Sub Recommend()
			Dim TopicIDlist,SQL,RS,ScoreToQuestionUser,ScoreToAnswerUser
			TopicIDlist=KS.FilterIds(Request("topicid"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			ScoreToQuestionUser=KS.ChkClng(KS.ASetting(33))
			ScoreToAnswerUser=KS.ChkClng(KS.ASetting(34))
			SQL="SELECT * FROM KS_AskTopic Where recommend=0 and TopicID in(" & TopicIDList & ")"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQL,Conn,1,3
			Do While Not RS.Eof
			  
			  RS("Recommend")=1
			  RS.Update
			  
			   '给提问者加积分
			  If ScoreToQuestionUser>0 Then
				 Call KS.ScoreInOrOut(RS("UserName"),1,ScoreToQuestionUser,"系统","问吧问题[" & rs("title") & "]被管理员推荐!",0,0)
			  End If
			   '给最佳回答者加积分
			  If ScoreToAnswerUser>0 Then
			     Dim rsb:set rsb=Conn.Execute("select username From KS_AskAnswer Where TopicID=" & RS("TopicID") & " and AnswerMode=1")
				 if not rsb.eof then
				 Call KS.ScoreInOrOut(rsb(0),1,ScoreToAnswerUser,"系统","问吧问题[" & rs("title") & "]最佳答案被管理员推荐!",0,0)
				 end if
				 rsb.close:set rsb=nothing
			  
			  End If
			  
			  RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'取消推荐问题
		Sub UnRecommend()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("topicid"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			SQL="SELECT * FROM KS_AskTopic Where recommend=1 and TopicID in(" & TopicIDList & ")"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQL,Conn,1,3
			Do While Not RS.Eof
			  RS("Recommend")=0
			  RS.Update
			  RS.MoveNext
			Loop
			RS.Close
			Set RS=Nothing
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub

        '批量审核
		Sub Verify()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("topicid"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			Conn.Execute("Update KS_AskTopic Set LockTopic=0 Where TopicID in(" & TopicIDList & ")")
			Conn.Execute("Update KS_AskPosts1 Set LockTopic=0 Where PostsMode=0 and TopicID in(" & TopicIDList & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
        '取消审核
		Sub UnVerify()
			Dim TopicIDlist,SQL,RS
			TopicIDlist=KS.FilterIds(Request("topicid"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End
			Conn.Execute("Update KS_AskTopic Set LockTopic=1 Where TopicID in(" & TopicIDList & ")")
			Conn.Execute("Update KS_AskPosts1 Set LockTopic=1 Where PostsMode=0 and TopicID in(" & TopicIDList & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'设为最佳答案
		Sub SetSatis()
		   Dim Rs,SQL,i,SQLArry,postsid,ClassID
			Dim TopicID,userName,k,TopicUseTable
			TopicID=KS.ChkClng(Request("topicid"))
            Postsid=KS.ChkClng(Request("postsid"))
			SQL="SELECT TopicID,userName,PostTable,TopicMode,classid FROM KS_AskTopic WHERE topicid="&TopicID
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			TopicUseTable=Rs(2)
			UserName=Rs(1)
			ClassID=RS(4)
			Set Rs=Nothing
			
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT postsid,TopicID,username,topic FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and PostsMode=1 and LockTopic=0 and postsid="& Postsid
			Rs.Open SQL,Conn,1,1
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('友情提示!\n\n请选择正确的问题ID!');history.back();</script>"
				Response.End
			Else
				Do While Not Rs.EOF
					Conn.Execute ("UPDATE ["&TopicUseTable&"] SET satis=1,DoneTime="& SqlNowString &" WHERE postsid="& Rs(0))
					
					If KS.ChkClng(KS.ASetting(31))>0 Then
				    Call KS.ScoreInOrOut(Rs(2),1,KS.ChkClng(KS.ASetting(31)),"系统","您的对问题[" & rs("topic") & "]的回答被设为最佳答案!",0,0)
					End If

					Conn.Execute ("UPDATE KS_AskAnswer SET AnswerMode=1 WHERE topicid="&topicid&" and username='"& Rs(2) & "'")
					Rs.movenext
				Loop
			End If
			Rs.Close:Set Rs = Nothing
		
			Conn.Execute ("UPDATE KS_AskTopic SET LastPostTime="& SqlNowString &",TopicMode=1 WHERE topicid="&topicid&" and username='"& UserName &"' and Closed=0 and LockTopic=0")
			Conn.Execute ("UPDATE KS_AskAnswer SET TopicMode=1 WHERE topicid="&topicid)
			
			'Conn.Execute ("UPDATE KS_User SET Score=Score+" & KS.ChkClng(KS.ASetting(32)) & " WHERE username='"& UserName & "'")
			Conn.Execute ("UPDATE KS_AskClass SET AskPendNum=AskPendNum-1,AskDoneNum=AskDoneNum+1 WHERE classid="& classid)
			Call KS.Alert("恭喜您！设置最佳答案成功!","KS.AskList.asp?action=asked&topicid=" & topicid)
		End Sub
		
		
		Sub delTopic()
			Dim Rs,SQL,i,SQLArry
			Dim TopicIDlist,userName,k
			Dim MinusPoints,ClassNumStr,parentArr
			TopicIDlist=KS.FilterIds(Request("topicid"))
			If TopicIDlist="" Then 	Call KS.AlertHintScript("你没有选择问题!"):Response.End

			SQL="SELECT TopicID,userName,PostTable,TopicMode,classid FROM KS_AskTopic WHERE topicid in("&TopicIDlist&")"
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			SQLArry=Rs.GetRows(-1)
			Set Rs=Nothing
		
			If IsArray(SQLArry) Then
				For i=0 To Ubound(SQLArry,2)
					topicid=CLng(SQLArry(0,i))
					userName=SQLArry(1,i)
					TopicUseTable=Trim(SQLArry(2,i))
					TopicMode=CLng(SQLArry(3,i))
					parentArr=split(conn.execute("select parentstr from KS_askclass where classid=" & SQLArry(4,i))(0),",")
					Select Case TopicMode
						Case 1
							MinusPoints=KS.ChkCLng(KS.ASetting(39))
							ClassNumStr="AskDoneNum=AskDoneNum-1 Where AskDoneNum>0"
						Case Else
							MinusPoints=KS.ChkClng(KS.ASetting(40))
							ClassNumStr="AskPendNum=AskPendNum-1 Where AskPendNum>0"
					End Select
					'Conn.Execute ("UPDATE KS_User SET score=score-"&MinusPoints&" WHERE username='"&username & "'")
					Conn.Execute("DELETE FROM KS_UploadFiles WHERE channelid=1032 and infoid in(select postsid from "&TopicUseTable&" WHERE topicid="&topicid & ")")
					Conn.Execute("DELETE FROM KS_AskTopic WHERE topicid="&topicid)
					Conn.Execute("DELETE FROM KS_AskAnswer WHERE topicid="&topicid)
					Conn.Execute("DELETE FROM "&TopicUseTable&" WHERE topicid="&topicid)
					For K=0 To Ubound(parentarr)-1
					Conn.Execute("Update KS_AskClass Set " & ClassNumStr & " and classid=" & parentarr(k))
					Next
				Next
				SQLArry=Null
			End If
			if instr(lcase(REQUEST.SERVERVARIABLES("HTTP_REFERER")),"index.asp")=0 then
			Call KS.AlertHintScript("恭喜您！数据删除成功!")
			else
			Call KS.Alert("恭喜您！数据删除成功!","KS.AskList.asp")
			end if
		End Sub
		
		Sub delAsked()
			Dim Rs,SQL,postsid
			Dim SQLArry,userName,PostNum
			Dim MinusPoints,MinusExperience
			Dim satis,PostsMode
			postsid=KS.ChkClng(Request("postsid"))
			topicid=KS.ChkClng(Request("topicid"))
			SQL="SELECT TopicID,username,PostTable,TopicMode,PostNum FROM KS_AskTopic WHERE topicid="&topicid
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs=Nothing
				Call KS.AlertHistory("错误的系统参数!",-1)
				Response.End
			End If
			SQLArry=Rs.GetRows(1)
			Set Rs=Nothing
			If IsArray(SQLArry) Then
				topicid=CLng(SQLArry(0,0))
				userName=SQLArry(1,0)
				TopicUseTable=Trim(SQLArry(2,0))
				TopicMode=CLng(SQLArry(3,0))
				PostNum=CLng(SQLArry(4,0))
			Else
				Call KS.AlertHintScript("错误的系统参数!")
				Response.End
			End If
			SQLArry=Null
			If PostNum>0 Then
				SQL="SELECT postsid,username,satis,PostsMode FROM "&TopicUseTable&" WHERE postsid="&postsid
				Set Rs = Conn.Execute(SQL)
				If Rs.BOF And Rs.EOF Then
					Set Rs=Nothing
					Call KS.AlertHintScript("错误的系统参数!")
				    Response.End
				End If
				SQLArry=Rs.GetRows(1)
				Set Rs=Nothing
				If IsArray(SQLArry) Then
					postsid=CLng(SQLArry(0,0))
					username=SQLArry(1,0)
					satis=CLng(SQLArry(2,0))
					PostsMode=CLng(SQLArry(3,0))
					If satis=0 Then
						MinusPoints=KS.ChkCLng(KS.ASetting(38))
					Else
						MinusPoints=KS.ChkClng(KS.ASetting(37))
					End If
					If PostsMode>0 Then
						Conn.Execute("DELETE FROM KS_AskAnswer WHERE topicid="&topicid&" And username='"&username&"' And AnswerNum<2")
						Conn.Execute("DELETE FROM "&TopicUseTable&" WHERE postsid="&postsid)
						'Conn.Execute ("UPDATE KS_User SET score=score-"&MinusPoints&" WHERE username='"&username & "'")
						Conn.Execute ("UPDATE KS_AskAnswer SET AnswerNum=AnswerNum-1 WHERE topicid="&topicid&" And username='"&username & "'")
					End If
				End If
				SQLArry=Null
			End If
			Call KS.Alert("恭喜您！数据删除成功!","KS.AskList.asp?action=asked&topicid=" & topicid)
		End Sub
End Class
%>
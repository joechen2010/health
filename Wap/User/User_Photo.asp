<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="UpFileSave.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的相册">
<p>
<%
Set KS=New PublicCls
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If

Action=Trim(Request("Action"))

If KS.SSetting(0)=0 Then
   Response.write "对不起，本站关闭个人空间功能！<br/>"
ElseIF Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)=0 Then
   Response.write "您还没有开通个人空间！<br/>"
ElseIF Conn.Execute("Select status From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)<>1 Then
   Response.write "对不起，你的空间还没有通过审核或被锁定！<br/>"
Else
       Select Case Action
          Case "Gallery"
			  Call Gallery()'上传须知
			 Case "Delxc"
			  Call Delxc()'删除相册
			 Case "Delzp"
			   Call Delzp()'删除照片
			 Case "Editzp" 
			  Call Editzp()'编辑照片
			Case "Addzp"
			 Call Addzp()'上传照片
			 Case "AddSave"
			  Call AddSave()'相片保存
			 Case "EditSave"
			  Call EditSave()'相片修改成功
			 Case "ViewZP"
			  Call ViewZP()'查看照片
			 Case "Editxc","Createxc"
			  Call Managexc() '相册，添加／修改
			 Case "photoxcsave"
			  Call photoxcsave() '保存相册
			 Case Else
			  Call PhotoxcList()'相册列表
		End Select
End If

'相册列表============================================
Sub PhotoxcList()
status=KS.S("status")
%>
=我的相册=<br/>
相册列表|
<a href="User_UpFile.asp?ChannelID=9997&amp;<%=KS.WapValue%>">WAP2.0上传</a><br/>
<a href="User_UpFile.asp?ChannelID=9998&amp;<%=KS.WapValue%>">创建相册</a>|
<a href="User_Photo.asp?action=Gallery&amp;<%=KS.WapValue%>">照片上传须知</a><br/>
<% if instr(Request.ServerVariables("http_accept"),"ucweb")<5  then Response.write GuangGao %>
---------<br/>
<%Select Case status
Case "1"
%>
已审[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=1")(0)%>]
<%Case Else%>
<a href="User_Photo.asp?status=1&amp;<%=KS.WapValue%>">已审[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=1")(0)%>]</a>
<%End Select
Select Case status
Case "0"%>
未审[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=0")(0)%>]
<%Case Else%>
<a href="User_Photo.asp?status=0&amp;<%=KS.WapValue%>">未审[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=0")(0)%>]</a>
<%End Select
Select Case status
Case "2"%>
锁定[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=2")(0)%>]
<%Case Else%>
<a href="User_Photo.asp?status=2&amp;<%=KS.WapValue%>">锁定[<%=conn.execute("select count(id) from ks_photoxc where username='"&KSUser.UserName&"' and status=2")(0)%>]</a>
<%End Select%>
<br/>---------<br/>
<%
If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If


set rs=server.createobject("adodb.recordset")
Param=" Where UserName='"&KSUser.UserName&"'"
IF status<>"" Then
   Param=""&Param&" And status="&status&""
End if
sql = "select * from KS_Photoxc "&Param&" order by AddDate DESC"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
   response.write "您还没有创建相册,马上<a href=""User_UpFile.asp?ChannelID=9998&amp;" & KS.WapValue & """>创建相册</a>!<br/>"
else
   MaxPerPage =2
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   Do While Not RS.Eof
    '*******************************************************
    If KS.BusinessVersion = 1 Then
       response.write "<img src='../JpegMore.asp?JpegSize=128x0&amp;JpegUrl="&Rs("photourl")&"'/><br/>"
    Else
       response.write "<img src='" &rs("photourl")& "'  /><br/>"
    End if
    '*******************************************************
	response.write "<a href=""User_Photo.asp?action=ViewZP&amp;xcid="&rs("ID")&"&amp;" & KS.WapValue & """>"&rs("xcname")&"</a><br/>"
	select case rs("status")
	 case 1:response.write "(已审/"
	 case 0:response.write "(未审/"
	end select
	select case rs("flag")
	 case 1
	 response.write "公开"
	 case 2
	 response.write "会员"
	 case 3
	 response.write "密码"
	 case 4
	 response.write "稳私"
	end select
	response.write "/"&rs("xps")&"张/"&rs("hits")&"次)<br/>"
	response.write "<a href=""User_Photo.asp?action=Editxc&amp;id="&rs("ID")&"&amp;" & KS.WapValue & """>修改相册</a> "
	response.write "<a href=""User_Photo.asp?action=Delxc&amp;id="&rs("ID")&"&amp;" & KS.WapValue & """>删除</a><br/>"

	  Rs.Movenext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Photo.asp", True, "本相册", CurrentPage, "Status=" & Status &"&amp;" & KS.WapValue & "")
end if

rs.close
Response.Write "<br/>---------<br/>"
End Sub

'删除相册=====================================================
Sub Delxc()
	id=KS.S("id")
	Dim rs:Set rs=server.createobject("adodb.recordset")
	rs.open "select PhotoUrl from KS_Photoxc where ID="&ID&"",conn,1,1
	if not rs.eof then
	KS.DeleteFile(rs("photourl"))
	Conn.Execute("Delete From KS_Photoxc Where ID="&ID&"")
	end if
	rs.close:set rs=nothing
	Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where xcid in(" &id & ")",conn,1,1
	if not rs.eof then
	  do while not rs.eof
	   KS.DeleteFile(rs("photourl"))
	   rs.movenext
	   loop
	end if
	Conn.execute("delete from ks_photozp where xcid in(" & id& ")")
	rs.close:set rs=nothing
		  Response.write "相册删除成功。<br/>"
		  Response.Write "<a href=""User_Photo.asp?action=PhotoxcList&amp;" & KS.WapValue & """>照片列表</a><br/>"
End Sub

'删除照片=====================================================
  Sub Delzp()
	id=KS.S("id")
	xcid=KS.S("xcid")
	Dim RS:Set rs=server.createobject("adodb.recordset")
	rs.open "select * from ks_photozp where id="&id&"",conn,1,1
	if not rs.eof then
	   KS.DeleteFile(rs("photourl"))
	   Conn.execute("update ks_photoxc set xps=xps-1 where id="&rs("xcid"))
	end if
	Conn.execute("delete from ks_photozp where id="&id&"")
	rs.close:set rs=nothing
		  Response.write "照片删除成功。<br/>"
		  Response.Write "<a href=""User_Photo.asp?action=ViewZP&amp;xcid="&xcid&"&amp;" & KS.WapValue & """>照片列表</a><br/>"
  End Sub

'查看照片=====================================================
Sub ViewZP()
xcid=KS.S("xcid")
%>
=查看照片=<br/>
<a href="User_Photo.asp?action=PhotoxcList&amp;<%=KS.WapValue%>">相册列表</a>|
<a href="User_UpFile.asp?ChannelID=9997&amp;<%=KS.WapValue%>">WAP2.0上传</a><br/>
<a href="User_UpFile.asp?ChannelID=9998&amp;<%=KS.WapValue%>">创建相册</a>|
<a href="User_Photo.asp?action=Gallery&amp;<%=KS.WapValue%>">照片上传须知</a><br/>
---------<br/>
<%
Set rs=Server.CreateObject("ADODB.RECORDSET")
rs.Open "select xcname from KS_Photoxc WHERE ID="&XCID&" order by AddDate desc",CONN,1,1
if rs.Eof And rs.Bof Then 
   rs.close:set rs=nothing
   response.write "参数传递出错！"
   response.end
else
   title=rs(0)
   rs.close
end if

If KS.S("page") <> "" Then
   CurrentPage = KS.ChkClng(KS.S("page"))
Else
   CurrentPage = 1
End If

set rs=server.createobject("adodb.recordset")
sql="select * from KS_PhotoZP Where xcid="& xcid&" order by AddDate desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   response.write "该相册下没有照片,<a href=""User_UpFile.asp?ChannelID=9997&amp;" & KS.WapValue & """>WAP2.0上传</a><br/>"
else
   MaxPerPage =2
   totalPut = RS.RecordCount
   If CurrentPage < 1 Then	CurrentPage = 1
   If (CurrentPage - 1) * MaxPerPage > totalPut Then
      If (totalPut Mod MaxPerPage) = 0 Then
	     CurrentPage = totalPut \ MaxPerPage
	  Else
	     CurrentPage = totalPut \ MaxPerPage + 1
	  End If
   End If
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   
   do while not rs.eof
      '*******************************************************
	  If KS.BusinessVersion = 1 Then
      response.write "<img src='../JpegMore.asp?JpegSize=128x0&amp;JpegUrl="&rs("photourl")&"'/><br/>"
	  Else
      response.write "<img src='" &rs("photourl")& "'  /><br/>"
	  End if
	  '*******************************************************
	  response.write "照片名称:"&rs("title")&"<br/>"
	  response.write "浏览次数:"&rs("hits")&"<br/>"
	  response.write "创建日期:"&rs("adddate")&"<br/>"
	  response.write "图片大小："&rs("photosize")&"<br/>"
	  'response.write "相片地址："&rs("photourl")&"<br/>"
	  response.write "相片描述："&rs("descript")&"<br/>"
	  response.write "所属相册："&conn.execute("select xcname from ks_photoxc where id=" & xcid)(0)&"<br/>"
	  response.write "<a href=""User_Photo.asp?action=Editzp&amp;id="&rs("ID")&"&amp;" & KS.WapValue & """>修改照片</a>|"
	  response.write "<a href=""User_Photo.asp?action=Delzp&amp;xcid="&xcid&"&amp;id="&rs("ID")&"&amp;" & KS.WapValue & """>删除照片</a><br/>"
      rs.movenext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Photo.asp", True, "张照片", CurrentPage, "Action=ViewZP&amp;xcid=" & xcid &"&amp;" & KS.WapValue & "")
End If
Rs.Close
response.Write "<br/>---------<br/>"
End Sub

 '相册，添加／修改==================================================================
Sub Managexc()
    On Error Resume Next
    If KS.S("Action")="Createxc" Then
       sTemp = Kesion
	   sTempArr=split(sTemp,"|||")
	   If Cbool(sTempArr(0))=False Then
	      Response.Write sTempArr(1)
	   Else
	      PhotoUrl = Replace(sTempArr(1),"|","")
	   End if
    End if
	
    If KS.S("ID")<>"" Then
	   Set Rs=Server.CreateObject("ADODB.RECORDSET")
	   sql = "select * from KS_Photoxc Where ID="&KS.S("ID")&""
	   Rs.Open sql,Conn,1,1
	   If Not RS.EOF Then
		  xcname=rs("xcname")
		  ClassID=rs("ClassID")
		  Descript=rs("Descript")
		  flag=rs("Flag")
		  If PhotoUrl<>"" Then
		     PhotoUrl=PhotoUrl
		  Else
		     PhotoUrl=rs("PhotoUrl")
		  End if
		  PassWord=rs("PassWord")
		  OpStr="确定修改"
		  TipStr="修改我的相册"
		  rs.Close:Set rs=Nothing
	   End if
	Else
	   xcname=FormatDatetime(Now,2)
	   ClassID="0"
	   flag="1"
	   password=""
	   OpStr="立即创建"
	   TipStr="创建我的相册"
	End if


%>
=<%=TipStr%>=<br/>
相册封面<br/>
<img src="<%=PhotoUrl%>"/><br/>
相册名称:<input name="xcname<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=xcname%>"/><br/>
相册分类:<select name="ClassID">
<option value="0">选择类别</option>
<% Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
If Not RS.EOF Then
Do While Not RS.Eof 
Response.Write "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
RS.MoveNext
Loop
End If
RS.Close
Set RS=Nothing
%>
</select> <br/>             
是否公开:<select name="flag">
<option value="1">完全公开</option>
<option value="2">会员开见</option>
<option value="3">密码共享</option>
<option value="4">隐私相册</option>
</select><br/>
访问密码:<input name="password<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=password%>"/><br/>
相册介绍:<input name="Descript<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=Descript%>"/><br/>
<anchor><%=OpStr%><go href="User_Photo.asp?Action=photoxcsave&amp;id=<%=id%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
<postfield name="xcname" value="$(xcname<%=minute(now)%><%=second(now)%>)"/>
<postfield name="ClassID" value="$(ClassID)"/>
<postfield name="flag" value="$(flag)"/>
<postfield name="password" value="$(password<%=minute(now)%><%=second(now)%>)"/>
<postfield name="PhotoUrl" value="<%=PhotoUrl%>"/>
<postfield name="Descript" value="$(Descript<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>
<%
response.Write "---------<br/>"
End Sub

 '保存相册===============================================
Sub photoxcsave()
    id=KS.S("id")
	xcname=KS.S("xcname")
	ClassID=KS.S("ClassID")
	flag=KS.S("flag")
	password=KS.S("password")
	PhotoUrl=KS.S("PhotoUrl")
	Descript=KS.S("Descript")
	If xcname="" Then
	   Response.write "出错提示，请输入相册名称！"
	   Response.Write "<a href=""User_Photo.asp?action=Createxc&amp;id="&id&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	ElseIF ClassID=0 Then
       Response.write "出错提示，请选择相册类型！"
	   Response.Write "<a href=""User_Photo.asp?action=Createxc&amp;id="&id&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
    Else
      Set rs=Server.CreateObject("ADODB.RECORDSET")
	  If id<>"" Then
	     rs.open "select * from KS_Photoxc Where id="&id&"",conn,1,3
	     If rs.Eof And rs.Bof Then
	        rs("UserName")=KSUser.UserName
		    rs("xcname")=xcname
			rs("ClassID")=ClassID
			rs("Descript")=Descript
			rs("Flag")=Flag
			rs("Password")=PassWord
			rs("PhotoUrl")=PhotoUrl
			rs.Update
		 End If
	  Else
	     rs.open "select * from KS_Photoxc",conn,1,3
		 rs.AddNew
		 rs("AddDate")=now
		 If KS.SSetting(4)=1 Then
			rs("Status")=0 '设为已审
		 Else
			rs("Status")=1 '设为已审
		 End If
		 rs("UserName")=KSUser.UserName
		 rs("xcname")=xcname
		 rs("ClassID")=ClassID
		 rs("Descript")=Descript
		 rs("Flag")=Flag
		 rs("Password")=PassWord
		 rs("PhotoUrl")=PhotoUrl
		 rs.Update
	  End If
		  Response.write "相册创建/修改成功。<br/>"
		  response.Write "<a href=""User_Photo.asp?action=PhotoxcList&amp;" & KS.WapValue & """>我的相册</a><br/>"
		  rs.Close:Set rs=Nothing
	End If
End Sub

'上传照片==================================
Sub Addzp()
    PhotoSize=KS.S("PhotoSize")
	PhotoUrl=KS.S("PhotoUrl")
	
    On Error Resume Next
	sTemp = Kesion
	sTempArr=split(sTemp,"|||")
	If Cbool(sTempArr(0))=False Then
	   Response.Write sTempArr(1)
	Else
	   Response.write "=增加相片=<br/>"
	   Response.write "恭喜您,照片上传成功!请按发布按钮进行保存.<br/>"
	   Dim I,TempFileArr
	   TempFileStr = sTempArr(1)
	   TempFileStr=Left(TempFileStr,Len(TempFileStr)-1)
	   TempFileArr=split(TempFileStr,"|")
	   For I=Lbound(TempFileArr) To Ubound(TempFileArr)
	       Response.Write "" & TempFileArr(i) & "<br/>"
	   Next
	   %>
       选择相册:<select name="XCID" >
       <option value="0">选择相册</option>
	   <%
	   Set RS=Server.CreateObject("ADODB.RECORDSET")
	   RS.Open "Select * From KS_Photoxc where username='" & KSUser.UserName & "' order by id desc",Conn,1,1
	   If Not RS.EOF Then
	      Do While Not RS.Eof 
	      Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
		  RS.MoveNext
		  Loop
	   End If
	   RS.Close:Set RS=Nothing
	   %>
       </select><br/>
       照片名称:<input name="Title<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=Title%>"/><br/>
       照片介绍:<input name="Descript<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=Descript%>"/><br/>
       <anchor>立即发布<go href="User_Photo.asp?Action=AddSave&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
       <postfield name="PhotoUrls" value="<%=TempFileStr%>"/>
       <postfield name="XCID" value="$(XCID)"/>
       <postfield name="Title" value="$(Title<%=minute(now)%><%=second(now)%>)"/>
       <postfield name="Descript" value="$(Descript<%=minute(now)%><%=second(now)%>)"/>
       </go></anchor>
       <br/>
       <%
	End if
End Sub


'编辑照片==================================
Sub Editzp()
    id=KS.S("id")
	Set rs=Server.CreateObject("ADODB.RECORDSET")
	rs.Open "Select * From KS_PhotoZp Where ID="&id,Conn,1,1
	If Not rs.Eof Then
	   XCID=rs("XCID")
	   Title=rs("Title")
	   UserName=rs("UserName")
	   descript=rs("descript")
	   PhotoUrl=rs("PhotoUrl")
	 End If
	 rs.Close:Set rs=Nothing
%>
=编辑照片=<br/>
<img src="<%=PhotoUrl%>"/><br/>
选择相册:<select name="XCID" >
<option value="0">选择相册</option>
<%
Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select * From KS_Photoxc where username='"&KSUser.UserName&"' order by id desc",conn,1,1
If Not RS.EOF Then
Do While Not RS.Eof 
Response.Write "<option value=""" & RS("ID") & """>" & RS("XCName") & "</option>"
RS.MoveNext
Loop
End If
RS.Close:Set RS=Nothing
%>
</select><br/>
照片名称:<input name="Title<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=Title%>"/><br/>
照片介绍:<input name="Descript<%=minute(now)%><%=second(now)%>" type="text" maxlength="20" size="20" value="<%=Descript%>"/><br/>
<anchor>编辑照片<go href="User_Photo.asp?Action=EditSave&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
<postfield name="PhotoUrl" value="<%=PhotoUrl%>"/>
<postfield name="XCID" value="$(XCID)"/>
<postfield name="Title" value="$(Title<%=minute(now)%><%=second(now)%>)"/>
<postfield name="Descript" value="$(Descript<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>
<%
End Sub

'相片修改成功==================================
Sub EditSave()
    ID=KS.S("ID")
	XCID=KS.S("XCID")
	Title=KS.S("Title")
	Descript=KS.S("Descript")
	PhotoUrl=KS.S("PhotoUrl")
	If XCID=0 Then
	   Response.write "出错提示，你没有选择相册。<br/>"
	   Response.Write "<a href=""User_Photo.asp?action=Editzp&amp;id="&id&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	ElseIF Title="" Then
       Response.write "出错提示，你没有输入相片名称。<br/>"
	   Response.Write "<a href=""User_Photo.asp?action=Editzp&amp;id="&id&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	Else
	   Set RSObj=Server.CreateObject("Adodb.Recordset")
	   RSObj.Open "Select * From KS_PhotoZP Where ID="&ID,Conn,1,3
	   If Not RSObj.Eof Then
	      RSObj("Title")=Title
		  RSObj("XCID")=XCID
		  RSObj("PhotoUrl")=PhotoUrl
		  RSObj("Descript")=Descript
		  'RSObj("PhotoSize") =KS.GetFieSize(Server.Mappath(PhotoUrls))
		  RSObj.Update
	   End If  
	   RSObj.Close:Set RSObj=Nothing
	   Response.write "相片修改成功。<br/>"
	   Response.Write "<a href=""User_Photo.asp?action=ViewZP&amp;XCID="&XCID&"&amp;" & KS.WapValue & """>相册分娄</a><br/>"
	End If
End Sub

'相片保存==================================
Sub AddSave()
    XCID=KS.S("XCID")
	Title=KS.S("Title")
	Descript=KS.S("Descript")
	PhotoUrls=KS.S("PhotoUrls")
	
    If PhotoUrls="" Then 
	   Response.Write "你没有上传相片!<br/>"
	   Response.Write "<a href=""User_Photo.asp?action=Addzp&amp;PhotoUrl="&PhotoUrl&"&amp;PhotoSize="&PhotoSize&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	ElseIF XCID=0 Then
	   Response.write "出错提示，你没有选择相册。<br/>"
	   Response.Write "<a href=""User_Photo.asp?action=Addzp&amp;PhotoUrl="&PhotoUrl&"&amp;PhotoSize="&PhotoSize&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	ElseIF Title="" Then
	   Response.write "出错提示，你没有输入相片名称。<br/>"
       Response.Write "<a href=""User_Photo.asp?action=Addzp&amp;PhotoUrl="&PhotoUrl&"&amp;PhotoSize="&PhotoSize&"&amp;" & KS.WapValue & """>返回重写</a><br/>"
	Else
	   PhotoUrlArr=Split(PhotoUrls,"|")
	   Set RSObj=Server.CreateObject("Adodb.Recordset")
	   RSObj.Open "Select * From KS_PhotoZP",Conn,1,3
	   For I=0 to ubound(PhotoUrlArr)
	       RSObj.AddNew
		   RSObj("PhotoSize")=KS.GetFieSize(Server.Mappath(PhotoUrlArr(I)))
		   RSObj("Title")=Title
		   RSObj("XCID")=XCID
		   RSObj("UserName")=KSUser.UserName
		   RSObj("PhotoUrl")=PhotoUrlArr(I)
		   RSObj("Adddate")=Now
		   RSObj("Descript")=Descript
		   RSObj.Update
		Next
		RSObj.Close:Set RSObj=Nothing
		Conn.Execute("update KS_Photoxc set xps=xps+" & Ubound(PhotoUrlArr)+1 & " where id="&xcid)
		Response.write "相片保存成功。<br/>"
		Response.Write "<a href=""User_Photo.asp?action=ViewZP&amp;XCID="&XCID&"&amp;" & KS.WapValue & """>我的相册</a><br/>"
	End If
End Sub

'上传须知=========================================
Sub Gallery()
%>
=上传须知=<br/>
<a href="User_Photo.asp?action=PhotoxcList&amp;<%=KS.WapValue%>">相册列表</a>|
<a href="User_UpFile.asp?ChannelID=9997&amp;<%=KS.WapValue%>">WAP2.0上传</a><br/>
<a href="User_UpFile.asp?ChannelID=9998&amp;<%=KS.WapValue%>">创建相册</a>|
照片上传须知<br/>
---------<br/>
我们力求打造一个真实的交友环境，因此只允许上传照片，以下类型的均会被客服删除：<br/>
1.不属于会员照片，比如动物、风景、卡通等图片。<br/>
2.明星照等其他一切此类型的照片。<br/>
3.重复或者模糊不清的照片。<br/>
4.其他我们认为不符合要求的照片。<br/>
5.本站支持照片格式为jpg/gif/png。其他格式接收不到。<br/>
6.严禁上传黄色淫秽图片，网站留有上传者的IP等个人信息记录，以便有问题时配合相关部门做好取证，请友友们自重！<br/>
你可以上传你家人,或者朋友的相片,但一定要在主题或简介里注明，否则我们会认为不符合规定。<br/>
---------<br/>
<%End Sub%>

<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>

<%
Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p>
</card>
</wml>

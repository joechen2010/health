<!--#include file="../Conn.asp"-->
<!--#include file="Kesion.FileIcon.asp"-->
<%
'Dim S
'Set S=New WebFilesCls
'call S.execute("/","",20,"网站文件管理")
'Set S=nothing
				
Class WebFilesCls
        Private KS  
		Private MaxPerPage
		Private OpTypeStr,TopDir,action
		Private Fso,FsoFile,AllFileSize,WebDir
		Private CurrentDir,DirFiles,DirFolder,strTitle
		Private TotalPut,CurrentPage,TotalPages
        Private  ComeUrl,SQL,Rs,i,ChannelID
		
		Private Sub Class_Initialize()
		  MaxPerPage=30
			ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		'ChannelID-频道ID，DirStr--操作的目录,OpStr--操作类型(select 或 空),PerPage--每页显示的文件数,标题,样式文件-完整路径
		Function Kesion(CID,DirStr,OpStr,PerPage,Title,CssStr)
		   ChannelID=CID:strtitle=Title:TopDir=DirStr:MaxPerPage=PerPage:OpTypeStr=OpStr
		%>
				<html>
				<head>
				<title>文件管理</title>
				<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
				<link href="<%=CssStr%>" rel="stylesheet" type="text/css">
				<SCRIPT language=javascript src="<%=KS.GetDomain%>KS_Inc/showtitle.js"></script>
				<base target="_self">
				</head>
				<body leftmargin="0" topmargin="0">
				<%
				webDir=KS.Setting(3)
				TopDir=Replace(WebDir&Topdir,"//","/")
				action=LCase(Trim(KS.G("action")))
				CurrentDir=Trim(Replace(KS.G("CurrentDir"),"../",""))
				CurrentPage=KS.ChkClng(KS.G("page"))
				
				if CurrentDir<>"" then
					CurrentDir=Replace(CurrentDir & "/","//","/")
				end if
				if instr(currentdir,".")<> 0 or instr(topdir,".")<>0 then
				  response.write "<script>alert('非法路径');window.close();</script>"
				  response.end
				end if
				Set Fso=KS.InitialObject(Trim(KS.Setting(99)))
				Select Case action
				Case "del"
					Call DelAll
				Case "rname"
					Call Rname
				Case Else
					Call Main
				End Select
			
				Set Fso = Nothing
				
				%>
				<br>
				</body>
				</html>
				<%
				End Function
				
				Sub Main()
				  on error resume next
				 'response.write topdir
				' response.end
					Set FsoFile = Fso.GetFolder(Server.MapPath(TopDir))
						if Err then
							Set	FsoFile = Nothing
							Response.Write "找不到目录，可能参数配置错误！"
							Response.End
						end if
						AllFileSize = FsoFile.size
					Set	FsoFile = Nothing

				
					Set	FsoFile = Fso.GetFolder(Server.MapPath(TopDir & CurrentDir))
					Dim FolderNuns,FileNums
					FolderNuns=FsoFile.SubFolders.count
					FileNums=FsoFile.Files.count
					TotalPut=FolderNuns+FileNums
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					Dim j,k
					j=0
				%>
				<script language=javascript>
				function Checked()
				{
					var j = 0
					for(i=0;i < document.form.elements.length;i++){
						if(document.form.elements[i].name == "FileId" || document.form.elements[i].name == "FolderId"){
							if(document.form.elements[i].checked){
								j++;
							}
						}
					}
					return j;
				}
				function CheckAll1()
				{
					for(i=0;i<document.form.elements.length;i++)
					{
						if(document.form.elements[i].checked){
							document.form.elements[i].checked=false;
							document.form.CheckAll.checked=false;
						}
						else{
							document.form.elements[i].checked = true;
							document.form.CheckAll.checked = true;
						}
					}
				}
				function DelAll()
				{
					if(Checked()  <= 0){
						alert("您必须选择其中的一个文件或文件夹");
					}	
					else{
						if(confirm("确定要删除选择的文件或文件夹么？\n此操作不可以恢复！")){
							form.action="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Del&OpTypeStr=<%=OpTypeStr%>";
							form.submit();
						}
					}
				}
				function Rname()
				{
					if(Checked() == 0){
						alert("您必须选择一个文件或文件夹");
					}
					else{
						if(Checked() != 1){
							alert("只能选择一个文件或一个文件夹");
						}
						else{
							for(i=0;i < document.form.elements.length;i++){
								if(document.form.elements[i].name == "FolderId" && document.form.elements[i].checked){
									var j = prompt("请输入新文件夹名",document.form.elements[i].value)
									break;
								}
								else if(document.form.elements[i].name == "FileId" && document.form.elements[i].checked){
									var j = prompt("请输入新文件名",document.form.elements[i].value.split(".")[0])
									break;
								}
							}
							if(j != "" && j != null){
								if(IsStr(j) == j.length){
									form.action="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Rname&OpTypeStr=<%=OpTypeStr%>&NewName=" + j;
									form.target="_self";
									form.submit();
								}
								else{
									alert("新名称不符合标准，只能是字母、数字和下划线的组合,\n不能含有汉字、空格、点和其他符号");
								}
							}
						}
					}
				}
				function IsStr(w)
				{
					var str = "abcdefghijklmnopqrstuvwxyz_1234567890"
					 w = w.toLowerCase();
					var j = 0;
					for(i=0;i < w.length;i++){
						if(str.indexOf(w.substr(i,1)) != -1){
							j++;
						}
					}
					return j;
				}
				function setReturn(v)
				{
				  if (document.all)
				  {
				  window.returnValue=v;
				  }else
				  { 
				   parent.window.opener.setVal(v);
				  }
				  top.close();
				}
				</script>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
				 <tr class="Title"> 
				  <td align="center" colspan="2"><B><%=strTitle%></B></td>
				 </tr>
				 <tr class="Title2" height=23> 
				  <td>主目录占用空间：<font color="#ff0000"><%=GetSize(AllFileSize,"b")%></font></td><td align="right">&nbsp;&nbsp;<a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Main&OpTypeStr=<%=OpTypeStr%>&CurrentDir=" title=返回到主目录><font color=FF0000>返回主目录</font></a></td>
				 </tr>
				 <tr height=23> 
				  <td style="border-bottom:1px dashed #a7a7a7">当前目录：<%=TopDir%><%=CurrentDir%>&nbsp;&nbsp;&nbsp;&nbsp;占用空间：<font color="#ff0000"><%=GetSize(FsoFile.size,"b")%></font>&nbsp;&nbsp;文件夹：<font color=blue><%=FolderNuns%></font>&nbsp;个，文件：<font color=blue><%=FileNums%></font>&nbsp;个</td>
				  <td style="border-bottom:1px dashed #a7a7a7" align="right" width="80"><a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Main&OpTypeStr=<%=OpTypeStr%>&CurrentDir=<%=GetUpDir%>"><font color=FF0000>↑上一目录</font></a></td>
				 </tr>
				</table>
				<br>
				<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
					<form name="form" method="post" >
					 <tr class="title">
					  <td width="48" height="25" align="center">选中</td>
					  <td width="318" align="center">文件/文件夹名</td>
					  <td width="197" align="center">文件大小</td>
					  <td width="178" align="center">最后修改时间</td>
					  <td width="198" align="center">可用操作</td>
					 </tr>
					 <%
					For Each DirFolder in FsoFile.SubFolders%>
					 <tr bgcolor="#ffffff" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
					  <td height="22" align="center"> 
					   <input type="checkbox" name="FolderId" value="<%=DirFolder.name%>"></td>
					  <td>&nbsp;<a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Main&OpTypeStr=<%=OpTypeStr%>&CurrentDir=<%=CurrentDir & DirFolder.name%>"><img src="<%=WebDir%>KS_Editor/images/FileIcon/folder.gif" border=0 width="16" height="16" align="absmiddle"></a>&nbsp;<a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Main&OpTypeStr=<%=OpTypeStr%>&CurrentDir=<%=CurrentDir & DirFolder.name%>"><%=DirFolder.name%></a></td>
					  <td width="197" align="center"><%=GetSize(DirFolder.size,"b")%></td>
					  <td align="center" nowrap>&nbsp;<%=DirFolder.DateLastModified%></td>
					  <td width="198" align="center"><a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Main&OpTypeStr=<%=OpTypeStr%>&CurrentDir=<%=CurrentDir & DirFolder.name%>">打开</a></td>
					 </tr>
					 <tr><td colspan=6 background='images/line.gif'></td></tr>
					 <%
					Next
				
					For Each DirFiles in FsoFile.Files
					k=k+1
					if j>=MaxPerPage then
						exit for
					elseif k>MaxPerPage*(CurrentPage-1) then
					%>
					 <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
					  <td height="22" align="center"> 
					   <input type="checkbox" name="FileId" value="<%=DirFiles.name%>"></td>
					  <td>&nbsp;
					  <%if OpTypeStr="select" then%>
					  <a href="#" title="<table width=80 border=0 align=center><tr><td><img src='<%=TopDir & CurrentDir & DirFiles.name%>' border=0 width='130' height='80'></td></tr></table>" onClick="setReturn('<%=TopDir & CurrentDir & DirFiles.name%>')">
					  <%else%>
					  <a title="<table width=80 border=0 align=center><tr><td><img src='<%=TopDir & CurrentDir & DirFiles.name%>' border=0 width='130' height='80'></td></tr></table>" href="<%=TopDir & CurrentDir & DirFiles.name%>" target="_blank">
					  <!--<a href="<%=TopDir & CurrentDir & DirFiles.name%>" target="_blank">-->
					  <%end if%>
					  <img src="<%=WebDir%>KS_Editor/images/FileIcon/<%=GetFileIcon(DirFiles.name)%>" border=0 width="16" height="16" align="absmiddle" alt="<%=DirFiles.type%>">&nbsp;<%=DirFiles.name%></a></td>
					  <td width="197" align="center"><%=GetSize(DirFiles.size,"b")%></td>
					  <td align="center" nowrap><%=DirFiles.DateLastModified%></td>
					  <td width="198" align="center">
					  <%if OpTypeStr="select" then%>
					  <a href="#" onClick="setReturn('<%=TopDir & CurrentDir & DirFiles.name%>');">选择</a>
					  <%else%>
					  <a href="<%=TopDir & CurrentDir & DirFiles.name%>" target="_blank">浏览</a> | <a href="?ChannelID=<%=ChannelID%>&topdir=<%=topdir%>&action=Del&OpTypeStr=<%=OpTypeStr%>&CurrentDir=<%=CurrentDir%>&FileId=<%=DirFiles.name%>" onClick="return confirm('确定要删除选择的文件么？\n此操作不可以恢复！')">删除</a> 
					  <%end if%>
					  </td>
					 </tr>
					  <tr><td colspan=6 background='images/line.gif'></td></tr>
					 <%
					j=j+1
					end if
					Next
					if OpTypeStr<>"select" then
					%>
					 <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
					  <td height="25" align="center" > 
					   <input type="checkbox" name="CheckAll" value="checkbox" onClick="CheckAll1()" title=全部选择 style="cursor:pointer"></td>
					  <td height="30" colspan="5">&nbsp;
					   <input type="button" name="Submit" value="重命名" class=button onClick="Rname()"  title=重命名>
					   <input type="button" name="Submit2" value=" 删 除" class=button onClick="DelAll()"  title=删除>
					   <input type="hidden" name="CurrentDir" value="<%=CurrentDir%>">
					  </td>
					 </tr>
					  <tr><td colspan=6 background='images/line.gif'></td></tr>
					<%end if%>
				
					</form>
					<tr> 
					  <td colspan="6" height="25" align="right" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
						<%
						Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "个文件", CurrentPage, "channelid=" & channelid & "&action=Main&OpTypeStr="&OpTypeStr&"&CurrentDir="&CurrentDir)
						%>
					  </td>
				  </tr> 
				</table>
				<%
				Set FsoFile = Nothing
				End Sub

				Public Function GetSize(size,unit)
						if isEmpty(size) or Not Isnumeric(size) then Exit Function
						size=CheckUnit(size,unit)
						if size>1024 then
							size=(size/1024)
							getsize=formatnumber(size,2) & " MB"
						else
							getsize=size & " KB"
							Exit Function
						end if
						if size>1024 then
							size=(size/1024)
							getsize=formatnumber(size,2) & " GB"
						end if
					End Function
					Public Function CheckUnit(size,unit)
						Select Case Lcase(Unit)
						Case "b"
							CheckUnit = formatnumber(size/1024,2)
						Case "k"
							CheckUnit = size
						Case "m"
							CheckUnit = (size*1024)
						Case "g"
							CheckUnit = (size*1024*1024)
						Case Else
							CheckUnit = size
						End Select
					End Function
					Public Sub DelFiles(strFiles)
						if strFiles="" then Exit Sub
						dim fso,arrFiles,i
						On Error Resume Next
						Err=0
						Set fso = KS.InitialObject(Trim(KS.Setting(99)))
							if fso.FileExists(server.MapPath(strFiles)) then
								fso.DeleteFile(server.MapPath(strFiles))
								if 0=Err then
									Response.Write "<br>清除文件（"&strFiles&"）成功！"
								else
									Response.Write "<br>清除文件（"&strFiles&"）失败！"
								end if
							end if
						Set fso = Nothing
						Err=0
					End Sub
					Function GetUpDir()
					Dim strDir,strDir2,i
					strDir=""
					If CurrentDir = "" then Exit Function
					strDir2=CurrentDir
					strDir2=Split(strDir2,"/")
					for i=0 to Ubound(strDir2)-1
						if i<Ubound(strDir2)-1 then strDir=strDir & strDir2(i) & "/"
					next
					GetUpDir=strDir
				End Function
				
				Sub DelAll()
					Dim FolderId,FileId,FileNum,FolderNum,FilePath,FolderPath
					Dim FsoFolder,sSize
					FolderId = Split(Request.Form("FolderId"),",")
					FileId = Trim(Request("FileId"))
					FileNum=0
					FolderNum=0
					If instr(FileId,",")>0 then 
						FileId = Split(FileId,",")
						If Ubound(FileId) > -1 then
							For i = 0 to Ubound(FileId)
								FilePath = Server.MapPath(TopDir & CurrentDir & Trim(FileId(i)))
								If Fso.FileExists(FilePath) then
									Fso.DeleteFile FilePath,true
									FileNum = FileNum + 1
								End If
							Next
						End If
					else
						FilePath = Server.MapPath(TopDir & CurrentDir & FileId)
						If Fso.FileExists(FilePath) then
							Fso.DeleteFile FilePath,true
							FileNum = FileNum + 1
						End If
					end if
					If Ubound(FolderId) > -1 then
						For i = 0 to Ubound(FolderId)
							FolderPath = Server.MapPath(TopDir & CurrentDir & Trim(FolderId(i)))
							If Fso.FolderExists(FolderPath) then
								Set FsoFolder = Fso.GetFolder(FolderPath)
								if FsoFolder.size <=0 then
								Fso.DeleteFolder FolderPath,true
								FolderNum = FolderNum + 1
								end if
							End If
						Next
					End If
					Response.Write("<script>alert('\n成功删除 "& FileNum &" 个文件\n成功删除 "& FolderNum &" 个文件夹');location.href='" & ComeUrl & "';</script>")
				End Sub
				Sub Rname()
					Dim FolderName,FileName,NewName,OldName,strNewName
					FolderName = Trim(Request("FolderId"))
					FileName = Trim(Request("FileId"))
					NewName = Trim(Request("NewName"))
					if instr(newname,".")<>0 then
					  Call  Response.Write("<script>alert('输入的文件名不符合规范！');location.href='" & ComeUrl & "';</script>")
					  Response.End
					end if
					if NewName="" then Call  Response.Write("<script>alert('请输入新文件名！');location.href='" & ComeUrl & "';</script>")
					If len(FolderName) <> 0 then
						strNewName = Server.MapPath(TopDir & CurrentDir & NewName)
						OldName = Server.MapPath(TopDir & CurrentDir & FolderName)
						If not Fso.FolderExists(strNewName) then
							Fso.MoveFolder OldName,strNewName
							Response.Write("<script>alert('文件夹“"& FolderName &"”已经成功改名为“"& NewName &"”');location.href='" & ComeUrl & "';</script>")
						Else
							 Response.Write("<script>alert('有同名文件夹，请换个文件夹名！');location.href='" & ComeUrl & "';</script>")
						End If
					End If
					If len(FileName) <> 0 then
						Dim FileExt,NewFileExt
						'Response.Write FileName
						FileExt=Split(FileName,".")
						NewFileExt=Trim(FileExt(Ubound(FileExt)))
						if Instr(NewName,".")>0 then
							Response.Write("<script>alert('文件名中不能带“.”，换个文件名！');location.href='" & ComeUrl & "';</script>")
							Response.End
						end if
						NewName=NewName & "." & NewFileExt
						strNewName = Server.MapPath(TopDir & CurrentDir & NewName)
						OldName = Server.MapPath(TopDir & CurrentDir & FileName)
						If not Fso.FileExists(strNewName) then
							Fso.MoveFile OldName,strNewName
							Response.Write("<script>alert('文件“"& FileName &"”已经成功改名为“"& NewName &"”');location.href='" & ComeUrl & "';</script>")
							
						Else
							Response.Write("<script>alert('有同名文件，请换个文件名!');location.href='" & ComeUrl & "';</script>")
							
						End If
					End If
				End Sub
End Class
%> 

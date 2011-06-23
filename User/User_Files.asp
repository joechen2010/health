<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.WebFilesCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Frame
        Private KS,KSUser
		Private TopDir
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
		  Response.Write "<script>window.close();</script>"
		  Exit Sub
		End If
		TopDir=KSUser.GetUserFolder(ksuser.username)
		if KS.S("action")="show" then
		  call showframe()
		else
		  call filelist()
		end if
		end sub
		
		sub showframe()
		 Call KSUser.Head()
		 Call KSUser.InnerLocation("我的文件管理")
		 Call KS.CreateListFolder(TopDir)
        %>
		
		<div class="tabs">	
			<ul>
	        <li class="select">我的文件夹</li>
			
			</ul>
		</div>						  
		


<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="1">
                                <tr>
												<td height='25' align='center'>
												
												<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" style="display:nowrap">
						<tr class="tdbg">
						<td width="157" align="right">您的总空间 <font color=red><%=round(KSUser.SpaceSize/1024,2)%>M</font>,使用情况：</td>
						<td><img src="images/bar.gif" width="0" height="16" id="Sms_bar" align="absmiddle" /></td>
						<td width="211"  align="center" id="Sms_txt">100%</td>
						</tr></table>
		 <%
        response.write showtable("Sms_bar","Sms_txt",KS.GetFolderSize(TopDir)/1024,KSUser.SpaceSize)
%>
												
												</td></tr>
												<tr class='tdbg'>
												 <td>
												 <div id="rssbody" style="overflow-y:scroll;height:500; width:100%;"> 
				                               <iframe src="user_files.asp" style="width:100%;height:100%" frameborder="0"></iframe>
						<div></div>
						    </div>
						   </td>
	                      </tr>
                        </table>
						 <div style="padding:8px;color:red">温馨提醒：为免浪费您的保贵空间，请及时删除无用的文件！</div>
						</div>
		<%
		end sub
		
		sub filelist()
		 Response.Buffer = True
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"
		Dim WFCls:Set WFCls = New WebFilesCls
		Call WFCls.Kesion(0,TopDir,"",20,"","Images/Css.css")
		Set WFCls = Nothing
	  
      End Sub
	   '（图片对象名称，标题对象名称，更新数，总数）
		Function ShowTable(SrcName,TxtName,str,c)
		Dim Tempstr,Src_js,Txt_js,TempPercent
		If C = 0 Then C = 99999999
		Tempstr = str/C
		TempPercent = FormatPercent(tempstr,0,-1)
		Src_js = "document.getElementById(""" + SrcName + """)"
		Txt_js = "document.getElementById(""" + TxtName + """)"
			ShowTable = VbCrLf + "<script>"
			ShowTable = ShowTable + Src_js + ".width=""" & FormatNumber(tempstr*600,0,-1) & """;"
			ShowTable = ShowTable + Src_js + ".title=""容量上限为："&c/1024&" MB，已用（"&FormatNumber(str/1024,2)&"）MB！"";"
			ShowTable = ShowTable + Txt_js + ".innerHTML="""
			If FormatNumber(tempstr*100,0,-1) < 80 Then
				ShowTable = ShowTable + "已使用:" & TempPercent & """;"
			Else
				ShowTable = ShowTable + "<font color=\""red\"">已使用:" & TempPercent & ",请赶快清理！</font>"";"
			End If
			ShowTable = ShowTable + "</script>"
		End Function
		
End Class
%> 

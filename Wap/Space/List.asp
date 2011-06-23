<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
Dim KSCls
Set KSCls = New List
KSCls.Kesion()
Set KSCls = Nothing

Class List
        Private KS,KSBCls,KSRFObj
		Private RS,ID
		Private UserName,UserType,Template,BlogName
		Private MaxPerPage,CurrentPage
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
		       Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			
			ID=KS.ChkClng(KS.S("ID"))
			UserName=KS.S("UserName")
			If UserName="" Then Response.End()
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Blog Where UserName='" & UserName & "'",Conn,1,1
			If RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
		       Call KS.ShowError("该用户没有开通空间站点！","该用户没有开通空间站点！")
			End If
			If RS("Status")=0 Then
			   RS.Close:Set RS=Nothing
			   Call KS.ShowError("该空间站点尚未审核！","该空间站点尚未审核！")
			ElseIf RS("Status")=2 Then
			   RS.Close:Set RS=Nothing
			   Call KS.ShowError("该空间站点已被管理员锁定！","该空间站点已被管理员锁定！")
			End If
			BlogName=RS("BlogName")
		    UserType=KS.ChkClng(Conn.Execute("Select UserType From KS_User Where UserName='" & UserName & "'")(0))
			Dim MainTemplate,TemplateSub
			If UserType=1 Then
			   MainTemplate=KSRFObj.LoadTemplate(KS.WSetting(23))'企业主模板
			   TemplateSub=KSRFObj.LoadTemplate(KS.WSetting(27))'企业日志副模板
			Else
		       MainTemplate=KSRFObj.LoadTemplate(KS.WSetting(20))'个人主模板
			   TemplateSub=KSRFObj.LoadTemplate(KS.WSetting(28))'个人日志副模板
			End If
			MainTemplate=KSRFObj.KSLabelReplaceAll(MainTemplate)
			MainTemplate=KSBCls.ReplaceBlogLabel(RS,MainTemplate)
			MainTemplate=KSBCls.ReplaceAllLabel(UserName,MainTemplate)
			RS.Close

			If KSUser.GroupID<>4 Then
			   RS.Open "Select * from KS_BlogInfo Where ID=" & ID & " And Status=0",Conn,1,3
			Else
			   RS.Open "Select * from KS_BlogInfo Where ID=" & ID,Conn,1,3
			End If
			If RS.EOF And RS.BOF Then
			   Call KS.ShowError("参数传递出错或该日志为草稿！","参数传递出错或该日志为草稿！")
			End If
			RS("Hits")=RS("Hits")+1
			RS.Update
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""" & BlogName & "-" & RS("Title") & """>" &vbcrlf
			Template=Template & MainTemplate
			Template=Replace(Template,"{$BlogMain}","" & ReplaceLabel(TemplateSub,RS) & "")
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
			RS.Close:Set  RS=Nothing
		End Sub
		
		Function ReplaceLabel(Byval Template,RS)
		    If KS.S("Action")="CommentSave" Then 
	    	   Dim HomePage,InsertFace,Content,Anonymous,Title
			   If KSUser.UserLoginChecked = True Then
			      AnounName=KSUser.UserName
			   Else
			      AnounName="游客"
			   End If
			   HomePage="http://wap.kesion.com/"
			   InsertFace=KS.S("InsertFace")
			   Content=KS.S("Content")
			   Title=KS.S("Title")
			   If Title="" Then Title="回复本文主题"
			   If AnounName="" Then 
			      ReplaceLabel="<br/><br/>请填写你的昵称!<br/><anchor><prev/>返回重写</anchor><br/><br/>"
				  Exit Function
			   End if
			   If Content="" Then 
			      ReplaceLabel="<br/><br/>请填写评论内容!<br/><anchor><prev/>返回重写</anchor><br/><br/>"
				  Exit Function
			   End if
			   Dim RSS:Set RSS=Server.CreateObject("ADODB.RECORDSET")
			   RSS.Open "Select * From KS_BlogComment",Conn,1,3
			   RSS.AddNew
			   RSS("LogID")=ID
			   RSS("AnounName")=AnounName
			   RSS("Title")=Title
			   RSS("UserName")=KS.S("UserName")
			   RSS("HomePage")=HomePage
			   RSS("Content")=InsertFace&Content
			   RSS("UserIP")=KS.GetIP
			   RSS("AddDate")=Now
			   RSS.UpDate
			   RSS.Close:Set RSS=Nothing
			   Template=Replace(Template,"{$ShowLogWriteComment}","你的评论发表成功!")
			End if
			Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/Images/face/" & RS("Face") & ".gif"" alt=""""/>"
			Dim TagList,TagsArr:TagsArr=Split(RS("Tags")," ")
			If RS("Tags")<>"" Then
			   TagList="<b>标签：</b>"
			   For I=0 To Ubound(TagsArr)
			       If TagsArr(i)<>"" Then
				      TagList=TagList &"<a href=""Blog.asp?UserName=" & UserName & "&amp;Tag=" & TagsArr(i) &"&amp;" & KS.WapValue & """>" & TagsArr(i) & "</a> "
				   End If
			   Next
			   TagList=TagList &"<br/>"
			End If
			
		   	Dim ContentStr
			Dim JFStr:If RS("Best")="1" Then JFStr="  <img src=""../images/jh.gif"" alt=""""/>" Else JFStr=""

		    If IsNull(RS("Password")) Or RS("PassWord")="" Then 
			   ContentStr=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RS("Content"))))))
			ElseIf KS.S("Pass")<>"" Then
			   If KS.S("Pass")=RS("password") Then
			      ContentStr=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RS("Content"))))))
			   Else
			      ReplaceLabel="<br/><br/>出错啦,您输入的日志密码有误!<br/><br/>"
				  Exit Function
			   End if
		    Else
		       ReplaceLabel="<br/><br/>请输入日志的查看密码：<input name=""Pass"" maxlength=""30"" value="""" emptyok=""false""/><a href=""List.asp?ID="&ID&"&amp;UserName="&UserName&"&amp;Pass=$(Pass)&amp;" & KS.WapValue & """>查看</a><br/><br/>"
			   Exit Function
		    End IF
			Template=Replace(Template,"{$ShowLogFace}",EmotSrc)'仅显示日志心情
			Template=Replace(Template,"{$ShowLogTitle}",RS("Title"))'仅显示日志标题
			Template=Replace(Template,"{$ShowLogBest}",JFStr)
			Template=Replace(Template,"{$ShowLogUserName}",RS("UserName"))'仅显示日志作者
			Template=Replace(Template,"{$ShowLogAddDate}",KS.DateFormat(Rs("AddDate"),17))'
			Template=Replace(Template,"{$ShowLogText}",KS.ReplaceInnerLink(Replace(KS.ContentPagination(ContentStr,200,"List.asp?UserName="&UserName&"&amp;ID="&ID&"&amp;Pass="&KS.S("Pass")&"&amp;" & KS.WapValue & "",False,False),"&","&amp;")))'正文
			Template=Replace(Template,"{$ShowLogTags}",TagList)'标签
			Template=Replace(Template,"{$ShowLogHits}",RS("Hits"))'阅读次数
			Template=Replace(Template,"{$ShowLogReturn}",Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &RS("id"))(0))'回复数
			Template=Replace(Template,"{$ShowLogPrev}",ReplacePrevNextArticle(UserName,RS("ID"),"Prev"))'上一篇
			Template=Replace(Template,"{$ShowLogNext}",ReplacePrevNextArticle(UserName,RS("ID"),"Next"))'下一篇
			Template=Replace(Template,"{$ShowWeather}",KSBCls.GetWeather(RS))'仅显示日志天气
			Template=Replace(Template,"{$ShowLogContent}",ShowBlogComment)
			Template=Replace(Template,"{$ShowLogWriteComment}",GetWriteComment)  
		    ReplaceLabel=Template
		End Function
		
		Function ShowBlogComment()
		    MaxPerPage = 5
			If KS.S("Page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("Page"))
			Else
			   CurrentPage = 1
			End If
			Dim RSP:set RSP=Server.Createobject("adodb.recordset")
			RSP.Open "Select * from KS_BlogComment Where LogID="&ID&" order by AddDate DESC",Conn,1,1
			If RSP.EOF And RSP.BOF Then
			   ShowBlogComment="没有任何回复评论!<br/>"
			Else
			   Dim TotalPut:TotalPut = RSP.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSP.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Do While Not RSP.EOF
			      ShowBlogComment=ShowBlogComment & ReplaceFace(RSP("Content"))
				  If RSP("UserName")="游客" Then
				     ShowBlogComment=ShowBlogComment & RSP("UserName") &"("&KS.DateFormat(RSP("AddDate"),17)&")<br/>"
				  Else 
				     ShowBlogComment=ShowBlogComment & "<a href=""Space.asp?UserName=" & RSP("UserName") & "&amp;" & KS.WapValue & """>" & KS.GetUserRealName(RSP("UserName")) & "</a>("&KS.DateFormat(RSP("AddDate"),17)&")<br/>"
				  End If
				  If Not IsNull(RSP("Replay")) or RSP("Replay")<>"" Then
				     ShowBlogComment=ShowBlogComment & "主人回复:"&KS.LoseHtml(RSP("Replay"))&"<br/>" 
				  End If
			      RSP.Movenext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   ShowBlogComment=ShowBlogComment & KS.ShowPagePara(TotalPut, MaxPerPage, "List.asp", False, "个", CurrentPage, "UserName="&UserName&"&amp;ID="&ID&"&amp;Pass="&KS.S("Pass")&"&amp;" & KS.WapValue & "")
			   ShowBlogComment=ShowBlogComment & "<br/>"
			End If
			RSP.close:set RSP=nothing
		End Function
		
		Function ReplaceFace(C)
		    Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
			Dim strArr:strArr=Split(str,"|")
			Dim K
			For K=0 To 19
			    C=Replace(C,"[e"&K &"]","<img src=""" & KS.Setting(3) & "Images/Emot/" & K & ".gif"" alt=""""/>")
			Next
			ReplaceFace=C
		End Function
		
		Function GetWriteComment()
		    Dim k,str:str="惊讶|撇嘴|色色|发呆|得意|流泪|害羞|闭嘴|睡觉|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
			Dim strArr:strArr=Split(str,"|")
			GetWriteComment = "<select name=""InsertFace"">"
			GetWriteComment = GetWriteComment & "<option value="""">无</option>"
			For k=0 to 19 
			    GetWriteComment = GetWriteComment & "<option value=""[e"&K&"]"">" & strArr(k) & "</option>"
			Next
			GetWriteComment = GetWriteComment & "</select> "
			Dim reSayArry:reSayArry = Array("好帖，要顶!","看帖回帖是美德!","你牛，我顶!","这帖不错，该顶!","支持你!","反对你!")
			Randomize
			GetWriteComment = GetWriteComment & "<input name=""Content" & Minute(Now) & Second(Now) & """ type=""text"" maxlength=""500"" size=""20"" value="""&reSayArry(Int(Ubound(reSayArry)*Rnd))&"""/> "
			GetWriteComment = GetWriteComment & "<anchor>提交<go href=""List.asp?Action=CommentSave&amp;UserName=" & UserName & "&amp;ID=" & ID & "&amp;Pass="&KS.S("Pass")&"&amp;" & KS.WapValue & """ method=""post"">"
            GetWriteComment = GetWriteComment & "<postfield name='AnounName' value='$(AnounName" & Minute(Now) & Second(Now) & ")'/>"
            GetWriteComment = GetWriteComment & "<postfield name='HomePage' value='$(HomePage)'/>"
            GetWriteComment = GetWriteComment & "<postfield name='InsertFace' value='$(InsertFace)'/>"
            GetWriteComment = GetWriteComment & "<postfield name='Content' value='$(Content" & Minute(Now) & Second(Now) & ")'/>"
            GetWriteComment = GetWriteComment & "</go></anchor><br/>"
		End Function

		Function ReplacePrevNextArticle(UserName,NowID,TypeStr)
		    Dim SqlStr
			If Trim(TypeStr) = "Prev" Then
			   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where UserName='" & UserName & "' And ID<" & NowID & " And Status=0 Order By ID Desc"
			ElseIf Trim(TypeStr) = "Next" Then
			   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where UserName='" & UserName & "' And ID>" & NowID & " And Status=0 Order By ID Desc"
			Else
			   ReplacePrevNextArticle = ""
			   Exit Function
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			RS.Open SqlStr, Conn, 1, 1
			If RS.EOF And RS.BOF Then
			   ReplacePrevNextArticle = "没有了"
			Else
			  ReplacePrevNextArticle = "<a href=""" & KSBCls.GetCurrLogUrl(RS("ID"),UserName) & """ title=""" & RS("Title") & """>" & RS("title") & "</a>"
			End If
			RS.Close:Set RS = Nothing
		End Function
End Class
%>
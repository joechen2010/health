<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS,KSBCls
Set KS=New PublicCls
Set KSBCls=New BlogCls
Dim ID,TemplateID,RS,CommentStr,N
Dim totalPut, CurrentPage, MaxPerPage,PageNum,SqlStr
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

Call show
Sub Show()
    ID=KS.ChkClng(KS.S("ID"))
	MaxPerPage=5    '每页显示评论条数
	If KS.S("page") <> "" Then
		CurrentPage = KS.ChkClng(KS.S("page"))
	Else
		 CurrentPage = 1
	End If

	SqlStr="Select * From KS_BlogComment Where LogID=" & ID & " Order By AddDate Desc"
	 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	 RS.Open SqlStr,Conn,1,1

  IF Not Rs.Eof Then
		 totalPut = RS.RecordCount
				If CurrentPage < 1 Then	CurrentPage = 1
							If (totalPut Mod MaxPerPage) = 0 Then
									PageNum = totalPut \ MaxPerPage
							Else
									PageNum = totalPut \ MaxPerPage + 1
						   End If
		
				If CurrentPage = 1 Then
						Call showContent(rs)
				Else
						If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									Call showContent(rs)
						Else
									CurrentPage = 1
									Call showContent(rs)
				        End If
				End If
  End If
  Response.Write "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|条||2"
  rs.close:set rs=nothing
End Sub

Sub ShowContent(rs)
 CommentStr="&nbsp;&nbsp;本文有 <font color=red>" & totalPut & " </font> 条评论，共分 <font color=red>" & pagenum & "</font> 页,第 <font color=red>" & CurrentPage & "</font> 页<br />"
    CommentStr=CommentStr & "<table  width='99%' border='0' align='center' cellpadding='0' cellspacing='1'>"
    If CurrentPage=1 Then
	 N=TotalPut
	 Else
	 N=totalPut-MaxPerPage*(CurrentPage-1)
	 End IF
  Dim FaceStr,Publish,i
  Do While Not RS.Eof 
   FaceStr=KS.Setting(3) & "images/face/0.gif"

    Publish=RS("AnounName")
	If not Conn.Execute("Select UserFace From KS_User Where UserName='"& Publish & "'").eof Then
      FaceStr=Conn.Execute("Select UserFace From KS_User Where UserName='"& Publish & "'")(0)
	  If lcase(left(FaceStr,4))<>"http" then FaceStr=KS.Setting(3) & FaceStr
   End IF

	
   CommentStr=CommentStr & "<tr>"
   CommentStr=CommentStr & "<td width='70' rowspan='3' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><img width=""60"" height=""60"" src=""" & facestr & """ border=""1""></td>"
   CommentStr=CommentStr & "<td height='25' width=""70%"">"
   CommentStr=CommentStr & RS("Title")
   CommentStr=CommentStr  & "  </td><td width=""30"" align=""right""><font style='font-size:32px;font-family:""Arial Black"";color:#EEF0EE'> " & N & "</font></td>"
   CommentStr=CommentStr & "</tr>"
   CommentStr=CommentStr & "<tr>"
   CommentStr=CommentStr & "<td height='25' colspan='2'>" & ReplaceFace(RS("Content"))
   		 If Not IsNull(RS("Replay")) Or Rs("Replay")<>"" Then
		 CommentStr=CommentStr&"<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><b>以下为space主人的回复:</b><br>" & RS("Replay") & "<br><div align=right>时间:" & rs("replaydate") &"</div></div>"
         End If
   CommentStr=CommentStr & "	 </td>"
   CommentStr=CommentStr & "</tr>"
   CommentStr=CommentStr & "<tr>"
   
   			 Dim MoreStr,KSUser,LoginTF
			 Set KSUser=New UserCls
			 LoginTF=Cbool(KSUser.UserLoginChecked)
			 IF LoginTF=true and KSUser.UserName=RS("UserName") Then
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> | <a href='../User/user_message.asp?Action=CommentDel&id=" & RS("ID") & "' onclick=""return(confirm('确定删除该留言吗?'));"">删除</a> | <a href='../User/?User_message.asp?id=" & RS("ID") & "&Action=ReplayComment' target='_blank'>回复</a>"
			 Else
                 MoreStr="<a href='" & RS("HomePage") & "' target='_blank'>主页</a>| <a href='#'>顶部</a> "
			 End If
			 Set KSUser=Nothing

   CommentStr=CommentStr & "<td align='right' colspan='2' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><font color='#999999'>(" & publish & " 发表于：" & RS("AddDate") &")</font>&nbsp;&nbsp;" & MoreStr & " </td>"
   CommentStr=CommentStr & "</tr>"
   N=N-1
   RS.MoveNext
		I = I + 1
	  If I >= MaxPerPage Then Exit Do
  loop
 CommentStr=CommentStr & "</table>"

 response.write CommentStr
End Sub

Function ReplaceFace(c)
		 Dim str:str="惊讶|撇嘴|色|发呆|得意|流泪|害羞|闭嘴|睡|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title=""" & strarr(k) & """ src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">")
		 Next
		 ReplaceFace=C
End Function

Call CloseConn()
Set KS=Nothing
Set KSBCls=Nothing
%>

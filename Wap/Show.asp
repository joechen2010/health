<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：内容页
'* 演示地址: http://wap.kesion.com/
'********************************
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New InfoCls
KSCls.Kesion()
Set KSCls = Nothing

Class InfoCls
        Private KS,KSRFObj
		Private RS,SQLStr,DomainStr,UserLoginTF,ID,ChannelID,PayTF
		Private InfoPurview,ReadPoint,ChargeType,PitchTime,ReadTimes,ClassPurview,UserName
		Private strContent,FileContent
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSRFObj = New Refresh
		End Sub
		'Call Kesion()
        Private Sub Class_Terminate()
		    Call CloseConn()
			Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
			strContent=false
		    DomainStr=KS.GetDomain
			UserLoginTF=Cbool(KSUser.UserLoginChecked)
			ID=KS.ChkClng(KS.S("ID"))
			ChannelID=KS.ChkClng(KS.S("ChannelID"))
			PayTF=KS.S("PayTF")
			IF ID=0 Then Exit Sub
			Select Case KS.C_S(ChannelID,6)
			    Case 1
				SqlStr= "Select top 1 a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join KS_Class b on a.tid=b.id Where a.ID=" & ID
				Case 2
				SqlStr= "Select top 1 a.*,ClassPurview,ClassID,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID
				Case 3
				SqlStr= "Select top 1 * from "&KS.C_S(ChannelID,2)&" Where ID=" & ID
				Case 5
				SqlStr= "Select top 1 * from "&KS.C_S(ChannelID,2)&"  Where verific=1 And ID=" & ID
				Case 7
				SqlStr= "Select * from KS_Movie Where verific=1 And ID=" & ID
				Case 8
				SqlStr= "Select b.WapTemplateID,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.ID=" & ID
			End Select
			Set RS=Server.CreateObject("Adodb.Recordset")
			RS.Open SqlStr,Conn,1,3
			IF RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   Select Case KS.C_S(ChannelID,6)
				   Case "1","2","3"
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数！")
				   Case "5"
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要查看的" & KS.C_S(ChannelID,3) & "已删除或是未通过暂停销售！")
				   Case "7"
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要观看的影片已删除。或是没有通过审核！")
				   Case "8"
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要查看的信息已删除。或是您非法传递注入参数！")
			   End Select
			Else
			   Call FCls.SetContentInfo(ChannelID,RS("Tid"),RS("ID"))
			   Select Case KS.C_S(ChannelID,6)
			       '=======================================================
			       Case 1
				      If RS("Verific")<>1 And UserLoginTF=False And KSUser.UserName<>RS("Inputer") Then
					     Call KS.ShowError("系统提示！","对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核！")
					  End If
					  InfoPurview=Cint(RS("InfoPurview"))
					  ReadPoint=Cint(RS("ReadPoint"))
					  ChargeType=Cint(RS("ChargeType"))
					  PitchTime=Cint(RS("PitchTime"))
					  ReadTimes=Cint(RS("ReadTimes"))
					  ClassPurview=Cint(RS("ClassPurview"))
					  UserName=RS("Inputer")  
					  '增加用户查看文章次数
					  Conn.Execute("UPDATE " & KS.C_S(ChannelID,2) & " SET Hits=Hits+1 WHERE ID="&ID)
					  Call PowerLimit()
					  FileContent = KSRFObj.LoadTemplate(RS("WapTemplateID"))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  If InStr(FileContent,"[KS_Charge]")=0 Then
					     Dim HtmlLabel,HtmlLabelArr,I
						 HtmlLabel = KSRFObj.GetFunctionLabel(FileContent,"{=GetArticleContent")
						 HtmlLabelArr=Split(HtmlLabel,"@@@")
						 For I=0 To Ubound(HtmlLabelArr)
							 FileContent = Replace(FileContent,HtmlLabelArr(I),"[KS_Charge]"&HtmlLabelArr(I)&"[/KS_Charge]")
						 Next
					  End If
					  '替换文章内容页标签为内容
					  FileContent = KSRFObj.ReplaceNewsContent(ChannelID,RS, FileContent, "")
					  If strContent<>"True" Then
					     Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "[KS_Charge]", "[/KS_Charge]", 0)
						 FileContent=Replace(FileContent,"[KS_Charge]" & ChargeContent &"[/KS_Charge]",strContent)
					  Else
					     FileContent=Replace(FileContent,"[KS_Charge]","")
					     FileContent=Replace(FileContent,"[/KS_Charge]","")
					  End If
				   '=======================================================
				   Case 2
				      If RS("Verific")<>1 And UserLoginTF=False And KSUser.UserName<>RS("Inputer") Then
					     Call KS.ShowError("系统提示！","对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核！")
					  End If
					  InfoPurview=Cint(RS("InfoPurview"))
					  ReadPoint=Cint(RS("ReadPoint"))
					  ChargeType=Cint(RS("ChargeType"))
					  PitchTime=Cint(RS("PitchTime"))
					  ReadTimes=Cint(RS("ReadTimes"))
					  ClassPurview=Cint(RS("ClassPurview"))
					  RS("Hits") = RS("Hits") + 1
					  If DateDiff("D", RS("LastHitsTime"), Now()) <= 0 Then
					     RS("HitsByDay") = RS("HitsByDay") + 1
					  Else
					     RS("HitsByDay") = 1
					  End If
					  If DateDiff("ww", RS("LastHitsTime"), Now()) <= 0 Then
					     RS("HitsByWeek") = RS("HitsByWeek") + 1
					  Else
					     RS("HitsByWeek") = 1  
					  End If
					  If DateDiff("m", RS("LastHitsTime"), Now()) <= 0 Then
					     RS("HitsByMonth") = RS("HitsByMonth") + 1
					  Else
					     RS("HitsByMonth") = 1
					  End If
					  RS("LastHitsTime") = Now()
					  RS.Update
					  Call PowerLimit()
					  FileContent = KSRFObj.LoadTemplate(RS("WapTemplateID"))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  If Cbool(strContent)=true Then
						 FileContent = KSRFObj.ReplacePictureContent(ChannelID,RS, FileContent, GetPictureByPage(ID,ChannelID,RS("PicUrls")))
					  Else
					     FileContent = KSRFObj.ReplacePictureContent(ChannelID,RS, FileContent,"")
						 FileContent = Replace(FileContent,"{$GetPictureByPage}",strContent)
					  End If
				   '=======================================================
				   Case 3
					  FileContent = KSRFObj.LoadTemplate(RS("WapTemplateID"))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  FileContent = KSRFObj.ReplaceDownLoadContent(ChannelID,RS, FileContent)
				   '=======================================================
				   Case 5
					  FileContent = KSRFObj.LoadTemplate(RS("WapTemplateID"))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  FileContent = KSRFObj.ReplaceProductContent(ChannelID,RS, FileContent)
				   Case 7
					  FileContent = KSRFObj.LoadTemplate(RS("WapTemplateID"))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  FileContent = KSRFObj.ReplaceMovieContent(ChannelID,RS, FileContent) 
				   Case 8
					  FileContent = KSRFObj.LoadTemplate(RS(0))
					  FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
					  FileContent = KSRFObj.ReplaceGQContent(ChannelID,RS, FileContent)
			   End Select
			   FileContent = KS.GetEncodeConversion(FileContent)
			   Response.Write FileContent
			   RS.Close:Set RS=Nothing
			End If
		End Sub
		
		Sub PowerLimit()
		    If InfoPurview=2 or ReadPoint>0 Then
			   IF UserLoginTF=False Then
			      Call GetNoLoginInfo'登录
			   Else
			      IF KS.FoundInArr(RS("ArrGroupID"),KSUser.GroupID,",")=False and readpoint=0 Then
				     strContent="<br/><b>对不起，你所在的用户组没有查看本" & KS.C_S(ChannelID,3) & "的权限！</b><br/>"
				  Else
				     Call PayPointProcess()
			      End If
			   End If
			ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
		       If UserLoginTF=False Then
			      Call GetNoLoginInfo'登录
			   Else
			      '============继承栏目收费设置时,读取栏目收费配置===========
				  ReadPoint=Cint(RS("DefaultReadPoint"))   
				  ChargeType=Cint(RS("DefaultChargeType"))
				  PitchTime=Cint(RS("DefaultPitchTime"))
				  ReadTimes=Cint(RS("DefaultReadTimes"))
				 '============================================================
			      If ClassPurview=2 Then
			         IF KS.FoundInArr(RS("DefaultArrGroupID"),KSUser.GroupID,",")=False Then
				        strContent="<br/><b>对不起，你所在的用户组没有查看的权限！</b><br/>"
				     Else
				        Call PayPointProcess()
				     End If
			      Else    
			         Call PayPointProcess()
				  End If
			   End If
			Else
		       Call PayPointProcess()
			End If
		End Sub
		

	   '收费扣点处理过程
	   Sub PayPointProcess()
	       Dim UserChargeType:UserChargeType=KSUser.ChargeType
		   If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) And KSUser.UserName<>UserName Then
		      If UserChargeType=1 Then
			     Select Case ChargeType
				     Case 0:Call CheckPayTF("1=1")
					 Case 1:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime)
					 Case 2:Call CheckPayTF("Times<" & ReadTimes)
					 Case 3:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
					 Case 4:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
					 Case 5:Call PayConfirm()
				 End Select
			  Elseif UserChargeType=2 Then
			     If KSUser.GetEdays <=0 Then
				    strContent="<br/>对不起，你的账户已过期 "&KSUser.GetEdays&" 天,此文需要在有效期内才可以查看！<br/><br/>"
					strContent=strContent&"充值有效期方法<br/>"
					strContent=strContent&"1.请用购买到神州行充值卡充值,点击进入<a href=""User/User_CardOnline.asp?"&KS.WapValue&""">神州行充值...</a><br/>"
				 Else
				    Call GetContent()
				 End If
			  Else
			     Call GetContent()
			  End If
		   Else
		      Call GetContent()
		   End IF
	   End Sub

	   '检查是否过期，如果过期要重复扣点券
	   '返回值 过期返回 true,未过期返回false
	   Sub CheckPayTF(Param)
	       Dim SqlStr:SqlStr="Select top 1 Times From KS_LogPoint Where ChannelID=" & ChannelID & " And InfoID=" & ID & " And InOrOutFlag=2 and UserName='" & KSUser.UserName & "' And (" & Param & ") Order By ID"
		   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open SqlStr,Conn,1,3
		   IF RS.Eof And RS.Bof Then
		      Call PayConfirm()	
		   Else
		      RS.Movelast
			  RS(0)=RS(0)+1
			  RS.Update
			  Call GetContent()
		   End IF
		   RS.Close:Set RS=nothing
	   End Sub
	   
	   Sub PayConfirm()
	       If UserLoginTF=false Then Call GetNoLoginInfo():Exit Sub
		   If ReadPoint<=0 Then Call GetContent():Exit Sub
			  If Cint(KSUser.Point)<ReadPoint Then
			     strContent="<br/>对不起，你的可用"&KS.Setting(45)&"不足！阅读本文需要 "&ReadPoint&" "&KS.Setting(46)&KS.Setting(45)&"，你还有"&KSUser.Point&" "&KS.Setting(46)&KS.Setting(45)&"！<br/><br/>"
				 strContent=strContent&"购买"&KS.Setting(45)&"方法<br/>"
				 strContent=strContent&"请用购买到神州行充值卡充值你的"&KS.Setting(45)&"，点击进入<a href=""User/User_CardOnline.asp?"&KS.WapValue&""">神州行充值...</a><br/>" 
			  Else
			     If PayTF="yes" Then
			        IF Cbool(KS.PointInOrOut(ChannelID,RS("ID"),KSUser.UserName,2,ReadPoint,"系统","阅读收费"&KS.C_S(ChannelID,3)&"：<br/>"&RS("Title")))=True Then
					   '支付投稿者提成
					   Dim PayPoint:PayPoint=(ReadPoint*KS.C_C(RS("Tid"),11))/100
					   If PayPoint>0 Then
					      Call KS.PointInOrOut(ChannelID,RS("ID"),RS("Inputer"),1,PayPoint,"系统",KS.C_S(ChannelID,3) & "“" & RS("Title") & "”的提成")
					   End If
					   Call GetContent()
					End If
				 Else
			       strContent="<br/>阅读本文需要消耗 "&ReadPoint&" "&KS.Setting(46)&KS.Setting(45)&"，你目前尚有"&KSUser.Point&""&KS.Setting(46)&KS.Setting(45)&"可用，阅读本文后，您将剩下"&KSUser.Point-ReadPoint&" "&KS.Setting(46)&KS.Setting(45)&"<br/>"
				   strContent=strContent&"你确实愿意花"&ReadPoint&" "&KS.Setting(46)&KS.Setting(45)&"来阅读此文吗？<br/>"
				   strContent=strContent&"<a href=""Show.asp?ChannelID="&ChannelID&"&amp;ID="&ID&"&amp;PayTF=yes&amp;"&KS.WapValue&""">我愿意</a> "
				   strContent=strContent&"<a href=""Show.asp?ChannelID="&ChannelID&"&amp;ID="&ID&"&amp;"&KS.WapValue&""">我不愿意</a><br/>"
			   End If
		   End If
	   End Sub
	   
	   Sub GetNoLoginInfo()
		   strContent="<br/>对不起，你还没有登录，本文至少要求本站的注册会员才可查看!<br/>"
		   strContent=strContent&"如果你还没有注册，请<a href=""User/Reg/?../../Show.asp?ChannelID="&ChannelID&"&amp;ID="&ID&""">点此注册</a>吧!<br/>"
		   strContent=strContent&"如果您已是本站注册会员，赶紧<a href=""User/Login/?../../Show.asp?ChannelID="&ChannelID&"&amp;ID="&ID&""">点此登录</a>吧！<br/>"
	   End Sub
	   Sub GetContent()
	       strContent=true
	   End Sub
	   
	   '**************************************************
	   '函数名：GetPictureByPage
	   '作  用：取出查看图片内容（上一页、下一页方式） 
	   '**************************************************
	   Function GetPictureByPage(ID,ChannelID,PhotoContent)
	       On Error Resume Next
		   Dim CurrPage,PicUrlsArr,TotalPage,Cols,Tpage,PageStr,n,C
		   CurrPage=KS.ChkClng(KS.S("Page"))
		   If CurrPage<=0 Then CurrPage=1
		   PicUrlsArr = Split(PhotoContent, "|||")
		   TotalPage = Cint(UBound(PicUrlsArr) + 1)
		   Cols=KS.ChkClng(KS.S("Cols"))
		   If Cols<=0 Then Cols=2
		   If ((Ubound(PicUrlsArr)+1) Mod cols)=0 Then
		      Tpage=(Ubound(PicUrlsArr)+1)\cols
		   Else
		      Tpage=(Ubound(PicUrlsArr)+1)\cols + 1
		   End If	
		   If TPage<>1 Then
		      If CurrPage=1 Then
			     PageStr = PageStr & "每页显:"
				 If Cols=2 Then
				    PageStr = PageStr & "2 "
				 Else
				    PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols=2&"&KS.WapValue&""" >2</a> "
				 End If
				 If Cols=4 Then
				    PageStr = PageStr & "4 "
				 Else
				    PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols=4&"&KS.WapValue&""" >4</a> "
				 End If
				 If Cols=6 Then
				    PageStr = PageStr & "6"
				 Else
				    PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols=6&"&KS.WapValue&""" >6</a>"
				 End If
				 PageStr = PageStr & KS.C_S(ChannelID,4)&"<br/>"
			  End If
		   End If
		   If TotalPage > 2 Then
		      If KS.BusinessVersion = 1 Then
			     PageStr = PageStr & "【<a href=""Plus/PhotoBroadcast.asp?ID="&ID&"&ChannelID="&ChannelID&"&"&KS.WapValue&""">自动播放</a>】<br/>"
		      End if
		   End if
		   If KS.ChkClng(KS.S("Page"))<=1 Then
		      n=0
		   Else
		      n=cols*(CurrPage-1)
		   End If
		   For c=1 To Cols
		       If n<=Ubound(PicUrlsArr) Then
			      dim url:url=Split(PicUrlsArr(n),"|")(2)
				  if left(url,1)="/" then url=right(url,len(url)-1)
				  if lcase(left(url),4)<>"http" then url=KS.Setting(2) & KS.Setting(3) & url
			      PageStr = PageStr &  "<img src="""&url&""" alt="""" /><br/>"
				  PageStr = PageStr & Split(PicUrlsArr(n),"|")(0) & "<br/><a href="""&url&""">下载</a><br/>"
			   Else
			      PageStr = PageStr & ""
			   End If
			   n=n+1
		   Next
		   Dim startpage,k
		   startpage=1:k=0
		   If TPage<>1 Then
		      If (CurrPage>=10) Then startpage=(CurrPage\10-1)*10+CurrPage Mod 10+2
			  If CurrPage <>tpage Then
			     PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols="&Cols&"&Page="&currpage+1&"&"&KS.WapValue&""">下页</a> "
			  End If
			  PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols="&Cols&"&page="&tpage&"&"&KS.WapValue&""" >末页</a> "
			  PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols="&Cols&"&"&KS.WapValue&""" >首页</a> "
			  If CurrPage<>1 Then
			     PageStr = PageStr & "<a href=""Show.asp?ID="&ID&"&ChannelID="&ChannelID&"&Cols="&Cols&"&Page="&CurrPage-1&"&"&KS.WapValue&""">上页</a> "
			  End If
			  PageStr = PageStr & "<br/>本"&KS.C_S(ChannelID,3)&"共 "&TPage&"/"&CurrPage&"页"
		   End If
		   GetPictureByPage=PageStr
	   End Function
End Class
%>

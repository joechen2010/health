<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer=true
%>
<%Response.ContentType="text/vnd.wap.wml; charset=utf-8" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New DownLoad
KSCls.Kesion()
Set KSCls = Nothing
%>


<%
Class DownLoad
        Private KS,KSUser, KSRFObj
		Private FileContent,RSObj,SqlStr,ShowInfoStr,InfoPurview,ReadPoint,ChargeType,PitchTime,ReadTimes,DownUrl
		Private DomainStr,ID,ChannelID,ClassPurview,UserLoginTF,PayTF,DownUrlTF,TitleStr,Rs,SQL,FoundErr,SoftName,DownID,Hits

		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSUser=New UserCls
		    Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing:Set KSUser=Nothing
		End Sub
		
		Public Sub Kesion()
		    DownUrlTF=false
			DomainStr=KS.GetDomain
		    UserLoginTF=Cbool(KSUser.UserLoginChecked)
			ID = KS.ChkClng(KS.S("ID"))
			ChannelID = KS.ChkClng(KS.S("ChannelID"))
			DownID = KS.ChkClng(KS.S("DownID"))
			PayTF=KS.S("PayTF")
			
			If ID = 0 Then
			   TitleStr="下载错误提示"
			   ShowInfoStr = ShowInfoStr & "错误的系统参数!请输入正确的" & KS.C_S(ChannelID,3) & "ID<br/>"
			   FoundErr=True
			End If
			If DownID = 0 Then
			   TitleStr="下载错误提示"
			   ShowInfoStr = ShowInfoStr & "错误的系统参数!请输入正确的" & KS.C_S(ChannelID,3) & "ID<br/>"
			   FoundErr=True
			End If

			If FoundErr Then Call ShowInfo :Exit Sub
			SqlStr= "Select a.*,ClassPurview From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID
			Set RSObj=Server.CreateObject("Adodb.Recordset")
			RSObj.Open SqlStr,Conn,1,3
			IF RSObj.Eof And RSObj.Bof Then
			   TitleStr="下载错误提示"
			   ShowInfoStr = ShowInfoStr & "找不到你要下载的" & KS.C_S(ChannelID,3) & "！<br/>"
			   FoundErr=True:Call ShowInfo :Exit Sub
			End IF
			ID=RSObj("ID")
			InfoPurview=Cint(RSObj("InfoPurview"))
			ReadPoint=Cint(RSObj("ReadPoint"))
			ChargeType=Cint(RSObj("ChargeType"))
			PitchTime=Cint(RSObj("PitchTime"))
			ReadTimes=Cint(RSObj("ReadTimes"))
			ClassPurview=Cint(RSObj("ClassPurview"))
		    
			If InfoPurview=2 or ReadPoint>0 Then
			   IF UserLoginTF=false Then
				  Call GetNoLoginInfo
			   Else
			      IF KS.FoundInArr(RSObj("ArrGroupID"),KSUser.GroupID,",")=false and readpoint=0 Then
					 ShowInfoStr = ShowInfoStr & "对不起，你没有下载本" & KS.C_S(ChannelID,3) & "的权限!<br/>"
					 FoundErr=True:Call ShowInfo :Exit Sub
				  Else
				     Call PayPointProcess()
				  End If
			   End If
		    ElseIF InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2) Then 
			   If UserLoginTF=false Then
			      Call GetNoLoginInfo
			   Else     
			  	 '============继承栏目收费设置时,读取栏目收费配置===========
			      ReadPoint=Cint(RSObj("DefaultReadPoint"))   
				  ChargeType=Cint(RSObj("DefaultChargeType"))
				  PitchTime=Cint(RSObj("DefaultPitchTime"))
				  ReadTimes=Cint(RSObj("DefaultReadTimes"))
				  '============================================================
				  If ClassPurview=2 Then
					 IF KS.FoundInArr(RSObj("ArrGroupID"),KSUser.GroupID,",")=false Then
					    ShowInfoStr="对不起，你所在的用户组没有下载的权限!<br/>"
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
		    If DownUrlTF=true Then
		       RSObj("Hits") = RSObj("Hits") + 1
			   If DateDiff("D", RSObj("LastHitsTime"), Now()) <= 0 Then
			      RSObj("HitsByDay") = RSObj("HitsByDay") + 1
			   Else
			      RSObj("HitsByDay") = 1
			   End If
			   If DateDiff("ww", RSObj("LastHitsTime"), Now()) <= 0 Then
			      RSObj("HitsByWeek") = RSObj("HitsByWeek") + 1
			   Else
			      RSObj("HitsByWeek") = 1  
			   End If
			   If DateDiff("m", RSObj("LastHitsTime"), Now()) <= 0 Then
			      RSObj("HitsByMonth") = RSObj("HitsByMonth") + 1
			   Else
			      RSObj("HitsByMonth") = 1
			   End If
			   RSObj("LastHitsTime") = Now()
			   RSObj.Update
			   
			   On Error Resume Next
		       Dim DownArr:DownArr=Split(Split(RSObj("DownUrls"),"|||")(DownID-1),"|")
			   If Err Then
			      TitleStr="下载错误提示"
				  ShowInfoStr = ShowInfoStr & "非法访问！<br/>"
				  Call ShowInfo :Exit Sub
			   End If
			   
			   If DownArr(0)="0" Then
			      DownUrl=replace(DownArr(2),"&","&amp;")
				  If lcase(left(DownUrl,4))<>"http" Then DownUrl=KS.Setting(2) & KS.Setting(3) & DownUrl
			      Response.Write "<wml>"
				  Response.Write "<head>"
				  Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>"
				  Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>"
				  Response.Write "</head>"
				  Response.Write "<card id=""main"" title=""" & TitleStr & """ ontimer="""& DownUrl&"""><timer value=""3""/>"
				  Response.Write "<p align=""center"">"
				  Response.Write "请稍候正在下载...<br/>"
				  Response.Write "如果你的手机没能下载，请点击<a href="""&DownUrl&""">这里</a>下载<br/>"
				  Response.Write "<anchor>点击返回<go href=""../Show.asp?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""" method=""post""></go></anchor><br/>"
				  Response.Write "<anchor>返回首页<go href="""&KS.GetGoBackIndex&""" method=""post""></go></anchor><br/>"
				  Response.Write "</p>"
				  Response.Write "</card>"
				  Response.Write "</wml>"
				  Exit Sub
			   Else
			      Set Rs = Server.CreateObject("ADODB.Recordset")
				  SQL = "SELECT top 1 AllDownHits,DayDownHits,HitsTime FROM KS_DownSer WHERE downid="& KS.ChkClng(KS.S("Sid"))
				  Rs.Open SQL,Conn,1,3
				  If Not(Rs.BOF And Rs.EOF) Then
				     hits = CLng(Rs("AllDownHits"))+1
					 Rs("AllDownHits").Value = hits
					 If DateDiff("D", Rs("HitsTime"), Now()) <= 0 Then
					    Rs("DayDownHits").Value = Rs("DayDownHits").Value + 1
					 Else
					    Rs("DayDownHits").Value = 1
						Rs("HitsTime").Value = Now()
					 End If
					 Rs.Update
				  End If
				  Rs.Close:Set Rs = Nothing
				  
				  Dim RS_S:Set RS_S=Server.CreateObject("ADODB.RECORDSET")
				  RS_S.Open "Select IsOuter,DownloadPath,UnionID From KS_DownSer Where DownID=" & KS.ChkClng(KS.S("Sid")),Conn,1,1
				  If Not RS_S.Eof Then
			         Response.Write "<wml>"
					 Response.Write "<head>"
					 Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>"
					 Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>"
					 Response.Write "</head>"
					 Select Case RS_S(0)
					     Case 0
						    Response.Write "<card id=""main"" title=""" & TitleStr & """ ontimer="""&RS_S(1) & DownArr(2)&"""><timer value=""3""/>"
							Response.Write "<p align=""center"">"
							Response.Write "请稍候正在下载...<br/>"
							Response.Write "如果你的手机没能下载，请点击<a href="""&DownArr(2)&""">这里</a>下载<br/>"
						 Case 2
						    Response.Write "<card id=""main"" title=""操作提示"">"
							Response.Write "<p align=""center"">"
							Response.Write "WEB迅雷专用下载地址，请返回选择其它下载地址...<br/>"
						 Case 3
						    Response.Write "<card id=""main"" title=""操作提示"">"
							Response.Write "<p align=""center"">"
							Response.Write "FLASHGET(快车)专用下载地址，请返回选择其它下载地址...<br/>"
					 End Select
					 Response.Write "<anchor>点击返回<go href=""Show.asp?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""" method=""post""></go></anchor><br/>"
					 Response.Write "<anchor>返回首页<go href="""&KS.GetGoBackIndex&""" method=""post""></go></anchor><br/>"
					 Response.Write "</p>"
					 Response.Write "</card>"
					 Response.Write "</wml>"
					 Exit Sub
				  End If
				  RS_S.Close:Set RS_S=Nothing
			   End If
		   Else
		     TitleStr="操作提示"
		   End If
		   Call ShowInfo()
		   RSObj.Close:Set RSObj=Nothing
	   End Sub


      '收费扣点处理过程
	   Sub PayPointProcess()
	       If Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2)) Then
		   IF UserLoginTF=false Then Call GetNoLoginInfo :Exit Sub
		   Dim UserChargeType:UserChargeType=KSUser.ChargeType
		   If UserChargeType=1 Then
		      Select Case ChargeType
			      Case 0:Call CheckPayTF("1=1")
				  Case 1
				     If DataBaseType=1 Then
					    Call CheckPayTF("datediff(hour,AddDate," & SqlNowString & ")<" & PitchTime)
					 Else
					    Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime)
					 End If
				  Case 2:Call CheckPayTF("Times<" & ReadTimes)
				  Case 3
				     If DataBaseType=1 Then
					    Call CheckPayTF("datediff(hour,AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
					 Else
					    Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
					 End If
				  Case 4
				     If DataBaseType=1 Then
					    Call CheckPayTF("datediff(hour,AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
					 Else
					    Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
					 End If
				  Case 5:Call PayConfirm()
			  End Select
		   Elseif UserChargeType=2 Then
		      If KSUser.GetEdays <=0 Then
			     ShowInfoStr="对不起，你的账户已过期" & KSUser.GetEdays & "天,此" & KS.C_S(ChannelID,3) & "需要在有效期内才可以下载，请及时与我们联系！<br/>"
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
		   RS.Open SqlStr,conn,1,3
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
		   If ReadPoint=0 Then GetContent():Exit Sub
		   If Cint(KSUser.Point)<ReadPoint Then
		      ShowInfoStr="对不起，你的可用" & KS.Setting(45) & "不足!下载本" & KS.C_S(ChannelID,3) & "需要" & ReadPoint & "" & KS.Setting(46) & KS.Setting(45) &",你还有" & KSUser.Point & "" & KS.Setting(46) & KS.Setting(45) & ",请及时与我们联系！<br/>" 
		   Else
		      If PayTF="yes" Then
		         IF Cbool(KS.PointInOrOut(ChannelID,RSObj("ID"),KSUser.UserName,2,ReadPoint,"系统","下载收费" & KS.C_S(ChannelID,3) & "：<br>" & RSObj("Title")))=True Then
				    '支付投稿者提成
					Dim PayPoint:PayPoint=(ReadPoint*KS.C_C(RSObj("Tid"),11))/100
					If PayPoint>0 Then
					   Call KS.PointInOrOut(ChannelID,RSObj("ID"),RSObj("Inputer"),1,PayPoint,"系统",KS.C_S(ChannelID,3) & "“" & RSObj("Title") & "”的提成")
					End If
				    Call GetContent()
				 End If
		      Else
		         ShowInfoStr="下载本软件需要消耗" & ReadPoint & "" & KS.Setting(46) & KS.Setting(45) &",你目前尚有" & KSUser.Point & "" & KS.Setting(46) & KS.Setting(45) &"可用,下载本" & KS.C_S(ChannelID,3) & "后，您将剩下" & KSUser.Point-ReadPoint & "" & KS.Setting(46) & KS.Setting(45) &"<br/>你确实愿意花" & ReadPoint & "" & KS.Setting(46) & KS.Setting(45) & "来下载本" & KS.C_S(ChannelID,3) & "吗?<br/><a href=""?ID=" & ID & "&PayTF=yes&DownID=" & DownID & """>我愿意</a> <a href=""" &DomainStr & """>我不愿意</a><br/>"
			   End If
			End If
	   End Sub
	   Sub GetNoLoginInfo()
		   ShowInfoStr="对不起，你还没有登录，本" & KS.C_S(ChannelID,3) & "至少要求本站的注册会员才可下载!<br/>如果你还没有注册，请<a href=""" & DomainStr & "User/Reg/"">点此注册</a>吧!<br/>如果您已是本站注册会员，赶紧<a href=""" & domainstr & "User/Login/"">点此登录</a>吧！<br/>"
	   End Sub
	   Sub GetContent()
		   TitleStr=RSObj("Title")
		   DownUrlTF=True
	   End Sub
			
	   Function ShowInfo()
		   Response.Write "<wml>" &vbcrlf
		   Response.Write "<head>" &vbcrlf
		   Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
		   Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
		   Response.Write "</head>" &vbcrlf
		   Response.Write "<card id=""main"" title=""" & TitleStr & """>" &vbcrlf
		   Response.Write "<p align=""center"">" &vbcrlf
		   Response.Write ""&ShowInfoStr&"" &vbcrlf
		   Response.Write "<anchor>点击返回<go href=""Show.asp?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""" method=""post""></go></anchor><br/>" &vbcrlf
		   Response.Write "<anchor>返回首页<go href="""&KS.GetGoBackIndex&""" method=""post""></go></anchor><br/>" &vbcrlf
		   Response.Write "</p>" &vbcrlf
		   Response.Write "</card>" &vbcrlf
		   Response.Write "</wml>"
	   End Function
End Class
%>
 

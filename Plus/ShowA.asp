<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Advertise
KSCls.Kesion()
Set KSCls = Nothing

Class Advertise
        Private KS
		Private getplace,getshow,adsrs,adssql,adsrsp,adssqlp,adsrss,adssqls,getip,getggwlxsz,getggwhei,getggwwid
        Private ttarg,DomainStr,GaoAndKuan,advertvirtualvalue
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		  Select Case KS.S("Action")
		   Case "Daima"
		     Call AdvertiseDaima()
		   Case "Adurl"
		     Call AdvertiseAdurl()
		   Case "AdOpen"
		     Call AdvertiseAdOpen()
		   Case Else
		    Call AdvertiseMain()
		  End Select
		End Sub
		Sub AdvertiseMain()
		DomainStr=KS.GetDomain
		getplace=KS.ChkClng(KS.S("i"))
		
		dim GaoAndKuan
		Dim adsrs1:Set adsrs1=server.createobject("adodb.recordset")
		adsrs1.open "select * From KS_ADPlace where show_flag=1 and place="&getplace,Conn,1,1
		if not adsrs1.eof then
		getggwlxsz=adsrs1(2)
		else
		getggwlxsz=0
		end if
		getggwhei=adsrs1(3)
		getggwwid=adsrs1(4)
		
		GaoAndKuan=""
		
		if getggwhei<>"" then GaoAndKuan=" height="&getggwhei&" "
		if getggwwid<>"" then GaoAndKuan=GaoAndKuan&" width="&getggwwid&" "
		
		adsrs1.close:Set adsrs1=nothing
		
		''''''''''''''''''''''''''''''''每次显示广告位前，检测其中的各广告条是否过期，并更新状态''''''''''''''''''''''''''''''''
		set adsrsp=server.createobject("adodb.recordset")
		adssqlp="Select * from KS_Advertise where act=1 and class <> 0 and  place="&getplace&" order by time"
		adsrsp.open adssqlp,Conn,1,3
		
		while not adsrsp.eof
		
		advertvirtualvalue=0
		
		if adsrsp("class")=1 then
		if adsrsp("click")>=adsrsp("clicks") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=2 then
		if adsrsp("show")>=adsrsp("shows") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=3 then
		if now()>=adsrsp("lasttime") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=4 then
		if adsrsp("click")>=adsrsp("clicks") then
		advertvirtualvalue=1
		end if
		if adsrsp("show")>=adsrsp("shows") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=5 then
		if adsrsp("click")>=adsrsp("clicks") then
		advertvirtualvalue=1
		end if
		if now()>=adsrsp("lasttime") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=6 then
		if adsrsp("show")>=adsrsp("shows") then
		advertvirtualvalue=1
		end if
		if now()>=adsrsp("lasttime") then
		advertvirtualvalue=1
		end if
		
		elseif adsrsp("class")=7 then
		if adsrsp("click")>=adsrsp("clicks") then
		advertvirtualvalue=1
		end if
		if adsrsp("show")>=adsrsp("shows") then
		advertvirtualvalue=1
		end if
		if now()>=adsrsp("lasttime") then
		advertvirtualvalue=1
		end if
		end if
		
		if advertvirtualvalue>=1 then
		adsrsp("act")=2
		adsrsp.update
		end if
		adsrsp.movenext
		wend
		adsrsp.close:set adsrsp=nothing 
		'''''''''''''''''''''''''''''''''''''''''''''''结束 检测、更新''''''''''''''''''''''''''''''''
		set adsrs=server.createobject("adodb.recordset")
		set adsrs1=server.createobject("adodb.recordset")
		adsrs1.open "select * From KS_ADPlace where place="&getplace,Conn,1,1
        ''''''''''''''''''''''''''''''''''''''''根据显示方式的不同进行显示''''''''''''''''''''''''
Select Case getggwlxsz

       Case 1 
       adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3
       Call DggtXs()
       adsrs.close

       Case 2 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3
       do while not adsrs.eof 
       Call DggtXs()
       adsrs.movenext
       Response.Write "document.write('<br>');"
       loop
       adsrs.close
       
       Case 3 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3
       do while not adsrs.eof 
       Call DggtXs()
       adsrs.movenext
       Response.Write "document.write('&nbsp;&nbsp;');"
       loop
       adsrs.close

       Case 4 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3
       Response.Write "document.write('<marquee  direction=up"&GaoAndKuan&">');"
       do while not adsrs.eof
       Call DggtXs()
       adsrs.movenext
       Response.Write "document.write('<br><br>'); "
       loop
       Response.Write "document.write('</marquee>');"
       adsrs.close 

       Case 5 
       
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3

       
       Response.Write "document.write('<marquee"&GaoAndKuan&">');"
       do while not adsrs.eof
       Call DggtXs()
       adsrs.movenext
       Response.Write "document.write('&nbsp;&nbsp;');"
       loop
       Response.Write "document.write('</marquee>');"
       adsrs.close 

       Case 6 
       adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3
       do while not adsrs.eof
       call gaokuan()
       Response.Write "window.open('"&DomainStr&"plus/ShowA.asp?Action=AdOpen&i="&adsrs("id")&"','" & KS.Setting(0) & "广告服务"&adsrs("id")&"','"&GaoAndKuan&"');"
       adsrs.movenext
       loop
       adsrs.close 

       Case 7 
       adssql="Select top 1 id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where act=1 and place="&getplace&" order by time"
       adsrs.open adssql,Conn,1,3  
       call gaokuan()
       Response.Write "window.open('"&DomainStr&"plus/ShowA.asp?Action=AdOpen&i="&adsrs("id")&"','" & KS.Setting(0) & "广告服务','"&GaoAndKuan&"');"
       adsrs.close 
       
   End Select 
		set adsrs=nothing
		Conn.close:set Conn=nothing 
		End Sub
	
	 ''''''''''''''''''''''''''''显示单个广告条 '''''''''''''''''''''''''''''''''''''''''''''' 
		
		Sub DggtXs() 
		adsrs("show")=adsrs("show")+1
		adsrs("time")=now()
		adsrs.Update
		if adsrs("window")=0 then
		ttarg = "_blank"
		else 
		ttarg="_top" 
		end if
		
		if isnumeric(adsrs("hei")) then
		GaoAndKuan=" height="&adsrs("hei")&" "
		else
		
		if right(adsrs("hei"),1)="%" then
		if isnumeric(Left(len(adsrs("hei"))-1))=true then
		 GaoAndKuan=" height="&adsrs("hei")&" "
		end if
		end if
		
		end if
		
		
		if isnumeric(adsrs("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
		else
		if right(adsrs("wid"),1)="%" then
		if isnumeric(Left(len(adsrs("wid"))-1))=true then 
		GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
		end if
		end if
		end if
		
		Select Case adsrs("xslei")
		   Case "txt"%>document.write('<a title=\"<%=adsrs("sitename")%>\"  href=\"<%=DomainStr%>plus/ShowA.asp?Action=Adurl&id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><%=Replace(Replace(Replace(Replace(adsrs("intro"), Chr(13)& Chr(10), ""),"'","\'"),"""","\"""),vbcrlf,"") %></a>');
		<% Case "gif"%>document.write('<a href=\"<%=DomainStr%>plus/ShowA.asp?Action=Adurl&id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><img  alt=\"<%=adsrs("sitename")%>\"  border=0 <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>');
		<% Case "swf"%>document.write('<EMBED src=<%=adsrs("gif_url")%>   <%=GaoAndKuan%>  quality=high TYPE=\"application/x-shockwave-flash\"></EMBED>');
		<% Case "dai"%>document.write('<iframe marginwidth=0 marginheight=0  frameborder=0 bordercolor=000000 scrolling=no  name=\"广告\" src=\"<%=DomainStr%>plus/ShowA.asp?Action=Daima&id=<%=adsrs("id")%>\"  <%=GaoAndKuan%> ></iframe>');
		<% Case else%>document.write('<a href=\"<%=DomainStr%>plus/ShowA.asp?Action=Adurl&id=<%=adsrs("id")%>\" target=\"<%=ttarg%>\"><img alt=\"<%=adsrs("sitename")%>\"  border=0 <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>');
		<%End Select
		'暂且关闭记录IP功能
		'getip=request.ServerVariables("REMOTE_ADDR")
		'set adsrss=server.createobject("adodb.recordset")
		'adssqls="select * from KS_Adiplist"
		'adsrss.open adssqls,Conn,1,3
		'adsrss.AddNew
		'adsrss("adid") =adsrs("id")
		'adsrss("time") = now()
		'adsrss("ip") = getip
		'adsrss("class") = 1
		'adsrss.update
		'adsrss.close
		'set adsrss=nothing
		
		
		End Sub
		
		Sub gaokuan() 
		adsrs("show")=adsrs("show")+1
		adsrs("time")=now()
		adsrs.Update
		if adsrs("window")=0 then
		ttarg = "_blank"
		else 
		ttarg="_top" 
		end if
		
		if adsrs("hei")<>"" then
		
		if isnumeric(adsrs("hei")) then
		GaoAndKuan=" height="&adsrs("hei")&" "
		else
		
		 if right(adsrs("hei"),1)="%" then
		   if isnumeric(Left(len(adsrs("hei"))-1))=true then
			 GaoAndKuan=" height="&adsrs("hei")&" "
		   end if
		 end if
		
		end if
		
		
		if isnumeric(adsrs("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
		else
		if right(adsrs("wid"),1)="%" then
		if isnumeric(Left(len(adsrs("wid"))-1))=true then 
		GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
		end if
		end if
		end if
		else 
		end if
	End Sub
	function UBBCode(strContent)
	on error resume next
	strContent = KS.HTMLEncode(strContent)
	dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =true
	objRegExp.Global=True

   
	objRegExp.Pattern="(\[color=(.*)\])(.*)(\[\/color\])"
	strContent=objRegExp.Replace(strContent,"<font color=$2>$3</font>")
	objRegExp.Pattern="(\[face=(.*)\])(.*)(\[\/face\])"
	strContent=objRegExp.Replace(strContent,"<font face=$2>$3</font>")
	objRegExp.Pattern="(\[align=(.*)\])(.*)(\[\/align\])"
	strContent=objRegExp.Replace(strContent,"<div align=$2>$3</div>")

	objRegExp.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=1 face=""Verdana, Arial"">quote:</font><HR>$2<HR></BLOCKQUOTE>")

    
	objRegExp.Pattern="(\[i\])(.*)(\[\/i\])"
	strContent=objRegExp.Replace(strContent,"<i>$2</i>")
	objRegExp.Pattern="(\[u\])(.*)(\[\/u\])"
	strContent=objRegExp.Replace(strContent,"<u>$2</u>")
	objRegExp.Pattern="(\[b\])(.*)(\[\/b\])"
	strContent=objRegExp.Replace(strContent,"<b>$2</b>")


	objRegExp.Pattern="(\[size=1\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=1>$2</font>")
	objRegExp.Pattern="(\[size=2\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=2>$2</font>")
	objRegExp.Pattern="(\[size=3\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=3>$2</font>")
	objRegExp.Pattern="(\[size=4\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=4>$2</font>")

	strContent = doCode(strContent, "[list]", "[/list]", "<ul>", "</ul>")
	strContent = doCode(strContent, "[list=1]", "[/list]", "<ol type=1>", "</ol id=1>")
	strContent = doCode(strContent, "[list=a]", "[/list]", "<ol type=a>", "</ol id=a>")
	strContent = doCode(strContent, "[*]", "[/*]", "<li>", "</li>")
	strContent = doCode(strContent, "[code]", "[/code]", "<pre id=code><font size=1 face=""Verdana, Arial"" id=code>", "</font id=code></pre id=code>")

	set objRegExp=Nothing
	UBBCode=strContent
	end function
 
 '代码
  Sub AdvertiseDaima()
         response.write "<body>"
  	    if KS.S("id")<>"" and isnumeric(KS.S("ID")) then
			dim adssql
			dim adsrs:set adsrs=server.createobject("adodb.recordset")
			adssql="Select top 1 intro from KS_Advertise where id="&KS.ChkClng(KS.S("id"))&" order by time"
			adsrs.open adssql,conn,1,1       
			if not adsrs.eof then
			response.write adsrs(0)
			end if
			adsrs.close:set adsrs=nothing
			conn.close:set conn=nothing
		else
			response.write "<center><br><br>无效广告。</center>"
		end if
		response.write "</body>"
  End Sub

 Sub AdvertiseAdurl()
 		dim Url,getid,getclick,geturl,adssql,RSObj,SqlStr,getip,sitename
		getid=KS.ChkClng(KS.S("id"))
		set RSObj=server.createobject("adodb.recordset")
		adssql="Select id,url,click,sitename from KS_Advertise where id="&getid
		RSObj.open adssql,Conn,1,3
		getclick=RSObj(2)+1
		sitename=RSOBJ(3)
		RSObj(2)=getclick
		RSObj.Update
		Url=RSObj(1)
		RSObj.Close
		'暂且关闭记录IP功能
		'getip=request.ServerVariables("REMOTE_ADDR")
		'SqlStr="select * from KS_Adiplist"
		'RSObj.open SqlStr,Conn,1,3
		'RSObj.AddNew
		'RSObj("adid") =getid
		'RSObj("time") = now()
		'RSObj("ip") = getip
		'RSObj("class") = 2
		'RSObj.update
		'RSObj.close
		'set RSObj=nothing 
		
		'========点广告加积分==================
		 if KS.Setting(166)="1" And KS.ChkClng(KS.Setting(167))>0 Then
		   If KS.C("UserName")<>"" Then
			  If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & KS.C("UserName") & "' and year(adddate)=year(" & SQLNowString  &") and month(adddate)=month(" & SQLNowString &") and day(adddate)=day(" & SQLNowString & ") and channelid=1000 and infoid=" & getid).Eof Then
			  	 '判断有没有到达每天增加的总限
				 Dim TodayScore:TodayScore=0
				 If KS.ChkClng(KS.Setting(165))<>0 Then
				  TodayScore=KS.ChkClng(Conn.Execute("select sum(Score) from ks_logscore where InOrOutFlag=1 and year(adddate)=year(" & SQLNowString & ") and month(adddate)=month(" & SQLNowString & ") and day(adddate)=day(" & SQLNowString & ") and username='" & ks.c("UserName") & "'")(0))
				 End If
                 If TodayScore+KS.ChkClng(KS.Setting(167))<KS.ChkClng(KS.Setting(165)) Then

                      Conn.Execute("Update KS_User Set Score=Score+" & KS.ChkClng(KS.Setting(167)) & " Where UserName='" & KS.C("UserName") & "'")
					  'on error resume next
					  Dim CurrScore:CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & KS.C("UserName") & "'")(0)
					  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,Channelid,InfoID) values('" & KS.C("UserName") & "',1," & KS.ChkClng(KS.Setting(167)) & ","&CurrScore & ",'系统','点击广告[" & sitename & "(" & url & ")]所得!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "',1000," & getid & ")")

				   
				 End If

			  End If
			  
		   End If
		 End If
		'=====================================
		
		
		
		geturl=Url
		set Conn=nothing
		Response.Redirect geturl
 End Sub
 Sub AdvertiseAdOpen()
 %>
     <html>
	 <head></head>
	 <body topmargin="0" leftmargin="0">
	<%
	Dim DomainStr:DomainStr=KS.GetDomain
	Dim ttarg:ttarg="_top"
	Dim GaoAndKuan:GaoAndKuan=""
	Dim Adsrs:Set adsrs=server.createobject("adodb.recordset")
	Dim adssql:adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei from KS_Advertise where id="&KS.Chkclng(KS.S("i"))
	adsrs.open adssql,Conn,3,3
	adsrs("show")=adsrs("show")+1
	adsrs("time")=now()
	adsrs.Update
	if adsrs("window")=0 then
	ttarg = "_blank"
	end if
	
	if isnumeric(adsrs("hei")) then
	GaoAndKuan=" height="&adsrs("hei")&" "
	else
	
	if right(adsrs("hei"),1)="%" then
	if isnumeric(Left(len(adsrs("hei"))-1))=true then
	 GaoAndKuan=" height="&adsrs("hei")&" "
	end if
	end if
	
	end if
	
	
	if isnumeric(adsrs("wid")) then
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	else
	if right(adsrs("wid"),1)="%" then
	if isnumeric(Left(len(adsrs("wid"))-1))=true then 
	GaoAndKuan=GaoAndKuan&" width="&adsrs("wid")&" "
	end if
	end if
	end if
	
	
			   Select Case adsrs("xslei")
		
				Case "txt"%><a title="<%=adsrs("intro")%>" href="?Action=AdUrl&id=<%=adsrs("id")%>" target="<%=ttarg%>"><font color=red><%=adsrs("sitename")%></font></a>
	<%          Case "gif"%><a title="<%=adsrs("intro")%>" href="?Action=AdUrl&id=<%=adsrs("id")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a> 
	<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=adsrs("gif_url")%>"><param name=quality value=high>
	  <%          Case "dai"%><%=adsrs("intro")%>
	  <embed src="<%=adsrs("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed></object>
	<%          Case else%><a title="<%=adsrs("intro")%>" href="?Action=AdUrl&id=<%=adsrs("id")%>" target="<%=ttarg%>"><img border=0  <%=GaoAndKuan%> src="<%=adsrs("gif_url")%>"></a>
	<%
			   End Select%><%
	adsrs.close
	set adsrs=nothing
	Conn.close
	set Conn=nothing 
	%>
	 </body>
	</html>
<%
 End Sub
End Class
 %>  

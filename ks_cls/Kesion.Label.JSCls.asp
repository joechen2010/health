<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim JSCls:Set JSCls=New JSCommonCls
Class JSCommonCls
        Dim KS,Temps,DomainStr
		Private Sub Class_Initialize()
		Set KS=New PublicCls
		DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set JSCls=Nothing
		 Set KS=Nothing
		End Sub
		'====================================================�滻ͨ��JS=============================================
		'˵�����������������չ���ѳ��õ�JS������з�װ���Ϳ������ñ�ǩ����
		'============================================================================================================
		Sub echo(str)
		  temps=temps & str
		End Sub
		Sub echoln(str)
		  temps=temps & str & vbcrlf
		End Sub
		Sub Run(sTemp,ByRef Templates)
		 dim RCls: set RCls=New Refresh
		  temps=Templates
		 select case Lcase(sTemp)
		   case "js_time1" : echo  "<script src=""" & DomainStr & "ks_inc/time/1.js"" type=""text/javascript""></script>"
		   case "js_time2" : echo  "<script src=""" & DomainStr & "ks_inc/time/2.js"" type=""text/javascript""></script>"
		   case "js_time3" : echo  "<script src=""" & DomainStr & "ks_inc/time/3.js"" type=""text/javascript""></script>"
		   case "js_time4" : echo  "<div id=""kstime""></div><script>setInterval(""kstime.innerHTML=new Date().toLocaleString()+' ����'+'��һ����������'.charAt (new Date().getDay());"",1000);</script>"
		   case "js_language" : echo "<script src=""" & DomainStr & "KS_Inc/language.js"" type=""text/javascript""></script>"
		   case "js_collection": echo "<a href=""#"" onclick=""javascript:window.external.addFavorite('http://'+location.hostname+(location.port!=''?':':'')+location.port,'" & KS.Setting(0) &"');"">�����ղ�</a>"
		   case "js_homepage" : echo "<a onclick=""this.style.behavior='url(#default#homepage)';this.setHomePage('http://'+location.hostname+(location.port!=''?':':'')+location.port);"" href=""#"">��Ϊ��ҳ</a>"
		   case "js_contactwebmaster" : echo "<a href=""mailto:" & KS.Setting(11) & """>��ϵվ��</a>"
		   case "js_nosave" : echo "<NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT>"
		   case "js_goback" : echo "<a href=""javascript:history.back(-1)"">������һҳ</a>"
		   case "js_windowclose" : echo "<a href=""javascript:window.close();"">�رմ���</a>"
		   case "js_noiframe": echo "<script type=""text/javascript"">if(self!=top){top.location=self.location;}</script>"
		   case "js_nocopy" : echoln "<script type=""text/javascript"">" 
		                      echoln "document.oncontextmenu=new Function(""event.returnValue=false;"");"  
							  echoln "document.onselectstart=new Function(""event.returnValue=false;"");"
							  echoln "</script>"
		   case "js_dcroll" : echoln "<script type=""text/javascript"">"
		                      echoln "var currentpos,timer; " 
							  echoln "function initialize(){ timer=setInterval(""scrollwindow()"",30);} " 
							  echoln "function sc(){clearInterval(timer);}" 
							  echoln "function scrollwindow(){ "
							  echoln "if (document.documentElement && document.documentElement.scrollTop){"
							  echoln " currentpos=document.documentElement.scrollTop;window.scroll(0,++currentpos); "
							  echoln " if (currentpos != document.documentElement.scrollTop) sc();}"
							  echoln "else if (document.body){"
							  echoln "	currentpos=document.body.scrollTop; window.scroll(0,++currentpos);"
							  echoln "if (currentpos != document.body.scrollTop) sc(); }"
							  echoln "} "
							  echoln "document.onmousedown=sc"
							  echoln "document.ondblclick=initialize"
							  echoln "</script>"
		 end select
		  Templates=temps
		End Sub
		
		Sub Equal(stemp,Param,ByRef Templates)
		  dim RCls: set RCls=New Refresh
		  temps=Templates
		  select case Lcase(stemp)
		    case "js_ad" '�������
			  echo "<script>var delta=" & Param(3) & ";var closeSrc='" & Param(2) & "';var rightSrc='" & Param(1) & "';var leftSrc='" & Param(0) & "';</script><script src=""" & DomainStr & "ks_inc/ad/1.js"" type=""text/javascript""></script>" 
		    case "js_status1" '״̬��Ŀ����Ч��
			  echo "<script type=""text/javascript"">var msg = '" & Param(0) & "' ;var interval = " & Param(1) & ";</script><script src=""" & DomainStr & "ks_inc/status/1.js"" type=""text/javascript""></script>"
			case "js_status2" '������״̬���ϴ�������ѭ����ʾ
			  echo "<script>var speed = " & Param(1) &";var m1 = '" & Param(0) & "' ;</script><script src=""" & DomainStr & "ks_inc/status/2.js"" type=""text/javascript""></script>"
			case "js_status3" '������״̬���ϴ���֮���ƶ���ʧ
			  echo "<script>var speed = " & Param(1) &";var Message = '" & Param(0) & "' ;</script><script src=""" & DomainStr & "ks_inc/status/3.js"" type=""text/javascript""></script>"
		  end select
		  Templates=temps
		End Sub
		
		
        '====================================================�滻ͨ��JS����=============================================

End Class
%> 

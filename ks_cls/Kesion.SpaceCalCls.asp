<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class CalendarCls
		Private Sub Class_Initialize()
		End Sub
        Private Sub Class_Terminate()
		End Sub
		
		Sub calendar(byref DateStr,username)
		dim ca_bg_color,ca_head_color,ca_week_color,ca_day_color,ca_nowday_color
		dim thisbgcolor,ca_today_color,ca_headtext_color
		dim c_year,c_month,c_day,logdate,today,tomonth,toyear,sql,s,count,b,c
		dim thismonth,thisdate,thisyear,startspace,nextmonth,nextyear,promonth,proyear,linktrue
		dim rs,selectdate,blogmonth,linkcount
		redim linkdays(2,0)
		'=============================��������===================
		ca_bg_color=""  '����������������ɫ
		ca_head_color=""  'ͷ����ɫ
		ca_week_color=""  '��������ɫ
		ca_day_color=""  '������ɫ
		ca_nowday_color="#ff0000"  '���ղ鿴ʱ���鿴�յ���ɫ
		ca_today_color=""  '�������ɫ
		ca_headtext_color="" '�ײ��������ֵ���ɫ
		
		selectdate=Request("date")
		if selectdate="" or not isdate(selectdate) then
			c_year=year(now())
			c_month=month(now())
			c_day=0
		else
			c_year=year(selectdate)
			c_month=month(selectdate)
			c_day=day(selectdate)
		end if
		
		blogmonth=trim(request.Form("month"))
		if blogmonth<>"" then
		   If IsDate(blogmonth) then
			c_year=year(blogmonth)
			c_month=month(blogmonth)
			c_day=day(blogmonth)
		   end if
		end if
		
		logdate=c_year&"-"&c_month
		c_year=cint(c_year)
		c_month=cint(c_month)
		c_day=cint(c_day)
		
		'===============================================�������====================
		if DataBaseType=1 then
			sql="SELECT adddate FROM ks_bloginfo WHERE datediff(month,"&logdate&",adddate)>0 and username='"&username&"'"
		else
			sql="SELECT adddate FROM ks_bloginfo WHERE datediff('n','"&logdate&"',adddate)>0 and username='"&username&"'"
		end if
			set rs=conn.EXECUTE(sql)
		
		dim theday
		theday=0
		
		do while not rs.eof
			if day(rs("adddate"))<>theday then
				theday=day(rs("adddate"))
				redim preserve linkdays(2,linkcount)
				linkdays(0,linkcount)=month(rs("adddate"))
				linkdays(1,linkcount)=day(rs("adddate"))
				linkdays(2,linkcount)="$('#date').val('"&logdate&"-"&theday & "');$('#calqform').submit();"
				linkcount=linkcount+1
			end if
			rs.MoveNext
		Loop
		set rs=nothing
		'=========================================================================
		
		dim mname(12) 
		mname(0)=""
		mname(1)="January "
		mname(2)="February "
		mname(3)="Mar."
		mname(4)="April "
		mname(5)="may "
		mname(6)="June "
		mname(7)="July "
		mname(8)="August "
		mname(9)="September "
		mname(10)="October "
		mname(11)="November "
		mname(12)="December "
		
		dim mdays(12)
		mdays(0)=""
		mdays(1)=31
		mdays(2)=28
		mdays(3)=31
		mdays(4)=30
		mdays(5)=31
		mdays(6)=30
		mdays(7)=31
		mdays(8)=31
		mdays(9)=30
		mdays(10)=31
		mdays(11)=30
		mdays(12)=31
		
		
		'�����������
		today=day(now()) 
		tomonth=month(now())
		toyear=year(now())
		
		'ָ���������ռ�����
		
		thismonth=c_month
		thisdate=c_day
		thisyear=c_year
		If IsDate("February 29, " & thisyear) Then mdays(2)=29
		
		'ȷ������1�ŵ�����
		startspace=weekday( thismonth&"-1-"&thisyear )-1
		
		nextmonth=c_month+1
		nextyear=c_year
		if nextmonth>12 then 
		nextmonth=1
		nextyear=nextyear+1
		end if
		promonth=c_month-1
		proyear=c_year
		if promonth<1 then 
		promonth=12
		proyear=proyear-1
		end if
		
		DateStr="<table border='0' width='105%' align='center' cellspacing='1' cellpadding='1' style='background: url(images/month/" & thismonth & ".gif);background-position: center; background-repeat: no-repeat;' bgcolor='"&ca_bg_color&"'>"
		
		DateStr=DateStr&"<div style='display:none'><form id='calqform' action='../space/?" & username & "/blog' method='post'><input type='text' name='date' id='date'></form></div>"
		DateStr=DateStr&"<div style='display:none'><form id='calform' action='#' method='post'><input type='text' name='month' id='month'></form></div>"
		
		DateStr=DateStr&"<tr><td colspan='1'  bgcolor='"&ca_head_color&"' style='font-size:16px; font-family:;text-align :right'><a href='javascript:void(0)' onclick=""$('#month').val('"&proyear&"-"&promonth&"');$('#calform').submit();"">&laquo;</a></td><td colspan='5' style='color:"&ca_headtext_color&";font-size:14px;font-family:;text-align :center'><b>"&mname(thismonth)& thisyear&"</b></td><td colspan='1' bgcolor='"&ca_head_color&"' style='font-size:16px; font-family:;text-align :left';><a href='javascript:void(0)' onclick=""$('#month').val('"&nextyear&"-"&nextmonth&"');$('#calform').submit()"">&raquo;</a></td></tr><tr>"
		
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>һ</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td>"
		DateStr=DateStr&"<td align='center' bgcolor='"&ca_week_color&"'>��</td></tr><tr>"
		
		for s=0 to startspace-1
			DateStr=DateStr&"<td bgcolor='"&ca_day_color&"'></td>"
		next
		
		count=1
		while count<=mdays(thismonth)
			 for b=startspace to 6
				 thisbgcolor=ca_day_color
				 if count=today and thisyear=toyear and thismonth=tomonth then thisbgcolor=ca_today_color
				 if count=thisdate then thisbgcolor=ca_nowday_color
				 DateStr=DateStr&"<td align='center' bgcolor='"&thisbgcolor&"' style='font-size:10px;font-family:'>"
				 linktrue="false"
				 for c=0 to ubound(linkdays,2)
					 if linkdays(0,c)<>"" then
						if linkdays(0,c)=thismonth and linkdays(1,c)=count then
						   
						   DateStr=DateStr&"<a href='javascript:void(0)' onclick="""&linkdays(2,c)&""">"
						   linktrue="true"
						end if
					 end if
				 next
				 if count<=mdays(thismonth) then DateStr=DateStr&count
				 if linktrue="true" then DateStr=DateStr&"</a>"
				 DateStr=DateStr&"</td>"
		
				 count=count+1
			 next
			 DateStr=DateStr&"</tr>"
			 startspace=0
		wend
		
		DateStr=DateStr&"<tr><td colspan='7' bgcolor='"&ca_week_color&"' align='center'>"
		DateStr=DateStr&"</tr></table>"
		
		end sub
End Class
%>

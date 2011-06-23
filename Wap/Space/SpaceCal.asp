<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="高级日历">
<p>
<%
UserName=Request("UserName")
Redim LinkDays(2,0)
'=============================日历代码===================
SelectDate=Request("date")
If SelectDate="" Or Not Isdate(SelectDate) Then
   C_Year=Year(Now())
   C_Month=Month(Now())
   C_Day=0
Else
   C_Year=Year(SelectDate)
   C_Month=Month(SelectDate)
   C_Day=Day(SelectDate)
End If
		
BlogMonth=Trim(Request.QueryString("Month"))
If BlogMonth<>"" Then
   If IsDate(BlogMonth) then
      C_Year=Year(BlogMonth)
	  C_Month=Month(BlogMonth)
	  C_Day=Day(BlogMonth)
   End If
End If
		
LogDate=C_Year&"-"&C_Month
C_Year=Cint(C_Year)
C_Month=Cint(C_Month)
C_Day=Cint(C_Day)
		
'===============================================添加连接====================
If DataBaseType=1 Then
   sql="SELECT adddate FROM ks_bloginfo WHERE datediff(Month,"&LogDate&",adddate)>0 and UserName='"&UserName&"'"
Else
   sql="SELECT adddate FROM ks_bloginfo WHERE datediff('n','"&LogDate&"',adddate)>0 and UserName='"&UserName&"'"
End If
set RS=Conn.EXECUTE(sql)
dim TheDay
TheDay=0
do while not RS.eof
   If Day(RS("adddate"))<>TheDay Then
      TheDay=Day(RS("adddate"))
	  redim preserve LinkDays(2,LinkCount)
	  LinkDays(0,LinkCount)=Month(RS("adddate"))
	  LinkDays(1,LinkCount)=Day(RS("adddate"))
	  LinkDays(2,LinkCount)="blog.asp?UserName=" & UserName & "&amp;date=" & LogDate & "-" & TheDay & "&amp;" & KS.WapValue & ""
	  LinkCount=LinkCount+1
   End If
   RS.MoveNext
Loop
set RS=nothing
'=========================================================================
dim MDays(12)
MDays(0)=""
MDays(1)=31
MDays(2)=28
MDays(3)=31
MDays(4)=30
MDays(5)=31
MDays(6)=30
MDays(7)=31
MDays(8)=31
MDays(9)=30
MDays(10)=31
MDays(11)=30
MDays(12)=31
		
'今天的年月日
ToDay=Day(Now()) 
ToMonth=Month(Now())
ToYear=Year(Now())
		
'指定的年月日及星期
ThisMonth=C_Month
thisdate=C_Day
ThisYear=C_Year
If IsDate("February 29, " & ThisYear) Then MDays(2)=29
		
'确定日历1号的星期
startspace=WeekDay( ThisMonth&"-1-"&ThisYear )-1
NextMonth=C_Month+1
NextYear=C_Year
If NextMonth>12 Then 
   NextMonth=1
   NextYear=NextYear+1
End If
ProMonth=C_Month-1
ProYear=C_Year
If ProMonth<1 Then 
   ProMonth=12
   ProYear=ProYear-1
End If
		
Response.Write "<a href='SpaceCal.asp?UserName="&UserName&"&amp;Month="&ProYear&"-"&ProMonth&"&amp;" & KS.WapValue & "'>"&ProMonth&"月</a>"
Response.Write " "& ThisYear &"年" & ThisMonth & "月 "
Response.Write "<a href='SpaceCal.asp?UserName="&UserName&"&amp;Month="&NextYear&"-"&NextMonth&"&amp;" & KS.WapValue & "'>"&NextMonth&"月</a><br/>"
Response.Write "日 一 二 三 四 五 六<br/>"
For S=0 To startspace-1
    Response.Write "　 "
Next
Count=1
While Count<=MDays(ThisMonth)
      For B=startspace To 6
	      Response.Write " "
		  LinkTrue="False"
		  For C=0 To Ubound(LinkDays,2)
		      If LinkDays(0,C)<>"" Then
			     If LinkDays(0,C)=ThisMonth And LinkDays(1,C)=Count Then
				    Response.Write "<a href='"&LinkDays(2,C)&"'>"
				    LinkTrue="True"
				 End If
			  End If
		  Next
		  If Count<=MDays(ThisMonth) Then
		     If Count="1" or Count="2" or Count="3" or Count="4" or Count="5" or Count="6" or Count="7" or Count="8" or Count="9" Then
			    Response.Write "0"&Count
			 Else
			    Response.Write Count
			 End If
		  End If
		  If LinkTrue="True" Then Response.Write "</a>"
		  Count=Count+1
	  Next
	  Response.Write "<br/>"
	  startspace=0
Wend
Response.write "---------<br/>"
IF Cbool(KSUser.UserLoginChecked)=True Then Response.write "<a href=""KS_Cls/Index.asp?" & KS.WapValue & """>我的地盘</a>"
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>"
Response.write KS.CopyRight
Call CloseConn
Set KSUser=Nothing
Set KS=Nothing
%>
</p>
</card>
</wml>

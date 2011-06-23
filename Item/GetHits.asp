<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS
Set KS=New PublicCls
Dim ChannelID,ID,Hits,RS,SqlStr,HitsByDay,HitsByWeek,HitsByMonth,Action

ChannelID=KS.ChkClng(KS.S("M"))
ID = KS.ChkClng(KS.S("ID"))
Action=KS.G("Action")
 If ID = 0 Or ChannelID=0 Then
        Hits = 0
 Else
       Set RS = Server.CreateObject("ADODB.Recordset")
        SqlStr = "SELECT top 1 Hits,HitsByDay,HitsByWeek,HitsByMonth,LastHitsTime FROM [" & KS.C_S(ChannelID,2) & "] Where ID=" & ID
	   If KS.C_S(ChannelID,6)=3 Then
        RS.Open SqlStr, conn, 1, 3
        If RS.bof And RS.EOF Then
            Hits = 0
        Else
			Hits=rs(0)
			HitsByDay=rs(1)
			HitsByWeek=rs(2)
			HitsByMonth=rs(3)
        End If
       Else
        RS.Open SqlStr, conn, 1, 3
        If RS.bof And RS.EOF Then
            Hits = 0
        Else
		    IF Action="Count" Then
             rs(0) = rs(0) + 1
             If KS.ChkClng(DateDiff("Ww", rs(4), Now())) <= 0 Then
                rs(2) = rs(2) + 1
			 Else
			    rs(2) = 1
             End If
             If DateDiff("M", rs(4), Now()) <= 0 Then
                rs(3) = rs(3) + 1
			 Else
			    rs(3) = 1
             End If
			 If DateDiff("D", rs(4), Now()) <= 0 Then
                rs(1) = rs(1) + 1
			 Else
			    rs(1) = 1
				rs(4) = Now()
             End If
            rs.Update
			Conn.Execute("Update [KS_ItemInfo] Set Hits=" & RS(0) & ",HitsByDay=" & RS(1) & ",HitsByWeek=" & RS(2) & ",HitsByMonth=" & RS(3) & ",LastHitsTime=" & SQLNowString&" Where channelid=" & ChannelID & " and InfoID=" & ID)
		   End IF
			Hits=rs(0)
			HitsByDay=rs(1)
			HitsByWeek=rs(2)
			HitsByMonth=rs(3)
        End If

	 End If
	 rs.Close:Set rs = Nothing 
End If

	Select Case  KS.ChkClng(KS.S("GetFlag"))
	 Case 0
	  Response.Write "document.write('" & Hits & "');"
	 Case 1
	  Response.Write "document.write('" & HitsByDay & "');"
	 Case 2
	  Response.Write "document.write('" & HitsByWeek & "');"
	 Case 3
	  Response.Write "document.write('" & HitsByMonth & "');"
	End Select


Call CloseConn()
Set KS=Nothing
%> 

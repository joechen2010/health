<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit
response.Buffer=true
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="base64.asp"-->
<%
Dim KSCls
Set KSCls = New DownLoad
KSCls.Kesion()
Set KSCls = Nothing

Class DownLoad
        Private KS,KSUser, KSRFObj,ChannelID
		Private FileContent,RSObj,SqlStr,ShowInfoStr,InfoPurview,ReadPoint,ChargeType,PitchTime,ReadTimes
		Private DomainStr,ID,ClassPurview,UserLoginTF,PayTF,DownUrlTF,TitleStr,Rs,SQL,FoundErr,SoftName,DownID,Hits

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
		    ChannelID=KS.ChkClng(Request("m"))
			If ChannelID=0 Then Response.End()
		    DownUrlTF=false
			DomainStr=KS.GetDomain
		    UserLoginTF=Cbool(KSUser.UserLoginChecked)
			ID = KS.ChkClng(KS.S("ID"))
			DownID = KS.ChkClng(KS.S("DownID"))
			PayTF=KS.S("PayTF")
			
			If ID = 0 Then
			    TitleStr="���ش�����ʾ"
				ShowInfoStr = ShowInfoStr & "<li>�����ϵͳ����!��������ȷ��" & KS.C_S(ChannelID,3) & "ID</li>"
				FoundErr=True
			End If
			If DownID = 0 Then
			    TitleStr="���ش�����ʾ"
				ShowInfoStr = ShowInfoStr & "<li>�����ϵͳ����!��������ȷ��" & KS.C_S(ChannelID,3) & "ID</li>"
				FoundErr=True
			End If
			If Not KS.CheckOuterUrl Then
				ShowInfoStr = ShowInfoStr & "<li>�Ƿ����أ��벻Ҫ������վ��Դ��</li>"
				FoundErr=True
			End If
			
			 If FoundErr Then Call ShowInfo :Exit Sub
			 SqlStr= "Select a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID
			 Set RSObj=Server.CreateObject("Adodb.Recordset")
			 RSObj.Open SqlStr,Conn,1,3
			 IF RSObj.Eof And RSObj.Bof Then
			      TitleStr="���ش�����ʾ"
				  ShowInfoStr = ShowInfoStr & "<li>�Ҳ�����Ҫ���ص�" & KS.C_S(ChannelID,3) & "��</li>"
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
					   ShowInfoStr = ShowInfoStr & "<li>�Բ�����û�����ر�" & KS.C_S(ChannelID,3) & "��Ȩ��!</li>"
					   FoundErr=True:Call ShowInfo :Exit Sub
					 Else
						  Call PayPointProcess()
					 End If
			   End If
		 ElseIF InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else
			  	'============�̳���Ŀ�շ�����ʱ,��ȡ��Ŀ�շ�����===========
			     ReadPoint=Cint(RSObj("DefaultReadPoint"))   
				 ChargeType=Cint(RSObj("DefaultChargeType"))
				 PitchTime=Cint(RSObj("DefaultPitchTime"))
				 ReadTimes=Cint(RSObj("DefaultReadTimes"))
				 '============================================================
        
				 If ClassPurview=2 Then
					 IF KS.FoundInArr(RSObj("DefaultArrGroupID"),KSUser.GroupID,",")=false Then
					    ShowInfoStr="<div align=center>�Բ��������ڵ��û���û�����ص�Ȩ��!</div>"
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
		   
			   on error resume next
		       Dim DownArr:DownArr=Split(Split(RSObj("DownUrls"),"|||")(DownID-1),"|")
			   if err then
			     response.write "�Ƿ�����"
				 response.end
			   end if
			   If DownArr(0)="0" Then
			    '	ShowInfoStr = "<a href=""" & DownArr(2) & """><font color=blue>�������� --- " & RSObj("Title") & "</font></a>"
			     Response.Redirect(DownArr(2))
			   Else
					Set Rs = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT AllDownHits,DayDownHits,HitsTime FROM KS_DownSer WHERE downid="& KS.ChkClng(KS.S("Sid"))
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
			   
			   
			    Dim Url
			     Dim RS_S:Set RS_S=Server.CreateObject("ADODB.RECORDSET")
				 RS_S.Open "Select IsOuter,DownloadPath,UnionID From KS_DownSer Where DownID=" & KS.ChkClng(KS.S("Sid")),conn,1,1
				 If Not RS_S.Eof Then
				  url=DownArr(2)
				  if left(lcase(url),4)<>"http" then url=RS_S(1) & URL
				  Select Case RS_S(0)
				   Case 0
				   	   Response.Redirect(URL)
				   Case 2
					 Call ThunderDownloadUrl(ThunderEncode(URL),RS_S(2))
				   Case 3
					 Call FlashGetDownloadUrl(URL,RS_S(2))
				  End Select
				 End If
				 RS_S.Close:Set RS_S=Nothing
			   End If
			 
		   Else
		     TitleStr="������ʾ"
		   End If
		   Call ShowInfo()
		   RSObj.Close:Set RSObj=Nothing
	   End Sub


      '�շѿ۵㴦�����
	   Sub PayPointProcess()
					    If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) and KSUser.UserName<>RSObj("Inputer") Then
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
						     ShowInfoStr="<div align=center>�Բ�������˻��ѹ��� <font color=red>" & KSUser.GetEdays & "</font> ��,��" & KS.C_S(ChannelID,3) & "��Ҫ����Ч���ڲſ������أ��뼰ʱ��������ϵ��</div>"
						  Else
						   Call GetContent()
						  End If
						Else
						 Call GetContent()
						end if   	  
					   Else
						  Call GetContent()
					   End IF
	   End Sub
	   '����Ƿ���ڣ��������Ҫ�ظ��۵�ȯ
	   '����ֵ ���ڷ��� true,δ���ڷ���false
	   Sub CheckPayTF(Param)
	    Dim SqlStr:SqlStr="Select top 1 Times From KS_LogPoint Where ChannelID=" & ChannelID & " And InfoID=" & ID & " And InOrOutFlag=2 and UserName='" & KSUser.UserName & "' And (" & Param & ") Order By ID"
		'response.write sqlstr 
		'response.end
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
					 ShowInfoStr="<div align=center>�Բ�����Ŀ���" & KS.Setting(45) & "����!���ر�" & KS.C_S(ChannelID,3) & "��Ҫ <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",�㻹�� <font color=green>" & KSUser.Point & "</font> " & KS.Setting(46) & KS.Setting(45) & "</div>,�뼰ʱ��������ϵ��" 
			 Else
					If PayTF="yes" Then
						IF Cbool(KS.PointInOrOut(ChannelID,RSObj("ID"),KSUser.UserName,2,ReadPoint,"ϵͳ","�����շ�" & KS.C_S(ChannelID,3) & "��" & RSObj("Title") & "��",0))=True Then 
						'֧��Ͷ�������
						 Dim PayPoint:PayPoint=(ReadPoint*KS.C_C(RSObj("Tid"),11))/100
						 If PayPoint>0 Then
						 Call KS.PointInOrOut(ChannelID,RSObj("ID"),RSObj("Inputer"),1,PayPoint,"ϵͳ",KS.C_S(ChannelID,3) & "��" & RSObj("Title") & "�������",0)
						 End If
						Call GetContent()
						End If
					Else
						ShowInfoStr="<div align=center>���ر������Ҫ���� <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",��Ŀǰ���� <font color=green>" & KSUser.Point & "</font> " & KS.Setting(46) & KS.Setting(45) &"����,���ر�" & KS.C_S(ChannelID,3) & "������ʣ�� <font color=blue>" & KSUser.Point-ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &"</div><div align=center>��ȷʵԸ�⻨ <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) & "�����ر�" & KS.C_S(ChannelID,3) & "��?</div><div>&nbsp;</div><div align=center><a href=""?m=" &ChannelID & "&ID=" & ID & "&PayTF=yes&DownID=" & DownID & """>��Ը��</a>    <a href=""" &DomainStr & """>�Ҳ�Ը��</a></div>"
					End If
			 End If
	   End Sub
	   Sub GetNoLoginInfo()
		   ShowInfoStr="<div align=center>�Բ����㻹û�е�¼����" & KS.C_S(ChannelID,3) & "����Ҫ��վ��ע���Ա�ſ�����!</div><div align=center>����㻹û��ע�ᣬ��<a href=""" & DomainStr & "User/reg/""><font color=red>���ע��</font></a>��!</div><div align=center>��������Ǳ�վע���Ա���Ͻ�<a href=""" & domainstr & "user/login/""><font color=red>��˵�¼</font></a>�ɣ�</div>"
	   End Sub
	   Sub GetContent()
		 TitleStr=RSObj("Title")
		 DownUrlTF=True
	   End Sub
			
	  Function ShowInfo()
			   With Response
				.Write "<html><head><title>" & TitleStr & "</title>" & vbNewLine
				.Write "<script>"&vbnewline
                .Write " <!--" & vbNewLine
                .Write " window.moveTo(100,100);" & vbNewLine
                .Write " window.resizeTo(550,400);" & vbNewLine
                .Write "//-->" & vbNewLine
                .Write "</script>" & vbNewLine
				.Write "<meta http-equiv=Content-Type content=text/html; charset=gb2312>" & vbNewLine
				.Write "<style type=""text/css"">" & vbNewLine
				.Write "body {font-size: 12px;font-family: ����;}" & vbNewLine
				.Write "td {font-size: 12px; font-family: ����; line-height: 18px;table-layout:fixed;word-break:break-all}" & vbNewLine
				.Write "a {color: #555555; text-decoration: none}" & vbNewLine
				.Write "a:hover {color: #FF8C40; text-decoration: underline}" & vbNewLine
				.Write "th{ background-color: #0A95D2;color: white;font-size: 12px;font-weight:bold;height: 25;}" & vbNewLine
				.Write ".TableRow1 {background-color:#F7F7F7;}" & vbNewLine
				.Write ".TableRow2 {background-color:#F0F0F0;}" & vbNewLine
				.Write ".TableBorder {border: 1px #3795D2 solid ; background-color: #FFFFFF;font: 12px;}" & vbNewLine
				.Write "</style>" & vbNewLine
				.Write "</head><body><br /><br />" & vbNewLine
				.Write "<table width=500 border=0 align=center cellpadding=0 cellspacing=0 class=TableBorder>"
				.Write "<tr>"
				.Write "  <th>ϵ ͳ �� ʾ</th>"
				.Write "</tr>"
				.Write "<tr height=110>"
				.Write "<td class=TableRow1 align=center>"  & ShowInfoStr & "</td>"
				.Write "</tr>"
				.Write "<tr height=22><td align=center class=TableRow2><a href=""" & KS.GetDomain & """>������ҳ...</a> | <a href=javascript:window.close()>�رձ�����...</a></td></tr>"
				.Write "</table>"
				.Write "<br /><br /></body></html>"
			  End With
	End Function
			
Function ThunderDownloadUrl(url,unionid)
	Response.Write "<script src='http://pstatic.xunlei.com/js/webThunderDetect.js'></script>" & vbNewLine
	Response.Write "<script>OnDownloadClick('" & url & "','',location.href,'" & UnionID & "',false)</script>" & vbNewLine
	Response.Write "<script>window.setInterval(""window.close()"",100);</script>" & vbCrLf
End Function

Function FlashGetDownloadUrl(url,unionid)
	Dim m_strFlashGetUrl,m_strDownUrl
	m_strDownUrl = url   
	m_strFlashGetUrl = FlashgetEncode(m_strDownUrl,UnionID)
	Response.Write "<script src=""http://ufile.kuaiche.com/Flashget_union.php?fg_uid=" & UnionID & """></script>" & vbCrLf
	Response.Write "<script>function ConvertURL2FG(url,fUrl,uid){	try{		FlashgetDown(url,uid);	}catch(e){		location.href = fUrl;		}}"& vbCrLf
	Response.Write "function Flashget_SetHref(obj){obj.href = obj.fg;}</script>"& vbCrLf
	Response.Write "<script>ConvertURL2FG('" & m_strFlashGetUrl & "','" & m_strDownUrl & "'," & UnionID & ")</script>" & vbCrLf
	Response.Write "<script>window.setInterval(""window.close()"",100);</script>" & vbCrLf
End Function

End Class
			%>
 

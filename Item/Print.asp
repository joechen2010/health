<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS:Set KS=New PublicCls
Dim Channelid,ID,RS,ArticleContent,PayTF

ChannelID=KS.ChkClng(KS.S("M"))
ID=KS.ChkClng(KS.S("ID"))
if ID=0 Or ChannelID=0 then
	Response.Write"<script>alert(""����Ĳ�����"");location.href=""javascript:history.back()"";</script>"
    Response.End
end if
Set RS=Server.CreateObject("ADODB.RECORDSET")
RS.Open "Select a.*,ClassPurview From "& KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID,Conn,1,1
IF RS.EOF AND RS.BOF THEN
  RS.CLOSE:SET RS=NOthing
  Call CloseConn()
  Set KS=Nothing
 	Response.Write"<script>alert(""����Ĳ�����"");location.href=""javascript:history.back()"";</script>"
    Response.End
END IF
	Dim InfoPurview:InfoPurview=Cint(RS("InfoPurview"))
	Dim ReadPoint:ReadPoint=Cint(RS("ReadPoint"))
	Dim ChargeType:ChargeType=Cint(RS("ChargeType"))
	Dim PitchTime:PitchTime=Cint(RS("PitchTime"))
	Dim ReadTimes:ReadTimes=Cint(RS("ReadTimes"))
	Dim ClassID:ClassID=RS("Tid")
	Dim KSUser:Set KSUser=New UserCls
	Dim UserLoginTF:UserLoginTF=Cbool(KSUser.UserLoginChecked)
	    
		 If ReadPoint>0 Then
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
				 Call PayPointProcess()
			   End If
		 ElseIf InfoPurview=2  Then 
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
					 IF InStr(RS("ArrGroupID"),KSUser.GroupID)=0 Then
					   ArticleContent="<div align=center>�Բ�����û�в鿴���ĵ�Ȩ��!</div>"
					 Else
						  Call PayPointProcess()
					 End If
			   End If
		 ElseIF InfoPurview=0 And (RS("ClassPurview")=1 Or RS("ClassPurview")=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else        
				Call PayPointProcess()
			  End If
		 Else
		   Call PayPointProcess()
		 End If   

	   '�շѿ۵㴦�����
	   Sub PayPointProcess()
	     Dim UserChargeType:UserChargeType=KSUser.ChargeType
					   If Cint(ReadPoint)>0 Then
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff('h',AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If KSUser.GetEdays <=0 Then
						     ArticleContent="<div align=center>�Բ�������˻��ѹ��� <font color=red>" & Edays & "</font> ��,������Ҫ����Ч���ڲſ��Բ鿴���뼰ʱ��������ϵ��</div>"
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
			 If Cint(KSUser.Point)<ReadPoint Then
					 ArticleContent="<div align=center>�Բ�����Ŀ���" & KS.Setting(45) & "����!�Ķ�������Ҫ <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",�㻹�� <font color=green>" & KSUser.Point & "</font> " & KS.Setting(46) & KS.Setting(45) & "</div>,�뼰ʱ��������ϵ��" 
			 Else
					If PayTF="yes" Then
						IF Cbool(KS.PointInOrOut(ChannelID,RS("ID"),KSUser.UserName,2,ReadPoint,"ϵͳ","�Ķ��շ�" & KS.C_S(ChannelID,3) & "��<br>" & RS("Title"),0))=True Then Call GetContent()
					Else
						ArticleContent="<div align=center>�Ķ�������Ҫ���� <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",��Ŀǰ���� <font color=green>" & KSUser.Point & "</font> " & KS.Setting(46) & KS.Setting(45) &"����,�Ķ����ĺ�����ʣ�� <font color=blue>" & KSUser.Point-ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &"</div><div align=center>��ȷʵԸ�⻨ <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) & "���Ķ�������?</div><div>&nbsp;</div><div align=center><a href=""?ID=" & ID & "&PayTF=yes&Page=" & CurrPage &""">��Ը��</a>    <a href=""" &DomainStr & """>�Ҳ�Ը��</a></div>"
					End If
			 End If
	   End Sub
	   Sub GetNoLoginInfo()
		   ArticleContent="<div align=center>�Բ����㻹û�е�¼����������Ҫ��վ��ע���Ա�ſɲ鿴!</div><div align=center>����㻹û��ע�ᣬ��<a href=""../User/reg/""><font color=red>���ע��</font></a>��!</div><div align=center>��������Ǳ�վע���Ա���Ͻ�<a href=""../user/login/""><font color=red>��˵�¼</font></a>�ɣ�</div>"
	   End Sub
	   Sub GetContent()
	     ArticleContent=Replace(RS("ArticleContent"),"[NextPage]","")
	   End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD><TITLE><%=rs("title")%>-��ӡ����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript type=text/JavaScript>
function resizepic(thispic)
{
if(thispic.width>700) thispic.width=700;
}
//˫����������Ļ�Ĵ���
var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",30);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<META content="MSHTML 6.00.3790.2577" name=GENERATOR></HEAD>
<BODY onmouseup=document.selection.empty() oncontextmenu="return false" onselectstart="return false" ondragstart="return false"onbeforecopy="return false" oncopy=document.selection.empty() leftMargin=0 topMargin=0 onselect=document.selection.empty() marginheight="0" marginwidth="0">
<TABLE width=760 height="100%" border=0 align=center cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF" class=center_tdbgall style="WORD-BREAK: break-all">
  <TBODY>
  <TR>
    <TD class=main_title_760 align=right height=20><A class=class 
      href="javascript:window.print()"><IMG src="../Images/Default/printpage.gif" alt=��ӡ���� border=0 align=absMiddle>&nbsp;��ӡ����</A>&nbsp;&nbsp;<IMG alt=�رմ��� src="../Images/Default/pageclose.gif" align=absMiddle border=0>&nbsp;<A 
      class=class href="javascript:window.close()">�رմ���</A>&nbsp;&nbsp; </TD>
  </TR>
  <TR>
    <TD height="25" align=middle class=main_ArticleTitle><B><%=RS("Title")%></B></TD>
  </TR>
  <TR>
    <TD height="25" 
      align=middle class=Article_tdbgall>���ߣ�<%=RS("Author")%>&nbsp;&nbsp;������Դ��<%=RS("Origin")%>&nbsp;&nbsp;�����
      <%=RS("Hits")%>&nbsp;&nbsp;����ʱ�䣺<%=RS("AddDate")%>&nbsp;&nbsp;����¼�룺<%=RS("Inputer")%>
  </TD></TR>
  <TR>
    <TD height="25">
      <HR align=center width="100%" color=#8ea7cd noShade SIZE=1>
    </TD>
  </TR>
  <TR>
    <TD valign="top"><%=KS.HtmlCode(ArticleContent)%></TD>
  </TR> 
</TBODY>
</TABLE>
</BODY>
</HTML> 
<%set ks=nothing
conn.close
set conn=nothing
%>
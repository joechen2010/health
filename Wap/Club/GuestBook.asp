<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.PublicCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 4.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,ListStr
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno
	    Private KeyWord, SearchType,GuestCheckTF,SqlStr
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If KS.Setting(56)="0" Then response.write "��վ�ѹر����Թ���":response.end
			If KS.Setting(59)="0" Then response.redirect("index.asp")
			GuestCheckTF=KS.Setting(52)
			KeyWord = KS.R(Trim(KS.S("keyword")))
			SearchType = KS.R(Trim(KS.S("SearchType")))
		    Dim FileContent,KMRFObj
			Set KMRFObj = New Refresh
		          If KS.Setting(114)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
				   FileContent = KMRFObj.LoadTemplate(KS.Setting(114))
				   If Trim(FileContent) = "" Then FileContent = "ģ�岻����!"
				   FCls.RefreshType = "guestindex" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   Call GetList()
				   FileContent=Replace(FileContent,"{$GetGuestList}",ListStr & PageList())
				  ' FileContent=Replace(FileContent,"{$PageStr}",PageList())
				   FileContent=Replace(FileContent,"{$GuestTitle}","��վ����")
				   response.write FileContent
		End Sub
		
	  Sub GetList()
		 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.RecordSet")
		 Dim Param
		 If GuestCheckTf=0 Then
		 Param=" where 1=1"
		 Else
		 Param=" where verific=1"
		 END iF
		 If KeyWord<>"" Then
		   If SearchType="1" Then
		    Param=Param & " and subject like '%" & keyword & "%'"
		   Else
		    Param=Param & " and username like '%" & keyword & "%'"
		   End If
		 End If
		 
		
		SqlStr = "SELECT * From KS_GuestBook " & Param & " ORDER BY ID DESC" 
			
	RSObj.Open SqlStr,Conn,1,1
	
	Dim Pmcount:Pmcount = KS.Setting(51)
	If KS.ChkClng(Pmcount) < 1 Then Pmcount = 10

	RSObj.Pagesize = Pmcount
	TotalPut = RSObj.RecordCount	'��¼���� 
	TotalPage = RSObj.PageCount	    '�õ���ҳ��
	MaxPerPage = RSObj.PageSize	    '����ÿҳ��
		
	CurrentPage = KS.ChkClng(KS.S("Page"))
	
	If CDbl(CurrentPage) < 1 Then CurrentPage = 1
	If CDbl(CurrentPage) > CDbl(TotalPage) Then CurrentPage = TotalPage

	If RSObj.Eof or RSObj.Bof Then 
		ListStr = "<div style='color:#FF0000;margin:10px;text-align:center;border:1px solid #efefef;height:50px;line-height:50px'>��ʱ��û���κ����ԣ�</div>"
	Else
		RSObj.Absolutepage = CurrentPage	'��ָ������ָ��ҳ�ĵ�һ����¼
		Loopno = MaxPerPage
		i = 0
Do While Not RSObj.Eof and Loopno > 0
          ListStr = ListStr & " <table width='100%' border='1' cellspacing='0' cellpadding='2' align='center' bordercolordark='#FFFFFF' bordercolorlight='#DDDDDD' style='word-break:break-all;font-family: Arial, Helvetica, sans-serif;'>" & vbcrlf
          ListStr = ListStr & " <tr>"
          ListStr = ListStr & "<td width='100' align='center' bgcolor='#F5F5F5' rowspan='3' ><font face='Arial, Helvetica, sans-serif'>��<font color='#FF0000'>" & ((TotalPut)-(MaxPerPage)*(CurrentPage-1))-i & "</font>������<br><img src='../images/Face/" & RSObj("Face") & "'><br></font>" &vbcrlf
          ListStr = ListStr & " <table width='98%'  border='0' align='center' cellpadding='0' cellspacing='0' bgcolor='#FFFFFF'>" & vbcrlf
          ListStr = ListStr & "                       <tr>"
          ListStr = ListStr & "                        <td align='center' bgcolor='#F5F5F5'><font face='Arial, Helvetica, sans-serif'>"& RSObj("UserName") & "</font></td></tr></table></td>" & vbcrlf
          ListStr = ListStr & "                <td height='25' valign='middle'><table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbcrlf
          ListStr = ListStr & "                    <tr>"
          ListStr = ListStr & "                      <td width='49%'><img src='../images/Face1/" & RSObj("TxtHead") & "' align='absmiddle'> ���⣺" & RSObj("Subject") & "</td><td width='51%' align='right'>" & vbcrlf
		    If RSObj("HomePage") <> "" and RSObj("HomePage") <> "http://" Then
          ListStr = ListStr & "      <a href='" & RSObj("HomePage") & "' target='_blank'><img src='images/home.gif' width='16' height='16' border='0' align='absmiddle' alt='��ҳ:[ " & RSObj("HomePage") & " ]'></a>"
            Else
          ListStr = ListStr & "      <a href='#'><img src='images/home-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='��ҳ'></a>" &vbcrlf
            End If
             ListStr = ListStr & "                     |" 
             If RSObj("Email") <> "" Then
           ListStr = ListStr & "                       <a href='mailto:" & RSObj("Email") & "' target='_blank'><img src='images/email.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�:[ " & RSObj("Email") &" ]'></a>" & vbcrlf
             Else
           ListStr = ListStr & "                       <a href='#'><img src='images/email-gray.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�'></a>" & vbcrlf
            End If
             ListStr = ListStr & "                     |" 
            If RSObj("Oicq") <> "" and RSObj("Oicq") <> "0" Then
            ListStr = ListStr & " <a href='#'><img src='images/qq.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����:[ " & RSObj("Oicq") & " ]'></a>"
            Else
            ListStr = ListStr & "  <a href='#'><img src='images/qq-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����'></a>" & vbcrlf
            End If
             ListStr = ListStr & "                     |" 
             If RSObj("GuestIP") <> "" Then
            ListStr = ListStr & " <a href='#'><img src='images/ip.gif' width='16' height='16' border='0' align='absmiddle' alt='���ԣ�[ " & RSObj("GuestIP") & " ]'></a>" & vbcrlf
             Else
            ListStr = ListStr & " <a href='#'><img src='images/ip-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='���ԣ�'></a>" & vbcrlf
            End If
             ListStr = ListStr & "                     &nbsp; </td>"
             ListStr = ListStr & "                 </tr>"
             ListStr = ListStr & "             </table></td>"
             ListStr = ListStr & "           </tr>"
             ListStr = ListStr & "           <tr>"
             ListStr = ListStr & "             <td height='45'>&nbsp;" & KS.HtmlCode(RSObj("Memo"))& " </td></tr>" & vbcrlf
             ListStr = ListStr & "           <tr><td height='20' align='right'>����ʱ�䣺" & RSObj("AddTime") & "&nbsp; </td></tr>" & vbcrlf
			 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			 rs.open "select content,replaytime,txthead from ks_guestreply where topicid=" & rsobj("id"),conn,1,1
			 if not rs.eof then
             ListStr = ListStr & "           <tr>"
             ListStr = ListStr & "             <td width='100' align='center' bgcolor='#F5F5F5'>����ظ���</td>" & vbcrlf
             ListStr = ListStr & "             <td><img src='../images/Face1/face" & rs(2) & ".gif' align='absmiddle'>&nbsp;<font color=red>" & RS(0) & "</font><div align=right>�ظ�ʱ�䣺&nbsp;" & Rs(1) & "</div></td></tr>"
             End If
			 rs.close:set rs=nothing
             ListStr = ListStr & "         </table><br>" & vbcrlf

	RSObj.MoveNext
	Loopno = Loopno-1
	i = i+1
	Loop
End if
	RSObj.Close:Set RSObj=Nothing
 End Sub
 
 Function PageList()
    PageList= "<table width=""100%"" aling=""center""><tr><td align=right>" & KS.ShowPagePara(totalPut, MaxPerPage, "", True, "������", CurrentPage, KS.QueryParam("Page")) & "</td></tr></table>"
 End Function
					  
End Class
%>

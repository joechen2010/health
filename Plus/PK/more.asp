<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,c_str,PKID,Template,role
		Private TotalPut,CurrentPage,MaxPerPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Dim I
		   PKID=KS.ChkClng(Request("pkid"))
		   Role=KS.ChkClng(Request("role"))
		   If PKID=0 Then 
		     ks.die "�Ƿ�����!"
		   End If
		   Template = KSR.LoadTemplate(KS.Setting(104))
		   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		   Call GetSubject()
		   if role=1 then
		     Template=replace(template,"{$GDType}","�����۵�")
		   elseif role=2 then
		     Template=replace(template,"{$GDType}","�����۵�")
		   else
		     Template=replace(template,"{$GDType}","�������۵�")
		   end if
		   ShowMessageList
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		
		Sub GetSubject()
		      Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			  RS.Open "select top 1 * from KS_PKZT where id=" & PKID,conn,1,1
			  If RS.Eof And RS.Bof Then
			    RS.Close
				Set RS=Nothing
				KS.Die "�Ҳ���PK����!"
			  End If
			  Template=replace(template,"{$GetPKID}",rs("id"))
			  Template=replace(template,"{$GetPKTitle}",rs("title"))
			  If KS.IsNul(rs("newslink")) Then
			  Template=replace(template,"{$GetBackGroundNews}","")
			  Else
			  Template=replace(template,"{$GetBackGroundNews}","<a href='" & rs("newslink") & "' target='_blank'>����鿴�������� >></a>")
			  End If
			  Template=replace(template,"{$GetZFTips}",rs("zftips"))
			  Template=replace(template,"{$GetFFTips}",rs("fftips"))
		End Sub
		
		
		Sub ShowMessageList()
		  CurrentPage=KS.ChkClng(request("page"))
		  If CurrentPage=0 Then CurrentPage=1
		  dim rs,UserIP,ipstr,i,content,FaceStr
		  set rs=server.createobject("adodb.recordset")
		  rs.open "select a.*,b.userface from KS_PKGD a left join ks_user b on a.username=b.username where pkid=" & pkId &" and role=" & role & " order by id desc",conn,1,1
		   if rs.eof then
			 c_str=c_str & "û�����ύ���ۣ�"
		   else
		 		    TotalPut= rs.recordcount
					If CurrentPage < 1 Then CurrentPage = 1
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (TotalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage>1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
					 dim n:n=0
					Do While Not RS.Eof
						UserIP=split(rs("userip"),".")
						IpStr=""
						for i=0 to ubound(UserIP)
						   if i=3 then
							ipstr=ipstr &"*"
						   else
							ipstr=ipstr &UserIP(i)&"."
						   end if
						next
					   if rs("status")="0" then
						content="�˹۵�δͨ�����!"
					   else
						content=rs("content")
					   end if
					   FaceStr=KS.Setting(2) & "/images/face/0.gif"
					   If Not KS.IsNul(rs("userface")) then
					   	FaceStr=rs("userface")
			            If lcase(left(FaceStr,4))<>"http" then FaceStr=KS.Setting(2) & "/" & FaceStr
                       End If
					
						c_str=c_str &"<div class='Articial'><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" &vbcrlf
						c_str=c_str & " <tr>"
						c_str=c_str &  "	<td width=""5%"" rowspan=""4"" align=""left"" valign=""top""><img src=""" & FaceStr & """ width=""32"" height=""32"" /></td>" &vbcrlf
						c_str=c_str &  "	<td width=""95%"">" &vbcrlf
					    c_str=c_str &   "	<div style=""float:left""><span class=""STYLE1"">��" & rs("username") & "��</span></div></td></tr>" &vbcrlf
						c_str=c_str &   "<tr><td><span class='STYLE2'>" & ipstr & " ����" & rs("adddate") & "</span></td></tr>" & vbcrlf
						c_str=c_str &   "<tr><td height='5'></td></tr>"
						c_str=c_str &   "<tr><td valign='top' class='neirong'>" & content & "</td></tr>" & vbcrlf
						c_str=c_str &   "</table></div>" & vbcrlf
						n=n+1
						if n>=maxperpage or rs.eof then exit do
						RS.MoveNext
				  loop
	       end if
		   rs.close
		   set rs=nothing
		   Template=Replace(Template,"{$ShowCommentList}",c_str)
		   Template=Replace(Template,"{$TotalPut}",totalput)
		   Template=Replace(Template,"{$ShowPage}","<div style='text-align:right'>" &  ShowPagePara(totalput, MaxPerPage, "", true,"��", CurrentPage,KS.QueryParam("page,submit")) & "</div>")
		   
		End Sub
		
		Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>9</span>" '�����һҳ��ť��ʾ��ʽ
				Const Btn_Prev = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>3</span>" '����ǰһҳ��ť��ʾ��ʽ
				Const Btn_Next = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>4</span>" '������һҳ��ť��ʾ��ʽ
				Const Btn_Last = "<span style='font-family:webdings;font-size:14px' title='���һҳ'>:</span>" '�������һҳ��ť��ʾ��ʽ
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
					PageStr = PageStr & ("<div class='showpage' style='height:20px'><form action=""" & FileName & "?" & ParamterStr & """ name=""myform"" method=""post"">�� <font color=red>" & totalnumber & "</font> " & strUnit & "  �� <font color=red>" & N & "</font> ҳ ÿҳ <font color=red>" & MaxPerPage &"</font> " & strUnit &" | ��ǰ�� <font color=red>" & CurrentPage & "</font> ҳ &nbsp;&nbsp;&nbsp;")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("ת��:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  PageStr = PageStr & "</form></div>"
			 ShowPagePara = PageStr
	End Function

		
End Class
%>

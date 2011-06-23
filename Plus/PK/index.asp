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
        Private KS, KSR,str,c_str,ClassID,Template,categoryname
		Private TotalPut,CurrentPage,MaxPerPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  MaxPerPage=10
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Dim I
		   ClassID=KS.ChkClng(KS.S("ClassID"))
	
		   Template = KSR.LoadTemplate(KS.Setting(102))
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   Call GetPKList()
		   
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		
		Sub GetPKList()
		   CurrentPage=KS.ChkClng(request("page"))
		   If CurrentPage=0 Then CurrentPage=1
		  dim rs,UserIP,ipstr,i,content,FaceStr,param
		  if ClassID<>0 then
		    param=" inner join ks_class b on a.classid=b.id where b.ClassID=" & classid
		  end if
		  set rs=server.createobject("adodb.recordset")
		  rs.open "select a.* from KS_PKZT a " & param & " order by a.id desc",conn,1,1
		   if rs.eof then
			 c_str=c_str & "没有PK主题！"
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
					 dim str,url,agreeNum ,argueNum,Total,zf,ff,m
					 m=(currentpage-1)*maxperpage
					Do While Not RS.Eof
					 n=n+1
					 m=M+1
					 agreeNum =rs("zfvotes")
					 argueNum = rs("ffvotes")
					 Total=agreeNum + argueNum+0.002
					 zf=formatpercent((agreeNum+0.001)/Total,2)
					 ff=formatpercent((argueNum+0.001)/Total,2)
					 
					 url="pk.asp?id=" & rs("id")
					 str=str &"<div class='listPk'>" & vbcrlf
					 str=str &"<table width='100%' border='0' cellspacing='0' cellpadding='0'>" & vbcrlf
					 str=str &"	  <tr valign='top'>" & vbcrlf
					 str=str & "	<td class='s14'>" & M &".</td>" & vbcrlf
					 str=str & "		<td ><h3><a href='" & url & "' target='_blank'>" & rs("title") & "</a></h3>" &vbcrlf
					 
					 str=str & "						 <div class='titleN'><span class='line18'>" & rs("adddate") & "</span></div>" &vbcrlf
					 str=str & "							 <table border='0' cellspacing='0' cellpadding='0' height='19' class='number' style='margin:0px;'>" &vbcrlf
					 str=str & "   <tr>" &vbcrlf
					 str=str & "	  <td width='35' align='center' valign='bottom'><span class='red'><a href='" & url & "' target='_blank'>YES</a></span></td>" &vbcrlf
					 str=str &"		  <td width='30' align='center'><h5 class='red'>" & agreeNum & "</h5></td>"
					 str=str & "	  <td width='83'><div class='exponentBj'><table width='70' border='0' align='center' cellpadding='0' cellspacing='0'>" &vbcrlf
					 str=str & "   <tr>" &vbcrlf
					 str=str & "   <td width='" & zf & "' style='border-right:1px #fff solid;'><div class='zhengfang' style='width:100%;'></div></td>" &vbcrlf
					 str=str & "   <td width='" & ff & "'><div class='fanfang' style='width:100%;'></div></td>"&vbcrlf
					 
					 str=str & "  </tr>" &vbcrlf
					 str=str & "</table>" &vbcrlf
					 str=str & "</div></td>"
					 str=str & " <td width='30' align='center'><h5 class='LightGrey01'>" & argueNum &"</h5></td>"&vbcrlf
					 str=str & " <td width='35' align='center' valign='bottom'><span class='LightGrey01'><a href='" & url & "' target='_blank'>NO</a></span></td>" &vbcrlf
					 str=str & "</tr>"
					 str=str & "</table>" &vbcrlf
					 str=str & "<div class='clear'></div>"
					 str=str & "<div class='btx1' style='padding:0px;'><p><a href='" & url & "' target='_blank'>我来参与</a></p></div>" &vbcrlf
					 str=str & "<div class='clear'></div>	</td>"
					 str=str & "</tr>"
					 str=str & "</table>"
					 str=str & "</div>"
					 if n>=maxperpage or rs.eof then exit do
					  RS.MoveNext
					Loop
					RS.Close
					Set RS=Nothing
		  end if
		   Template=Replace(Template,"{$GetPKList}",str)
		   Template=Replace(Template,"{$ShowPage}","<div style='text-align:right'>" &  ShowPagePara(totalput, MaxPerPage, "", true,"条", CurrentPage,KS.QueryParam("page,submit,SB")) & "</div>")
		   	
		End Sub
		
		Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "<span style='font-family:webdings;font-size:14px' title='第一页'>9</span>" '定义第一页按钮显示样式
				Const Btn_Prev = "<span style='font-family:webdings;font-size:14px' title='上一页'>3</span>" '定义前一页按钮显示样式
				Const Btn_Next = "<span style='font-family:webdings;font-size:14px' title='下一页'>4</span>" '定义下一页按钮显示样式
				Const Btn_Last = "<span style='font-family:webdings;font-size:14px' title='最后一页'>:</span>" '定义最后一页按钮显示样式
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
					PageStr = PageStr & ("<div class='showpage' style='height:20px'><form action=""" & FileName & "?" & ParamterStr & """ name=""myform"" method=""post"">共 <font color=red>" & totalnumber & "</font> " & strUnit & "  分 <font color=red>" & N & "</font> 页 每页 <font color=red>" & MaxPerPage &"</font> " & strUnit &" | 当前第 <font color=red>" & CurrentPage & "</font> 页 &nbsp;&nbsp;&nbsp;")
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
						PageStr = PageStr & ("转到:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  PageStr = PageStr & "</form></div>"
			 ShowPagePara = PageStr
	End Function
		
		
End Class
%>

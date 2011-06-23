<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
Dim KSCls
Set KSCls = New Show_Product
KSCls.Kesion()
Set KSCls = Nothing

Class Show_Product
        Private KS,KSBCls,KSRFObj,str
		Private UserName,Template,BlogName
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			UserName=KS.S("i")
			If UserName="" Then Response.End()
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",Conn,1,1
			If RS.Eof And RS.Bof Then
			   Call KS.ShowError("该用户没有开通空间站点！","该用户没有开通空间站点！")
			End If
			BlogName=RS("BlogName")
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""{$ShowTitle}-" & BlogName & """>" &vbcrlf
			Template=Template & KSRFObj.LoadTemplate(KS.WSetting(23))'企业主模板
			Template=KSBCls.ReplaceBlogLabel(RS,Template)
			Template=KSBCls.ReplaceAllLabel(UserName,Template)
			Template=Replace(Template,"{$BlogMain}",ShowNews)
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
			RS.Close:Set  RS=Nothing
		End Sub
		
		Function ShowNews()
		    Dim SQL,i,RS,PhotoUrl
			str="【查看新闻】<br/>"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_EnterPriseNews Where UserName='" & UserName & "' and ID=" & KS.ChkClng(KS.S("ID"))  ,conn,1,1
			If RS.EOF and RS.Bof  Then
			   str=str & "参数传递出错！<br/>"
			Else
			   Template=Replace(Template,"{$ShowTitle}",RS("Title"))
			   str=str & "" & RS("Title") & "<br/>"
			   str=str & "作者：" & UserName & " 时间:" & RS("AddDate") & "<br/>"
			   Dim Content
			   Content=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RS("Content"))))))
			   str=str & KS.ContentPagination(Content,200,"Show_News.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") & "&amp;" & KS.WapValue & "",True,True)
			   str=str & " <br/><anchor>[返回公司动态]<prev/></anchor><br/>"
			End If
			str=str &"<br/>"   
			ShowNews=str
			RS.Close:Set RS=Nothing
		End Function
End Class
%>
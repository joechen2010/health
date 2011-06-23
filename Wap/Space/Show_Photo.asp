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
		Private RS
		Private UserName,Template,BlogName
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
		    Set KSRFObj=New Refresh
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			UserName=KS.S("i")
			If UserName="" Then Response.End()
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",conn,1,1
			If RS.Eof And RS.Bof Then
			   Call KS.ShowError("该用户没有开通空间站点！","该用户没有开通空间站点！")
			End If
			BlogName=RS("BlogName")
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""" & BlogName & "-作品展示"">" &vbcrlf
			Template=Template & KSRFObj.LoadTemplate(KS.WSetting(23))'企业主模板
			Template=KSBCls.ReplaceBlogLabel(RS,Template)
			Template=KSBCls.ReplaceAllLabel(UserName,Template)
			Template=Replace(Template,"{$BlogMain}",ShowPhoto)
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
			RS.Close:Set  RS=Nothing
		End Sub
		
		Function ShowPhoto()
		    Dim SQL,n,RS,PhotoUrlArr,PhotoUrl,t
			str="【查看作品】<br/>"
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Photo Where Inputer='" & UserName & "' and ID=" & KS.ChkClng(KS.S("ID"))  ,Conn,1,1
			If RS.EOF and RS.Bof  Then
			   str=str & "参数传递出错！<br/>"
			Else
			   PhotoUrlArr=Split(RS("PicUrls"),"|||")
			   N=KS.ChkClng(KS.S("N"))
			   If N<0 Then N=0
			   T=Ubound(PhotoUrlArr)
			   If N>=t Then n=0
			   If T=0 Then T=1
			   PhotoUrl=Split(PhotoUrlArr(N),"|")(1)
				If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
				if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
				if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl

			   str=str & "<a href=""Show_Photo.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") & "&amp;N=" & n+1 &"&amp;" & KS.WapValue & """><img src=""" & PhotoUrl &""" alt=""""/></a><br/>"
			   str=str &"浏览:"&RS("Hits")&" 投票:<a href=""../plus/PhotoVote.asp?ID=" & KS.S("ID") & "&amp;" & KS.WapValue & """>投它一票("&RS("Score")&")</a><br/>"
			   str=str & "第" & N+1 & "/" & T & "张 "
			   str=str & "<a href=""Show_Photo.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") &"&amp;N=0"">首页</a> "
			   str=str & "<a href=""Show_Photo.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") &"&amp;N=" & N-1 & "&amp;" & KS.WapValue & """>上张</a> "
			   str=str & "<a href=""Show_Photo.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") &"&amp;N=" & N+1 & "&amp;" & KS.WapValue & """>下张</a> "
			   str=str & "<a href=""Show_Photo.asp?i=" & UserName & "&amp;ID=" & KS.S("ID") &"&amp;N=" & T-1 & "&amp;" & KS.WapValue & """>尾页</a><br/>"
			End If
			str=str &"<br/>"   
			ShowPhoto=str
		End Function
		
  
End Class
%>
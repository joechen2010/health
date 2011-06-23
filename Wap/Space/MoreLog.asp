<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New SpaceMore
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceMore
        Private KS,KSRFObj
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
			Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		    Dim FileContent
			FileContent = KSRFObj.LoadTemplate(KS.WSetting(30))
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			FileContent = Replace(FileContent,"{$ShowMain}",GetLogList())
			Response.Write FileContent  
	    End Sub
		
		Function GetLogList()
		    MaxPerPage = 15
		    Dim ClassID:ClassID=KS.Chkclng(KS.S("ClassID"))
		    Dim IsBest:IsBest=KS.Chkclng(KS.S("IsBest"))
		    If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
		    Else
			   CurrentPage = 1
		    End If
		    str = "【日志查找】<br/>"
		    str = str & "<select name=""ClassID"">"
		    str = str & "<option value=""0"">所有分类</option>"
		    Dim RSC:set RSC=Conn.Execute("select TypeName,TypeID from KS_BlogType order by OrderID")
		    If Not RSC.EOF Then
		     Do While Not RSC.EOF
			    str = str & "<option value=""" & RSC(1) & """>" & RSC(0) & "</option>"
				RSC.Movenext
			 Loop
		    End If
		    RSC.Close:set RSC=Nothing
		    str = str & "</select>"
		  
		    str = str & "标题:<input type=""text"" size=""12"" name=""key""/>"
		    str = str & "<anchor>查找<go href=""MoreLog.asp?"&KS.WapValue&""" method=""post"">"
		    str = str & "<postfield name=""ClassID"" value=""$(ClassID)""/>"
		    str = str & "<postfield name=""key"" value=""$(key)""/>"
		    str = str & "</go></anchor><br/>"

		    str = str & "【日志列表】<br/>"
		    Dim Param:Param=" where status=0"
		    If ClassID<>0 Then Param=Param & " and a.TypeID=" & ClassID
		    If IsBest<>0 Then Param=Param & " and best=1"
		    If KS.S("key")<>"" Then Param=Param & " and Title like '%" & KS.R(KS.S("key")) &"%'"
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "select a.*,b.TypeName from KS_BlogInfo a inner join KS_BlogType b on a.typeid=b.TypeID " & Param & " order by Adddate desc" ,Conn,1,1
		    If RSObj.EOF and RSObj.BOF  Then
		       str = str & "对不起，没有找到日志文章! <br/>"
		    Else
		       TotalPut = RSObj.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSObj.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I
			   Do While Not RSObj.EOF
			      str = str & "·<a href=""List.asp?ID="&RSObj("ID")&"&amp;UserName="&RSObj("UserName")&"&amp;"&KS.WapValue&""">"&KS.GotTopic(RSObj("Title"),32)&"(" & FormatDateTime(RSObj("AddDate"),2) &")</a>"
			      If RSObj("best")=1 Then str = str & "[精]"
                  str = str & "<br/>"
			      RSObj.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   str = str & KS.ShowPagePara(TotalPut, MaxPerPage, "MoreLog.asp", True, "篇", CurrentPage, KS.QueryParam("page"))
		    End If
		    GetLogList = str & "<br/>"
		    RSObj.Close:Set RSObj=Nothing
		End Function
End Class
%>

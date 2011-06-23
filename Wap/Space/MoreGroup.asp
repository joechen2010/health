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
			FileContent = Replace(FileContent,"{$ShowMain}",GetGroupList())
			Response.Write FileContent  
	    End Sub
		
		Function GetGroupList()
		    MaxPerPage = 15
			Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
			Dim recommend:recommend=KS.ChkClng(KS.S("recommend"))
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
			Else
			   CurrentPage = 1
			End If
			str = "【圈子查找】<br/>"
			str = str & "<select name=""ClassID"">"
			str = str & "<option value=""0"">所有分类</option>"
			Dim RSC:set RSC=Conn.Execute("select ClassName,ClassID from KS_TeamClass order by OrderID")
			If Not RSC.EOF Then
			   Do While Not RSC.EOF
			      str = str & "<option value=""" & RSC(1) & """>" & RSC(0) & "</option>"
			      RSC.Movenext
			   Loop
			End If
			RSC.Close:set RSC=Nothing
			str = str & "</select> "
			str = str & "名称:<input type=""text"" size=""12"" name=""key""/> "
			str = str & "<anchor>查找<go href=""MoreGroup.asp?"&KS.WapValue&""" method=""post"">"
			str = str & "<postfield name=""ClassID"" value=""$(ClassID)""/>"
			str = str & "<postfield name=""key"" value=""$(key)""/>"
			str = str & "</go></anchor><br/>"
			str = str & "【圈子列表】<br/>"
			
			Dim Param:Param=" where verific=1"
			If ClassID<>0 Then Param=Param & " and a.ClassID=" & ClassID
			If Recommend<>0 Then Param=Param & " and  rRecommend=1"
			If KS.S("key")<>"" Then Param=Param & " and TeamName like '%" & KS.R(KS.S("key")) &"%'"
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "select a.*,b.classname from KS_team a inner join ks_teamclass b on a.classid=b.classid " & Param & " order by id desc",Conn,1,1
			If RSObj.EOF and RSObj.Bof  Then
			   str = str & "没有用户创建圈子！<br/>"
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
			      str = str & "<a href=""Group.asp?ID="&RSObj("ID")&"&amp;"&KS.WapValue&""">"&RSObj("TeamName")&" ("&Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RSObj("ID") & "  and ParentID=0")(0)&"/"&Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RSObj("ID"))(0)&"/"&Conn.Execute("select Count(UserName)  from KS_TeamUsers where Status=3 and TeamID=" & RSObj("ID"))(0)&")</a>"
				  str = str & "<br/>"
				  RSObj.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   str = str & KS.ShowPagePara(TotalPut, MaxPerPage, "MoreGroup.asp", True, "个", CurrentPage, "ClassID="&ClassID&"&amp;key="&KS.S("key")&"&amp;" & KS.WapValue & "")
			End If
			GetGroupList = str & "<br/>"
			RSObj.Close:Set RSObj=Nothing
		End Function
End Class
%>

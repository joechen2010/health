<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New AddLabelSave
KSCls.Kesion()
Set KSCls = Nothing

Class AddLabelSave
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'主体部分
		Public Sub Kesion()
		With KS
		 .echo "<script src='../../../jquery.js'></script>"
		 .echo ("<body bgcolor=#EAF0F5 scroll=no>")
		Dim TempClassList, InstallDir, CurrPath, FolderID, LabelContent
		Dim LabelID, LabelRS, SQLStr, LabelName, Descript, ParentID, Action, RSCheck, FileUrl, LabelFlag
		  FileUrl = Request("FileUrl") '便于添加完毕后返回
		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
		
		If FolderID = "" Then FolderID = "0"
		Select Case Trim(Request.Form("Action"))
		 Case "Add"
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Descript")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
		
			If LabelFlag = "" Then LabelFlag = 0
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   Response.End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  Response.End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select LabelName From [KS_Label] Where LabelName='" & LabelName & "'", Conn, 1, 1
		
			If Not LabelRS.EOF Then
			  .echo ("<script>alert('标签名称已经存在!');location.href='" & FileUrl & "?Action=Add&FolderID=" & ParentID & "';</script>")
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [KS_Label] Where (ID is Null)", Conn, 1, 3
				LabelRS.AddNew
				  Do While True
					'生成ID  年+10位随机
					LabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					 End If
				  Loop
				 LabelRS("ID") = LabelID
				 LabelRS("LabelName") = LabelName
				 LabelRS("Description") = Descript
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("FolderID") = ParentID
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 0 '指定为系统函数标签
				 LabelRS("OrderID") = 1
				 LabelRS.Update
				 Call KS.FileAssociation(1021,1,LabelContent,0)
				.echo ("<script>if (confirm('成功提示:\n\n添加标签成功,继续添加标签吗?')){parent.location.href='" & FileUrl & "?Action=Add&FolderID=" & ParentID & "';}else{top.frames['MainFrame'].location.href='../Label_Main.asp?FolderID=" & ParentID & "';top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & ParentID & "&OpStr=标签管理 >> 系统函数标签&ButtonSymbol=FunctionLabel';}</script>")
			End If
		Case "Edit"
			LabelID = Trim(Request.Form("LabelID"))
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Descript")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelFlag = "" Then LabelFlag = 0
			If LabelName = "" Then
			   Call KS.AlertHistory("标签名称不能为空!", -1)
			   Set KS = Nothing
			   Response.End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("标签内容不能为空!", -1)
			  Set KS = Nothing
			  Response.End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select LabelName From [KS_Label] Where ID <>'" & LabelID & "' AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  .echo ("<script>alert('标签名称已经存在!');location.href='" & FileUrl & "?LabelID=" & LabelID & "';</script>")
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("Description") = Descript
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("FolderID") = ParentID
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 
				 '遍历所有标签内容，找出所有标签的图片
				 Dim Node,UpFiles,RCls
				 UpFiles=LabelContent
				 if Not IsObject(Application(KS.SiteSN&"_labellist")) Then
				     Set RCls=New Refresh
				     Call Rcls.LoadLabelToCache()
					 Set Rcls=Nothing
				 End If
				 For Each Node in Application(KS.SiteSN&"_labellist").DocumentElement.SelectNodes("labellist")
					   UpFiles=UpFiles & Node.Text
				 Next
				 Call KS.FileAssociation(1021,1,UpFiles,1)
				 '遍历及入库结束
				 
				.echo ("<script>alert('成功提示:\n\n标签修改成功!');top.frames['MainFrame'].location.href='../Label_Main.asp?FolderID=" & ParentID & "';top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & ParentID & "&OpStr=标签管理 >> 系统函数标签&ButtonSymbol=FunctionLabel';</script>")
			End If
		End Select
		 End With
		End Sub
End Class
%> 

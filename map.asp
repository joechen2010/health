<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteMaps
KSCls.Kesion()
Set KSCls = Nothing

Class SiteMaps
        Private KS, KSR,Maps
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
		           Dim FileContent
		           Dim MapTemplatePath:MapTemplatePath=KS.Setting(3) & KS.Setting(90) & "map.html"  '模板地址
				   FileContent = KSR.LoadTemplate(MapTemplatePath)    
				   FCls.RefreshType = "map" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   Call MapList()
				   FileContent=Replace(FileContent,"{$ShowMap}",Maps)
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   response.write FileContent
		End Sub
		
		Sub MapList()
				Dim RS,TreeStr,ID,SqlStr,ClassXml,Node,TJ,SpaceStr,k
				Set  RS=Server.CreateObject("ADODB.Recordset")
				SQLstr = "select a.ID,a.FolderName,a.FolderOrder,a.ClassType,a.ChannelID,a.tj,a.tn,a.adminpurview from KS_Class a inner join ks_channel b on a.channelid=b.channelid where b.channelstatus=1 Order BY root,folderorder"
				RS.Open SQLstr, Conn, 1, 1
				If Not RS.Eof Then Set ClassXml=KS.RsToXml(RS,"row","")
				RS.Close
				Set RS=Nothing
				If IsOBject(ClassXml) Then
				  For Each Node In ClassXML.DocumentElement.SelectNodes("row")
				      TJ=Node.SelectSingleNode("@tj").text
					  If tJ=1 Then
				        TreeStr = TreeStr  & "<li>" & KS.GetClassNP(Node.SelectSingleNode("@id").text)& "</li><br>"
					  Else
						SpaceStr=""
						For k = 1 To TJ - 1
						  SpaceStr = SpaceStr & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
						Next
	                    TreeStr = TreeStr & SpaceStr & "·" & KS.GetClassNP(Node.SelectSingleNode("@id").text) & "<br>"
				      End If
				  Next
				End If
			 Maps=TreeStr
	End Sub
End Class
%>

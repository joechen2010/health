<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/template.asp"-->
<!--#include file="function.asp"-->
<%

Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,KSUser,UserLoginTF,AnonymScore
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
		  Call showmain()
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
		<%
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.ASetting(20))    
			 FCls.RefreshType = "askIndex" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			 FileContent=KSR.KSLabelReplaceAll(FileContent)
			 FileContent=Replace(FileContent,"{$AskMenuList}",ACls.IndexMenulist)
			 Immediate=false
			 Scan FileContent
			 Response.write RexHtml_IF(Templates)
		End Sub
		
		Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
				Case "ask"  
				  echo ACls.ReturnAskConfig(sTokenName)
		    End Select 
        End Sub 
        
End Class
%>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%option explicit%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.StaticCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing
Const AllowSecondDomain=true       '是否允许开启空间二级域名 true-开启 false-不开启


Class SiteIndex
        Private KS, KSR
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
		    If AllowSecondDomain=True Then 
			 SecondDomain
			Else
			 Run
			End If
		End Sub
		
		Sub Run()
			   Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			   If QueryStrings<>"" And Ubound(Split(QueryStrings,"-"))>=1 Then
				 Call StaticCls.Run()
			   Else
				  Dim Template,FsoIndex:FsoIndex=KS.Setting(5)
				  IF Split(FsoIndex,".")(1)<>"asp" Then
					  Response.Redirect KS.Setting(5):Exit Sub
				  Else
						   Template = KSR.LoadTemplate(KS.Setting(110))
						   FCls.RefreshType = "INDEX" '设置刷新类型，以便取得当前位置导航等
						   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
						   Template=KSR.KSLabelReplaceAll(Template)
				 End IF
				 Response.Write Template  
			  End If
			  Set StaticCls=Nothing
		End Sub
		
		Public Sub SecondDomain()
		dim From,gourl,sdomain,title,username,domain
		From = LCase(Request.ServerVariables("HTTP_HOST"))
		
		sdomain = LCase(KS.SSetting(15))
		sdomain = Replace(sdomain,"http://","")
		sdomain = Replace(sdomain,"/","")
		
		dim domain1,domain2
		domain = LCase (from)
		domain = Replace (domain,"http://","")
		domain = Replace (domain,"/","")
			
		if sdomain=domain and sdomain<>"" then
			  title=KS.Setting(1) & "-空间" 
			  gourl="space/index.asp"
		else
			 domain1= Replace (Left (domain,InStr (domain,".")),".","")
			 if Trim (domain1)="" or domain1="www" Then 
			      Run : Exit Sub
			 Else
				  '=====================这里定义其它系统非个人空间的二级域名转向，如论坛等=============================
				  if instr(Request.ServerVariables("SERVER_NAME"),"bbs.kesion.com")>0 then
					 response.redirect KS.GetDomain & "bbs/index.asp"
				  elseif instr(Request.ServerVariables("SERVER_NAME"),"news.kesion.com")>0 then
					 response.redirect KS.GetDomain & "news/"
				  elseif instr(Request.ServerVariables("SERVER_NAME"),"help.kesion.com")>0 then
					 response.redirect KS.GetDomain & "help/"
				  end if
				  '============================================================================
			 End If
			 
			 
			 dim rs:set rs=conn.execute("select top 1 username,blogname from ks_blog where [domain]='" & KS.R(domain1) & "'")
			 if rs.eof and rs.bof then
			     rs.close:set rs=nothing
			     Run : Exit Sub
			 end if
			 title=rs("blogname")
			 domain1=rs("username")
			 rs.close:set rs=nothing
			 domain2= Right(domain,Len(domain)-InStr(domain,"."))
			 gourl="space/?" & domain1
			end if
		   
		   Response.Write ("<html>") & vbcrlf
		   Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />") & vbcrlf
		   Response.Write ("<title>" & title & "</title>") & vbcrlf
		   Response.Write ("<head>") & vbcrlf
		   Response.Write ("</head>") & vbcrlf
		   Response.Write( "<frameset><frame src="""&KS.GetDomain & gourl&"""></frameset>")
		 End Sub
		
End Class
%>

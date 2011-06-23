<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New FullScreen
KSCls.Kesion()
Set KSCls = Nothing

Class FullScreen
        Private KS
		Private Domain,AdminDir
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub  Kesion()
			Dim KSLoginCls
				Set KSLoginCls = New LoginCheckCls1
				KSLoginCls.Run()
			Set KSLoginCls= Nothing
			Dim Style,DomainStr,InstallStr,AdminDirStr,ChannelID
			Style=KS.S("Style")
			ChannelID=KS.S("ChannelID")
			Domain=KS.GetDomain
			AdminDir=KS.Setting(89)
			DomainStr=Replace(KS.Setting(2),"/","\\/")
			InstallStr=Replace(KS.Setting(3),"/","\\/")
			AdminDirStr=Replace(KS.Setting(89),"/","\\/")
			%>
			<HTML>
			<HEAD>
			<TITLE>科汛在线编辑器 - 全屏编辑</TITLE>
			<META http-equiv=Content-Type content="text/html; charset=gb2312">
			<style type="text/css">
			body {	margin: 0px; border: 0px; background-color: buttonface; }
			</style>
			</HEAD>
			<BODY scroll="no" onUnload="Minimize()">
			<form  method="post">
			<input type="hidden" id="Content" name="Content">
			<iframe id="ContentFrame" name="ContentFrame" src="<%=Domain & AdminDir%>KS.Editor.asp?FullScreenFlag=1&ID=Content&Style=<%=Style%>&ChannelID=<%=ChannelID%>" frameborder=0 scrolling=no width="100%" height="100%"></iframe>
			</form>
			<script language=javascript>
			<!--
			 <%IF Style=1 Then%>
			   var TempContentArray=opener.ArticleContentArray;
				 opener.SaveCurrPage();
				 document.forms[0].Content.value='';
				 for (var i=0;i<TempContentArray.length;i++)
				{
					if (TempContentArray[i]!='')
					{
						if (document.forms[0].Content.value=='') document.forms[0].Content.value=TempContentArray[i];
						else document.forms[0].Content.value=document.forms[0].Content.value+'[NextPage]'+TempContentArray[i];
					} 
				}
			<%ELSE%>
			document.forms[0].Content.value =opener.KS_EditArea.document.documentElement.innerHTML;
			<%end if%>
			//alert(document.forms[0].Content.value);
			function Minimize() {
			 frames[0].setMode('EDIT',<%=Style%>,'<%=DomainStr%>','<%=InstallStr%>','<%=AdminDirStr%>');  //返回设定为编辑状态
			<%IF Style=1 Then%>
				 frames[0].SaveCurrPage();
				var TempContentArray=frames[0].ArticleContentArray;
				document.forms[0].Content.value='';
				for (var i=0;i<TempContentArray.length;i++)
				{
					if (TempContentArray[i]!='')
					{
						if (document.forms[0].Content.value=='') document.forms[0].Content.value=TempContentArray[i];
						else document.forms[0].Content.value=document.forms[0].Content.value+'[NextPage]'+TempContentArray[i];
					} 
				}
				  for (var i=0; i<opener.parent.frames.length; i++)
					  {
						if (opener.parent.frames[i].document==opener.document)
						  {
						   for (var j=0;j<opener.parent.document.forms.length;j++)
							   if (opener.parent.document.forms[j].Content!=null)
								{  
							   opener.parent.document.forms[j].Content.value=document.forms[0].Content.value;
							   }
							   opener.parent.frames[i].location.reload();
						 }
					 }
			<%ELSE%>
			   opener.KS_EditArea.document.body.innerHTML=frames[0].KS_EditArea.document.body.innerHTML;
			   opener.ShowTableBorders();
			<%END IF%>   
			   self.close();
			}
			//-->
			</script>
			</BODY>
			</HTML>
<%
 End Sub
End Class
%> 

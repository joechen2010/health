<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Down_Param
KSCls.Kesion()
Set KSCls = Nothing

Class Down_Param
        Private KS,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 With KS
			.echo "<html>"
			.echo "<title>���ػ�����������</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<style type=""text/css"">" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo ".STYLE1 {color: #FF0000}" & vbCrLf
			.echo "-->" & vbCrLf
			.echo "</style>" & vbCrLf
			.echo "</head>"
			
			Dim RS, Action, SQLStr
			Dim DownLb, DownYY, DownSQ, DownPT, JyDownUrl, JyDownWin
			Action = KS.G("Action")
			ChannelID= KS.ChkClng(KS.G("ChannelID"))
			If ChannelID=0 Then ChannelID=3
			If Not KS.ReturnPowerResult(0, "KMST20001") Then Call KS.ReturnErr(1, "")   '���ػ�����������Ȩ�޼��
			
			SQLStr = "Select * From KS_DownParam Where ChannelID=" & ChannelID
			Set RS = Server.CreateObject("Adodb.RecordSet")
			If Action = "save" Then
			  RS.Open SQLStr, conn, 1, 3
			  If RS.Eof Then
			   RS.AddNew
			   RS("ChannelID")=ChannelID
			  End If
			  RS("DownLb") = KS.G("DownLB")
			  RS("DownYY") = KS.G("DownYY")
			  RS("DownSQ") = KS.G("DownSQ")
			  RS("DownPT") = KS.G("DownPT")
			  RS.Update
			  .echo ("<script>alert('���ز����޸ĳɹ�!');</script>")
			  RS.Close
			End If
			 RS.Open SQLStr, conn, 1, 1
			  If Not RS.EOF Then
			   DownLb = RS("DownLb")
			   DownYY = RS("DownYy")
			   DownSQ = RS("DownSQ")
			   DownPT = RS("DownPT")
			  End If
			RS.Close
			
			Set RS = Nothing
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.echo "      <div class='topdashed sort'>"
			.echo "      ���ز�������"
			.echo "      </div>"
			
			.echo "<br /><strong>��ģ������:</strong><select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.echo " <option value='0'>---��ѡ��ģ��---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks6=3]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			    .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.echo "</select>"

			.echo "<form action=""?ChannelID=" & ChannelID &"&Action=save"" method=""post"" name=""DownParamForm"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" colspan=""4"" class='clefttitle'>&nbsp;<font color=""#000080""><b>��������Զ�</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" height=""200"" align=""center"" class='tdbg'>�趨���<br>"
			.echo "        <textarea name=""DownLb"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownLb & "</textarea>"
			.echo "        <br>"
			.echo "        <span class=""STYLE1"">˵����ÿһ�����Ϊһ��</span><br></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>�趨���ԣ�<br>"
			.echo "      <textarea name=""DownYy"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownYY & "</textarea>"
			.echo "      <br>"
			.echo "      <span class=""STYLE1"">˵����ÿһ������Ϊһ��</span></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>��Ȩ��ʽ�� <br>"
			.echo "      <textarea name=""DownSq"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownSQ & "</textarea>"
			.echo "        <br>"
			.echo "        <span class=""STYLE1"">˵����ÿһ����Ȩ��ʽΪһ��</span></td>"
			.echo "      <td width=""25%"" align=""center"" class='tdbg'>����ƽ̨��<br>"
			.echo "      <textarea name=""DownPt"" cols=""20"" rows=""10"" style=""border-style: solid; border-width: 1"">" & DownPT & "</textarea>"
			.echo "      <br>"
			.echo "      <span class=""STYLE1"">˵����ÿһ������ƽ̨Ϊһ��</span></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "</form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language=""javascript"">"
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.DownParamForm;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "    return true;" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->"
			.echo "</Script>"
			End With
		End Sub

End Class
%> 

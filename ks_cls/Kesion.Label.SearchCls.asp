<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'��Ѵ��վ����ϵͳ,ͨ��ˢ����
'����:������ �汾 V 6.0
'-----------------------------------------------------------------------------------------------
Dim ShCls:Set ShCls=New RefreshSearchCls
Class RefreshSearchCls
		Private KS  
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set ShCls=Nothing
		End Sub
		
		'�滻��վ����������
		Function Run(byVal tag)
		 tag=Lcase(tag)
		 if tag="getsearchbydate" then
		   Run=GetSearchByDate()
		 elseif tag="getsearch" then
		   Run=GetSearch()
		 else
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.DocumentElement.SelectNodes("channel")
					if tag=lcase("get" & Node.SelectSingleNode("@ks10").text & "search") then
					  run="<script src=""" & KS.Setting(3) & KS.Setting(93) & "S_" & Node.SelectSingleNode("@ks10").text & ".js""></script>"
					end if
				Next
		 end if
		End Function
		
		'ȡ�ø߼���������
		Function GetSearchByDate()
		 GetSearchByDate="<iframe id=gToday:normal:agenda.js style=""BORDER-RIGHT: 0px ridge; BORDER-TOP: 0px ridge; BORDER-LEFT: 0px ridge; BORDER-BOTTOM: 0px ridge"" name=gToday:normal:agenda.js src=""" & KS.Setting(3) & "KS_Inc/iflateng.htm?../plus/search/?m=1&stype=100"" frameBorder=0 width=160 scrolling=no height=170></iframe>"
		End Function
		'ȡ��������
		Function GetSearch()
			   GetSearch = "<form id=""SearchForm"" name=""SearchForm"" method=""Get"" action=""" & KS.Setting(3) &"plus/search/"">" & vbCrLf
			   GetSearch = GetSearch & "<div class=""searchsd"">" & vbCrLf
			   GetSearch = GetSearch & "������ <input name=""key"" type=""text"" class=""textbox"" value=""������ؼ���"" onfocus=""this.select();""/><span>" & vbCrLf
			   GetSearch = GetSearch & "<select style=""width:120px"" name=""m"">" & vbCrLf
			   GetSearch = GetSearch & "<option value=""0"">ȫ��</option>" & vbCrLf 
			   If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.DocumentElement.SelectNodes("channel")
			     if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				 GetSearch = GetSearch & "<option value=""" &Node.SelectSingleNode("@ks0").text & """>" & Node.SelectSingleNode("@ks3").text & "</option>" & vbCrLf
				 End If
				Next

			   GetSearch = GetSearch & "</select>" & vbCrLf 
			   GetSearch = GetSearch & "<input type=""submit"" class=""inputButton"" name=""Submit1"" value=""վ������"" />" & vbCrLf
			   GetSearch = GetSearch & "</span>" & vbCrLf
			   GetSearch = GetSearch & "</div>" & vbCrLf
			   GetSearch = GetSearch & "</form>" & vbCrLf
		End Function

End Class
%> 
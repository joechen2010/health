<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
<!--#include file="config.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,c_str,curr_tips,pid,ads_str,s_str,ID,ClassName,S,showStr
		Private TotalPut,MaxPerPage,CurrentPage,Key
		Private url,spaceurl,msgurl,contacturl,node,logo
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  MaxPerPage=10
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If KS.S("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
			Else
			  CurrentPage = 1
			End If
			Key=KS.CheckXSS(KS.S("Key"))
			Pid=KS.ChkClng(KS.S("id"))
			If Pid=0 and Key="" Then KS.Die "������!û��ָ����Ŀ!"
            KS.LoadClassConfig()
			if Key="" then
				Dim Node,Xml:Set Xml=Application(KS.SiteSN&"_class")
				Set Node=Xml.DocumentElement.SelectSingleNode("class[@ks9=" & pid & "]")
				If Node Is Nothing Then ks.die "������,�Ƿ�����!"
				ID=Node.SelectSingleNode("@ks0").text
				ClassName=Node.SelectSingleNode("@ks1").text
			else
			    ClassName=Key
		    end if
			

		           Dim Template
				   Template = KSR.LoadTemplate(KS.Setting(3) & KS.Setting(90) & "��ҵ�ռ�/product_list.html")
				   FCls.RefreshType = "enterpriseprolist" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = ID '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   Fcls.Locationstr=className
				   Template=Replace(Template,"{$ShowClassName}",className)
				   Template=Replace(Template,"{$ShowClassID}",pid)
				   call getcategory()
				   Template=Replace(Template,"{$ShowSmallClass}",str)
				   call GetProductList()
				   call getsearchlist()
				   Template=Replace(Template,"{$ShowProductList}",c_str)
				   Template=Replace(Template,"{$ShowSearch}",s_str)
				   Template=Replace(Template,"{$ShowPage}",KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false))
				   Template=KSR.KSLabelReplaceAll(Template)
		 Response.Write Template  
		End Sub
		
		Sub getcategory()
		 Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr
		 Str=""
		 For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 and @ks12=5 and @ks10>=2]")
		      SpaceStr=""
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>2 Then
				Str = Str & "<a class=""item"" href='?id=" & Node.SelectSingleNode("@ks9").text & "'>" & Node.SelectSingleNode("@ks1").text & " </a>"
			  Else
				Str = Str & "<div class=""category""><a href='?id=" & Node.SelectSingleNode("@ks9").text & "'>" & Node.SelectSingleNode("@ks1").text & "</a></div>"
			  End If
		Next
		End Sub
		
		
		Sub GetSearchList()
		  s_str="<form action='?' name='psform' method='get'>"
		  s_str=s_str & "��Ʒ������<input type='text' name='key' size='30'>"
		  s_str=s_str & "&nbsp;<select name='t'><option value='0'>��ʾ���в�Ʒ</option><option value='1'>��ʾ��������</option><option value='3'>��ʾ���3��</option><option value='5'>��ʾ���5��</option><option value='7'>��ʾ���7��</option><option value='15'>��ʾ���15��</option><option value='30'>��ʾ���30��</option><option value='90'>��ʾ���������</option><option value='180'>��ʾ�������</option></select>"
		  s_str=s_str & "&nbsp;<select name='pid'>"
		  
		  
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
		KS.LoadClassConfig()
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1 and @ks12=5]")
		       SpaceStr=""
			   If trim(pid)=trim(Node.SelectSingleNode("@ks9").text) then pstr=" selected" else pstr=""
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "����"
				 Next
				s_str=s_str & "<option value='" & Node.SelectSingleNode("@ks9").text & "'" &pstr &">" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
			  Else
				s_str=s_str & "<option value='" & Node.SelectSingleNode("@ks9").text & "'" &pstr &">" & Node.SelectSingleNode("@ks1").text & " </option>"
			  End If
		Next
		  
		  
		  
		  
		  s_str=s_str & "</select>&nbsp;<input onclick=""if(document.psform.key.value==''){alert('������ؼ���!');document.psform.key.focus();return false;}"" type='image' src='../images/vblog_lmy_002.gif' align='absmiddle'>"
		  s_str=s_str & "</form>"
		End Sub
		
		
		
		Sub GetProductList()
		  c_str="<img src='../images/order_arrow.gif' align='absmiddle'><a href='?"&KS.QueryParam("page,popular,recommend") & "'>Ĭ������</a> <img src='../images/order_arrow.gif' align='absmiddle'><a href='?recommend=1&"& KS.QueryParam("page,popular,recommend") & "'>�Ƽ���Ʒ</a> <img src='../images/order_arrow.gif' align='absmiddle'><a href='?popular=1&"&KS.QueryParam("page,popular,recommend") & "'>���Ų�Ʒ</a>"

		 Dim Param:Param=" where a.verific=1"
		 If Key<>"" Then 
		  Param=Param & " and a.title like '%" & Key & "%'"
		 Else
		  Param=Param & " and tid in(select id from ks_class where ts like '%" & id & "%')"
		 End If
		 If KS.S("Recommend")="1" Then Param =Param & " and a.recommend=1"
		 If KS.S("Popular")="1" Then Param=Param & " and a.popular=1"
		 
		 If KS.ChkClng(KS.S("T"))<>0 Then
			  Param=Param & " and datediff("& DataPart_D&",a.AddDate," &SqlNowString & ")<" & KS.ChkClng(KS.S("T"))
		 End If
		 Dim RS,SqlStr,OrderStr,XML,Node
		 OrderStr=" order by a.istop desc,a.id desc"
		 SqlStr="select b.CompanyName,b.province,b.city,a.inputer,a.id,a.price_member,a.price_market,a.price,a.title,a.tid,a.prointro,a.PhotoUrl,a.recommend,a.popular,a.promodel,a.rank,a.adddate from KS_Product a left join ks_enterprise b on a.inputer=b.username "&param& OrderStr
		 Set RS=Server.CreateObject("adodb.recordset")
		 rs.open SqlStr,conn,1,1
		 IF RS.Eof And RS.Bof Then
			  totalput=0
			  exit sub
		  Else
							TotalPut= Conn.Execute("Select count(*) from KS_Product a " & Param)(0)
							If CurrentPage < 1 Then CurrentPage = 1
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","")
							If IsObject(XML) Then
							  	 Call ShowByList(Xml)
						    End If
			End IF
			
			
			RS.Close
			Set RS=Nothing
		End Sub
		
		Sub GetUrl()
		  If KS.SSetting(21)="1" Then 
		  SpaceUrl="../space/" & node.selectsinglenode("@inputer").text & KS.SSetting(22) 
		  url="../space/show-product-" & node.selectsinglenode("@inputer").text & "-" & node.selectsinglenode("@id").text & KS.SSetting(22)
		  msgUrl="../space/message-" & node.selectsinglenode("@inputer").text & KS.SSetting(22) 
		  contacturl="../space/info-" & node.selectsinglenode("@inputer").text & KS.SSetting(22) 
		 Else 
		  SpaceUrl="../space/?" & node.selectsinglenode("@inputer").text
		  url="../space/?" & node.selectsinglenode("@inputer").text & "/showproduct/" & node.selectsinglenode("@id").text
		  msgUrl="../space/?" & node.selectsinglenode("@inputer").text & "/message"
		  contacturl="../space/?" & node.selectsinglenode("@inputer").text & "/info"
		 End If
		End Sub
		
		
		Sub ShowByList(Xml)
		 Dim I,n
		 c_str=c_str & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">" & vbcrlf
         c_str=c_str & "<tr bgcolor=""#E7E7E7"">"
         c_str=c_str & "<td width=""111"" height=""26"" align=""center"">��ƷͼƬ</td>"
         c_str=c_str & "<td width=""200"" align=""center"">��Ʒ/��˾</td>"
         c_str=c_str & "<td width=""85"" align=""center"">�����ص�</td>"
         c_str=c_str & "<td align=""center"">����ѯ��</div></td>"
         c_str=c_str & "</tr>"
		 For Each Node In XML.DocumentElement.SelectNodes("row")
		 logo=trim(Node.SelectSingleNode("@photourl").text)
		 if KS.isnul(logo) then 
		  logo="/images/nophoto.gif"
		 end if
		 dim str:str=""
		 if node.selectsinglenode("@recommend").text="1" then str="<font color=green>��</font>"
		 if node.selectsinglenode("@popular").text="1" then str= str & " <font color=red>��</font>"
		
		 GetUrl
         n=n+1
		 if n mod 2=0 then
		 c_str=c_str & "<tr bgcolor=""#f6f6f6"">"
		 else
         c_str=c_str & "<tr>"
		 end if
         c_str=c_str & "<td height=""125"" align=""center""><a class=""pic"" href='" & url & "' target='_blank'><img src=""" & logo & """ width=90 height=80 border='0'/></a><br/>" & formatdatetime(Node.SelectSingleNode("@adddate").text,2) & "</td>"
         c_str=c_str & "<td  valign='top' style=""padding:5px;WORD-BREAK: break-all""><a href=""" & url & """ target=""_blank"" class='productname'>" & node.selectsinglenode("@title").text &"</a> " & str & "<br/><span class='attribute'>���:" & KS.C_C(Node.SelectSingleNode("@tid").text,1) & " ��Ʒ�ͺ�:" &Node.SelectSinglenode("@promodel").text & " | ��Ʒ�ȼ�: " & Node.SelectSingleNode("@rank").text & " <br/>�ο��۸�:��" & Node.SelectSinglenode("@price_market").text & " | �Żݼ۸�:��" & Node.SelectSingleNode("@price_member").text & "<br/>����:" & KS.Gottopic(KS.LoseHtml(KS.HtmlCode(node.selectsinglenode("@prointro").text)),120) & "...</span>"
		 If Not KS.IsNul(node.selectsinglenode("@companyname").text) Then
		 c_str=c_str & "<br/><b>��˾����:</b><a href='" & SpaceUrl & "' target='_blank'>" & node.selectsinglenode("@companyname").text  &"</a> (<a href='" & contacturl & "' target='_blank'>�鿴�ù�˾��ϵ��ʽ</a>)"
		 End If
		 c_str=c_str &"</td>"
         c_str=c_str & "<td align=""center"">" & node.selectsinglenode("@province").text & node.selectsinglenode("@city").text & " </td>"
         c_str=c_str & "<td align=""center""><a href='" & msgUrl &"' target='_blank'>����ѯ��</a><br/><br/><a href='../User/index.asp?User_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & node.selectsinglenode("@id").text & "' target='_blank'>�ղ�</a> <a href='../plus/Comment.asp?ChannelID=5&InfoID=" & node.selectsinglenode("@id").text & "' target='_blank'>����</a></td>"
         c_str=c_str & "</tr>"
		 I=I+1
		 Next
         c_str=c_str & "</table>"
		End Sub
		
End Class
%>

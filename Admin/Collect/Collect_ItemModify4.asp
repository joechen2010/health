<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Collect_ItemModify4
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemModify4
        Private KS
		Private KMCObj
		Private ConnItem,ThumbType,TbsString,TboString
		Private RsItem, SqlItem, FoundErr, ErrMsg, Action, ItemID
		Private LoginType, LoginUrl, LoginPostUrl, LoginUser, LoginPass, LoginFalse, LoginResult, LoginData
		Private ListStr, LsString, LoString, ListPageType, LPsString, LPoString, ListPageStr1, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3, HsString, HoString, HttpUrlType, HttpUrlStr
		Private TsString, ToString, CsString, CoString, DateType, DsString, DoString, AuthorType, AsString, AoString, AuthorStr, CopyFromType, FsString, FoString, CopyFromStr, KeyType, KsString, KoString, KeyStr, NewsPageType, NPsString, NPoString, NewsPageStr, NewsPageEnd,CharsetCode
		Private ListUrl, ListCode, NewsArrayCode, NewsArray, UrlTest, NewsCode,ThumbArrayCode,ThumbArray
		Private Testi,ChannelID,Tp_Lists,Tp_Listo,Tp_Srcs,Tp_Srco,Tp_Is,Tp_Io
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
		Action = Trim(Request("Action"))
		ItemID = Trim(Request("ItemID"))
		FoundErr = False
		
		If ItemID = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "�����������ĿID����Ϊ��\n"
		Else
		   ItemID = CLng(ItemID)
		End If
		
		If Action = "SaveEdit" And FoundErr <> True Then
		   Call SaveEdit
		End If
		
		If FoundErr <> True Then
		   Call GetTest
		End If
		
		If FoundErr = True Then
		   Call KS.AlertHistory(ErrMsg,-1)
		Else
		   Call Main
		End If
		End Sub
		
		Sub Main()
		Dim TitleStr,BodyStr,TempStr
		if KS.C_S(ChannelID,6)<4 then
		TitleStr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")(0)
		BodyStr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")(9)
		Else
		 TitleStr="����" : BodyStr="����"
		end if
		  
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>�ɼ�ϵͳ</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "<div class='topdashed'>"& KMCObj.GetItemLocation(4,ItemID) &"</div>"
		Response.Write "<form method=""post"" action=""Collect_ItemModify5.asp"" name=""form1"">"
		Response.Write "<input type='hidden' name='ChannelID' value='" & ChannelID & "'>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" class=""ctable"">"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""20%"" align=""center"" class='clefttitle'>" & TitleStr &"��ʼ��ǣ�<p>��</p>"
		Response.Write "      " & TitleStr & "������ǣ�</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""TsString"" cols=""49"" rows=""3"">" & TsString & "</textarea><br>"
		 Response.Write "     <textarea name=""ToString"" cols=""49"" rows=""3"">" & ToString & "</textarea></td>"
		 Response.Write "    </tr>"
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>" & BodyStr & "��ʼ��ǣ�<br/>" & TempStr & "<br/>"
		 Response.Write "    " & BodyStr &"������ǣ�<br/>" & TempStr & "</td>"
		 Response.Write "    <td width=""75%"">"
		 Response.Write "     <textarea name=""CsString"" cols=""49"" rows=""3"">" & CsString & "</textarea><br>"
		 Response.Write "    <textarea name=""CoString"" cols=""49"" rows=""3"">" & CoString & "</textarea></td>"
		 Response.Write "   </tr>"
		 
		  Dim RSM,Xml,Node
		 If KS.C_S(ChannelID,6)="2" Then 'ͼƬģ��
				  Set RSM=connItem.Execute("Select * From KS_FieldRules Where ChannelID=" & ChannelID &" And ItemID=" & ItemID & " And FieldID=0 order by id")
				  If Not RSM.Eof Then
					Set Xml=KS.RstoXml(RSM,"row","")
				  End If
				  RSM.Close
				  Set RSM=Nothing
				  If IsObject(xml) Then
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_List']/@beginstr")
				   if not Node Is Nothing Then
					Tp_lists=Node.Text
				   End If
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_List']/@endstr")
				   if not Node Is Nothing Then
					Tp_listo=Node.Text
				   End If
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_Src']/@beginstr")
				   if not Node Is Nothing Then
					Tp_Srcs=Node.Text
				   End If
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_Src']/@endstr")
				   if not Node Is Nothing Then
					Tp_Srco=Node.Text
				   End If
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_I']/@beginstr")
				   if not Node Is Nothing Then
					Tp_Is=Node.Text
				   End If
				   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='Tp_I']/@endstr")
				   if not Node Is Nothing Then
					Tp_Io=Node.Text
				   End If
				   
				  End If
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'><b>ͼƬ��ַ�б�ʼ��ǣ�</b></td>"
				 Response.Write "    <td width=""75%""><textarea name=""Tp_listBeginStr"" cols=""49"" rows=""3"">" & Tp_lists & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'><b>ͼƬ��ַ�б������ǣ�</b></td>"
				 Response.Write "    <td width=""75%""><textarea name=""Tp_listEndStr"" cols=""49"" rows=""3"">" & Tp_listo & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>����ͼƬ���ã�</td>"
				 Response.Write "    <td width=""75%"">"
				 Response.Write "     <table border='0' width='100%'>"
				 Response.Write "       <tr><td><font color=blue>����ͼƬ��ַ��ʼ���</font></td><td><textarea name=""Tp_srcBeginStr"" cols=""49"" rows=""3"">" & Tp_srcs & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>����ͼƬ��ַ�������</font></td><td><textarea name=""Tp_srcEndStr"" cols=""49"" rows=""3"">" & Tp_srco & "</textarea></td></tr>"
				 Response.Write "       <tr><td><font color=blue>����ͼƬ���ܿ�ʼ���</font></td><td><textarea name=""Tp_iBeginStr"" cols=""49"" rows=""3"">" & Tp_is & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>����ͼƬ���ܽ������</font></td><td><textarea name=""Tp_iEndStr"" cols=""49"" rows=""3"">" & Tp_io & "</textarea></td></tr>"
				 Response.Write "    </table>"
				 Response.Write "   </tr>"
		 ElseIf KS.C_S(ChannelID,6)="5" Then '�̳�ģ��
		         Dim Shop_BigPhotoBeginStr,Shop_BigPhotoEndStr,Shop_BigPhotoSrcBeginStr,Shop_BigPhotoSrcEndStr,Shop_UnitBeginStr,Shop_UnitEndStr,Shop_OriginPriceBeginStr,Shop_OriginPriceEndStr,Shop_PriceBeginStr,Shop_PriceEndStr,Shop_MarketPriceBeginStr,Shop_MarketPriceEndStr,Shop_MemberPriceBeginStr,Shop_MemberPriceEndStr,Shop_ProModelBeginStr,Shop_ProModelEndStr,Shop_ProSpecificatBeginStr,Shop_ProSpecificatEndStr,Shop_ProducerNameBeginStr,Shop_ProducerNameEndStr,Shop_TrademarkNameBeginStr,Shop_TrademarkNameEndStr
				Dim FieldNameList,FieldNameArr
				 FieldNameList="Shop_BigPhoto,Shop_BigPhotoSrc,Shop_Unit,Shop_OriginPrice,Shop_Price,Shop_MarketPrice,Shop_MemberPrice,Shop_ProModel,Shop_ProSpecificat,Shop_ProducerName,Shop_TrademarkName"
				 FieldNameArr=Split(FieldNameList,",")
				 Set RSM=connItem.Execute("Select * From KS_FieldRules Where ChannelID=" & ChannelID &" And ItemID=" & ItemID & " And FieldID=0 order by id")
				  If Not RSM.Eof Then
					Set Xml=KS.RstoXml(RSM,"row","")
				  End If
				  RSM.Close
				  Set RSM=Nothing
				  If IsObject(xml) Then
				    For I=0 To Ubound(FieldNameArr)
					   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='" & FieldNameArr(i) & "']/@beginstr")
					   if not Node Is Nothing Then
						Execute(FieldNameArr(i)&"beginstr=Node.Text")
					   End If
					   Set Node=XML.DocumentElement.SelectSingleNode("row[@fieldname='" & FieldNameArr(i) & "']/@endstr")
					   if not Node Is Nothing Then
						Execute(FieldNameArr(i)&"endstr=Node.Text")
					   End If
					
					Next
				  End If
				 
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ��ͼ���ã�</td>"
				 Response.Write "    <td width=""75%"">"
				 Response.Write "     <table border='0' width='100%'>"
				 Response.Write "       <tr><td><font color=blue>��Ʒ��ͼ���뿪ʼ���</font></td><td><textarea name=""Shop_BigPhotoBeginStr"" cols=""49"" rows=""2"">" & Shop_BigPhotoBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>��Ʒ��ͼ����������</font></td><td><textarea name=""Shop_BigPhotoEndStr"" cols=""49"" rows=""2"">" & Shop_BigPhotoEndStr & "</textarea></td></tr>"
				 Response.Write "       <tr><td><font color=blue>��Ʒ��ͼSrc��ʼ���</font></td><td><textarea name=""Shop_BigPhotoSrcBeginStr"" cols=""49"" rows=""3"">" & Shop_BigPhotoSrcBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>��Ʒ��ͼSrc�������</font></td><td><textarea name=""Shop_BigPhotoSrcEndStr"" cols=""49"" rows=""3"">" & Shop_BigPhotoSrcEndStr & "</textarea></td></tr>"
				 Response.Write "    </table>"
				 Response.Write "   </tr>"
				 
			
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>������λ��ʼ��ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_UnitBeginStr"" cols=""49"" rows=""2"">" & Shop_UnitBeginStr & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>������λ������ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_UnitEndStr"" cols=""49"" rows=""2"">" & Shop_UnitEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>�۸����ã�</td>"
				 Response.Write "    <td width=""75%"">"
				 Response.Write "     <table border='0' width='100%'>"
				 Response.Write "       <tr><td><font color=blue>��Ա�ۿ�ʼ���</font></td><td><textarea name=""Shop_MemberPriceBeginStr"" cols=""49"" rows=""2"">" & Shop_MemberPriceBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>��Ա�۽������</font></td><td><textarea name=""Shop_MemberPriceEndStr"" cols=""49"" rows=""2"">" & Shop_MemberPriceEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td></tr>"
				 Response.Write "       <tr><td><font color=blue>ԭʼ���ۼۿ�ʼ���</font></td><td><textarea name=""Shop_OriginPriceBeginStr"" cols=""49"" rows=""2"">" & Shop_OriginPriceBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>ԭʼ���ۼ۽������</font></td><td><textarea name=""Shop_OriginPriceEndStr"" cols=""49"" rows=""2"">" & Shop_OriginPriceEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td></tr>"
				 Response.Write "       <tr><td><font color=blue>��ǰ���ۼۿ�ʼ���</font></td><td><textarea name=""Shop_PriceBeginStr"" cols=""49"" rows=""2"">" & Shop_PriceBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>��ǰ���ۼ۽������</font></td><td><textarea name=""Shop_PriceEndStr"" cols=""49"" rows=""2"">" & Shop_PriceEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td></tr>"
				 Response.Write "       <tr><td><font color=blue>�г��ۿ�ʼ���</font></td><td><textarea name=""Shop_MarketPriceBeginStr"" cols=""49"" rows=""2"">" & Shop_MarketPriceBeginStr & "</textarea></td><tr>"
				 Response.Write "       <tr><td><font color=blue>�г��۽������</font></td><td><textarea name=""Shop_MarketPriceEndStr"" cols=""49"" rows=""2"">" & Shop_MarketPriceEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td></tr>"
				 Response.Write "    </table>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ�ͺſ�ʼ��ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProModelBeginStr"" cols=""49"" rows=""2"">" & Shop_ProModelBeginStr & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ�ͺŽ�����ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProModelEndStr"" cols=""49"" rows=""2"">" & Shop_ProModelEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ���ʼ��ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProSpecificatBeginStr"" cols=""49"" rows=""2"">" & Shop_ProSpecificatBeginStr & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ��������ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProSpecificatEndStr"" cols=""49"" rows=""2"">" & Shop_ProSpecificatEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>�����̿�ʼ��ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProducerNameBeginStr"" cols=""49"" rows=""2"">" & Shop_ProducerNameBeginStr & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>�����̽�����ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_ProducerNameEndStr"" cols=""49"" rows=""2"">" & Shop_ProducerNameEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ�̱꿪ʼ��ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_TrademarkNameBeginStr"" cols=""49"" rows=""2"">" & Shop_ProducerNameBeginStr & "</textarea></td>"
				 Response.Write "   </tr>"
				 Response.Write "   <tr class='tdbg'>"
				 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��Ʒ�̱������ǣ�</td>"
				 Response.Write "    <td width=""75%""><textarea name=""Shop_TrademarkNameEndStr"" cols=""49"" rows=""2"">" & Shop_ProducerNameEndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
				 Response.Write "   </tr>"
				 
		 End If
		 
		 Dim RS,SQL,I,BeginStr,EndStr
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select FieldID,FieldTitle,FieldName,BeginStr,EndStr From KS_FieldItem Where ShowType=0 and ChannelID=" &ChannelID & " order by orderid",ConnItem,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) Then
		   For I=0 To Ubound(SQL,2)
			 Response.Write "   <tr class='tdbg'>"
			 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>" & SQL(1,I) & "��ʼ��ǣ�<br/><br/>"
			 Response.Write "    " & SQL(1,I) &"������ǣ�<br/></td>"
			 Response.Write "    <td width=""75%"">"
			   Dim RSV:Set RSV=Server.CreateObject("ADODB.RECORDSET")
			   RSV.Open "Select BeginStr,EndStr From KS_FieldRules Where ItemID=" & ItemID & " And channelid=" & ChannelID & " and FieldName='" & SQL(2,I) &"'",ConnItem,1,1
			   If Not RSV.Eof Then
			     BeginStr=RSV(0)
				 EndStr=RSV(1)
			   Else
			     BeginStr=""
				 EndStr=""
			   End If
			   RSV.Close:Set RSV=Nothing
			 Response.Write "     <textarea name=""begin" & SQL(2,I) & """ cols=""49"" rows=""3"">" & BeginStr & "</textarea><br>"
			 Response.Write "    <textarea name=""end" & SQL(2,I) & """ cols=""49"" rows=""3"">" & EndStr & "</textarea><br/><font color=red>Tips:�������������ʱ,����ȡ��ʼ�����ΪĬ��ֵ.</font></td>"
			 Response.Write "   </tr>"
		   Next
		 End If
		 
		  Response.Write "   <tr class='tdbg'>"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'>ʱ&nbsp; ��&nbsp;"
		  Response.Write "      ��&nbsp; �ã�</td>"
		  Response.Write "    <td width=""75%"">"
		  Response.Write "      <input type=""radio"" value=""0"" name=""DateType"" "
		  If DateType = 0 Then Response.Write "checked"
		  Response.Write " onClick=""Date1.style.display='none'"">��������&nbsp;"
		  Response.Write "  <input type=""radio"" value=""1"" name=""DateType"" "
		  If DateType = 1 Then Response.Write "checked"
		  Response.Write " onClick=""Date1.style.display=''"">���ñ�ǩ&nbsp;    </tr>"
		  Response.Write "  <tr class='tdbg' id=""Date1"" style=""display:'"
		  If DateType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>ʱ�俪ʼ��ǣ�</font>"
		  Response.Write "      <p>��</p>"
		  Response.Write "      <font color=blue>ʱ�������ǣ�</font></td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""DsString"" cols=""49"" rows=""3"">" & DsString & "</textarea><br>"
		 Response.Write "     <textarea name=""DoString"" cols=""49"" rows=""3"">" & DoString & "</textarea></td>"
		Response.Write "    </tr>"
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>��&nbsp; ��&nbsp;"
		 Response.Write "       ��&nbsp; �ã�</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "       <input type=""radio"" value=""0"" name=""AuthorType"" "
		 If AuthorType = 0 Then Response.Write "checked"
		 Response.Write "  onClick=""Author1.style.display='none';Author2.style.display='none'"">��������&nbsp;"
		  Response.Write "      <input type=""radio"" value=""1"" name=""AuthorType"" "
		  If AuthorType = 1 Then Response.Write "checked"
		  Response.Write " onClick=""Author1.style.display='';Author2.style.display='none'"">���ñ�ǩ&nbsp;"
		  Response.Write "      <input type=""radio"" value=""2"" name=""AuthorType"" "
		  If AuthorType = 2 Then Response.Write "checked"
		  Response.Write " onClick=""Author1.style.display='none';Author2.style.display=''"">ָ������</td>"
		 Response.Write "   </tr>"
		  Response.Write "  <tr class='tdbg' id=""Author1"" style=""display:'"
		  If AuthorType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>���߿�ʼ��ǣ�</font>"
		  Response.Write "      <p>��</p>"
		   Response.Write "     <font color=blue>���߽�����ǣ�</font></td>"
		   Response.Write "   <td width=""75%"">"
		   Response.Write "   <textarea name=""AsString"" cols=""49"" rows=""3"">" & AsString & "</textarea><br>"
		  Response.Write "    <textarea name=""AoString"" cols=""49"" rows=""3"">" & AoString & "</textarea></td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr  class='tdbg' id=""Author2"" style=""display:'"
		  If AuthorType <> 2 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>��ָ�����ߣ�</font></td>"
		  Response.Write "    <td width=""75%"">"
		  Response.Write "    <input name=""AuthorStr"" type=""text"" id=""AuthorStr"" value=""" & AuthorStr & """>      </td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'>"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'>��&nbsp; Դ&nbsp;"
		  Response.Write "      ��&nbsp; �ã�</td>"
		 Response.Write "     <td width=""75%"">"
		  Response.Write "      <input type=""radio"" value=""0"" name=""CopyFromType"" "
		  If CopyFromType = 0 Then Response.Write "checked"
		  Response.Write " onClick=""CopyFrom1.style.display='none';CopyFrom2.style.display='none'"">��������&nbsp;"
		  Response.Write "      <input type=""radio"" value=""1"" name=""CopyFromType"" "
		  If CopyFromType = 1 Then Response.Write "checked"
		  Response.Write " onClick=""CopyFrom1.style.display='';CopyFrom2.style.display='none'"">���ñ�ǩ&nbsp;"
		  Response.Write "      <input type=""radio"" value=""2"" name=""CopyFromType"" "
		  If CopyFromType = 2 Then Response.Write "checked"
		  Response.Write " onClick=""CopyFrom1.style.display='none';CopyFrom2.style.display=''"">ָ����Դ</td>"
		   Response.Write " </tr>"
		  Response.Write "  <tr  class='tdbg' id=""CopyFrom1"" style=""display:'"
		  If CopyFromType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "   <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>��Դ��ʼ��ǣ�</font>"
		  Response.Write "      <p>��</p>"
		  Response.Write "      <font color=blue>��Դ������ǣ�</font></td>"
		  Response.Write "    <td width=""75%"">"
		  Response.Write "    <textarea name=""FsString"" cols=""49"" rows=""3"">" & FsString & "</textarea><br>"
		  Response.Write "    <textarea name=""FoString"" cols=""49"" rows=""3"">" & FoString & "</textarea></td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg' id=""CopyFrom2"" style=""display:'"
		  If CopyFromType <> 2 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>��ָ����Դ��</font></td>"
		  Response.Write "    <td width=""75%"">"
		   Response.Write "   <input name=""CopyFromStr"" type=""text"" id=""CopyFromStr"" value=""" & CopyFromStr & """>      </td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'>"
		 Response.Write "     <td width=""20%"" align=""center""  class='clefttitle'>�ؼ��ִ����ã�</td>"
		  Response.Write "    <td width=""75%"">"
		 Response.Write "       <input type=""radio"" value=""0"" name=""KeyType"" "
		 If KeyType = 0 Then Response.Write "checked"
		 Response.Write " onClick=""Key1.style.display='none';Key2.style.display='none'"">��������&nbsp;"
		 Response.Write "       <input type=""radio"" value=""1"" name=""KeyType"" "
		 If KeyType = 1 Then Response.Write "checked"
		 Response.Write " onClick=""Key1.style.display='';Key2.style.display='none'"">��ǩ����&nbsp;"
		Response.Write "        <input type=""radio"" value=""2"" name=""KeyType"" "
		If KeyType = 2 Then Response.Write "checked"
		Response.Write " onClick=""Key1.style.display='none';Key2.style.display=''"">�Զ���ؼ���</td>"
		 Response.Write "   </tr>"
		Response.Write "    <tr class='tdbg' id=""Key1"" style=""display:'"
		If KeyType <> 1 Then Response.Write "none"
		Response.Write "'"">"
		Response.Write "      <td width=""20%"" align=""center""  class='clefttitle'><font color=blue>�ؼ��ʿ�ʼ��ǣ�</font>"
		Response.Write "        <p>��</p>"
		Response.Write "        <font color=blue>�ؼ��ʽ�����ǣ�</font></td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "     <textarea name=""KsString"" cols=""49"" rows=""3"">" & KsString & "</textarea><br>"
		  Response.Write "    <textarea name=""KoString"" cols=""49"" rows=""3"">" & KoString & "</textarea></td>"
		 Response.Write "   </tr>"
		  Response.Write "  <tr class='tdbg' id=""Key2"" style=""display:'"
		  If KeyType <> 2 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center"" class='clefttitle' ><font color=blue>��ָ���ؼ���</font></td>"
		  Response.Write "    <td width=""75%"">"
		 Response.Write "     <input name=""KeyStr"" type=""text"" id=""KeyStr"" value=""" & KeyStr & """>      </td>"
		 Response.Write "   </tr>"
		
		' If ChannelID<>1 Then
		'   Response.Write "<tr class='tdbg' style='display:none'>"
		' Else
		  Response.Write "   <tr class='tdbg'>"
		' End If
		 Response.Write "     <td width=""20%"" align=""center"" class='clefttitle'>"
		 If KS.C_S(ChannelID,6)="2" Then
		  Response.Write "ͼƬ��ҳ"
		 Else
		  Response.Write "���ķ�ҳ"
		 End If
		 Response.Write "���ã�</td>"
		 Response.Write "     <td width=""75%"">"
		 Response.Write "       <input type=""radio"" value=""0"" name=""NewsPageType"" "
		 If NewsPageType = 0 Then Response.Write "checked"
		 Response.Write " onClick=""NewsPage1.style.display='none';NewsPage12.style.display='none';NewsPage13.style.display='none';NewsPage2.style.display='none'"">��������&nbsp;"
		 Response.Write "       <input type=""radio"" value=""1"" name=""NewsPageType"" "
		 If NewsPageType = 1 Then Response.Write "checked"
		 Response.Write " onClick=""NewsPage1.style.display='';NewsPage12.style.display='';NewsPage13.style.display='';NewsPage2.style.display='none'"">Դ�����л�ȡ��ҳURL &nbsp;"
		 Response.Write "       <input style=""display:none"" type=""radio"" value=""2"" name=""NewsPageType"" "
		 If NewsPageType = 2 Then Response.Write "checked"
		 Response.Write " onClick=""NewsPage1.style.display='none';NewsPage12.style.display='none';NewsPage13.style.display='none';NewsPage2.style.display=''"">"
	'	 Response.Wreite "�ֶ�����"
		 Response.Write "   </td></tr>"
		  Response.Write "  <tr class='tdbg' id=""NewsPage1"" style=""display:'"
		  If NewsPageType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center"" class='clefttitle'><font color=blue>��ҳ���뿪ʼ��</font>"
		  Response.Write "      <p>��</p>"
		  Response.Write "      <font color=blue>��ҳ���������</font></td>"
		  Response.Write "    <td width=""75%"">"
		   Response.Write "     <textarea name=""NPsString"" cols=""49"" rows=""3"">" & server.htmlencode(NPsString) & "</textarea><br>"
		  Response.Write "      <textarea name=""NPoString"" cols=""49"" rows=""3"">" & server.htmlencode(NPoString) & "</textarea></td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'  id=""NewsPage12"" style=""display:'"
		  If NewsPageType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center"" class='clefttitle'><font color=""#0000FF"">��ҳURL��ʼ���룺</font></td>"
		  Response.Write "    <td width=""75%"">"
		  Response.Write "      <input name=""NewsPageStr"" type=""text"" size=""58"" value=""" & server.htmlencode(NewsPageStr) & """></td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg' id=""NewsPage13"" style=""display:'"
		  If NewsPageType <> 1 Then Response.Write "none"
		  Response.Write "'"">"
		  Response.Write "    <td width=""20%"" align=""center"" class='clefttitle'><font color=""#0000FF"">��ҳURL�������룺</font></td>"
		  Response.Write "    <td width=""75%"">"
		   Response.Write "     <input name=""NewsPageEnd"" type=""text"" size=""58"" value=""" & server.htmlencode(NewsPageEnd) & """></td>"
		   Response.Write " </tr>"
		
		   Response.Write " <tr class='tdbg'  id=""NewsPage2"" style=""display:'"
		   If NewsPageType <> 2 Then Response.Write "none"
		   Response.Write "'"">"
		   Response.Write "   <td width=""20%"" align=""center""><font color=blue>��&nbsp; ��&nbsp; ��&nbsp; �ã�</font></td>"
		   Response.Write "   <td width=""75%"">"
		   Response.Write "     <input name=""NewsPageStr2"" type=""text"" value=""Ԥ������"" size=""58"">      </td>"
		   Response.Write " </tr>"
		
		   Response.Write " <tr class='tdbg'>"
		   Response.Write "   <td height=""30"" colspan=""2"" align=""center""><br>"
		   Response.Write "     <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveEdit"">"
		   Response.Write "     <input name=""ItemID"" type=""hidden"" id=""ItemID"" value=""" & ItemID & """>"
		   Response.Write "     <input  type=""button"" class='button' name=""button1"" value=""��&nbsp;һ&nbsp;��"" onClick=""window.location.href='javascript:history.go(-1)'""  >"
		   Response.Write "     &nbsp;&nbsp;&nbsp;&nbsp;"
			Response.Write "  <input  type=""submit"" class='button' name=""Submit"" value=""��&nbsp;һ&nbsp;��""></td>"
			Response.Write "    <input type=""hidden"" name=""UrlTest"" id=""UrlTest"" value=""" & UrlTest & """ >"
			Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</form>"
		Response.Write "<br>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""ctable"" >"
		Response.Write "  <tr>"
		 Response.Write "   <td height=""22"" colspan=""2"" class=""sort""><div align=""center""><strong>�� �� �� �� �� �� �� ��</strong></div></td>"
		Response.Write "  </tr>"
		Response.Write "  <tr>"
		 Response.Write "   <td height=""30"" colspan=""2"">"
		 Response.Write "<font color=red>�б����ӵ�ַ:</font>"
		 Response.Write "<select name=""link1"" onchange=""window.open(this.value);"">"	
			For Testi = 0 To UBound(NewsArray)
				Response.Write "<option value='" & NewsArray(Testi) & "'>" & NewsArray(Testi) & "</option>"
			Next
		 Response.Write "</select>"
         
		 If IsArray(ThumbArray) Then
			 response.write "<br><br>"
			 Response.Write "<font color=red>�б�����ͼ��ַ:</font><select name=""link2"" onchange=""window.open(this.value);"">"
			For Testi = 0 To UBound(ThumbArray)
			   Response.Write "<option value='" & ThumbArray(Testi) & "'>" & ThumbArray(Testi) & "</option>"
			Next
			  Response.Write "</select>" 
		End If
		
		  Response.Write "      <br></td>"
		  Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		Sub SaveEdit()
		   HsString = Request.Form("HsString")
		   HoString = Request.Form("HoString")
		   HttpUrlType = Trim(Request.Form("HttpUrlType"))
		   HttpUrlStr = Trim(Request.Form("HttpUrlStr"))
		   ThumbType=Request.Form("ThumbType")
		   TbsString=Request.Form("TbsString")
		   TboString=Request.Form("TboString")
		
		   If HsString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "�����ӿ�ʼ��ǲ���Ϊ��\n"
		   End If
		   If HoString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "�����ӽ�����ǲ���Ϊ��\n"
		   End If
		   
		   If ThumbType=1 Then
		        If TbsString="" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "���б�����ͼ��ʼ��ǲ���Ϊ��\n"
				End If
				If TboString="" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "���б�����ͼ������ǲ���Ϊ��\n"
				End If
		   End If
		   
		   If HttpUrlType = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "����ѡ�����Ӵ�������\n"
		   Else
			  HttpUrlType = CLng(HttpUrlType)
			  If HttpUrlType = 1 Then
				 If HttpUrlStr = "" Then
					FoundErr = True
					ErrMsg = ErrMsg & "�������þ������ӵ�ַ\n"
				 Else
					If Len(HttpUrlStr) < 15 Then
					   FoundErr = True
					   ErrMsg = ErrMsg & "��������ӵ�ַ���ò���ȷ(����15���ַ�)\n"
					End If
				 End If
			  End If
		   End If
		
		   If FoundErr <> True Then
			  SqlItem = "Select ItemID,HsString,HoString,HttpUrlType,ThumbType,TbsString,TboString,HttpUrlStr,ChannelID From KS_CollectItem Where ItemID=" & ItemID
			  Set RsItem = Server.CreateObject("adodb.recordset")
			  RsItem.Open SqlItem, ConnItem, 2, 3
			  RsItem("HsString") = HsString
			  RsItem("HoString") = HoString
			  RsItem("HttpUrlType") = HttpUrlType
			  RsItem("ThumbType")=ThumbType
			  If ThumbType=1 Then
			  RsItem("TbsString")=TbsString
			  RsItem("TboString")=TboString
			  End If
			  If HttpUrlType = 1 Then
				 RsItem("HttpUrlStr") = HttpUrlStr
			  End If
			  RsItem.Update
			  
			      Dim RS,SQL,I,RSV
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select FieldTitle,FieldName,ChannelID,FieldID,OrderID,ShowType From KS_FieldItem Where ShowType=1 and ChannelID=" &RsItem("ChannelID") & " order by orderid",ConnItem,1,1
				 Do While Not RS.Eof 
				   Set RSV=Server.CreateObject("ADODB.RECORDSET")
				   RSV.Open "Select * From KS_FieldRules Where ItemID=" & ItemID & " and channelid=" & RS(2) & " and fieldid=" & rs(3),connItem,1,3
				   If RSV.Eof And RSV.Bof Then
				    RSV.AddNew
				   End If
				   RSV("ItemID")=ItemID
				   RSV("ChannelID")=rs(2)
				   RSV("FieldID")=rs(3)
				   RSV("FieldName")=rs(1)
				   RSV("OrderID")=rs(4)
				   RSV("BeginStr")=Request.Form("begin"&rs(1))
				   RSV("EndStr")=Request.Form("end"&rs(1))
				   RSV("ShowType")=rs(5)
				   RSV.Update
				   rsv.close
				   RS.MoveNext
				 Loop
				 RS.Close:Set RS=Nothing
			   RsItem.Close
			   Set RsItem = Nothing
		   End If
		End Sub
		
		Sub GetTest()
		   SqlItem = "Select * From KS_CollectItem Where ItemID=" & ItemID
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   RsItem.Open SqlItem, ConnItem, 1, 1
		   If RsItem.EOF And RsItem.BOF Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "�����������ĿID����Ϊ��\n"
		   Else
			  LoginType = RsItem("LoginType")
			  LoginUrl = RsItem("LoginUrl")
			  LoginPostUrl = RsItem("LoginPostUrl")
			  LoginUser = RsItem("LoginUser")
			  LoginPass = RsItem("LoginPass")
			  LoginFalse = RsItem("LoginFalse")
			  ListStr = RsItem("ListStr")
			  LsString = RsItem("LsString")
			  LoString = RsItem("LoString")
			  ListPageType = RsItem("ListPageType")
			  LPsString = RsItem("LPsString")
			  LPoString = RsItem("LPoString")
			  ListPageStr1 = RsItem("ListPageStr1")
			  ListPageStr2 = RsItem("ListPageStr2")
			  ListPageID1 = RsItem("ListPageID1")
			  ListPageID2 = RsItem("ListPageID2")
			  ListPageStr3 = RsItem("ListPageStr3")
			  HsString = RsItem("HsString")
			  HoString = RsItem("HoString")
			  HttpUrlType = RsItem("HttpUrlType")
			  HttpUrlStr = RsItem("HttpUrlStr")
			  TsString = RsItem("TsString")
			  ToString = RsItem("ToString")
			  CsString = RsItem("CsString")
			  CoString = RsItem("CoString")
			  
			  DateType = RsItem("DateType")
			  DsString = RsItem("DsString")
			  DoString = RsItem("DoString")
			  
			  AuthorType = RsItem("AuthorType")
			  AsString = RsItem("AsString")
			  AoString = RsItem("AoString")
			  AuthorStr = RsItem("AuthorStr")
		
			  CopyFromType = RsItem("CopyFromType")
			  FsString = RsItem("FsString")
			  FoString = RsItem("FoString")
			  CopyFromStr = RsItem("CopyFromStr")
		
			  KeyType = RsItem("KeyType")
			  KsString = RsItem("KsString")
			  KoString = RsItem("KoString")
			  KeyStr = RsItem("KeyStr")
		
			  NewsPageType = RsItem("NewsPageType")
			  NPsString = RsItem("NPsString")
			  NPoString = RsItem("NPoString")
			  NewsPageStr = RsItem("NewsPageStr")
			  NewsPageEnd = RsItem("NewsPageEnd")
			  
			  ChannelID=RsItem("ChannelID")
			  ThumbType=RsItem("ThumbType")
		      TbsString=RsItem("TbsString")
		      TboString=RsItem("TboString")
			  
			  CharsetCode=RsItem("CharsetCode")
			  
		   End If
		   RsItem.Close
		   Set RsItem = Nothing
		    if isnull(NewsPageEnd) then NewsPageEnd=""
			if isnull(NewsPageStr) then NewsPageStr=""
			if isnull(NPsString)   then NPsString=""
			if isnull(NPoString)   then NPoString=""
		
		   If LsString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "���б�ʼ��ǲ���Ϊ�գ�\n"
		   End If
		   If LoString = "" Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "���б������ǲ���Ϊ�գ�\n"
		   End If
		   
		   If ThumbType=1 Then
			   If TbsString = "" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "���б�����ͼ��ʼ��ǲ���Ϊ�գ�\n"
			   End If
			   If TboString = "" Then
				  FoundErr = True
				  ErrMsg = ErrMsg & "���б�����ͼ������ǲ���Ϊ�գ�\n"
			   End If
		   End If
		   
		   If ListPageType = 0 Or ListPageType = 1 Then
			  If ListStr = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���б�����ҳ����Ϊ�գ�\n"
			  End If
			  If ListPageType = 1 Then
				 If LPsString = "" Or LPoString = "" Then
					FoundErr = True
					ErrMsg = ErrMsg & "��������ҳ��ʼ��������ǲ���Ϊ�գ�\n"
				 End If
			  End If
			  If ListPageStr1 <> "" And Len(ListPageStr1) < 15 Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "��������ҳ�ض������ò���ȷ��\n"
					End If
		   ElseIf ListPageType = 2 Then
			  If ListPageStr2 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "����������ԭ�ַ�������Ϊ�գ�\n"
			  End If
			  If IsNumeric(ListPageID1) = False Or IsNumeric(ListPageID2) = False Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���������ɵķ�Χֻ�������֣�\n"
			  Else
				 ListPageID1 = CLng(ListPageID1)
				 ListPageID2 = CLng(ListPageID2)
				 If ListPageID1 = 0 And ListPageID2 = 0 Then
					FoundErr = True
					ErrMsg = ErrMsg & "���������ɵķ�Χ����ȷ��\n"
				 End If
			  End If
		   ElseIf ListPageType = 3 Then
			  If ListPageStr3 = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "��������ҳ����Ϊ�գ�\n"
			  End If
		   Else
			  FoundErr = True
			  ErrMsg = ErrMsg & "����ѡ�񷵻���һ������������ҳ����\n"
		   End If
		 
		   If LoginType = 1 Then
			  If LoginUrl = "" Or LoginPostUrl = "" Or LoginUser = "" Or LoginPass = "" Or LoginFalse = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���뽫��¼��Ϣ��д����\n"
			  End If
		   End If
		
		   If FoundErr <> True Then
			  Select Case ListPageType
			  Case 0, 1
					ListUrl = ListStr
			  Case 2
				 ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1))
			  Case 3
				 If InStr(ListPageStr3, "|") > 0 Then
					ListUrl = Left(ListPageStr3, InStr(ListPageStr3, "|") - 1)
				 Else
					ListUrl = ListPageStr3
				 End If
			  End Select
		   End If
		
			  If FoundErr <> True And Action <> "SaveEdit" And LoginType = 1 Then
			  LoginData = KMCObj.UrlEncoding(LoginUser & "&" & LoginPass)
			  LoginResult = KMCObj.PostHttpPage(LoginUrl, LoginPostUrl, LoginData)
			  If InStr(LoginResult, LoginFalse) > 0 Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���¼��վʱ����������ȷ�ϵ�¼��Ϣ����ȷ�ԣ�\n"
			  End If
			  End If
			  
		   If FoundErr <> True Then
			  ListCode = KMCObj.GetHttpPage(ListUrl,CharsetCode)
			  If ListCode <> "Error" Then
				 ListCode = KMCObj.GetBody(ListCode, LsString, LoString, False, False)
				 If ListCode = "Error" Then
					FoundErr = True
					ErrMsg = ErrMsg & "���ڽ�ȡ�б�ʱ��������\n"
				 End If
			  Else
				 FoundErr = True
				 ErrMsg = ErrMsg & "���ڻ�ȡ:" & ListUrl & "��ҳԴ��ʱ��������\n"
			  End If
		   End If
		
		   If FoundErr <> True Then
			  NewsArrayCode = KMCObj.GetArray(ListCode, HsString, HoString, False, False)
			  If NewsArrayCode = "Error" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���ڷ�����" & ListUrl & "�����б�ʱ��������\n"
			  Else
				 NewsArray = Split(NewsArrayCode, "$Array$")
				 If IsArray(NewsArray) = True Then
					For Testi = 0 To UBound(NewsArray)
					   If HttpUrlType = 1 Then
						  NewsArray(Testi) = Replace(HttpUrlStr, "{$ID}", NewsArray(Testi))
					   Else
						  NewsArray(Testi) = KMCObj.DefiniteUrl(NewsArray(Testi), ListUrl)
					   End If
					Next
					UrlTest = NewsArray(0)
					NewsCode = KMCObj.GetHttpPage(UrlTest,CharsetCode)
				 Else
					FoundErr = True
					ErrMsg = ErrMsg & "���ڷ�����" & ListUrl & "�����б�ʱ��������\n"
				 End If
			  End If
			  If ThumbType=1 Then
				  ThumbArrayCode = KMCObj.GetArray(ListCode, TbsString, TboString, False, False)
				  If ThumbArrayCode = "Error" Then
					 FoundErr = True
					 ErrMsg = ErrMsg & "���ڷ�����" & ListUrl & "�б������ͼʱ��������\n"
				  Else
					 ThumbArray = Split(ThumbArrayCode, "$Array$")
					 If IsArray(ThumbArray) = True Then
						For Testi = 0 To UBound(ThumbArray)
							 ThumbArray(Testi) = KMCObj.DefiniteUrl(ThumbArray(Testi), ListUrl)
						Next
					 Else
						FoundErr = True
						ErrMsg = ErrMsg & "���ڷ�����" & ListUrl & "�б������ͼʱ��������\n"
					 End If
				  End If
			  End If
			  
			  
			   '==============�б��Զ����ֶβɼ����Թ���Ϸ���================================
			     Dim RS,SQL,I
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select FieldID,FieldName,BeginStr,EndStr From KS_FieldRules Where ShowType=1 and ChannelID=" &ChannelID & " order by orderid",ConnItem,1,3
				 Do While Not RS.Eof
				   If rs(2)<>"" And rs(3)<>"" Then
				    Dim DiyField:DiyField = KMCObj.GetArray(ListCode, RS(2), RS(3), False, False)
				    Dim DiyFieldArr:DiyFieldArr = Split(DiyField, "$Array$")
					If Ubound(DiyFieldArr)<>Ubound(NewsArray) Then
						ErrMsg = ErrMsg & "���ڷ�����" & ListUrl & "�б���Զ����ֶ�[" & RS(1) & "]����ʱ��������\n"
						RS.Close:Set RS=Nothing
						FoundErr = True
					    Exit Sub
					End If
				   End If
				   RS.MoveNext
				 Loop
				 RS.Close:Set RS=Nothing
			  '==================================================================================
			  
			  
			  
		   End If
		End Sub
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file=../"Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS:Set KS=New PublicCLs
Dim KSUser:Set KSUser = New UserCls
Dim LoginTF:LoginTF=KSUser.UserLoginChecked
Dim SQLStr,RS,xml,node,Url,SignUser


'���ղ���
Dim Num:Num=KS.ChkClng(Request("num"))                                '�г����� 
Dim TitleLen:TitleLen=KS.ChkClng(Request("titlelen"))                 '��������
Dim Tid:Tid=KS.G("Tid")                                               '���õ���ĿID,������
Dim ShowClassName:ShowClassName=KS.ChkClng(request("showclassname"))  '��ʾ��Ŀ���� 1��ʾ 0����ʾ
Dim ShowDate:ShowDate=KS.ChkClng(Request("showdate"))                 '��ʾʱ�� 1��ʾ 0����ʾ


If Num=0 Then Num=10
Dim Param:Param=" Where Verific=1"
If Tid<>"" Then
  Param=Param & " and tid='" & tid & "'"
End If
SqlStr= "Select top " &num & " id,tid,title,adddate,fname,issign,signuser From KS_Article " & Param&" order by id desc"
Set RS=Server.CreateObject("adodb.recordset")
RS.Open SQLStr,conn,1,1
If Not RS.Eof Then
  Set xml=KS.RsToXml(rs,"row","")
End If
RS.Close
Set RS=Nothing
If Not IsObject(xml) Then KS.Die ""

For Each Node In Xml.DocumentElement.SelectNodes("row")
  Url=KS.GetItemUrl(1,Node.selectsinglenode("@tid").text,node.selectsinglenode("@id").text,node.selectsinglenode("@fname").text)
  SignUser=Node.selectsinglenode("@signuser").text
  KS.Echo "document.write('<li>"
  If ShowClassName=1 Then    '��ʾ��Ŀ����
   KS.Echo "<span class=""category"">[" & KS.GetClassNP(Node.SelectSingleNode("@tid").text) &"]</span>"
  End If
  KS.Echo "<a href=""" & url &""" target=""_blank"">" & KS.Gottopic(Node.SelectSingleNode("@title").text,TitleLen) &"</a>"
  If ShowDate=1 Then   '��ʾ����
    KS.Echo " " & year(node.selectsinglenode("@adddate").text) & "��" &month(node.selectsinglenode("@adddate").text)& "��" &day(node.selectsinglenode("@adddate").text) &"��"
  End If
  
  If node.selectsinglenode("@issign").text="1" and Not KS.IsNul(signuser) then
	  If LoginTF=True Then
	     If KS.FoundInArr(signuser,KSUser.UserName,",")=true Then   '��鵱ǰ�û��Ƿ���ǩ���û��б���
		       if conn.execute("select top 1 username from ks_itemsign where username='" & ksuser.username & "' and channelid=1 and infoid=" & node.selectsinglenode("@id").text).eof then
			     KS.Echo " <a href=""" & url & """ target=""_blank""><span class=""qs"">ǩ��</span></a>"
			   end if
		 End If
	  End If
  End If
  KS.Echo "</li>');" &vbcrlf
Next

%>

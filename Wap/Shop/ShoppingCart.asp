<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：我的购物车
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的购物车">
<p>
<%
Dim KSCls
Set KSCls = New ShoppingCart
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class ShoppingCart
        Private KS,DomainStr
		Private ProductList,LoginTF
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    LoginTF=KSUser.UserLoginChecked
		    ProductList = Session("ProductList")
			Dim FileContent,Products,i,RS,strsql,CarListStr
			Products = Split(Replace(KS.S("id")," ",""), ",")
			If Replace(KS.S("id")," ","")="" And KS.S("action")="set" Then 
			   ProductList=""
			ElseIf KS.S("Action")<>"set" Then 
			   For I=0 To UBound(Products)
				   PutToShopBag Products(I), ProductList,I
			   Next
			End iF
			If KS.S("Action")="Del" Then  DelProduct()
			Session("ProductList") = KS.FilterIds(ProductList)
			If Cbool(KSUser.UserLoginChecked)=False Then
			   Response.Write "温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""../User/Login/?../shop/ShoppingCart.asp?ID="&KS.S("ID")&""">登录</a>或<a href=""../User/Reg/?../shop/ShoppingCart.asp?ID="&KS.S("ID")&""">注册</a>成为商城会员！<br/>" &vbcrlf
			Else
			   Response.Write "亲爱的"&KSUser.UserName&"<br/>" &vbcrlf
			   Response.Write "【个人信息】<br/>" &vbcrlf
			   Response.Write "用户组:"&KS.GetUserGroupName(KSUser.GroupID)&"<br/>" &vbcrlf
			   Response.Write "可用资金:" & KSUser.Money & "元 " & KS.Setting(45) & ":" & KSUser.Point & "" & KS.Setting(46)&" 积分:" & KSUser.Score & "分<br/>" &vbcrlf
			End iF
			Response.Write "【购 物 车】<br/>" &vbcrlf
			
			Set RS=Server.CreateObject("ADODB.RecordSet") 
			If Session("ProductList")<>"" Then
			   strsql="select ID,Title,ProductType,Price_Original,Price,Price_Member,Discount,TotalNum,GroupPrice from KS_Product where ID in ("&Session("ProductList")&") order by ID"
			Else
			   strsql="select ID,Title,ProductType,Price_Original,Price,Price_Member,Discount,TotalNum,GroupPrice from KS_Product where 1=0 order by ID"
			End If
			RS.Open strsql,Conn,1,1
			Dim TotalPrice,RealPrice,Price_Original,Discount,Amount
			If Not RS.Eof Then
			   Amount = 1
			   Do While Not RS.EOF
			      Amount = KS.ChkClng(KS.S( "Q_" & RS("ID")))
				  If Amount <= 0 Then 
			         Amount = KS.ChkClng(Session("Amount"&RS("ID")))
				     If Amount <= 0 Then Amount = 1
				  End If
				  Session("Amount"&RS("ID")) = Amount
				  IF RS("TotalNum") < Amount Then
			         Amount = 1
					 Session("Amount"&RS("ID")) = 1
					 Response.Write "对不起，"&RS("Title")&"暂时库存不足，请过段时间再来购买该商品！<br/>" 
					 Exit Sub
				  End IF
				  IF Cbool(LoginTF)=true Then
			         If RS("GroupPrice")=0 Then
				        RealPrice=RS("Price_Member")
					 Else
					    Dim RSP:Set RSP=Conn.Execute("Select top 1 Price From KS_ProPrice Where GroupID=" & KSUser.GroupID & " And ProID=" & RS("ID"))
						If RSP.Eof Then
					       RealPrice=RS("Price_Member")
						Else
					       RealPrice=RSP(0)
						End If
						RSP.Close:Set RSP=Nothing
					 End If
				  Else
			         RealPrice=RS("Price")
				  End If
				  TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
				  Response.Write "编 号:"&RS("ID")&"<br/>" &vbcrlf
				  Response.Write "商品名称:"&RS("Title")&"<br/>" &vbcrlf
				  Response.Write "数 量:<input type=""Text"" name=""Q_"&RS("ID")&""&Minute(Now)&Second(Now)&""" value="""&Amount&""" size=""5""/>" &vbcrlf
				  Response.Write "<anchor>调整<go href=""ShoppingCart.asp?Action=set&amp;ID="&RS("ID")&"&amp;"&KS.WapValue&""" method=""post"">" &vbcrlf
				  Response.Write "<postfield name=""Q_"&RS("ID")&""" value=""$(Q_"&RS("ID")&""&Minute(Now)&Second(Now)&")""/>" &vbcrlf
                  Response.Write "</go></anchor>" &vbcrlf
				  Response.Write "<br/>" &vbcrlf
				  Response.Write "原 价:￥" & RS("Price_Original") & " 折 扣:" & RS("Discount") & "折 实 价:￥" & RealPrice & " 总 计:￥" & Round(RealPrice*Amount,2)  & "<br/>" &vbcrlf
				  Response.Write "操 作:<a href=""ShoppingCart.asp?Action=Del&amp;ID="&RS("ID")&"&amp;"&KS.WapValue&""">删除</a> <a href=""../User/User_Favorite.asp?Action=Add&amp;ChannelID=5&amp;InfoID="&RS("ID")&"&amp;"&KS.WapValue&""">收藏</a><br/><br/>" &vbcrlf
				  RS.MoveNext
			   Loop
			   Else
			      Response.Write "您的购物车没有商品!<br/>" &vbcrlf
			   End If
			   RS.close:set RS=nothing
			   Dim ID:id=KS.ChkClng(KS.S("ID"))
			   Response.Write "<br/>" &vbcrlf
  			   If ID=0 Then
			   Response.Write "<a href=""PayMent.asp?ID="&ProductList&"&amp;"&KS.WapValue&""">去收银台</a><br/>" &vbcrlf
               Else
			   Response.Write "<a href=""PayMent.asp?ID="&ProductList&"&amp;"&KS.WapValue&""">去收银台</a> <a href=""../Show.asp?ID="&KS.S("ID")&"&amp;ChannelID=5&amp;"&KS.WapValue&""">继续购物</a><br/>" &vbcrlf
			   End If
			   Response.Write "合计：￥" & Round(TotalPrice,2) & "元!<br/>" &vbcrlf
			   Response.Write "<br/><br/>" &vbcrlf
			   Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a>"
	    End Sub
		
		
		Sub PutToShopBag( Prodid, ProductList ,I)
		    If KS.S("Action")="set" Then
			   If i = 0 Then
				  ProductList =Prodid
			   ElseIf KS.FoundInArr( ProductList, Prodid,",")=false Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		   Else
			   If Len(ProductList) = 0 Then
				  ProductList =Prodid
			   ElseIf InStr( ProductList, Prodid ) <= 0 Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		    End If
        End Sub
		
	  Sub DelProduct()
	   Dim i,Parr:Parr=Split(ProductList,",")
	   Dim DelID:DelID=KS.S("ID")
	   Dim NewPList
	   For i=0 To Ubound(Parr)
	    If trim(Parr(i))<>trim(DelID) Then
		 If NewPlist="" Then
		  NewPlist=Parr(i)
		 Else
		  NewPlist=NewPlist & "," & Parr(I)
		 End If
		End If
	   Next
	   ProductList=NewPlist
	  End Sub
End Class
%>
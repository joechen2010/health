<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

Dim KS:Set KS=New PublicCls
Dim Action
Action=KS.S("Action")
Select Case Action
 Case "Ctoe" CtoE
 Case "GetTags" GetTags
 Case "GetRelativeItem" GetRelativeItem
 Case "Shop_GetCoupon" Shop_GetCoupon
 Case "Shop_ValidateCoupon" Shop_ValidateCoupon
 Case "Shop_BrandOption" Shop_BrandOption
 Case "Shop_CheckProID" Shop_CheckProID
 Case "GetClassOption" GetClassOption
 Case "GetFieldOption" GetFieldOption
 Case "SpecialSubList" SpecialSubList
 Case "GetArea" GetArea
 Case "GetFunc" GetFunc
 Case "AddFriend" AddFriend
 Case "MessageSave" MessageSave
 Case "CheckMyFriend" CheckMyFriend
 Case "SendMsg" SendMsg
 Case "SearchUser" SearchUser
 Case "CheckLogin" CheckLogin
 Case "relativeDoc" relativeDoc
 Case "getModelType" getModelType
 Case "getDocImage" getDocImage
 Case "checkDocFname" checkDocFname
 Case "addCart" addShoppingCart
 Case "GetClubBoard" GetClubBoard
 Case "GetPackagePro" GetPackagePro
 Case "GetSupplyContact" GetSupplyContact
End Select
Set KS=Nothing
CloseConn()

Sub getModelType()
 Dim ChannelID:ChannelID=KS.ChkClng(Request("channelid"))
 If ChannelID<>0 Then KS.Echo KS.C_S(Channelid,6)
End Sub

'自动关联文档
Sub relativeDoc()
 Dim NowID
 If Request("flag")="begin" Then
	 Response.Cookies(KS.SiteSn)("relative_startid")=1
 Else
     Response.Cookies(KS.SiteSn)("relative_startid")=KS.ChkClng(KS.C("relative_startid"))+1
 End If
 NowID=KS.ChkClng(KS.C("relative_startid"))
 getRelativeDoc NowID
 response.write KS.C("relative_count") & "|" & NowID
End Sub

Sub getRelativeDoc(n)
     Dim KeyWords,ChannelID,InfoID,Param,TopStr,SqlStr,TotalPut
	 If KS.ChkClng(Request("num"))<>0 Then TopStr=" top " & KS.ChkClng(Request("num"))
	 If KS.ChkClng(Request("ChannelID"))<>0 Then Param=" Where ChannelID=" &KS.ChkClng(Request("ChannelID")) 
     Dim RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
	 SqlStr="Select " & TopStr & " KeyWords,ChannelID,InfoID,title From KS_ItemInfo" & Param & " Order By Id Desc"
	 RSI.Open SqlStr,conn,1,1
	 If Not RSI.Eof Then 
	    TotalPut=RSI.Recordcount
   		If N=1 Then Response.Cookies(KS.SiteSn)("relative_count")=TotalPut
		If n>TotalPut Then n=TotalPut
		RSI.Move(n-1)
		KeyWords=RSI(0)
		ChannelID=RSI(1)
		InfoID=RSI(2)
		RSI.Close : Set RSI=Nothing
		If KeyWords<>"" And Not Isnull(KeyWords) Then 
			Dim KeyWordsArr, I, SqlKeyWordStr
			KeyWordsArr = Split(KeyWords, ",")
			 For I = 0 To UBound(KeyWordsArr)
				 If DataBaseType=0 Then
						 If SqlKeyWordStr = "" Then
								SqlKeyWordStr = " instr(keywords,'" & KeyWordsArr(I) & "')>0 "
						 Else
								SqlKeyWordStr = SqlKeyWordStr & "or instr(keywords,'" & KeyWordsArr(I) & "')>0 "
						 End If
				 Else
					 If SqlKeyWordStr = "" Then
							SqlKeyWordStr = " charindex('" & KeyWordsArr(I) & "',keywords)>0 "
					 Else
							SqlKeyWordStr = SqlKeyWordStr & "or charindex('" & KeyWordsArr(I) & "',keywords)>0 "
					 End If
				 End If
			Next
			
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 30 ChannelID,InfoID,Title From KS_ItemInfo Where ChannelID=" & ChannelID & " And InfoID<>" & InfoID & " and (" & SqlKeyWordStr & ")",conn,1,1
			Do While Not RS.Eof
			  Conn.Execute("Delete From KS_ItemInfoR Where ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and RelativeID=" & RS(1) & " And RelativeChannelID=" & RS(0))
			  Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & InfoID & "," & RS(0) & "," & RS(1) & ")")
			 RS.MoveNext
			Loop
		    RS.Close:Set RS=Nothing
	  End If
	 Else
	  RSI.Close:Set RSI=Nothing
	 End If
End Sub

'提取文档第一张图片
Sub getDocImage()
 Dim ChannelID:ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
 If ChannelID=0 Then Response.End
 Dim NowID
 If Request("flag")="begin" Then
	 Response.Cookies(KS.SiteSn)("relative_startid")=1
	 Response.Cookies(KS.SiteSn)("relative_has")=0
 Else
     Response.Cookies(KS.SiteSn)("relative_startid")=KS.ChkClng(KS.C("relative_startid"))+1
 End If
 NowID=KS.ChkClng(KS.C("relative_startid"))
 setDocImage ChannelID,NowID
 response.write KS.C("relative_count") & "|" & NowID & "|" & KS.C("relative_has")
End Sub
Sub setDocImage(ChannelID,n)
     Dim Content,PhotoUrl,TopStr
	 If KS.ChkClng(Request("num"))<>0 Then TopStr=" top " & KS.ChkClng(Request("num"))
	 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
	 RS.Open "Select " & TopStr &" PhotoUrl,PicNews,ArticleContent,ID From " & KS.C_S(Channelid,2) & " Where PicNews=0 Order by Id desc",conn,1,3
	 If Not RS.Eof Then
			If N=1 Then Response.Cookies(KS.SiteSn)("relative_count")=RS.Recordcount
			RS.Move(n-1)
			
			Dim regEx:Set regEx = New RegExp
			  regEx.IgnoreCase = True
			  regEx.Global = True
			  regEx.Pattern = "src\=.+?\.(gif|jpg)"
			  Content=KS.HtmlCode(rs(2))
			  Set Matches = regEx.Execute(Content)
			  If regEx.Test(Content) Then
			   Response.Cookies(KS.SiteSn)("relative_has")=KS.ChkClng(KS.C("relative_has"))+1
			   PhotoUrl=Lcase(Matches(0).value)
			   PhotoUrl=replace(PhotoUrl,"src=","")
			   PhotoUrl=replace(PhotoUrl,"""","")
			   PhotoUrl=replace(PhotoUrl,"'","")
			   RS(0)=PhotoUrl
			   RS(1)=1
			   RS.Update
			   Conn.Execute("Update KS_ItemInfo Set PhotoUrl='" & PhotoUrl & " Where ChannelID=" &ChannelID & " And InfoId=" & RS(3))
			  End If
	 End If
 RS.Close : Set RS=Nothing
End Sub

'检测文档文件名
Function checkDocFname()
 Dim ChannelID:ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
 Dim NowID
 If Request("flag")="begin" Then
	 Response.Cookies(KS.SiteSn)("relative_startid")=1
	 Response.Cookies(KS.SiteSn)("relative_has")=0
 Else
     Response.Cookies(KS.SiteSn)("relative_startid")=KS.ChkClng(KS.C("relative_startid"))+1
 End If
 NowID=KS.ChkClng(KS.C("relative_startid"))
 beginCheckDocFname ChannelID,NowID
 response.write KS.C("relative_count") & "|" & NowID & "|" & KS.C("relative_has")
End Function
Function beginCheckDocFname(ChannelID,N)
 Dim Param,SqlStr,TopStr
 If KS.ChkClng(Request("num"))<>0 Then TopStr=" top " & KS.ChkClng(Request("num"))
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 If DataBaseType=1 Then
  Param=" where fname is null or charindex('.',fname)=0"
 Else
  Param=" where fname is null or instr(fname,'.')=0"
 End If
 If ChannelID=0 Then
  SqlStr="Select" & TopStr & " Fname,InfoID,ChannelID From KS_ItemInfo " & Param & " Order By ID Desc"
 Else
  SqlStr="Select" & TopStr & " Fname,ID From " & KS.C_S(ChannelID,2) & " " & Param & " Order By ID Desc"
 End If
 RS.Open SqlStr,conn,1,3
 If Not (RS.Eof or rs.bof) Then
        If N=1 Then Response.Cookies(KS.SiteSn)("relative_count")=RS.Recordcount
		Do While Not RS.Eof
			RS(0)=RS(1) & ".html"
			RS.Update
			If ChannelID=0 Then
			 Conn.Execute("Update " & KS.C_S(RS(2),2) & " Set Fname='" & RS(0) & "' Where ID=" & RS(1))
			Else
			 Conn.Execute("Update KS_ItemInfo Set Fname='" & RS(0) & "' Where ChannelID=" & ChannelID & " And InfoID=" & RS(1))
			End If
			Response.Cookies(KS.SiteSn)("relative_has")=KS.ChkClng(KS.C("relative_has"))+1
		 RS.MoveNext
		Loop
 Else
  Response.Cookies(KS.SiteSn)("relative_count")=0
 End If
 RS.Close
 Set RS=Nothing
End Function


'取中文首字母
Sub Ctoe()
 Dim FolderName:FolderName=UnEscape(KS.G("FolderName"))
 Dim CE:Set CE=New CtoECls
 Response.Write Escape(CE.CTOE(FolderName))
 Set CE=Nothing
End Sub

'取关键词tags
Sub GetTags()
 Dim Text:Text=UnEscape(KS.G("Text"))
 If Text<>"" Then
     Dim MaxLen:MaxLen=KS.ChkClng(KS.S("MaxLen"))
	 Dim WS:Set WS=New Wordsegment_Cls
	 Response.Write Escape(WS.SplitKey(text,4,MaxLen))
	 Set WS=Nothing
 End If
End Sub


'相关信息
Sub GetRelativeItem()
 Dim Key:Key=UnEscape(KS.S("Key"))
 Dim Rtitle:rtitle=lcase(KS.G("rtitle"))
 Dim RKey:Rkey=lcase(KS.G("Rkey"))
 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("Channelid"))
 Dim ID:ID=KS.ChkClng(KS.G("ID"))
 Dim Param,RS,SQL,k,SqlStr
 If Key<>"" Then
   If (Rtitle="true" Or RKey="true") Then
	 If Rtitle="true" Then
	   param=Param & " title like '%" & key & "%'"
	 end if
	 If Rkey="true" Then
	   If Param="" Then
	     Param=Param & " keywords like '%" & key & "%'"
	   Else
	     Param=Param & " or keywords like '%" & key & "%'"
	   End If
	 End If
 Else
    Param=Param & " keywords like '%" & key & "%'"
 End If
End If

 
 If Param<>"" Then 
  	Param=" where InfoID<>" & id & " and (" & param & ")"
 else
    Param=" where InfoID<>" & id
 end if
 
  If ChannelID<>0 Then Param=Param & " and ChannelID=" & ChannelID


 SqlStr="Select top 30 ChannelID,InfoID,Title From KS_ItemInfo " & Param & " order by id desc"
 Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open SqlStr,conn,1,1
 If Not RS.Eof Then
  SQL=RS.GetRows(-1)
 End If
 RS.Close
 Set RS=Nothing
 If IsArray(SQL) Then
	 For k=0 To Ubound(SQL,2)
	   Response.Write "<option value='" & SQL(0,K) & "|" & SQL(1,K) & "'>" & SQL(2,K) & "</option>" 
	 Next
 End If
End Sub

Sub Shop_GetCoupon()
  Dim CouponUserID:CouponUserID=KS.ChkClng(KS.S("CouponID"))
  If CouponUserID=0 Then Exit Sub
  Dim RS:Set RS=Conn.Execute("SELECT Top 1 FaceValue,MinAmount,MaxDiscount,b.AvailableMoney,a.EndDate,a.status FROM KS_ShopCoupon A Inner Join KS_ShopCouponUser B ON A.ID=B.CouponID Where b.id=" & CouponUserID)
  If Not RS.Eof Then
   If DateDiff("s",RS("EndDate"),Now)>0 Then
    Response.Write "对不起,您输入的优惠券已过使用期限!"
   ElseIf RS("Status")=0 Then
    Response.Write "对不起,您输入的优惠券已被锁定!"
   ElseIf RS("AvailableMoney")<=0 Then
	Response.Write "对不起,您输入的优惠券已用完!"
   Else
    Response.Write RS(0) & "|" & RS(1) & "|" & RS(2)&"|"&RS(3)
   End If
  End If
  RS.Close:Set RS=Nothing
End Sub

Sub Shop_ValidateCoupon()
  Dim CouponNum:CouponNum=KS.S("CouponNum")
  If CouponNum="" Then Exit Sub
  Dim RS:Set RS=Conn.Execute("SELECT Top 1 A.FaceValue,A.MinAmount,A.MaxDiscount,b.AvailableMoney,a.BeginDate,a.EndDate,a.status FROM KS_ShopCoupon A Inner Join KS_ShopCouponUser B On A.ID=B.CouponID Where B.CouponNum='" & CouponNum & "'")
  If Not RS.Eof Then
       If DateDiff("s",RS("BeginDate"),Now)<0 Then
		Response.Write "对不起,您输入的优惠券需要" & RS("BeginDate") & "后才能使用!"
	   ElseIf DateDiff("s",RS("EndDate"),Now)>0 Then
		Response.Write "对不起,您输入的优惠券已过使用期限!"
	   ElseIf RS("Status")=0 Then
		Response.Write "对不起,您输入的优惠券已被锁定!"
	   ElseIf RS("AvailableMoney")<=0 Then
		Response.Write "对不起,您输入的优惠券已用完!"
	   Else
		Response.Write RS(0) & "|" & RS(1) & "|" & RS(2)&"|"&RS(3)
	   End If
  End If
  RS.Close:Set RS=Nothing
End Sub

'根据栏目ID得品牌列表
Sub Shop_BrandOption()
  Dim SQL,K
  Dim ClassID:ClassID=KS.G("ClassID")
  If ClassID="" Or ClassID="0"  Then Response.Write Escape("请先选择栏目!"):Response.End
  Dim Str:Str=GetBrandByClassID(ClassID,0)
  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
  If Str="Null" Then
     Response.Write Escape("&nbsp;<font color=blue>该栏目下没有添加品牌，请先</font><a onclick=""window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=管理中心 >> 品牌管理 >> <font color=red>新增品牌</font>&ButtonSymbol=GO'"" href='KS.ShopBrand.asp?action=Add&classid=" & classid & "'><font color=red>添加</font></a>")
  Else
	  Response.Write Escape(Str)
  End If
End Sub
		
Function GetBrandByClassID(ClassID,BrandID)
		  Dim SQL,K
		  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select B.ID,B.BrandName From KS_ClassBrand B inner join KS_ClassBrandR R On B.id=R.BrandID where R.classid='" & classid & "' order by B.orderid",conn,1,1
		  If Not RS.Eof  Then SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		  If Not IsArray(SQL) Then
		   GetBrandByClassID="Null" 
		  Else
		     GetBrandByClassID = "<select name='brandid'>"
			 GetBrandByClassID = GetBrandByClassID & "<option value='0'>-请选择品牌-</option>"
		     For K=0 To Ubound(SQL,2)
			  If BrandID=SQL(0,K) Then
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "' selected>" & sql(1,k) & "</option>"
			  Else
			  GetBrandByClassID=GetBrandByClassID & "<option value='" & sql(0,k) & "'>" & sql(1,k) & "</option>"
			  End If
			 Next
			 GetBrandByClassID = GetBrandByClassID &  "</select>"
			 Erase Sql
		  End If
End Function

'检查商品ID是否可用
Sub Shop_CheckProID()
 Dim proid:proid=UnEscape(KS.S("proid"))
 Dim ID:ID=KS.ChkClng(KS.S("ID"))
 Dim SQLStr
 If ProID="" Then 
   Response.Write Escape("你没有输入商品编号!")
 Else
   If Id=0 Then
    SqlStr="Select ProID From KS_Product Where ProID='" & ProID & "'"
   Else
    SqlStr="Select ProID From KS_Product Where ID<>" & ID & " and ProID='" & ProID & "'"
   End IF
   If Conn.Execute(SqlStr).Eof Then
    Response.Write Escape("恭喜,该商品编号可用!")
   Else
    Response.Write Escape("对不起,该商品编号已存在!")
   End If
 End If
End Sub

'检查是否登录
Sub CheckLogin()
  If KS.C("UserName")="" Then
   KS.Echo "false"
  Else
   KS.Echo "true"
  End If
End Sub

'取栏目选项
Sub GetClassOption()
 Dim ChannelID:ChannelID=KS.ChkCLng(Request.Querystring("ChannelID"))
 'If ChannelID=0 Then Exit Sub
  KS.Echo Escape(KS.LoadClassOption(ChannelID))
End Sub

Sub SpecialSubList()
	  Dim ClassID, RS,SpecialXML,Node
	  ClassID=KS.ChkClng(Request.QueryString("ClassID"))
	  If ClassID=0 Then Exit Sub
	  Set RS=Conn.Execute("Select SpecialID,SpecialName from KS_Special Where ClassID=" & ClassID & " Order BY SpecialAddDate Desc")
	  If Not RS.Eof Then Set SpecialXML=KS.RsToXml(RS,"row","xmlroot")
	  RS.Close:Set RS=Nothing
	  If IsObject(SpecialXml) Then
	  	For Each node in SpecialXml.DocumentElement.SelectNodes("row")
		  KS.Echo Escape("<div><img src=""images/folder/Special.gif"" align='absmiddle'>")
          KS.Echo Escape("<a href=""#"">"  & Trim(Node.SelectSingleNode("@specialname").text) & "</a><input type='checkbox' onclick=""set(" & Node.SelectSingleNode("@specialid").text & ",'" & Node.SelectSingleNode("@specialname").text & "');"" value='" & Node.SelectSingleNode("@specialid").text & "'></div>")
	    Next
		 Set SpecialXml=Nothing
      End If
End Sub

Sub GetFieldOption()
    Dim Node,ChannelID
	ChannelID=KS.ChkClng(Request.QueryString("ChannelID"))
	If ChannelID=0 Then Exit Sub
    If Not IsObject(Application(KS.SiteSN & "_ChannelField")) then KS.LoadChannelField
	If IsObject(Application(KS.SiteSN & "_ChannelField")) Then
	For Each Node In Application(KS.SiteSN & "_ChannelField").DocumentElement.SelectNodes("row[@channelid=" & ChannelID&"]")
    KS.Echo "<li class='diyfield' title=""" & Node.SelectSingleNode("@title").text &""" onclick=""InsertLabel('{@" & Node.SelectSingleNode("@fieldname").text & "}')"">" & Node.SelectSingleNode("@title").text & "</li>"
	Next
	End If
	Set Node=Nothing
End Sub

'取得ajax选项
sub GetArea()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,City FROM KS_Province " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'取得职能
sub GetFunc()
Dim Parentid:parentid=KS.ChkClng(Request("parentid"))
Dim Param:Param="where parentid=0"
if parentid<>0 then param=" where parentid=" & parentid
If Parentid<>0 Then
  response.write escape("<div><a href='javascript:void(0)' onclick='goBack()'>返回上一级</a></div>")
End If
Dim ors : set ors=Conn.Execute("select ID,hymc FROM KS_Job_hyzw " & Param & " order by orderid")
 do while not ors.eof
  if parentid=0 then
  response.write escape("<label><input type='checkbox' name='province' onclick='loadSecond(" & ors(0) & ",""" & ors(1) & """)' value='" & ors(1) & "'>" & ors(1) &" </label>")
  else
  response.write escape("<label><input type='checkbox' name='province' onclick='addPreItem()' value='" & ors(1) & "'>" & ors(1) &" </label>")
  end if
 ors.movenext
 loop
 ors.close
 set ors=nothing

end sub

'请求加为好友
Sub AddFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=UnEscape(KS.S("UserName"))
 Dim Message:Message=UnEscape(KS.S("Message"))
 If Len(Message)>255 Then 
   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
   exit sub
 End If
 If UserName="" Then KS.Echo escape("没有输入好友名称!") : Exit Sub
 call saveFriend(username,message,0)
 Set KSUser=New UserCls
 Call KSUser.AddLog(KS.C("UserName"),"给<a href=""{$GetSiteUrl}space/?" & username & """ target=""_blank"">" & username & "</a>发送加为好友请求!",106)
 Set KSUser=Nothing
 KS.Echo "success"
End Sub
'检查是否好友
Sub CheckMyFriend()
 If KS.C("UserName")="" Then KS.Echo "nologin" : Response.End
 Dim UserName:UserName=UnEscape(KS.S("UserName"))
 Dim RS:Set RS=Conn.Execute("Select accepted from KS_Friend Where UserName='" & KS.C("UserName") & "' and friend='" & username & "'")
 If rs.eof then
  KS.Echo "false"
 Else
  If rs(0)="1" then
   KS.Echo "true"
  Else
   KS.Echo "verify"
  End If
 End If
 RS.Close:Set RS=Nothing
End Sub

sub saveFriend(username,message,accepted)
		dim incept,i,sql,rs
		incept=KS.R(username)
		incept=split(incept,",")
		set rs=server.createobject("adodb.recordset")
		for i=0 to ubound(incept)
			sql="select top 1 UserName from KS_User where UserName='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				rs.close:set rs=nothing
				KS.Echo escape("系统没有（"&incept(i)&"）这个用户，操作未成功。")
				Set KS=Nothing
				Response.End
			end if
			set rs=Nothing
			
			if KS.C("UserName")=Trim(incept(i)) then
			   KS.Echo escape("不能把自已添加为好友。")
			   Set KS=Nothing
			   Response.End
			end if
			
			sql="select id,friend,accepted from KS_Friend where username='"&KS.C("UserName")&"' and  friend='"&incept(i)&"'"
			set rs=Conn.Execute(sql)
			if rs.eof and rs.bof then
				sql="insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&KS.C("UserName")&"','"&Trim(incept(i))&"',"&SqlNowString&",1,'" & replace(message,"'","") & "'," & accepted & ")"
				set rs=Conn.Execute(sql)
			else
			    if rs("accepted")=0 then
				  conn.execute("update ks_friend set message='" & replace(message,"'","") & "' where id=" & rs("id"))
				end if
			end if
			rs.close
		
		next
		set rs=nothing
end sub
'发送短消息
Sub SendMsg()
     If KS.C("UserName")="" Then Response.End
	 Dim UserName:UserName=UnEscape(KS.S("UserName"))
	 Dim Message:Message=UnEscape(KS.S("Message"))
	 If Len(Message)>255 Then 
	   KS.Echo escape("附言字数太多,最多只能输入255个字符!")
	   exit sub
	 End If

     Call KS.SendInfo(UserName,KS.C("UserName"),KS.Gottopic(Message,100),Message)
	 Set KSUser=New UserCls
     Call KSUser.AddLog(KS.C("UserName"),"给<a href="""  & KS.GetDomain & "space/?" & username & """ target=""_blank"">" & username & "</a>发送了一条消息!",107)
	 Set KSUser=Nothing
	 KS.Echo "success"
End Sub

'搜索好友
Sub SearchUser()
 Dim Page:Page=KS.ChkClng(Request("Page")) : If Page= 0 Then Page=1
 Dim Province:Province=UnEscape(KS.S("Province"))
 Dim City:City=UnEscape(KS.S("City"))
 Dim Birth_Y:Birth_Y=KS.ChkClng(Request("Birth_Y"))
 Dim Birth_M:Birth_M=KS.ChkClng(Request("Birth_M"))
 Dim Birth_D:Birth_D=KS.ChkClng(Request("Birth_D"))
 Dim RealName:RealName=UnEscape(KS.S("RealName"))
 Dim Sex:Sex=UnEscape(KS.S("Sex"))
 Dim RS:Set RS=Server.CreateObject("Adodb.recordset")
 Dim Param,SQLStr,XML,Node,totalPut,MaxPerPage,TotalPage,N
 MaxPerPage=10
 Param="Where locked=0"
 If Province<>"" Then Param=Param &" and Province='"& Province & "'"
 If City<>"" Then Param=Param & " and city='" & city & "'"
 If Sex<>"" Then Param=Param & " and sex='" & Sex & "'"
 If RealName<>"" Then Param=Param & " and realname like '%" & RealName & "%'"
 If Birth_Y<>0 Then Param=Param & " and year(birthday)=" & Birth_Y & ""
 If Birth_M<>0 Then Param=Param & " and month(birthday)=" & Birth_m & ""
 If Birth_D<>0 Then Param=Param & " and day(birthday)=" & Birth_d & ""

 
 SQLStr="Select userid,username,realname,sex,birthday,province,city,userface,isonline from ks_user " & param & " order by userid desc"
 'response.write sqlstr
 RS.Open SQLStr,conn,1,1
 If RS.Eof And RS.Bof Then
   RS.Close: Set RS=Nothing
    KS.Echo Escape("<div style='text-align:center'>对不起,找不到您要查找的用户!请更换查询条件,重新检索!</div>")
 Else
    totalPut = Conn.Execute("Select Count(*) From KS_User " & Param)(0)
	If Page < 1 Then	Page = 1
	If (totalPut Mod MaxPerPage) = 0 Then
		TotalPage = totalPut \ MaxPerPage
	Else
		TotalPage = totalPut \ MaxPerPage + 1
	End If
	
	If Page > 1  and (Page - 1) * MaxPerPage < totalPut Then
		RS.Move (Page - 1) * MaxPerPage
	Else
		Page = 1
	End If
	Set XML=KS.ArrayToXML(RS.GetRows(MaxPerPage),RS,"row","")
	RS.Close : Set RS=Nothing
	If IsObject(XML) Then
	  Dim user_face,UserName
	 For Each Node In XML.DocumentElement.SelectNodes("row")
	  user_face=node.selectsinglenode("@userface").text
	  If user_face="" then 
	    if node.selectSingleNode("@sex").text="男" then  user_face="images/face/0.gif" else user_face="images/face/girl.gif"
	  End If
	  If lcase(left(user_face,4))<>"http" then user_face=KS.Setting(2) & "/" & user_face
      username=Node.selectsinglenode("@username").text
	  KS.Echo "<li>"
	  KS.Echo "<table border='0' width='100%'>"
	  KS.Echo "<tr><td width='120' align='center' class='face'> <a href='" & KS.Setting(3) & "space/?" & username & "' target='_blank'><img src='" & user_face & "' alt='" & username & "' /></a></td>"
	  KS.Echo " <td align='left'>"
	  KS.Echo   Escape(Username & "(" & Node.SelectSingleNode("@realname").text & ")")
	  KS.Echo Escape(" <br />性别：" & Node.SelectSingleNode("@sex").text & "　出生：" & Node.SelectSingleNode("@birthday").text)
	  KS.Echo Escape(" <br />来自：" & Node.SelectSingleNode("@province").text & Node.SelectSingleNode("@city").text)
	  KS.Echo Escape(" <br />状态：")
	  If Node.SelectSingleNode("@isonline").text="1" Then KS.Echo escape("<font color='red'>在线</font>") else KS.Echo Escape("离线")
	  KS.Echo Escape(" <br /><img src='" & KS.Setting(3) & "images/user/log/106.gif' border='0'><a href='javascript:void(0)' onclick=""addF(event,'" & username & "')"">加为好友</a> <img src='" & KS.Setting(3) & "images/user/mail.gif'><a href=""javascript:void(0)"" onClick=""sendMsg(event,'" & username & "')"">发送消息</a>")
	  KS.Echo " </td>"
	  KS.Echo "</tr>"
	  KS.Echo "</table>"
	  KS.Echo "</li>"
	 Next
	End If
 End If
 If TotalPut<>0 Then
	 KS.Echo "<div id=""pageNext"" style=""text-align:center;clear:both;"">"
	 KS.Echo "<table align=""center""><tr><td>"
	 If Page>=2 Then
	  KS.Echo Escape("<a class='prev' href='javascript:void(0)' onclick=""query.page(" & Page-1 & ")"">上一页</a>")
	 End If
	 
	 If Page>=10 Then
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(1)"">1</a> <a class=""num"" href=""javascript:void(0)"" onclick=""query.page(2)"">2</a> <a href='#' class='dh'>...</a>"
	 End If
	 
	 Dim StartPage,EndPage
	 If TotalPage<10 Or Page<10 Then
	  StartPage=1
	  If Page<10 Then EndPage=10 Else  EndPage=TotalPage
	 ElseIf Page>=10 Then
	  StartPage=Page-4
	  EndPage=Page+4
	 ElseIf Page<TotalPage Then
	  StartPage=TotalPage-10
	  EndPage=TotalPage
	 End If
	 If EndPage>TotalPage Then EndPage=TotalPage : StartPage=TotalPage-10
	 If StartPage<0 Then StartPage=1
	 For N=StartPage To EndPage
	  If N=Page Then
	   KS.Echo "<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</a> "
	  Else
	   KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & n &")"">" & N & "</a> "
	  End If
	 Next
	 
	 If TotalPage>10 And Page<TotalPage-4 Then
	  KS.Echo "<a href='#' class='dh'>...</a>"
	  KS.Echo "<a class=""num"" href=""javascript:void(0)"" onclick=""query.page(" & TotalPage-1 & ")"">" & TotalPage-1 & "</a> <a href=""javascript:void(0)"" class=""num"" onclick=""query.page(" & TotalPage & ")"">" & TotalPage & "</a>"
	 End If
	 If Page<>TotalPage Then
	  KS.Echo Escape("<a class='next' href='javascript:void(0)' onclick=""query.page(" & Page+1 & ")"">下一页</a>")
	 End If
	 KS.Echo "</td></tr></table>"
	 
	 KS.Echo "</div>"
	End If
End Sub

'保存空间留言
Sub MessageSave()
		 Dim Content:Content=Request("Content")
		 Dim AnounName:AnounName=KS.S("AnounName")
         Dim HomePage:HomePage=KS.S("HomePage")
         Dim Title:Title=KS.S("Title")
		if AnounName="" Then 
		 Response.Write("请填写你的昵称!")
		 Response.End
		End if
		if Title="" Then 
		 Response.Write("请填写留言主题!")
		 Response.End
		End if
		if Content="" Then 
		 Response.Write("请填写留言内容!")
		 Response.End
		End if
		IF Trim(KS.S("Verifycode"))<>Trim(Session("Verifycode")) Then
		 Response.Write ("你输入的认证码不正确!")
		 Response.End
		End If
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_BlogMessage where 1=0",Conn,1,3
		RS.AddNew
		 RS("AnounName")=AnounName
		 RS("Title")=Title
		 RS("UserName")=KS.S("UserName")
		 RS("HomePage")=HomePage
		 RS("Content")=Content
		 RS("UserIP")=KS.GetIP
		 RS("AddDate")=Now
		RS.UpDate
		 RS.Close:Set RS=Nothing
		 If KS.C("UserName")<>"" Then
		  Set KSUser=New UserCls
		  Call KSUser.AddLog(KS.C("UserName"),"给用户<a href=""{$GetSiteUrl}space/?" & KS.S("UserName") & """ target=""_blank"">" & KS.S("UserName") &"</a>发了一条留言!",100)
		  Set KSUser=Nothing
		 End If
		 response.write "ok"
End Sub 

'加到购物车
Sub addShoppingCart()
   Dim Prodid:Prodid=KS.ChkClng(request("id"))
   if Prodid=0 then KS.Die ""
   Dim ProductList:ProductList=Session("ProductList")
   Dim Num:Num=KS.ChkClng(Request("Num"))
   Session("Amount"&Prodid) = Num
   If KS.S("AttributeCart")<>"" Then Session("AttributeCart"&Prodid)=UnEscape(KS.S("AttributeCart"))

   If Len(ProductList) = 0 Then
	  ProductList =Prodid
   ElseIf KS.FoundInarr(ProductList, Prodid,",") =false Then
	  ProductList = ProductList&","&Prodid &""
   End If
   ProductList=KS.FilterIds(ProductList)
   Session("ProductList")=ProductList
   If ProductList=""  Then Exit Sub
   Dim RS,RealPrice,n
   Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "Select id,title,GroupPrice,Price_Member,Price from KS_Product where id in (" & ProductList & ") and verific=1 order by id desc",conn,1,1
   if not rs.eof then
      KS.echo Escape("购物车中共有<font color=red>" & rs.recordcount & "</font>样商品!")
	  KS.Echo "<table border=0>"
	  n=1
	   Do While Not RS.Eof
		IF KS.C("UserName")<>"" Then
		  If RS("GroupPrice")=0 Then
		   RealPrice=RS("Price_Member")
		  Else
		   Dim RSP:Set RSP=Conn.Execute("Select Price From KS_ProPrice Where GroupID=(select groupid from ks_user where username='" & KS.C("UserName") & "') And ProID=" & RS("ID"))
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
	    Num=KS.ChkClng(Session("Amount"&rs(0)))
		If Num=0 Then Num=1
	    KS.Echo Escape("<tr><td style=""border-bottom:1px dashed #ccc;font-size:14px;font-weight:bold""><input type='hidden' name='id' value='" & rs(0) & "'>" & n & "、<font color=brown>" & rs(1) & "</font></td><td width='60' style=""border-bottom:1px dashed #ccc;color:#ff6600;"">￥" & RealPrice & "×" & Num & "</td></tr>")
		n=n+1
	   RS.MoveNext
	   Loop
	  KS.Echo "</table><br/>"
   end if
   RS.Close
   Set RS=Nothing
End Sub

Sub GetClubBoard()
 Call KS.LoadClubBoard()
   Dim node,Xml,n
   Set Xml=Application(KS.SiteSN&"_ClubBoard")
        KS.Echo Escape("<select name=""boardid"">")
   for each node in xml.documentelement.selectnodes("row[@parentid=0]")
		KS.Echo Escape("<optgroup label='" & node.selectsinglenode("@boardname").text &"'>")
		for each n in xml.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
		   KS.Echo Escape("<option value='" & N.SelectSingleNode("@id").text & "'>---" & n.selectsinglenode("@boardname").text &"</option>")
		next
	next
	KS.Echo Escape("</select>")
    Set Xml=Nothing
End Sub

Sub GetPackagePro()
    Dim RS,Key,pricetype,tid,minPrice,maxPrice,param,sqlstr,xml,node
	dim id:id=ks.chkclng(request("id"))
	dim proid:proid=ks.s("proid")
	Key=unescape(KS.S("Key"))
	pricetype=KS.ChkClng(KS.S("pricetype"))
	tid=KS.S("tid")
	minPrice=KS.S("minPrice"):If Not Isnumeric(minPrice) Then minPrice=0
	maxPrice=KS.S("maxPrice"):If Not Isnumeric(maxPrice) Then maxPrice=0
	param=" where verific=1"
	if tid<>"" and tid<>"0" then param=param & " and tid in(" & KS.GetFolderTid(TID) &")"
	if proid<>"" then param=param & " and proid='"& proid & "'"
    if id<>0 then param=param & " and id<>" & id 

	If PriceType<>0 Then
	  Select Case PriceType
	   case 1 : param=param & " and price>=" & minPrice & " and price<=" & maxPrice
	   case 2 : param=param & " and Price_Original>=" & minPrice & " and Price_Original<=" & maxPrice
	   case 3 : param=param & " and Price_Member>=" & minPrice & " and Price_Member<=" & maxPrice
	  End Select
	End If
	if key<>"" Then
	  Param=Param & " and title like '%" & key & "%'"
	End If
	sqlstr="select top 500 id,title from ks_product" & param & " order by id desc"
	
	
	set rs=conn.execute(sqlstr)
	if not rs.eof then
	 set xml=KS.RstoXml(rs,"row","")
	end if
	rs.close:set rs=nothing
	if isobject(xml) then
	  for each node in xml.documentelement.selectnodes("row")
       ks.echo "<option value='" & node.selectsinglenode("@id").text & "'>" & node.selectsinglenode("@title").text & "</option>"
	  next
    end if
End Sub

'查看联系信息
Sub GetSupplyContact()
 Dim ID:ID=KS.ChkClng(Request("id"))
 Set RS=Server.CreateObject("Adodb.Recordset")
 RS.Open "Select top 1 b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.verific=1 and a.ID=" & ID,Conn,1,1
 if rs.eof and rs.bof then
   rs.close:set rs=nothing
   ks.echo escape("加载出错!")
 else
   Dim KSUser:Set KSUser=New UserCls
   Dim UserLoginTF:UserLoginTF=KSUser.UserLoginChecked
    Dim ClassPurView:ClassPurview=rs("classpurview")
	Dim DefaultArrGroupID:DefaultArrGroupID=rs("defaultarrgroupid")
	 If ClassPurView="2" And Not KS.IsNul(DefaultArrGroupID) And DefaultArrGroupID<>"0" Then
		 IF UserLoginTF=false Then
		        response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您还没有登录，请<a href='" & KS.Setting(2) & "/user/login/' target='_blank'>登录</a>后再查看联系信息。</div>")
				rs.close:set rs=nothing
				response.end
		 ElseIf KS.FoundInArr(DefaultArrGroupID,KSUser.GroupID,",")=false Then
		        response.write ("<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您的级别不够,无法查看联系信息!得到更好服务,请联系本站管理员。</div>")
				rs.close:set rs=nothing
				response.end
		 End If
	 End If
   
   
   
   Dim template:template=LFCls.GetConfigFromXML("supply","/labeltemplate/label","contactinfo")
   template=replace(template,"{$GetContactMan}",rs("contactman"))
   template=replace(template,"{$GetContactTel}",rs("tel"))
   template=replace(template,"{$GetFax}",rs("fax"))
   template=replace(template,"{$GetEmail}",rs("email"))
   template=replace(template,"{$GetHomePage}",rs("homepage"))
   template=replace(template,"{$GetAddress}",rs("address"))
   ks.echo (template)   
 end if
 rs.close:set rs=nothing
End Sub

%>
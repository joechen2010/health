<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：栏目列表
'* 演示地址: http://wap.kesion.com/
'********************************
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New ClassCls
KSCls.Kesion()
Set KSCls = Nothing

Class ClassCls
        Private KS,KSRFObj,KMRFObj,PageStyle
		Private FileContent,SqlStr,ID,ChannelID,CurrPage,RSObj,PerPageNumber
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSRFObj = New Refresh
			Set KMRFObj= New RefreshFunction
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		    Set KSRFObj=Nothing
			Set KMRFObj=Nothing
		End Sub

        Public Sub Kesion()
		    ID=KS.ChkClng(KS.S("ID"))
			IF ID=0 Then
			   Call KS.ShowError("非法参数！","非法参数！")
			End If
			Set RSObj=Server.CreateObject("Adodb.Recordset")
			SqlStr= "Select ID,ClassPurview,TN,WapFolderTemplateID,FolderDomain,DefaultArrGroupID,ChannelID From KS_Class Where ClassID=" & ID
			RSObj.Open SqlStr,Conn,1,1
			IF RSObj.Eof And RSObj.Bof Then
			   RSObj.Close:Set RSObj=Nothing
			   Call KS.ShowError("非法参数！","非法参数！")
			End If
			ChannelID=RSObj("ChannelID")

			Select Case KS.C_S(ChannelID,6)
			    Case 1,2,3
				If RSObj("ClassPurview")=2 Then
				   If Cbool(KSUser.UserLoginChecked)=false Then 
				      RSObj.Close:Set RSObj=Nothing
					  Call KS.ShowError("对不起！","本栏目为认证栏目，至少要求本站的注册会员才能浏览！")
				   ElseIF KS.FoundInArr(RSObj("DefaultArrGroupID"),KSUser.GroupID,",")=false Then
			          RSObj.Close:Set RSObj=Nothing
					  Call KS.ShowError("对不起！","对不起，你所在的用户级没有权限浏览！")
				   End If
				End If
		    End Select
			Call FCls.SetClassInfo(ChannelID,RSObj("ID"),RSObj("TN"))
			FileContent = KSRFObj.LoadTemplate(RSObj("WapFolderTemplateID"))			
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			
			Dim PageParamArr:PageParamArr=Split(Application("PageParam"),",")
			If Ubound(PageParamArr)>0  Then
			   If PageParamArr(0)="GetShowClassCent" Then
				  PageStyle=PageParamArr(7)'分页样式
				  
				  Dim sPai,sType,sTypeNum
				  Dim TempStr
				  Randomize()
				  CurrPage=KS.ChkClng(KS.S("Page"))
				  If CurrPage<=0 Then CurrPage=CurrPage+1
				  sPai=KS.ChkClng(KS.S("sPai"))
				  If sPai<=0 Then sPai=2'排列
				  
				  sType=KS.ChkClng(KS.S("sType"))
				  If sType=0 Then sType=KS.ChkClng(PageParamArr(3))
				  
				  If sType=2 Then
				     sTypeNum=KS.ChkClng(KS.S("sTypeNum"))
					 If sTypeNum<=0 Then sTypeNum=PageParamArr(5)'豪华详版重复显示条数
				  Else
				     sTypeNum=PageParamArr(4)'文字简版重复显示条数
				  End If
				  
				  If Cbool(PageParamArr(1))=True Then
				     If sType=1 Then
					    TempStr = TempStr & "文字简版 "
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sType=1&amp;" & KS.WapValue & """>文字简版</a> "
					 End If
					 If sType=2 Then
					    TempStr = TempStr & "豪华详版<br/>"
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sTypeNum=4&amp;sType=2&amp;" & KS.WapValue & """>豪华详版</a><br/>"
					 End If
				  End If
				  
				  If Cbool(PageParamArr(2))=True Then
				     If sPai=1 Then
					    TempStr = TempStr & "最早 "
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=1&amp;sTypeNum=" & sTypeNum & "&amp;sType=" & sType & "&amp;" & KS.WapValue & """>最早</a> "
					 End If
					 If sPai=2 Then
					    TempStr = TempStr & "最新 "
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=2&amp;sTypeNum=" & sTypeNum & "&amp;sType=" & sType & "&amp;" & KS.WapValue & """>最新</a> "
					 End If
					 If sPai=3 Then
					    TempStr = TempStr & "热门 "
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=3&amp;sTypeNum=" & sTypeNum & "&amp;sType=" & sType & "&amp;" & KS.WapValue & """>热门</a> "
					 End If
					 If sPai=4 Then
					    TempStr = TempStr & "随机<br/>"
					 Else
					    TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=4&amp;sTypeNum=" & sTypeNum & "&amp;sType=" & sType & "&amp;" & KS.WapValue & """>随机</a><br/>"
					 End If
				  End If
				  
				  If sType=2 Then
				     If CurrPage=1 Then
					    TempStr = TempStr & "每页显:"
						If sTypeNum=2 Then
						   TempStr = TempStr & "2 "
						Else
						   TempStr = TempStr & "<a href=""list.asp?ID=" & ID &"&amp;sPai=" & sPai & "&amp;sTypeNum=2&amp;sType=" & sType & "&amp;" & KS.WapValue & """>2</a> "
						End If
						If sTypeNum=4 Then
						   TempStr = TempStr & "4 "
						Else
						   TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai="&sPai&"&amp;sTypeNum=4&amp;sType=" & sType & "&amp;" & KS.WapValue & """>4</a> "
						End If
						If sTypeNum=6 Then
						   TempStr = TempStr & "6 "
						Else
						   TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=" & sPai & "&amp;sTypeNum=6&amp;sType=" & sType & "&amp;" & KS.WapValue & """>6</a> "
						End If
						If sTypeNum=8 Then
						   TempStr = TempStr & "8 "
						Else
						   TempStr = TempStr & "<a href=""list.asp?ID=" & ID & "&amp;sPai=" & sPai & "&amp;sTypeNum=8&amp;sType=" & sType & "&amp;" & KS.WapValue & """>8</a> "
						End If
						TempStr = TempStr & KS.C_S(ChannelID,4)&"<br/>"
					 End If
				  Else
				      If TempStr<>"" Then TempStr = TempStr & "---------<br/>" 
				  End If
				  
				  Dim Param,Asort,FolderID:FolderID=RSObj("ID")
				  Param = " Where Tid='" & FolderID & "' AND Verific=1 AND DelTF=0"
				  Select Case sPai
				      Case "1":ASort=" order by IsTop Desc,AddDate asc"
					  Case "2":ASort=" order by IsTop Desc,AddDate desc"
					  Case "3":ASort=" order by IsTop Desc,Hits desc"
					  Case "4"
					  If DataBaseType=0 then
					     ASort=" order by IsTop Desc,Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
					  Else
					     ASort=" order by IsTop Desc,newid()"
					  End if
				  End Select
				  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				  RS.Open "SELECT ID FROM " & KS.C_S(ChannelID,2) & Param & ASort, Conn, 1, 1
				  If RS.EOF And RS.BOF Then
				     TempStr = "此栏目下没有" & KS.C_S(ChannelID,4) & "内容!<br/>"
				  Else
				     PerPageNumber = KS.ChkClng(sTypeNum)
				     Dim PageNum,totalput,TempIDArrStr
					 TotalPut = Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & Param)(0)
					 If (TotalPut Mod PerPageNumber)=0 Then
					    PageNum = TotalPut \ PerPageNumber
					 Else
					    PageNum = TotalPut \ PerPageNumber + 1
					 End If
					 If CurrPage = 1 Then
					    TempIDArrStr=GetTempIDArrStr(RS)
					 Else
					    If (CurrPage - 1) * PerPageNumber < totalPut Then
						   RS.Move (CurrPage - 1) * PerPageNumber
						   TempIDArrStr=GetTempIDArrStr(RS)
						Else
						   CurrPage = 1
						   TempIDArrStr=GetTempIDArrStr(RS)
						End If
					 End If
					 Select Case KS.C_S(ChannelID,6)
					     Case 1
						 SqlStr = "SELECT ID,Title,TitleType,Intro,ArticleContent,AddDate,PhotoUrl,Popular,Fname,Changes FROM " & KS.C_S(ChannelID,2) & " Where ID in (" & TempIDArrStr & ") AND Verific=1 AND DelTF=0 order by IsTop Desc"
						 If sType=1 Then
						    TempStr = TempStr & KMRFObj.KS_A_L(ChannelID, SqlStr,False,PageParamArr(6), True,True,True)
						 Else
						    TempStr = TempStr & KMRFObj.KS_PicA_L(ChannelID,SqlStr,128,128,True,4,50,PageParamArr(6))
						 End If
						 
						 Case 2 
						 SqlStr = "SELECT ID,Title,Tid,ReadPoint,InfoPurview,Fname,AddDate,PhotoUrl,PictureContent FROM " & KS.C_S(ChannelID,2) & " Where ID in (" & TempIDArrStr & ") AND Verific=1 AND DelTF=0 order by IsTop Desc"
						 If sType=1 Then
						    TempStr = TempStr & KMRFObj.KS_P_L(2,SqlStr,"128","128",3,PageParamArr(6),"")
						 Else
						    TempStr = TempStr & KMRFObj.KS_P_L(2,SqlStr,"128","128",2,PageParamArr(6),"")
						 End If
						 
						 Case 3
						 SqlStr = "SELECT ID,Title,Tid,DownVersion,PhotoUrl,DownContent FROM " & KS.C_S(Channelid,2) & " Where ID in (" & TempIDArrStr & ") AND Verific=1 AND DelTF=0 order by IsTop Desc" 
						 If sType=1 Then
						    TempStr = TempStr & KMRFObj.KS_D_L(ChannelID,SqlStr,PageParamArr(6),1,"")
						 Else
						    TempStr = TempStr & KMRFObj.KS_C_PicD_L(ChannelID,SqlStr,128,128,True,4,50,PageParamArr(6))
						 End If
						 
						 Case 5
						 SqlStr = "SELECT ID,Title,Tid,Fname,AddDate,PhotoUrl,Discount,Price_Original,Price,Price_Market,Price_Member FROM KS_Product Where ID in (" & TempIDArrStr & ") AND Verific=1 AND DelTF=0 order by IsTop Desc"
						 If sType=1 Then
						    TempStr = TempStr & KMRFObj.KS_Pro_L(SqlStr,1,7,0,True,128,128,PageParamArr(6))
						 Else
						    TempStr = TempStr & KMRFObj.KS_Pro_L(SqlStr,6,7,0,True,128,128,PageParamArr(6))
						 End If
						 
						 Case 8
						 SqlStr = "SELECT ID,Title,Tid,Fname,AddDate,PhotoUrl,GQContent,TypeID,Province,City FROM KS_GQ Where ID in (" & TempIDArrStr & ") AND Verific=1 AND DelTF=0 order by IsTop Desc"
						 If sType=1 Then
						    TempStr = TempStr & KMRFObj.KS_S_L(SqlStr,1,128,128,60,30,"",False,True,True)
						 Else
						    TempStr = TempStr & KMRFObj.KS_S_L(SqlStr,4,128,128,60,30,"",False,True,True)
						 End If
					 End Select
					 TempStr = TempStr & KS.GetPrePageList(PageStyle,KS.C_S(ChannelID,4),PageNum,CurrPage,TotalPut,PerPageNumber)
					 TempStr = TempStr & KS.GetPageList("list.asp?" & KS.QueryParam("page") &"&amp;",PageStyle,CurrPage,PageNum, True)
				  End If
				  RS.Close:set RS=Nothing
		       End If
			End If
			FileContent = Replace(FileContent,Application("PageParam"),TempStr)
			FileContent = KS.GetEncodeConversion(FileContent)
			Response.Write FileContent
			RSObj.Close:Set RSObj=Nothing
	    End Sub
		
	    Function GetTempIDArrStr(RS)
	        Dim I,K,SQL
		    SQL=RS.GetRows(-1)
		    For K=0 To Ubound(SQL,2)
			    GetTempIDArrStr = GetTempIDArrStr &SQL(0,K) & ","
			    I = I + 1
			    If I >= PerPageNumber Then Exit For
		    Next
		    GetTempIDArrStr = Left(GetTempIDArrStr, Len(GetTempIDArrStr) - 1)
	    End Function
End Class
%>

<%@ Language="VBSCRIPT" codepage="936" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KS:Set KS=New PublicCls
Dim KSCls:Set KSCls=New DIYCls

Dim CurrPage,Tconn
Dim SqlLabel:SqlLabel=KS.S("LabelID")
CurrPage=KS.ChkClng(KS.S("curpage"))
If CurrPage<=0 Then CurrPage=1

		  Dim I,KS_RS_Obj,LabelName,UserParamArr,FunctionLabelParamArr,CirLabelContent,FunctionSQL,LabelContent,TempCirContent
		  Dim FunctionLabelType,ItemName,PageStyle,PerPageNumber,TotalPut,PageNum,J,TempStr,Ajax,DataSourceType,DataSourceStr

Function GetPageContent()
          if KS.IsNul(request.ServerVariables("http_referer")) Then
		     KS.Die "�벻Ҫ�Ƿ�����!"
		  End If

		  LabelName    = Replace(Replace(Split(SqlLabel,"(")(0),"'",""),"""","")
		  '�û���������
		  UserParamArr = Split(Replace(Replace(Replace(Replace(SqlLabel,LabelName&"(",""),")}",""),"""",""),"'",""),",")   
		  
		   Dim L_Description:L_Description=KSCls.G_S_P(LabelName &"}",1)
		   If L_Description="" Then
		    GetPageContent="�Բ��𣬱�ǩ������!":exit function
		   Else
		    FunctionLabelParamArr = Split(L_Description,"@@@")
		    LabelContent          = Replace(KSCls.G_S_P(LabelName &"}",2),Chr(10) ,"$KS:Page$")
		   End If
		  
		   FunctionSQL=FunctionLabelParamArr(0)           '��ѯ���
		   FunctionSQL=Replace(FunctionSQL,"{$CurrClassID}",KS.S("classID"))
		   FunctionSQL=Replace(FunctionSQL,"{$CurrClassChildID}",KS.GetFolderTid(KS.S("classID")))
		   LabelContent = KSCls.ReplaceRequest(LabelContent)    '�滻request��ֵ
		   FunctionSQL = KSCls.ReplaceRequest(FunctionSQL)    '�滻request��ֵ

		   For I=0 To Ubound(UserParamArr)
		    FunctionSQL  = Replace(FunctionSQL,"{$Param("&I&")}",KS.DelSQL(UserParamArr(I)))
			LabelContent = Replace(LabelContent,"{$Param("&I&")}",KS.DelSQL(UserParamArr(I)))
		   Next
		   FunctionLabelType=FunctionLabelParamArr(2)
		   If Not Isnumeric(FunctionLabelType) Then FunctionLabelType=0
		   Ajax=FunctionLabelParamArr(5)
           		   
		   ItemName=FunctionLabelParamArr(3)
		   PageStyle=FunctionLabelParamArr(4)
		   DataSourceType=FunctionLabelParamArr(6)
		   DataSourceStr=FunctionLabelParamArr(7)
		   if DataSourceType=1 Or DataSourceType=5 Or DataSourceType=6 then	DataSourceStr=LFCls.GetAbsolutePath(DataSourceStr)
		   If OpenExtConn=false Then GetPageContent="�ⲿ���ݿ����ӳ���!":exit function
		   on error resume next
		   Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
           If DataSourceType=0 Then
		    KS_RS_Obj.Open FunctionSQL,Conn,1,1
		   Else
		    KS_RS_Obj.Open FunctionSQL,TConn,1,1
		   End IF
		   if err then 
		    err.clear
			KS_RS_Obj.close: set KS_RS_Obj=nothing
			KS.Die "�Ƿ�����!"
		   end if
		   
		   If Not KS_RS_Obj.Eof Then
			    Dim regEx, Matches, Match,LoopTimes
				Set regEx = New RegExp
				regEx.Pattern = "\[loop=\d*].+?\[/loop]"
				regEx.IgnoreCase = True
				regEx.Global = True
				Set Matches = regEx.Execute(LabelContent)
				If FunctionLabelType=1 Then                  '��ҳ��ǩ
				         PerPageNumber=0
				         For Each Match In Matches
							PerPageNumber=PerPageNumber+KSCls.GetLoopNum(Match.Value)   'ÿҳ��¼��
						 Next
                         If PerPageNumber=0 Then GetPageContent= "�Զ��庯����ǩ��ѭ�������������0":exit function
						 
				  		TotalPut = KS_RS_Obj.recordcount
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
							 TempCirContent    = LabelContent
							 KS_RS_Obj.Move (CurrPage - 1) * PerPageNumber
						     For Each Match In Matches
								  LoopTimes=KSCls.GetLoopNum(Match.Value)   'ѭ������
								  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
								  TempCirContent    = Replace(TempCirContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",KSCls.GetCirLabelContent(CirLabelContent,KS_RS_Obj,LoopTimes),1,1)

								  If KS_RS_Obj.Eof Then Exit For
							 Next
							  TempStr = TempCirContent
						      GetPageContent=Replace(KSCls.CleanLabel(TempStr),"$KS:Page$",vbcrlf)

				End If		 
		   Else
		     GetPageContent="�Բ���û������!":exit function
		   End if
		   KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   If DataSourceType=0 Then
		   Conn.Close:Set Conn=Nothing
		   Else
		   TConn.Close:Set TConn=Nothing
		   End If
   End Function


		Function OpenExtConn()
		 If DataSourceType=0 Then
		   OpenExtConn=True
		 Else
			on error resume next
		    Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open datasourcestr
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			   OpenExtConn=False
			Else 
			   OpenExtConn=true
			End If
		 End If
    	End Function

Response.Write GetPageContent
Response.Write "{ks:page}" & TotalPut & "|" & PerPageNumber & "|" & PageNum & "|" & ItemName & "||" & PageStyle
closeconn
%>

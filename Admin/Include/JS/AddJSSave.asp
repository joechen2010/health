<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
Dim TempClassList,InstallDir,CurrPath,JSConfig,KS,KSRObj,FolderID,TempSymbol
Dim JSID,JSRS,SQLStr,JSName,JSFunctionFlag,Descript,Action,RSCheck,FileUrl,JSType,JSFileName
Dim KeyWord,SearchType,StartDate,EndDate
  
'�ռ���������
KeyWord=Request("KeyWord")
SearchType=Request("SearchType")
StartDate = Request("StartDate")
EndDate = Request("EndDate")

FileUrl=Request("FileUrl") '���������Ϻ󷵻�
Set KS=New PublicCls
Set KSRObj=New Refresh
	JSFileName=Replace(Replace(Trim(Request.Form("JSFileName")),"""",""),"'","")
	if instr(JSFileName,";")<>0 or instr(lcase(JSFileName),".asp")<>0 or instr(lcase(JSFileName),".php")<>0 or instr(lcase(JSFileName),".cer")<>0 or instr(lcase(JSFileName),".asa")<>0 then
       Call KS.AlertHistory("JS���Ƹ�ʽ���Ϸ�!",-1)
	   Set KS=Nothing
	   Response.End
	end if

Set JSRS=Server.CreateObject("Adodb.RecordSet")
Select Case Request.Form("Action")
 Case "Add" 
    JSName= Replace(Replace(Trim(Request.Form("JSName")),"""",""),"'","")
    Descript=Replace(Trim(Request.Form("Descript")),"'","")
    JSConfig=Trim(Request.Form("JSConfig"))
	JSType=Request.Form("JSType")
	FolderID=Request.Form("ParentID")
    IF FolderID="" Then FolderID="0"
	IF JSType="" Then JSType=0
    IF JSName="" THEN
       Call KS.AlertHistory("JS���Ʋ���Ϊ��!",-1)
	   Set KS=Nothing
	   Response.End
    END IF
	IF UCASE(Right(JSFileName,3))<>".JS" THEN
	  Call KS.AlertHistory("JS�ļ�������չ��������.js",-1)
	  Set KS=Nothing
	  Response.End
	END IF
    IF JSConfig="" THEN
      Call KS.AlertHistory("JS���ݲ���Ϊ��!",-1)
	  Set KS=Nothing
	  Response.End
    END IF
	JSName="{JS_" & JSName & "}"
	JSRS.Open "Select JSName From [KS_JSFile] Where JSName='" & JSName & "' Or JSFileName='" & JSFileName &"'",Conn,1,1
	IF Not JSRS.EOF Then
	  if Trim(JSRS("JSName"))=JSName Then
	   Response.Write("<script>alert('JS�����Ѿ�����!');location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID &"';</script>")
	  else
	   Response.Write("<script>alert('JS�ļ����Ѿ�����!');location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID &"';</script>")
	  end if
	  JSRS.Close
	  Conn.Close
	  Set JSRS=Nothing
	  Set Conn= Nothing
	  Set KS=Nothing
	  Response.End
	ELSE
	    JSRS.Close
		JSRS.Open "Select * From [KS_JSFile] Where (JSID is Null)",Conn,1,3
		JSRS.AddNew
		  Do While True
		    '����ID  ��+6λ���
            JSID = Year(Now()) & KS.MakeRandom(6)
            Set RSCheck = conn.execute("Select JSID from [KS_JSFile] Where JSID='" & JSID & "'")
             If RSCheck.EOF And RSCheck.BOF Then
              RSCheck.Close
			  Set RSCheck=Nothing
              Exit Do
             End If
          Loop
		 JSRS("JSID")=JSID
		 JSRS("JSName")=JSName
		 JSRS("JSFileName")=JSFileName
		 JSRS("Description")=Descript
		 JSRS("JSConfig")=JSConfig
		 JSRS("JSType")=JSType
		 JSRS("AddDate")=now
		 JSRS("OrderID")=1
		 JSRS("FolderID")=FolderID
		 JSRS.Update
		IF JSType=0 Then
		    TempSymbol="&OpStr=JS����  >> ϵͳJS &ButtonSymbol=SysJSList"
		 ELSE
		    TempSymbol="&OpStr=JS����  >> ����JS &ButtonSymbol=FreeJSList"
		 END IF
		 KSRObj.RefreshJS(JSName)
		 JSRS.Close
		 Set JSRS=Nothing
		 Set KSRObj=Nothing
     	Response.Write("<script>if (confirm('�ɹ���ʾ:\n\n���JS�ɹ�,�������JS��?')){location.href='" & FileUrl & "?Action=Add&FolderID=" & FolderID & "';}else{top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?FolderID=" & FolderID &"&JSType=" & JSType & "';}</script>") 
	END IF
Case "Edit"
    Dim Page
	Page=Request.Form("Page")
    JSID=Trim(Request.Form("JSID"))
    JSName= Replace(Replace(Trim(Request.Form("JSName")),"""",""),"'","")
    Descript=Replace(Trim(Request.Form("Descript")),"'","")
    JSConfig=Trim(Request.Form("JSConfig"))
	JSType=Request.Form("JSType")
	FolderID=Request.Form("ParentID")
	IF FolderID="" Then FolderID="0"
	IF JSType="" Then JSType=0
    IF JSName="" THEN
       Call KS.AlertHistory("JS���Ʋ���Ϊ��!",-1)
	   Set KS=Nothing
	   Response.End
    END IF
    IF JSConfig="" THEN
      Call KS.AlertHistory("JS���ݲ���Ϊ��!",-1)
	  Set KS=Nothing
	  Response.End
    END IF
	JSName="{JS_" & JSName & "}"
	JSRS.Open "Select JSName From [KS_JSFile] Where JSID <>'" & JSID &"' AND JSName='" & JSName & "'",Conn,1,1
	IF Not JSRS.EOF Then
	  Response.Write("<script>alert('JS�����Ѿ�����!');location.href='" & FileUrl & "?Page=" & Page & "&JSID=" & JSID & "';</script>")
	  JSRS.Close
	  Conn.Close
	  Set JSRS=Nothing
	  Set Conn= Nothing
	  Set KS=Nothing
	  Response.End
	ELSE
	    JSRS.Close
		JSRS.Open "Select * From [KS_JSFile] Where JSID='" & JSID &"'",Conn,1,3
		 JSRS("JSName")=JSName
		 JSRS("Description")=Descript
		 JSRS("JSConfig")=JSConfig
		 JSRS("JSType")=JSType
		 JSRS("FolderID")=FolderID
		 JSRS.Update
		 KSRObj.RefreshJS(JSName)

		 IF KeyWord="" Then
		    IF JSType=0 Then
		        TempSymbol="&OpStr=JS����  >> ϵͳJS &ButtonSymbol=SysJSList"
		    ELSE
		        TempSymbol="&OpStr=JS����  >> ����JS &ButtonSymbol=FreeJSList"
		    END IF
     	   Response.Write("<script>alert('�ɹ���ʾ:\n\nJS�޸ĳɹ�!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?FolderID="& FolderID & "&Page=" & Page & "&JSType=" & JSType & "';</script>") 
		 ELSE
		    IF JSType=0 Then
		        TempSymbol="OpStr=JS����  >> <font color=red>����ϵͳJS���</font>&ButtonSymbol=SysJSSearch"
		    ELSE
		        TempSymbol="OpStr=JS����  >> <font color=red>��������JS���</font>&ButtonSymbol=FreeJSSearch"
		    END IF
     	   Response.Write("<script>alert('�ɹ���ʾ:\n\nJS�޸ĳɹ�!');top.frames['BottomFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?" &TempSymbol &"';top.frames['MainFrame'].location.href='" & KS.Setting(3) & KS.Setting(89) & "include/JS_Main.asp?KeyWord="& KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate &"&EndDate=" & EndDate &"&Page=" & Page & "&JSType=" & JSType & "';</script>") 
		 END IF
	END IF
		 JSRS.Close
		 Set JSRS=Nothing
		 Set KSRObj=Nothing
End Select
%>
 

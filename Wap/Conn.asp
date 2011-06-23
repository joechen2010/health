<%
Dim SqlNowString,DataPart_D,DataPart_Y,DataPart_H
Dim Conn,DBPath,CollectDBPath,DataServer,DataUser,DataBaseName,DataBasePsw,ConnStr,CollcetConnStr
Const DataBaseType=0                   '系统数据库类型，"1"为MS SQL2000数据库，"0"为MS ACCESS 2000数据库
Const MsxmlVersion=".3.0"                '系统采用XML版本设置 

Const WapCharset="GB2312"   '主Web程序用的编码类型,请填写GB2312或UTF-8,如果是GB2312,那么Wap模板请保存成Gb2312格式,反之请保存成UTF-8格式
Const G_Domain=""      'IIS里绑定的独立域名或二级域名,必须以"/"结束;没有绑定请留空,否则可能导致WAP无法使用
 
If G_Domain<>"" And Lcase(GetAutoDomain)<>Lcase(G_Domain) Then
  Response.Redirect G_Domain
End If
 
If DataBaseType=0 then
	'如果是ACCESS数据库，请认真修改好下面的数据库的文件名
	DBPath       = GetMapPath & "\KS_Data\KesionCMS6.mdb"      'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径
Else
	 '如果是SQL数据库，请认真修改好以下数据库选项
	 DataServer   = "(local)"                                  '数据库服务器IP
	 DataUser     = "sa"                                       '访问数据库用户名
	 DataBaseName = "KesionCMS"                                '数据库名称
	 DataBasePsw  = "989066"                                   '访问数据库密码 
End if

'=============================================================== 以下代码请不要自行修改========================================
Call OpenConn
Sub OpenConn()
    On Error Resume Next
    If DataBaseType = 1 Then
       ConnStr="Provider = Sqloledb; User ID = " & datauser & "; Password = " & databasepsw & "; Initial Catalog = " & databasename & "; Data Source = " & dataserver & ";"
	   SqlNowString = "getdate()"
	   DataPart_D   = "d"
	   DataPart_Y   = "y"
	   DataPart_H   = "hour"
    Else
       ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath
	   SqlNowString = "Now()"
	   DataPart_D   = "'d'"
	   DataPart_Y   = "'yyyy'"
	   DataPart_H   = "'h'"
    End If
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open ConnStr
    If Err Then Err.Clear:Set conn = Nothing:Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。":Response.End
End Sub
Sub CloseConn()
    On Error Resume Next
	Conn.close:Set Conn=nothing
End sub

'获取实现的物理路径
Function GetMapPath()
  Dim sPath:sPath=Server.MapPath("/")
  If G_Domain="" Then 
    GetMapPath=sPath
  Else
	  Dim I,Arr,L,P
	  Arr=Split(sPath,"\"):L=Ubound(Arr)
	  For i=0 To L
		If I<>L Then
			 If i=0 Then
			   p=Arr(i)
			 Else
			   p=p & "\" & Arr(i)
			 End If
		End If
	  Next
	  GetMapPath=p 
 End If
End Function

'**************************************************
'函数名：GetAutoDoMain()
'作  用：取得当前服务器IP 如：http://127.0.0.1
'参  数：无
'**************************************************
Public Function GetAutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetAutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			GetAutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(GetAutoDomain),"/W3SVC")<>0 Then
			   GetAutoDomain=Left(GetAutoDomain,Instr(GetAutoDomain,"/W3SVC"))
		 End If
		 GetAutoDomain = "http://" & GetAutoDomain & "/"
End Function
%>

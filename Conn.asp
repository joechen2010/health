<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
 

<%

Dim SqlNowString,DataPart_D,DataPart_Y,DataPart_H,DataPart_S
Dim Conn,DBPath,CollectDBPath,DataServer,DataUser,DataBaseName,DataBasePsw,ConnStr,CollcetConnStr
Const DataBaseType=0                   'ϵͳ���ݿ����ͣ�"1"ΪMS SQL2000���ݿ⣬"0"ΪMS ACCESS 2000���ݿ�
Const MsxmlVersion=".3.0"                'ϵͳ����XML�汾���� 

Const EnableSiteManageCode = True        '�Ƿ����ú�̨������֤�� �ǣ� True  �� False 
Const SiteManageCode = "8888"      '��̨������֤�룬���޸ģ�������ʹ����֪�������ĺ�̨�û���������Ҳ���ܵ�¼��̨

 
If DataBaseType=0 then
	'�����ACCESS���ݿ⣬�������޸ĺ���������ݿ���ļ���
	DBPath       = "/KS_Data/KesionCMS6.mdb"      'ACCESS���ݿ���ļ�������ʹ���������վ��Ŀ¼�ĵľ���·��
Else
	 '�����SQL���ݿ⣬�������޸ĺ��������ݿ�ѡ��
	 DataServer   = "(local)"                                  '���ݿ������IP
	 DataUser     = "sa"                                       '�������ݿ��û���
	 DataBaseName = "KesionCMS6"                                '���ݿ�����
	 DataBasePsw  = "989066"                                   '�������ݿ����� 
End if

 '�ɼ����ݿ�·��
 CollectDBPath="\KS_Data\Collect\KS_Collect.Mdb"

'=============================================================== ���´����벻Ҫ�����޸�========================================
Call OpenConn
Sub OpenConn()
    On Error Resume Next
    If DataBaseType = 1 Then
       ConnStr="Provider = Sqloledb; User ID = " & datauser & "; Password = " & databasepsw & "; Initial Catalog = " & databasename & "; Data Source = " & dataserver & ";"
	   SqlNowString = "getdate()"
	   DataPart_D   = "d"
	   DataPart_Y   = "y"
	   DataPart_H   = "hour"
	   DataPart_S   = "s"
    Else
       ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
	   SqlNowString = "Now()"
	   DataPart_D   = "'d'"
	   DataPart_Y   = "'yyyy'"
	   DataPart_H   = "'h'"
	   DataPart_S   = "'s'"
    End If
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open ConnStr
    If Err Then Err.Clear:Set conn = Nothing:Response.Write "���ݿ����ӳ�������Conn.asp�ļ��е����ݿ�������á�":Response.End
	CollcetConnStr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(CollectDBPath)
End Sub
Sub CloseConn()
    On Error Resume Next
	Conn.close:Set Conn=nothing
End sub

'==============================================ȫ����ʱ�����࿪ʼ==============================
Dim GCls:Set GCls=New GlobalVarCls
Class GlobalVarCls
    Public StaticPreList,StaticPreContent,StaticExtension
	Private Sub Class_Initialize()
	   StaticPreList    = "list"                 rem α��̬�б�ǰ׺ ���ܰ���"?"��"-"
	   staticPreContent = "thread"              rem α��̬����ǰ׺ 
	   StaticExtension  = ".html"                rem α��̬��չ��
	End Sub
    Private Sub Class_Terminate()
		 Set GCls=Nothing
	End Sub
	
	Public Function Execute(Command)
		If Not IsObject(Conn) Then OpenConn()
		On Error Resume Next
		Set Execute = Conn.Execute(Command)
		If Err Then
				Response.Write("��ѯ���Ϊ��" & Command & "<br>")
				Response.Write("������ϢΪ��" & Err.Description & "<br>")
			Err.Clear
			Set Execute = Nothing
			Response.End()
		End If
		'Sql_Use = Sql_Use + 1
	End Function
	
	Function GetUrl() 
		On Error Resume Next 
		Dim strTemp 
		If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
		 strTemp = "http://"
		Else 
		 strTemp = "https://"
		End If 
		strTemp = strTemp & Request.ServerVariables("SERVER_NAME") 
		If Request.ServerVariables("SERVER_PORT") <> 80 Then 
		 strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT") 
		end if
		strTemp = strTemp & Request.ServerVariables("URL") 
		If Trim(Request.QueryString) <> "" Then 
		 strTemp = strTemp & "?" & Trim(Request.QueryString) 
		end if
		GetUrl = strTemp 
	End Function

	'====================��־���õ�ַ================
	Public Property Let ComeUrl(ByVal strVar) 
			Session("M_ComeUrl") = strVar 
	End Property 
			
	Public Property Get ComeUrl
			ComeUrl= Session("M_ComeUrl")
	End Property 
	'================================================
End Class
'==============================================ȫ����ʱ���������==============================
%>

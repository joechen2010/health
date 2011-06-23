<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.Commoncls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

'是否允许自动检测最新版本 true 允许 false 不允许
Dim EnabledAutoUpdate:EnabledAutoUpdate=true 
'网站程序使用的编码,一定要设置正确,否则可能导致网站出现乱码
const Encoding="gb2312"
'官方远程文件版本地址
const Kesion_Version_XmlUrl="http://www.kesion.com/websystem/version.xml" 

'官方远程文件更新列表地址,必须/结束   
const Kesion_Update_FileUrl="http://www.kesion.com/websystem/"   


Dim SuccNum,ErrNum
Dim KS:Set KS=New PublicCls
Dim xmlObj : set xmlObj = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)

'普通管理员屏蔽自动升级
If KS.C("SuperTF")<>"1" And EnabledAutoUpdate=true Then EnabledAutoUpdate=false

select case Request("action")
  case "check"  checkIsNewestVersion
  case "update" updateProcess
end select

Set KS=Nothing
CloseConn

function checkIsNewestVersion()
  If EnabledAutoUpdate=false Then KS.Die "enabled"
  Dim LocalVersion,RemoteVersion
  on error resume next
  xmlObj.load(server.mappath("include/version.xml"))
  if isObject(xmlObj) then
    LocalVersion=xmlObj.getElementsByTagName("kesioncms/version")(0).Text
  end if
  if err.number<>0 then
    err.clear
    KS.Die "localversionerr"
  end if
 
  Dim CCls : Set CCls=New CollectPublicCls
  Dim XmlStr:XmlStr=CCls.GetHttpPage(Kesion_Version_XmlUrl,"gbk")
  Set CCls=Nothing
  xmlObj.loadXML(xmlstr)
  if isObject(xmlObj) then
    RemoteVersion=xmlObj.getElementsByTagName("kesioncms/version")(0).Text
  end if
  if err.number<>0 then
    err.clear
    KS.Die "remoteversionerr"
  end if 
  if RemoteVersion>LocalVersion Then
    If xmlObj.getElementsByTagName("kesioncms/allowupdateonline")(0).Text="false"Then   
	 KS.Echo "unallow"
	ElseIf KS.ChkClng(split(RemoteVersion,".")(0))>KS.ChkCLng(split(LocalVersion,".")(0)) Then '增加判断是不是同一版本号的
	 KS.Echo "unallowversion"
	Else
     KS.Echo Escape(xmlObj.getElementsByTagName("kesioncms/message")(0).Text)
	End If
  else
    KS.Echo "false"
  end if
end function

sub updateProcess()
  Dim FileList,FileLen,FileArr,I,Node,RemoteVersion
  on error resume next
  Dim CCls : Set CCls=New CollectPublicCls
  Dim XmlStr:XmlStr=CCls.GetHttpPage(Kesion_Version_XmlUrl,"gbk")
  Set CCls=Nothing
  xmlObj.loadXML(xmlstr)
  if isObject(xmlObj) then
     RemoteVersion=xmlObj.getElementsByTagName("kesioncms/version")(0).Text
     FileList=xmlObj.getElementsByTagName("kesioncms/filelist")(0).Text
  end if
  if err.number<>0 then
    err.clear
    KS.Die "remoteversionerr"
  end if 
  
  FileList=Replace(Replace(FileList," ",""),vbcrlf,"")
  FileArr=Split(FileList,",")
  FileLen=Ubound(FileArr)
  SuccNum=0 : ErrNum=0
  For I=0 To FileLen
    If getFileAndUpdate(trim(FileArr(i)),RemoteVersion) Then 
	  SuccNum=SuccNum+1
      KS.Echo Escape("<font color=green>" & Replace(FileArr(i),".txt",".asp") & "更新完毕!</font><br>")
	Else
	  ErrNum=ErrNum+1
      KS.Echo Escape("<font color=red>" & Replace(FileArr(i),".txt",".asp") & "没有成功更新!</font><br>")
	End If
  Next
  If ErrNum>0 Then
   KS.Echo Escape("<font color=blue>发现有文件没有成功更新,可能是您的部分目录没有写入权限,建议到官方网站下载补丁手工升级!</font>")
  Else
       '更新当前版本号
       Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	   Doc.async = false
	   Doc.setProperty "ServerHTTPRequest", true 
	   Doc.load(Server.MapPath("include/version.xml"))
	   doc.documentElement.selectSingleNode("//kesioncms/version").text=RemoteVersion
	   doc.save(Server.MapPath("include/version.xml"))
			
    KS.Echo Escape("恭喜,本次成功更新了" & SuccNum & "个文件!!!")
  End If
  KS.Echo Escape("<div style='text-align:center'><input type='button' style='height:25px;background:#efefef;border:1px solid #000' value=' 关 闭 ' onclick=""closeWindow()""></div>")
end sub

Function getFileAndUpdate(fileUrl,RemoteVersion)
  Dim CCls : Set CCls=New CollectPublicCls
  Dim BodyText:BodyText=CCls.GetHttpPage(Kesion_Update_FileUrl & Encoding & "/" & RemoteVersion & FileUrl,"gb2312")
  Set CCls=Nothing
  
  If BodyText<>"" Then
    FileUrl=Replace(FileUrl,".txt",".asp")
	If Left(FileUrl,1)="/" Then FileUrl=Right(FileUrl,Len(FileUrl)-1)
	FileUrl=KS.Setting(3) & replace(FileUrl,"admin/",KS.Setting(89))
    getFileAndUpdate=KS.WriteTOFile(FileUrl, BodyText)
  Else
    getFileAndUpdate=false
  End If
End Function

%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Server.ScriptTimeout=9999999

Dim KSCls
Set KSCls = New DB_BackUp
KSCls.Kesion()
Set KSCls = Nothing

Class DB_BackUp
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		
		 With KS
		If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			  .echo "<script>alert('�벻Ҫ���Ƿ��ύ��');history.back();</script>"
			Response.end
		 End If
		 
		 
		  .echo "<html>"
		  .echo "<head>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		  .echo "<title>�������ݿ�</title>"
		  .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		  .echo "<script src=""../ks_inc/jquery.js""></script>"
		if KS.G("Action")<>"ExecSql" then
		  .echo ("<body oncontextmenu=""return false;"" scroll=no>")
		  .echo "<ul id='menu_top'>"
		  .echo "<li class='parent' onclick=""location.href='?Action=BackUp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�������ݿ�</span></li>"
		  .echo "<li class='parent' onclick=""location.href='?Action=Restore';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>�ָ����ݿ�</span></li>"
		  .echo "<li class='parent' onclick=""location.href='?Action=Compact';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/verify.gif' border='0' align='absmiddle'>"
		If DataBaseType=1 Then
		   .echo "MSSQL���ݿ���־����"
		Else
		   .echo "ѹ���޸����ݿ�"
		End If
		  .echo "</span></li>"
		  .echo "</ul>"
	    elseif ks.g("flag")<>"Result" then
		  .echo ("<body oncontextmenu=""return false;"" scroll=no>")
		  .echo "<ul id='menu_top'>"
		  .echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		  .echo "  <tr>"
		  .echo "    <td height=""23"" align=""left"" valign=""top"">"
		  .echo "	<td align=""center""><strong>����ִ��SQL���</strong></td>"
		  .echo "    </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</ul>"
		end if
		   Select Case KS.G("Action") 
		    Case "BackUp"
			   If Not KS.ReturnPowerResult(0, "KMST10007") Then                '������Ա�����(���͸�)��Ȩ�޼��
		          Call KS.ReturnErr(1, "")
				  Response.End
			   Else
			     Call Db_BackUp()
			  End If
		   Case "Restore"
			   If Not KS.ReturnPowerResult(0, "KMST10007") Then                '���ָ����ݿ��Ȩ��
				  Call KS.ReturnErr(1, "")
				  Response.End
			   Else
			     Call Db_Restore()
			  End If
		   Case "Compact"
		     	If Not KS.ReturnPowerResult(0, "KMST10007") Then                '���ѹ�����ݿ��Ȩ��
				  Call KS.ReturnErr(1, "")
				  Response.End
				Else
				  Call Db_Compact()
			  End If
		   Case "ExecSql"
		     If Not KS.ReturnPowerResult(0, "KMST10009") Then                '�������ִ��SQL���
				  Call KS.ReturnErr(1, "")
			  Response.End
			  Else
			    Call Db_ExecSQL()
		      End If
		  End Select
		  End With
		End Sub

		
		'����
	  Sub Db_BackUp()
		With KS
		  .echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		  .echo "  <tr> "
		  .echo "    <td align=""center"" valign=""top"">"
		  .echo "       <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
		  .echo "        <tr> "
		  .echo "          <td height=""22""> "
		
		 Dim CurrPath,BackPath,tempArr,bkdbname    
		if request("Flag")="Backup" then
		   bkdbname=request.form("bkdbname")
           If DataBaseType=0 Then
			  CurrPath=request.form("Dbpath")
			  TempArr=replace(CurrPath,"/","\")
			  TempArr=split(TempArr,"\")
			  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
			  if KS.backupdata(CurrPath,BackPath & bkdbname)=true then
			     .echo "<div align=center><font color=green>ϵͳ�����ݿⱸ�ݳɹ�!</font></div><div align=center>���ݵ������ݿ�Ϊ:" & backpath & Bkdbname & "</div>"
			  Else
			    .echo ("<font color=red>����ʧ��!</font>")
			  End IF
		  Else
			  If Left(bkdbname,1)<>"/" and Left(bkdbname,1)<>"\" Then bkdbname="/" & bkdbname
			  CurrPath=bkdbname
			  TempArr=replace(CurrPath,"/","\")
			  TempArr=split(TempArr,"\")
			  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
			  KS.CreateListFolder BackPath
			  conn.execute   "backup database  " & DataBaseName &"  to  disk='"& Server.MapPath(bkdbname) &"'" 
			  on   error   resume   next   
			  If   err   Then   
				   .echo ("<font color=red>����ʧ��!</font>")
			  Else   
				   .echo "<div align=center><font color=green>ϵͳ�����ݿⱸ�ݳɹ�!</font></div><div align=center>���ݵ������ݿ�Ϊ:" & Bkdbname & "</div>"
			  End   If   
                   .echo "<script>$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle(false);</script>"
		  End If
		elseif request("Flag")="Backup1" then
		  CurrPath=request.form("Dbpath")
		  TempArr=replace(CurrPath,"/","\")
		  TempArr=split(TempArr,"\")
		  BackPath=Replace(CurrPath,TempArr(Ubound(TempArr)),"")
		  bkdbname=request.form("bkdbname")
		  if KS.backupdata(CurrPath,BackPath & bkdbname)=true then
		     .echo "<div align=center><font color=green>ϵͳ�ɼ����ݿⱸ�ݳɹ�!</font></div><div align=center>���ݵĲɼ����ݿ�Ϊ:" & backpath & Bkdbname & "</div>"
		  Else
		     .echo ("<font color=red>����ʧ��!</font>")
		 End IF
		end if
		
		  .echo "</td>"
		  .echo "        </tr>"
		  .echo "        <tr> "
		  .echo "		     <td> "
		
		if DataBaseType=0 then
				  .echo "              <fieldset>"
				  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup"">"
				  .echo "	<legend>ϵͳ�����ݿ�</legend>"
				  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""> ��ǰ���ݿ�·��</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" &server.mappath(DBPath) & """ readonly></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22""></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center"">�������ݿ�����[�������Ŀ¼���ڸ��ļ������ǣ������Զ�����]</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""Data(" & date() & ").bak""></td>"
				  .echo "                </tr>"
				  .echo "              </table>"
				  .echo "			  </fieldset>"
				  .echo "			  <table width=""100%"" border=""0"">"
				  .echo "			   <tr>"
				  .echo "			   <td height=""50"" align=center>"
				  .echo "			     <input type=submit value=""ȷ������"" class=""button"">"
				  .echo "			   </td>"
				  .echo "			   </tr>"
				  .echo "			   </form></table>"
		Else
				  .echo "              <fieldset>"
				  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup"">"
				  .echo "	<legend>ϵͳ�����ݿ�(MSSQL)</legend>"
				  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""> ��ǰ���ݿ�</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" & DataBaseName & """ readonly></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22""></td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center"">�������ݿ�����[�������Ŀ¼���ڸ��ļ������ǣ������Զ�����]</td>"
				  .echo "                </tr>"
				  .echo "                <tr> "
				  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""/KS_Data/SQL(" & date() & ").bak""></td>"
				  .echo "                </tr>"
				  .echo "              </table>"
				  .echo "			  </fieldset>"
				  .echo "			  <table width=""100%"" border=""0"">"
				  .echo "			   <tr>"
				  .echo "			   <td height=""50"" align=center>"
				  .echo "			     <input type=submit onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle(true);"" value=""ȷ������"" class=""button"">"
				  .echo "			   </td>"
				  .echo "			   </tr>"
				  .echo "			   </form></table>"
		end if
		
		  .echo "              <fieldset>"
		
		  .echo "          <form method=""post"" action=""?Action=BackUp&Flag=Backup1"">"
		  .echo "	<legend>ϵͳ�ɼ����ݿ�</legend>"
		  .echo "	<table width=""91%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""> ��ǰ���ݿ�·��</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=DBpath value=""" &server.mappath(CollectDBPath) & """ readonly></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22""></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center"">�������ݿ�����[�������Ŀ¼���ڸ��ļ������ǣ������Զ�����]</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""center""><input type=text size=50 name=bkDBname value=""Collect(" & date() & ").bak""></td>"
		  .echo "                </tr>"
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr>"
		  .echo "			   <td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""ȷ������"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "          </form>"
		  .echo "			   </table>"
		
		if DataBaseType=0 then			  
		  .echo "			  �����ݿ�����·��Ϊ��&nbsp;&nbsp;&nbsp;<font color=red>" & server.mappath(dbpath) & "</font><br>"
		end if
		  .echo "              �ɼ����ݿ�����·��Ϊ��<font color=red>" & server.mappath(CollectDBPath) & "</font><br></td>"
		  .echo "        </tr>"
		  .echo "      </table>"
		  .echo "     </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		 End With
		End Sub
		
		'�ָ�
		Sub Db_Restore()
		  With KS
			  .echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			  .echo "  <tr> "
			  .echo "    <td align=""center"" valign=""top""> <br> <strong><br>"
			  .echo "      </strong> <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
			  .echo "        <tr> "
			  .echo "          <td height=""25"" align=""center""> "
					
			if request("submit1")="�ָ�ѡ�еı����ļ�" then
				if Request.Form("backname")="0" then
				    .echo ("<script>alert('û�б����ļ�!');history.back();</script>")
				  Response.End
				end if
			   if  RestoreDatabase(Request.Form("backname"),Request("Flag"))=true then
				  if request("Flag")="main" then
				   .echo "<div align=center><font color=green>�����ɹ���</font></div><div align=center>�����ݿ��Ѵ�<font color=red>" & Request.Form("backname") & "</font>�����лָ�!</div>"
				  else
				   .echo "<div align=center><font color=green>�����ɹ���</font></div><div align=center>�ɼ����ݿ��Ѵ�<font color=red>" & Request.Form("backname") & "</font>�����лָ�!</div>"
				  end if
			   else
				   .echo "<font color=red>����ʧ��!</font>"
			   end if
			elseif request("submit1")="ɾ��ѡ�еı����ļ�" then
			   if Request.Form("backname")="0" then
				    .echo ("<script>alert('û�б����ļ�!');history.back();</script>")
				  Response.End
				end if
			  if  DeleteFile(Request.Form("backname"))=true then
				   .echo "<div align=center><font color=green>�����ɹ���</font></div><div align=center>�����ļ�<font color=red>" & Request.Form("backname") & "</font>��ɾ��!</div>"
			   else
				   .echo "<font color=red>����ʧ��!</font>"
			   end if
			end if  
			
			
			  .echo "</td>"
			  .echo "        </tr>"
			  .echo "        <tr> "
					 
			  .echo "           <td>"
			
			if DataBaseType=0 then
			  .echo "          <form method=""post"" name=""restoreform"" action=""KS.Database.asp?Action=Restore&Flag=main"">"
			  .echo "<br> <fieldset>"
			  .echo "	<legend>�����ݿ��ѱ��ݵ��ļ�</legend>"
						
						
			  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			  .echo "                  <td height=""22"" align=""center""><strong>ѡ�񱸷��ļ���</strong>"
			
								dim  tempStr,strCurDir,CurrDataBase,CurrLdb,Fso,Dir,s
								dim havebackfile:havebackfile=false
								tempStr=replace(dbpath,"/","\")
								tempStr=split(tempStr,"\")
								strCurDir=replace(dbpath,tempStr(ubound(tempStr)),"")
								strCurDir=server.mappath(strCurDir)
								
								 CurrDataBase=tempStr(ubound(tempStr))
								 CurrLdb=left(CurrDataBase,len(CurrDataBase)-4) & ".ldb"
								
			  .echo "				    <select name=""backname"">"
								 
							  set fso = KS.InitialObject(KS.Setting(99))
							  set dir = fso.GetFolder(strCurDir)
							  for each s in dir.Files
								 if s.name<>CurrDataBase and s.name<>Currldb and lcase(right(s.name,4))<>".dat" then
								  havebackfile=true
								  .echo "<option value=""" & strCurDir &"\" & s.name & """>" & s.name & "</option>"
								end if
							  next
							  if havebackfile=false then
							     .echo "<option value=""0"">---��û�б��ݵ������ݿ��ļ�---</option>"
							   end if
			  .echo "		            </select></td>"
			  .echo "               </tr>"
			  .echo "              </table>"
	
			  .echo "			  </fieldset>"
			  .echo "			  <table width=""100%"" border=""0"">"
			  .echo "			   <tr><td height=""50"" align=center>"
			  .echo "			     <input type=""submit"" name=""submit1"" "
			if havebackfile=false then   .echo "disabled"
			  .echo " value=""�ָ�ѡ�еı����ļ�"" class=""button"" onclick=""return(confirm('ȷ���ָ����ݿ��𣿴˲���������'))"">"
			  .echo "			     <input name=""submit1"" type=""submit"" "
			if havebackfile=false then   .echo " disabled" 
			  .echo " value=""ɾ��ѡ�еı����ļ�"" class=""button"" onclick=""return(confirm('ȷ��ɾ��ѡ�еı����ļ��𣿴˲���������'))""/>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			  .echo "          </form>"
			  .echo "		    </td>"
			  .echo "        </tr>"
			  .echo "      </table>"
			  .echo "     </td>"
			  .echo "  </tr>"
			  .echo "</table>"
			
			end if
			  .echo "             <br><table width=""560"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			
			  .echo "          <form method=""post"" name=""restoreform"" action=""KS.Database.asp?Action=Restore&Flag=collect"">"
					 
			  .echo "          <td><fieldset>"
			  .echo "	<legend>�ɼ����ݿ��ѱ��ݵ��ļ�</legend>"
						
						
			  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
			  .echo "                <tr> "
			  .echo "                  <td height=""22"" align=""center""><strong>ѡ�񱸷��ļ���</strong>"
			
								
								havebackfile=false
								tempStr=replace(CollectDBPath,"/","\")
								tempStr=split(tempStr,"\")
								strCurDir=replace(CollectDBPath,tempStr(ubound(tempStr)),"")
								strCurDir=server.mappath(strCurDir)
								
								 CurrDataBase=tempStr(ubound(tempStr))
								 CurrLdb=left(CurrDataBase,len(CurrDataBase)-4) & ".ldb"
								
			  .echo "				    <select name=""backname"">"
								 
							  set fso = KS.InitialObject(KS.Setting(99))
							  set dir = fso.GetFolder(strCurDir)
							  for each s in dir.Files
								 if s.name<>CurrDataBase and s.name<>Currldb then
								  havebackfile=true
								  .echo "<option value=""" & strCurDir &"\" & s.name & """>" & s.name & "</option>"
								end if
							  next
							  if havebackfile=false then
							     .echo "<option value=""0"">---��û�б��ݵĲɼ����ݿ��ļ�---</option>"
							   end if
			  .echo "		            </select></td>"
			  .echo "               </tr>"
			  .echo "              </table>"
			  .echo "			  </fieldset>"
			
			  .echo "			  <table width=""100%"" border=""0"">"
			  .echo "			   <tr><td height=""50"" align=center>"
			  .echo "			     <input type=""submit"" name=""submit1"" "
			if havebackfile=false then   .echo "disabled"
			  .echo " value=""�ָ�ѡ�еı����ļ�"" class=""button"" onclick=""return(confirm('ȷ���ָ����ݿ��𣿴˲���������'))"">"
			  .echo "			     <input name=""submit1"" type=""submit"" "
			if havebackfile=false then   .echo " disabled" 
			  .echo " value=""ɾ��ѡ�еı����ļ�"" class=""button"" onclick=""return(confirm('ȷ��ɾ��ѡ�еı����ļ��𣿴˲���������'))""/>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			
			  .echo "          </form>"
			  .echo "			   </td>"
			  .echo "			   </tr>"
			  .echo "		      </table>"
			  .echo "</body>"
			  .echo "</html>"
			End With
		End Sub
		' �ָ����ݿ�
		Public Function RestoreDatabase(BackName,Flag)
				dim fso,sFileName
				RestoreDatabase=false
				on error resume next
				set fso = KS.InitialObject(KS.Setting(99))
				IF Flag="main" Then  '�����ݿ�
				  sFileName = DbPath
				  Conn.Close
				  fso.CopyFile BackName, server.mappath(DbPath), True
				  if err then
					RestoreDatabase=false
				  else
					RestoreDatabase=true
				  end if
				  conn.Open ConnStr
				Elseif Flag="collect" Then  '�ɼ����ݿ�
				 
				 sFileName = CollectDBPath
				  fso.CopyFile BackName, server.mappath(CollectDBPath), True
				  if err then
					RestoreDatabase=false
				  else
					RestoreDatabase=true
				  end if
				Else 
				  RestoreDatabase=false
				  Exit Function
				End IF
			   IF err Then
				RestoreDatabase=false
			   End IF
			End Function
		
		'**************************************************
		'��������DeleteFile
		'��  �ã�ɾ��ָ���ļ�
		'��  ����FileStrҪɾ�����ļ�
		'����ֵ���ɹ�����true ���򷵻�Flase
		'**************************************************
		Function DeleteFile(FileStr)
		   Dim fso
		   On Error Resume Next
		   Set fso = KS.InitialObject(KS.Setting(99))
			If fso.FileExists(FileStr) Then
				fso.DeleteFile FileStr, True
			Else
			DeleteFile = True
			End If
		   Set fso = Nothing
		   If Err.Number <> 0 Then
		   Err.Clear
		   DeleteFile = False
		   Else
		   DeleteFile = True
		   End If
		End Function


       '�ָ�
	   Sub Db_Compact()
	    With KS
		  .echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		  .echo "  <tr> "
		  .echo "    <td align=""center"" valign=""top""> <br> <strong><br>"
		  .echo "      </strong> <table width=""560"" border=""0"" cellpadding=""2"" cellspacing=""1"">"
		  .echo "        <tr> "
		  .echo "          <td height=""25"" align=""center""> "
				  
		if request("Flag")="Backup" then
		  If DataBaseType=0 Then
		   if  CompactDatabase(DBPath,ConnStr)=true then
			   .echo "<font color=green>�����ݿ�ѹ�����޸��ɹ�!</font>"
		   else
			   .echo "<font color=red>����ʧ��!</font>"
		   end if
		  Else
			 conn.execute("DUMP TRANSACTION " & DataBaseName & " WITH  NO_LOG")
			 conn.execute("DBCC SHRINKDATABASE(" & DataBaseName & ")")
			  on   error   resume   next   
			  If   err   Then   
				   .echo ("<font color=red>����ʧ��!</font>")
			  Else   
				   .echo "<div align=center><font color=green>����mssql���ݿ���־�����!</font></div>"
			  End   If   
                 .echo "<script>$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle(false);</script>"
		  End If
		elseif Request("Flag")="Backup1" then
		   if  CompactCollectDatabase(CollectDBPath,CollcetConnStr)=true then
			   .echo "<font color=green>�ɼ����ݿ�ѹ�����޸��ɹ�!</font>"
		   else
			   .echo "<font color=red>����ʧ��!</font>"
		   end if
		end if
		
		  .echo "</td>"
		  .echo "        </tr>"
		
		if DataBaseType=0 then
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup"">"
				 
		  .echo "            <td> <fieldset>"
		  .echo "	<legend>�����ݿ���Ϣ</legend>"
					
					dim filesize:filesize=KS.GetFieSize(server.mappath(DBPath))
					dim ReclaimedSpace:ReclaimedSpace=CLng(conn.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value)
					dim LocaleIdentifier:LocaleIdentifier=Conn.Properties("Locale Identifier").Value
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td width=""23%"" height=""22"" align=""right""><strong>���ݿ�·����</strong></td>"
		  .echo "                  <td width=""77%""><font color=#ff6600>" & server.mappath(DBPath) & "</font></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>ѹ��ǰ��С��</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize, 0, False, False, True) & " �ֽ�</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>ѹ�����С��</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize - ReclaimedSpace, 0, False, False, True) & " �ֽ� (�ܼƿ��Լ���" & FormatNumber(ReclaimedSpace, 0, True, False, True)& " �ֽ�)</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>������ʶ����</strong></td>"
		  .echo "                  <td height=""22"">" & GetLocaleName(LocaleIdentifier) & "</td>"
		  .echo "                </tr>"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""��ʼѹ��"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"
		Else
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup"">"
				 
		  .echo "            <td> <fieldset>"
		  .echo "	<legend>�����ݿ���Ϣ(MSSQL)</legend>"
					
		
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		on error resume next
		Dim RS,I
		Set RS=Conn.Execute("select name, convert(float,size) * (8192.0/1024.0)/1024. from [kesioncms5].dbo.sysfiles")			
		For I=0 To 1
			  .echo "                <tr> "
			  .echo "                  <td height=""22""><strong>�ļ�" & RS(0) & "��С��</strong>" & rs(1) & " MB</td>"
			  .echo "                </tr>"
		 RS.MoveNext
	    Next
	   RS.Close:Set RS=Nothing
		if err then KS.AlertHintScript "�Բ���,���ķ�������֧�ִ˲���!"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""��ʼ������־"" onclick=""$(parent.frames['FrameTop'].document).find('#ajaxmsg').toggle(true);"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"		
		end if
		  .echo "        <tr> "
		  .echo "          <form method=""post"" action=""?Action=Compact&Flag=Backup1"">"
				 
		  .echo "            <td><br> <fieldset>"
		  .echo "	<legend>�ɼ����ݿ���Ϣ</legend>"
		
						 conn.close
						Set conn = KS.InitialObject("ADODB.Connection")
						conn.open CollcetConnStr
		
					filesize=KS.GetFieSize(server.mappath(CollectDBPath))
					ReclaimedSpace=CLng(conn.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value)
					LocaleIdentifier=Conn.Properties("Locale Identifier").Value
		  .echo "              <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""2"">"
		  .echo "                <tr> "
		  .echo "                  <td width=""23%"" height=""22"" align=""right""><strong>���ݿ�·����</strong></td>"
		  .echo "                  <td width=""77%""><font color=#ff6600>" & server.mappath(CollectDBPath) & "</font></td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>ѹ��ǰ��С��</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize, 0, False, False, True) & " �ֽ�</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>ѹ�����С��</strong></td>"
		  .echo "                  <td height=""22"">" & FormatNumber(filesize - ReclaimedSpace, 0, False, False, True) & " �ֽ� (�ܼƿ��Լ���" & FormatNumber(ReclaimedSpace, 0, True, False, True)& " �ֽ�)</td>"
		  .echo "                </tr>"
		  .echo "                <tr> "
		  .echo "                  <td height=""22"" align=""right""><strong>������ʶ����</strong></td>"
		  .echo "                  <td height=""22"">" & GetLocaleName(LocaleIdentifier) & "</td>"
		  .echo "                </tr>"
					   
		  .echo "              </table>"
		  .echo "			  </fieldset>"
		  .echo "			  <table width=""100%"" border=""0"">"
		  .echo "			   <tr><td height=""50"" align=center>"
		  .echo "			     <input type=submit value=""��ʼѹ��"" class=""button"">"
		  .echo "			   </td>"
		  .echo "			   </tr>"
		  .echo "			   </table>"
		  .echo "			  </td>"
		  .echo "          </form>"
		  .echo "        </tr>"
		
		
		  .echo "      </table>"
		  .echo "	  ˵�������ⲻ��Ԥ��Ĵ�����������ѹ��֮ǰ����ԭʼ���ݿ⣡"
		  .echo "     </td>"
		  .echo "  </tr>"
		  .echo "</table>"
		  .echo "</body>"
		  .echo "</html>"
		End With
		End Sub
		
		'**********************************************************************
		'��������CompactDatabase
		'���ã�ѹ�������ݿ�
		'������DBPath--���ݿ�λ��,ConnStr---���ݿ������ַ���
		'**********************************************************************   
		 Public Function CompactDatabase(DBPath, ConnStr)
				On Error Resume Next
				Dim strTempFile, fso, jro, ver, strCon, strTo, LCID
				Set fso = KS.InitialObject(KS.Setting(99))
				strTempFile = DBPath
				strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
				Set jro = KS.InitialObject("JRO.JetEngine")
				LCID = Conn.Properties("Locale Identifier").Value
				'�ر����ݿ�
				Conn.Close
				strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & Server.MapPath(strTempFile) & "; Jet OLEDB:Engine Type=" & ver
				
				jro.CompactDatabase ConnStr, strTo
				CompactDatabase = False
				If Err Then
					fso.DeleteFile Server.MapPath(strTempFile)
				Else
					fso.DeleteFile Server.MapPath(DBPath)
					fso.MoveFile Server.MapPath(strTempFile), Server.MapPath(DBPath)
					If Err Then
						fso.DeleteFile Server.MapPath(strTempFile)
					Else
						CompactDatabase = True
					End If
				End If
				Set jro = Nothing
				Set fso = Nothing
				'���´����ݿ�
				Conn.Open ConnStr
		End Function
		'**********************************************************************
		'��������CompactDatabase
		'���ã�ѹ���ɼ����ݿ�
		'������DBPath--���ݿ�λ��,ConnStr---���ݿ������ַ���
		'**********************************************************************   
		 Public Function CompactCollectDatabase(DBPath, ConnStr)
				On Error Resume Next
				Dim strTempFile, fso, jro, ver, strCon, strTo, LCID
				
				Set conn = KS.InitialObject("ADODB.Connection")
				conn.open CollcetConnStr
				
				Set fso = KS.InitialObject(KS.Setting(99))
				strTempFile = DBPath
				strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
				Set jro = KS.InitialObject("JRO.JetEngine")
				LCID = Conn.Properties("Locale Identifier").Value
				'�ر����ݿ�
				Conn.Close
				strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & Server.MapPath(strTempFile) & "; Jet OLEDB:Engine Type=" & ver
				
				jro.CompactDatabase ConnStr, strTo
				CompactCollectDatabase = False
				If Err Then
					fso.DeleteFile Server.MapPath(strTempFile)
				Else
					fso.DeleteFile Server.MapPath(DBPath)
					fso.MoveFile Server.MapPath(strTempFile), Server.MapPath(DBPath)
					If Err Then
						fso.DeleteFile Server.MapPath(strTempFile)
					Else
						CompactCollectDatabase = True
					End If
				End If
				Set jro = Nothing
				Set fso = Nothing
				'���´����ݿ�
				Conn.Open ConnStr
		End Function
		
		'�õ����ݿ�ĵ�����ʶ��	
		Function GetLocaleName(lcid)
				Select Case lcid
					Case 1033	GetLocaleName = "����"
					Case 2052	GetLocaleName = "���ı��"
					Case 133124	GetLocaleName = "���ıʻ�"
					Case 1028	GetLocaleName = "���ıʻ�(̨��)"
					Case 197636	GetLocaleName = "����ƴ��(̨��)"
					Case 1050	GetLocaleName = "���޵�����"
					Case 1029	GetLocaleName = "�ݿ���"
					Case 1061	GetLocaleName = "��ɳ������"
					Case 1036	GetLocaleName = "����"
					Case 66615	GetLocaleName = "��³������(�ִ�)"
					Case 66567	GetLocaleName = "����(�绰��)"
					Case 1038	GetLocaleName = "��������"
					Case 66574	GetLocaleName = "��������(��������)"
					Case 1039	GetLocaleName = "������"
					Case 1041	GetLocaleName = "����"
					Case 66577	GetLocaleName = "����(Unicode)"
					Case 1042	GetLocaleName = "����"
					Case 66578	GetLocaleName = "����(Unicode)"
					Case 1062	GetLocaleName = "����ά����"
					Case 1036	GetLocaleName = "��������"
					Case 1071	GetLocaleName = "FYRO �������"
					Case 1044	GetLocaleName = "Ų����/������"
					Case 1045	GetLocaleName = "������"
					Case 1048	GetLocaleName = "����������"
					Case 1051	GetLocaleName = "˹�工����"
					Case 1060	GetLocaleName = "˹����������"
					Case 1034	GetLocaleName = "��������(��ͳ)"
					Case 3082	GetLocaleName = "��������(������)"
					Case 1053	GetLocaleName = "�����/������"
					Case 1054	GetLocaleName = "̩����"
					Case 1055	GetLocaleName = "��������"
					Case 1058	GetLocaleName = "�ڿ�����"
					Case 1066	GetLocaleName = "Խ����"
					Case Else	GetLocaleName = "δ֪"
				End Select
			End Function

       '����ִ��SQL
	   Sub Db_ExecSQL()
	   With KS
		  .echo "<html>"
		  .echo "<head>"
		  .echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		  .echo "<title>����ִ��SQL���</title>"
		  .echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		  .echo "<script src=""../ks_inc/jquery.js""></script>"
		Dim Flag:Flag=KS.G("Flag")
		IF Flag="Result" Then 
		  .echo ("<body style=""margin:1;"">")
		 Call ExeSQL
		Else
		  .echo ("<body scroll=no>")
    %>
		<script language="javascript">
	<!--
	 function CheckForm()
	 {
	 if ($('textarea[name=Sql]').val()=='')
	  {
	  alert('������SQL��ѯ��䣡');
	  $('textarea[name=Sql]').focus();
	  return false;
	  }
	  //alert(escape($F('Sql')));
	  ExeSQLFrame.location.href="KS.Database.asp?Action=ExecSql&Flag=Result&SQL="+escape($('textarea[name=Sql]').val().replace('+','ksaddks'));
	  return false;
	  }
	-->
	</script>
	<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<form name="SqlForm" method="post" Action="?Action=ExecSql" onsubmit="return CheckForm()">
	<tr height="50">
	  <td>
	  <textarea name="Sql" rows="5" wrap="OFF" style="width:100%;"></textarea>
	  <input type="hidden" name="Flag" value="Exec">
	  </td>
	</tr>
	<tr height="25">
	 <td align="center">
	  <input type="submit" name="submit1" class="button" value="����ִ��"><span style="color:red">һ�ο���ִ�ж���SQL��䣬����������ûس����и����������û��һ����SQL���������鲻Ҫʹ�ã�</span>
	  </td>
	</tr>
	</form>
	  <tr> 
		<td valign="_top"><iframe id="ExeSQLFrame" scrolling="auto" src="KS.Database.asp?Action=ExecSql&Flag=Result" style="width:100%;height:93%" frameborder=1></iframe></td>
	  </tr>
	</table>
	<% End iF%>
	</BODY>
	</HTML>
<% End With
  End Sub
  Sub ExeSQL()
        Dim SelectSQLTF,ExecSQLErrorTF,ExeResultNum,ExeResult,FiledObj,i
		Dim Sql:Sql =replace(request.querystring("Sql"),"ksaddks","+")
	    if SQL="" Then Exit Sub
		sql=split(sql,vbcrlf)
		For I=0 To Ubound(sql)
		  if (Sql(i)<>"") Then
				If Instr(1,lcase(Sql(i)),"delete from ks_log")<>0 then
					Call KS.AlertHistory("�Բ��𣬲���ɾ����־�����ݣ�",-1)
						Exit Sub
				End If
				SelectSQLTF = (LCase(Left(Trim(Sql(i)),6)) = "select")
				Conn.Errors.Clear
				On Error Resume Next
				if SelectSQLTF = True then
					  Set ExeResult = Conn.Execute(Sql(i),ExeResultNum)
				else
					  Conn.Execute Sql(i),ExeResultNum
				end if
				 
				If Conn.Errors.Count<>0 Then
					  ExecSQLErrorTF = True
					  Set ExeResult = Conn.Errors
				Else
					  ExecSQLErrorTF = False
				End If
				if ExecSQLErrorTF = True then
				%>
				<table width="100%" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
				  <tr bgcolor="F4F4EA"> 
					<td height="20" nowrap> 
					  <div align="center">�����</div></td>
					<td height="20" nowrap> 
					  <div align="center">��Դ</div></td>
					<td height="20" nowrap> 
					  <div align="center">����</div></td>
					<td height="20" nowrap> 
					  <div align="center">����</div></td>
					<td height="20" nowrap> 
					  <div align="center">�����ĵ�</div></td>
				  </tr>
				  <tr height="20" bgcolor="#FFFFFF"> 
					<td nowrap> 
					  <% = Err.Number %> </td>
					<td nowrap> 
					  <% = Err.Description %> </td>
					<td nowrap> 
					  <% = Err.Source %> </td>
					<td nowrap> 
					  <% = Err.Helpcontext %> </td>
					<td nowrap> 
					  <% = Err.HelpFile %> </td>
				  </tr>
				</table>
				<%
				else
				%>
				<table border="0" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
				  <%
					if SelectSQLTF = True then
				%>
				  <tr>
				<%
						For Each FiledObj In ExeResult.Fields
				%>
					<td nowrap bgcolor="F4F4EA" height="26"><div align="center">
						<% = FiledObj.name %>
					  </div></td>
				<%
						next
				%>
				  </tr>
				<%
						do while Not ExeResult.Eof
				%>
				  <tr height="20" nowrap bgcolor="#ffffff" onMouseOver="this.style.background='#F5f5f5'" onMouseOut="this.style.background='#FFFFFF'">
				<%
							For Each FiledObj In ExeResult.Fields
				%>
					<td> 
					  <div align="center">
						<%
						 if IsNull(FiledObj.value) then
							KS.echo("&nbsp;")
						 else
							KS.Echo (FiledObj.value)
						 end if
						 %>
					  </div></td>
				<%
							next
				%>
				  </tr>
				<%
							ExeResult.MoveNext
						loop
					else
				%>
				  <tr>
					<td bgcolor="F4F4EA" height="26">
				<div align="center">ִ�н��</div></td>
				  </tr>
				  <tr>
					<td height="20" bgcolor="#FFFFFF">
				<div align="center">
						<% = ExeResultNum & "����¼��Ӱ��"%>
					  </div></td>
				  </tr>
				<%
					end if
				%>
				</table>
				<%
				  end if
			 end if
		  Next
		 End Sub
End Class
%> 

<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KS:Set KS=New PublicCls
Dim ID:ID=KS.ChkClng(KS.S("ID"))
Dim Action:Action=KS.S("Action")
if check=false then response.write "error":response.end
Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
If KS.S("Type")="0" Then
 RS.Open "Select Content From KS_GuestReply Where ID=" & ID,conn,1,3
Else
 RS.Open "Select [Memo],Subject From KS_GuestBook Where ID=" & ID,conn,1,3
End If
If RS.Eof And RS.Bof Then
 RS.Close:Set RS=Nothing
 Response.End
End If
If Action="show" Then
%>
<table width="100%" border="0">
<%If KS.S("Type")="1" Then%>
  <tr>
    <td align='right'><strong>标题:</strong></td>
    <td><input type='text' name='etitle' id='etitle' size='50' value="<%=rs(1)%>"></td>
  </tr>
<%end if%>
  <tr>
    <td align='right'><strong>内容:</strong></td>
    <td><textarea  ID='Content<%=ID%>' name='Content<%=ID%>' cols=90 rows=6 style='display:none'><%=RS(0)%></textarea><iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content<%=ID%>&amp;Toolbar=Basic" width="90%" height="180" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>

<div align="center"><input type="button" onClick="EditSave(<%=id%>)" style='padding:4px' value="确定修改" name="submit1"></div>
<%ElseIf Action="save" then
 Dim Content:Content=KS.HTMLEncode(Request("Content"))
 RS(0)=Content
 If KS.S("Type")="1" and KS.S("title")<>"" Then
 RS(1)=KS.S("title")
 End If
 RS.Update
 	'关联上传文件
 If KS.S("Type")="0" Then
	Call KS.FileAssociation(1036,ID,Content,1)
 Else
	Call KS.FileAssociation(1035,ID,Content,1)
 End If
 
 Response.Write KS.Htmlcode(content)
End If
RS.Close:Set RS=Nothing

function check()
	 	Dim KSLoginCls
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
		    master=LFCls.GetSingleFieldValue("select master from ks_guestboard where id=" & KS.ChkClng(FCls.RefreshFolderID))
			Dim KSUser:Set KSUser=New UserCls
			If Cbool(KSUser.UserLoginChecked)=false Then 
			  check=false
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			   if check=false then
			        '检查是不是用户自己的发帖
			     	dim rs:set rs=server.CreateObject("adodb.recordset")
					If KS.S("Type")="0" Then
					 RS.Open "Select top 1 username From KS_GuestReply Where ID=" & ID & " and username='" & KS.C("UserName") & "'",conn,1,1
					Else
					 RS.Open "Select top 1 username From KS_GuestBook Where ID=" & ID & " and username='" & KS.C("UserName") & "'",conn,1,1
					End If
					if rs.eof then
					  check=false
					else
					  check=true
					end if
					rs.close
					set rs=nothing
                   
			   end if
			End If
		end if
		

 End function			

%>
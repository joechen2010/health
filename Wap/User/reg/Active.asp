<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Md5.asp"-->
<%
Dim KSCls
Set KSCls = New UserReg
KSCls.Kesion()
Set KSCls = Nothing

Class UserReg
        Private KS
		Private FileContent,Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    IF KS.Setting(21)=0 Then
			   Call KS.ShowError("暂停注册！","对不起，本站暂停新会员注册！")
		    End IF
		    Response.Write "<wml>" &vbcrlf
			Response.Write "<head>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Response.Write "</head>" &vbcrlf
			Response.Write "<card id=""main"" title=""会员验证激活"">" &vbcrlf
			Response.Write "<p align=""left"">" &vbcrlf
			
			If KS.S("Action")="Check" Then 
			 Call Check()
			Else
			%>
            用户名称<br/>
            <input name="UserName" type="text" maxlength="<%=KS.Setting(30)%>" value="<%=KS.G("UserName")%>" emptyok="false"/><br/>
            验证码<br/>
            <input name="CheckNum" type="text" maxlength="30" value="<%=KS.G("CheckNum")%>" emptyok="false"/><br/>
            
            <anchor>马上验证激活<go href="?action=Check" method="post">
            <postfield name="UserName" value="$(UserName)"/>
            <postfield name="CheckNum" value="$(CheckNum)"/>
            </go>
            </anchor>
			<br/>
			<%
			End If
			Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a>"
			Response.Write "</p>" &vbcrlf
			Response.Write "</card>" &vbcrlf
			Response.Write "</wml>"
		End Sub

        
		'会员类型
		Function Check()
		    Dim UserName:UserName=trim(KS.S("UserName"))
			Dim CheckNum:CheckNum=trim(KS.S("CheckNum"))
			If UserName="" Or CheckNum="" Then
			  Call KS.ShowError("错误提示","用户名及验证都必须输入！")
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 CheckNum,locked,wap From KS_User Where UserName='" & UserName & "'",Conn,1,3
			If RS.Eof And RS.Bof Then
			 response.write "对不起，您输入的用户名不存在！<br/><anchor>返回上一页<prev/></anchor><br/>"
			else
			  if rs("checknum")<>checknum then
			   response.write "激活码有误，请重新输入！<br/><anchor>返回上一页<prev/></anchor><br/>"
			  else
			   rs("locked")=0
			   dim wp:wp=MD5(KS.MakeRandomChar(20),32)
			   RS("wap")= wp
			   RS.Update
			   response.write "恭喜您,账号激活成功,您现在可以正常登录了！<br/>"
			   %>
			    <anchor>现在登录<go href="../index.asp?wap=<%=wp%>" method="get">
					</go>
					</anchor>
                 <br/>
			   <%
			  end if
			end if
			rs.close:set rs=nothing
		End Function
End Class
%>

 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New JobManage
KSCls.Kesion()
Set KSCls = Nothing

Class JobManage
        Private KS,KSUser
		Private Descript,OrderID
		Private ComeUrl
		Private totalPut,currentpage,MaxPerPage
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
   
       If KS.S("Action")<>"View" Then
		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
	        <li<%if KS.S("status")="" then response.write " class='select'"%>>投诉/意见管理</li>
			</ul>
	    </div>
	 <%
	 End If
	 	Call KSUser.InnerLocation("投诉管理")
	 	KSUser.CheckPowerAndDie("s17")

	 
	 	Select Case KS.S("Action")
			  Case "Show" Call View()
			  case "del" call FeedBackDel()
			  case "Add" call Add()
			  case "DoSave" call Addsave()
		      Case Else  Call JobList()
		End Select					  
	End Sub
	
	Sub JobList()
      %>
	  <script language=javascript>
		function selectall(chkval) {
			with($('myform')) {
				for (var i=0;i<elements.length;i++) {
					if (elements[i].type=="checkbox") 
						elements[i].checked= chkval;
						
				}
			}
		}	
		function chkselect() {
			var selnum=0
			with($('myform')) {
				for (var i=0;i<elements.length;i++) {
					if (elements[i].type=="checkbox") 
						if(elements[i].checked==true)
							selnum++;
						
				}
				if(selnum==0) {
					alert("请选择要操作的投诉记录！");
					return false;
				}
			}
			
			return true;
		}
		function view(id)
		{
		var phx=window.open('../job/showtraining.asp?id='+id,'new','width=560,height=420,resizable=no,scrollbars=yes,left=280,top=150');
         phx.moveTo((screen.width-560)/2,(screen.height-420)/2);
		}
		</script>
		 <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="?Action=Add"><span style="font-size:14px;color:#ff3300">我要投诉</span></a></div>
	   <table width="98%"  border="0" align="center" cellpadding="1" cellspacing="1">
        <tr align="center" class="title">
			 <td width="5%" height="28" align="center"><strong>编号</strong></td>
			 <td><strong>主题</strong></td>
			 <td align="center"><strong>对象</strong></td>
			 <td width="10%" align="center"><strong>投诉时间</strong></td>
			 <td width="10%" align="center"><strong>处理人</strong></td>
			 <td width="12%" align="center"><strong>处理时间</strong></td>
			 <td width="10%" align="center"><strong>状态</strong></td>
			 <td><strong>操作</strong></td>
         </tr>
		   <%
							MaxPerPage=10
							If KS.S("page") <> "" Then
							   CurrentPage = KS.ChkClng(KS.S("page"))
							Else
							  CurrentPage = 1
							End If
							 Dim Param:Param=" where UserName='" & KSUser.UserName & "'"
							  
							  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_FeedBack " & Param & " order By ID",conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td colspan='10'class='splittd' align='center' colspan=2 height=30 valign=top>您没有发表任意见或投诉!</td></tr>"
								 Else
									totalPut = RS.RecordCount
									If CurrentPage < 1 Then	CurrentPage = 1
									If (CurrentPage - 1) * MaxPerPage > totalPut Then
										If (totalPut Mod MaxPerPage) = 0 Then
											CurrentPage = totalPut \ MaxPerPage
										Else
											CurrentPage = totalPut \ MaxPerPage + 1
										End If
									End If
			
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
							        Call showJobList(RS)
				End If
     %>                 
            </table>
						<div style="margin:8px">
						<strong>操作说明</strong><br /><font color='#777777'>
投诉/意见管理放置的是您投诉及对本站的建议记录；<br>您可以删除未处理的记录。
</font>
					</div>
	  <%
	End Sub
	
	Sub showJobList(RS)
	  Dim str,i
	  Do While Not RS.Eof
	      dim bh:bh=rs("id")
		  IF LEN(BH)=1 THEN 
			  BH="00"& bh
		  ElseIf LEN(BH)=2 Then
			  Bh="0" & bh
		  End If
		  bh="YJ" & year(rs("adddate")) & month(rs("adddate")) & bh
          response.write "<tr bgcolor=#ffffff>"
          Response.Write "<td height='30' class='splittd' align='center'>" & bh & "</td>"
          Response.Write "<td class='splittd' align='center'>" 
		  
		   response.write rs("title")
		  response.write "</td>"
          Response.Write "<td class='splittd' align='center'>" & rs("object") & "</td>"
          Response.Write "<td class='splittd' align='center'>" & formatdatetime(rs("adddate"),2) & "</td>"
		  
          Response.Write "<td class='splittd' align='center'>"
		  Dim AcceptTime,Delstr,strs
		  if rs("Accepted")="" or isnull(rs("accepted")) then
		   response.write "未处理"
		   AcceptTime="---"
		   Delstr="<a onclick=""return(confirm('确定删除吗?'))"" href='?action=del&id=" & rs("id") & "'>删除</a>"
		   strs="<font color=red>待受理</font>"
		  else
		   response.write rs("Accepted")
		   AcceptTime=RS("AcceptTime")
		   Delstr="<a href='#' disabled>删除</a>"
		   strs="<font color=green>已受理</font>"
		  end if
		  response.write "</td>"
          Response.Write "<td class='splittd' align='center'>" & AcceptTime & "</td>"
          Response.Write "<td class='splittd' align='center'>" & strs & "</td>"
          Response.Write "<td class='splittd' align='center'><a href='?action=Show&id=" & rs("id") & "'>查看详情</a>  " & delstr & "</td>"

           Response.Write "</tr>"
	   
	  	RS.MoveNext
		I = I + 1
		If I >= MaxPerPage Then Exit Do
	 Loop
	 response.write str
	 %>
	 						 </form>
								 <tr>
								  <td align="right">
								  
								  <%=KS.ShowPagePara(totalPut, MaxPerPage, "", True, "位", CurrentPage, "")%>
								  </td>
								 </tr>
							    </table>
                              </td>
                            </tr>

	 <%
		
	End Sub
	
	Sub Add()       
	   Dim ID,RS,RealName,Tel,Sex
	   ID=KS.ChkClng(KS.S("ID"))
	   Call KSUser.InnerLocation("我要投诉")
	   
	%>
			
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			                    <tr>

        <tr>
         <td align="center">
		         <script>
				  function checkform()
				  {
				   if ($('Title').value==''){
				    alert('请输入投诉主题!');
					$Foc('Title');
					return false;
				   }
				   if ($('Content').value==''){
				    alert('请输入投诉内容!');
					$Foc('Content');
					return false;
				   }
				  }
				 </script>
                  
                <table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="normaltext">
				 <form name="bmform" action="?action=DoSave" method="post">
				  <input type="hidden" name="TrainID" value="<%=ID%>">
                      <td width="145" align="right" class='splittd' height="25"><strong>意见主题：</strong></td>
                      <td width="797" class='splittd'> 
					  <input type="text" name="Title" class="textbox" size="30">
					 
				      </td>
				  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25"><strong>意见对象：</strong></td>
                      <td class='splittd' height="25"> <input type="text" name="Object" class="textbox" size="30"> </td>
                  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25"><strong>意见内容：</strong></td>
                      <td class='splittd' height="25"> 
					  <textarea name="content" style="width:450px;height:100px"></textarea>
				     </td>
                  </tr>
				  <tr>
                      <td width="145" align="right" class='splittd' height="25"><strong>期望解决方案：</strong></td>
                      <td class='splittd' height="25"> 
					  <textarea name="Hopesolution" style="width:450px;height:100px"></textarea>
				    </td>
                  </tr>
                    
                   
           </table>
                <br><div style="text-align:center">
				
				&nbsp;<input type="Submit" class="button" onClick="return(checkform())" value=" 立即投诉 ">
				
				</div>
                
		 
		 </td>
       </tr>
	    </form>

     </table>
	 <br><br><br><br>
	 <%'RS.Close:Set RS=Nothing
	End Sub
	
	Sub Addsave()
	    if ks.s("title")="" then
		 response.write "<script>alert('请输入主题!');history.back();</script>"
		 exit sub
		end if
	    if ks.s("content")="" then
		 response.write "<script>alert('请输入内容!');history.back();</script>"
		 exit sub
		end if
		
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select * from ks_feedback where username='" & KSUser.UserName & "' and id=" & ID,conn,1,3
		 If RS.EOf Then
		  rs.addnew
		  rs("adddate")=now
		 end if
		 rs("username")=ksuser.username
		 rs("title")=ks.s("title")
		 rs("object")=ks.s("object")
		 rs("content")=ks.s("content")
		 rs("hopesolution")=ks.s("hopesolution")
		 rs.update
		 rs.close
		 set rs=nothing
		 response.write "<script>alert('你的投诉已提交，请耐心等待处理结果!');location.href='User_FeedBack.asp';</script>"
	End Sub
	
	Sub View()
	   Call KSUser.InnerLocation("查看投诉详情")
       Dim ID,RS
	   ID=KS.ChkClng(KS.S("ID"))
	   Set RS=Server.CreateOBject("ADODB.RECORDSET")
	   RS.Open "Select * from ks_feedback where id=" & ID,conn,1,1

	   IF RS.Eof Then
	     RS.CLOSE:Set RS=Nothing
		 Response.Write "<script>alert('出错了!');window.close();</script>"
	   End If
	%>
          <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
			<title>投诉</title>
			<link href="images/css.css" type="text/css" rel="stylesheet" />
			<script src="../ks_inc/common.js"></script>
			</head>
			<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">	 
			
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
         <td align="center">
                  <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr align="center">
                      <td height="30"><strong>查看投诉详情</strong></td>
                    </tr>
           </table>
                <table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="normaltext">
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">意见主题：</td>
                      <td width="797" class='splittd'> 
					  &nbsp;<%=RS("title")%>
					 
				      </td>
				  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25">意见对象：</td>
                      <td class='splittd' height="25">&nbsp; <%=RS("object")%> </td>
                  </tr>
				   <tr>
                      <td width="145" align="right" class='splittd' height="25">意见内容：</td>
                      <td class='splittd' height="25">&nbsp; <%=RS("content")%> </td>
                  </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">希望处理结果：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("hopesolution")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理人：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("accepted")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理时间：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("accepttime")%></td>
                      
                    </tr>
                    <tr>
                      <td width="145" align="right" class='splittd' height="25">处理结果：</td>
                      <td class='splittd' height="25">&nbsp;<%=RS("acceptresult")%></td>
                      
                    </tr>
                    
                   
           </table>
                <br><div style="text-align:center">
				<input type="button" class="button" value=" 返 回 " onClick="history.back();">
				&nbsp;
				
				</div>
                
		 
		 </td>
       </tr>

     </table>
	 <%RS.Close:Set RS=Nothing

	 End Sub
	
	Sub FeedBackDel()
	  Conn.Execute("Delete from ks_FeedBack where (Accepted='' or Accepted is null ) and username='" & KSUser.UserName &"' and id=" & KS.ChkClng(KS.S("ID")))
	  Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	End Sub
	
End Class
%> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"--><%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New DefineForm
KSCls.Kesion()
Set KSCls = Nothing

Class DefineForm
        Private KS,F_Str,ID,TableName,Step,PostByStep,StepNum,ToUserEmail,Template,FormName,Temp
		Private Title,TimeLimit,StartDate,ExpiredDate,AllowGroupID,status,useronce,onlyuser,ShowNum
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
			dim Action,RS
			Action    = KS.S("Action")
			ID        = KS.ChkCLng(KS.S("ID"))
            Step      = KS.ChkCLng(KS.S("Step"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open "select * from ks_form where id=" & id,conn,1,1
			if not rs.eof then
			 FormName    = rs("FormName")
			 status      = rs("status")
			 TableName   = rs("TableName")
			 title       = rs("formname")
			 TimeLimit   = rs("TimeLimit")
			 StartDate   = rs("StartDate")
			 ExpiredDate = rs("ExpiredDate")
			 AllowGroupID= rs("AllowGroupID")
			 UserOnce    = rs("UserOnce")
			 OnlyUser    = rs("OnlyUser")
			 ShowNum     = rs("ShowNum")
			 PostByStep  = rs("PostByStep")
			 StepNum     = rs("StepNum")
			 ToUserEmail = rs("ToUserEmail")
			 IF Action="Save" Then 
				 Call LoadSave()
			 Else
			    Temp=RS("Template")
				If IsNull(Temp) Or Temp="" Then Temp=" "
			    Template=Split(Temp,"$aaa$")(step)
				If Step>0 and PostByStep=1 Then 
				  Call CollectHiddenFiled()
				End If
				F_Str=Template
			 End IF
			else
			 F_Str= "��Ч����"
			end if
			rs.Close():Set RS = Nothing
			If PostByStep=0 Then
			 F_Str=Replace(Replace(F_Str,"'","\'"),"""","\""")
			 F_Str=ReplaceJsBr(F_Str)
			Else
			%>
		    <html>
			<head>
			<title>�Զ����</title>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<style type="text/css">
			<!--
			td{FONT-FAMILY:����;FONT-SIZE: 9pt;line-height: 130%;}
			a{text-decoration: none;} /* �������»���,��Ϊunderline */
			a:link {color: #000000;} /* δ���ʵ����� */
			a:visited {color: #333333;} /* �ѷ��ʵ����� */
			a:hover{COLOR: #AE0927;} /* ����������� */
			a:active {color: #0000ff;} /* ����������� */
			-->
			</style>
			</head>
			<body style="background-color:transparent" oncontextmenu="return false">
			<%
			End If
			response.write F_Str
			If PostByStep=1 Then
			%>
			</body>
			</html>
			<script language="javascript" type="text/javascript">    
			function IFrameAutoFit()
			{
				try
				{
					if(window!=parent)
					{
						var a = parent.document.getElementsByTagName("IFRAME");
						for(var i=0; i<a.length; i++)
						{
							if(a[i].contentWindow == window)
							{
								var h1=0, h2=0,w1,w2;
								a[i].parentNode.style.height = a[i].offsetHeight +"px";
								a[i].parentNode.style.width  = a[i].offsetWidth +"px";
								a[i].style.height = "10px";
								if(document.documentElement && document.documentElement.scrollHeight)
								{
									h1 = document.documentElement.scrollHeight;
								}
								if(document.body)
								{
									h2=document.body.scrollHeight;
								}
								var h = Math.max(h1, h2);
								if(document.all) 
								{
									h += 4;
								}
								if(window.opera) 
								{
									h += 1;
								}
								
								if(document.documentElement && document.documentElement.scrollWidth)
								{
									w1 = document.documentElement.scrollWidth;
								}
								if(document.body)
								{
									w2=document.body.scrollWidth;
								}
								var w = Math.max(w1, w2);
								a[i].style.height = a[i].parentNode.style.height = h +"px";
								a[i].style.width=a[i].parentNode.style.width=w+"px";
							}
						}
					}
				}
				catch (ex)
				{
				}
			}
			if(window.attachEvent)
			{
				window.attachEvent("onload",  IFrameAutoFit);
			}
			else if(window.addEventListener)
			{
				window.addEventListener('load',  IFrameAutoFit,  false);
			}    
		</script>  

			<%
		 End If
    End Sub
     
	 '�ռ��û��ύ�������ֶμ����ύ	
	 Sub CollectHiddenFiled()
	   Dim HiddenFields,SQL,K,RS
	   Set RS=conn.execute("select FieldName,title,MustFillTF,FieldType from ks_formfield where itemid=" & id & " and ShowOnForm=1 and Step<=" & Step & " order by orderid")
	   If Not RS.Eof Then SQL=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
	     For K=0 To Ubound(SQL,2)
			if sql(2,k)=1 and KS.S(sql(0,k))="" then call KS.AlertHistory(sql(1,k) & "������д��",-1):exit sub
			select case sql(3,k)
			   case 8
				If Not KS.IsValidEmail(KS.S(sql(0,k))) Then  Call KS.AlertHistory("Email��ʽ����ȷ!",-1):Exit Sub
			   case 4
				If Not isnumeric(KS.S(sql(0,k))) Then  Call KS.AlertHistory("���ָ�ʽ����ȷ!",-1):Exit Sub
			   case 5
				If Not ISDate(KS.S(sql(0,k))) Then  Call KS.AlertHistory(sql(1,k) &"��ʽ����ȷ!",-1):Exit Sub
			end select
		 Next
		 
		 for k=0 to ubound(sql,2)
				HiddenFields=HiddenFields & "<input type=""hidden"" value=""" & Request.Form(sql(0,k)) & """ name=""" & sql(0,k) & """>" & vbcrlf
		next
	   End If
	   Template=Replace(Template,"{$HiddenFields}",HiddenFields)
	 End Sub
	 
	 Sub LoadSave()
	    Dim KSUser:Set KSUser=New UserCls
		Dim LoginTf:LoginTF=KSUser.UserLoginChecked
		
	   if status=0 then call KS.AlertHistory("�Բ��𣬸ñ���������",-1):exit sub
	   if TimeLimit=1 then
		 if now<StartDate then call KS.AlertHistory("�Բ�������" & startdate & "�������ύ��",-1):exit sub
		 if now>ExpiredDate then call KS.AlertHistory("�Բ��𣬸ñ�����" & expireddate & "���ڣ�",-1):exit sub
	   end if
	   
	   If (PostByStep=1 And Step=StepNum) Or PostByStep=0 Then
		   IF Trim(KS.S("Verifycode"))="" And Shownum=1 then Set KSUser=Nothing:call KS.AlertHistory("��֤��������룡",-1):exit sub
		   IF Trim(KS.S("Verifycode"))<>Trim(Session("Verifycode")) And Shownum=1 then Set KSUser=Nothing:call KS.AlertHistory("��֤�벻��ȷ��",-1):exit sub
	   End If
	   
	   IF onlyuser=1 and Cbool(LoginTf)=false Then Set KSUser=Nothing:call KS.AlertHistory("�Բ��𣬸ñ���Ҫ��¼������ύ��",-1):exit sub

	   if AllowGroupID<>"" then
	    if KS.FoundInArr(AllowGroupID,KSUser.groupid,",")=false then  Set KSUser=Nothing:call KS.AlertHistory("�Բ��������ڵ��û��鲻�ܲ���ñ����ύ��",-1):exit sub
	   end if
	   
	   if useronce=1 then
	    if not conn.execute("select username from " & TableName & " where username='" & ksuser.username &"'").eof then
		call KS.AlertHistory("�Բ��������ύ�����ñ�ֻ����һ����Ա�ύһ�Σ�",-1):exit sub
		end if
	   end if
	   
	   Dim S_Content,sql,k,email,ReturnInfo,UpFiles
	   Dim rs:set rs=conn.execute("select FieldName,title,MustFillTF,FieldType from ks_formfield where itemid=" & id & " and ShowOnForm=1 order by orderid")
	   if rs.eof then rs.close:set rs=nothing:call KS.AlertHistory("�����ύ����",-1):exit sub
	   sql=rs.getrows(-1):rs.close
	   s_content="<table border=0 cellpadding=0 cellspacing=0>" & vbcrlf
	   for k=0 to ubound(sql,2)
	    if sql(2,k)=1 and KS.S(sql(0,k))="" then call KS.AlertHistory(sql(1,k) & "������д��",-1):exit sub
		select case sql(3,k)
		   case 8
		    If Not KS.IsValidEmail(KS.S(sql(0,k))) Then  Call KS.AlertHistory("Email��ʽ����ȷ!",-1):Exit Sub
			email=KS.S(sql(0,k))
	       case 4
		    If Not isnumeric(KS.S(sql(0,k))) Then  Call KS.AlertHistory("���ָ�ʽ����ȷ!",-1):Exit Sub
		   case 5
		    If Not ISDate(KS.S(sql(0,k))) Then  Call KS.AlertHistory(sql(1,k) &"��ʽ����ȷ!",-1):Exit Sub
		end select
		s_content=s_content &"<tr>" & vbcrlf
	    s_content=s_content & "<td width=120 align=right>" & sql(1,k) & ":</td>" & vbcrlf
		s_content=s_content & "<td>" 
		
		If sql(3,k)=10 Then
		 s_content=s_content & KS.ClearBadChr(Request.Form(sql(0,k)))
		Elseif sql(3,k)=9 Then
		 s_content=s_content & "<a href='" & KS.S(sql(0,k)) & "' target='_blank'>������</a>"
		Else
		 s_content=s_content & KS.S(sql(0,k))
		End If
		s_content=s_content & "</td>" & vbcrlf
		s_content=s_content & "</tr>" & vbcrlf
	   next

	    s_content=s_content &"</table>"
		
	   rs.open "select * from " & TableName & " where 1=0",conn,1,3
	   rs.addnew
		rs("userip")=ks.getip
		rs("adddate")=now
		rs("username")=KSUser.UserName
		rs("status")=0
		for k=0 to ubound(sql,2)
		 If sql(3,k)=10 Then
			 rs(sql(0,k))=KS.ClearBadChr(Request.Form(sql(0,k)))
			 UpFiles=UpFiles&KS.S(sql(0,k))
		 Elseif sql(3,k)=9 Then
		     rs(sql(0,k))="<a href='" & KS.S(sql(0,k)) & "' target='_blank'>������</a>"
			 UpFiles=UpFiles&KS.S(sql(0,k))
		 Else
			 rs(sql(0,k))=KS.ClearBadChr(KS.S(sql(0,k)))
		 End If
		next
	  rs.update
	  rs.movelast
	  Call KS.FileAssociation(1016,RS("ID"),UpFiles,1)
	  rs.close
	  set rs=nothing
	  If ToUserEmail="1" Then
	      s_content="�𾴵��û������ã�<br />&nbsp;&nbsp;&nbsp;&nbsp;����������<font color=""red"">"  &KS.Setting(0) & "</font>�ύ[" & FormName & "]����Ϣ:<br />" & s_content
           ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "�û��ύ[" & FormName & "]����Ϣ", KS.Setting(11),KS.Setting(0), s_content,KS.Setting(11))
	     If Email="" Then Email=KSUser.Email
	     If Email<>"" Then
           ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "����" & KS.Setting(0) & "�ύ[" & FormName & "]����Ϣ", Email,KS.Setting(0), s_content,KS.Setting(11))
		   If ReturnInfo="OK" Then
		    ReturnInfo="�ѽ��ύ������͵���������" & Email & "!"
		   Else
		    ReturnInfo=""
		   End If
		 End If
	  End If
	  Set KSUser=Nothing
	  If PostByStep=1 Then
	  response.write "<script>alert('��ϲ��������Ϣ���ύ�ɹ���" & ReturnInfo & "');location.href='form.asp?id=" & ID& "';</script>"
	  Else
	  response.write "<script>alert('��ϲ��������Ϣ���ύ�ɹ���" & ReturnInfo & "');location.href='" & request.servervariables("http_referer") & "';</script>"
	  End If
	 End Sub
	 
	 Function ReplaceJsBr(Content)
		 Dim i
		 Dim JsArr:JSArr=Split(Content,Chr(13) & Chr(10))
		 For I=0 To Ubound(JsArr)
		   ReplaceJsBr=ReplaceJsBr & "document.writeln('" & JsArr(I) &"')" & vbcrlf 
		 Next
	End Function
End Class
%>
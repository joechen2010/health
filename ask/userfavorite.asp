<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New Ask_Fav
KSCls.Kesion()
Set KSCls = Nothing

Class Ask_Fav
        Private KS, KSR,KSUser,UserLoginTF,AnonymScore
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Sub Kesion()
		    Dim UserLoginTF:UserLoginTF=Cbool(KSUser.UserLoginChecked)
			If UserLoginTF=false Then
				Response.Write "<script>alert('������ʾ!\n\n�����¼��ſ����ղ�!');parent.ShowLogin();</script>"
				Exit Sub
			End If
		    Dim TopicID:TopicID = KS.ChkClng(Request("TopicID"))
			Dim Rs,SQL,m_strTitle,favtotalrec
			favtotalrec = Conn.Execute("SELECT COUNT(*) FROM KS_AskFavorite WHERE username='"&KSUser.UserName&"'")(0)
			If favtotalrec > 500 Then
				Response.Write "<script>alert('������ʾ!\n\n����ղؼ�¼�Ѿ��ﵽ����!');</script>"
				Exit Sub
			End If
			Set Rs = Conn.Execute("SELECT TopicID,title,username FROM KS_AskTopic WHERE topicid="&topicid&" And LockTopic=0")
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('������ʾ!\n\nû���ҵ���Ҫ�ղص���Ϣ!');</script>"
				Exit Sub
			Else
				If Rs("userName") = KSUser.UserName Then
					Response.Write "<script>alert('������ʾ!\n\n�����ղ��Լ�������!');</script>"
					Exit Sub
				End If
			End If
			Rs.Close:Set Rs = Nothing
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM KS_AskFavorite WHERE UserName='"&KSUser.UserName&"' And topicid="&topicid
			Rs.Open SQL,Conn,1,3
			If Rs.BOF And Rs.EOF Then
				Rs.Addnew
				Rs("username") = KSUser.UserName
				Rs("TopicID") = TopicID
				Rs("FavorTime") = Now()
				Rs.Update
			Else
				Response.Write "<script>alert('������ʾ!\n\n���Ѿ��ղ��˸�����!');</script>"
				Exit Sub
			End If
			Rs.Close:Set Rs = Nothing
			Response.Write "<script>alert('������ʾ!\n\n�ղسɹ�!');</script>"

		End Sub
End Class
%>
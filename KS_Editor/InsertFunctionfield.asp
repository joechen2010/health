<%@language=vbscript codepage="936" %>
<%
Option Explicit
Response.Buffer = True
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include File="../KS_Cls/Kesion.CommonCls.asp" -->
<!-- #include File="../conn.asp" -->
<!-- #include File="../Plus/Session.asp" -->
<%
Dim Login:Set Login=New LoginCheckCls1
Call Login.Run()
Set Login=Nothing
Dim KS:Set KS=New PublicCls
Dim ID, sql, rs

ID = KS.R(KS.S("id"))
Call Main

Call CloseConn
Set KS=Nothing
Sub Main()
    Set rs=Conn.Execute("select * from KS_Label where ID='" & ID & "'") 
    If rs.bof and rs.EOF Then
        response.write "��ǩ������"
    Else
%>
<html>
<head>
<title>�Զ��庯����ǩ���������</title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<script src="../ks_Inc/common.js" language="javascript"></script>
<script language="javascript">
function objectTag(itotal) {
        var TempStr="";
        for(i=0;i<itotal;i++){
		    if ($F('Field'+ i)=='')
			 {
			 alert('������'+$('Param'+i).innerHTML);
			 $('Field'+i).focus();
			 return false;
			 }
            if(i<itotal-1){
                TempStr =TempStr + $F('Field'+ i) + ","; 
            }else{
                TempStr=TempStr + $F('Field'+ i); 
            }
        }
	    var reval = '<%=Replace(rs("LabelName"),"}","") %>('+TempStr+')}';  
	    window.returnValue = reval;
	    window.close();
}
</script>
<link href='Editor.css' rel='stylesheet' type='text/css'>
</head>
<body>
<form name="myform">
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
  <tr>
    <td colspan="2" align="center"><strong>�����붯̬������ǩ����</strong><hr></td>
  </tr>
<%
   Dim arrFieldParam,FieldParams,FieldParam, i
   FieldParams=Split(rs("Description"),"@@@")
   If Ubound(FieldParams)>0 and FieldParams(1)<>"" Then
       FieldParam=FieldParams(1)
       ArrFieldParam=Split(FieldParam,vbcrlf)
       For i = 0 To UBound(arrFieldParam)
          response.write "<tr><td align='right'><span id='Param" & I &"'>" & arrFieldParam(i) & "</span>��</td><td><input type=""text"" id='Field" & i & "' name='Field" & i & "'></td></tr>"
       Next
    response.write "<tr><td colspan=2 align='center'><input TYPE='button' value=' ȷ �� ' onCLICK='objectTag(" & UBound(arrFieldParam)+1 & ")'></td>" 
 Else
 %>
 <script>window.returnValue='<%=Replace(rs("LabelName"),"}","") %>()}';window.close();</script>"
 <%
 response.end
 End If  
%>
  </tr>
</table>
<br>
<hr>
<font color=red>˵�����Զ��庯����ǩ�ĵ��ø�ʽ���£�<br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{SQL_��ǩ����(<font color=blue>����1,����2...</font>)}</font>

</form>
</body>
</html>
<%

    End If
    Set rs = Nothing
End Sub
%>
 

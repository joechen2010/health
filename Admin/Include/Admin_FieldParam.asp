<%
response.Expires = -1
response.ExpiresAbsolute = Now() - 1
response.Expires = 0
response.CacheControl = "no-cache"
Response.CodePage=936
Response.Charset="gb2312"

Dim fieldname, num, dbname, dbtype, isknow,isidarr,isid,datasourcetype

fieldname = Trim(Request("fieldname"))
dbname = Trim(Request("dbname"))
datasourcetype=request("datasourcetype")
isidarr=split(fieldname,".")
isid=false
if ubound(isidarr)=1 then
  if lcase(isidarr(1))="id" and datasourcetype="0" then
    isid=true
  end if
end if

If dbname = "" Then dbname = 0
dbtype = Trim(Request("dbtype"))
If dbtype = "" Then dbtype = 0
isknow = False
%>
<html>
<head>
<title>�ֶ���������</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
<script language = 'JavaScript'>
function changemode(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
    input1.style.display='';
    }else{
    input1.style.display='none';
    }
    if(dbname=='Num'){
    input2.style.display='';
    }else{
    input2.style.display='none';
    }
    if(dbname=='Date'){
    input3.style.display='';
    }else{
    input3.style.display='none';
    }
    if(dbname=='GetInfoUrl'|dbname=='GetClassUrl'){
    input5.style.display='';
    }else{
    input5.style.display='none';
    }
}
function changeDate(){
    var dbname=document.myform.Datetype.value;
    if(dbname=='3'){
    document.myform.Datemb.value="2";
    }else{
        document.myform.Datemb.value="YY-MM-DD hh:mm:ss";
    }
}
function submitdate(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
        dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + document.myform.CatNum.value + "," + document.myform.CatType.value + "," + document.myform.OutSplit.value + ","+document.myform.NullChar.value+")}";
    }
    if(dbname=='Num'){
	    for (var i=0;i<document.myform.OutType.length;i++){
            if (document.myform.OutType[i].checked){
                var OutType=document.myform.OutType[i].value;
        }
        }
       dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + OutType + "," + document.myform.XiaoShu.value + ")}";
    }
    if(dbname=='Date'){
    dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + document.myform.Datemb.value + ")}";
    }
    if(dbname=='GetInfoUrl'||dbname=='GetClassUrl'){
	    for (var i=0;i<document.myform.outype.length;i++){
            if (document.myform.outype[i].checked){
                var outype=document.myform.outype[i].value;
        }
        }
        dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + <%=dbname%> + "," + outype + ")}";
    }
    parent.InsertValue(dbname);
	parent.closeWindow();
}
</script>
</head>
<body>
<table width="100%">
<form method='post' action='' name='myform'>
    <tr class="tdbg"><td><strong>�ֶ����ƣ�</strong><input name='FieldName' type='text' id='FieldName' size='20' value="<% =fieldname %>" readonly></td></tr>
    <tr class="tdbg"><td><strong>������ͣ�</strong><select name="ftype" style="width:150" onChange="changemode()">
	<option value='Text'>�ı���</option>
<%
If (dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17 Then
    response.write "<option value='Num' selected>������</option>"
    isknow = True
Else
    response.write "<option value='Num'>������</option>"
End If
If dbtype = 7 Then
    response.write "<option value='Date' selected>ʱ����</option>"
    isknow = True
Else
    response.write "<option value='Date'>ʱ����</option>"
End If

	If (LCase(fieldname)<>"ks_class.id" and (LCase(fieldname) = "id" or cbool(isid) or LCase(fieldname) = "newsid" Or LCase(fieldname) = "picid" Or LCase(fieldname) = "downid" Or LCase(fieldname) = "flashid" Or LCase(fieldname) = "proid" or LCase(fieldname) = "movieid") or LCase(fieldname) = "gqid" or LCase(fieldname) = "classid") and datasourcetype="0" Then
        response.write "<option value='GetInfoUrl' selected>����URL��(ϵͳ����)</option>"
        isknow = True
    Else
       ' response.write "<option value='GetInfoUrl'>����URL(ϵͳ����)</option>"
    End If

    If Lcase(FieldName)="tid" or Lcase(fieldname)="ks_class.id" Then
        response.write "<option value='GetClassUrl' selected>��Ŀ|Ƶ��URL��(ϵͳ����)</option>"
        isknow = True
    Else
       ' response.write "<option value='GetClassUrl'>��Ŀ|Ƶ��URL(ϵͳ����)</option>"
    End If
%>
</select></td></tr>
<%
If isknow = False Then
    response.write "<tbody id='input1' style='display:'>"
Else
    response.write "<tbody id='input1' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>������ȣ�</strong><input name='CatNum' type='text' id='gotopic' size='6' value=0>
    &nbsp;&nbsp;&nbsp;<font color='#FF0000'>Ϊ0�򲻽ض�</font></td></tr>
	<tr class="tdbg"><td><strong>�ض���ʾ��</strong><Input name='CatType' type='text' value='...' size="6">
	  &nbsp;&nbsp;&nbsp;<font color='#FF0000'>Ϊ0����ʾ�κα�־</font></td>
	</tr>
    <tr class="tdbg"><td><strong>���˴���</strong><select name='OutSplit'><option value='0' selected>����HTML���</option><option value='1'>������HTML���</option><option value='2'>����HTML���</option></select></td></tr>
	    <tr class="tdbg"><td><strong>�ֶ�ֵΪ��ʱ�����</strong><input title='(��ͼƬֵΪ�գ������һ��Ĭ�ϵ�ͼƬ "/Images/defaule.gif")' name='NullChar' type='text' id='NullChar' size='20' value=""></td></tr>

</tbody>

<%
If ((dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17) And Not (LCase(fieldname) = "id") And Not (LCase(fieldname) = "classid")  and not isid   Then
    response.write "<tbody id='input2' style='display:'>"
Else
    response.write "<tbody id='input2' style='display:none'>"
End If
%>
    <tr class="tdbg"><td><strong>�����ʽ��</strong><Input type='radio' name='OutType' value='0' checked onClick="input21.style.display='';input22.style.display='none'">
    ԭ�� 
        <Input type='radio' name='OutType' value='1' onClick="input21.style.display='none';input22.style.display=''">С�� <Input type='radio' name='OutType' value='2' onClick="input21.style.display='none';input22.style.display='none'">�ٷ���</td></tr>
<%
        If ((dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17) And Not (LCase(fieldname) = "id") Then
        response.write "<tbody id='input21' style='display:'>"
        Else
        response.write "<tbody id='input21' style='display:none'>"
        End If
%>
</tbody>
    <tbody id='input22' style='display:none'><tr class="tdbg"><td><strong>С��λ����</strong><input name='XiaoShu' type='text' id='XiaoShu' size='4' value=2></td></tr></tbody>
</tbody>


<%
If dbtype = 7 Or dbtype = 135 Then
    response.write "<tbody id='input3' style='display:'>"
Else
    response.write "<tbody id='input3' style='display:none'>"
End If
%>
    
    <tr class="tdbg">
      <td><strong>�����ʽ��</strong>
        <input name='Datemb' type='text' id='Datemb' size='28' value="YYYY-MM-DD">
		<br>
		<font color=red>YYYY:��(4λ) YY:��(2λ) ��MM:�� ��DD:��<br>
		hh:ʱ�� mm:�֡� ss:��</font></td>
    </tr>
</tbody>


<%
If dbtype = 11 Then
    response.write "<tbody id='input4' style='display:'>"
Else
    response.write "<tbody id='input4' style='display:none'>"
End If
%>
    
</tbody>


<%
If (LCase(fieldname) = "tid" or LCase(fieldname) = "ks_class.id" or LCase(fieldname) = "id"  or isid or LCase(fieldname) = "newsid" Or LCase(fieldname) = "picid" Or LCase(fieldname) = "downid" Or LCase(fieldname) = "flashid" Or LCase(fieldname) = "proid" or LCase(fieldname) = "movieid" or LCase(fieldname) = "gqid" or LCase(fieldname) = "classid") and datasourcetype="0" Then
    response.write "<tbody id='input5' style='display:'>"
Else
    response.write "<tbody id='input5' style='display:none'>"
End If

%>
<tr class="tdbg"><td><strong>�����ʽ��</strong>
<Input type='radio' name='outype' value=0>
��� 
<Input type='radio' name='outype' value='1' checked>
����Url 
<Input type='radio' name='outype' value='2'> 
�ֶ�ֵ </td>
</tr>
</tbody>

<tr class="tdbg" align="center"><td><input type='button' value="����" onClick="submitdate();">&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' value="ȡ��" onClick="parent.closeWindow();"></td></tr>
<tr class="tdbg" height="100%"><td>&nbsp;<input name='Fieldnum' id='Fieldnum' value="<% =num %>" type='hidden'><br>&nbsp;<br>&nbsp;</td></tr>
</form>
</table>
</body>
</html>
 

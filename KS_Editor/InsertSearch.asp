<HTML>
<HEAD>
<TITLE>���� / �滻</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<script language="JavaScript">
var SelectionObj;
SelectionObj = dialogArguments.document.selection.createRange();
function SearchType()
{
    var MatchCase = 0;
    var MatchWord = 0;
    if (document.SearchForm.BtnMatchCase.checked) MatchCase = 4;
    if (document.SearchForm.BtnMatchWord.checked) MatchWord = 2;
    return(MatchCase+MatchWord);
}
function CheckInput()
{
    if (document.SearchForm.TxtSearch.value.length < 1) 
	{
        alert("�������������");
        return false;
    } 
	else 
	{
        return true;
    }
}
function SearchText()
{
    if (CheckInput()) 
	{
        var SearchValue = document.SearchForm.TxtSearch.value;
        SelectionObj.collapse(false);
        if (SelectionObj.findText(SearchValue,100,SearchType())) SelectionObj.select();
		else 
		{
            if (confirm("������ɣ��Ƿ�Ҫ�Ӷ�����ʼ����������")==true) 
			{
                SelectionObj.expand("textedit");
                SelectionObj.collapse();
                SelectionObj.select();
                SearchText();
            }
        }
    }
}
function SearchNextText()
{
	SearchText();
}
function ReplaceText()
{
    if (CheckInput()) 
	{
        if (document.SearchForm.BtnMatchCase.checked)
		{
            if (SelectionObj.text == document.SearchForm.TxtSearch.value) SelectionObj.text = document.SearchForm.TxtReplace.value;
        } 
		else 
		{
            if (SelectionObj.text.toLowerCase() == document.SearchForm.TxtSearch.value.toLowerCase()) SelectionObj.text = document.SearchForm.TxtReplace.value;
        }
        SearchText();
    }
}
function ReplaceAllText()
{
    if (CheckInput()) 
	{
        var SearchValue = document.SearchForm.TxtSearch.value;
        var WordCount = 0;
        var Massage = "";
        SelectionObj.expand("textedit");
        SelectionObj.collapse();
        SelectionObj.select();
        while (SelectionObj.findText(SearchValue,100,SearchType()))
		{
            SelectionObj.select();
            SelectionObj.text = document.SearchForm.TxtReplace.value;
            WordCount++;
        }
        if (WordCount == 0) Massage = "Ҫ���ҵ�����û���ҵ�"
        else Massage = WordCount + " ���ı����滻�ɹ�";
        alert(Massage);
		window.close();
    }
}
</script>

<link href="Editor.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY bgcolor="menu" leftmargin="10" topmargin="10" marginwidth="10" marginheight="10">
<FORM NAME="SearchForm" method="post" action="">
  <br>
  <TABLE border="0" align="center" cellpadding="0" CELLSPACING="0">
    <TR>
      <TD width="220"><fieldset>
        <legend><b>���� / �滻����</b></legend>
        <TABLE CELLSPACING="0" cellpadding="5" border="0">
          <TR> 
            <TD VALIGN="top" align="left" nowrap> <label for="strSearch">��������</label>
              <br> 
              <INPUT TYPE=TEXT SIZE=40 NAME=TxtSearch id="TxtSearch" style="width : 200px;"> 
              <br> <label for="strReplace">�滻����</label> <br> <INPUT TYPE=TEXT SIZE=40 NAME=TxtReplace id="TxtReplace" style="width : 200px;"> 
              <br> </td>
          </tr>
          <TR> 
            <TD VALIGN="top" align="left" nowrap><input type=Checkbox size=40 name=BtnMatchCase id="BtnMatchCase"> 
              <label for="blnMatchCase">���ִ�Сд</label> <input type=Checkbox size=40 name=BtnMatchWord id="BtnMatchWord"> 
              <label for="blnMatchWord">ȫ��ƥ��</label></td>
          </tr>
        </table>
        </fieldset></td>
      <td>&nbsp;</td>
<td rowspan="2" valign="top">
    <input type=button style="width:80px;margin-top:15px" name="btnFind" onClick="SearchText();" value="������һ��"><br>
    <input type=button style="width:80px;margin-top:5px" name="btnCancel" onClick="window.close();" value="�ر�"><br>
    <input type=button style="width:80px;margin-top:5px" name="btnReplace" onClick="ReplaceText();" value="�滻"><br>
    <input type=button style="width:80px;margin-top:5px" name="btnReplaceall" onClick="ReplaceAllText();" value="ȫ���滻"><br>
</td>
</tr>
</table>
</FORM>
</BODY>
</HTML>
 

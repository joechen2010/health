//ʹ��˵��:
//openBg(ѡ�����,'ajax�����ַ','Action','��ťID','��ֵID','����ֵID','��ѡ�����','����ѡ��ĳ���');
var lenMax=2;                   //�������ѡ��
var selectCount = 0;            //��ѡ����
var tempData=new Array();
var previewItem=new Array();
var goHistory=0;
var ajaxUrl=""; 
var action="";
var buttonID="";
var valueID="";
var parentId="";             //��¼��ID
var parentText=new Array();
var returnParentId=new Array();   //���ظ�ID
var parentValId=""           //�游��Ŀֵ��ID��

function showBox(){
 var str=('<DIV id="bg" style="display: none; Z-INDEX: 1; BACKGROUND: #ccc; FILTER: alpha(opacity=70); LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; opacity: 0.7"></DIV>');
 str+=('<DIV style="display:none;BORDER-RIGHT: #000 1px solid; BORDER-TOP: #000 1px solid; MARGIN-TOP: 10px; Z-INDEX: 2; BACKGROUND: #fff; OVERFLOW: hidden; BORDER-LEFT: #000 1px solid; WIDTH: 400px; BORDER-BOTTOM: #000 1px solid; POSITION: absolute; TOP: 0px" id="selectItem">');
 str+=('	<DIV style="BACKGROUND: #ccc;PADDING-LEFT: 10px; MARGIN: 1px; LINE-HEIGHT: 20px; HEIGHT: 20px">');
 str+=('	 <span style="float:left;font-weight:bold" id="boxTitle"></span>');
 str+=('	 <SPAN style="float:right;cursor:pointer" onclick="closeBg();">[ȡ��]</SPAN>');
 str+=('	 <SPAN style="float:right;cursor:pointer" onclick="makeSure();">[ȷ��]</SPAN>');
 str+=('	</DIV>');
 str+=('   <DIV style="clear: both; FONT-SIZE: 0px; OVERFLOW: hidden; HEIGHT: 0px"></DIV>');
 str+=('	<DIV style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px">');
 str+=('    <DIV id="selectSub">������...</DIV>');
 str+=('	</DIV>');
 str+=('	<DIV id="preview" style="BORDER-RIGHT: #ccc 1px solid; BORDER-TOP: #ccc 1px solid; MARGIN: 1px; BORDER-LEFT: #ccc 1px solid; BORDER-BOTTOM: #ccc 1px solid">');
 str+=('		<DIV id="boxTitle1" style="font-weight:bold;color:#999;BACKGROUND: #eee;PADDING-LEFT: 10px; MARGIN: 1px; LINE-HEIGHT: 20px; HEIGHT: 20px">');
 str+=('		</DIV>');
 str+=('		<DIV class=cont id="previewItem">&nbsp;</DIV>');
 str+=('	</DIV>');
 str+=('</DIV>');
 $("body").append(str);
}

function openBg(lenmax,url,act,btid,pvalId,svalid,t1,t2){ //���մ򿪹رտ���
	if ($("#bg")[0]==undefined){
	  showBox();
	}
	$("#boxTitle").html(t1);
	$("#boxTitle1").html(t2);
	lenMax=lenmax;
	ajaxUrl=url;
	Action=act;
	buttonID=btid;
	
	if (previewItem[buttonID]==undefined)
	{ $("#previewItem").html('&nbsp;');
	 }else{
	 $("#previewItem").html(previewItem[buttonID]);
	}

	valueID=svalid;
	if (pvalId!='')
	{
	 returnParentId[buttonID]=true;
	 parentValId=pvalId;
	}else{
	 returnParentId[buttonID]=false;
	}
	loadFirstData()
	$("#bg").css("display","block");
	var h = document.body.offsetHeight > document.documentElement.offsetHeight ? document.body.offsetHeight : document.documentElement.offsetHeight;
	$("#bg").css("height", h + "px");
	
	$("#selectItem").css("display","block");
	$("#selectItem").css("left",($("#bg")[0].offsetWidth - $("#selectItem")[0].offsetWidth)/2 + "px");
	$("#selectItem").css("top",document.body.scrollTop + 200 + "px");
}

function closeBg()
{
	$("#bg").css("display","none");
	$("#selectItem").css("display","none");
}
function loadFirstData()
{   
    if (tempData[buttonID]==null||tempData[buttonID]==''||goHistory==1){
		 $.get(ajaxUrl,{action:Action},function(d){
		  $("#selectSub").html(unescape(d));
		  tempData[buttonID]=unescape(d);
		 });
    }else{
	 $("#selectSub").html(tempData[buttonID]);
	}
	
	//��Ĭ��ѡ��
	var items = $("#previewItem")[0].getElementsByTagName("input");
	var len��= 0 ;
	for(var i = 0 ; i < items.length ; i++)
	{
	  if(items[i].checked == true){
	  same(items[i]);
	 // $("#selectSub").find(":input[type=checkbox][value="+items[i].value+"]").attr("checked",true);
	  }
	}
}
function loadSecond(id,text)
{
 if (action=='GetArea'){
  parentId=id;
 }else{
  parentId=text;
 }
 parentText[buttonID]=text;
 $.get(ajaxUrl,{action:Action,parentid:id},function(d){
  $("#selectSub").html(unescape(d));
  tempData[buttonID]=unescape(d);
 });
}

function open(id,state){ //��ʾ���ؿ���
if(state == 1)
	$("#"+id).css("display","block");
	$("#"+id).css("diaplay","none");
}
function addPreItem(){
	$("#previewItem").html("");
	var items = $("#selectSub")[0].getElementsByTagName("input");
	var len��= 0 ;
	for(var i = 0 ; i < items.length ; i++)
	{
	  if(items[i].checked == true)
		{
		len++;
		if(len > lenMax)
		{
		items[i].checked=false;
		alert('�Բ���,���ֻ��ѡ��'+lenMax+'��!');
		return false;
		}
		var mes = "<label><input type='checkbox' checked='true' value='"+ items[i].value +"' onclick='same(this);'>" + items[i].nextSibling.nodeValue+"</label>";
		$("#previewItem").html($("#previewItem").html()+mes);
		previewItem[buttonID]=$("#previewItem").html();
		}
	}
	previewItem[buttonID]=$("#previewItem").html();

}
function same(ck){
	var items = $("#selectSub")[0].getElementsByTagName("input");
	for(var i = 0 ; i < items.length ; i++)
	{
		if(ck.value == items[i].value)
		 {
		  items[i].checked = ck.checked;
		 }
	}
}
String.prototype.trim = function()   
{   
    return this.replace(/(^\s*)|(\s*$)/g, "");   
}
function makeSure(){
    var items = $("#previewItem")[0].getElementsByTagName("input");
	var len��= 0 ;
	var val='';
	var text='';
	for(var i = 0 ; i < items.length ; i++)
	{
	  if(items[i].checked == true){
	   if (val=='') {
	    text=items[i].nextSibling.nodeValue.trim();
	    val=items[i].value;
		}else{
		text+=","+items[i].nextSibling.nodeValue.trim();
	    val+=","+items[i].value;
		}
	  }
	}
	if (text!=''){
	 $("#"+buttonID).val(text);
	 $("#"+valueID).val(val);
	}
	if (returnParentId[buttonID]==true&&text!='')
	{
	 if (parentText[buttonID]!=''&&parentText[buttonID]!=undefined){
	 $("#"+buttonID).val(parentText[buttonID]+"->"+text);
	 }	
	 $("#"+parentValId).val(parentId);	
	}
	closeBg();
}
function goBack()
{goHistory=1;
 loadFirstData()
}

/*
KesionCMSͨ�ýű���������������2008-9-28
*/
$=function(obj){
 return document.getElementById(obj);
}
$F=function(obj){
 return document.getElementById(obj).value;
}
$Foc=function(obj){
 return document.getElementById(obj).focus();
}
//�ݴ�ű�
ResumeError=function ()
 {
        return true;
}
window.onerror = ResumeError;

 //����Ƿ������ַ�
is_zw=function(str){
	exp=/[0-9a-zA-Z_.,#@!$%^&*()-+=|\?/<>]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}
//��֤�Ƿ��������
CheckBadChar=function (Obj,AlertStr)
{
	exp=/[,��]/g;
	if(Obj.value.search(exp) != -1)
	{   alert(AlertStr+"���ܰ�������");
	    Obj.value="";
		Obj.focus();
		return false;
	}
	return true;
}
// ����Ƿ���Ч����չ��
IsExt=function(FileName, AllowExt){
		var sTemp;
		var s=AllowExt.toUpperCase().split("|");
		for (var i=0;i<s.length ;i++ ){
			sTemp=FileName.substr(FileName.length-s[i].length-1);
			sTemp=sTemp.toUpperCase();
			s[i]="."+s[i];
			if (s[i]==sTemp){
				return true;
				break;
			}
		}
		return false;
}
//����Ƿ����ַ���һ
is_number=function(str){
	exp=/[^0-9()-]/g;
	if(str.search(exp) != -1)
	{
		return false;
	}
	return true;
}
//������ַ�����
CheckNumber=function(Obj,DescriptionStr){
	if (Obj.value!='' && (isNaN(Obj.value) || Obj.value<0))
	{
		alert(DescriptionStr+"Ӧ����Ч���֣�");
		Obj.value="";
		Obj.focus();
		return false;
	}
	return true;
}
//�������ʼ���Ч��
is_email=function(str){ 
if((str.indexOf("@")==-1)||(str.indexOf(".")==-1)){
	return false;
	}
	return true;
}
//������ڸ�ʽ�Ƿ�Ϊ2008-01-01 13:01:01
is_date=function(str){   
var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/; 
var r = str.match(reg); 
if(r==null)return is_shortdate(str); 
var d= new Date(r[1], r[3]-1,r[4],r[5],r[6],r[7]); 
var v=(d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]&&d.getHours()==r[5]&&d.getMinutes()==r[6]&&d.getSeconds()==r[7]);
if (v==false)
  return is_shortdate(str)
 else
 return true;
}
////������ڸ�ʽ�Ƿ�Ϊ2008-01-01
is_shortdate=function(str){
var r = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/); 
if(r==null)return false; 
var d= new Date(r[1], r[3]-1, r[4]); 
return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]);	
}

/* ���ѡ�б�li */
function chk_iddiv(id){
	var objc=document.getElementById("c"+id); //��ѡ��
	var obju=document.getElementById("u"+id);//ul
	if (objc.checked==''){
		objc.checked='checked';
		obju.style.background='#EEF8FE';
		//obju.className='listmouseover';
	}else{
		objc.checked='';
		obju.style.background='';
		//obju.className='list';
	}
}
/**/
function chk_idBatch(form,askString){
	var bCheck;
	bCheck=false;
	for (var i=0;i < form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name == "id"){
       if (e.checked ==1){
       		bCheck=true;
       		break;
       	}
		}
	}
	
	if (bCheck==false){
		alert("��ѡ��Ҫ����������!")
		return false;
		}
	else{
		return confirm('ȷ��Ҫ'+askString+"?");
		}
}
function get_Ids(form)
{
	var ids='';
	for (var i=0;i < form.elements.length;i++)
	{
			var e = form.elements[i];
			if (e.name == "id"){
			   if (e.checked ==1){
			      if (ids=='')
				   ids=e.value;
				  else
					ids+=","+e.value;
				  }
				}
	}
	return ids;
}
function Select(flag)
{  
  $("input[type=checkbox]").each(function(){
  if ($(this).attr("name")=="id"){
	var objc=$("#c"+$(this).val()); 
	var obju=$("#u"+$(this).val());
	switch (flag){
	  case 0:  //ȫѡ
	   objc.attr("checked",true);
	   obju.attr("style","background:#eef8fe");
	   break;
	  case 1: //��ѡ
		if (objc.attr("checked")==false){
			objc.attr("checked",true);
			obju.attr("style","background:#eef8fe");
		}else{
			objc.attr("checked",false);
	    	obju.attr("style","background:");
		}
		break;
	 case 2:  //��ѡ
		objc.attr("checked",false);
	    obju.attr("style","background:");
		break;
	 }
  }
 })
}


// utility function called by getCookie( )
 function getCookieVal(offset) {
			var endstr = document.cookie.indexOf (";", offset);
			if (endstr == -1) {
				endstr = document.cookie.length;
			}
		    return unescape(document.cookie.substring(offset, endstr));
}
// primary function to retrieve cookie by name
function getCookie(name) {
			var arg = name + "=";
			var alen = arg.length;
			var clen = document.cookie.length;
			var i = 0;
			while (i < clen) {
				var j = i + alen;
				if (document.cookie.substring(i, j) == arg) { 
					return getCookieVal(j);
				}
				i = document.cookie.indexOf(" ", i) + 1;
				if (i == 0) break; 
			}
			return "";
}
// store cookie value with optional details as needed
function setCookie(name, value) {
			document.cookie = name + "=" + escape (value)
		}
		// remove the cookie by setting ancient expiration date
		function deleteCookie(name,path,domain) {
			if (getCookie(name)) {
				document.cookie = name + "=" 
			}
		}

function CheckAll(form)
{
	 for (var i=0;i<form.elements.length;i++)
	 {
		var e = form.elements[i];
		if (e.Name != 'chkAll'&&e.disabled==false)
		e.checked = form.chkAll.checked;
	}
 } 
function OpenWindow(Url,Width,Height,WindowObj){
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	return ReturnStr;
}
var obj=null;
function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj){
	if (document.all){
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
	}else{
		 
	 obj=SetObj;
	 Width=Width+180;
	 Height=Height+80;
	 window.open(Url,'newWin','modal=yes,width='+Width+',height='+Height+',resizable=no,scrollbars=no');
	}
}
function setVal(v)
{ 
 obj.value=v;
}
function CheckEnglishStr(Obj,DescriptionStr)
{
	var TempStr=Obj.value,i=0,ErrorStr='',CharAscii;
	if (TempStr!='')
	{
		for (i=0;i<TempStr.length;i++)
		{
			CharAscii=TempStr.charCodeAt(i);
			if (CharAscii>=255||CharAscii<=31)
			{
				ErrorStr=ErrorStr+TempStr.charAt(i);
			}
			else
			{
				if (!CheckErrorStr(CharAscii))
				{
					ErrorStr=ErrorStr+TempStr.charAt(i);
				}
			}
		}
		if (ErrorStr!='')
		{
			alert("������Ϣ:\n\n"+DescriptionStr+'���ַǷ��ַ�:'+ErrorStr);
			Obj.focus();
			return false;
		}
		if (!(((TempStr.charCodeAt(0)>=48)&&(TempStr.charCodeAt(0)<=57))||((TempStr.charCodeAt(0)>=65)&&(TempStr.charCodeAt(0)<=90))||((TempStr.charCodeAt(0)>=97)&&(TempStr.charCodeAt(0)<=122))))
		{
			alert(DescriptionStr+'���ַ�ֻ�ܹ�Ϊ���ֻ�����ĸ');
			Obj.focus();
			return false;
		}
	}
	return true;
}
function CheckErrorStr(CharAsciiCode)
{
	var TempArray=new Array(34,47,92,42,58,60,62,63,124);
	for (var i=0;i<TempArray.length;i++)
	{
		if (CharAsciiCode==TempArray[i]) return false;
	}
	return true;
}
//Obj�����Ķ���,OpStr--BottomFrame��ʾ��ǰ��������ʾ��Ϣ,ButtonSymbol��ť״̬,MainUrl--MainFrame������
function SelectObjItem1(Obj,OpStr,ButtonSymbol,MainUrl,ChannelID)
{   if (OpStr!='')
    {
		window.parent.parent.frames['BottomFrame'].location.href='KS.Split.asp?ChannelID='+escape(ChannelID)+'&OpStr='+escape(OpStr)+'&ButtonSymbol='+escape(ButtonSymbol);
		}
	if(MainUrl!='')
	{window.parent.parent.frames['MainFrame'].location.href=MainUrl;
	}

}
function FolderClick(Obj,el)
{   	var i=0;
  for (var i=0;i<document.all.length;i++)
	   {
		if (document.all(i).className=='FolderSelected') document.all(i).className='';
	    }
	         Obj.className='FolderSelected';
	  
              for (i=0;i<DocElementArr.length;i++)
			{
				if (el==DocElementArr[i].Obj)
				{
					if (DocElementArr[i].Selected==false)
					{
						DocElementArr[i].Obj.className='FolderSelectItem';
						DocElementArr[i].Selected=true;
					}
					else
					{
						DocElementArr[i].Obj.className='FolderItem';
						DocElementArr[i].Selected=false;
					}
				}
			}
}
function InsertKeyWords(obj,KeyWords)
{ 
	if (KeyWords!='')
	{
		if (obj.value.search(KeyWords)==-1)
		{
			if (obj.value=='') obj.value=KeyWords;
			else obj.value=obj.value+','+KeyWords;
			
		}
	}
	if (KeyWords == 'Clean')
	{
		obj.value = '';
	}
	return;
}
//���Ͳ���������Frames����
function SendFrameInfo(MainUrl,LeftUrl,ControlUrl){
	location.href=MainUrl;
    parent.frames['LeftFrame'].LeftInfoFrame.location.href=LeftUrl;
	 $(parent.document).find('#BottomFrame')[0].src=ControlUrl;
}

// ���ļ���չ��ȡͼ������������
			function getFilePic(url)
			{
				var sExt;
				sExt=url.substr(url.lastIndexOf(".")+1);
				sExt=sExt.toUpperCase();
				var sPicName;
				switch(sExt)
				{
				case "TXT":
					sPicName = "txt.gif";
					break;
				case "CHM":
				case "HLP":
					sPicName = "hlp.gif";
					break;
				case "DOC":
					sPicName = "doc.gif";
					break;
				case "PDF":
					sPicName = "pdf.gif";
					break;
				case "MDB":
					sPicName = "mdb.gif";
					break;
				case "GIF":
					sPicName = "gif.gif";
					break;
				case "JPG":
					sPicName = "jpg.gif";
					break;
				case "BMP":
					sPicName = "bmp.gif";
					break;
				case "PNG":
					sPicName = "pic.gif";
					break;
				case "ASP":
				case "JSP":
				case "JS":
				case "PHP":
				case "PHP3":
				case "ASPX":
					sPicName = "code.gif";
					break;
				case "HTM":
				case "HTML":
				case "SHTML":
					sPicName = "htm.gif";
					break;
				case "ZIP":
				case "RAR":
					sPicName = "zip.gif";
					break;
				case "EXE":
					sPicName = "exe.gif";
					break;
				case "AVI":
					sPicName = "avi.gif";
					break;
				case "MPG":
				case "MPEG":
				case "ASF":
					sPicName = "mp.gif";
					break;
				case "RA":
				case "RM":
					sPicName = "rm.gif";
					break;
				case "MP3":
					sPicName = "mp3.gif";
					break;
				case "MID":
				case "MIDI":
					sPicName = "mid.gif";
					break;
				case "WAV":
					sPicName = "audio.gif";
					break;
				case "XLS":
					sPicName = "xls.gif";
					break;
				case "PPT":
				case "PPS":
					sPicName = "ppt.gif";
					break;
				case "SWF":
					sPicName = "swf.gif";
					break;
				default:
					sPicName = "unknow.gif";
					break;
				}
				return sPicName;
			
			}
function Getcolor(img_val,Url,input_val){
	var arr = showModalDialog(Url, "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr != null){
		document.getElementById(input_val).value = arr;
		img_val.style.backgroundColor = arr;
		}
}
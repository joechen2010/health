/*================================================================
Created:2007-5-29
Autor��linwenzhong
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS V4.0
Service QQ��111394,54004407
==================================================================*/
//ajax �ؼ�
function PageAjax(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
var loader=new PageAjax;
function ajaxLoadPage(url,request,method,fun)
{ 
	method=method.toUpperCase();
	if (method=='GET')
	{
		urls=url.split("?");
		if (urls[1]=='' || typeof urls[1]=='undefined')
		{
			url=urls[0]+"?"+request;
		}
		else
		{
			url=urls[0]+"?"+urls[1]+"&"+request;
		}
		
		request=null;
	}
	loader.open(method,url,true);
	if (method=="POST")
	{
		loader.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	}
	loader.onreadystatechange=function(){
	     eval(fun+'()');
	}
	loader.send(request);
 }
  //����֧��
  function Support(id,typeid,installdir)
  { 
    try{
        ajaxLoadPage(installdir+'plus/Comment.asp','action=Support&Type='+typeid+'&id=' +id,'post','callback');
    }
	catch(e){
		CreateJs(installdir+'plus/Comment.asp?action=Support&Type='+typeid+'&id=' +id);
	}
     
  }
  function callback()
  {
  if (loader.readyState==4){
	var s=loader.responseText;
	ShowSupportMessage(s);
  }
  }
  function ShowSupportMessage(s)
  {
	if (s=='good'||s=='bad'){
	   loadDate(_page);
	 }
	else
	alert(s);
}
 //�ظ�
function reply(channelid,quoteId,installdir){
	popup("<b>���ûظ�</b>","<div style='height:200px;text-align:center'><iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='"+installdir+"plus/comment.asp?action=QuoteSave' method='post'><input type='hidden' name='channelid' value='"+channelid+"'><input type='hidden' name='quoteId' value='"+quoteId+"'><textarea name='quotecontent' style='width:300px;height:150px'></textarea><br><label><input type='checkbox' value='1' name='Anonymous'>��������</label> <input type='submit' value='����'></form></div>",400);
}
 //��ǰҳ,Ƶ��ID,��ĿID����ϢID,Action,InstallDir
function Page(curPage,channelid,infoid,action,installdir)
   {
   this._channelid = channelid;
   this._infoid    = infoid;
   this._action    = action;
   this._url       = installdir +"plus/comment.asp";
   
   this._c_obj="c_"+infoid;
   this._p_obj="p_"+infoid;
   this._installdir=installdir;
   this._page=curPage;
     loadDate(1);
   }
function loadDate(p)
{  this._page=p;
   var loadurl=_url+"?channelid="+_channelid+"&infoid="+_infoid+"&action=" +_action+"&page="+p;
 try{
   var xhr=new PageAjax();
   xhr.open("get",loadurl,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1){
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3){
			  }
			  else if(xhr.readyState==4){
				 if (xhr.status==200)
				 {   
					  show(xhr.responseText);
				 }
			}
	   }
    xhr.send(null); 
	}catch(e){
		CreateJs(loadurl);
	}
}
function CreateJs(loadurl)
{
	 var head = document.getElementsByTagName("head")[0];        
	 var js = document.createElement("script"); 
	 js.src = loadurl+'&printout=js'; 
	 head.appendChild(js);   
}

function show(text)
{
  var pagearr=text.split("{ks:page}")
  var pageparamarr=pagearr[1].split("|");
  count=pageparamarr[0];    
  perpagenum=pageparamarr[1];
  pagecount=pageparamarr[2];
  itemunit=pageparamarr[3];   
  itemname=pageparamarr[4];
  pagestyle=pageparamarr[5];
  document.getElementById(_c_obj).innerHTML=pagearr[0];
  pagelist();
}

function pagelist()
{
 var n=1;	
 var statushtml=null;
 switch(parseInt(this.pagestyle))
 {
  case 1:	
     statushtml="��"+this.count+this.itemunit+" <a href=\"javascript:homePage();\" title=\"��ҳ\">��ҳ</a> <a href=\"javascript:previousPage()\" title=\"��һҳ\">��һҳ</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"��һҳ\">��һҳ</a> <a href=\"javascript:lastPage();\" title=\"���һҳ\">βҳ</a> ҳ��:<font color=red>"+this._page+"</font>/"+this.pagecount+"ҳ "+this.perpagenum+this.itemunit+this.itemname+"/ҳ";
		break;
  case 2:
	 statushtml="��"+this.pagecount+"ҳ/"+this.count+this.itemunit+this.itemname+" <a href=\"javascript:homePage();\" title=\"��ҳ\"><span style=\"font-family:webdings;font-size:14px\">9</span></a> <a href=\"javascript:previousPage()\" title=\"��һҳ\"><span style=\"font-family:webdings;font-size:14px\">7</span></a>&nbsp;";
	 var startpage=1;
	 if (this._page>10)
	   startpage=(parseInt(this._page/10)-1)*10+parseInt(this._page%10)+1;
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this._page)
		   statushtml+="<a href=\"#\"><font color=\"#ff0000\">"+i+"</font></a>&nbsp;"
		  else
			statushtml+="<a href=\"javascript:turn("+i+")\">["+i+"]</a>&nbsp;"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"��һҳ\"><span style=\"font-family:webdings;font-size:14px\">8</span></a> <a href=\"javascript:lastPage();\" title=\"���һҳ\"><span style=\"font-family:webdings;font-size:14px\">:</span></a>";
	 statushtml="<span class='kspage'>"+statushtml+"</span>";
	break;	 
  case 3:
     statushtml="��<font color=#ff000>"+this._page+"</font>ҳ ��"+this.pagecount+"ҳ <a href=\"javascript:homePage();\" title=\"��ҳ\"><<</a> <a href=\"javascript:previousPage()\" title=\"��һҳ\"><</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"��һҳ\">></a> <a href=\"javascript:lastPage();\" title=\"���һҳ\">>></a> "+this.perpagenum+this.itemunit+this.itemname+"/ҳ";
  case 4:
     statushtml="ҳ��:<font color=red>"+this._page+"</font>/"+this.pagecount+"ҳ [ <a href=\"javascript:homePage();\" title=\"��ҳ\">��ҳ</a> <a href=\"javascript:previousPage()\" title=\"��һҳ\">��һҳ</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"��һҳ\">��һҳ</a> <a href=\"javascript:lastPage();\" title=\"���һҳ\">βҳ</a> ]";
   break;
 }
	 statushtml+="&nbsp;��ת����<select name=\"goto\" onchange=\"turn(parseInt(this.value));\">";
	  for(var i=1;i<=this.pagecount;i++){
		 if (i==this._page)
		 statushtml+="<option value='"+i+"' selected>"+i+"</option>";
		 else
		 statushtml+="<option value='"+i+"'>"+i+"</option>";
	  }	
	 statushtml+="</select>ҳ";
	 
	 if (this.pagecount!=""&&this.count!=0)
	 {
	 document.getElementById(this._p_obj).innerHTML='<div style="margin-top:8px">'+statushtml+'</div>';
	 }
}
function homePage()
{
   if(this._page==1)
    alert("�Ѿ�����ҳ�ˣ�")
   else
   loadDate(1);
} 
function lastPage()
{
   if(this._page==this.pagecount)
    alert("�Ѿ������һҳ�ˣ�")
   else
   loadDate(this.pagecount);
} 
function previousPage()
{
   if (this._page>1)
      loadDate(this._page-1);
   else
      alert("�Ѿ��ǵ�һҳ��");      
}

function nextPage()
{
   if(this._page<this.pagecount)
      loadDate(this._page+1);
   else
      alert("�Ѿ������һҳ��");
}
function turn(i)
{
      loadDate(i);
}
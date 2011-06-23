/*================================================================
Created:2007-5-29
Autor：linwenzhong
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS V4.0
Service QQ：111394,54004407
==================================================================*/
//ajax 控件
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
  //评论支持
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
 //回复
function reply(channelid,quoteId,installdir){
	popup("<b>引用回复</b>","<div style='height:200px;text-align:center'><iframe style='display:none' src='about:blank' id='_framehidden' name='_framehidden' width='0' height='0'></iframe><form name='rform' target='_framehidden' action='"+installdir+"plus/comment.asp?action=QuoteSave' method='post'><input type='hidden' name='channelid' value='"+channelid+"'><input type='hidden' name='quoteId' value='"+quoteId+"'><textarea name='quotecontent' style='width:300px;height:150px'></textarea><br><label><input type='checkbox' value='1' name='Anonymous'>匿名发表</label> <input type='submit' value='发表'></form></div>",400);
}
 //当前页,频道ID,栏目ID，信息ID,Action,InstallDir
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
     statushtml="共"+this.count+this.itemunit+" <a href=\"javascript:homePage();\" title=\"首页\">首页</a> <a href=\"javascript:previousPage()\" title=\"上一页\">上一页</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">下一页</a> <a href=\"javascript:lastPage();\" title=\"最后一页\">尾页</a> 页次:<font color=red>"+this._page+"</font>/"+this.pagecount+"页 "+this.perpagenum+this.itemunit+this.itemname+"/页";
		break;
  case 2:
	 statushtml="共"+this.pagecount+"页/"+this.count+this.itemunit+this.itemname+" <a href=\"javascript:homePage();\" title=\"首页\"><span style=\"font-family:webdings;font-size:14px\">9</span></a> <a href=\"javascript:previousPage()\" title=\"上一页\"><span style=\"font-family:webdings;font-size:14px\">7</span></a>&nbsp;";
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
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"下一页\"><span style=\"font-family:webdings;font-size:14px\">8</span></a> <a href=\"javascript:lastPage();\" title=\"最后一页\"><span style=\"font-family:webdings;font-size:14px\">:</span></a>";
	 statushtml="<span class='kspage'>"+statushtml+"</span>";
	break;	 
  case 3:
     statushtml="第<font color=#ff000>"+this._page+"</font>页 共"+this.pagecount+"页 <a href=\"javascript:homePage();\" title=\"首页\"><<</a> <a href=\"javascript:previousPage()\" title=\"上一页\"><</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">></a> <a href=\"javascript:lastPage();\" title=\"最后一页\">>></a> "+this.perpagenum+this.itemunit+this.itemname+"/页";
  case 4:
     statushtml="页次:<font color=red>"+this._page+"</font>/"+this.pagecount+"页 [ <a href=\"javascript:homePage();\" title=\"首页\">首页</a> <a href=\"javascript:previousPage()\" title=\"上一页\">上一页</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"下一页\">下一页</a> <a href=\"javascript:lastPage();\" title=\"最后一页\">尾页</a> ]";
   break;
 }
	 statushtml+="&nbsp;跳转到第<select name=\"goto\" onchange=\"turn(parseInt(this.value));\">";
	  for(var i=1;i<=this.pagecount;i++){
		 if (i==this._page)
		 statushtml+="<option value='"+i+"' selected>"+i+"</option>";
		 else
		 statushtml+="<option value='"+i+"'>"+i+"</option>";
	  }	
	 statushtml+="</select>页";
	 
	 if (this.pagecount!=""&&this.count!=0)
	 {
	 document.getElementById(this._p_obj).innerHTML='<div style="margin-top:8px">'+statushtml+'</div>';
	 }
}
function homePage()
{
   if(this._page==1)
    alert("已经是首页了！")
   else
   loadDate(1);
} 
function lastPage()
{
   if(this._page==this.pagecount)
    alert("已经是最后一页了！")
   else
   loadDate(this.pagecount);
} 
function previousPage()
{
   if (this._page>1)
      loadDate(this._page-1);
   else
      alert("已经是第一页了");      
}

function nextPage()
{
   if(this._page<this.pagecount)
      loadDate(this._page+1);
   else
      alert("已经到最后一页了");
}
function turn(i)
{
      loadDate(i);
}
function $$(_sId){return document.getElementById(_sId)}

var ksblog = new Object();
ksblog._url='spaceajax.asp';
ksblog._mainlist='blogmain';
ksblog._pagelist='kspage';
ksblog._usernmae=null;
ksblog.loading = function(tag,username) {
	this._username=username;
	//alert(tag);
	//document.getElementById(ksblog._mainlist).innerHTML=tag;
	//return;
	title=document.title.split('-')[0];
	switch (tag)
	 {
	  case 'intro':
	     document.title=title+'-公司简介';
		 this.loadintro();
		 break;
	  case 'product':
	     document.title=title+'-主营产品';
		 this.loadproduct();
		 break;
	  case 'news':
	     document.title=title+'-公司动态';
		 this.loadnews();
		 break;
	  case 'job':
	    document.title=title+'-公司招聘';
		this.loadjob();
		break;
	  case 'log':
	    document.title=title+'-日志列表';
	    this.loadlog();
		break;
	  case 'guest':
	    document.title=title+'-留言信息';
		this.loadguest();
		break;
	  case 'listxx':
 	   var _request='channelid=1&action='+tag+'&username='+escape(this._username);
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setxx");
	   break;
	  default:
	  {
	   document.title=title+'-联系档案';
	   $$(ksblog._pagelist).style.display='none';
 	   var _request='action='+tag+'&username='+escape(this._username);
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setObj");
	 }
	 }
}

ksblog.checkmsg=function(){
		     var message=escape($$("s_message").value);
			 var username=escape($$("s_username").value);
			 if (username==''){
			  alert('参数传递出错!');
			  closeWindow();
			 }
			 if (message==''){
			   alert('请输入消息内容!');
			   $$("s_message").focus();
			   return false;
			 }
			 	var ksxhr=new ksblog.Ajax;
				var senddata="../plus/ajaxs.asp?action=SendMsg&username="+username+"&message="+message;
				ksxhr.open("get",senddata,true);
				ksxhr.onreadystatechange=function(){
					  if(ksxhr.readyState==4)
					  {
								 if (ksxhr.status==200)
								 { var s=ksxhr.responseText;
								   if (s!='success'){
										alert(r);
									 }else{
										 alert('恭喜，您的消息已发送,对方登录后将看到您的消息!');
										 closeWindow();
									 }
								 }
							  }
							}
				ksxhr.send(null);  


}
ksblog.sendMsg=function(ev,username)
{ 
	 mousepopup(ev,"<img src='../images/user/mail.gif' align='absmiddle'>发送消息","对方登录后可以看到您的消息(可输入255个字符)<br /><textarea name='message' id='s_message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(ksblog.checkmsg())' value=' 确 定 ' class='button'><input type='hidden' value="+username+" name='username' id='s_username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='button'></div>",350);
    ksblog.checkIsLogin();
}
ksblog.checkIsLogin=function(){
	var ksxhr=new ksblog.Ajax;
	var senddata="../plus/ajaxs.asp?action=CheckLogin";
	ksxhr.open("get",senddata,true);
    ksxhr.onreadystatechange=function(){
		  if(ksxhr.readyState==4)
		  {
					 if (ksxhr.status==200)
					 { var s=ksxhr.responseText;
					   if (s!='true'){
							 ksblog.ShowLogin();
						 }
					 }
				  }
				}
	ksxhr.send(null);  
}

ksblog.ShowLogin=function(){ 
  popupIframe('会员登录','../user/userlogin.asp?Action=Poplogin',397,184,'no');
}

ksblog.addF=function(ev,username){ 
	mousepopup(ev,"<img src='../images/user/log/106.gif'>添加好友","通过对方验证才能成为好友(可输入255个字符)<br /><textarea name='message' id='f_message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(ksblog.checkAddF())' value=' 确 定 ' class='button'><input type='hidden' value="+username+" name='username' id='f_username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='button'></div>",350);

	var isMyFriend=false;
	var ksxhr=new ksblog.Ajax;
	var senddata="../plus/ajaxs.asp?action=CheckMyFriend&username="+escape(username);
				ksxhr.open("get",senddata,true);
				ksxhr.onreadystatechange=function(){
					  if(ksxhr.readyState==4)
					  {
								 if (ksxhr.status==200)
								 { var b=ksxhr.responseText;
								   if (b=='nologin'){
									  closeWindow();
									  ksblog.ShowLogin();
									}else if (b=='true'){
									  closeWindow();
									  alert('用户['+username+']已经是您的好友了！');
									  return false;
									 }else if(b=='verify'){
									  closeWindow();
									  alert('您已邀请过['+username+'],请等待对方的认证!');
									  return false;
									 }else{
									 }
								 }
							  }
							}
				ksxhr.send(null);  
}
ksblog.checkAddF=function(){
		 var message=escape($$("f_message").value);
		 var username=escape($$("f_username").value);
		 if (username==''){
		  alert('参数传递出错!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('请输入附言!');
		   $$("f_message").focus();
		   return false;
		 }
	var ksxhr=new ksblog.Ajax;
	var senddata="../plus/ajaxs.asp?action=AddFriend&username="+username+"&message="+message;
	ksxhr.open("get",senddata,true);
    ksxhr.onreadystatechange=function(){
		  if(ksxhr.readyState==4)
		  {
					 if (ksxhr.status==200)
					 { var r=ksxhr.responseText;
					   r=unescape(r);
					   if (r!='success'){
						alert(r);
					   }else{
						 alert('您的请求已发送,请等待对方的确认!');
						 closeWindow();
					   }
					 }
				  }
				}
	ksxhr.send(null);  
		 
}



ksblog.Ajax=function(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
var loader=new ksblog.Ajax;
ksblog.ajaxLoadPage=function(url,request,method,fun)
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
ksblog.formToRequestString=function(form_obj)
{
    var query_string='';
    var and='';
    for (var i=0;i<form_obj.length;i++ )
    {
        e=form_obj[i];
        if (e.name) {
            if (e.type=='select-one') {
                element_value=e.options[e.selectedIndex].value;
            } else if (e.type=='select-multiple') {
                for (var n=0;n<e.length;n++) {
                    var op=e.options[n];
                    if (op.selected) {
                        query_string+=and+e.name+'='+escape(op.value);
                        and="&"
                    }
                }
                continue;
            } else if (e.type=='checkbox' || e.type=='radio') {
                if (e.checked==false) {   
                    continue;   
                }   
                element_value=e.value;
            } else if (typeof e.value != 'undefined') {
                element_value=e.value;
            } else {
                continue;
            }
            query_string+=and+e.name+'='+escape(element_value);
            and="&"
        }

    }
    return query_string;
}
ksblog.ajaxFormSubmit=function(form_obj,fun)
{
	ksblog.ajaxLoadPage(form_obj.getAttributeNode("action").value,ksblog.formToRequestString(form_obj),form_obj.method,fun)
}

ksblog._setObj=function(){
  if (loader.readyState==4)
  {
	var s=loader.responseText;
	document.getElementById(ksblog._mainlist).innerHTML=s;
	document.getElementById(ksblog._pagelist).innerHTML='';
	}
}
ksblog._setxx=function(){
  if (loader.readyState==4)
  {
	var s=loader.responseText;
	document.getElementById("xxlist").innerHTML=s;
	}
}

ksblog.loadlog=function()
{
	document.getElementById(ksblog._pagelist).style.display='';
	Page(1,"log",this._username);
}
ksblog.loadguest=function()
{
	document.getElementById(ksblog._pagelist).style.display='';
   Page(1,"guest",this._username);	
}
ksblog.loadxx=function(channelid,username)
{  
	document.getElementById(ksblog._pagelist).style.display='';
   Page(1,"xx&channelid="+channelid,username);	
}
ksblog.loadintro=function()
{
	   document.getElementById(ksblog._pagelist).style.display='none';
 	   var _request='action=intro&username='+this._username;
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setObj");
}
ksblog.loadproduct=function()
{	    
	   document.getElementById(ksblog._pagelist).style.display='none';
 	   var _request='action=product&username='+this._username;
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setObj");
}
ksblog.loadjob=function()
{	    
	   document.getElementById(ksblog._pagelist).style.display='none';
 	   var _request='action=job&username='+this._username;
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setObj");
}
ksblog.loadnews=function()
{
	document.getElementById(ksblog._pagelist).style.display='';
   Page(1,"news",this._username);	
}
ksblog.loadnewscontent=function(username,newsid)
{
	   document.getElementById(ksblog._pagelist).style.display='none';
 	   var _request='action=newscontent&username='+username+'&newsid='+newsid;
       var _method='post';
       this.ajaxLoadPage(ksblog._url,_request,_method,"ksblog._setObj");
}
ksblog.loadshortintro=function(username)
{   this._username=username;
	this.loadintro();
}
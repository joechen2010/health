/*================================================================
Created:2007-5-29
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS V6.0
Service QQ��111394,54004407
==================================================================*/

var installdir='/';           //ϵͳ��װĿ¼������ȷ��д���簲װ��kesionĿ¼�£�����Ϊ installdir='/kesion/'
function LabelAjax()
{
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
function getlabeltag(){
    var labelItem = document.getElementsByTagName("span"); 
    for(var i=0; i<labelItem.length; i++){
        var obj = labelItem[i];   
		if (typeof(obj.id)!="undefined"&&(obj.id.substring(0,2)=="ks"||obj.id.substring(0,3)=="SQL"))
		{
		  if (obj.id.substring(0,2)=="ks")
		   {
			  var idarr=obj.id.split('_');
			  var labelid=idarr[0].replace("ks","");
			  var typeid=idarr[1];
			  var classid=idarr[2];
			  var infoid=idarr[3];
			  var channelid=idarr[4];
			  try{  
			  getlabelcontent("plus/ajax.asp",obj,labelid,"Label",typeid,channelid,classid,infoid)
			   }catch(e){}
		   }
		   else if (obj.id.substring(0,3)=="SQL")
		   {
			   var p=obj.id.substring(obj.id.indexOf("ksr")+3);
			   var parr=p.split('p');
			   var classid=0;
			   var infoid=0;
			   var channelid=0;
			   if (p!='') 
			   {  infoid=parr[0];
			      classid=parr[1];
			   }
			try{getlabelcontent("plus/ajax.asp",obj,obj.id,"SQL",0,channelid,classid,infoid);   
			 }catch(e){}
		   }
		}
  }
}
function getlabelcontent(posturl,obj,labelid,action,typeid,channelid,classid,infoid)
{ 
  try{
	var ksxhr=new LabelAjax(); 
	var senddata="?action="+action+"&labelid="+escape(labelid)+"&labtype="+typeid+"&channelid=" +channelid+"&classid="+classid+"&infoid="+infoid+getUrlParam();
	ksxhr.open("get",installdir+posturl+senddata,true);
    ksxhr.onreadystatechange=function(){
		  if(ksxhr.readyState==1)
				  {
					obj.innerHTML="<span align='center'><img src='"+installdir+"images/loading.gif'>���ڼ�������...</span>";
				  }
				  else if(ksxhr.readyState==2 || ksxhr.readyState==3)
				  {
				   obj.innerHTML="<span align='center'><img src='"+installdir+"images/loading.gif'>���ڶ�ȡ����...</span>";
				   }
				  else if(ksxhr.readyState==4)
				  {
					  
					 if (ksxhr.status==200)
					 {var s=ksxhr.responseText;
					  obj.innerHTML=s;
					 }
				  }
				}
	ksxhr.send(null);  
  }
  catch(e)
  {}
}
//ȡurl���Ĳ���
function getUrlParam()
{
	var URLParams = new Object() ;
	var aParams = document.location.search.substr(1).split('&') ;//substr(n,m)��ȡ�ַ���n��m,split('o')��oΪ���,�ָ��ַ���Ϊ����
	if(aParams!=''&&aParams!=null){
		var sum=new Array(aParams.length);//��������
		for (i=0 ; i < aParams.length ; i++) {
		sum[i]=new Array();
		var aParam = aParams[i].split('=') ;//�ԵȺŷָ�
		URLParams[aParam[0]] = aParam[1] ;
		sum[i][0]=aParam[0];
		sum[i][1]=aParam[1];
		}
		var p='';
		for(i=0;i<sum.length;i++)
		{
		  p=p+'&'+sum[i][0]+"="+sum[i][1]
		}
	   return p;
	}else{
	   return "";
	}
}
getlabeltag();
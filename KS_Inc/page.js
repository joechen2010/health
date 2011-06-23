/*================================================================
Created:2007-5-29
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

function Page(curPage,labelid,classid,installdir,url,refreshtype,specialid)
   {
   this.labelid=labelid;
   this.classid=classid;
   this.url=url;
   if (labelid.substring(0,5)=="{SQL_")
   {
	var slabelid=labelid.split('(')[0];
    slabelid=slabelid.replace("{","");
    this.c_obj="c_"+slabelid;
    this.p_obj="p_"+slabelid;
   }
   else
   {
   this.c_obj="c_"+labelid;
   this.p_obj="p_"+labelid;
   }
   this.installdir=installdir;
   this.refreshtype=refreshtype;
   this.specialid=specialid;
   this.page=curPage;
   loadData(1);
   }
function loadData(p)
{  this.page=p;
   var xhr=new PageAjax();
   var senddata=installdir+url+"?labelid="+escape(labelid)+"&classid="+classid+"&refreshtype="+refreshtype+"&specialid=" +specialid+"&curpage="+p+getUrlParam();
   xhr.open("get",senddata,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1)
			  {
				 if (p==1)
				eval('document.all.'+c_obj).innerHTML="<div align='center'><img src='"+installdir+"images/loading.gif'>�������ӷ�����...</div>";
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3)
			  {
				if (p==1)
				eval('document.all.'+c_obj).innerHTML="<div align='center'><img src='"+installdir+"images/loading.gif'>���ڶ�ȡ����...</div>";
			  }
			  else if(xhr.readyState==4)
			  {
			 if (xhr.status==200)
			 {
				  var pagearr=xhr.responseText.split("{ks:page}")
				  var pageparamarr=pagearr[1].split("|");
				  count=pageparamarr[0];    
				  perpagenum=pageparamarr[1];
				  pagecount=pageparamarr[2];
				  itemunit=pageparamarr[3];   
				  itemname=pageparamarr[4];
				  pagestyle=pageparamarr[5];
				  getObject(c_obj).innerHTML=pagearr[0];
				  pagelist();
			 }
			}
	   }
    xhr.send(null); 
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
function getObject(id) 
{
	if(document.getElementById) 
	{
		return document.getElementById(id);
	}
	else if(document.all)
	{
		return document.all[id];
	}
	else if(document.layers)
	{
		return document.layers[id];
	}
}

function pagelist()
{
 var n=1;	
 var statushtml=null;
 switch(parseInt(this.pagestyle))
 {
  case 1:	
     statushtml="��"+this.count+this.itemunit+" <a href=\"javascript:homePage(1);\" title=\"��ҳ\">��ҳ</a> <a href=\"javascript:previousPage()\" title=\"��һҳ\">��һҳ</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"��һҳ\">��һҳ</a> <a href=\"javascript:lastPage();\" title=\"���һҳ\">βҳ</a> ҳ��:<font color=red>"+this.page+"</font>/"+this.pagecount+"ҳ "+this.perpagenum+this.itemunit+this.itemname+"/ҳ";
		break;
  case 2:
	 statushtml="<a href='#'>"+this.pagecount+"ҳ/"+this.count+this.itemunit+"</a> <a href=\"javascript:homePage(1);\" title=\"��ҳ\"><span style='font-family:webdings;font-size:14px'>9</span></a> <a href=\"javascript:previousPage()\" title=\"��һҳ\"><span style='font-family:webdings;font-size:14px'>7</span></a>&nbsp;";
	 var startpage=1;
	 if (this.page==10)
	   startpage=2;
	 else if(this.page>10)
	   startpage=eval((parseInt(this.page/10)-1)*10+parseInt((this.page)%10)+2);
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this.page)
		   statushtml+="<a href=\"#\"><font color=\"#ff0000\">"+i+"</font></a>&nbsp;"
		  else
			statushtml+="<a href=\"javascript:turn("+i+")\">"+i+"</a>&nbsp;"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"��һҳ\"><font face=webdings>8</font></a> <a href=\"javascript:lastPage();\" title=\"���һҳ\"><span style='font-family:webdings;font-size:14px'>:</span></a>";
	 statushtml="<span class='kspage'>"+statushtml+"</span>";
	break;	 
  case 4:
	 statushtml="<table border='0' align='right'><tr><td><a class='prev' href='javascript:previousPage();'>��һҳ</a>";
	 statushtml+="<a class='prev' href='javascript:nextPage();'>��һҳ</a>";
	 statushtml+="<a class='prev' href='javascript:homePage(1);'>�� ҳ</a>";
	 var startpage=1;
	 if (this.page>7) startpage=page-5;
	 if (this.pagecout-this.page<5) startpage=this.pagecount-9;
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this.page)
		   statushtml+="<a href='javascript:void(0)' class='curr'><font color=\"#ff0000\">"+i+"</font></a>"
		  else
			statushtml+="<a class='num' href=\"javascript:turn("+i+")\">"+i+"</a>"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:lastPage();\" class='next' title=\"���һҳ\">ĩ ҳ</a><span>����" +this.pagecount+"ҳ</td></tr></table>";
	break;	 
  case 3:
     statushtml="��<font color=#ff000>"+this.page+"</font>ҳ ��"+this.pagecount+"ҳ <a href=\"javascript:homePage(1);\" title=\"��ҳ\"><<</a> <a href=\"javascript:previousPage()\" title=\"��һҳ\"><</a>&nbsp;<a href=\"javascript:nextPage()\" title=\"��һҳ\">></a> <a href=\"javascript:lastPage();\" title=\"���һҳ\">>></a> "+this.perpagenum+this.itemunit+this.itemname+"/ҳ";
   break;
 }
  if (parseInt(this.pagestyle)!=4){
	 statushtml+="&nbsp;��<select name=\"goto\" onchange=\"turn(parseInt(this.value));\">";
	  for(var i=1;i<=this.pagecount;i++){
		 if (i==this.page)
		 statushtml+="<option value='"+i+"' selected>"+i+"</option>";
		 else
		 statushtml+="<option value='"+i+"'>"+i+"</option>";
	  }	
	 statushtml+="</select>ҳ";
  }
	 getObject(this.p_obj).innerHTML=statushtml;
}
function homePage()
{
   if(this.page==1)
    alert("�Ѿ�����ҳ�ˣ�")
   else
   loadData(1);
} 
function lastPage()
{
   if(this.page==this.pagecount)
    alert("�Ѿ������һҳ�ˣ�")
   else
   loadData(this.pagecount);
} 
function previousPage()
{
   if (this.page>1)
      loadData(this.page-1);
   else
      alert("�Ѿ��ǵ�һҳ��");      
}

function nextPage()
{
   if(this.page<this.pagecount)
      loadData(this.page+1);
   else
      alert("�Ѿ������һҳ��");
}
function turn(i)
{
     loadData(i);
}
/*================================================================
Created:2007-6-15
Autor��linwenzhong
Copyright:www.Kesion.com  bbs.kesion.com
Version:KesionCMS V4.0
Service QQ��111394,54004407
==================================================================*/
 //�ռ����
 function SpacePage(curPage,action)
 {
   this._username = null;
   this._action   = action;
   this._c_obj    = "spacemain";
   this._p_obj    = "kspage";
   this._page     = curPage;
   this._url      = "ajax.asp";
   loadDate(1);
 }


//��ǰҳ,�������û���
function Page(curPage,action,username)
   {
   this._username = username;
   this._action   = action;
   this._c_obj    = ksblog._mainlist;
   this._p_obj    = ksblog._pagelist;
   this._page     = curPage;
   this._url      = ksblog._url;
   loadDate(1);
   }
 //��ǰҳ,��־ID
function CommentPage(curPage,id)
   {
   this._id        = id;
   this._action    = "Show&id="+id;
   this._url       = "Getcomment.asp";
   this._username  ="";
   this._c_obj="commentmainlist";
   this._p_obj="commentpagelist";
   this._page=curPage;
   loadDate(1);
 }
function GuestPage(curPage,action,username)
   {
   this._username = username;
   this._action   = action;
   this._c_obj    = "guestmain";
   this._p_obj    = "guestpage";
   this._page     = curPage;
   this._url      = ksblog._url;
   loadDate(1);
 }
 //Ȧ������
 function TeamPage(curPage,action)
 {
   this._username = null;
   this._action   = action;
   this._c_obj    = "teammain";
   this._p_obj    = "kspage";
   this._page     = curPage;
   this._url      = "groupajax.asp";
   loadDate(1);
 }
function loadDate(p)
{  this._page=p;
   var xhr=new ksblog.Ajax();
   xhr.open("get",_url+"?action="+_action+"&username="+escape(_username)+"&page="+p,true);
   xhr.onreadystatechange=function (){
	         if(xhr.readyState==1)
			  {
				document.getElementById(_c_obj).innerHTML="<div align='center'><img src='images/loading.gif'>���ڼ���...</div>";
			  }
			  else if(xhr.readyState==2 || xhr.readyState==3)
			  {
				  if (p==1)
				 document.getElementById(_c_obj).innerHTML="<div align='center'><img src='images/loading.gif'>���ڶ�ȡ����...</div>";
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
				  document.getElementById(_c_obj).innerHTML=pagearr[0];
				  pagelist();
			 }
			}
	   }
    xhr.send(null);  	
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
	 statushtml="��"+this.pagecount+"ҳ/"+this.count+this.itemunit+this.itemname+" <a href=\"javascript:homePage();\" title=\"��ҳ\"><font face=webdings>9</font></a> <a href=\"javascript:previousPage()\" title=\"��һҳ\"><font face=webdings>7</font></a>&nbsp;";
	 var startpage=1;
	 if (this._page>10)
	   startpage=(parseInt(this._page/10)-1)*10+parseInt(this._page%10)+1;
	  for(var i=startpage;i<=this.pagecount;i++){ 
		  if (i==this._page)
		   statushtml+="<a href=\"javascript:turn("+i+")\"><font color=\"#ff0000\">"+i+"</font></a>&nbsp;"
		  else
			statushtml+="<a href=\"javascript:turn("+i+")\">"+i+"</a>&nbsp;"
			n=n+1;
		  if (n>10) break;
	  }
	 statushtml+="<a href=\"javascript:nextPage()\" title=\"��һҳ\"><font face=webdings>8</font></a> <a href=\"javascript:lastPage();\" title=\"���һҳ\"><font face=webdings>:</font></a>";
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
	// if (this.pagecount!="")
	// {
	 document.getElementById(this._p_obj).innerHTML=statushtml;
	// }
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
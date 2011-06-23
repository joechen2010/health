bInitialized=false;
var DocPopupContextMenu = null;
DocPopupContextMenu = window.createPopup();
var SelectedTD=null;
var SelectedTR=null;
var SelectedTBODY=null;
var SelectedTable=null;
function SetEditAreaStyle()
{  KS_EditArea.document.open();
   KS_EditArea.document.write('<head><link href="'+InstallDir+'KS_Editor/Editor_Style.CSS" type="text/css" rel="stylesheet"></head>');
   KS_EditArea.document.close();
}
function InitBtn(btn) 
{
	btn.onmouseover = ImageBtnMouseOver;
	btn.onmouseout = BtnMouseOut;
	btn.onmousedown = BtnMouseDown;
	btn.onmouseup = BtnMouseOut;
	btn.ondragstart = YCancelEvent;
	btn.onselectstart = YCancelEvent;
	btn.onselect = YCancelEvent;
	btn.disabled=false;
	return true;
}
function InsertMorePage(Index)
{
	KS_EditArea.focus();
	InsertHTMLStr('[NextPage]');
	KS_EditArea.focus();
}

function ImageBtnMouseOver() 
{
	var image = event.srcElement;
	image.className = "ToolBtnMouseOver";
	event.cancelBubble = true;
}

function BtnMouseOut() 
{
	var image = event.srcElement;
	image.className = "Btn";
	event.cancelBubble = true;
}

function BtnMouseDown() 
{
	var image = event.srcElement;
	image.className = "ToolBtnMouseDown";
	event.cancelBubble = true;
	event.returnValue=false;
	return false;
}

function YCancelEvent() 
{
	event.returnValue=false;
	event.cancelBubble=true;
	return false;
}
function LoadEditFile(Style)
{   //接收Style确定右键是否出现新闻属性
	KS_EditArea.document.body.contentEditable="true";
	KS_EditArea.document.onmouseup=new Function("return SearchObject(KS_EditArea.event);");
	KS_EditArea.document.oncontextmenu=new Function("return ShowMouseRightMenu(KS_EditArea.event,"+Style+");");
	KS_EditArea.focus();
}
var CurrMode='EDIT';
function setMode(NewMode,Style,ReplaceStr,InstallStr,AdminDirStr)   
{  var AdminBadStr="\\/"+AdminDirStr+"([A-Z]|[a-z]|\\/|)*.asp([A-Z]|[a-z]||\\?|=|&|;|[0-9])*#";
	 if (NewMode!=CurrMode)   
	 {   
		if (NewMode=='TEXT')
		{if (!confirm("警告！切换到纯文本模式会丢失您所有的HTML格式，您确认切换吗？")) return false;}
		    var sBody='';
		    switch(CurrMode)
		     {
			   case "CODE":
				    if (NewMode=="TEXT") sBody=KS_EditArea.document.body.innerText;
				    else sBody=KS_EditArea.document.body.innerText;
				    break;
			   case "TEXT":
				   sBody=KS_EditArea.document.body.innerText;
				   sBody=HTMLEncode(sBody);
				break;
			   case "EDIT":
			   case "VIEW":
				  if (NewMode=="TEXT") sBody=KS_EditArea.document.body.innerText;
				   else sBody=KS_EditArea.document.body.innerHTML;
				  break;
		    }
			sBody=sBody.replace(eval("/"+ReplaceStr+InstallStr+"/g"),'/');
			sBody=sBody.replace(eval("/"+ReplaceStr+"/g"),'');
			sBody=sBody.replace(eval("/"+AdminBadStr+"/g"),'#');
		    document.all["editor_CODE"].className='ModeBarBtnOff';
	     	document.all["editor_EDIT"].className='ModeBarBtnOff';
		    document.all["editor_TEXT"].className='ModeBarBtnOff';
		    document.all["editor_VIEW"].className='ModeBarBtnOff';
		    document.all["editor_"+NewMode].className='ModeBarBtnOn';
	 switch (Style)
     { case 2:
	   case 3:
	   case 4:
		var documentElementStr='<html>'+KS_EditArea.document.documentElement.innerHTML+'</html>';
		switch (NewMode)
		{
			case "CODE":
				KS_EditArea.document.designMode="On";
				KS_EditArea.document.open();
				KS_EditArea.document.write(documentElementStr);
				KS_EditArea.document.body.innerText=sBody;
				KS_EditArea.document.body.contentEditable="true";
				KS_EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "EDIT":
				KS_EditArea.document.designMode="On";
				KS_EditArea.document.open();
				KS_EditArea.document.write(documentElementStr);
				KS_EditArea.document.body.innerHTML=sBody;
				KS_EditArea.document.body.contentEditable="true";
				KS_EditArea.document.close();
				ShowTableBorders();
				DisabledAllBtn(false);
				break;
			case "TEXT":
				KS_EditArea.document.designMode="On";
				KS_EditArea.document.open();
				KS_EditArea.document.write(documentElementStr);
				KS_EditArea.document.body.innerText=sBody;
				KS_EditArea.document.body.contentEditable="true";
				KS_EditArea.document.close();
				DisabledAllBtn(true);
				break;
			case "VIEW":
				KS_EditArea.document.designMode="off";
				KS_EditArea.document.open();
				KS_EditArea.document.write(documentElementStr);
				KS_EditArea.document.body.innerHTML=sBody;
				KS_EditArea.document.body.contentEditable="false";
				KS_EditArea.document.close();
				DisabledAllBtn(true);
				break;
		}
		CurrMode=NewMode;
		if (NewMode!='EDIT') EmptyShowObject(true);
		else {EmptyShowObject(false);LoadEditFile(Style);}
		break;
	  default :
		    switch (NewMode)
		    {
		    	case "CODE":
				  KS_EditArea.document.designMode="On";
				  KS_EditArea.document.open();
				  KS_EditArea.document.write("<head><link href="+InstallDir+"KS_Editor/Editor_Style.CSS type=text/css rel=stylesheet></head><body MONOSPACE>"+sBody);
				  sBody=FormatHtml(ReplaceUrl(ReplaceImgToScript(Resumeblank(sBody))));
				  KS_EditArea.document.body.innerText=sBody;
				  KS_EditArea.document.body.contentEditable="true";
				  LoadEditFile(Style);
				  KS_EditArea.document.close();
				  DisabledAllBtn(true);
				  break;
			    case "EDIT":
				  KS_EditArea.document.designMode="On";
				  KS_EditArea.document.open();
				  KS_EditArea.document.write("<head><link href="+InstallDir+"KS_Editor/Editor_Style.CSS type=text/css rel=stylesheet></head><body MONOSPACE>"+sBody);
			      KS_EditArea.document.body.contentEditable="true";
				  KS_EditArea.document.execCommand("2D-Position",true,true);
				  KS_EditArea.document.execCommand("MultipleSelection", true, true);
				  KS_EditArea.document.execCommand("LiveResize", true, true);
				  LoadEditFile(Style);
				  KS_EditArea.document.close();
				  ShowTableBorders();
				  DisabledAllBtn(false);
				  break;
			    case "TEXT":
				  KS_EditArea.document.designMode="On";
				  KS_EditArea.document.open();
			      KS_EditArea.document.write("<head><link href="+InstallDir+"KS_Editor/Editor_Style.CSS type=text/css rel=stylesheet></head><body MONOSPACE>"+sBody);
				  KS_EditArea.document.body.innerText=sBody;
				  KS_EditArea.document.body.contentEditable="true";
				  LoadEditFile(Style);
				  KS_EditArea.document.close();
				  DisabledAllBtn(true);
				  break;
			    case "VIEW":
				  KS_EditArea.document.designMode="off";
				  KS_EditArea.document.open();
				  KS_EditArea.document.write("<head><link href="+InstallDir+"KS_Editor/Editor_Style.CSS type=text/css rel=stylesheet></head><body MONOSPACE>"+sBody);
				  KS_EditArea.document.body.contentEditable="false";
				  KS_EditArea.document.close();
				  DisabledAllBtn(true);
				break;
		    }
		  CurrMode=NewMode;
		  EmptyShowObject(true);
    }
   }
	KS_EditArea.focus();
}
function EmptyShowObject(Flag)
{
	document.all.ShowObject.disabled=Flag;
}
	// 替换特殊字符
function HTMLEncode(text){
	text = text.replace(/&/g, "&amp;") ;
	text = text.replace(/"/g, "&quot;") ;
	text = text.replace(/</g, "&lt;") ;
	text = text.replace(/>/g, "&gt;") ;
	text = text.replace(/'/g, "&#146;") ;
	text = text.replace(/\ /g,"&nbsp;");
	text = text.replace(/\n/g,"<br>");
	text = text.replace(/\t/g,"&nbsp;&nbsp;&nbsp;&nbsp;");
	return text;
}
function ShowMouseRightMenu(event,Style)
{
	var width=0;
	var height=0;
	var lefter=event.clientX;
	var topper=event.clientY;
	var ObjPopDocument=DocPopupContextMenu.document;
	var ObjPopBody=DocPopupContextMenu.document.body;
	var MenuStr='';
	width+=138;
	height+=120;
		//判断是否是文章管理调用,若是增加一些功能
   if ((Style==1)&&(CurrMode=='EDIT'))
   	{	
       //判断是否有图片选中，有则初始化快捷菜单
       if (ImageSelected())
		{
         height+=16;
	     MenuStr+=GetMenuFunRowStr("SetPicArticle()", "设置为图片文章","Pic_Aritlce.gif");
		 MenuStr+=GetMenuSplitRowStr();	
	     MenuStr+=GetMenuFunRowStr("SetEditPic()", "编辑图片属性","Pic_Aritlce.gif");
		 MenuStr+=GetMenuSplitRowStr();	
		 
		}
		//判断是否有选中的文本,有则初始化快捷菜单
        var RangeType = KS_EditArea.document.selection.type;	
        if (RangeType=="Text")  
	    { height+=155;
          MenuStr+=GetMenuFunRowStr("SetNewsAttribute(1)", "设置为正标题","News_Title.gif");
		  MenuStr+=GetMenuFunRowStr('SetNewsAttribute(2)','设置为完整标题','News_SubTitle.gif');
    	  MenuStr+=GetMenuFunRowStr('SetNewsAttribute(3)','设置为关键字','News_Keyword.gif');
		  MenuStr+=GetMenuSplitRowStr();
		  MenuStr+=FormatMenuRow("bold", "加粗","bold.gif");
		  MenuStr+=FormatMenuRow("italic", "斜体","italic.gif");
		  MenuStr+=FormatMenuRow("underline", "下划线","underline.gif");
	      MenuStr+=GetMenuFunRowStr("TextColor()", "文字颜色","TextColor.gif");
		  MenuStr+=GetMenuFunRowStr("TextBGColor()", "文字背景色","fgbgcolor.gif");
		  MenuStr+=GetMenuSplitRowStr();
	   }
	}
    MenuStr+=FormatMenuRow("cut", "剪切","Cut.gif");
	MenuStr+=FormatMenuRow("copy", "复制","Copy.gif");
	MenuStr+=FormatMenuRow("paste", "常规粘贴","Paste.gif");
	MenuStr+=FormatMenuRow("delete", "删除","Del.gif");
	MenuStr+=GetMenuSplitRowStr();

	MenuStr+=FormatMenuRow("selectall", "全选","SelectAll.gif");
	MenuStr+=GetMenuFunRowStr("SearchStr('"+InstallDir+"')",'查找替换...','find.gif');
	MenuStr="<TABLE border=0 cellpadding=0 cellspacing=0 width=111><tr><td width=111 class=RightBg><TABLE border=0 cellpadding=0 cellspacing=0 style=\"font-size:9pt;\">"+MenuStr
	MenuStr=MenuStr+"<\/TABLE><\/td><\/tr><\/TABLE>";
	MenuStr="<TABLE Border=0 cellpadding=0 cellspacing=0 class=Menu width=138><tr><td width=20 bgcolor=#0072BC><img src="+InstallDir+"KS_Editor/images/about.gif><\/td><td>"+MenuStr+"<\/td><\/tr><\/TABLE>"
	ObjPopDocument.open();
	ObjPopDocument.write("<head><link href='"+InstallDir+"KS_Editor/ContextMenu.css' type=\"text/css\" rel=\"stylesheet\"></head><body scroll=\"no\" onConTextMenu=\"event.returnValue=false;\">"+MenuStr);
	ObjPopDocument.close();
	height+=5;
	if(lefter+width > document.body.clientWidth) lefter=lefter-width;
	DocPopupContextMenu.show(lefter, topper, width, height, KS_EditArea.document.body);
	return false;
}
//返用户自定义的方法或函数
function GetMenuFunRowStr(FunStr,MenuDescription,MenuImage)
{  var MenuFunRowStr='';
	MenuFunRowStr="<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=111><tr><td valign=middle height=18 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
	MenuFunRowStr+=" onClick=\"parent."+FunStr+";parent.DocPopupContextMenu.hide();\">";
	if (MenuImage!="")
	{
		MenuFunRowStr+="&nbsp;<img border=0 src='"+InstallDir+"KS_Editor/images/"+MenuImage+"' align=absmiddle>&nbsp;";
	}
	else
	{
		MenuFunRowStr+="&nbsp;";
	}
	MenuFunRowStr+=MenuDescription+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
	return MenuFunRowStr;

}
//分隔线
function GetMenuSplitRowStr()
{
  	var MenuSplitRowStr='';
    MenuSplitRowStr="<tr><td height=1 bgcolor=#cccccc></td></tr>";
	return  MenuSplitRowStr;
}
//返回系统方法或函数
function GetMenuRowStr(DisabledStr, MenuOperation, MenuImage, MenuDescripion)
{
	var MenuRowStr='';
	MenuRowStr="<tr><td align=center valign=middle><TABLE border=0 cellpadding=0 cellspacing=0 width=111><tr "+DisabledStr+"><td valign=middle height=18 class=MouseOut onMouseOver=this.className='MouseOver'; onMouseOut=this.className='MouseOut';";
	if (DisabledStr==''){
		MenuRowStr += " onclick=\"parent."+MenuOperation+";parent.DocPopupContextMenu.hide();\"";
	}
	MenuRowStr+=">"
	if (MenuImage!="")
	{
		MenuRowStr+="&nbsp;<img border=0 src='"+InstallDir+"KS_Editor/images/"+MenuImage+"' align=absmiddle "+DisabledStr+">&nbsp;";
	}
	else
	{
		MenuRowStr+="&nbsp;";
	}
	MenuRowStr+=MenuDescripion+"<\/td><\/tr><\/TABLE><\/td><\/tr>";
	return MenuRowStr;

}
function FormatMenuRow(MenuStr,MenuDescription,MenuImage)
{
	var DisabledStr='';
	var ShowMenuImage='';
	if (!KS_EditArea.document.queryCommandEnabled(MenuStr))
	{
		DisabledStr="disabled";
	}
	var MenuOperation="Format('"+MenuStr+"')";
	if (MenuImage)
	{
		ShowMenuImage=MenuImage;
	}
	return GetMenuRowStr(DisabledStr,MenuOperation,ShowMenuImage,MenuDescription)
}
function SearchObject()
{
	UpdateToolbar();
}
function MouseRightMenuItem(CommandString, CommandId)
{
	this.CommandString = CommandString;
	this.CommandId = CommandId;
}
function GetKS_EditAreaSelectionType()
{
	return KS_EditArea.document.selection.type;
}
var ContextMenu = new Array();
function ExeEditAttribute(InstallDir)
{
	OpenWindow(InstallDir+'KS_Editor/AttributeWindow.asp',360,120,window)
	KS_EditArea.focus();
}
function InsertHTMLStr(Str)
{
	KS_EditArea.focus();
	if (KS_EditArea.document.selection.type.toLowerCase() != "none")
	{
		KS_EditArea.document.selection.clear() ;
	}
	KS_EditArea.document.selection.createRange().pasteHTML(Str) ; 
	KS_EditArea.focus();
	ShowTableBorders();
}
function QueryCommand(CommandID)
{
	var State=KS_EditArea.QueryStatus(CommandID)
	if (State==3) return true;
	else return false;
}
function Format(Operation,Val) 
{
	KS_EditArea.focus();
	if (Val=="RemoveFormat")
	 {
		Operation=Val;
		Val=null;
	}
	if (Val==null) KS_EditArea.document.execCommand(Operation);
	else KS_EditArea.document.execCommand(Operation,"",Val);
	KS_EditArea.focus();
}
//设图片为图片文章的地址
function SetPicArticle()
{
	if (KS_EditArea.document.selection.type=="Control")
	{
		var oControlRange=KS_EditArea.document.selection.createRange();
		if (oControlRange(0).tagName.toUpperCase()=="IMG")
		{    
			selectedImage=KS_EditArea.document.selection.createRange()(0);
			//alert(selectedImage.src);
			//parent.document.NewsForm.PicNews.checked=true;
			//parent.ShowPicNews();
	        parent.document.getElementById('PhotoUrl').value=selectedImage.src;
		}	
	}
}
//编辑图片属性
function SetEditPic()
{
	if (KS_EditArea.document.selection.type=="Control")
	{
		var oControlRange=KS_EditArea.document.selection.createRange();
		if (oControlRange(0).tagName.toUpperCase()=="IMG")
		{    
			selectedImage=KS_EditArea.document.selection.createRange()(0);
			var ReturnValue=OpenWindow('../KS_Editor/InsertPicture.asp?FromUrl=1&ChannelID=1&action=edit&src='+selectedImage.src+'&width='+selectedImage.width+'&height='+selectedImage.height+'&border='+selectedImage.border+'&alt='+selectedImage.alt+'&vspace='+selectedImage.vspace+'&hspace='+selectedImage.hspace+'&align='+selectedImage.align,420,370,window);
			if (ReturnValue!='')
			{
				var TempArray=ReturnValue.split("$$$");
				InsertHTMLStr(TempArray[0]);
			}

		}	
	}
}
//常规粘贴
function Paste()
{ 
  KS_EditArea.focus();
  KS_EditArea.document.execCommand('Paste');
  KS_EditArea.focus();
}
//纯文本粘贴
function PasteText()
{
	KS_EditArea.focus();
	var sText = HTMLEncode(clipboardData.getData("Text")) ;
	InsertHTMLStr(sText);
	KS_EditArea.focus();
}
//设置新闻属性
function SetNewsAttribute(OpID)
{
	KS_EditArea.focus();
	var EditRange = KS_EditArea.document.selection.createRange();
	if (OpID==1)
	parent.document.NewsForm.title.value=EditRange.text;
	else if(OpID==2)
	parent.document.NewsForm.Fulltitle.value=EditRange.text;
	else if(OpID==3)
	InsertKeyWords(parent.document.NewsForm.KeyWords,EditRange.text);
	KS_EditArea.focus();
}
// 全屏编辑
function FullScreen(InstallDir,Style,ChannelID)
{
	
	if (CurrMode!='EDIT') 
	 {
	  alert('需转换为编辑状态才能使用全屏编辑功能!');
	  return;
	 }
	window.open(InstallDir+'KS_Editor/fullscreen.asp?ChannelID='+ChannelID+'&Style='+Style, '','toolbar=no, menubar=no, top=0,left=0,width=1024,height=768, scrollbars=no, resizable=no,location=no, status=no');
}
function TextBGColor()
{
	KS_EditArea.focus();
	var EditRange = KS_EditArea.document.selection.createRange();
	var RangeType = KS_EditArea.document.selection.type;
	if (RangeType!="Text")
	{
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/SelectColor.asp',230,190,window);
	if (ReturnValue!=null)
	{
		EditRange.pasteHTML("<span style='background-color:"+ReturnValue+"'>"+EditRange.text+"</span> ");
		EditRange.select();
	}
	KS_EditArea.focus();
}
function Print(CommandID)
{
	KS_EditArea.focus();
	//alert(KS_EditArea.QueryStatus(CommandID));
	if (KS_EditArea.QueryStatus(CommandID)!=3) KS_EditArea.ExecCommand(CommandID,0);
	KS_EditArea.focus();
}
function InsertTable(InstallDir)
{
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertTable.asp',250,220,window);
	InsertHTMLStr(ReturnValue);
	KS_EditArea.focus();
}
function InsertPage(InstallDir)
{
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertPage.asp',320,110,window);
	InsertHTMLStr(ReturnValue);
	KS_EditArea.focus();
}
function InsertExcel()
{
	KS_EditArea.focus();
	var TempStr="<object classid='clsid:0002E510-0000-0000-C000-000000000046' id='Spreadsheet1' codebase='file:\\Bob\software\office2000\msowc.cab' width='100%' height='250'><param name='EnableAutoCalculate' value='-1'><param name='DisplayTitleBar' value='0'><param name='DisplayToolbar' value='-1'><param name='ViewableRange' value='1:65536'></object>";
	InsertHTMLStr(TempStr);
	KS_EditArea.focus();
}
function InsertMarquee(InstallDir)
{
	KS_EditArea.focus();
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertMarquee.asp',260,50,window); 
	InsertHTMLStr(ReturnValue);
	KS_EditArea.focus();
}
function Calculator(InstallDir)
{
	KS_EditArea.focus();
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/Calculator.asp',160,180,window);
	if (ReturnValue!=null)
	{
		var TempArray,ParameterA,ParameterB;
		TempArray=ReturnValue.split("*")
		ParameterA=TempArray[0];
		InsertHTMLStr(ParameterA);
	}
	KS_EditArea.focus();
}
function InsertDate()
{
	KS_EditArea.focus();
	var NowDate = new Date();
	var FormateDate=NowDate.getYear()+"年"+(NowDate.getMonth() + 1)+"月"+NowDate.getDate() +"日";
	InsertHTMLStr(FormateDate);
	KS_EditArea.focus();
}
function InsertTime()
{
	KS_EditArea.focus();
	var NowDate=new Date();
	var FormatTime=NowDate.getHours() +":"+NowDate.getMinutes()+":"+NowDate.getSeconds();
	InsertHTMLStr(FormatTime);
	KS_EditArea.focus();
}
function InsertFrame(InstallDir)
{
	KS_EditArea.focus();
	var ReturnVlaue =OpenWindow(InstallDir+'KS_Editor/InsertFrame.asp',280,118,window);
	if (ReturnVlaue != null)
	{
		InsertHTMLStr(ReturnVlaue);
	}
	KS_EditArea.focus();
}
function InsertBR(Index)
{
	KS_EditArea.focus();
	InsertHTMLStr('<br>');
	KS_EditArea.focus();
}
function DelAllHtmlTag()
{
	var TempStr;
	TempStr=KS_EditArea.document.body.innerHTML;
	var re=/<\/*[^<>]*>/ig
	TempStr=TempStr.replace(re,'');
	KS_EditArea.document.body.innerHTML=TempStr;
}
function AbortArticle(InstallDir)
{
  var arr = OpenWindow(InstallDir+'KS_Editor/Abort.asp',220,100,window);
}
function InsertSymbol(InstallDir)
{
  var ReturnValue = OpenWindow(InstallDir+'KS_Editor/InsertTsfh.asp',300,190,window); 
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  KS_EditArea.focus();
}
function InsertPictureFromUp(ImgSrc)
{
InsertHTMLStr('<img src="'+ImgSrc+'" border="0"/>');	
}
function InsertFileFromUp(FileList,InstallDir)
{
	Files=FileList.split("|");
	for(var i=0;i<Files.length-1;i++)
	{     var ext=getFilePic(Files[i]);
	      var files=Files[i].split('/');
		  var file=files[files.length-1];
		  var br='';
		  if (i!=Files.length-1) br='<br />';
          InsertHTMLStr("<img border=0 src='"+InstallDir+"KS_Editor/images/FileIcon/"+ext+"'> <a href="+Files[i]+" class='newsContent' target='_blank'>[点击浏览该文件:"+file+"]</a>"+br);	
     }
}
// 按文件扩展名取图，并产生链接
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




function InsertPicture(FromUrl,InstallDir,ChannelID)
{		
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertPicture.asp?FromUrl='+FromUrl+'&ChannelID='+ChannelID,420,370,window);
	if (ReturnValue!='')
	{
		var TempArray=ReturnValue.split("$$$");
		InsertHTMLStr(TempArray[0]);
	}
}
function InsertFlash(FromUrl,InstallDir,ChannelID)
{ 
  var ReturnValue = OpenWindow(InstallDir+'KS_Editor/InsertFlash.asp?FromUrl='+FromUrl+'&ChannelID='+ChannelID,400,320,window); 
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  KS_EditArea.focus();
}
function InsertVideo(FromUrl,InstallDir,ChannelID)
{
  var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertVideo.asp?FromUrl='+FromUrl+'&ChannelID='+ChannelID,400,320,window);
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  KS_EditArea.focus();
}
function InsertRM(FromUrl,InstallDir,ChannelID)
{
  var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertRM.asp?FromUrl='+FromUrl+'&ChannelID='+ChannelID,400,320,window);
  if (ReturnValue!='')
  {
    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }
  KS_EditArea.focus();
}
function InsertUpFile(FromUrl,InstallDir,ChannelID)
{
  var ReturnValue=OpenWindow(InstallDir+'KS_Editor/InsertUpFile.asp?FromUrl='+FromUrl+'&ChannelID='+ChannelID,400,120,window);
  if (ReturnValue!='')
  {    var TempArray=ReturnValue.split("$$$");
    InsertHTMLStr(TempArray[0]);
  }  KS_EditArea.focus();
} 
function SpecialHR(InstallDir)
{
	KS_EditArea.focus();
	var ReturnValue = OpenWindow(InstallDir+'KS_Editor/InsertSpecialHR.asp',320,120,window); 
	if (ReturnValue!= null) InsertHTMLStr(ReturnValue);
	KS_EditArea.focus();
}
function InsertHR()
{
	KS_EditArea.focus();
	InsertHTMLStr('<hr>');
	KS_EditArea.focus();
}
var BorderShown=1;
function ShowTableBorders()
{
	AllTables=KS_EditArea.document.body.getElementsByTagName("TABLE");
	for(var i=0;i<AllTables.length;i++)
	{
		if ((AllTables[i].border==null)||(AllTables[i].border=='0'))
		{
			AllTables[i].runtimeStyle.borderTop=AllTables[i].runtimeStyle.borderLeft="1px dotted #709FCB";
			AllRows = AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++)
			{
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++)
				{
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom="1px dotted #709FCB";
				}
			}
		}
		else
		{
			AllTables[i].runtimeStyle.borderTop='';
			AllRows=AllTables[i].rows;
			for(var y=0;y<AllRows.length;y++)
			{
				AllCells=AllRows[y].cells;
				for(var x=0;x<AllCells.length;x++)
				{
					AllCells[x].runtimeStyle.borderRight=AllCells[x].runtimeStyle.borderBottom='';
				}
			}
		}
	}
  BorderShown=BorderShown?0:1;
}
function ImageSelected()
{
	KS_EditArea.focus();
	if (KS_EditArea.document.selection.type=="Control")
	{
		var oControlRange=KS_EditArea.document.selection.createRange();
		if (oControlRange(0).tagName.toUpperCase()=="IMG")
		{
			selectedImage=KS_EditArea.document.selection.createRange()(0);
			return true;
		}	
	}
}
function TextColor()
{	
	KS_EditArea.focus();
	var EditRange = KS_EditArea.document.selection.createRange();
	var RangeType = KS_EditArea.document.selection.type;
	if (RangeType!="Text")
	{
		alert("请先选择一段文字！");
		return;
	}
	var ReturnValue=OpenWindow(InstallDir+'KS_Editor/SelectColor.asp',230,190,window);
	if (ReturnValue!=null)
	{
		EditRange.pasteHTML("<font color='"+ReturnValue+"'>"+EditRange.text+"</font>");
		EditRange.select();
	}
	KS_EditArea.focus();
}
function PicAndTextArrange(InstallDir)
{
	if(ImageSelected())
	{
		sPrePos=selectedImage.style.position;
		var ReturnValue=OpenWindow(InstallDir+'KS_Editor/SelectPicStyle.asp',380,130,window);
		if(ReturnValue)
		{
			for(key in ReturnValue)
			if(key=='style') for(sub_key in ReturnValue.style) selectedImage.style[sub_key]=ReturnValue.style[sub_key];
			else selectedImage[key]=ReturnValue[key];
			if(!ReturnValue.align) selectedImage.removeAttribute('align');
			if(sPrePos.match(/^absolute$/i) && !selectedImage.style.position.match(/^absolute$/i))
			{
				sFired = selectedImage.parentElement;
				while(!sFired.tagName.match(/^table$|^body$/i))
				sFired = sFired.parentElement;
				if(sFired.tagName.match(/^table$/i) && sFired.style.position.match(/absolute/i));
				sFired.outerHTML=selectedImage.outerHTML;
			}
			else
			{
				if(!sPrePos.match(/^absolute$/i) && selectedImage.style.position.match(/^absolute$/i)) selectedImage.outerHTML='<table style="position: absolute;"><tr><td>' + selectedImage.outerHTML + '</td></tr></table>';
			}
		}
	}
	else alert('请选择图片');
}
function GetAllAncestors()
{
	var p = GetParentElement();
	var a = [];
	while (p && (p.nodeType==1)&&(p.tagName.toLowerCase()!='body'))
	{
		a.push(p);
		p=p.parentNode;
	}
	a.push(KS_EditArea.document.body);
	return a;
}
function GetParentElement()
{
	var sel=GetSelection();
	var range=CreateRange(sel);
	switch (sel.type)
	{
		case "Text":
		case "None":
			return range.parentElement();
		case "Control":
			return range.item(0);
		default:
			return KS_EditArea.document.body;
	}
}
function GetSelection()
{
	return KS_EditArea.document.selection;
}
function CreateRange(sel)
{
	return sel.createRange();
}
function UpdateToolbar()
{
	var ancestors=null;
	ancestors=GetAllAncestors();
	document.all.ShowObject.innerHTML='&nbsp;';
	for (var i=ancestors.length;--i>=0;)
	{
		var el = ancestors[i];
		if (!el) continue;
		var a=document.createElement("span");
		a.href="#";
		a.el=el;
		a.editor=this;
		if (i==0)
		{
			a.className='AncestorsMouseUp';
			EditControl=a.el;
		}
		else a.className='AncestorsStyle';
		a.onmouseover=function()
		{
			if (this.className=='AncestorsMouseUp') this.className='AncestorsMouseUpOver';
			else if (this.className=='AncestorsStyle') this.className='AncestorsMouseOver';
		};
		a.onmouseout=function()
		{
			if (this.className=='AncestorsMouseUpOver') this.className='AncestorsMouseUp';
			else if (this.className=='AncestorsMouseOver') this.className='AncestorsStyle';
		};
		a.onmousedown=function(){this.className='AncestorsMouseDown';};
		a.onmouseup=function(){this.className='AncestorsMouseUpOver';};
		a.ondragstart=YCancelEvent;
		a.onselectstart=YCancelEvent;
		a.onselect=YCancelEvent;
		a.onclick=function()
		{
			this.blur();
			SelectNodeContents(this);
			return false;
		};
		var txt='<'+el.tagName.toLowerCase();
		a.title=el.style.cssText;
		if (el.id) txt += "#" + el.id;
		if (el.className) txt += "." + el.className;
		txt=txt+'>';
		a.appendChild(document.createTextNode(txt));
		document.all.ShowObject.appendChild(a);
	}
}
function SelectNodeContents(Obj,pos)
{
	var node=Obj.el;
	EditControl=node;
	for (var i=0;i<document.all.ShowObject.children.length;i++)
	{
		if (document.all.ShowObject.children(i).className=='AncestorsMouseUp') document.all.ShowObject.children(i).className='AncestorsStyle';
	}
	//Obj.className='AncestorsMouseUp';
	KS_EditArea.focus();
	var range;
	var collapsed=(typeof pos!='undefined');
	range = KS_EditArea.document.body.createTextRange();
	range.moveToElementText(node);
	(collapsed) && range.collapse(pos);
	range.select();
}
function DeleteHTMLTag()
{
	var AvaiLabelDeleteTagName='p,a,div,span';
	if (EditControl!=null)
	{
		var DeleteTagName=EditControl.tagName.toLowerCase();
		if (AvaiLabelDeleteTagName.indexOf(DeleteTagName)!=-1)
		{
			EditControl.parentElement.innerHTML=EditControl.innerHTML;
		}
	}
	UpdateToolbar();
	ShowTableBorders();
}
function InsertRow()
{
	if (CursorInTableCell())
	{
		var SelectColsNum=0;
		var AllCells=SelectedTR.cells;
		for (var i=0;i<AllCells.length;i++)
		{
		 	SelectColsNum=SelectColsNum+AllCells[i].getAttribute('colSpan');
		}
		var NewTR=SelectedTable.insertRow(SelectedTR.rowIndex);
		for (i=0;i<SelectColsNum;i++)
		{
		 	NewTD=NewTR.insertCell();
			NewTD.innerHTML="&nbsp;";
		}
	}
	ShowTableBorders();	
}
function InsertColumn()
{
   	if (CursorInTableCell())
	{
		var MoveFromEnd=(SelectedTR.cells.length-1)-(SelectedTD.cellIndex);
		var AllRows=SelectedTable.rows;
		for (i=0;i<AllRows.length;i++)
		{
			RowCount=AllRows[i].cells.length-1;
			Position=RowCount-MoveFromEnd;
			if (Position<0)
			{
				Position=0;
			}
			NewCell=AllRows[i].insertCell(Position);
			NewCell.innerHTML="&nbsp;";
		}
		ShowTableBorders();
	}	
}
function DeleteRow()
{
	if (CursorInTableCell())
	{
		SelectedTable.deleteRow(SelectedTR.rowIndex);
	}
}
function DeleteColumn()
{
   	if (CursorInTableCell())
	{
		var MoveFromEnd=(SelectedTR.cells.length-1)-(SelectedTD.cellIndex);
		var AllRows=SelectedTable.rows;
		for (var i=0;i<AllRows.length;i++)
		{
			var EndOfRow=AllRows[i].cells.length-1;
			var Position=EndOfRow-MoveFromEnd;
			if (Position<0) Position=0;
			var AllCellsInRow=AllRows[i].cells;
			if (AllCellsInRow[Position].colSpan>1) AllCellsInRow[position].colSpan=AllCellsInRow[position].colSpan-1;
			else AllRows[i].deleteCell(Position);
		}
	}
}
function MergeColumn()
{
	if (CursorInTableCell())
	{
		var RowSpanTD=SelectedTD.getAttribute('rowSpan');
		AllRows=SelectedTable.rows;
		if (SelectedTR.rowIndex+1!=AllRows.length)
		{
			var AllCellsInNextRow=AllRows[SelectedTR.rowIndex+SelectedTD.rowSpan].cells;
			var AddRowSpan=AllCellsInNextRow[SelectedTD.cellIndex].getAttribute('rowSpan');
			var MoveTo=SelectedTD.rowSpan;
			if (!AddRowSpan) AddRowSpan=1;
			SelectedTD.rowSpan=SelectedTD.rowSpan+AddRowSpan;
			AllRows[SelectedTR.rowIndex+MoveTo].deleteCell(SelectedTD.cellIndex);
		}
		else alert('请重新选择');
	}
	ShowTableBorders();
}
function MergeRow()
{
	if (CursorInTableCell())
	{
		var ColSpanTD=SelectedTD.getAttribute('colSpan');
		var AllCells=SelectedTR.cells;
		if (SelectedTD.cellIndex+1!=SelectedTR.cells.length)
		{
			var AddColspan=AllCells[SelectedTD.cellIndex+1].getAttribute('colSpan');
			SelectedTD.colSpan=ColSpanTD+AddColspan;
			SelectedTR.deleteCell(SelectedTD.cellIndex+1);
		}	
	}
}
function SplitRows()
{
	if (!CursorInTableCell()) return;
	var AddRowsNoSpan=1;
	var NsLeftColSpan=0;
	for (var i=0;i<SelectedTD.cellIndex;i++) NsLeftColSpan+=SelectedTR.cells[i].colSpan;
	var AllRows=SelectedTable.rows;
	while (SelectedTD.rowSpan>1&&AddRowsNoSpan>0)
	{
		var NextRow=AllRows[SelectedTR.rowIndex+SelectedTD.rowSpan-1];
		SelectedTD.rowSpan-=1;
		var NcLeftColSpan=0;
		var Position=-1;
		for (var n=0;n<NextRow.cells.length;n++)
		{
			NcLeftColSpan+=NextRow.cells[n].getAttribute('colSpan');
			if (NcLeftColSpan>NsLeftColSpan)
			{
				Position=n;
				break;
			}
		}
		var NewTD=NextRow.insertCell(Position);
		NewTD.innerHTML="&nbsp;";
		AddRowsNoSpan-=1;
	}
	for (var n=0;n<AddRowsNoSpan;n++)
	{
		var numCols=0
		allCells=SelectedTR.cells
		for (var i=0;i<allCells.length;i++) numCols=numCols+allCells[i].getAttribute('colSpan')
		var newTR=SelectedTable.insertRow(SelectedTR.rowIndex+1)
		for (var j=0;j<SelectedTR.rowIndex;j++)
		{
			for (var k=0;k<AllRows[j].cells.length;k++)
			{
				if ((AllRows[j].cells[k].rowSpan>1)&&(AllRows[j].cells[k].rowSpan>=SelectedTR.rowIndex-AllRows[j].rowIndex+1)) AllRows[j].cells[k].rowSpan+=1;
			}
		}
		for (i=0;i<allCells.length;i++)
		{
			if (i!=SelectedTD.cellIndex) SelectedTR.cells[i].rowSpan+=1;
			else
			{
				NewTD=newTR.insertCell();
				NewTD.colSpan=SelectedTD.colSpan;
				NewTD.innerHTML="&nbsp;";
			}
		}
	}
	ShowTableBorders();
}
function SplitColumn()
{
	if (!CursorInTableCell()) return;
	var AddColsNoSpan=1;
	var NewCell=null;
	var NsLeftColSpan=0;
	var NsLeftRowSpanMoreOne=0;
	for (var i=0;i<SelectedTD.cellIndex;i++)
	{
		NsLeftColSpan+=SelectedTR.cells[i].colSpan;
		if (SelectedTR.cells[i].rowSpan>1) NsLeftRowSpanMoreOne+=1;
	}
	var AllRows=SelectedTable.rows;
	while (SelectedTD.colSpan>1&&AddColsNoSpan>0)
	{
		NewCell=SelectedTR.insertCell(SelectedTD.cellIndex+1);
		NewCell.innerHTML="&nbsp;"
		selectedTD.colSpan-=1;
		AddColsNoSpan-=1;
	}
	for (i=0;i<AllRows.length;i++)
	{
		var ncLeftColSpan=0;
		var position=-1;
		for (var n=0;n<AllRows[i].cells.length;n++)
		{
			ncLeftColSpan+=AllRows[i].cells[n].getAttribute('colSpan');
			if (ncLeftColSpan+NsLeftRowSpanMoreOne>NsLeftColSpan)
			{
				position=n;
				break;
			}
		}
		if (SelectedTR.rowIndex!=i)
		{
			if (position!=-1) AllRows[i].cells[position+NsLeftRowSpanMoreOne].colSpan+=AddColsNoSpan;
		}
		else
		{
			for (var n=0;n<AddColsNoSpan;n++)
			{
				NewCell=AllRows[i].insertCell(SelectedTD.cellIndex+1)
				NewCell.innerHTML="&nbsp;"
				NewCell.rowSpan=SelectedTD.rowSpan;
			}
		}
	}
	ShowTableBorders();
}
function InsertHref(Operation)
{
	KS_EditArea.focus();
	KS_EditArea.document.execCommand(Operation,true);
	KS_EditArea.focus();
}
function Pos()   
{
	var ObjReference=null;
	var RangeType=KS_EditArea.document.selection.type;
	if (RangeType!="Control")
	{
		alert('你选择的不是对象！');
		return;
	}
	var SelectedRange= KS_EditArea.document.selection.createRange();
	for (var i=0; i<SelectedRange.length; i++)
	{
		ObjReference = SelectedRange.item(i);
		if (ObjReference.style.position != 'absolute') 
		{
			ObjReference.style.position='absolute';
		}
		else
		{
			ObjReference.style.position='static';
		}
	}
	KS_EditArea.content = false;
}
function CursorInTableCell()
{
	if (KS_EditArea.document.selection.type!="Control")
	{
		var SelectedElement=KS_EditArea.document.selection.createRange().parentElement();
		while (SelectedElement.tagName.toUpperCase()!="TD"&&SelectedElement.tagName.toUpperCase()!="TH")
		{
			SelectedElement=SelectedElement.parentElement;
			if (SelectedElement==null) break;
		}
		if (SelectedElement)
		{
			SelectedTD=SelectedElement;
			SelectedTR=SelectedTD.parentElement;
			SelectedTBODY=SelectedTR.parentElement;
			SelectedTable=SelectedTBODY.parentElement;
			return true;
		}
	}
}
function PasteText()
{
	KS_EditArea.focus();
	var sText = HTMLEncode(clipboardData.getData("Text")) ;
	InsertHTMLStr(sText);
	KS_EditArea.focus();
}
function SearchStr(InstallDir)
{
  var Temp=window.showModalDialog(InstallDir+"KS_Editor/InsertSearch.asp", window, "dialogWidth:340px; dialogHeight:190px; help: no; scroll: no; status: no");
}
//Change EditArea Height
function ChangeEditAreaHeight(KS_EditAreaHeight)
{ 
  for (var i=0; i<parent.frames.length; i++)
  {
			if (parent.frames[i].document==self.document)
			  {
				var obj=parent.frames[i].frameElement;
				var height = parseInt(obj.offsetHeight)+KS_EditAreaHeight;
				if (height+KS_EditAreaHeight>=100)
				{
				obj.height = height;
				}
			  }
	}
}
//editor Btn Click Event Function End.
function DisabledAllBtn(Flag)
{
	var AllBtnArray=document.body.getElementsByTagName('IMG'),CurrObj=null;
	for (var i=0;i<AllBtnArray.length;i++)
	{
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='Btn') CurrObj.disabled=Flag;
	}
	AllBtnArray=document.body.getElementsByTagName('SELECT');
	for (var i=0;i<AllBtnArray.length;i++)
	{
		CurrObj=AllBtnArray[i];
		if (CurrObj.className=='ToolSelectStyle') CurrObj.disabled=Flag;
	}
}
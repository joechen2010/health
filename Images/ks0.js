function getNames(obj,name,tij)
	{	
		var p = document.getElementById(obj);
		var plist = p.getElementsByTagName(tij);
		var rlist = new Array();
		for(i=0;i<plist.length;i++)
		{
			if(plist[i].getAttribute("name") == name)
			{
				rlist[rlist.length] = plist[i];
			}
		}
		return rlist;
	}

	function fod(obj,tag,name,showCss1,showCss2,unShowCss1,unShowCss2)
	{
		var p = getNames(tag,"t","div");
		var p1 = getNames(name,"f","div"); // document.getElementById(name).getElementsByTagName("div");
		for(i=0;i<p1.length;i++)
		{
			if(obj==p[i])
			{
				p[i].className = showCss1;
				p1[i].className = showCss2;
			}
			else
			{
				p[i].className = unShowCss1;
				p1[i].className = unShowCss2;
			}
		}
	}
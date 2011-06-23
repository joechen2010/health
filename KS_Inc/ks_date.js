						var lunarInfo=new Array(
						0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,
						0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,
						0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,
						0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,
						0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,
						0x06ca0,0x0b550,0x15355,0x04da0,0x0a5d0,0x14573,0x052d0,0x0a9a8,0x0e950,0x06aa0,
						0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,
						0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b5a0,0x195a6,
						0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,
						0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06b58,0x055c0,0x0ab60,0x096d5,0x092e0,
						0x0c960,0x0d954,0x0d4a0,0x0da50,0x07552,0x056a0,0x0abb7,0x025d0,0x092d0,0x0cab5,
						0x0a950,0x0b4a0,0x0baa4,0x0ad50,0x055d9,0x04ba0,0x0a5b0,0x15176,0x052b0,0x0a930,
						0x07954,0x06aa0,0x0ad50,0x05b52,0x04b60,0x0a6e6,0x0a4e0,0x0d260,0x0ea65,0x0d530,
						0x05aa0,0x076a3,0x096d0,0x04bd7,0x04ad0,0x0a4d0,0x1d0b6,0x0d250,0x0d520,0x0dd45,
						0x0b5a0,0x056d0,0x055b2,0x049b0,0x0a577,0x0a4b0,0x0aa50,0x1b255,0x06d20,0x0ada0,
						0x14b63);
						var solarMonth=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
						var Gan=new Array("甲","乙","丙","丁","戊","己","庚","辛","壬","癸");
						var Zhi=new Array("子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥");
						var nStr1 = new Array('日','一','二','三','四','五','六','七','八','九','十');
						var nStr2 = new Array('初','十','廿','卅','　');
						var nStr3 = new Array('日','一','二','三','四','五','六','七','八','九','十','十一','十二');
						function lYearDays(y) {
						   var i, sum = 348;
						   for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0;
						   return(sum+leapDays(y));
						}
						function leapDays(y) {
						   if(leapMonth(y))  return((lunarInfo[y-1900] & 0x10000)? 30: 29);
						   else return(0);
						}
						function leapMonth(y) {
						   return(lunarInfo[y-1900] & 0xf);
						}
						function monthDays(y,m) {
						   return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 );
						}
						function Lunar(objDate) {
						   var i, leap=0, temp=0;
						   var offset   = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate()) - Date.UTC(1900,0,31))/86400000;
						
						   this.dayCyl = offset + 40;
						   this.monCyl = 14;
						
						   for(i=1900; i<2050 && offset>0; i++) {
							  temp = lYearDays(i);
							  offset -= temp;
							  this.monCyl += 12;
						   }
						
						   if(offset<0) {
							  offset += temp;
							  i--;
							  this.monCyl -= 12;
						   }
						
						   this.year = i;
						   this.yearCyl = i-1864;
						
						   leap = leapMonth(i); //闰哪个月
						   this.isLeap = false;
						
						   for(i=1; i<13 && offset>0; i++) {
							  //闰月
							  if(leap>0 && i==(leap+1) && this.isLeap==false)
								 { --i; this.isLeap = true; temp = leapDays(this.year); }
							  else
								 { temp = monthDays(this.year, i); }
						
							  if(this.isLeap==true && i==(leap+1)) this.isLeap = false;
						
							  offset -= temp;
							  if(this.isLeap == false) this.monCyl ++;
						   }
						
						   if(offset==0 && leap>0 && i==leap+1)
							  if(this.isLeap)
								 { this.isLeap = false; }
							  else
								 { this.isLeap = true; --i; --this.monCyl;}
						
						   if(offset<0){ offset += temp; --i; --this.monCyl; }
						
						   this.month = i;
						   this.day = offset + 1;
						}
						function solarDays(y,m) {
						   if(m==1)
							  return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
						   else
							  return(solarMonth[m]);
						}
						function cyclical(num) {
						   return(Gan[num%10]+Zhi[num%12]);
						}
						function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,cDay) {
						
							  this.isToday    = false;
							  //公历
							  this.sYear      = sYear;
							  this.sMonth     = sMonth;
							  this.sDay       = sDay;
							  this.week       = week;
							  //农历
							  this.lYear      = lYear;
							  this.lMonth     = lMonth;
							  this.lDay       = lDay;
							  this.isLeap     = isLeap;
							  //干支
							  this.cYear      = cYear;
							  this.cMonth     = cMonth;
							  this.cDay       = cDay;
						
						}
						function calendar(y,m) {
						   var sDObj, lDObj, lY, lM, lD=1, lL, lX=0, tmp1, tmp2;
						   var lDPOS = new Array(3);
						   var n = 0;
						   var firstLM = 0;
						   sDObj = new Date(y,m,1,0,0,0,0);    //当月一日日期
						
						   this.length    = solarDays(y,m);    //公历当月天数
						   this.firstWeek = sDObj.getDay();    //公历当月1日星期几
						
						   for(var i=0;i<this.length;i++) {
						
							  if(lD>lX) {
								 sDObj = new Date(y,m,i+1);    //当月一日日期
								 lDObj = new Lunar(sDObj);     //农历
								 lY    = lDObj.year;           //农历年
								 lM    = lDObj.month;          //农历月
								 lD    = lDObj.day;            //农历日
								 lL    = lDObj.isLeap;         //农历是否闰月
								 lX    = lL? leapDays(lY): monthDays(lY,lM); //农历当月最後一天
						
								 if(n==0) firstLM = lM;
								 lDPOS[n++] = i-lD+1;
							  }
						
							  this[i] = new calElement(y, m+1, i+1, nStr1[(i+this.firstWeek)%7],
													   lY, lM, lD++, lL,
													   cyclical(lDObj.yearCyl) ,cyclical(lDObj.monCyl), cyclical(lDObj.dayCyl++) );
						   }
						
						}
						function cDay(d){
						   var s;
						
						   switch (d) {
							  case 10:
								 s = '初十'; break;
							  case 20:
								 s = '二十'; break;
								 break;
							  case 30:
								 s = '三十'; break;
								 break;
							  default :
								 s = nStr2[Math.floor(d/10)];
								 s += nStr1[d%10];
						   }
						   return(s);
						}
						function cMon(d){
						   var s;
						   switch (d) {
							  case 1:
								 s = '正'; break;
							  case 12:
								 s = '腊'; break;
							  default :
								 s = nStr3[d];
								 break;
						   }
						   return(s);
						}
						
						var ctime=new Date();
						var nYear = ctime.getYear();
						var nMonth = ctime.getMonth();
						var d = ctime.getDate() - 1;
						var ddd =new calendar(nYear,nMonth);
						
						var bsWeek=ddd[d].week;
						var bsDate=ddd[d].sYear+"年"+ddd[d].sMonth+"月";
						var bsDate2=ddd[d].sDay+"日";
						var bsYear=ddd[d].cYear + "年" + cMon(ddd[d].lMonth) + "月" + cDay(ddd[d].lDay);
						
						function CAL(){
						document.write("<table width='120' height='60' cellpadding='1'");
						document.write("<tr><td align='center'><b><font color=#008040>"+bsDate+bsDate2+"</font><br><font color=#FF7530><span style='FONT-SIZE: 10pt'>");
						document.write("星期" + bsWeek+"</span>"+"<br></b><font color=#9B4E00>");
						document.write(bsYear + "</td></tr></table>");
						}
						CAL();
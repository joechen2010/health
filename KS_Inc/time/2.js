tmpDate = new Date();
date = tmpDate.getDate();
month= tmpDate.getMonth() + 1 ;
year= tmpDate.getFullYear();
document.write(year);
document.write("��");
document.write(month);
document.write("��");
document.write(date);
document.write("�� ");
myArray=new Array(6);
myArray[0]="������"
myArray[1]="����һ"
myArray[2]="���ڶ�"
myArray[3]="������"
myArray[4]="������"
myArray[5]="������"
myArray[6]="������"
weekday=tmpDate.getDay();
if (weekday==0 | weekday==6)
{
document.write(myArray[weekday])
}
else
{document.write(myArray[weekday])
};
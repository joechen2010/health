var FlashSrc; //���嶯����ַ
var total;//����flashӰƬ��֡��
var frame_number;//����flashӰƬ��ǰ֡��

//�����ǹ�����ͼƬ�϶�����
var dragapproved=false;
var z,x,y
//�ƶ�����
function move(){
if (event.button==1&&dragapproved){
y=temp1+event.clientX-x;
//�����ǿ����ƶ��ķ�Χ
if(y<0)
 y=0;
if(y>500)
 y=500;

z.style.pixelLeft=y
movie.GotoFrame(y/500*total);//�ƶ���ĳһλ�ã�flashӰƬ���ŵ�ĳ��λ��
return false
}
}
//����϶�ǰ��ʼ���ݵĺ���
function drags(){
if (!document.all)
return
if (event.srcElement.className=="drag"){
dragapproved=true
z=event.srcElement
temp1=z.style.pixelLeft
x=event.clientX
document.onmousemove=move
}
}

//��̬��ʾ����ӰƬ�ĵ�ǰ֡/��֡��
function ShowCount(){
 frame_number=movie.CurrentFrame();
 frame_number++;
 frameCount.innerText=frame_number+"/"+movie.TotalFrames;
 element.style.pixelLeft=480*(frame_number/movie.TotalFrames)-15;//������ͼƬ��֮����Ӧ��λ��
 if(frame_number==movie.TotalFrames)
  clearTimeout(tn_ID);
 else
  var tn_ID=setTimeout('ShowCount();',1000);
}
//ʹӰƬ���ص�һ֡ 
function Rewind(){
 if(movie.IsPlaying()){
 Pause();
 }
 movie.Rewind();
 element.style.pixelLeft=0;
 frameCount.innerText="1/"+total;
}
//����ӰƬ 
function Play(){
 movie.Play();
 ShowCount();
}
//��ͣ����
function Pause(){
 movie.StopPlay();
}

//������ĩ֡
function GoToEnd(){
 if(movie.IsPlaying())
  Pause();
 movie.GotoFrame(total);
 element.style.pixelLeft=500;
 frameCount.innerText=total+"/"+total;
}
//����ӰƬ
function Back()
{
 if(movie.IsPlaying())
  Pause();
 frame_number=frame_number-50;
 movie.GotoFrame(frame_number);
 Play();
}
//���ӰƬ
function Forward()
{
 if(movie.IsPlaying())
  Pause();
 frame_number=frame_number+50;
 movie.GotoFrame(frame_number);
 Play();
}
//���²���ӰƬ
function Replay(){
 if(movie.IsPlaying()){
 Pause();
 movie.Rewind();
 Play();
 }
 else
 {
 movie.Rewind();
 Play(); 
 }
}
//ֹͣ����ӰƬ���ص���һ֡
function Stop(){
 if(movie.IsPlaying()){
 Pause();
 movie.Rewind();
 }
 else
 {
 movie.Rewind();
 }
}
//ȫ���ۿ�
function FullScreen()
{
 window.open(FlashSrc);	
}
//��ʾӰƬ������ȣ���ȫ�������ư�ť����
function Loading(){
	
 var in_ID;
 bar.style.width=Math.round(movie.PercentLoaded())+"%";
 frameCount.innerText=Math.round(movie.PercentLoaded())+"%";
 if(movie.PercentLoaded() >= 100){
  PlayerButtons.document.all.tags('IMG')[0].disabled=false;
  PlayerButtons.document.all.tags('IMG')[1].disabled=false;
  PlayerButtons.document.all.tags('IMG')[2].disabled=false;
  PlayerButtons.document.all.tags('IMG')[3].disabled=false;
  PlayerButtons.document.all.tags('IMG')[4].disabled=false;
  PlayerButtons.document.all.tags('IMG')[5].disabled=false;
  PlayerButtons.document.all.tags('IMG')[6].disabled=false;
  PlayerButtons.document.all.tags('IMG')[7].disabled=false;
  PlayerButtons.document.all.tags('IMG')[8].disabled=false;

total=movie.TotalFrames;
  frame_number++;
  frameCount.innerText=frame_number+"/"+total;
  bar.style.background="";
  bar.innerHTML='<img src="/Images/Default/posbar1.gif" style="POSITION:relative;cursor:pointer;border:0;" id="element" class="drag" OnMouseOver="fnOnMouseOver()" OnMouseOut="fnOnMouseOut()">';
  document.onmousedown=drags
  document.onmouseup=new Function("dragapproved=false;Play()")
  ShowCount();
  clearTimeout(in_ID);
 }
 else
  in_ID=setTimeout("Loading();",1000);
}

//��ʼ����flashӰƬ����������У����ſ��ư�ť������
function LoadFlashUrl(FlashUrl,FlashWidth,FlashHeight){
 FlashSrc=FlashUrl;
 movie.LoadMovie(0, FlashUrl);
 movie.width=FlashWidth;
 movie.height=FlashHeight;
 PlayerButtons.document.all.tags('IMG')[0].disabled=true;
 PlayerButtons.document.all.tags('IMG')[1].disabled=true;
 PlayerButtons.document.all.tags('IMG')[2].disabled=true;
 PlayerButtons.document.all.tags('IMG')[3].disabled=true;
 PlayerButtons.document.all.tags('IMG')[4].disabled=true;
 PlayerButtons.document.all.tags('IMG')[5].disabled=true;
 PlayerButtons.document.all.tags('IMG')[6].disabled=true;
 PlayerButtons.document.all.tags('IMG')[7].disabled=true;
 PlayerButtons.document.all.tags('IMG')[8].disabled=true;

 frame_number=movie.CurrentFrame();
 Loading();
}
//��ʾ�㺯��
function showMenu(menu){
menu.style.display='block';
}

//������������ϵ�λ�ã�ӰƬ��Ӧ���ŵ��Ǹ�λ��
function Jump(fnume){
 if(movie.IsPlaying()){
 Pause();
 movie.GotoFrame(fnume);
 Play();
 }
 else
 {
 movie.GotoFrame(fnume);
 Play();
 }
}

//��������������ͼƬ�л�����
function fnOnMouseOver(){
 element.src = "/Images/Default/posbar.gif";
}

function fnOnMouseOut(){
 element.src = "/Images/Default/posbar1.gif";
}


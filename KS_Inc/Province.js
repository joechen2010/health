
cityareaname=new Array(35);
cityareacode=new Array(35);
 function first(preP,preC,formname,prov,city)
   {
     a=0;
if (prov=='01')
  { a=1;tempoption=new Option('����','01',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[1]=tempoption;');
cityareacode[0]=new Array('������','������','������','������','������','������','��̨��','ʯ��ɽ');
cityareaname[0]=new Array('������','������','������','������','������','������','��̨��','ʯ��ɽ');
if (prov=='02')
  { a=2;tempoption=new Option('����','02',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[2]=tempoption;');
cityareacode[1]=new Array('�޺���','������','��ɽ��','������','������','������');
cityareaname[1]=new Array('�޺���','������','��ɽ��','������','������','������');
if (prov=='03')
  { a=3;tempoption=new Option('�Ϻ�','03',false,true); }
else
  { tempoption=new Option('�Ϻ�','�Ϻ�'); }
eval('document.all.'+preP+'.options[3]=tempoption;');
cityareacode[2]=new Array('��ɽ','��ɽ','����','����','����','����','����','¬��','�ɽ�','����','�ֶ�','����','���','����','բ��','����','����','���','�ζ�','�ϻ�');
cityareaname[2]=new Array('��ɽ','��ɽ','����','����','����','����','����','¬��','�ɽ�','����','�ֶ�','����','���','����','բ��','����','����','���','�ζ�','�ϻ�');
if (prov=='04')
  { a=4;tempoption=new Option('����','04',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[4]=tempoption;');
cityareacode[3]=new Array('����','����','ɳƺ��','�ϰ�','������','��ɿ�');
cityareaname[3]=new Array('����','����','ɳƺ��','�ϰ�','������','��ɿ�');
if (prov=='05')
  { a=5;tempoption=new Option('���','05',false,true); }
else
  { tempoption=new Option('���','���'); }
eval('document.all.'+preP+'.options[5]=tempoption;');
cityareacode[4]=new Array('��ƽ','�ӱ�','����','�Ӷ�','�Ͽ�','����','����','����','���','����','����','����','����','����','����');
cityareaname[4]=new Array('��ƽ','�ӱ�','����','�Ӷ�','�Ͽ�','����','����','����','���','����','����','����','����','����','����');
if (prov=='06')
  { a=6;tempoption=new Option('�㶫','06',false,true); }
else
  { tempoption=new Option('�㶫','�㶫'); }
eval('document.all.'+preP+'.options[6]=tempoption;');
cityareacode[5]=new Array('����','�麣','��ɽ','��ɽ','��ݸ','��Զ','����','����','տ��','ï��','�ع�','����','��Դ','��β','��ͷ','÷��');
cityareaname[5]=new Array('����','�麣','��ɽ','��ɽ','��ݸ','��Զ','����','����','տ��','ï��','�ع�','����','��Դ','��β','��ͷ','÷��');
if (prov=='07')
  { a=7;tempoption=new Option('�ӱ�','07',false,true); }
else
  { tempoption=new Option('�ӱ�','�ӱ�'); }
eval('document.all.'+preP+'.options[7]=tempoption;');
cityareacode[6]=new Array('ʯ��ׯ','��ɽ','�ػʵ�','����','��̨','�żҿ�','�е�','�ȷ�','����','����','��ˮ');
cityareaname[6]=new Array('ʯ��ׯ','��ɽ','�ػʵ�','����','��̨','�żҿ�','�е�','�ȷ�','����','����','��ˮ');
if (prov=='08')
  { a=8;tempoption=new Option('ɽ��','08',false,true); }
else
  { tempoption=new Option('ɽ��','ɽ��'); }
eval('document.all.'+preP+'.options[8]=tempoption;');
cityareacode[7]=new Array('̫ԭ','��ͬ','��Ȫ','˷��','����','�ٷ�','����');
cityareaname[7]=new Array('̫ԭ','��ͬ','��Ȫ','˷��','����','�ٷ�','����');
if (prov=='09')
  { a=9;tempoption=new Option('���ɹ�','09',false,true); }
else
  { tempoption=new Option('���ɹ�','���ɹ�'); }
eval('document.all.'+preP+'.options[9]=tempoption;');
cityareacode[8]=new Array('���ͺ���','��ͷ','�ں�','�ٺ�','��ʤ','����','���ֺ���','ͨ��','���','������','��������');
cityareaname[8]=new Array('���ͺ���','��ͷ','�ں�','�ٺ�','��ʤ','����','���ֺ���','ͨ��','���','������','��������');
if (prov=='10')
  { a=10;tempoption=new Option('����','10',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[10]=tempoption;');
cityareacode[9]=new Array('����','����','��ɽ','����','����','�̽�','����','��˳','Ӫ��','����','����','��Ϫ','����','��«��');
cityareaname[9]=new Array('����','����','��ɽ','����','����','�̽�','����','��˳','Ӫ��','����','����','��Ϫ','����','��«��');
if (prov=='11')
  { a=11;tempoption=new Option('����','11',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[11]=tempoption;');
cityareacode[10]=new Array('����','����','��ƽ','��Դ','ͨ��','��ɽ','��ԭ','�׳�','�ӱ�');
cityareaname[10]=new Array('����','����','��ƽ','��Դ','ͨ��','��ɽ','��ԭ','�׳�','�ӱ�');
if (prov=='12')
  { a=12;tempoption=new Option('������','12',false,true); }
else
  { tempoption=new Option('������','������'); }
eval('document.all.'+preP+'.options[12]=tempoption;');
cityareacode[11]=new Array('������','�������','ĵ����','��ľ˹','����','����','�ں�','����','�׸�','˫Ѽɽ','��̨��','�绯','���˰���');
cityareaname[11]=new Array('������','�������','ĵ����','��ľ˹','����','����','�ں�','����','�׸�','˫Ѽɽ','��̨��','�绯','���˰���');
if (prov=='13')
  { a=13;tempoption=new Option('����','13',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[13]=tempoption;');
cityareacode[12]=new Array('�Ͼ�','����','����','����','��','���Ƹ� ','����','���� ','��ͨ','�γ�','����','̩��','��Ǩ');
cityareaname[12]=new Array('�Ͼ�','����','����','����','��','���Ƹ� ','����','���� ','��ͨ','�γ�','����','̩��','��Ǩ');
if (prov=='14')
  { a=14;tempoption=new Option('�㽭','14',false,true); }
else
  { tempoption=new Option('�㽭','�㽭'); }
eval('document.all.'+preP+'.options[14]=tempoption;');
cityareacode[13]=new Array('����','����','��ˮ','����','����','��ɽ','����','��','̨��','����','����');
cityareaname[13]=new Array('����','����','��ˮ','����','����','��ɽ','����','��','̨��','����','����');
if (prov=='15')
  { a=15;tempoption=new Option('����','15',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[15]=tempoption;');
cityareacode[14]=new Array('�Ϸ�  ','�ߺ� ','���� ','���� ','���� ','���� ','��ɽ ','���� ','���� ','���� ','��ɽ ','ͭ��','���� ','���� ','���� ','���� ','����');
cityareaname[14]=new Array('�Ϸ�  ','�ߺ� ','���� ','���� ','���� ','���� ','��ɽ ','���� ','���� ','���� ','��ɽ ','ͭ��','���� ','���� ','���� ','���� ','����');
if (prov=='16')
  { a=16;tempoption=new Option('����','16',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[16]=tempoption;');
cityareacode[15]=new Array('���� ','���� ','Ȫ�� ','���� ','���� ','��ƽ ','���� ','���� ','����');
cityareaname[15]=new Array('���� ','���� ','Ȫ�� ','���� ','���� ','��ƽ ','���� ','���� ','����');
if (prov=='17')
  { a=17;tempoption=new Option('����','17',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[17]=tempoption;');
cityareacode[16]=new Array('�ϲ�','������','�Ž�','Ƽ��','����','ӥ̶','����','�˴�','����','����','����');
cityareaname[16]=new Array('�ϲ�','������','�Ž�','Ƽ��','����','ӥ̶','����','�˴�','����','����','����');
if (prov=='18')
  { a=18;tempoption=new Option('ɽ��','18',false,true); }
else
  { tempoption=new Option('ɽ��','ɽ��'); }
eval('document.all.'+preP+'.options[18]=tempoption;');
cityareacode[17]=new Array('����','�ൺ','�Ͳ�','����','��̨','Ϋ��','����','̩��','����','����','����','��ׯ','����','����','�ĳ�','����','��Ӫ');
cityareaname[17]=new Array('����','�ൺ','�Ͳ�','����','��̨','Ϋ��','����','̩��','����','����','����','��ׯ','����','����','�ĳ�','����','��Ӫ');
if (prov=='19')
  { a=19;tempoption=new Option('����','19',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[19]=tempoption;');
cityareacode[18]=new Array('֣��','����','����','ƽ��ɽ','����','�ױ�','����','����','���','���','���','����Ͽ','����','����','�ܿ�','פ���','����','��Դ');
cityareaname[18]=new Array('֣��','����','����','ƽ��ɽ','����','�ױ�','����','����','���','���','���','����Ͽ','����','����','�ܿ�','פ���','����','��Դ');
if (prov=='20')
  { a=20;tempoption=new Option('����','20',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[20]=tempoption;');
cityareacode[19]=new Array('�人','��ʯ','ʮ��','����','�˲�','�差','����','����','Т��','�Ƹ�','����','��ʩ','����','����','����','Ǳ��','��ũ��');
cityareaname[19]=new Array('�人','��ʯ','ʮ��','����','�˲�','�差','����','����','Т��','�Ƹ�','����','��ʩ','����','����','����','Ǳ��','��ũ��');
if (prov=='21')
  { a=21;tempoption=new Option('����','21',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[21]=tempoption;');
cityareacode[20]=new Array('��ɳ','����','��̶','����','����','����','����','����','����','����','����','¦��','���� ');
cityareaname[20]=new Array('��ɳ','����','��̶','����','����','����','����','����','����','����','����','¦��','���� ');
if (prov=='22')
  { a=22;tempoption=new Option('����','22',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[22]=tempoption;');
cityareacode[21]=new Array('����','����','����','����','����','���Ǹ�','����','���','����','����','��ɫ','�ӳ�');
cityareaname[21]=new Array('����','����','����','����','����','���Ǹ�','����','���','����','����','��ɫ','�ӳ�');
if (prov=='23')
  { a=23;tempoption=new Option('����','23',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[23]=tempoption;');
cityareacode[22]=new Array('���� ','����','ͨʲ','��','��ɽ','�Ĳ�','����','����','����');
cityareaname[22]=new Array('���� ','����','ͨʲ','��','��ɽ','�Ĳ�','����','����','����');
if (prov=='24')
  { a=24;tempoption=new Option('�Ĵ�','24',false,true); }
else
  { tempoption=new Option('�Ĵ�','�Ĵ�'); }
eval('document.all.'+preP+'.options[24]=tempoption;');
cityareacode[23]=new Array('�ɶ�','�Թ�','��֦��','����','����','����','��Ԫ','����','�ڽ�','��ɽ','�ϳ�  ','�˱�','�㰲','�ﴨ','����','�Ű�','üɽ  ','���� ','���� ','��ɽ');
cityareaname[23]=new Array('�ɶ�','�Թ�','��֦��','����','����','����','��Ԫ','����','�ڽ�','��ɽ','�ϳ�  ','�˱�','�㰲','�ﴨ','����','�Ű�','üɽ  ','���� ','���� ','��ɽ');
if (prov=='25')
  { a=25;tempoption=new Option('����','25',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[25]=tempoption;');
cityareacode[24]=new Array('���� ','����ˮ','����','ͭ��','�Ͻ�','��˳','ǭ���� ','ǭ����','ǭ��');
cityareaname[24]=new Array('���� ','����ˮ','����','ͭ��','�Ͻ�','��˳','ǭ���� ','ǭ����','ǭ��');
if (prov=='26')
  { a=26;tempoption=new Option('����','26',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[26]=tempoption;');
cityareacode[25]=new Array('����','����','����','��Ϫ','��ͨ','˼é','�ٲ�','��ɽ','����','��ɽ ','��� ','��˫���� ','���� ','���� ','�º� ','ŭ��','����');
cityareaname[25]=new Array('����','����','����','��Ϫ','��ͨ','˼é','�ٲ�','��ɽ','����','��ɽ ','��� ','��˫���� ','���� ','���� ','�º� ','ŭ��','����');
if (prov=='27')
  { a=27;tempoption=new Option('����','27',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[27]=tempoption;');
cityareacode[26]=new Array('����','����','����','ɽ��','�տ���','����','��֥');
cityareaname[26]=new Array('����','����','����','ɽ��','�տ���','����','��֥');
if (prov=='28')
  { a=28;tempoption=new Option('����','28',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[28]=tempoption;');
cityareacode[27]=new Array('����','ͭ��','����','����','μ��','�Ӱ�','����','����','����','����');
cityareaname[27]=new Array('����','ͭ��','����','����','μ��','�Ӱ�','����','����','����','����');
if (prov=='29')
  { a=29;tempoption=new Option('����','29',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[29]=tempoption;');
cityareacode[28]=new Array('����','���','����','��ˮ','������','����','ƽ��','����','¤��','����','��Ҵ','��Ȫ','���� ','����');
cityareaname[28]=new Array('����','���','����','��ˮ','������','����','ƽ��','����','¤��','����','��Ҵ','��Ȫ','���� ','����');
if (prov=='30')
  { a=30;tempoption=new Option('�ຣ','30',false,true); }
else
  { tempoption=new Option('�ຣ','�ຣ'); }
eval('document.all.'+preP+'.options[30]=tempoption;');
cityareacode[29]=new Array('����','����',' ���� ','����','����','����','����','����');
cityareaname[29]=new Array('����','����',' ���� ','����','����','����','����','����');
if (prov=='31')
  { a=31;tempoption=new Option('����','31',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[31]=tempoption;');
cityareacode[30]=new Array('����','ʯ��ɽ','����','��ԭ');
cityareaname[30]=new Array('����','ʯ��ɽ','����','��ԭ');
if (prov=='32')
  { a=32;tempoption=new Option('�½�','32',false,true); }
else
  { tempoption=new Option('�½�','�½�'); }
eval('document.all.'+preP+'.options[32]=tempoption;');
cityareacode[31]=new Array('��³ľ��','��������','ʯ����','��³��','����','����','������','��ʲ','��������','��������','����','��������','����');
cityareaname[31]=new Array('��³ľ��','��������','ʯ����','��³��','����','����','������','��ʲ','��������','��������','����','��������','����');
if (prov=='33')
  { a=33;tempoption=new Option('���','33',false,true); }
else
  { tempoption=new Option('���','���'); }
eval('document.all.'+preP+'.options[33]=tempoption;');
cityareacode[32]=new Array();
cityareaname[32]=new Array();
if (prov=='34')
  { a=34;tempoption=new Option('����','34',false,true); }
else
  { tempoption=new Option('����','����'); }
eval('document.all.'+preP+'.options[34]=tempoption;');
cityareacode[33]=new Array();
cityareaname[33]=new Array();
if (prov=='35')
  { a=35;tempoption=new Option('̨��','35',false,true); }
else
  { tempoption=new Option('̨��','̨��'); }
eval('document.all.'+preP+'.options[35]=tempoption;');
cityareacode[34]=new Array();
cityareaname[34]=new Array();

eval('document.all.'+preP+'.options[a].selected=true;');

cityid=prov;
    if (cityid!='0')
      {
        b=0;for (i=0;i<cityareaname[cityid-1].length;i++)
           {
             if (city==cityareacode[cityid-1][i])
               {b=i+1;tempoption=new Option(cityareaname[cityid-1][i],cityareacode[cityid-1][i],false,true);}
             else
               tempoption=new Option(cityareaname[cityid-1][i],cityareacode[cityid-1][i]);
            eval('document.all.'+preC+'.options[i+1]=tempoption;');
           }
        eval('document.all.'+preC+'.options[b].selected=true;');
      }
    }
 function selectcityarea(preP,preC)
   {
     cityid=eval('document.all.'+preP+'.selectedIndex;');
     j=eval('document.all.'+preC+'.length;');
     for (i=1;i<j;i++)
        {eval('document.all.'+preC+'.options[j-i]=null;')}
     if (cityid!="0")
       {
         for (i=0;i<cityareaname[cityid-1].length;i++)
            {
             tempoption=new Option(cityareaname[cityid-1][i],cityareacode[cityid-1][i]);
             eval('document.all.'+preC+'.options[i+1]=tempoption;');
            }
       }
    }

first("prov","city",0,0);
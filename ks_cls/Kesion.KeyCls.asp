<%
'�ʿ��ļ�����config/zwck.txt,�ôʿ�Ŀǰ���213663����Ŀ,��Ҳ����������
'˵��:�˰汾����Ӣ�Ĵ���.
Class Wordsegment_Cls
	Private KeyList
	Private CacheCK
	Private Sub Class_Initialize() 
	  CacheCK=true                 '�Ƿ񻺴�ʿ�,true �� false��.���ڴ��Сʱ��������Ϊfalse,�����ܻ�Ӱ��ؼ��ʵļ����ٶ�
	End Sub
	
	Private Sub Class_Terminate()
	End Sub

    '�ʿ�����
	Sub LoadKey()
		Dim KS:Set KS=new publiccls
		If Application(KS.SiteSN & "ZWCK")="" Or CacheCK=False Then
		 KeyList=KS.ReadFromFile(KS.Setting(3) & "config/zwck.txt")
		 KeyList="|" & Replace(KeyList,vbcrlf,"|") & "|"
		 Application(KS.SiteSN & "ZWCK")=KeyList
		Else
		 KeyList=Application(KS.SiteSN & "ZWCK")
		End If
	End Sub
	
	'����:str ���������ַ��� x ȡ���� maxlen ��෵�س���
	Function SplitKey(str, x,maxlen)
	    LoadKey
		Dim a, b 'As String
		Dim i, j, flag, max, temp_str
	
		a = str
	
		'�ִ�
		For i = 1 To Len(a)
			For j = 1 To x
				a = a & Mid(a, i, j) & " "
			Next
		Next
	
		a = Split(a, " ")
		max = UBound(a)
	
		'�����ظ��ַ���
		For i = 0 To max - 1
			flag = a(i)
			If iscn(flag) Then
				For j = i + 1 To max - 1
					If a(j) = flag And flag <> "" Then
						a(j) = ""   
					End If
				Next
				If a(i) <> "" Then
					temp_str = temp_str & a(i) & " "
				End If
		   End If
		Next
		 
		 a = Split(temp_str, " ")
		 temp_str = ""
		 For i=0 to Ubound(a)
		  If Instr(KeyList,"|" & a(i) & "|")<>0 Then
		   temp_str =temp_str  & a(i) &  " "
		   If len(temp_str)>maxlen and maxlen<>0 then exit for
		  end if
		 Next
	   
		SplitKey=replace(trim(temp_str)," ",",")
	End Function
	
	'�ж�����
	function iscn(str) 
		Dim i 
		i = Len(str) 
		If i = 0 Then 
		   iscn = False 
		   Exit Function 
		End If 
		
		Do While i > 0 
		  If Asc(Mid(str, i, 1)) < 10000 And Asc(Mid(str, i, 1)) > -10000 Then 
		    iscn = False 
		    Exit Function 
		  End If 
		  i = i - 1 
		Loop 
		iscn = True 
	end function 

End Class

%>
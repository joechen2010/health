<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'����:��Ҫ���ڽ��JQuery���ı���

 '��javascript�е�escape()��Ч
 Function Escape(str)
        dim i,s,c,a 
        s="" 
        For i=1 to Len(str) 
            c=Mid(str,i,1)
            a=ASCW(c)
            If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
                s = s & c
            ElseIf InStr("@*_+-./",c)>0 Then
                s = s & c
            ElseIf a>0 and a<16 Then
                s = s & "%0" & Hex(a)
            ElseIf a>=16 and a<256 Then
                s = s & "%" & Hex(a)
            Else
                s = s & "%u" & Hex(a)
            End If
        Next
        Escape=s
    End Function
    '��javascript�е�unescape()��Ч
    Function UnEscape(str)
                    Dim x
        x=InStr(str,"%") 
        Do While x>0
            UnEscape=UnEscape&Mid(str,1,x-1)
            If LCase(Mid(str,x+1,1))="u" Then
                UnEscape=UnEscape&ChrW(CLng("&H"&Mid(str,x+2,4)))
                str=Mid(str,x+6)
            Else
                UnEscape=UnEscape&Chr(CLng("&H"&Mid(str,x+1,2)))
                str=Mid(str,x+3)
            End If
            x=InStr(str,"%")
        Loop
        UnEscape=UnEscape&str
    End Function
%>

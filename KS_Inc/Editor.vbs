Function Resumeblank(ByVal Content)
					if Content="" then 
						Resumeblank=Content 
						Exit Function
					end if
					Dim strHtml, strHtml2, Num, Numtemp, Strtemp, i
					strHtml = Replace(strHtml, "<SPAN", "<span") 
					strHtml = Replace(strHtml, "</SPAN>", "</span>")
					strHtml = Replace(Content, "<DIV", "<div")
					strHtml = Replace(strHtml, "</DIV>", "</div>")
					strHtml = Replace(strHtml, "<OBJECT", "<object")
					strHtml = Replace(strHtml, "</OBJECT>", "</object>")
					strHtml = Replace(strHtml, "<PARAM", "<param")
					strHtml = Replace(strHtml, "</PARAM>", "</param>")
					strHtml = Replace(strHtml, "<P", "<p")
					strHtml = Replace(strHtml, "</P>", "</p>")
					strHtml = Replace(strHtml, "<BR", "<br /")
					strHtml = Replace(strHtml, "<HR", "<hr")
					strHtml = Replace(strHtml, "<STRONG", "<strong")
					strHtml = Replace(strHtml, "</STRONG>", "</strong>")
					strHtml = Replace(strHtml, "<EM", "<em")
					strHtml = Replace(strHtml, "</EM>", "</em>")
					strHtml = Replace(strHtml, "<UL", "<ul")
					strHtml = Replace(strHtml, "</UL>", "</ul>")
					strHtml = Replace(strHtml, "<OL", "<ol")
					strHtml = Replace(strHtml, "</OL>", "</ol>")
					strHtml = Replace(strHtml, "<LI", "<li")
					strHtml = Replace(strHtml, "</LI>", "</li>")
					strHtml = Replace(strHtml, "<U", "<u")
					strHtml = Replace(strHtml, "</U>", "</u>")
					strHtml = Replace(strHtml, "<A", "<a")
					strHtml = Replace(strHtml, "</A>", "</a>")
					strHtml = Replace(strHtml, "<IMG", "<img")
					strHtml = Replace(strHtml,"VALUE=","value=")
					strHtml = Replace(strHtml, "<FONT", "<font")
					strHtml = Replace(strHtml, "</FONT>", "</font>")
					strHtml = Replace(strHtml, "<TABLE", "<table")
					strHtml = Replace(strHtml, "</TABLE>", vbCrLf & "</table>" & vbCrLf)
					strHtml = Replace(strHtml, "<TBODY>", "")
					strHtml = Replace(strHtml, "</TBODY>", "" & vbCrLf)
					strHtml = Replace(strHtml, "<TR", "<tr")
					strHtml = Replace(strHtml, "</TR>", vbCrLf & "</tr>" & vbCrLf)
					strHtml = Replace(strHtml, "<TD", "<td")
					strHtml = Replace(strHtml, "</TD>", "</td>")
					strHtml = Replace(strHtml, "<"&"!--", vbCrLf & "<"&"!--")
					strHtml = Replace(strHtml, "<SELECT", vbCrLf & "<Select")
					strHtml = Replace(strHtml, "</SELECT>", vbCrLf & "</Select>")
					strHtml = Replace(strHtml, "<OPTION", vbCrLf & "  <Option")
					strHtml = Replace(strHtml, "</OPTION>", "</Option>")
					strHtml = Replace(strHtml, "<INPUT", vbCrLf & "  <Input")
					strHtml = Replace(strHtml, "<" & "script", vbCrLf & "<"&"script")
					strHtml = Replace(strHtml, "&amp;", "&")
					strHtml = Replace(strHtml, "{$--", vbCrLf & "<"&"!--$")
					strHtml = Replace(strHtml, "--}", "$--"&">")
					arrContent = Split(strHtml, vbCrLf)
					For i = 0 To UBound(arrContent)
						Numtemp = False
						If InStr(arrContent(i), "<table") > 0 Then
							Numtemp = True
							If Strtemp <> "<table" And Strtemp <> "</table>" Then
								Num = Num + 2
							End If
							Strtemp = "<table"
						ElseIf InStr(arrContent(i), "<tr") > 0 Then
							Numtemp = True
							If Strtemp <> "<tr" And Strtemp <> "</tr>" Then
								Num = Num + 2
							End If
							Strtemp = "<tr"
						ElseIf InStr(arrContent(i), "<td") > 0 Then
							Numtemp = True
							If Strtemp <> "<td" And Strtemp <> "</td>" Then
								Num = Num + 2
							End If
							Strtemp = "<td"
						ElseIf InStr(arrContent(i), "</table>") > 0 Then
							Numtemp = True
							If Strtemp <> "</table>" And Strtemp <> "<table" Then
								Num = Num - 2
							End If
							Strtemp = "</table>"
						ElseIf InStr(arrContent(i), "</tr>") > 0 Then
							Numtemp = True
							If Strtemp <> "</tr>" And Strtemp <> "<tr" Then
								Num = Num - 2
							End If
							Strtemp = "</tr>"
						ElseIf InStr(arrContent(i), "</td>") > 0 Then
							Numtemp = True
							If Strtemp <> "</td>" And Strtemp <> "<td" Then
								Num = Num - 2
							End If
							Strtemp = "</td>"
						ElseIf InStr(arrContent(i), "<"&"!--") > 0 Then
							Numtemp = True
						End If
				
						If Num < 0 Then Num = 0
						If Trim(arrContent(i)) <> "" Then
							If i = 0 Then
								strHtml2 = String(Num, " ") & arrContent(i)
							ElseIf Numtemp = True Then
								strHtml2 = strHtml2 & vbCrLf & String(Num, " ") & arrContent(i)
							Else
								strHtml2 = strHtml2 & vbCrLf & arrContent(i)
							End If
						End If
					Next
			Resumeblank = strHtml2
	End Function
	function FormatHtml(Content)
	  FormatHtml=Content
	  Exit Function
	 Dim regEx, Matches, Match, TempStr
				Set regEx = New RegExp
				regEx.Pattern = "(background=|SIZE=|color=|bgColor=|colSpan=|align=|width=|height=|cellSpacing=|cellPadding=|border=|class=| id=|target=)[^( |)>]*"
				regEx.IgnoreCase = True
				regEx.Global = True
				Set Matches = regEx.Execute(Content)
				For Each Match In Matches
				 If Instr(Match.Value,"""")=0 Then
					 If Instr(Match.Value,"class=")<>0 then TempStr=Replace(Match.Value,"class=","class=""")&""""
					 If Instr(Match.Value,"target=")<>0 then TempStr=Replace(Match.Value,"target=","target=""")&""""
					 If Instr(Match.Value," id=")<>0 then TempStr=Replace(Match.Value," id="," id=""")&""""
					 If Instr(Match.Value,"border=")<>0 then TempStr=Replace(Match.Value,"border=","border=""")&""""
					 If Instr(Match.Value,"cellPadding=")<>0 then TempStr=Replace(Match.Value,"cellPadding=","cellpadding=""")&""""
					 If Instr(Match.Value,"cellSpacing=")<>0 then TempStr=Replace(Match.Value,"cellSpacing=","cellspacing=""")&""""
					 If Instr(Match.Value,"width=")<>0 and Instr(Match.Value,"%")=0 then  TempStr=Replace(Match.Value,"width=","width=""")&""""
					  If Instr(Match.Value,"height=")<>0 and Instr(Match.Value,"%")=0 then  TempStr=Replace(Match.Value,"height=","height=""")&""""
					 If Instr(Match.Value,"align=")<>0 then  TempStr=Replace(Match.Value,"align=","align=""")&""""
					 If Instr(Match.Value,"colSpan=")<>0 then  TempStr=Replace(Match.Value,"colSpan=","colSpan=""")&""""
					 If Instr(Match.Value,"bgColor=")<>0 then  TempStr=Replace(Match.Value,"bgColor=","bgColor=""")&""""
					 If Instr(Match.Value,"color=")<>0 then  TempStr=Replace(Match.Value,"color=","color=""")&""""
					 If Instr(Match.Value,"SIZE=")<>0 then  TempStr=Replace(Match.Value,"SIZE=","size=""")&""""
					 If Instr(Match.Value,"background=")<>0 then  TempStr=Replace(Match.Value,"background=","background=""")&""""
					 
					 'If Instr(Match.Value,"http-equiv=")<>0 then  TempStr=Replace(Match.Value,"http-equiv=","http-equiv=""")&""""
					' If Instr(Match.Value,"rel=")<>0 then  TempStr=Replace(Match.Value,"rel=","rel=""")&""""
					 'If Instr(Match.Value,"type=")<>0 then  TempStr=Replace(Match.Value,"type=","type=""")&""""
					
					 If Instr(Match.value,"%")=0 Then Content=Replace(Content,Match.value,TempStr)
				 End If
				Next
				FormatHtml=Content
	End Function
	
	
	function ReplaceScriptToImg(Content)
		   Dim regEx,Match, Matches,strTemp 
		    Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
		    regEx.Pattern = "(\<Script)(.*?)(\<\/Script\>)"
        Set Matches = regEx.Execute(Content)
        For Each Match In Matches
            strTemp = Replace(Match.Value, "<", "[!")
            strTemp = Replace(strTemp, ">", "!]")
            strTemp = Replace(strTemp, "'", "¡ä")
            strTemp = "<IMG alt='#" & strTemp & "#' src=""" &domain&"KS_Editor/images/jscript.gif"" border=0 $>"
            Content = Replace(Content, Match.Value, strTemp)
        Next
		ReplaceScriptToImg=content
		'ReplaceScriptToImg=Replace(content,"""","[")
		End function
		function ReplaceImgToScript(Content)
		   Dim regEx,Matches,Match,strTemp2,Match2,strTemp
		   Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "\<IMG(.[^\<]*?)\$\>"
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
			regEx.Pattern = "\#\[(.*?)\]\#"
			Set strTemp = regEx.Execute(Match.Value)
				For Each Match2 In strTemp
					strTemp2 = Replace(Match2.Value, "&amp;", "&")
					strTemp2 = Replace(strTemp2, "#", "")
					strTemp2 = Replace(strTemp2,"&13;&10;",vbCrLf)
					strTemp2 = Replace(strTemp2,"&9;",vbTab)
					strTemp2 = Replace(strTemp2,"¡ä","'")
					strTemp2 = Replace(strTemp2, "[!", "<")
					strTemp2 = Replace(strTemp2, "!]", ">")
					Content = Replace(Content, Match.Value, strTemp2)
				Next
			 Next
			 		Content = Replace(Content, "<HEAD", "<head")
					Content = Replace(Content, "</HEAD>", "</head>")
					Content = Replace(Content, "<TITLE", "<title")
					Content = Replace(Content, "</TITLE>", "</title>")
					Content = Replace(Content, "<liNK", "<link")
					Content = Replace(Content, "><link", ">" & vbcrlf &"<link")
					Content = Replace(Content, "></head", ">" & vbcrlf &"</head")
					Content = Replace(Content, "<META", "<meta")
					Content = Replace(Content, "<BODY", "<body")
					Content = Replace(Content, "</BODY>", "</body>")
					'Content = Replace(Content, "<body contentEditable=true", "<body")
		    ReplaceImgToScript= Content
		End function
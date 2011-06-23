<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Function BytesToBstr(strBody,CodeBase)
       dim obj
       set obj=Server.CreateObject("Adodb.Stream")
       obj.Type=1
       obj.Mode=3
       obj.Open
       obj.Write strBody
       obj.Position=0
       obj.Type=2
       obj.Charset=CodeBase
       BytesToBstr=obj.ReadText
       obj.Close
       set obj=nothing
End Function
If Request.QueryString("Action")="utf8" then
	Response.contentType="application/xml"
	url=trim(request("feedurl"))
	Set http=Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET",url,False
	http.send
	if http.status="200" then
		response.Write http.responseText
	end if
Else
	Response.contentType="application/xml"
	url=request("feedurl")
	Set xml=Server.CreateObject("Microsoft.XMLHTTP")
	xml.Open "GET",url,False
	xml.send
	if xml.status="200" then
		response.Write BytesToBstr(xml.responseBody,"GB2312")
	end if
	set xml=nothing
End IF
%>
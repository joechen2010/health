<%
dim guestimagesnum,imagespath,emotnum,emotpath
imagespath="../Images/face/"
emotpath="../images/emot/"
guestimagesnum=56
emotnum=23

action=request("action")
select case action
	case "face"
		PageTitle="请选择头像"
		call guestface()
	case "emot"
		PageTitle="请选择表情"
		call emot()
end select

%>
<html>
<head>
<title><%=PageTitle%></title>
<script>
window.focus()
function changeimage(imagename)
{ 
	window.opener.document.myform.pic.value=imagename;
	window.opener.document.myform.showimages.src="<%=imagespath%>"+imagename+".gif";
}
</script>
<style type="text/css">
A{TEXT-DECORATION: none;}
A:hover{COLOR: #0099FF;}
A:link {color: #205064;}
A:visited {color: #006699;}
BODY{
FONT-FAMILY: 宋体;
FONT-SIZE: 9pt;
text-decoration: none;
line-height: 150%;
background-color: #FBFDFF;}
TD{
FONT-FAMILY:宋体;
FONT-SIZE: 9pt;}
Input{
FONT-SIZE: 9pt;
HEIGHT: 20px;}
Button{
FONT-SIZE: 9pt;
HEIGHT: 20px; }
Select{
FONT-SIZE: 9pt;
HEIGHT: 20px;}
.border{border: 1px solid #CCCCCC;}
.border2{
background:#fef8ed;
BORDER-RIGHT: #999999 1px solid; 
BORDER-LEFT: #999999 1px solid}
.title{background:#f6f6f6;}
</style>
</head>
<body>
<% sub guestface()%>
<table align=center width=95% cellpadding=5>
<tr>
 <td>
<%

for i=1 to 9
	Response.Write "<a href=""javascript:window.close();""><img src='"&imagespath&i&".gif' border=0 onclick=""changeimage("&i&") "" style=cursor:pointer></a> "
next
for i=10 to guestimagesnum
	Response.Write "<a href=""javascript:window.close();""><img src='"&imagespath&""&i&".gif' border=0 onclick=""changeimage("&i&") "" style=cursor:pointer></a> "
next
%>
</td></tr>
</table>
<%end sub%>

<% sub emot()%>
<table align=center width=95% cellpadding=5><td>
<%

for i=1 to emotnum
	Response.Write "<a href=""javascript:window.close();""><img src='"&emotpath&"emot"&i&".gif' border=0 onclick=""window.opener.document.formwrite.GuestContent.value+='[emot"&i&"]' "" style=cursor:pointer></a> "
	if i mod 6 =0 then
		Response.Write "<br>"
	end if
next
%>
</td></tr></table>
<%end sub%>
<div align="center"><font size="2">[<a href="javascript:window.close();">关闭窗口</a>]</font></div> 

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp" --> 
<!--#include file="md5.asp"-->
<!--#include file="string.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>登录中……</title>
</head>
<body>
<%
dim username
dim password
username=stringfilter(request.form("username"))
password=stringfilter(request.form("password"))
sql="select * from userinfo where username='"&username&"'"
rs.open sql,conn,1,3
if rs.recordcount>0 then
	if password<>"" and rs("password")=md5(password) then
		response.cookies("user")("username")=username
		response.cookies("user")("password")=md5(password)
		if request.form("zddl")="yes" then
			response.expires=#December 31,2099#
		end if
		response.redirect("index.asp")
	else
		response.redirect("error.asp?id=2")
	end if
else
	response.redirect("error.asp?id=2")
end if
%>
</body>
</html>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"--> 
<!--#include file="md5.asp"-->
<!--#include file="string.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>
注册中...
</title>
</head>
<body>
<center>
<font color="#FF0000">
<%
response.write("<a href=index.asp>首页</a>")
rs.open "select * from type",conn,3,1
rs.movefirst
if rs.recordcount>0 then
	for i=1 to rs.recordcount
		response.write("|<a href=type.asp?id="&rs("ID")&">"&rs("type")&"</a>") 
		rs.movenext
	next
end if
rs.close
%>
</font>
</center>
<%
	dim username
	dim password
	dim password2
	dim email
	username=stringfilter(request.form("username"))
	password=stringfilter(request.form("password"))
	password2=stringfilter(request.form("password2"))
	email=stringfilter(request.form("e-mail"))
	if username="" or password="" or password2="" or email="" then
		response.redirect("error.asp?id=7&url=javascript:history.back()")
	end if
	if password<>password2 then
		response.redirect("error.asp?id=9&url=javascript:history.back()")
	end if
	sql="select * from userinfo where username='"&username&"'"
	rs.open sql,conn,3,1
	if rs.recordcount>0 then
		rs.close
		response.redirect("error.asp?id=6&url=javascript:history.back()")
	end if
	rs.close
	rs.open "select * from userinfo where [e-mail]='"&email&"'",conn,3,1
	if rs.recordcount>0 then '邮箱已存在
		rs.close
		response.redirect("error.asp?id=13&url=javascript:history.back()")
	end if
	rs.close
	rs.open "select * from userinfo",conn,3,3
	conn.begintrans
		rs.addnew
		rs("id")=rs.recordcount+1
		rs("username")=username
		rs("password")=md5(password)
		rs("e-mail")=email
		rs.update
	If err.number = 0 then
		conn.committrans
		response.redirect("registersuccessful.asp?username="&username&"&e-mail="&email)
Else
		conn.rollbacktrans
		response.redirect("error.asp?id=8&url=javascript:history.back()")
End if
%>
</body>
</html>
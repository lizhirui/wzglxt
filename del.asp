<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="css.css"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<%
dim wzidh
rs.open "select * from wz",conn,3,3 
if rs.recordcount>0 then 
	rs.movefirst 
	for i=1 to rs.recordcount 
		if request.form("sc"&rs("id"))="删除" then
			wzidh=rs("id")
			exit for 
		end if 
		rs.movenext 
	next 
end if 
rs.close 
conn.begintrans
rs.open "select * from wz where id="&wzidh,conn,3,3
rs.delete
rs.update
rs.close
rs.open "select * from userinfo where username='"&request.cookies("user")("username")&"'",conn,3,3
rs("ftl")=rs("ftl")-1
rs.update
rs.close
if conn.errors.count>0 then
	conn.rollbacktrans
else
	conn.committrans
end if
response.redirect("information.asp")
%>
</body>
</html>

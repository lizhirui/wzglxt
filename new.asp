<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="css.css"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>
<body>
<%
if trim(request.form("wzname"))="" and trim(request.form("wztext"))="" then
	response.redirect("information.asp")
end if
conn.begintrans
rs.open "select * from wz",conn,1,3
if request.form("wz")="修改" then
	if rs.recordcount>0 then
		rs.movefirst
		for i=1 to rs.recordcount
			if val(rs("id"))=val(request.form("h_id")) then
				exit for
			end if
			rs.movenext
		next
	else
		response.redirect("information.asp")
	end if
		
else
	rs.addnew
	rs("id")=rs.recordcount+1
end if
rs("wzname")=request.form("wztitle")
rs("wzusername")=request.cookies("user")("username")
rs("wztext")=request.form("wztext")
rs("wztypeID")=request.form("wztype")
rs.update
rs.close
response.write(err.Description)
rs.open "select * from userinfo where username='"&request.cookies("user")("username")&"'",conn,1,3
rs("ftl")=rs("ftl")+1
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

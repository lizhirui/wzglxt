<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�һ��û���</title>
</head>
<body>
<div class="setcenter">
<img src="logo.JPG" />
</div>
<center>
<font color="#FF0000">
<%
response.write("<a href=index.asp>��ҳ</a>")
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
<form id="form1" name="form1" method="get" action="zhpassword2.asp">
�������û�����<input type="text" name="username" class="css_text"/>
<input name="qd" type="submit" value="ȷ��"/>
</form>
</body>
</html>

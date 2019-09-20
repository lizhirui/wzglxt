<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>退出</title>
</head>
<body>
<%
response.cookies("user").Expires=(now()-1)
response.Redirect("index.asp")
%>
</body>
</html>

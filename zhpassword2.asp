<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"-->
<!--#include file="string.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>更改密码</title>
</head>
<body>
<div class="setcenter">
<img src="logo.JPG" />
</div>
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
<form id="form1" name="form1" method="post" action="zhpassword3.asp">
<%
request.querystring("username")=stringfilter(request.querystring("username"))
randomize()
yzm=int((999999-100000+1)*rnd+100000)
response.session("zhpassword")("yzm")=yzm
rs.open "select * from userinfo where username='"&request.querystring("username")&"'",conn,3,3
if rs.recordcount<1 then '用户名不存在
	rs.close
	response.redirect("error.asp?id=11&url=zhpassword.asp")
end if
response.cookies("zhpassword")("username")=rs("username")

Set objMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.163.com"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "wzglxt@163.com"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "lizhirui"
objCDOSYSCon.Fields.Update
Set objMail.Configuration = objCDOSYSCon
objMail.From ="wzglxt@163.com"
objMail.Subject = "文章管理系统密码找回" 
objMail.To = rs("e-mail")
objMail.HtmlBody = "验证码为："&yzm
objMail.Send
Set objMail = Nothing
Set objCDOSYSCon = Nothing
rs.close
%>
验证码已经发入您的邮箱
<br>
&nbsp;&nbsp;验证码：<input type="text" name="yzm2" class="css_text"/>
<br>
&nbsp;&nbsp;新密码：<input type="password" name="password" class="css_text"/>
<br>
确认密码：<input type="password" name="password2" class="css_text"/>
<br>
<input name="qd" type="submit" value="确定"/>
</form>
</body>
</html>

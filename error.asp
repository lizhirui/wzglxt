<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>错误<%response.write request.querystring("id")%></title>
</head>
<body>
<%
select case request.querystring("id")
case 1
response.write("搜索选项错误！")
case 2
response.write("密码错误！")
case 3
response.write("文章号不是数字或为0！")
case 4
response.write("文章不存在!")
case 5
response.write("用户类型错误！")
case 6
response.write("用户名已存在！")
case 7
response.write("信息不完整！")
case 8
response.write("数据库读写错误！")
case 9
response.write("两次输入的密码不相同！")
case 10
response.write("类别不存在！")
case 11
response.write("邮箱不存在！")
case 12
response.write("用户名不存在！")
case 13
response.write("邮箱已存在！")
case 14
response.write("验证码错误！")
case 15
response.write("cookies中的用户名或密码为空！")
case 16
response.write("您不是管理员！")
case 17
response.write("检测到SQL注入！")
case else
response.write("错误码不存在！")
end select
%>
<br>
<%
if trim(request.querystring("url"))<>"" then
	response.write("<a href="&trim(request.querystring("url"))&""+">返回</a><br>")
end if
response.write("<a href=index.asp>返回主页</a>")
%>
</body>
</html>

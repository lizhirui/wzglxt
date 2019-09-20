<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>注册成功</title>
</head>
<body>
<table width="100%">	
	<tr>
		<td width="60%" height="20">
			<p>今天是
  <%	
		Response.Write(year(date) & "年" & month(date) & "月" & day(date) & "日<br>" )
		if hour(time)>4 and hour(time)<=7 then
		response.write("早上好！")
		elseif hour(time)>7 and hour(time)<12 then
		response.write("上午好！")
		elseif hour(time)=12 then
		response.write("中午好！")
		elseif hour(time)>12 and hour(time)<18 then
		response.write("下午好！")
		elseif hour(time)>=18 and hour(time)<=23 then
		response.write("晚上好！")
		else
		response.write("凌晨好！")
		end if
%>
</p>
	  </td>
		<td align="right">
		<%			
			if request.cookies("user")("username")<>"" and request.cookies("user")("password")<>"" then
				rs.open "select * from userinfo where [username]='"&request.cookies("user")("username")&"'",conn,1,3			
				if rs.recordcount>0 then
					if rs("password")<>request.cookies("user")("password") then
						response.write("<a href=login.asp>登录</a>")	
						response.write("&nbsp")
						response.write("<a href=register.asp>注册</a>")
					else
						if	rs("type")=0 then
							response.write(request.cookies("user")("username")&" 在线(普通用户)")
						elseif rs("type")=1 then
							response.write(request.cookies("user")("username")&" 在线(管理员)")
						else
							response.redirect("error.asp?id=5")
						end if
						response.write("&nbsp")
						response.write("<a href=information.asp>个人中心</a>")
						response.write("&nbsp")
						if rs("type")=1 then
							response.write("<a href=systemadmin.asp>系统管理</a>&nbsp;")
						end if
						response.write("<a href=exit.asp>退出</a>")
					end if
				else
					response.write("<a href=login.asp>登录</a>")
					response.write("&nbsp")
					response.Write("<a href=register.asp>注册</a>")
				end if
				rs.close
			else
				response.write("<a href=login.asp>登录</a>")
				response.write("&nbsp")
				response.write("<a href=register.asp>注册</a>")
				response.write("&nbsp")
				response.write("<a href=zhusername.asp>找回用户名</a>")
				response.write("&nbsp")
				response.write("<a href=zhpassword.asp>找回密码</a>")
			end if
		%>	
		</td>
	</tr>
</table>
<center>
<img src=logo.jpg>
</center>
<table width="100%" border="0">
<tr valign="top" height="1">
<td colspan="3">
<form id="form1" name="form1" method="get" action="wzsearch.asp">
搜索网页信息：
<input type="text" name="wzmc" class="css_text"/>
搜索方式：
<select name="searchoption">
  <option value="1" >文章名称</option>
  <option value="2">发表人</option>
</select>
<input type="submit" name="Submit" value="搜索" />
</form>
</tr>
<tr valign="top" height="20">
<td colspan="3" bgcolor="#00FF00" >
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
</td>
</tr>
</table>
<%
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
objMail.Subject = "文章管理系统注册信" 
objMail.To = request.querystring("e-mail")
objMail.HtmlBody = request.querystring("username")&"用户已注册成功！请保留好以下内容"&"<br>用户名:"&request.querystring("username")&"<br>E-mail:"&request.querystring("e-mail")
objMail.Send
Set objMail = Nothing
Set objCDOSYSCon = Nothing
response.write(request.querystring("username")&"用户已注册成功！请保留好以下内容，这些内容已被发送至您的邮箱"&"<br>用户名:"&request.querystring("username")&"<br>E-mail:"&request.querystring("e-mail")&"<br><a href=127.0.0.1/login.asp>登陆</a>")
%>
</body>
</html>

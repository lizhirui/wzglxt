<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>密码更改成功</title>
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
rs.open "select * from userinfo where username='"&request.cookies("zhpassword")("username")&"'",conn,3,3
if recordcount<0 then
	response.redirect("error.asp?id=12&url=zhpassword.asp")
else if request.session("zhpassword")("yzm")<>trim(request.form("yzm2")) then
	response.redirect("error.asp?id=14&url=zhpassword.asp")
else if trim(request.form("password"))="" or trim(request.form("password2"))="" then
	response.redirect("error.asp?id=7&url=zhpassword.asp")
else if trim(request.form("password"))<>trim(request.form("password2")) then
	response.redirect("error.asp?id=9&url=zhpassword.asp")
end if
end if
end if
end if
rs("password")=md5(trim(request.form("password")))
rs.update
rs.close
%>
密码更改成功！
<a href=login.asp>登录</a>
<a href=index.asp>返回主页</a>
</body>
</html>

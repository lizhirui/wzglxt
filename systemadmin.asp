<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="string.asp"-->
<!--#include file="conn_admin.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>系统管理</title>
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
<%
on error resume next
dim username
dim password
username=stringfilter(request.cookies("user")("username"))
password=stringfilter(request.cookies("user")("password"))
if username="" or password="" then
	response.redirect("../error.asp?id=15")
end if
rs.open "select * from userinfo where username='"&username&"'",conn,3,1
if rs.recordcount<=0 then
	response.redirect("../error.asp?id=12")
	rs.close
end if
if password<>stringfilter(rs("password")) then
	response.redirect("../error.asp?id=2")
	rs.close
end if
if rs("type")<>1 then
	response.redirect("../error.asp?id=16")
end if
rs.close
%>
<form id="form1" name="form1" method="post" action="systemadmin.asp">
SQL语句：
<p>
  <textarea name="sql" rows="15" cols="36"><% response.write(trim(request.form("sql"))) %></textarea>
</p>
<p>
<input type="submit" name="sqlexecute" value="执行"/>
</p>
</form>
<%
sqlvalue=trim(request.form("sql"))
if trim(request.form("sql"))<>"" then
	if left(trim(request.form("sql")),6)="select" then '查询语句
		rs.open trim(request.form("sql")),conn,3,3
		response.write("执行结果:<table border=1><tr>")
		for i=0 to rs.fields.count-1 '输出列名
			set sqlfields=rs.fields.item(i)
			response.write("<td>"&sqlfields.name&"</td>")
			set sqlfields=nothing
		next
		response.write("</tr>")
		if rs.recordcount>0 then
			rs.movefirst
			for i=1 to rs.recordcount
				response.write("<tr>")
				for j=0 to rs.fields.count-1
					response.write("<td>"&rs(j)&"</td>")
				next
				response.write("</tr>")
				rs.movenext
			next
		end if
		rs.close
		response.write("</table>")
	else
		conn.begintrans
		conn.execute(trim(request.form("sql")))
		if conn.errors.count>0 then
			conn.rollbacktrans
			for i=0 to conn.errors.count-1
				set connerr=conn.errors.item(i)
				response.write("执行结果:<br>错误码:"&connerr.number&"<br>错误信息:"&connerr.description)
				set connerr=nothing
			next
		else
			conn.committrans
			response.write("执行结果:执行成功")
		end if
	end if
end if
%>
</center>
</body>
</html>

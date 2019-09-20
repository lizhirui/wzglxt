<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp" --> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>
<%
Response.Write(Trim(request.form("wzmc")))
%>
</title>
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
<img src="logo.JPG" />
</center>
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
<table border="1">
<tr>
<td>文章名称</td>
<td>发表人</td>
<%
if request.querystring("searchoption")=1 then
	sql="SELECT * FROM wz where [wzname] like '%" & request.querystring("wzmc")& "%' order by [wzname] asc"
elseif request.querystring("searchoption")=2 then
	sql="SELECT * FROM wz where [wzusername] like '%" & request.querystring("wzmc")& "%' order by [wzname] asc"
else
	response.Redirect("error.asp?id=1") '"searchoption"值错误，转向错误页面，错误码：1。
end if
rs.open sql,conn,3,1
do while not rs.eof
	response.Write("<tr>")
	response.Write("<td>")
	response.write("<a href='readwz.asp?id="&rs("ID")&"' target='_blank'>") 
	response.write(rs("wzname"))
	response.Write("</a>")
	response.Write("</td>")
	response.Write("<td>")
	response.write(rs("wzusername"))
	response.Write("</td>")
	response.Write("</tr>")
	rs.MoveNext
	loop
%>
</tr>
</table>
</center>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="checklogin.asp" -->
<!--#include file="conn.asp"-->
<!--#include file="string.asp"-->
<!--#include file="functions.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>个人中心</title>
</head>
<body>
<center>
<%
on error resume next
dim username
dim password
dim wzidh
dim wztitle
dim wztype
dim wztext
rs.open "select * from wz",conn,3,3 
if rs.recordcount>0 then 
	rs.movefirst 
	for i=1 to rs.recordcount 
		if request.form("xg"&rs("id"))="修改" then
			wzidh=rs("id")
			wztitle=rs("wzname")
			wztype=rs("wztypeID")
			wztext=rs("wztext")
			exit for
		end if 
		rs.movenext 
	next 
end if 
rs.close 
username=stringfilter(request.cookies("user")("username"))
password=stringfilter(request.cookies("user")("password"))
if username="" or password="" then
	response.redirect("error.asp?id=15")
end if
rs.open "select * from userinfo where username='"&username&"'",conn,3,1
if rs.recordcount<=0 then
	rs.close
	response.redirect("error.asp?id=12")
end if
if password<>stringfilter(rs("password")) then
	rs.close
	response.redirect("error.asp?id=2")
end if
rs.close
if request.form("qd")="修改密码" then
	if stringfilter(request.form("oldpassword"))="" or stringfilter(request.form("password"))="" or stringfilter(request.form("password2"))="" then
		response.redirect("error.asp?id=7&url=information.asp")
	end if
	if stringfilter(request.form("password"))<>stringfilter(request.form("password2")) then
		response.redirect("error.asp?id=9&url=information.asp")
	end if
	rs.open "select * from userinfo where username='"&request.cookies("user")("username")&"'",conn,3,3
	if rs.recordcount>0 then
		if rs("password")<>md5(stringfilter(request.form("oldpassword"))) then
			rs.close
			response.redirect("error.asp?id=2&url=information.asp")
		else
			rs("password")=md5(stringfilter(request.form("password")))
			response.cookies("user")("password")=stringfilter(request.form("oldpassword"))
			rs.update
			rs.close
			response.redirect("exit.asp")
		end if
	else
		rs.close
		response.redirect("error.asp?id=12&url=information.asp")
	end if
end if
%>
<form id="passwordform" name="passwordform" method="post" action="information.asp" OnSubmit="return confirm('真的要修改密码吗?')">
旧密码：<input name="oldpassword" type="password" class="css_text"/> 
新密码：<input name="password" type="password" class="css_text"/> 
确认密码：<input name="password2" type="password" class="css_text" />
<input name="qd" type="submit" value="修改密码"/>
</form>
<form id="wzform" name="wzform" method="post" action="new.asp" OnSubmit="return confirm('真的要提交吗?')">
文章标题：<input name="wztitle" type="text" class="css_text" value=""/>
文章类型：
<select name="wztype">
<%
rs.open "select * from type"
if rs.recordcount>0 then
	rs.movefirst
	for i=1 to rs.recordcount
		response.write("<option value="&rs("ID")&">"&rs("type")&"</option>")
		if rs("ID")=wztype then
			wztype=i
		end if
		rs.movenext
	next
end if
rs.close
%>
</select>
<br><br>文章内容：<p><textarea name="wztext" rows="20%" cols="120%"></textarea>
<p>
<input name="wz" id="wz" type="submit" value="添加">
<a href=information.asp>返回</a>
</p>
<input type="hidden" name="h_id" id="h_id" value=<% response.write(wzidh) %>>
<input type="hidden" name="h_title" id="h_title" value=<% response.write(Server.UrlEncode(wztitle)) %>>
<input type="hidden" name="h_type" id="h_type" value=<% response.write(wztype) %>>
<input type="hidden" name="h_text" id="h_text"value=<% response.write(Server.UrlEncode(wztext)) %>>
</form>
<script language="VBScript">
wzid=document.wzform.h_id.value
wztitle=document.wzform.h_title.value
wztext=document.wzform.h_text.value
wztype=document.wzform.h_type.value
if wzid>0 then
	document.wzform.wztitle.value=URLDecode(wztitle)
	document.wzform.wztype.value=wztype
	document.wzform.wztext.value=URLDecode(wztext)
	document.wzform.wz.value="修改"
end if
</script>
<table border=1>
<tr>
<td>ID</td>
<td>文章名称</td>
<td>文章类型</td>
<td>修改</td>
<td>删除</td>
</tr>
<%
dim rscount
dim typeid
rs.open "select * from wz where wzusername='"&request.cookies("user")("username")&"'",conn,3,3
rscount=rs.recordcount
if rscount>0 then
	rs.movefirst
	for i=1 to rscount
		response.write("<tr>")
		response.write("<td>"&rs("ID")&"</td>")
		response.write("<td><a href=readwz.asp?id="&rs("id")&">"&rs("wzname")&"</a></td>")
		typeid=rs("wztypeID")
		rs.close
		rs.open "select * from type where id="&typeid,conn,3,3
		response.write("<td>"&rs("type")&"</td>")
		rs.close
		rs.open "select * from wz where wzusername='"&request.cookies("user")("username")&"'",conn,3,3
		rs.movefirst
		if i>1 then
			for j=1 to i-1
				rs.movenext
			next
		end if
		response.write("<td><form id="&chr(34)&"scform"&rs("id")&chr(34)&" name="&chr(34)&"scform"&rs("id")&chr(34)&" method="&chr(34)&"post"&chr(34)&" action="&chr(34)&"information.asp"&chr(34)&"><input name="&chr(34)&"xg"&rs("id")&chr(34)&" type="&chr(34)&"submit"&chr(34)&" value="&chr(34)&"修改"&chr(34)&"/></form></td>"&"<td><form id="&chr(34)&"scform"&rs("id")&chr(34)&" name="&chr(34)&"scform"&rs("id")&chr(34)&" method="&chr(34)&"post"&chr(34)&" action="&chr(34)&"del.asp"&chr(34)&" OnSubmit="&chr(34)&"return confirm('真的要删除吗?')"&chr(34)&"><input name="&chr(34)&"sc"&rs("id")&chr(34)&" type="&chr(34)&"submit"&chr(34)&" value="&chr(34)&"删除"&chr(34)&"/></form></td>")
		response.write("</tr>")
		rs.movenext
	next
end if
rs.close
%>
</table>
</center>
</body>
</html>

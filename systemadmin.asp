<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="string.asp"-->
<!--#include file="conn_admin.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ϵͳ����</title>
</head>
<body>
<table width="100%">	
	<tr>
		<td width="60%" height="20">
			<p>������
  <%	
		Response.Write(year(date) & "��" & month(date) & "��" & day(date) & "��<br>" )
		if hour(time)>4 and hour(time)<=7 then
		response.write("���Ϻã�")
		elseif hour(time)>7 and hour(time)<12 then
		response.write("����ã�")
		elseif hour(time)=12 then
		response.write("����ã�")
		elseif hour(time)>12 and hour(time)<18 then
		response.write("����ã�")
		elseif hour(time)>=18 and hour(time)<=23 then
		response.write("���Ϻã�")
		else
		response.write("�賿�ã�")
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
						response.write("<a href=login.asp>��¼</a>")	
						response.write("&nbsp")
						response.write("<a href=register.asp>ע��</a>")
					else
						if	rs("type")=0 then
							response.write(request.cookies("user")("username")&" ����(��ͨ�û�)")
						elseif rs("type")=1 then
							response.write(request.cookies("user")("username")&" ����(����Ա)")
						else
							response.redirect("error.asp?id=5")
						end if
						response.write("&nbsp")
						response.write("<a href=information.asp>��������</a>")
						response.write("&nbsp")
						if rs("type")=1 then
							response.write("<a href=systemadmin.asp>ϵͳ����</a>&nbsp;")
						end if
						response.write("<a href=exit.asp>�˳�</a>")
					end if
				else
					response.write("<a href=login.asp>��¼</a>")
					response.write("&nbsp")
					response.Write("<a href=register.asp>ע��</a>")
				end if
				rs.close
			else
				response.write("<a href=login.asp>��¼</a>")
				response.write("&nbsp")
				response.write("<a href=register.asp>ע��</a>")
				response.write("&nbsp")
				response.write("<a href=zhusername.asp>�һ��û���</a>")
				response.write("&nbsp")
				response.write("<a href=zhpassword.asp>�һ�����</a>")
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
SQL��䣺
<p>
  <textarea name="sql" rows="15" cols="36"><% response.write(trim(request.form("sql"))) %></textarea>
</p>
<p>
<input type="submit" name="sqlexecute" value="ִ��"/>
</p>
</form>
<%
sqlvalue=trim(request.form("sql"))
if trim(request.form("sql"))<>"" then
	if left(trim(request.form("sql")),6)="select" then '��ѯ���
		rs.open trim(request.form("sql")),conn,3,3
		response.write("ִ�н��:<table border=1><tr>")
		for i=0 to rs.fields.count-1 '�������
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
				response.write("ִ�н��:<br>������:"&connerr.number&"<br>������Ϣ:"&connerr.description)
				set connerr=nothing
			next
		else
			conn.committrans
			response.write("ִ�н��:ִ�гɹ�")
		end if
	end if
end if
%>
</center>
</body>
</html>

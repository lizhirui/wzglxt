<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"-->
<link href="css.css" rel="stylesheet" type="text/css">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ע��</title>
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
<div style="width:100%; text-align:center; clear:top">
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
<form id="form1" name="form1" method="post" action="register2.asp">
<table><tbody>
<tr><td>&nbsp;&nbsp;�û���:</td><td><input type="text" name="username" class="css_text" /></td><td>ע��50������</td></tr>
<tr><td>&nbsp;&nbsp;&nbsp;&nbsp;����:</td><td><input type="password" name="password" class="css_text"/></td><td>ע��������</td></tr>
<tr><td> ȷ������:</td><td><input type="password" name="password2" class="css_text"/></td><td>ע��Ҫ��������ͬ</td></tr>
<tr><td>&nbsp;&nbsp;E-mail:</td><td><input type="text" name="e-mail" class="css_text"/></td><td>ע������ȡ�������;�����벻Ҫ������</td></tr>
</tbody></table>
<input name="ע��" type="submit" value="ע��"/>
</form>
</div>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<!--#include file="conn.asp"--> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>���¹���ϵͳ</title>
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
<img src=logo.jpg>
</center>
<table width="100%" height="60%" border="0">
<tr valign="top" height="1">
<td colspan="3">
<form id="form1" name="form1" method="get" action="wzsearch.asp">
������ҳ��Ϣ��
<input type="text" name="wzmc" class="css_text"/>
������ʽ��
<select name="searchoption">
  <option value="1" >��������</option>
  <option value="2">������</option>
</select>
<input type="submit" name="Submit" value="����" />
</form>
</tr>
<tr valign="top" height="20">
<td colspan="3" bgcolor="#00FF00" >
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
</td>
</tr>
<tr valign="top">
<td width="24%" bgcolor="#0000FF">
<center><font size="5">��¼</font></center>
  <form id="form2" name="form2" method="post" action="login2.asp">
  <p>�û�����
    <input type="text" name="username" class="css_text" />
    </p>
  <p> &nbsp;&nbsp;���룺
    <input type="password" name="password" class="css_text" />
</p>
  <p>
    <input type="checkbox" name="zddl" value="yes" />
  �Զ���¼</p>
  <p>
    <input type="submit" name="login" value="��¼" />
  </p>
</form></td>
<td width="58%" bgcolor="#3366FF">
<font size="5">�������£�</font>
<br>
<%
rs.open "select top 10 * from wz order by llcs desc",conn,3,1
rs.movefirst
if rs.recordcount>0 then
	for i=1 to rs.recordcount
		response.write("<a href=readwz.asp?id="&rs("ID")&">"&rs("wzname")&"</a><br>")
		rs.movenext
	next
end if
rs.close
%>    </td>
	<td width="18%" bgcolor="#0000FF">
	<font size="5">������������<br>�û���,������</font>
	<br>
	<%
	rs.open "select top 10 * from userinfo order by ftl desc",conn,3,1
	rs.movefirst
	if rs.recordcount>0 then
	for i=1 to rs.recordcount
		response.write rs("username")&" "&rs("ftl")&"<br>"
		rs.movenext
	next
response.end
	end if
	rs.close
	%>	</td>
  </tr>
</table>
<hr align="center" size="1" />
<div style="width:100%; text-align:center; clear:top">
  <p><strong><font size=4>��ϵ��ʽ</font></strong><br>
    E-mail:exiis@126.com,578341256@qq.com,859067292@qq.com
    <br>
    QQ:
    <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=578341256&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:578341256:41 &r=0.4908671043357943" alt="578341256" title="578341256"></a>
  <a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=859067292&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:859067292:41" alt="859067292" title="859067292"></a></p>
</div> 
</body>
</html>

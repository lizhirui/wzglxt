<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="css.css"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����<%response.write request.querystring("id")%></title>
</head>
<body>
<%
select case request.querystring("id")
case 1
response.write("����ѡ�����")
case 2
response.write("�������")
case 3
response.write("���ºŲ������ֻ�Ϊ0��")
case 4
response.write("���²�����!")
case 5
response.write("�û����ʹ���")
case 6
response.write("�û����Ѵ��ڣ�")
case 7
response.write("��Ϣ��������")
case 8
response.write("���ݿ��д����")
case 9
response.write("������������벻��ͬ��")
case 10
response.write("��𲻴��ڣ�")
case 11
response.write("���䲻���ڣ�")
case 12
response.write("�û��������ڣ�")
case 13
response.write("�����Ѵ��ڣ�")
case 14
response.write("��֤�����")
case 15
response.write("cookies�е��û���������Ϊ�գ�")
case 16
response.write("�����ǹ���Ա��")
case 17
response.write("��⵽SQLע�룡")
case else
response.write("�����벻���ڣ�")
end select
%>
<br>
<%
if trim(request.querystring("url"))<>"" then
	response.write("<a href="&trim(request.querystring("url"))&""+">����</a><br>")
end if
response.write("<a href=index.asp>������ҳ</a>")
%>
</body>
</html>

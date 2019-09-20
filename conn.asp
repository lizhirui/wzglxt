<!--#include file="css.css"-->
<!--#include file="cookiecheck.asp"-->
<%
on error resume next
sub ca()
set rs=nothing
conn.close
set conn=nothing
end sub
Set conn=Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=; Data Source="&Server.MapPath("wz.mdb")
If Err Then
		Response.Write ""&IsSqlVer&"数据库连接出错，请检查连接字串。<br /><br />"&Err.Source&" ("&Err.Number&")"
		Set Conn = Nothing
		err.Clear
		Response.End
End If
%>

<%
code=array(",","'","<",">",chr(34)," ",";")
If Request.QueryString<>"" Then
	For Each SQL_Get In Request.QueryString
		For SQL_Data=0 To Ubound(code)
			if instr(Request.QueryString(SQL_Get),code(Sql_DATA))>0 Then
				response.redirect("error.asp?id=17")
			end if
		next
	Next
End If
If Request.Form<>"" Then
	For Each Sql_Post In Request.Form
		For SQL_Data=0 To Ubound(code)
			if instr(LCase(Request.Form(Sql_Post)),code(Sql_DATA))>0 Then
				if Sql_Post<>"h_text" and Sql_Post<>"h_title" and Sql_Post<>"wztitle" and Sql_Post<>"wztext" then
					response.redirect("error.asp?id=17")
				end if
			end if
		next
	next
end if
If Request.Cookies<>"" Then 
	For Each Fy_cook In Request.Cookies 
		For Fy_Xh=0 To Ubound(code) 
			If Instr(LCase(Request.Cookies(Fy_cook)),code(Fy_Xh))>0 Then 
				response.redirect("error.asp?id=17")
			End If 
		Next 
	Next 
end if
%>
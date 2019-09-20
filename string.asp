<%
'×Ö·û´®¹ýÂË
function stringfilter(byval sql_value)
	if trim(sql_value)<>"" then
		sql_value=trim(sql_value)
	else
		exit function
	end if
	code=array(",","'","<",">",chr(34)," ",";")
	for i=0 to ubound(code)
		for j=1 to len(sql_value)
			if mid(sql_value,j,1)=code(i) then
				response.redirect("error.asp?id=17")
			end if
		next
	next
	stringfilter=sql_value
end function
%>
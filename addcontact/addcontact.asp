<%
' Redirect user to correct page

Select Case Request.QueryString("QueryType")

	Case "" Response.Redirect("default.asp")
	Case "SearchContact" Response.Redirect("SearchContact.asp")

End Select




%>
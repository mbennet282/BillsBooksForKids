<%
if ubound(Session("shoppingCartItems")) <> -1 then
	Session("shoppingCartItems") = Array()
	Response.Redirect("shoppingcart.asp")
end if
%>
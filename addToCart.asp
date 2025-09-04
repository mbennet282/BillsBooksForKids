<%
Option Explicit
%>
<!--#include virtual="/inc/vars.inc"-->
<!--#include virtual="/inc/functions.inc"-->
<!--#include virtual="/inc/subs.inc"-->

<%
if IsNumeric(Request.QueryString("productid")) then
	Dim productId, rs, cmd
	productId = Request.QueryString("productid")
	AddToCart(productId)
	
else%>
	<p>productid not found or is not a numeric value</p>
<%end if%>
<!doctype html>
	<body>
		<script>
			window.history.back();
		</script>
	</body>
</html>
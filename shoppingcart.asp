<%Option Explicit%>
<!--#include virtual="/inc/vars.inc"-->
<!--#include virtual="/inc/functions.inc"-->
<!--#include virtual="/inc/subs.inc"-->

<%
title = "My shopping cart"
%>

<!--#include virtual="/inc/head.inc"-->
<%if ubound(Session("shoppingCartItems")) <> -1 then
	dim rs, total
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "select * from Books",Application("conn"),3,3

	'update shopping cart
	if Request.Form("updateCheckoutBtn") = "Update" then
		dim j, k, updatedCart, removedItems
		updatedCart = Array()
		removedItems = Array()

		'determine which items to remove
		if Request.Form("remove").Count > 0 then
			for j = 1 to Request.Form("remove").Count
				Redim Preserve removedItems(ubound(removedItems) + 1)
				removedItems(ubound(removedItems)) = CInt(Request.Form("remove")(j))
				
			next
		end if

		for j = 1 to Request.Form("id").Count
			Dim q
			if not Isnumeric(Request.Form("quantity")(j)) then
				q = 1
			else
				q = Request.Form("quantity")(j)
			end if
		
			for k = 1 to q
				if ubound(filter(removedItems,Request.Form("id")(j))) = -1 then
					Redim Preserve updatedCart(ubound(updatedCart) + 1)
					updatedCart(ubound(updatedCart)) = CInt(Request.Form("id")(j))
				end if
				
			next
		Next
		updatedCart = sort(updatedCart)
		Session("shoppingCartItems") = updatedCart
	end if

	'remove item from shopping cart
	
%>
	<form id="shoppingcart" name="shoppingcart" method="post">
		<table id="shoppingCartInfo">
			<tr>
				<th>ID</th>
				<th>Cover</th>
				<th>Author</th>
				<th>Title</th>
				<th>ISBN</th>
				<th>Price</th>
				<th><label for="quantity">Quantity</label></th>
				<th>Remove</th>
			</tr>
<%
		do until rs.EOF
			Dim i, items, quantity
			items = Session("shoppingCartItems")
			for i = 0 to ubound(items)
			dim continue
			continue = false
			if i > 0 then
				if items(i) = items(i - 1) then
					continue = true
				end if
			end if
			if not continue then
			if rs.Fields.Item("BookID") = items(i) then
				quantity = ubound(filter(items,items(i))) + 1
				total = total + (rs.Fields.Item("BookPrice") * quantity)
		%>
			<tr>
				<td><%=rs.Fields.Item("BookID")%></td>
				<td><img width="160" height="150" src="<%=rs.Fields.Item("BookCover")%>" alt="<%=rs.Fields.Item("BookTitle")%>"/></td>
				<td><%=rs.Fields.Item("AuthorSurname")%>, <%=rs.Fields.Item("AuthorFirstNameInitial")%>.</td>
				<td><%=rs.Fields.Item("BookTitle")%></td>
				<td><%=rs.Fields.Item("BookISBN")%></td>
				<td><%=formatcurrency(rs.Fields.Item("BookPrice"))%></td>
				<td><input type="number" min="1" max="99" name="quantity" value="<%=quantity%>" /></td>
				<td><input type="checkbox" name="remove" value="<%=rs.Fields.Item("BookID")%>"/></td>
				<input type="hidden" name="id" value="<%=rs.Fields.Item("BookID")%>"/>
			</tr>
		<%
			end if
			end if
			next
			rs.MoveNext
		loop%>
		</table>
		<input type="submit" name="updateCheckoutBtn" value="Update"/>
	</form>
	<p>Total amount due: <%=formatcurrency(total)%></p>
	
	<a href="clearcart.asp">Clear shopping cart</a>
<%
	rs.close
	set rs = nothing
else
	%>
	<p>Your shopping cart is empty</p>
	<%
end if%>
<!--#include virtual="/inc/foot.inc"-->
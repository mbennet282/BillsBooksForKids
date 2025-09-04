<%
sub queryAllBooks()
	' create a table containing book information
	Dim field, rs, cols, books, currentPage, startBookOffset, endBookOffset, remainingBooks, remainingBooksDisplay, pageNumbers, i
	i = 1
	cols = 0
	books = 0
	currentPage = 0
	startBookOffset = 0
	endBookOffset = 4

	if IsNumeric(Request.QueryString("page")) then
		if Request.QueryString("page") <= 0 then
			currentPage = 1
		else
			currentPage = Request.QueryString("page")
		end if
	else
		currentPage = 1
	end if

	endBookOffset = 4 * currentPage
	startBookOffset = endBookOffset - 5

	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open "select * from Books",Application("conn"),3,3

	remainingBooks = rs.RecordCount
	remainingBooksDisplay = rs.RecordCount

	remainingBooksDisplay = remainingBooksDisplay - endBookOffset%>
	<table>
		<%
			do until rs.EOF
			if books > startBookOffset and books < endBookOffset then
			if cols = 0 then%>
				<tr>
			<%end if%>
					<td><a href="bookDetails.asp?productid=<%=rs.Fields.Item("BookID")%>"><img width="160" height="150" src="<%=rs.Fields.Item("BookCover")%>" alt="<%=rs.Fields.Item("BookTitle")%>"/></a></td>
					<td>
						<table>
							<tr>
								<th><%=rs.Fields.Item("BookTitle")%></th>
							</tr>
							<tr>
								<td><strong>Author: </strong> <%=rs.Fields.Item("AuthorSurname")%>, <%=rs.Fields.Item("AuthorFirstNameInitial")%>.</td>
							</tr>
							<tr>
								<td><strong>Category: </strong> <%=rs.Fields.Item("BookCategory")%></td>
							</tr>
							<tr>
								<td><strong>Year: </strong> <%=rs.Fields.Item("BookYear")%></td>
							</tr>
							<tr>
								<td><strong>Cost: </strong> <%=FormatCurrency(rs.Fields.Item("BookPrice"))%></td>
							</tr>
							<tr>
								<td><strong>ISBN: </strong> <%=rs.Fields.Item("BookISBN")%></td>
							</tr>
							<tr>
								<td><a href="addToCart.asp?productid=<%=rs.Fields.Item("BookID")%>">Add to cart</a></td>
							</tr>
						</table>
					</td>
					<%cols = cols + 1
					%>
				<%if cols = 2 then%>
				</tr>
				<%cols = 0
				end if
			
			end if
			books = books + 1
			rs.MoveNext
			loop%>
			<tr>
				<td>
				
				<%if IsNumeric(Request.QueryString("page")) then
					if Request.QueryString("page") > 1 then%>
					<a href="?page=<%=currentPage - 1%>">< Back</a>
				<%
					end if
					while remainingBooks > 0
						remainingBooks = remainingBooks - 4
						if i = CInt(currentPage) then
							pageNumbers = pageNumbers & i & " "
						else
							pageNumbers = pageNumbers & "<a href='?page=" & i & "'>" & i & "</a> "
						end if
						i = i + 1
					wend
					%>
					<%=pageNumbers%>
					<%
					if remainingBooksDisplay > 0 then%>
					<a href="?page=<%=currentPage + 1%>">Next ></a>
					<%end if
				end if%>
					
				</td>
			</tr>
		</table>
<%
	rs.close
	set rs = nothing
end sub

sub queryBooks(title)
	' create a table containing book information from a search query
	Dim field, rs, cmd, cols, books, currentPage, startBookOffset, endBookOffset, remainingBooks, remainingBooksDisplay, pageNumbers, i
	i = 1
	cols = 0
	books = 0
	currentPage = 0
	startBookOffset = 0
	endBookOffset = 4

	if IsNumeric(Request.QueryString("page")) then
		if Request.QueryString("page") <= 0 then
			currentPage = 1
		else
			currentPage = Request.QueryString("page")
		end if
	else
		currentPage = 1
	end if

	endBookOffset = 4 * currentPage
	startBookOffset = endBookOffset - 5

	set cmd=Server.CreateObject("ADODB.Command")
	set cmd.ActiveConnection = Application("conn")

	cmd.CommandText = "SELECT * FROM [Books] WHERE [BookTitle] LIKE ?"
	cmd.CommandType = 1
	cmd.ActiveConnection.CursorLocation = 3
	cmd.Parameters.Append cmd.CreateParameter("title", 202, 1, 255)
	cmd.Parameters("title")="%" & title & "%"

	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open cmd.Execute

	if rs.BOF and rs.EOF then
	%>
	<p>No results found</p>
	<%
	rs.close
	set rs = nothing
	set cmd = nothing
	exit sub
	end if
	remainingBooks = rs.RecordCount
	remainingBooksDisplay = rs.RecordCount

	remainingBooksDisplay = remainingBooksDisplay - endBookOffset%>
	<table>
		<%
			do until rs.EOF
			if books > startBookOffset and books < endBookOffset then
			if cols = 0 then%>
				<tr>
			<%end if%>
					<td><a href="bookDetails.asp?productid=<%=rs.Fields.Item("BookID")%>"><img width="160" height="150" src="<%=rs.Fields.Item("BookCover")%>" alt="<%=rs.Fields.Item("BookTitle")%>"/></a></td>
					<td>
						<table>
							<tr>
								<th><%=rs.Fields.Item("BookTitle")%></th>
							</tr>
							<tr>
								<td><strong>Author: </strong> <%=rs.Fields.Item("AuthorSurname")%>, <%=rs.Fields.Item("AuthorFirstNameInitial")%>.</td>
							</tr>
							<tr>
								<td><strong>Category: </strong> <%=rs.Fields.Item("BookCategory")%></td>
							</tr>
							<tr>
								<td><strong>Year: </strong> <%=rs.Fields.Item("BookYear")%></td>
							</tr>
							<tr>
								<td><strong>Cost: </strong> <%=FormatCurrency(rs.Fields.Item("BookPrice"))%></td>
							</tr>
							<tr>
								<td><strong>ISBN: </strong> <%=rs.Fields.Item("BookISBN")%></td>
							</tr>
							<tr>
								<td><a href="addToCart.asp?productid=<%=rs.Fields.Item("BookID")%>">Add to cart</a></td>
							</tr>
						</table>
					</td>
					<%cols = cols + 1
					%>
				<%if cols = 2 then%>
				</tr>
				<%cols = 0
				end if
			
			end if
			books = books + 1
			rs.MoveNext
			loop%>
			<tr>
				<td>
				
				<%if IsNumeric(Request.QueryString("page")) then
					if Request.QueryString("page") > 1 then%>
					<a href='?query=<%=title%>&page=<%=currentPage - 1%>'>< Back</a>
				<%
					end if
					while remainingBooks > 0
						remainingBooks = remainingBooks - 4
						if i = CInt(currentPage) then
							pageNumbers = pageNumbers & i & " "
						else
							pageNumbers = pageNumbers & "<a href='?query=" & title & "&page=" & i & "'>" & i & "</a> "
						end if
						i = i + 1
					wend
					%>
					<%=pageNumbers%>
					<%
					if remainingBooksDisplay > 0 then%>
					<a href='?query=<%=title%>&page=<%=currentPage + 1%>'>Next ></a>
					<%end if
				end if%>
					
				</td>
			</tr>
		</table>
<%
	rs.close
	set rs = nothing
	set cmd = nothing
end sub

Sub AddToCart(id)
	'add item to cart
	set cmd = Server.CreateObject("ADODB.Command")
	set cmd.ActiveConnection = Application("conn")

	cmd.CommandText = "SELECT [BookID] FROM [Books] WHERE [BookID] = ?"
	cmd.CommandType = 1

	cmd.Parameters.Append cmd.CreateParameter("id", 3, 1, 255)
	cmd.Parameters("id")=productId
	
	set rs = cmd.Execute

	if rs.BOF and rs.EOF then
		%>
		<p>Book not found</p>
		<%
	else
		dim items,i
		items = Session("shoppingCartItems")
		Redim Preserve items(ubound(items) + 1)
		items(ubound(items)) = rs.Fields.Item("BookID")
		items = sort(items)
		Session("shoppingCartItems") = items
		
	end if
	rs.close
	set rs = nothing
	set cmd = nothing
End Sub

%>
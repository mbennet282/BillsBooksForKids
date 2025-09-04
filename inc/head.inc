<!DOCTYPE html>
<html lang="en">
	<head>
		<title>Bills Books for Kids - <%=title%></title>
		<meta charset="utf-8"/>
		<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
		<link rel="stylesheet" href="styles.css"/>
	</head>
	<body>
		<div id="container">
			<div id="header">
				<a href="default.asp" id="headerLogo"><h2>Bill's books for kids</h2></a> <br/>
				<div id="search">
					<form method="get" action="search.asp">
						Search book title <input type="text" name="query" id="searchBox"/>
						<input type="submit" id="btnSubmit" value="Go"/>
					</form>
				</div>
				<a href="shoppingcart.asp">View shopping cart</a>
				<div id="links">
					<a href="default.asp"><div class="link">Home</div></a>
					<a href="books.asp"><div class="link">Books</div></a>
					<a href="contact.asp"><div class="link">Contact</div></a>
					<a href="about.asp"><div class="link">About</div></a>
				</div>
				
			</div>
		
			<div id="maincontent">
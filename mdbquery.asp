<%

Option Explicit
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Provider = "Microsoft.Jet.OLEDB.4.0"
conn.open Server.MapPath("/bbk.mdb")

' create a table in the database
'conn.execute "CREATE TABLE Books (BookID AUTOINCREMENT PRIMARY KEY, BookTitle varchar(255) NOT NULL, BookCategory varchar(255) NOT NULL, BookISBN varchar(255) NOT NULL UNIQUE, AuthorSurname varchar(255) NOT NULL, AuthorFirstNameInitial char(1), BookPrice money NOT NULL, BookYear int NOT NULL, BookCover varchar(255))"

' drop (delete) a table in the database
'conn.execute "DROP TABLE Books"

' add a new book
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Amber Spyglass','Fantasy','0-590-54244-3','Pullman','P',14.99,2000,'/Images/BookCovers/Tas_pb_uk.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('Skellig','Fantasy','0-340-71600-2','Almond','D',12.99,1998,'/Images/BookCovers/Skellig_cover.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Gruffalo','Fantasy','0-333-71093-2','Donaldson','J',6.99,1999,'/Images/BookCovers/Fairuse_Gruffalo.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Very Hungry Caterpillar','Literature','0-399-22690-7','Carle','E',6.99,1969,'/Images/BookCovers/HungryCaterpillar.JPG')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Cat in the Hat','Literature','978-0-7172-6059-1','Seuss','D',9.99,1957,'/Images/BookCovers/The_Cat_in_the_Hat.png')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Snowman','Picture books','0241100046','Briggs','R',10.99,1978,'/Images/BookCovers/Raymond_Briggs_Snowman.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The BFG','Fantasy','0-224-02040-4','Dahl','R',15.99,1982,'/Images/BookCovers/bfg.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('War Horse','War','978-0-439-79664-4','Morpurgo','M',12.99,1982,'/Images/BookCovers/War_Horse.jpg')"
'conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('We re Going on a Bear Hunt','Literature','0689504764','Rosen','M',8.99,1989,'/Images/BookCovers/Were_Going_on_a_Bear_Hunt.jpg')"
conn.execute "INSERT INTO Books (BookTitle, BookCategory, BookISBN, AuthorSurname, AuthorFirstNameInitial, BookPrice, BookYear, BookCover) VALUES ('The Sheep-Pig','Picture books','0575033754','King-Smith','D',11.99,1983,'/Images/BookCovers/The_Sheep-Pig.jpg')"


' update book record
'conn.execute "UPDATE Books SET BookCover = '" + Server.MapPath("/Images/BookCovers/Tas_pb_uk.jpg") + "' WHERE BookID = 1"

conn.close
Set conn = nothing

%>

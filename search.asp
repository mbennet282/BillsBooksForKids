<%Option Explicit%>
<!--#include virtual="/inc/vars.inc"-->
<!--#include virtual="/inc/functions.inc"-->
<!--#include virtual="/inc/subs.inc"-->
<%

title = "Book search"


%>
<!--#include virtual="/inc/head.inc"-->
<%queryBooks(request.querystring("query"))%>
<!--#include virtual="/inc/foot.inc"-->
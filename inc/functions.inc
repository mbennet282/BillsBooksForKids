<%
function sort(arr)
	'https://www.experts-exchange.com/questions/20871620/Sorting-a-VBScript-Array.html
	'sort an array in ascending order
	dim i, j, temp
	for i = ubound(arr) - 1 to 0 step -1
		for j = 0 to i
			if arr(j) > arr(j + 1) then
				temp=arr(j + 1)
				arr(j + 1)=arr(j)
				arr(j)=temp
			end if
		next
	next
	sort = arr
end function
%>
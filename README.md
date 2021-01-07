# ClassicASP-Master-Library
For a better world...

# How To Use
## Array Sorting (array-sorting.asp)
Uses the System.Collections.ArrayList object to sort the array.
[Demo](https://aspmasterlibrary.adjans.com.tr/array-sorting.asp)

```asp
<%
MyArray = Array(1,5,9,7,3,2)

tmp_data = MyArray
Response.Write "<h4>Default Array</h4>"
Response.Write Join(tmp_data)
Response.Write "<hr>"

tmp_data = SortArray(MyArray, "ASC")
Response.Write "<h4>Sorted Array (ASC)</h4>"
Response.Write Join(tmp_data)
Response.Write "<hr>"

tmp_data = SortArray(MyArray, "DESC")
Response.Write "<h4>Sorted Array (ASC)</h4>"
Response.Write Join(tmp_data)
Response.Write "<hr>"
%>
```




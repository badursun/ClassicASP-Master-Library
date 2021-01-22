# ClassicASP-Master-Library
For a better world...

# List Of Library
To be prepared

# How To Use
### Array Sorting [array-sorting.asp](/blob/main/array-sorting.asp)
Uses the System.Collections.ArrayList object to sort the array.
[Demo](https://aspmasterlibrary.adjans.com.tr/array-sorting.asp)
* Check this repo for ready Class : [Sorting Class for ClassicASP](https://github.com/badursun/Sorting-Scripting-Dictionary-Classic-ASP)

#### Code Usage
```asp
<%
MyArray = Array(1,5,9,7,3,2)

tmp_data = MyArray
Response.Write "<h4>Default Array</h4>"
Response.Write Join(tmp_data) & "<hr>" 
'OUTPUT: 1 5 9 7 3 2

tmp_data = SortArray(MyArray, "ASC")
Response.Write "<h4>Sorted Array (ASC)</h4>"
Response.Write Join(tmp_data) & "<hr>" 
'OUTPUT: 1 2 3 5 7 9

tmp_data = SortArray(MyArray, "DESC")
Response.Write "<h4>Sorted Array (ASC)</h4>"
Response.Write Join(tmp_data) & "<hr>"
'OUTPUT: 9 7 5 3 2 1
%>
```



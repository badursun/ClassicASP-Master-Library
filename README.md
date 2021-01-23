# ClassicASP-Master-Library
> For a better world...

## List Of Library
> Main headings and details you will find in this repo. Some code samples and comments may not be complete. Check the checkboxes on the side.

- [x] Array Sorting
- [x] Object and Property Exist Control


## How To Use
### (1) Array Sorting [array-sorting.asp](array-sorting.asp)
Uses the System.Collections.ArrayList object to sort the array.
[Demo](https://aspmasterlibrary.adjans.com.tr/array-sorting.asp)
Check this repo for ready Class : [Sorting Class for ClassicASP](https://github.com/badursun/Sorting-Scripting-Dictionary-Classic-ASP)
<details>
<summary>
<a class="btnfire small stroke"><em class="fas fa-chevron-circle-down"></em>&nbsp;&nbsp;Show code usage</a> 
</summary>

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
</details>

### (2) Object and Property Exist Control [object-exist-checker.asp](object-exist-checker.asp)
	Uses Native Javascript runat Server method and return object exist. If object exist, return true, else return false value.
	[Demo](https://aspmasterlibrary.adjans.com.tr/object-exist-checker.asp)
<details>
<summary>
<a class="btnfire small stroke"><em class="fas fa-chevron-circle-down"></em>&nbsp;&nbsp;Show code usage</a> 
</summary>

```asp
<%
Set Checker = New CheckerClass
' Not Exist Class: SomeClass
'------------------------------------
Set objClass = Eval("New SomeClass")
If Checker.ClassExist(objClass) = True Then
	Response.Write "<span style=""color:green"">Exist</span><br>"
Else
	Response.Write "<span style=""color:red"">Not Exist</span><br>"
End If
Set objClass = Nothing

' Not Exist Property in Class: SomeClass.SomeProperty
'------------------------------------
Set objClass = Eval("New SomeClass")
If Checker.ObjectExist(objClass, "SomeProperty") = True Then
	Response.Write "<span style=""color:green"">Exist</span><br>"
Else
	Response.Write "<span style=""color:red"">Not Exist</span><br>"
End If
Set objClass = Nothing
Set Checker = Nothing
%>
```
</details>





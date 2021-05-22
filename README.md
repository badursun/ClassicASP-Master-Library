# ClassicASP-Master-Library
> For a better world...

## List Of Library
> Main headings and details you will find in this repo. Some code samples and comments may not be complete. Check the checkboxes on the side.

- [x] Array Sorting (CF)
- [x] Object and Property Exist Control (CL)
- [x] Do While Usage With MySQL (SC)
- [x] MySQL Connection On ClassicASP (SC)
- [x] Determine date is Weekend or Weekday (CF)

----------------
PS: CF=Custom Function F=Function CL=Class SC=Sample Code
----------------

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

### (3) Do While Usage With MySQL [do-while-with-table.asp](do-while-with-table.asp)
	It shows making a MySQL database connection and spreading a table to the screen with a loop.
	Demo Not found!
<details>
<summary>
<a class="btnfire small stroke"><em class="fas fa-chevron-circle-down"></em>&nbsp;&nbsp;Show code usage</a> 
</summary>

```asp
<%
Set rsObj = Conn.Execute("SELECT * FROM tbl_name ORDER BY col_name ASC")
If rsObj.Eof Then
	' No Record Found in tbl_name
Else
    Do While Not rsObj.Eof
    	' Your Looping Code
    rsObj.MoveNext : Loop
End If
rsObj.Close : Set rsObj = Nothing
%>
```
</details>

### (4) MySQL Connection On ClassicASP [mysql-connection.asp](mysql-connection.asp)
	How do I make a connection to a MySQL database using ASP?
<details>
<summary>
<a class="btnfire small stroke"><em class="fas fa-chevron-circle-down"></em>&nbsp;&nbsp;Show code usage</a> 
</summary>

```asp
<%
Set rsObj = Conn.Execute("SELECT * FROM tbl_name ORDER BY col_name ASC")
If rsObj.Eof Then
	' No Record Found in tbl_name
Else
    Do While Not rsObj.Eof
    	' Your Looping Code
    rsObj.MoveNext : Loop
End If
rsObj.Close : Set rsObj = Nothing
%>
```
</details>

### (5) Determine WeekEnd or WeekDay with Classic ASP Function [is-weekend-function.asp](is-weekend-function.asp)
	How do I determine the date is weekend or weekday?
<details>
<summary>
<a class="btnfire small stroke"><em class="fas fa-chevron-circle-down"></em>&nbsp;&nbsp;Show code usage</a> 
</summary>

```asp
<%
IsWeekend(Date()) ' return true Or false

If IsWeekend(Date()) = True Then 
	Response.Write "Yes, It's weekend"
Else
	Response.Write "No, It's weekday"
End If
%>
```
</details>









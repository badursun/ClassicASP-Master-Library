```asp
<%
	' https://www.dotnetperls.com/arraylist

    If (typeName(vArr) <> "Variant()" OR UBound(vArr) = 0) Then 
    	Exit Function
    End If
    If vSort = "" Then vSort = "ASC"

	Set outputLines = CreateObject("System.Collections.ArrayList")
		For iArr = 0 To UBound(vArr)
			outputLines.Add vArr(iArr)
		Next
	
		outputLines.Sort()
		
		Select Case vSort
			Case "DESC" : outputLines.Reverse()
			Case Else 
		End Select
		
		' SortArray = outputLines ' List Çıktı
		' SortArray = outputLines.ToString ' String Çıktı
		SortArray = outputLines.ToArray ' Array Çıktı

	Set outputLines = Nothing
End Function

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


' MyArray = Array(1,5,9,7,3,2)
' Set outputLines = CreateObject("System.Collections.ArrayList")
' 	outputLines.Add 5
' 	outputLines.Add 3
' 	outputLines.Add 7
' 	outputLines.Add 10
' 	outputLines.Sort()
' 	' outputLines.Reverse()
' 	For Each outputLine in outputLines
' 	    Response.Write outputLine & "<br>"
' 	Next
' Set outputLines = Nothing
%>
```

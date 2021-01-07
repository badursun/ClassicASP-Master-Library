<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
'**********************************************

Function SortArray(vArr, vSort)
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
		
		SortArray = outputLines.ToArray ' Array Çıktı

	Set outputLines = Nothing
End Function


'**********************************************
' Demo
'**********************************************
Dim MyArray
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
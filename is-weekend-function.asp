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

Function IsWeekend(SomeDate)
    Select Case Weekday(SomeDate)
        Case 1: IsWeekend = True  ' Pazar / Sunday
        Case 2: IsWeekend = False ' Pazartesi / Monday
        Case 3: IsWeekend = False ' Salı / Tuesday
        Case 4: IsWeekend = False ' Çarşamba / Wednesday
        Case 5: IsWeekend = False ' Perşembe / Thursday
        Case 6: IsWeekend = False ' Cuma / Friday
        Case 7: IsWeekend = True  ' Cumartesi / Saturday
    End Select
End Function


'**********************************************
' Demo
'**********************************************
Response.Write "<h4>The Date</h4>"
Response.Write Date() & " - " & WeekdayName( Weekday(Date()) )
Response.Write "<hr>"

Response.Write "<h4>Week Of Day</h4>"
Response.Write "Weekday( Date() ) = " & Weekday(Date())
Response.Write "<hr>"

Response.Write "<h4>Is Weekend ?</h4>"
Response.Write "IsWeekend( Date() ) = " & IsWeekend(Date()) 
Response.Write "<hr>"
%>
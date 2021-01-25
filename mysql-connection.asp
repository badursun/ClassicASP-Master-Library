<%
On Error Resume Next

' Define Your DB Conneciton
Const conn_db_name = "your_db_name"
Const conn_db_user = "your_db_user"
Const conn_db_pass = "your_db_pass"
Const conn_db_addr = "your_db_address" ' Usually localhost but use 127.0.0.1 for performance

' Set Conn object to db
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open = "DRIVER={MySQL ODBC 3.51 Driver};database="& Db &";server="& ServerIP &";uid="& User &";password="& Pass &";"
    Conn.Execute "SET NAMES 'utf8mb4'"
    Conn.Execute "SET CHARACTER SET utf8mb4"
    Conn.Execute "SET COLLATION_CONNECTION = 'utf8mb4_general_ci'"

If Err<> 0 Then 
    Response.Write "Database connection failed. Check your conn_ information"
    Repsonse.End
End If



' Do Some Stuff With Your DB. Can access db object via 'Conn' object



' Release Connection Object
Set Conn = Nothing
%>
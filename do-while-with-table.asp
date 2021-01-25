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
%>

<!-- Bootstrap Responsive Table -->
<div class="table-responsive">
    <table class="table table-striped table-bordered">
        <thead>
            <tr>
                <th>Col Name 1</th>
                <th>Col Name 2</th>
                <th>Col Name 3</th>
            </tr>
        </thead>
        <tbody>
<%
Set rsObj = Conn.Execute("SELECT * FROM tbl_name ORDER BY col_name ASC")
If rsObj.Eof Then
%>
            <tr>
                <td colspan="3" align="center">No Record Found</td>
            </tr>
<%
Else
    Do While Not rsObj.Eof
%>
            <tr>
                <td><%=rsObj("COL_1")%></td>
                <td><%=rsObj("COL_2")%></td>
                <td><%=rsObj("COL_3")%></td>
            </tr>
<%
    rsObj.MoveNext : Loop
End If
rsObj.Close : Set rsObj = Nothing
%>
        </tbody>
    </table>
</div>
<!-- Bootstrap Responsive Table -->

<%
' Release Connection Object
Set Conn = Nothing
%>
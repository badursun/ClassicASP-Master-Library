<%
DATE_NOW 	= Date()
MONTH_NOW 	= Month(DATE_NOW)
YEAR_NOW 	= Year(DATE_NOW)

MONTH_START = DateSerial(YEAR_NOW, MONTH_NOW, 1)
MONTH_END 	= DateSerial(YEAR_NOW, MONTH_NOW + 1, 0)

Response.Write "Month: "& MONTH_NOW &" <br>"
Response.Write "Year: "& YEAR_NOW &" <br>"
Response.Write "Date: "& DATE_NOW &" <br>"

Response.Write "This Month Start: "& MONTH_START &"<br>"
Response.Write "This Month End: "& MONTH_END &"<br>"
%>
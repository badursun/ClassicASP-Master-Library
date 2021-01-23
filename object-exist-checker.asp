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

Class CheckerClass
	private sub class_initialize()%>
		<script language="javascript" runat="server">
		function CheckObj(obj){return (typeof obj != "undefined");}
		function CheckProperty(obj, propName){return (typeof obj[propName] != "undefined");}
		</script>
	<%end sub

	Private sub class_terminate()
	End Sub

	Public Property Get ClassExist(obj)
		ClassExist = CheckObj(obj)
	End Property
	
	Public Property Get ObjectExist(obj, propName)
		ObjectExist = CheckProperty(obj, propName)
	End Property
End Class


'**********************************************
' Demo
'**********************************************
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
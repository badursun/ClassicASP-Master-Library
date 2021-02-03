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
Class RedirectWith301
	Private Parca3, Parca4, Parca5, strQuery, strsplit
	Private OldURL, NewURL, URLsSize, URLs, SendURL, SendMETHOD
	Private Debug, RunAs, ExecutePage, SimulateURL
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	private sub class_initialize()
		On Error Resume Next

		URLs 		= Array()
		URLsSize 	= 0
		Debug 		= False
		ExecutePage = False
		RunAs 		= False

		strQuery = Request.ServerVariables("QUERY_STRING")
		strsplit = Split(strQuery,"/")
		For i = 0 to UBound(strsplit)
			If strsplit(i) = "" Then
				strsplit(i) = ""
			Else
				strsplit(i) = strsplit(i)
			End If
		Next
		If UBound(strsplit) => 3 Then Parca3 = Trim(strsplit(3)) & ""
		If UBound(strsplit) => 4 Then Parca4 = Trim(strsplit(4)) & ""
		If UBound(strsplit) => 5 Then Parca5 = Trim(strsplit(5)) & ""

		OldURL 		= ""
		NewURL 		= ""
		SendURL		= "/index.asp"
		SimulateURL = ""
		SendMETHOD 	= False
	end sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let Execute(vVal)
		ExecutePage = vVal
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let DebugStatus(vVal)
		Debug = vVal
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let IfURL(vVal)
		OldURL = vVal
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Let SendTo(vVal)
		NewURL = vVal
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get AddRule()
	    ReDim PRESERVE URLs(URLsSize)
	    URLsSize=URLsSize+1
	    ReDim PRESERVE URLs( URLsSize )
	    URLs(URLsSize-1) = ""& OldURL &"|"& NewURL &"|"& ExecutePage &""
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function GetDomain()
	    If Request.ServerVariables("SERVER_PORT") = 443 Then 
	        str_http_status = "https://"
	    Else
	        If Request.ServerVariables("HTTP_X-Forwarded-Proto") = "https" Then
	            str_http_status = "https://"
	        Else 
	            str_http_status = "http://"
	        End If
	    End If
	    SiteURL = Request.ServerVariables("SERVER_NAME")
	    SiteURL = Replace(SiteURL ,"https://","",1,-1,1)
	    SiteURL = Replace(SiteURL ,"http://","",1,-1,1)
	    GetDomain = str_http_status & SiteURL
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function GetURL()
	    strSecureURL = Request.QueryString
	    strSecureURL = Replace(strSecureURL, GetDomain() ,"" )  
	    strSecureURL = Replace(strSecureURL, ":8080" ,"" )  
	    strSecureURL = Replace(strSecureURL, ":80" ,"" )  
	    strSecureURL = Replace(strSecureURL, ":443" ,"" )  
	    strSecureURL = Replace(strSecureURL, "404;" ,"" )  
	    GetURL = strSecureURL
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Property Get GoToURL(vVal, vMethod)		
		If Len(vVal)=0 Then 
			vVal = "/index.asp"
		End If

		If Debug = True Then 
			If vMethod = True Then 
				Response.Write "METHOD 		: Execute("& vVal &")" &vbcrlf
				Response.Write "RUN SCRIPT 	: Execute("& SimulateURL &")" &vbcrlf
				Response.Write "[id = '"& Session("TRANSFER_ID") &"']" &vbcrlf
			Else
				Response.Write "METHOD 		: Transfer("& vVal &")"&vbcrlf
			End If
			Response.End
			Exit Property
		Else
			If vMethod = True Then 
				Response.Status = "200 OK"
				' Response.Write SimulateURL & " => id=" & Session("TRANSFER_ID")
				Server.Transfer(SimulateURL)
				' Server.Execute(SimulateURL)
				' Response.Redirect(SimulateURL)
			Else
				Response.Status = "301 Moved Permanently"
				' Response.Redirect vVal
				Response.AddHeader "Location", ""& vVal &""
			End If
		End If
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get RunRedirector()
		If Debug = True Then Response.Write "<pre>"
		If UBound(URLs)<0 Then 
			Response.Write "No Added URL" & vbcrlf
			Exit Property
		End If

		If (Len(Parca3) = 0 OR Parca3 = "index.asp") Then 
			GoToURL "", False
		Else
			If Debug = True Then 
				Response.Write "********** Redirector Debug **********" & vbcrlf
				Response.Write "RequestURL is 		: '"& GetURL() &"'" & vbcrlf
				Response.Write "Total Rule 		: '"& UBound(URLs) &"'" & vbcrlf
				Response.Write "GetDomain() 		: '"& GetDomain() &"'" & vbcrlf
				Response.Write "GetURL() 		: '"& GetURL() &"'"& vbcrlf & vbcrlf & vbcrlf
				
				Response.Write "********** Redirector Rule List **********" & vbcrlf
			End If

			For i=0 To Ubound(URLs)-1
				SimulateURL = ""
				tmp_data 	= Split(URLs(i),"|")
				OLD_URL 	= tmp_data(0)
				NEW_URL 	= tmp_data(1)
				PG_METH 	= CBool(tmp_data(2))
				
				' Response.Write "1 ["& (Trim(NEW_URL)=Trim(GetURL()) AND PG_METH = True) &"]" &vbcrlf
				' Response.Write "2 ["& (Trim(OLD_URL)=Trim(GetURL()) AND PG_METH = False) &"]"&vbcrlf

				If Debug = True Then Response.Write "RULE #"& i &" 	: '"& OLD_URL &"' => '"& NEW_URL &"' "
				' Check Execute
				If ( Trim(NEW_URL)=Trim(GetURL()) ) Then 
					If Debug = True Then Response.Write "<span style=""color:green"">IS EXECUTED</span> 	" & vbcrlf
					
					SendMETHOD 				= PG_METH
					SimulateURL 			= Replace(OLD_URL, "/", "_")
					If Instr(1, OLD_URL, "?") <> 0 Then
						SimulateURL 			= Split(SimulateURL, "?")(0)
					End If
					Session("TRANSFER_ID") 	= URLStringParse(OLD_URL, "id")
					
					If Debug = False Then Exit For
				ElseIf ( Trim(OLD_URL)=Trim(GetURL()) ) Then 
					If Debug = True Then Response.Write "<span style=""color:green"">IS MATCHED</span> 	" & vbcrlf
					
					SendURL 	= NEW_URL
					
					If Debug = False Then Exit For
				Else
					If Debug = True Then Response.Write "<span style=""color:red"">IS NOT MATCHED</span> 	" & vbcrlf
				End If

				If Debug = True Then
					Response.Write "METHOD 		: '"& MethodName(PG_METH) &"' => '"& PG_METH &"'" & vbcrlf
					Response.Write "GetURL 		: '"& GetURL() &"'" & vbcrlf
					Response.Write "NEW_URL 	: '"& NEW_URL &"'"& vbcrlf
					Response.Write "OLD_URL 	: '"& OLD_URL &"'"& vbcrlf
					If PG_METH = True Then
					Response.Write "RUN SCRIPT 	: '"& SimulateURL &"'" &vbcrlf
					End If
				End If

				If Debug = True Then Response.Write vbcrlf 
			Next

			GoToURL SendURL, SendMETHOD
		End If

		If Debug = True Then Response.Write "</pre>"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function MethodName(v)
		If v = "True" Then 
			MethodName = "Execute"
		Else
			MethodName = "Transfer"
		End If
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function URLStringParse(DataVal, Hangisi)
		If DataVal = "" Then 
			DataVal = Request.ServerVariables("QUERY_STRING") & "&s=x"
		End If

		If Instr(1, DataVal, "?") <> 0 Then 
			DataVal = DataVal & "&s=x"
		End If
		Hangisi = Hangisi & "="

		dim sResult 
		dim lStart
		dim lEnd

		lStart = instr( 1, DataVal, Hangisi, 1 )
		if lStart > 0 then 
			lStart = lStart + len(Hangisi)
			lEnd = instr( lStart, DataVal, "&" )
			if lEnd = 0 then lEnd = len( DataVal )
			sResult = mid( DataVal, lStart, lEnd - lStart )
		end if

		URLStringParse = SQLInjectionBlocker(sResult)
	end function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private Function SQLInjectionBlocker(vData)
		vData = Replace(vData, "script", "&#115;cript", 1, -1, 0)
		vData = Replace(vData, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
		vData = Replace(vData, "Script", "&#083;cript", 1, -1, 0)
		vData = Replace(vData, "script", "&#083;cript", 1, -1, 1)
		vData = Replace(vData, "object", "&#111;bject", 1, -1, 0)
		vData = Replace(vData, "OBJECT", "&#079;BJECT", 1, -1, 0)
		vData = Replace(vData, "Object", "&#079;bject", 1, -1, 0)
		vData = Replace(vData, "object", "&#079;bject", 1, -1, 1)
		vData = Replace(vData, "document", "&#100;ocument", 1, -1, 0)
		vData = Replace(vData, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
		vData = Replace(vData, "Document", "&#068;ocument", 1, -1, 0)
		vData = Replace(vData, "document", "&#068;ocument", 1, -1, 1)
		vData = Replace(vData, "cookie", "&#099;ookie", 1, -1, 0)
		vData = Replace(vData, "COOKIE", "&#067;OOKIE", 1, -1, 0)
		vData = Replace(vData, "Cookie", "&#067;ookie", 1, -1, 0)
		vData = Replace(vData, "cookie", "&#067;ookie", 1, -1, 1)
		vData = Replace(vData, "applet", "&#097;pplet", 1, -1, 0)
		vData = Replace(vData, "APPLET", "&#065;PPLET", 1, -1, 0)
		vData = Replace(vData, "Applet", "&#065;pplet", 1, -1, 0)
		vData = Replace(vData, "applet", "&#065;pplet", 1, -1, 1)
		vData = Replace(vData, "UNION", "", 1, -1, 0)
		vData = Replace(vData, "union", "", 1, -1, 0)
		vData = Replace(vData, "Union", "", 1, -1, 0)
		vData = Replace(vData, "document.cookie", "&#068;ocument.cookie", 1, -1, 1)
		vData = Replace(vData, "javascript:", "javascript ", 1, -1, 1)
		vData = Replace(vData, "vbscript:", "vbscript ", 1, -1, 1)
		vData = Replace(vData, "'", "&apos;")
		vData = Replace(vData, chr(39), "&apos;")
		vData = Replace(vData, "%20", " ")
		SQLInjectionBlocker = vData
	End Function
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
End Class 
%>
<!--#include file="_redirector_class.asp"--><%
Set Redirector = New RedirectWith301
	With Redirector
		.DebugStatus= False

		' Rule #1
		'---------------------------
		.IfURL  	= "/hakkimda.html"
		.SendTo 	= "/marcus-greer-fitness.asp"
		.AddRule()

		' Rule #2
		'---------------------------
		.IfURL  	= "/mobil/baylar.asp"
		.SendTo 	= "/uzaktan-pt-paketleri.asp"
		.AddRule()

		' Rule #3 Double Redirect and Execute
		'---------------------------
		.IfURL  	= "/blog-detay.asp?id=15" ' Executed File start trail (blog-detay.asp => _blog-detay.asp )
		.SendTo 	= "/fit-bir-vucut-icin-bilmeniz-gereken-temel-maddeler.asp"
		.Execute 	= True
		.AddRule()
		.IfURL  	= "/blog-Fit-Bir-Vücut-İçin-Bilmemiz-Gereken-Temel-Maddeler!-no15.html"
		.SendTo 	= "/blog-detay.asp?id=15"
		.Execute 	= False
		.AddRule()

		' Rule #3
		'---------------------------
		.IfURL  	= "/blog-detay.asp?id=16"
		.SendTo 	= "/diyet-pankek-tarifi.asp"
		.Execute 	= True
		.AddRule()

		.RunRedirector()
	End with
Set Redirector = Nothing
%>
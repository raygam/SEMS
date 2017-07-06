<html>
	<head>
		<title>SEMS</title>
		<link href="..\css\Cortex.css" rel="stylesheet" type="text/css" />
		<link href="..\css\CortexDashboard.css" rel="stylesheet" type="text/css" />
		<script language="javascript"> 
			
			
			
			function toggle_it(itemID){ 
     
				if ((document.getElementById(itemID).style.display == 'none')) { 
					document.getElementById(itemID).style.display = 'block' 
					
				 } else { 
					document.getElementById(itemID).style.display = 'none'; 
					
				}    
			} 
			
		</script>

	</head>
	
	<body>


		<a href="../studies.asp"><img border="0" src="..\Images\SemsLogo.png"></a> &nbsp;&nbsp;<a href="javascript:window.print()"><font size=2>Print</font></a><br>
<%
Dim dtStart
dtStart = now()

dim db, dbSL, rsUsers, rs, SecCodes, i, CurrentPage, ShowReports
dim SLXDocumentLocation, uploadsdirvar
dim strCompanyLinks, CurrentCompany, strCurrency, strCurrencyCode, CompanyURLExtra
CurrentPage = Request.ServerVariables("SCRIPT_NAME") 

set db = server.CreateObject("ADODB.Connection")
set dbSL = server.CreateObject("ADODB.Connection")
set rsUsers = server.CreateObject("ADODB.Recordset")
set rs = server.CreateObject("ADODB.Recordset")
set rs2 = server.CreateObject("ADODB.Recordset")
set rs3 = server.CreateObject("ADODB.Recordset")
dim SLXSQLConnectionString, SLXProviderConnectionString, SalesLogixOppURL


'SLXSQLConnectionString = "Provider=SQLNCLI10;SERVER=SHWTetra;DATABASE=SMVEUCortex;UID=CortexApp;PWD=(#GetiN#)"
'SLXSQLConnectionString = "Provider=SQLNCLI10;SERVER=SHWTetra;DATABASE=SEMS;UID=CortexApp;PWD=(#GetiN#)"
SLXSQLConnectionString = "Provider=SQLNCLI10;SERVER=COMPUTER894\SQL2014;DATABASE=SEMS;UID=SEMS;PWD=SEMS"
'SLXSQLConnectionString = "Provider=SQLNCLI10;SERVER=eushwvdb1;DATABASE=SEMS;UID=SEMs_User;PWD=SQLdB82652#"


db.open SLXSQLConnectionString

if trim(session("LoggedIn")) = "" then
	dim DomainUser, DU
	DomainUser = Request.ServerVariables("AUTH_USER")
		DomainUser = "SSI\RGAMBARDELLA"

	du = Split(DomainUser,"\")
	
	if ucase(du(0))="SSI" then
		
		if ucase(du(1))="HDUNN" or ucase(du(1))="DATACRAFTADMIN" then
			'du(1)="CCooke"
			'du(1)="AFournier"
			'du(1)= "ENfon"
			'du(1)="KCocks"
			du(1)="FDavies"
			du(1)="DFairhurst"
			'du(1)="DO'Kelly"
		end if
		du(1)=replace(du(1),"'","''")
		
		rsUsers.open "SELECT * FROM Users WHERE IsActive = -1 AND Username = '" & du(1) & "'",db
		
		if rsUsers.eof=false or rsUsers.bof=false then
			'ACCOUNT EXISTS IN CORTEX
			
			session("UserWebKey")=rsUsers("WebKey")
			session("LoggedIn")=rsUsers("UserID")
			session("UserCode")=rsUsers("Username")
			Response.Cookies("UserCode")=rsUsers("Username")
			
			session("Fullname") = rsUsers("Firstname") & " " & rsUsers("Surname")
			session("SystemAdmin") = rsUsers("SystemAdmin")
			session("Archivist") = rsUsers("Archivist")
			session("DepartmentManager") = rsUsers("DepartmentManager")
			session("SponsorProjectManager") = rsUsers("SponsorProjectManager")
			session("BD") = rsUsers("BD")
			session("TFM") = rsUsers("TFM")
			session("LoggedInName") = rsUsers("Firstname") & " " & rsUsers("Surname")
			session("Department") = rsUsers("OADepartment") & ""
			session("DMDepartments") = rsUsers("DMDepartments") & ""
			Session("QA") = rsUsers("QA") & ""

			'MAKE ENTRY IN EM_AUDIT TABLE?
			
		else
			session("CRMLogin") = ""
			session("LoggedIn") = ""
			session("UserCode") = ""
			session("SystemAdmin") = ""
			session("Archivist") = ""
			session("StudyDirector") = ""
			session("QA") = ""
			
			session("ProjectManager") = ""
			session("SponsorProjectManager") = ""
			session("DepartmentManager") = ""
			session("BD") = ""
			session("TFM") = ""
			session("LoggedInName") = ""
			session("Department") = ""
			session("DMDepartments") = ""
			
		end if
		rsUsers.close
	else
		session("LoggedIn") = ""
	end if
end if




if session("LoggedIn")>"" then
	
	
else
	'REDIRECT TO LOGIN FAIL PAGE
	Response.Write "<br><blockquote>The user <b>" & du(0) & "\" & du(1) & "</b> has not been granted access to Cortex.</blockquote>"
	Response.End
	
end if


Response.Write "<table cellspacing=0 cellpadding=3 width=""100%"" bgcolor=""#59545c""><tr><td align=right><font color=white>Logged In: <b>" & session("LoggedInName") & "</b></td>"
Response.Write "</tr></table>"

Response.Write "<div id=navbar>"

Response.Write "<a href=""..\Studies.asp"">Studies</a>"
Response.Write "<a href=""..\Reports.asp"">Reports</a>"
if trim(session("SystemAdmin")) = "-1" then
	Response.Write "<a href=""..\users.asp"">Users</a>"
end if 

if trim(session("UserCode"))="FDavies" or lcase(trim(session("UserCode")))="jsumner" then
	Response.Write "<a href=""..\SetTriggers.asp"">Triggers</a>"
	
end if
if trim(session("UserCode"))="FDavies" then
	
	Response.Write "<a href=""..\ClientList.asp"">Client List</a>"
	Response.Write "<a href=""..\TestSubstanceList.asp"">Test Substance List</a>"
	
end if

if trim(session("SystemAdmin")) = "-1" then
	'Response.Write "<a href=""..\audit.asp"">Audit</a>"
end if

Response.Write "</div>"


Response.Write "<br/>"
%>
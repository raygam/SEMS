<%@ Language=VBScript %>
<%

%>
<!-- #include file = "includes\CortexConfig.asp"-->

<%

if cstr(session("SystemAdmin")) <> "-1" then
	Response.Redirect "Studies.asp"
end if


if trim(Request.QueryString("Key"))="" then

	path = left( monthname(month(now())) ,3) & year(now())
	
	WebKey = Request.Form("WebKey")
	Firstname = Request.Form("Firstname")
	Surname = Request.Form("Surname")
	Username = Request.Form("Username")
	Department = Request.Form("Department")
	SystemAdmin = Request.Form("SystemAdmin")
	DeopartmentManager = Request.Form("DepartmentManager")
	SponsorProjectManager = Request.Form("SponsorProjectManager")
	TFM = Request.Form("TFM")
	QA = Request.Form("QA")
	
	if SystemAdmin = "" then
		SystemAdmin = "0"
	end if
	Archivist = Request.Form("Archivist")
	if Archivist = "" then
		Archivist = "0"
	end if
	
	if QA = "" then
		QA = "0"
	end if
	StudyDirector = Request.Form("StudyDirector")
	if StudyDirector = "" then
		StudyDirector = "0"
	end if
	
	TFM = Request.Form("TFM")
	if TFM = "" then
		TFM = "0"
	end if
	
	DepartmentManager = Request.Form("DepartmentManager")
	if DepartmentManager = "" then
		DepartmentManager = "0"
	end if
	
	SponsorProjectManager = Request.Form("SponsorProjectManager")
	if SponsorProjectManager = "" then
		SponsorProjectManager = "0"
	end if
	
	BD = Request.Form("BD")
	if BD = "" then
		BD = "0"
	end if
	
	IsActive = Request.Form("IsActive")
	if IsActive = "" then
		IsActive = "0"
	end if
	
	ClientEdit = Request.Form("ClientEdit")
	Email = Request.Form("Email")
	b1 = Request.Form("b1")
	

	if b1 ="Cancel" then
		Response.Redirect "Users.asp"
	end if
	
		
	msg = ""
	if ClientEdit="" then
		UserID = Session("LoggedIn")
	end if
	
	if Firstname = "" or Surname="" then
		msg = "Please enter a Firstname and Surname. "
	end if
	if Department = "" then
		msg = msg & "Please enter a Department. "
	end if
	if Username = "" then 
		msg = msg & "Please enter a Username. "
	end if
	if msg = "" then
		'OK to Save
		strsql = "SELECT * FROM Users WHERE WebKey = '" & WebKey & "'"
		rs.open strsql,db,1,3
		if WebKey = "" then
			rs.addnew
			WebKey = CreateWebKey()
			rs("WebKey")= WebKey
			rs.update
			rs.close
			strsql = "SELECT * FROM Users WHERE WebKey = '" & WebKey & "'"
			rs.open strsql,db,1,3
			
		end if
		
		
		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Firstname", rs("Firstname") & "", Firstname 
		rs("Firstname") = Firstname
		
		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Surname", rs("Surname") & "", Surname 
		rs("Surname") =Surname
		
		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Username", rs("Username") & "", Username 
		rs("Username") = Username
		
		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "OADepartment", rs("Department") & "", Department 
		rs("OADepartment") = Department
		
		WriteToChangeLog "ChangeLogLarge", rs("UserID"), Session("LoggedIn"), "Users", "Email", rs("Email") & "", Email
		rs("Email") = Email
		
		
		
		
		
		
		
		if Session("SystemAdmin")="-1" then
			'if trim(StudyDirector) = "0" then
			'		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "StudyDirector", rs("StudyDirector") & "", "0"
			'		rs("StudyDirector")=0	
			'else
			'		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "StudyDirector", rs("StudyDirector") & "", "-1"
			'		rs("StudyDirector") = -1	
			'end if
		
			'if trim(QA) = "0" then
			'		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "0"
			'		rs("QA")=0	
			'else
			'		WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "-1"
			'		rs("QA") = -1	
			'end if
		
			if trim(Archivist) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Archivist", rs("Archivist") & "", "0"
					rs("Archivist")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Archivist", rs("Archivist") & "", "-1"
					rs("Archivist") = -1	
			end if
		
			if trim(TFM) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "TFM", rs("TFM") & "", "0"
					rs("TFM")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "TFM", rs("TFM") & "", "-1"
					rs("TFM") = -1	
			end if
		
			
			
			if trim(SystemAdmin) = "0" then
			
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SystemAdmin", rs("SystemAdmin") & "", "0"
					rs("SystemAdmin")=0	
					'Session("SystemAdmin")="0"
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SystemAdmin", rs("SystemAdmin") & "", "-1"
					rs("SystemAdmin") = -1	
			end if
			
			if trim(IsActive) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "IsActive", rs("IsActive") & "", "0"
					rs("IsActive")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "IsActive", rs("IsActive") & "", "-1"
					rs("IsActive") = -1	
			end if
		
			if trim(DepartmentManager) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "DepartmentManager", rs("DepartmentManager") & "", "0"
					rs("DepartmentManager")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "DepartmentManager", rs("DepartmentManager") & "", "-1"
					rs("DepartmentManager") = -1	
			end if
			
			if trim(SponsorProjectManager) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SponsorProjectManager", rs("SponsorProjectManager") & "", "0"
					rs("SponsorProjectManager")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SponsorProjectManager", rs("SponsorProjectManager") & "", "-1"
					rs("SponsorProjectManager") = -1	
			end if
			
			if trim(BD) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "BD", rs("BD") & "", "0"
					rs("BD")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "BD", rs("BD") & "", "-1"
					rs("BD") = -1	
			end if
		
			if trim(QA) = "0" then
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "0"
					rs("QA")=0	
			else
					WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "-1"
					rs("QA") = -1	
			end if
		
		end if
		
		
		
		
		
		
		
		rs("LastUpdated")=now()
		rs("LastUpdatedBy") = Session("LoggedIn")
		
		rs.update
		rs.close
		
		
		Response.Redirect "Users.asp"
	else
		'Response.Write "Message: " & msg
		'Error - back to form
	end if
else
	WebKey = replace(Request.querystring("Key"),"'","")
	strsql = "SELECT * FROM Users WHERE WebKey = '" & WebKey & "'"
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		UserID = rs("UserID")
		Firstname = rs("Firstname") & ""
		Surname = rs("Surname") & ""
		Department = rs("OADepartment") & ""
		Username = rs("Username") & ""
		Email = rs("Email") & ""
		SystemAdmin = rs("SystemAdmin") & ""
		Archivist = rs("Archivist") & ""
		TFM = rs("TFM") & ""
		
		
		'StudyDirector = rs("StudyDirector") & ""
		IsActive = rs("IsActive") & ""
		SponsorProjectManager = rs("SponsorProjectManager") & ""
		DepartmentManager = rs("DepartmentManager") & ""
		BD = rs("BD") & ""
		QA = rs("QA") & ""
		
		
	else
		Response.Redirect "Users.asp"
	end if
	rs.close
end if



UserTitle = FirstName & " " & Surname

ClientEdit = trim(ClientEdit)
Response.Write "<blockquote><h1>Edit "

If WebKey > "" then
	Response.Write UserTitle 
else
	Response.Write "New User"
end if
Response.Write "</h1>"
if msg>"" and ClientEdit > "" then
	Response.Write "<font color=red><b>" & msg & "</b></font>"
end if

Response.Write "<form method=""Post"" action=""EditUsers.asp"">"

Response.Write "<table align=""center"" cellpadding=10 border=0><tr><td valign=top>"

Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""3"">"


Response.Write "<tr><td valign=top>Firstname<font color=red>*</font></td>"
Response.Write "<td valign=top>" & TextBoxHTML("Firstname",Firstname,25,30,"") & "</td></tr>"

Response.Write "<tr><td valign=top>Surname<font color=red>*</font></td>"
Response.Write "<td valign=top>" & TextBoxHTML("Surname",Surname,25,30,"") & "</td></tr>"

Response.Write "<tr><td valign=top>Windows Username<font color=red>*</font></td>"
Response.Write "<td valign=top>" & TextBoxHTML("Username",Username,25,30,"") & "</td></tr>"

Response.Write "<tr><td valign=top>Department</td>"
Response.Write "<td valign=top>" & DataComboHTML("Department", Department, "SELECT LookupValue, LookupValue as D1 FROM LookupValues WHERE LookupID = 20 ORDER BY LookupValue ASC" , "", "") & "</td></tr>"

Response.Write "<tr><td valign=top>Email</td>"
Response.Write "<td valign=top>" & TextBoxHTML("Email",Email,40,100,"") & "</td></tr>"

Response.Write "</table>"

Response.Write "</td><td valign=""top"">"

Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""4"">"



Response.Write "<tr><td valign=top>System Admin</td>"
'Response.Write UserID & "/" & Session("LoggedIn")

if trim(UserID)=trim(Session("LoggedIn")) then
	Locked = "Yes"
else
	Locked = ""
end if
'Response.Write UserID & "/" & Session("LoggedIn") & "/" & Locked & "/" & session("SystemAdmin") & "/" & SystemAdmin
'if trim(Session("SystemAdmin")) = "-1" then
	Response.Write "<td valign=top>" & CheckBoxHTML("SystemAdmin",SystemAdmin,"") & "</td></tr>"
'else
'	Response.Write "<td valign=top>" & CheckBoxHTML("SystemAdmin",SystemAdmin,"Yes") & "</td></tr>"	
'end if

Response.Write "<tr><td valign=top>Archivist</td>"
'if Session("Archivist") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("Archivist",Archivist,"Yes") & "</td></tr>"
'else
	Response.Write "<td valign=top>" & CheckBoxHTML("Archivist",Archivist,"") & "</td></tr>"	
'end if

Response.Write "<tr><td valign=top>Department Manager</td>"
Response.Write "<td valign=top>" & CheckBoxHTML("DepartmentManager",DepartmentManager,"") & "</td></tr>"

Response.Write "<tr><td valign=top>Sponsor Project Manager</td>"
Response.Write "<td valign=top>" & CheckBoxHTML("SponsorProjectManager",SponsorProjectManager,"") & "</td></tr>"

Response.Write "<tr><td valign=top>BD</td>"
Response.Write "<td valign=top>" & CheckBoxHTML("BD",BD,"") & "</td></tr>"

Response.Write "<tr><td valign=top>TFM</td>"
Response.Write "<td valign=top>" & CheckBoxHTML("TFM",TFM,"") & "</td></tr>"

Response.Write "<tr><td valign=top>QA</td>"
Response.Write "<td valign=top>" & CheckBoxHTML("QA",QA,"") & "</td></tr>"



'Response.Write "<tr><td valign=top>QA</td>"
'if Session("QA") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("QA",QA,"Yes") & "</td></tr>"
'else
	'Response.Write "<td valign=top>" & CheckBoxHTML("QA",QA,"") & "</td></tr>"	
'end if

'Response.Write "<tr><td valign=top>Study Director</td>"
'if Session("StudyDirector") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("StudyDirector",StudyDirector,"Yes") & "</td></tr>"
'else
	'Response.Write "<td valign=top>" & CheckBoxHTML("StudyDirector",StudyDirector,"") & "</td></tr>"	
'end if

Response.Write "<tr><td valign=top>Is Active</td>"
'if trim(Session("SystemAdmin")) = "-1" then
	Response.Write "<td valign=top>" & CheckBoxHTML("IsActive",IsActive,"") & "</td></tr>"
'else
'	Response.Write "<td valign=top>" & CheckBoxHTML("IsActive",IsActive,"Yes") & "</td></tr>"	
'end if



Response.Write "</table>"


Response.Write "</td></tr>"




Response.Write "<tr><td colspan=""2""><hr></td></tr>"

Response.Write "<tr><td colspan=""2"" align=""center""><input type=""Submit"" name=""b1"" value=""Save"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""Submit"" name=""b1"" value=""Cancel""><input type=""hidden"" name=""ClientEdit"" value=""Yes""><input type=""hidden"" name=""WebKey"" value=""" & WebKey & """></td></tr>"

Response.Write "</table>"
Response.Write "</form>"


%>
<!-- #include file = "includes\footer.asp"-->
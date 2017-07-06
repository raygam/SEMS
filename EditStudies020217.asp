<%@ Language=VBScript %>
<%

%>
<!-- #include file = "includes\CortexConfig.asp"-->

<%


dim FieldData(50), SecurityData(50), FormData(50), FieldLength(50), CanEdit(50),FieldType(50), intFields, StudyID, NewNoteOrSite
set rsTrigger = server.CreateObject("ADODB.Recordset")
Response.Write "<script type='text/JavaScript' src='scw.js'></script>"
NewNoteOrSite = false
intFields = 0
rs.open "SELECT * FROM FieldData",db
do until rs.eof=true
	
	FieldData(rs("FieldID")) = rs("FieldName") & ""
	SecurityData(rs("FieldID")) = trim(rs("FieldSecurity") & "")
	FormData(rs("FieldID")) = Request.Form(FieldData(rs("FieldID")))
	FieldType(rs("FieldID")) = rs("FieldType")
	if isnumeric(rs("FieldLength"))=true then
		FieldLength(rs("FieldID")) = cint(rs("FieldLength"))
	else
		FieldLength(rs("FieldID")) = 0
	end if
	intFields = intFields + 1
	
	
	CanEdit(rs("FieldID"))="N"

	s = split(SecurityData(rs("FieldID")),";")
	for i = 0 to ubound(s)
		wk = trim(replace(Request.Form("WebKey"),"'",""))
		if wk="" then
			wk = trim(replace(Request.Querystring("Key"),"'",""))
		end if
		strSQL = "SELECT * FROM Studies WHERE WebKey = '" & wk & "'"
		rs2.open strsql, db
		if rs2.eof=false or rs2.bof=false then
			strStudyDirector = rs2("StudyDirector") & ""
			strDepartment = rs2("Department") & ""
			
		
		end if
		rs2.close
		
		
		select case s(i)
		case "SD"
			if trim(session("Fullname")) = strStudyDirector then
				CanEdit(rs("FieldID"))="Y"
			end if
		case "DM"
			'if session("ProjectManager") = "-1" and strDepartment = session("Department") then
			'if session("DepartmentManager") = "-1" and strDepartment = session("Department") then
			if session("DepartmentManager") = "-1" and instr(ucase(session("DMDepartments")),ucase(strDepartment))>0 then
			
				CanEdit(rs("FieldID"))="Y"
			end if
		case "SA"
			if session("SystemAdmin") = "-1" then
				CanEdit(rs("FieldID"))="Y"
			end if
		case "AR"
			if session("Archivist") = "-1" then
				CanEdit(rs("FieldID"))="Y"
			end if
		case "SPM"
			if session("SponsorProjectManager") = "-1" then
				CanEdit(rs("FieldID"))="Y"
			end if
		case "BD"
			if session("BD") = "-1" then
				CanEdit(rs("FieldID"))="Y"
			end if
		case "TFM"
			if session("TFM") = "-1" then
				CanEdit(rs("FieldID"))="Y"
			end if	
		end select
	next
	rs.movenext
loop
rs.close
NoteNew = Request.Form("NoteNew")
TestFacilityNew = Request.Form("TestFacilityNew")
TestSiteNew = Request.Form("TestSiteNew")
OffSiteSDNew = Request.Form("OffSiteSDNew")
SmithersSDNew = Request.Form("SmithersSDNew")
SmithersPINew = Request.Form("SmithersPINew")
OffSitePINew = Request.Form("OffSitePINew")



'SECURITY DETAILS
'Response.Write "(Archivist:" & session("Archivist") & ")" 

if trim(Request.QueryString("Key"))=""  then

	path = left( monthname(month(now())) ,3) & year(now())
	
	WebKey = Request.Form("WebKey")
	
	
	'Firstname = Request.Form("Firstname")
	'Surname = Request.Form("Surname")
	'Username = Request.Form("Username")
	'Department = Request.Form("Department")
	'SystemAdmin = Request.Form("SystemAdmin")
	'if SystemAdmin = "" then
	'	SystemAdmin = "0"
	'end if
	'Archivist = Request.Form("Archivist")
	'if Archivist = "" then
	'	Archivist = "0"
	'end if
	'QA = Request.Form("QA")
	'if QA = "" then
	'	QA = "0"
	'end if
	'StudyDirector = Request.Form("StudyDirector")
	'if StudyDirector = "" then
	'	StudyDirector = "0"
	'end if
	'IsActive = Request.Form("IsActive")
	'if IsActive = "" then
	'	IsActive = "0"
	'end if
	ClientEdit = Request.Form("ClientEdit")
	'Email = Request.Form("Email")
	b1 = Request.Form("b1")
	

	if b1 ="Cancel" then
		Response.Redirect "Studies.asp"
	end if
	
	if b1 ="Back" then
		Response.Redirect "Studies.asp"
	end if
		
	msg = ""
	if ClientEdit="" then
		UserID = Session("LoggedIn")
	end if
	
	'if Firstname = "" or Surname="" then
	'	msg = "Please enter a Firstname and Surname. "
	'end if
	'if Department = "" then
	'	msg = msg & "Please enter a Department. "
	'end if
	'if Username = "" then 
	'	msg = msg & "Please enter a Username. "
	'end if
	
	'CHECK FOR INVALID DATES
	InvalidDates = ";"
	for i = 0 to 50
		if CanEdit(i)="Y" then
			if FieldType(i)="Date" then
				if trim(FormData(i))>"" then
					if IsDate(FormData(i))=false then
						msg = msg & "You have entered invalid dates. "
						InvalidDates = InvalidDates & i & ";"
					else
						if cdate(FormData(i)) < cdate("01/Jan/2000") or cdate(FormData(i))> cdate("31/Dec/2099") then
							msg = msg & "You have entered 'out of range' dates. "
							InvalidDates = InvalidDates & i & ";"
						else
							if instr(lcase(FieldData(i)), "actual")>0 then
								if cdate(FormData(i))>Date then
									msg = msg & FieldData(i) & " cannot be set in the future. "
									InvalidDates = InvalidDates & i & ";"
								end if 
							end if
						end if
					end if
				end if 
			end if
		
		end if				
	next	
	
	strsql = "SELECT * FROM Studies WHERE WebKey = '" & WebKey & "'"
	rs.open strsql,db,1,3
	StudyID = rs("StudyID")	
	strStudyNumber = rs("StudyNumber")
	
	for i = 0 to 50
		if CanEdit(i)<>"Y" then
			if FieldData(i)>"" then
				FormData(i) = rs(FieldData(i))
			end if
		end if
	next			
	'Response.Write InvalidDates & " - " & msg & "<br>"
	
	'PROCESS EXISTING SITES
		if CanEdit(18)="Y" then
			rs3.open "SELECT * FROM MultiSite WHERE StudyID = " & StudyID,db,1,3
			do until rs3.eof=true
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "TestFacility", rs3("TestFacility") & "", Request.form("TestFacility" & rs3("MultiSiteID")) 
				
				if rs3("TestFacility") <> Request.form("TestFacility" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "TestFacility" & rs3("MultiSiteID") & ","
				end if
				rs3("TestFacility") = Request.form("TestFacility" & rs3("MultiSiteID"))
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "TestSite", rs3("TestSite") & "", Request.form("TestSite" & rs3("MultiSiteID")) 
				if rs3("TestSite") <> Request.form("TestSite" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "TestSite" & rs3("MultiSiteID") & ","
				end if
				rs3("TestSite") = Request.form("TestSite" & rs3("MultiSiteID"))
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "OffSiteSD", rs3("OffSiteSD") & "", Request.form("OffSiteSD" & rs3("MultiSiteID")) 
				if rs3("OffSiteSD") <> Request.form("OffSiteSD" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "OffSiteSD" & rs3("MultiSiteID") & ","
				end if
				rs3("OffSiteSD") = Request.form("OffSiteSD" & rs3("MultiSiteID"))
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "SmithersSD", rs3("SmithersSD") & "", Request.form("SmithersSD" & rs3("MultiSiteID")) 
				if rs3("SmithersSD") <> Request.form("SmithersSD" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "SmithersSD" & rs3("MultiSiteID") & ","
				end if
				rs3("SmithersSD") = Request.form("SmithersSD" & rs3("MultiSiteID"))
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "OffSitePI", rs3("OffSitePI") & "", Request.form("OffSitePI" & rs3("MultiSiteID")) 
				if rs3("OffSitePI") <> Request.form("OffSitePI" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "OffSitePI" & rs3("MultiSiteID") & ","
				end if
				rs3("OffSitePI") = Request.form("OffSitePI" & rs3("MultiSiteID"))
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "SmithersPI", rs3("SmithersPI") & "", Request.form("SmithersPI" & rs3("MultiSiteID")) 
				if rs3("SmithersPI") <> Request.form("SmithersPI" & rs3("MultiSiteID")) then
					SavedSiteFields = SavedSiteFields & "SmithersPI" & rs3("MultiSiteID") & ","
				end if
				rs3("SmithersPI") = Request.form("SmithersPI" & rs3("MultiSiteID"))
				
				if Request.Form("SITE" & rs3("MultiSiteID")) ="REMOVE" then
					rs3("InActive")=1
					WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "InActive", rs3("InActive") & "", "1" 
					strChangedFields = strChangedFields & "Site Deleted,"
				
				end if
				if SavedSiteFields>"" and instr(strChangedFields,"SiteData,")=0 then
					strChangedFields = strChangedFields & "SiteData,"
				end if
				rs3.update
				rs3.movenext
			loop
		
			'PROCESS NEW SITE
			if TestFacilityNew > "" or TestSiteNew > "" or OffSiteSDNew > "" or SmithersSDNew > "" or SmithersPINew > "" or OffSitePINew >"" then
				'CREATE NEW SITE
				rs3.addnew
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "TestFacility", "New Record", TestFacilityNew
				rs3("TestFacility") = TestFacilityNew
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "TestSite", "New Record", TestSiteNew
				rs3("TestSite") = TestSiteNew
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "OffSiteSD", "New Record", OffSiteSDNew
				rs3("OffSiteSD") = OffSiteSDNew
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "SmithersSD", "New Record", SmithersSDNew 
				rs3("SmithersSD") = SmithersSDNew
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "OffSitePI", "New Record", OffSitePINew
				rs3("OffSitePI") = OffSitePINew
				
				WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Multisite", "SmithersPI", "New Record", SmithersPINew 
				rs3("SmithersPI") = SmithersPINew
				
				rs3("StudyID") = StudyID
				rs3("StudyNumber") = rs("StudyNumber")
				rs3("LastUpdate")=now()
				rs3("LastUpdatedBy") = Session("LoggedIn")
				t = CreateWebKey()
				rs3("WebKey")= t
				NewNoteOrSite = true
				strChangedFields = strChangedFields & "NewSite,"
				rs3.update
				rs3.close	
				
				rs3.open "SELECT * FROM Multisite WHERE WebKey = '" & t & "'",db
				NewSiteID = rs3("MultiSiteID")
				
			end if
			rs3.close
		end if
		
		if NoteNew > "" then
			rs3.open "SELECT * FROM Notes WHERE StudyID = " & StudyID,db,1,3
			rs3.addnew
			WriteToChangeLog "ChangeLogLarge", rs("StudyID"), Session("LoggedIn"), "Notes", "Note", "New Record", NoteNew
			rs3("Note") = NoteNew
			
			rs3("StudyID") = StudyID
			rs3("StudyNumber") = rs("StudyNumber")
			
			rs3("LastUpdate")=now()
			rs3("LastUpdatedBy") =session("Fullname")
			rs3("CreatedBy") = session("Fullname")
			rs3.update
			rs3.close
			rs3.open "SELECT * FROM Notes WHERE StudyID = " & StudyID & " AND CreatedBy = '" & session("Fullname") & "' ORDER BY DateStamp DESC",db
			if rs3.eof=false or rs3.bof=false then
				AddedNoteID = rs3("NoteID")
			end if
			rs3.close
			NewNoteOrSite = true
			strChangedFields = strChangedFields & "NewNote,"
		end if
	rs.close
	
	'Response.Write "Getting here"
	
	if msg = "" then
		'OK to Save
		strsql = "SELECT * FROM Studies WHERE WebKey = '" & WebKey & "'"
		rs.open strsql,db,1,3
		'if WebKey = "" then
		'	rs.addnew
		'	WebKey = CreateWebKey()
		'	rs("WebKey")= WebKey
		'	rs.update
		'	rs.close
		'	strsql = "SELECT * FROM Users WHERE WebKey = '" & WebKey & "'"
		'	rs.open strsql,db,1,3
		StudyID = rs("StudyID")	
		'end if
		'strChangedFields = ""
		for i = 0 to 50
			session("FieldChange")="N"
			if CanEdit(i)="Y" then
				if FieldType(i)="Date" then
				
					if CompareDates(FormData(i), rs(FieldData(i)) & "") = false then
						WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Studies", FieldData(i), rs(FieldData(i)) & "", FormData(i)
					end if
					
					if trim(rs(FieldData(i)) & "")="" and trim(FormData(i))>"" then
						'CHECK IF TRIGGER SET
						rsTrigger.open "SELECT * FROM Triggers WHERE TriggerField = '" & FieldData(i) & "' AND StudyID = " & rs("StudyID") & " AND IsActive = 1", db
						if rsTrigger.eof=false or rsTrigger.bof=false then
							'TRIGGER SET
							strmsg = "<font face=arial size=2><b>" & FieldData(i) & "</b> for <b>" & rs("StudyNumberNew") & "</b> populated with <b>" & FormData(i) & "</b> by <b>" & session("FullName") & "</b><br/><br/>"
							strmsg = strmsg & "Financial Client Name: <b>" & rs("FinancialClientName") & "</b><br/>"
							strmsg = strmsg & "Study Director: <b>" & rs("StudyDirector") & "</b><br/>"
							strmsg = strmsg & "Study Description: <b>"  & rs("StudyDescription") & "</b><br/>"
							strmsg = strmsg & "Study Cost: <b>&pound;"  & rs("TotalStudyCost") & "</b><br/>"
							
							strmsg = strmsg & "</font>"
							'a = SendTriggerEmail(FieldData(i),rs("StudyNumberNew"),"humphrey@data-craft.co.uk",strmsg)
							a = SendTriggerEmail(FieldData(i),rs("StudyNumberNew"),"Harrogate.contracts@smithers.com",strmsg)
							
						end if
						rsTrigger.close
					end if
					'Response.Write Expiry & "<br>"
					if FormData(i) = "" then
						rs(FieldData(i)) = null
					else
						rs(FieldData(i)) = FormData(i)
					end if
				else
					if FieldLength(i) > 50 then
						WriteToChangeLog "ChangeLogLarge", rs("StudyID"), Session("LoggedIn"), "Studies", FieldData(i), rs(FieldData(i)) & "", FormData(i)
					else
						WriteToChangeLog "ChangeLog", rs("StudyID"), Session("LoggedIn"), "Studies",  FieldData(i), rs(FieldData(i)) & "", FormData(i)
					end if
					'Response.Write FieldType(i) & "-" & FieldData(i) & " - " & FormData(i) & " - " & len(FormData(i))  & "<br>"
					rs(FieldData(i)) = FormData(i)
				end if
			else
				'CAN'T EDIT FIELD
				if FieldData(i)>"" then
					FormData(i) = rs(FieldData(i))
				end if
			end if
			if session("FieldChange")="Y" then
				strChangedFields = strChangedFields & FieldData(i) & ","
			end if
		next
		if strChangedFields>"" then
			strChangedFields = left(strChangedFields,len(strChangedFields)-1)
		end if
		'WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Surname", rs("Surname") & "", Surname 
		'rs("Surname") =Surname
		
		'WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Username", rs("Username") & "", Username 
		'rs("Username") = Username
		
		'WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Department", rs("Department") & "", Department 
		'rs("Department") = Department
		
		'WriteToChangeLog "ChangeLogLarge", rs("UserID"), Session("LoggedIn"), "Users", "Email", rs("Email") & "", Email
		'rs("Email") = Email
		
		
		
		
		
		
		
		'if Session("SystemAdmin")="-1" then
		'	if trim(StudyDirector) = "0" then
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "StudyDirector", rs("StudyDirector") & "", "0"
		'			rs("StudyDirector")=0	
		'	else
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "StudyDirector", rs("StudyDirector") & "", "-1"
		'			rs("StudyDirector") = -1	
		'	end if
		'
		'	if trim(QA) = "0" then
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "0"
		'			rs("QA")=0	
		'	else
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "QA", rs("QA") & "", "-1"
		'			rs("QA") = -1	
		'	end if
		
		'	if trim(Archivist) = "0" then
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Archivist", rs("Archivist") & "", "0"
		'			rs("Archivist")=0	
		'	else
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "Archivist", rs("Archivist") & "", "-1"
		'			rs("Archivist") = -1	
		'	end if
		
		'	if trim(SystemAdmin) = "0" then
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SystemAdmin", rs("SystemAdmin") & "", "0"
		'			rs("SystemAdmin")=0	
		''	else
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "SystemAdmin", rs("SystemAdmin") & "", "-1"
		'			rs("SystemAdmin") = -1	
		'	end if
			
		'	if trim(IsActive) = "0" then
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "IsActive", rs("IsActive") & "", "0"
		'			rs("IsActive")=0	
		'	else
		'			WriteToChangeLog "ChangeLog", rs("UserID"), Session("LoggedIn"), "Users", "IsActive", rs("IsActive") & "", "-1"
		'			rs("IsActive") = -1	
		'	end if
		'end if
		
		
		
		rs("LastUpdate")=now()
		rs("LastUpdatedBy") = Session("LoggedIn")
		
		rs.update
		rs.close
		'Response.end
		if NewNoteOrSite  = false then
			if strChangedFields="" and SavedSiteFields="" then
				Response.Redirect "Studies.asp"
				'Response.Write "got here: " & SavedSiteFields
				'Response.End
			end if
		else
			'Response.Redirect "EditStudies.asp?NewNote=True&Key=" & wk
		end if
	else
		'Response.Write "Message: " & msg
		'Error - back to form
	end if
else
	WebKey = replace(Request.querystring("Key"),"'","")
	strsql = "SELECT * FROM Studies WHERE WebKey = '" & WebKey & "'"
	
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		StudyID = rs("StudyID")	
		for i = 0 to 50
			if FieldData(i)>"" then
			'Response.Write FieldData(i) & "/"
				FormData(i) = rs(FieldData(i)) & ""
			
			end if
		next
		strCurrency = DisplayCurrency(rs("StudyCurrency") & "")
		'UserID = rs("UserID")
		'Firstname = rs("Firstname") & ""
		'Surname = rs("Surname") & ""
		'Department = rs("Department") & ""
		'Username = rs("Username") & ""
		'Email = rs("Email") & ""
		'SystemAdmin = rs("SystemAdmin") & ""
		'Archivist = rs("Archivist") & ""
		'QA = rs("QA") & ""
		'StudyDirector = rs("StudyDirector") & ""
		'IsActive = rs("IsActive") & ""
		
		
		
	else
		Response.Redirect "Studies.asp"
	end if
	rs.close
end if



UserTitle = FirstName & " " & Surname

ClientEdit = trim(ClientEdit)
Response.Write "<blockquote><h1>Edit Study " 

if trim(FormData(31))>"" then
	Response.Write FormData(31)
else
	Response.Write strStudyNumber
end if

If WebKey > "" then
	Response.Write UserTitle 
else
	Response.Write "New User"
end if
Response.Write "</h1>"
if msg>"" and ClientEdit > "" then
	Response.Write "<font color=red><b>" & msg & "</b></font>"
end if
if strChangedFields>"" and ClientEdit > "" then
	Response.Write "<font size=2 color=green>The following fields were updated: " & strChangedFields & "</font>"
end if
Response.Write "<form method=""Post"" action=""EditStudies.asp"">"

Response.Write "<table align=""center"" cellpadding=10 border=0><tr><td valign=top>"

'Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""3"">"


'Response.Write "<tr><td valign=top>Firstname<font color=red>*</font></td>"
'Response.Write "<td valign=top>" & TextBoxHTML("Firstname",Firstname,25,30,"") & "</td></tr>"

'Response.Write "<tr><td valign=top>Surname<font color=red>*</font></td>"
'Response.Write "<td valign=top>" & TextBoxHTML("Surname",Surname,25,30,"") & "</td></tr>"

'Response.Write "<tr><td valign=top>Windows Username<font color=red>*</font></td>"
'Response.Write "<td valign=top>" & TextBoxHTML("Username",Username,25,30,"") & "</td></tr>"

'Response.Write "<tr><td valign=top>Department</td>"
'Response.Write "<td valign=top>" & DataComboHTML("Department", Department, "SELECT LookupValue, LookupValue as D1 FROM LookupValues WHERE LookupID = 7 ORDER BY LookupValue ASC" , "", "") & "</td></tr>"

'Response.Write "<tr><td valign=top>Email</td>"
'Response.Write "<td valign=top>" & TextBoxHTML("Email",Email,40,100,"") & "</td></tr>"

'Response.Write "</table>"

'Response.Write "</td><td valign=""top"">"

Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""4"">"

for Section = 1 to 4

	Response.write "<tr bgcolor=#59545c><td colspan=4><font face=arial color=white>"
	select case Section
	case 1
		Response.Write "Study Details"
	case 2
		Response.Write "Study Timeline"
	case 3
		Response.Write "QA Reporting"
	case 4
		Response.Write "Archiving"
	end select
	Response.Write "</td></tr>"
	
	rs.open "SELECT * FROM FieldData WHERE Section = " & section & " ORDER BY Sequence ASC",db
	c=0
	do until rs.eof=true
		if trim(CanEdit(rs("FieldID"))) = "Y" then
			EditDisabled = ""
		else
			EditDisabled = "Y" 
			
		end if
		
		if c/2 = int(c/2) then
			Response.Write "<tr>"
		end if
		Response.Write "<td valign=top>" 
		if instr("," & strChangedFields & ",","," & rs("FieldName") & ",") > 0 then
			Response.Write "<font color=green>"
		end if
		Response.write rs("FormLabel")
		if rs("Mandatory")=1 then
			'Response.Write "<font color=red>*</font>"
		end if 
		Response.Write "</td>"
		Response.Write "<td valign=top>"
		
		if rs("FormFieldType") = "Date" then
			if EditDisabled = "Y" then
				Response.Write "<input id=""" & FieldData(rs("FieldID")) & """ name=""" & FieldData(rs("FieldID")) & """ type=""text"" tabindex=""" & c & """ value=""" & formatdate(FormData(rs("FieldID"))) & """"
				Response.Write " disabled"
				Response.Write " />"
			else
				Response.Write "<input id=""" & FieldData(rs("FieldID")) & """ name=""" & FieldData(rs("FieldID")) & """ type=""text"" tabindex=""" & c & """ value=""" & formatdate(FormData(rs("FieldID"))) & """"
				if instr(InvalidDates, cstr(rs("FieldID")))>0 then
					Response.Write " style=""background-color: #FF0000; color: #FFFFFF"" "
				end if
				Response.Write " />"
				Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('" & FieldData(rs("FieldID")) & "'),event);"" />"
				
			end if
		end if
		
		if rs("FormFieldType") = "Text" then
			Response.Write TextBoxHTML(FieldData(rs("FieldID")),FormData(rs("FieldID")),25,rs("FieldLength"),EditDisabled) & "</td>"
		end if
		
		if rs("FormFieldType") = "Combo" then
			if FieldData(rs("FieldID")) = "StudyType" or FieldData(rs("FieldID")) = "TestSubstanceName" or FieldData(rs("FieldID")) = "RegulatoryStatus" or FieldData(rs("FieldID")) = "FinancialClientName2" then
				Response.Write TextDataComboHTML(FieldData(rs("FieldID")), FormData(rs("FieldID")), rs("OptionsQuery") , EditDisabled, "Y") & "</td>"
			else
				Response.Write TextDataComboHTML(FieldData(rs("FieldID")), FormData(rs("FieldID")), rs("OptionsQuery") , EditDisabled, "") & "</td>"
			
			end if
		end if
		
		if rs("FormFieldType") = "CheckBox" then
			Response.Write CheckBoxHTML("ExcludeFromOverdue", FormData(rs("FieldID")), EditDisabled)
		end if
		
		if rs("FormFieldType") = "Display" then
			if rs("FieldType")="Money" then
				if FormData(rs("FieldID")) > "" then
					Response.Write "<font color=black>" & strCurrency & formatnumber(FormData(rs("FieldID")),2)
				else
					Response.Write "<font color=black>Not Set"
				end if
			else
				if FormData(rs("FieldID")) > "" then
					Response.Write "<font color=black>" & FormData(rs("FieldID")) 
				else
					Response.Write "<font color=black>Not Set"
				end if
			end if
			Response.Write "</td>"
		end if
		
		Response.Write "</td>"
		'if rs("FieldName") = "ProjectManagementCode" then
		'	c=c+1
		'end if
		if c/2 <> int(c/2) then
			Response.Write "</tr>"
		end if
		c=c+1
		rs.movenext
	loop
	rs.close
next

'MULTISITE
if CanEdit(18)="Y" then
	EditDisabled = ""
else
	EditDisabled = "Y"	
end if 
Response.write "<tr bgcolor=#59545c><td colspan=4><font face=arial color=white>Sites &nbsp;<font size=1>"
if EditDisabled = "" then
	Response.Write "<a href=""#"" id=""addmultisite"" onClick=""toggle_it('trnewmultisite');toggle_it('addmultisite');""><u>Add Site</u></a>"
end if
Response.Write "</td><tr>"
Response.Write "<tr><td colspan=4>"
Response.Write "<table width=""100%"">"
Response.Write "<tr>"
Response.Write "<td>Test Facility</td>"
Response.Write "<td>Test Site</td>"
Response.Write "<td>Offsite SD</td>"
Response.Write "<td>Smithers SD</td>"
Response.Write "<td>Offsite PI</td>"
Response.Write "<td>Smithers PI</td>"
Response.Write "<td>Remove</td>"
Response.Write "</tr>"

rs.open "SELECT * FROM MultiSite WHERE InActive = 0 AND StudyID = " & StudyID
do until rs.eof=true
	Response.write "<tr>"
	'Response.Write "<td><input type=text maxlength=50 name=TestFacility" & rs("MultiSiteID") & " value=""" & rs("TestFacility") & """></td>"
	'Response.Write "<td><input type=text maxlength=50 name=TestSite" & rs("MultiSiteID") & " value=""" & rs("TestSite") & """></td>"
	'Response.Write "<td><input type=text maxlength=50 name=OffSiteSD" & rs("MultiSiteID") & " value=""" & rs("OffsiteSD") & """></td>"
	'Response.Write "<td><input type=text maxlength=50 name=SmithersSD" & rs("MultiSiteID") & " value=""" & rs("SmithersSD") & """></td>"
	'Response.Write "<td><input type=text maxlength=50 name=SmithersPI" & rs("MultiSiteID") & " value=""" & rs("SmithersPI") & """></td>"
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "TestFacility" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("TestFacility" & rs("MultiSiteID"),rs("TestFacility"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("TestFacility" & rs("MultiSiteID"),rs("TestFacility"),25,50,EditDisabled) & "</td>"
	end if
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "TestSite" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("TestSite" & rs("MultiSiteID"),rs("TestSite"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("TestSite" & rs("MultiSiteID"),rs("TestSite"),25,50,EditDisabled) & "</td>"
	end if
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "OffSiteSD" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("OffSiteSD" & rs("MultiSiteID"),rs("OffSiteSD"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("OffSiteSD" & rs("MultiSiteID"),rs("OffSiteSD"),25,50,EditDisabled) & "</td>"
	end if
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "SmithersSD" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("SmithersSD" & rs("MultiSiteID"),rs("SmithersSD"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("SmithersSD" & rs("MultiSiteID"),rs("SmithersSD"),25,50,EditDisabled) & "</td>"
	end if
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "OffSitePI" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("OffSitePI" & rs("MultiSiteID"),rs("OffSitePI"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("OffSitePI" & rs("MultiSiteID"),rs("OffSitePI"),25,50,EditDisabled) & "</td>"
	end if
	
	if NewSiteID = rs("MultiSiteID") or instr(SavedSiteFields, "SmithersPI" & rs("MultiSiteID") & ",")>0 then
		Response.write "<td>" & TextBoxHTMLGreen("SmithersPI" & rs("MultiSiteID"),rs("SmithersPI"),25,50,EditDisabled) & "</td>"
	else
		Response.write "<td>" & TextBoxHTML("SmithersPI" & rs("MultiSiteID"),rs("SmithersPI"),25,50,EditDisabled) & "</td>"
	end if
	
	Response.write "<td><input type=checkbox name=SITE" & rs("MultiSiteID") & " value=REMOVE>" & "</td>"
	
	
	
	Response.Write "</tr>"
	rs.movenext
loop
rs.close
Response.write "<tr id=trnewmultisite style=""display: none;"">"
	Response.Write "<td><input type=text name=TestFacilityNew value=""""></td>"
	Response.Write "<td><input type=text name=TestSiteNew value=""""></td>"
	Response.Write "<td><input type=text name=OffSiteSDNew value=""""></td>"
	Response.Write "<td><input type=text name=SmithersSDNew value=""""></td>"
	Response.Write "<td><input type=text name=OffSitePINew value=""""></td>"
	Response.Write "<td><input type=text name=SmithersPINew value=""""></td>"
	Response.Write "<td></td>"
	
	Response.Write "</tr>"
Response.Write "</table>"

Response.Write "</td></tr>"


'NOTES


if CanEdit(18)="Y" then
	EditDisabled = ""
else
	EditDisabled = "Y"	
end if 

Response.write "<tr bgcolor=#59545c><td colspan=4><font face=arial color=white>Notes &nbsp;<font size=1>"
if EditDisabled = "" then
	Response.Write "<a href=""#"" id=""addnote"" onClick=""toggle_it('trnewnote');toggle_it('addnote');""><u>Add Note</u></a>"
end if
Response.Write "</td><tr>"
Response.Write "<tr><td colspan=4>"
Response.Write "<table width=""100%"">"
Response.Write "<tr>"
Response.Write "<td>Note</td>"
Response.Write "<td></td>"
Response.Write "<td></td>"
Response.Write "</tr>"

rs.open "SELECT * FROM Notes WHERE StudyID = " & StudyID & " ORDER BY DateStamp DESC"
c=0
do until rs.eof=true
	Response.write "<tr>"
	Response.write "<td width=""75%"">" 
	if c=0 and instr("," & strChangedFields & ",",",NewNote,")>0 then
		Response.Write "<font color=green>"
	end if
	if AddedNoteID = rs("NoteID") then
		Response.Write "<font color=green>"
	end if
	Response.Write rs("Note") & "</td>"
	Response.write "<td>" 
	if AddedNoteID = rs("NoteID") then
		Response.Write "<font color=green>"
	end if
	Response.Write rs("CreatedBy") & "</td>"
	Response.write "<td align=right>" 
	if AddedNoteID = rs("NoteID") then
		Response.Write "<font color=green>"
	end if
	Response.write formatdate(rs("DateStamp")) & "</td>"
	
	Response.Write "</tr>"
	c=c+1
	rs.movenext
loop
rs.close
Response.write "<tr id=""trnewnote"" style=""display: none;"">"
	Response.Write "<td colspan=3><textarea style=""font-family: arial;width: 100%"" rows=3 name=NoteNew></textarea></td>"
	

	Response.Write "</tr>"
Response.Write "</table>"

Response.Write "</td></tr>"



Response.Write "</td></tr>"
Response.Write "</table>"
'Response.Write "<tr><td valign=top>System Admin</td>"
'if UserID=Session("LoggedIn") then
'	Locked = "Yes"
'else
'	Locked = ""
'end if
'if Session("SystemAdmin") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("SystemAdmin",SystemAdmin,Locked) & "</td></tr>"
'else
'	Response.Write "<td valign=top>" & CheckBoxHTML("SystemAdmin",SystemAdmin,"Yes") & "</td></tr>"	
'end if

'Response.Write "<tr><td valign=top>Archivist</td>"
'if Session("Archivist") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("Archivist",Archivist,"Yes") & "</td></tr>"
'else
'	Response.Write "<td valign=top>" & CheckBoxHTML("Archivist",Archivist,"") & "</td></tr>"	
'end if

'Response.Write "<tr><td valign=top>QA</td>"
'if Session("QA") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("QA",QA,"Yes") & "</td></tr>"
'else'
'	Response.Write "<td valign=top>" & CheckBoxHTML("QA",QA,"") & "</td></tr>"	
'end if

'Response.Write "<tr><td valign=top>Study Director</td>"
'if Session("StudyDirector") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("StudyDirector",StudyDirector,"Yes") & "</td></tr>"
'else'
'	Response.Write "<td valign=top>" & CheckBoxHTML("StudyDirector",StudyDirector,"") & "</td></tr>"	
'end if

'Response.Write "<tr><td valign=top>Is Active</td>"
'if Session("SystemAdmin") = "-1" then
'	Response.Write "<td valign=top>" & CheckBoxHTML("IsActive",IsActive,"") & "</td></tr>"
'else
'	Response.Write "<td valign=top>" & CheckBoxHTML("IsActive",IsActive,"Yes") & "</td></tr>"	
'end if



'Response.Write "</table>"


'Response.Write "</td></tr>"




Response.Write "<tr><td colspan=""2""><hr></td></tr>"

Response.Write "<tr><td colspan=""2"" align=""center""><input type=""Submit"" name=""b1"" value=""Save"">&nbsp;&nbsp;&nbsp;&nbsp;"
if strChangedFields>"" and ClientEdit > "" then
	Response.Write "<input type=""Submit"" name=""b1"" value=""Back"">"
else
	Response.Write "<input type=""Submit"" name=""b1"" value=""Cancel"">"
end if
Response.Write "<input type=""hidden"" name=""ClientEdit"" value=""Yes""><input type=""hidden"" name=""WebKey"" value=""" & WebKey & """></td></tr>"

Response.Write "</table>"
Response.Write "</form>"
'Response.Write "<font color=white>" & strChangedFields & "</font>"
'Response.write "DeptMan: " & session("DepartmentManager") & "<br>"
'Response.write "DMDepartments: " & session("DMDepartments") & "<br>"
'Response.write "strDept: " & ucase(strDepartment) & "<br>"



Function SendTriggerEmail(TriggerField,StudyNumber,Recipient,Message)

	Set JMail = Server.CreateObject("JMail.SMTPMail")        
    'JMail.ServerAddress = "10.10.10.16"
    JMail.ServerAddress = "mailuk.smithers.com"
    JMail.Charset = "utf-8"
    JMail.Sender = "Harrogate.contracts@smithers.com"
    JMail.SenderName = "SEMS Triggers"
    JMail.Subject = "SEMS Trigger Activated: " & TriggerField & " populated for " & StudyNumber        
                      
    Jmail.AddRecipient Recipient
                                                
    'JMail.HTMLBody = Message
    JMail.ContentType = "text/html"
	JMail.Body = Message
    JMail.Priority = 3
    JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
                
                                                                
    JMail.Execute
                                                
                                                
	JMail.close
	set Jmail = nothing
	SendTriggerEmail = true

End Function
%>
<!-- #include file = "includes\footer.asp"-->
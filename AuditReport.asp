
<!-- #include file = "includes\CortexConfig.asp"-->
<script type='text/JavaScript' src='scw.js'></script>

<%
	StartDate = Request.Form("StartDate")
	EndDate = Request.Form("EndDate")
	
	

	if IsDate(StartDate)=false then
		msg = msg & "The Start Date is not valid. "
		StartDate = ""
	
	end if	
	if IsDate(EndDate)=false then
		msg = msg & "The End Date is not valid. "
		EndDate=""
		
	end if
	if StartDate="" or EndDate = "" then
		StartDate = "01/" & month(dateadd("m",-1,now())) & "/" & year(dateadd("m",-1,now()))
		EndDate = "01/" & month(now()) & "/" & year(now())
		EndDate = DateAdd("d",-1,EndDate)
	end if
	EndDate = formatdate(EndDate)
	StartDate = formatdate(StartDate)
	Response.Write "<h1>Audit Report</h1>"	
	'Response.Write "<p>All studies where the <b>ReportFinalSignDate</b> field is populated and the <b>ArchiveActDate</b> field is blank.</p>"
	
	StudyDirector = Request.Form("StudyDirector")
	if StudyDirector = "" then
		StudyDirector = "0"
	end if
	
	Department = Request.Form("Department")
	If Department = "" then
		Department = "All Departments"
	end if	
	StudyNumber = Request.Form("StudyNumber")
	StudyDate = Request.Form("StudyDates")
	if StudyDate = "" then
		'StudyDate = "ProtocolIssuedDate"
	end if
	Response.Write "<form method=post action=AuditReport.asp>"
	Response.Write "<table cellspacing=0 cellpadding=3>"
	Response.Write "<td><b>Filter Users</b></td><td><select name=StudyDirector>"
	Response.Write "<option value=0>All Users</option>"
	rs.open "SELECT UserID,FirstName,Surname FROM Users ORDER BY Firstname ASC",db
	do until rs.eof=true
		Response.Write "<option value=" & rs("UserID")
		if trim(rs("UserID"))=trim(StudyDirector) then
			Response.Write " selected"
		end if
		Response.Write ">" & rs("Firstname") & " " & rs("Surname") & "</option>"	
		rs.movenext
	loop
	rs.close
	Response.Write "</select></td>"
	Response.Write " <td><b>Filter Study Number</b></td><td>"
	Response.Write "<input type=text name=StudyNumber value=""" & StudyNumber & """>"
	Response.Write "</td></tr>"
	
	rs.open "SELECT * FROM FieldData WHERE FieldID <> 17 and FieldID <> 19 Order By Fieldname ASC",db
	do until rs.eof=true
		strDates = strDates & rs("FieldName") & ","
		rs.movenext
	loop
	rs.close
	strDates = left(strDates, len(strDates)-1)
	
	'strDates="Protocol Issued Date,Quoted Experimental Start Date,In Life Completion Date,Estimated Experimental Start Date,Actual Experimental Start Date,Estimated Experimental End Date,Actual Experimental End Date,Estimated Draft Report Date,Actual Draft Report Date,Quoted Draft Report Date,Report Final Sign Date,Report Finalisation Deadline,Archive Act Date,OA Start Date,OA Actual End Date,Last Update,Date Stamp"
	StudyDates = split(strDates,",")
	Response.Write "<tr><td><b>Filter Field</b></td><td><select name=StudyDates><option>All Fields</option>"
	'Response.Write "<option>None</option>"
	for i = 0 to ubound(StudyDates)
		Response.Write "<option"
		if StudyDate = StudyDates(i) then
			Response.Write " selected"
		end if
		Response.Write ">" & StudyDates(i) & "</option>"
	next
	Response.Write "</select></td></tr>"
	
	Response.Write "<tr><td valign=top><b>Start Date</b></td>"
	Response.Write "<td valign=top><input id=""StartDate"" name=""StartDate"" type=""text"" tabindex=""130"" value=""" & StartDate & """ />"
	Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('StartDate'),event);"" /></td>"

	Response.Write "<td valign=top><b>End Date</b></td>"
	Response.Write "<td valign=top><input  id=""EndDate"" name=""EndDate"" type=""text"" tabindex=""131"" value=""" & EndDate & """ />"
	Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('EndDate'),event);"" /></td>"
	Response.Write "<td> <input type=submit name=b1 value=Refresh></td>"
	Response.Write "</tr></table>"
	Response.Write "</form>"
	
	
	
	
	
	strsql = "SELECT * FROM AuditView WHERE ChangeTable = 'Studies'"	
	
	if StudyDirector <> "0" then
		strsql = strsql & " AND ChangedBy = '" & StudyDirector & "' "
	end if
	
	if StudyNumber > "" then
		strsql = strsql & " AND StudyNumber LIKE '" & StudyNumber & "%' "
	end if
	
	if StudyDate <>"All Fields" then
		strsql = strsql & " AND ChangeField = '" & StudyDate & "'"
	end if
	if StartDate > "" and EndDate >"" then
		strsql = strsql & " AND DateStamp >= '" & StartDate & " 00:00:00' AND DateStamp <= '" & EndDate & " 23:59:59'"
	end if
	
	'Response.Write strsql & "<br>"
	
	strsql = strsql & " UNION ALL " & replace(strsql,"AuditView","AuditViewLarge")
	strsql = strsql & " ORDER BY StudyNumber ASC"
	DisplayFields = "Study Number;Changed Field;Old Value;New Value;Changed By;Date Changed"
	'if StudyDate<>"All Fields" then
	'	DisplayFields = DisplayFields & ";" & StudyDate
	'end if
'Response.Write strsql
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		
		
		f = split(DisplayFields,";")
		Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		for i = 0 to ubound(f)
			Response.Write "<td><b>" 
			
			Response.Write f(i)
			 
			Response.Write "</b></td>"
			f(i) = replace(f(i)," ","")
		next 
		
		
		Response.Write "</tr>"
		
			
		c=0
		do until rs.eof=true or c=500 
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			Response.Write "<td valign=top>" & rs("StudyNumber") & "</td>"
			Response.Write "<td valign=top>" & rs("ChangeField") & "</td>"
			Response.Write "<td valign=top>" & rs("OldValue") & "</td>"
			Response.Write "<td valign=top>" & rs("NewValue") & "</td>"
			Response.Write "<td valign=top>" & rs("Firstname") & " " & rs("Surname") & "</td>"
			
			Response.Write "<td valign=top>" & rs("DateStamp") & "</td>"
			
			
			
			
			
			
			
			'STRATEGIC PARTNER
			
			'Response.Write "<td valign=top><a href=""ViewStudy.asp?Key=" & rs("WebKey") & """>View</a></td>"
			'Response.Write "<td valign=top><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
			Response.Write "</tr>"
			rs.movenext
			c=c+1
		loop
		
		Response.Write "</table>"
	
	end if
	
	rs.close
	Response.Write "<br>Records: " & c


%>
<!-- #include file = "includes\footer.asp"-->
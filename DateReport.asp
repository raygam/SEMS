
<!-- #include file = "includes\CortexConfig.asp"-->
<script type='text/JavaScript' src='scw.js'></script>
<script src="scripts\sorttable.js"></script>
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
		EndDate = Date
	end if
	EndDate = formatdate(EndDate)
	StartDate = formatdate(StartDate)
	Response.Write "<h1>Date Report</h1>"	
	'Response.Write "<p>All studies where the <b>ReportFinalSignDate</b> field is populated and the <b>ArchiveActDate</b> field is blank.</p>"
	
	StudyDirector = Request.Form("StudyDirector")
	if StudyDirector = "" then
		StudyDirector = "All Study Directors"
	end if
	
	Department = Request.Form("Department")
	If Department = "" then
		Department = "All Departments"
	end if	
	
	StudyDate = Request.Form("StudyDates")
	if StudyDate = "" then
		StudyDate = "ProtocolIssuedDate"
	end if
	Response.Write "<form method=post action=DateReport.asp>"
	Response.Write "<table cellspacing=0 cellpadding=3>"
	Response.Write "<td><b>Filter Study Directors</b></td><td><select name=StudyDirector>"
	Response.Write "<option>All Study Directors</option>"
	rs.open "SELECT StudyDirector FROM Studies WHERE StudyDirector IS NOT NULL GROUP BY StudyDirector ORDER BY StudyDirector ASC",db
	do until rs.eof=true
		Response.Write "<option"
		if trim(rs("StudyDirector"))=trim(StudyDirector) then
			Response.Write " selected"
		end if
		Response.Write ">" & rs("StudyDirector") & "</option>"	
		rs.movenext
	loop
	rs.close
	Response.Write "</select></td>"
	Response.Write " <td><b>Filter Department</b></td><td><select name=Department>"
	Response.Write "<option>All Departments</option>"
	rs.open "SELECT Department FROM Studies WHERE Department IS NOT NULL GROUP BY Department ORDER BY Department ASC",db
	do until rs.eof=true
		Response.Write "<option"
		if trim(rs("Department"))=trim(Department) then
			Response.Write " selected"
		end if
		Response.Write ">" & rs("Department") & "</option>"	
		rs.movenext
	loop
	rs.close
	
	Response.Write "</select> </td></tr>"
	
	
	'strDates="Protocol Issued Date,Quoted Experimental Start Date,In Life Completion Date,Estimated Experimental Start Date,Actual Experimental Start Date,Estimated Experimental End Date,Actual Experimental End Date,Estimated Draft Report Date,Actual Draft Report Date,Quoted Draft Report Date,Report Final Sign Date,Report Finalisation Deadline,Archive Act Date,Open Acc Start Date,Open Acc Actual End Date,Last Update,Date Stamp"
	strDates="Quoted Experimental Start Date,Quoted Draft Report Date,Expected Receipt of Test Substance,Actual Receipt of Test Substance,Actual Protocol Issued Date,Actual Experimental Start Date,Expected Preliminary Initiation,"
	strDates = strDates & "Actual Preliminary Initiation,Expected Preliminary Termination,Actual Preliminary Termination,Expected Definitive Initiation,Actual Definitive Initiation,Expected Definitive Termination,Actual Definitive Termination,Actual In Life Completion,Actual Experimental End Date,Expected Unaudited Draft to Client,Actual Unaudited Draft to Client,Expected Draft Report to QA,Actual Draft Report to QA,Expected Audit Findings to SD,Actual Audit Findings to SD,Expected Audited Draft to Client,Actual Audited Draft to Client"
	strDates = strDates & ",Expected Final Report Issue,Actual Final Report Issue,Expected Archive Date,Actual Archive Date"


	
	
	
	
	StudyDates = split(strDates,",")
	Response.Write "<tr><td><b>Filter Date</b></td><td><select name=StudyDates>"
	'Response.Write "<option>None</option>"
	for i = 0 to ubound(StudyDates)
		Response.Write "<option"
		if StudyDate = StudyDates(i) then
			Response.Write " selected"
		end if
		Response.Write ">" & replace(StudyDates(i),"OA","Open Acc") & "</option>"
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
	
	
	
	
	
	strsql = "SELECT * FROM Studies WHERE StudyID > 0 "	
	
	if StudyDirector <> "All Study Directors" then
		strsql = strsql & " AND StudyDirector = '" & StudyDirector & "' "
	end if
	
	if Department <> "All Departments" then
		strsql = strsql & " AND Department = '" & Department & "' "
	end if
	
	
	
	
	StudyDate = replace(StudyDate,"Open Acc", "OA")
	StudyDate = replace(StudyDate,"Actual In Life Completion","InLifeCompletionDate")
	StudyDate = replace(StudyDate,"Expected Audited Draft to Client","EstimatedDraftReportDate")
	StudyDate = replace(StudyDate,"Actual Audited Draft to Client","ActualDraftReportDate")
	StudyDate = replace(StudyDate,"Expected Final Report Issue","ReportFinalisationDeadline")
	StudyDate = replace(StudyDate,"Actual Final Report Issue","ReportFinalSignDate")
	
	DisplayFields = "Study Number;Financial ClientName;Test Substance Name;Study Description;Study Director;Department;OA Status;Total Study Cost"
	if StudyDate>"" then
		DisplayFields = DisplayFields & ";" & StudyDate
	end if
	
	
	
	if StudyDate >"" then
		strsql = strsql & " AND " & replace(StudyDate," ","") & " >= '" & StartDate & " 00:00:00' AND " & replace(StudyDate," ","") & " <= '" & EndDate & " 23:59:59'"
	end if
	
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	
'Response.Write strsql
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		
		
		f = split(DisplayFields,";")
		Response.Write "<table class=""sortable"" width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		for i = 0 to ubound(f)
			Response.Write "<td style=""cursor: pointer""><b>" 
			t = f(i)
			t = replace(t,"InLifeCompletionDate","Actual In Life Completion")
			t = replace(t,"EstimatedDraftReportDate","Expected Audited Draft To Client")
			t = replace(t,"ActualDraftReportDate","Actual Audited Draft To Client")
			t = replace(t,"ReportFinalisationDeadline","Expected Final Report Issue")
			t = replace(t,"ReportFinalSignDate","Actual Final Report Issue")
			Response.Write replace(t,"OA","Open Acc")
			 
			Response.Write "</b></td>"
			f(i) = replace(f(i)," ","")
		next 
		
		
		Response.Write "</tr>"
		
			
		c=0
		do until rs.eof=true 
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			for i = 0 to ubound(f)
				'Response.Write "<font color=red>" & f(i) & "</font>"
				if f(i)<>"DaysSinceRFSD" then
					if isdate(rs(f(i)))= true then
						Response.Write "<td align=right valign=top>" & formatdate2(rs(f(i))) 
					else
					
						if f(i)="TotalStudyCost" then
							Response.Write "<td align=right valign=top>"
							if isnull(rs(f(i)))=false then
								
								Response.Write "&pound;" & formatnumber(rs(f(i)),2)
							end if
						else
							Response.Write "<td valign=top>"
							if f(i)="OAStatus" then
								Response.Write DisplayOAStatus(rs(f(i)) & "")
							else
								Response.Write rs(f(i)) 
							end if
						end if
					
					end if
				else
					Response.Write "<td align=right valign=top>"
					if isDate(rs("ReportFinalSignDate"))=true then
						
						d = cint(datediff("d",rs("ReportFinalSignDate"),now()))
						if d>=30 then
							Response.Write "<font color=red>"
						end if	
						Response.Write d
						if d>=30 then
							Response.Write "</font>"
						end if
					else
						Response.Write "NA"
					end if
				end if
				Response.Write "</td>"
			next
			
			
			
			
			'STRATEGIC PARTNER
			
			'Response.Write "<td valign=top><a href=""ViewStudy.asp?Key=" & rs("WebKey") & """>View</a></td>"
			Response.Write "<td valign=top><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
			Response.Write "</tr>"
			rs.movenext
			c=c+1
		loop
		
		Response.Write "</table>"
	
	end if
	
	rs.close
	Response.Write "<br>Records: " & c
	
	Function FormatDate2(strDate)
		s = trim(strDate)
		if s & "" = "" then 
			FormatDate2 = ""
		else
			if isdate(s)=false then
				FormatDate2 = "##/###/####"
			else	
				FormatDate2 = right("0" & day(s),2) & "/" & right("0" & month(s),2) & "/" & year(s)
			end if
		end if
	
	End Function
%>


%>
<!-- #include file = "includes\footer.asp"-->
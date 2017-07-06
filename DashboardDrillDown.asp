
<!-- #include file = "includes\CortexConfig.asp"-->


<%

	strsql = "SELECT * FROM Studies WHERE OAStatus = 'L' AND StudyDirector = '" & session("FullName") & "'"
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	d = Request.QueryString("d")
	d = ConvertToGetSearch(d)
	y = Request.QueryString("y")
	
	strsql = "SELECT * FROM Studies WHERE "
	Select Case y
	Case 1
		t = "All Live Studies"
		strsql = strsql & "OAStatus = 'L'" 
	Case 2
		t = "No Estimated Experimental Start Date"
		strsql = strsql & "OAStatus = 'L' AND EstimatedExperimentalStartDate IS NULL AND ActualExperimentalStartDate IS NULL"
	Case 3
		t = "Upcoming Estimated Experimental Start Date"
		strsql = strsql & "OAStatus = 'L' AND ActualExperimentalStartDate IS NULL AND EstimatedExperimentalStartDate > '" & FormatDate(now()) & "' AND EstimatedExperimentalStartDate < '" & FormatDate(dateadd("d",14,now())) & "'"
	Case 4
		t = "Overdue Experimental Start Date"
		strsql = strsql & "OAStatus = 'L' AND ActualExperimentalStartDate IS NULL AND EstimatedExperimentalStartDate < '" & FormatDate(now()) & "'"
	Case 5
		t = "Upcoming Estimated Experimental End Date"
		strsql = strsql & "OAStatus = 'L' AND ActualExperimentalEndDate IS NULL AND EstimatedExperimentalEndDate > '" & FormatDate(now()) & "' AND EstimatedExperimentalEndDate < '" & FormatDate(dateadd("d",14,now())) & "'"
	Case 6
		t = "Overdue Experimental End Date"
		strsql = strsql & "OAStatus = 'L' AND ActualExperimentalEndDate IS NULL AND EstimatedExperimentalEndDate < '" & FormatDate(now()) & "'"
	Case 7
		t = "Upcoming Estimated Draft Report Date"
		strsql = strsql & "OAStatus = 'L' AND ActualDraftReportDate IS NULL AND EstimatedDraftReportDate > '" & FormatDate(now()) & "' AND EstimatedDraftReportDate < '" & FormatDate(dateadd("d",14,now())) & "'"
	Case 8
		t = "Overdue Draft Report Date"
		strsql = strsql & "OAStatus = 'L' AND ActualDraftReportDate IS NULL AND EstimatedDraftReportDate < '" & FormatDate(now()) & "'"
	End Select
	if d <> "All Departments" then
		strsql = strsql & " AND Department = '" & d & "'"
	end if
	
	
	
	DisplayFields = "Study Number;Financial ClientName;Test Substance Name;Study Description;Study Director;Department;OA Status;Total Study Cost;Protocol Issued Date;Estimated Experimental Start Date;Estimated Experimental End Date"


	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		Response.Write "<h1>" & t & " (" & d & ")</h1>"	
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
		do until rs.eof=true 
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			for i = 0 to ubound(f)
				'Response.Write "<font color=red>" & f(i) & "</font>"
				if isdate(rs(f(i)))= true then
					Response.Write "<td align=right valign=top>" & formatdate(rs(f(i))) 
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
		Response.Write "<br><font face=arial size=2>Records: " & c & "</font>"
	end if
	
	rs.close



%>
<!-- #include file = "includes\footer.asp"-->
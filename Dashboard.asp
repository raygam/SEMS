
<!-- #include file = "includes\CortexConfig.asp"-->


<%

	Response.Write "<h1>My Live Studies</h1>"
	
	dim grid(20,50)
	for x = 0 to 20
		for y = 0 to 50
			Grid(x,y) = "0"	
		next
	next
	Grid(0,0)=""
	
	rs.open "SELECT Department FROM Studies WHERE Department > '' GROUP BY Department ORDER By Department ASC",db
	c=1
	do until rs.eof=true
		grid(0,c) = rs("Department")
		rs.movenext
		c=c+1
	loop 
	rs.close
	intCols = c - 1


	'ALL LIVE STUDIES
	r=1
	Grid(r,0)="All Live Studies"
	
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' GROUP BY Department"
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close

	'NO EXP START DATE
	r = 2
	Grid(r,0)="No Estimated Experimental Start Date"
	
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND EstimatedExperimentalStartDate IS NULL AND ActualExperimentalStartDate IS NULL GROUP BY Department"
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'UPCOMING ESTIMATED EXP START DATE
	r = 3
	Grid(r,0)="Upcoming Estimated Experimental Start Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualExperimentalStartDate IS NULL AND EstimatedExperimentalStartDate > '" & FormatDate(now()) & "' AND EstimatedExperimentalStartDate < '" & FormatDate(d) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'OVERDUE ESTIMATED EXP START DATE
	r = 4
	Grid(r,0)="Overdue Experimental Start Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualExperimentalStartDate IS NULL AND EstimatedExperimentalStartDate < '" & FormatDate(now()) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'UPCOMING ESTIMATED EXP END DATE
	r = 5
	Grid(r,0)="Upcoming Estimated Experimental End Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualExperimentalEndDate IS NULL AND EstimatedExperimentalEndDate > '" & FormatDate(now()) & "' AND EstimatedExperimentalEndDate < '" & FormatDate(d) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'OVERDUE ESTIMATED EXP START DATE
	r = 6
	Grid(r,0)="Overdue Experimental End Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualExperimentalEndDate IS NULL AND EstimatedExperimentalEndDate < '" & FormatDate(now()) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'UPCOMING DRAFT REPORT DATE
	r = 7
	Grid(r,0)="Upcoming Estimated Draft Report Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualDraftReportDate IS NULL AND EstimatedDraftReportDate > '" & FormatDate(now()) & "' AND EstimatedDraftReportDate < '" & FormatDate(d) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'OVERDUE DRAFT REPORT DATE
	r = 8
	Grid(r,0)="Overdue Draft Report Date"
	d = dateadd("d",14,now())
	strsql = "SELECT Department, COUNT(StudyID) AS CountOfStudyID FROM Studies WHERE OAStatus = 'L' AND ActualDraftReportDate IS NULL AND EstimatedDraftReportDate < '" & FormatDate(now()) & "' GROUP BY Department"
	'Response.Write strsql
	rs.open strsql,db
	do until rs.eof=true
		for i = 1 to intCols
			if Grid(0,i) = rs("Department") then
				Grid(r,i) = rs("CountOfStudyID")
			end if
		next
		rs.movenext
	loop
	rs.close
	
	'Grid(r+1,0)="Totals"
	for y = 1 to r
		t = 0
		for x = 1 to intCols
			t = t + cint(Grid(y,x))
		next
		Grid(y,intCols+1) = t
	next
	
	Grid(0,intCols+1) = "All Departments"
	Response.Write "<table width=100% cellspacing=0 cellpadding=3>"
	for y = 0 to r
		Response.Write "<tr"
		if y/2 = int(y/2) and y >1 then
			Response.Write " bgcolor=#E1E1FF"
		end if
		Response.Write "><td><b>" & grid(y,0) & "</td>"
		for x = 1 to intCols+1
			if y = 0 then
				Response.Write "<td align=center colspan=2"
				Response.Write "><b><font color=black>"
				
				Response.write Grid(y, x)
				Response.Write "</td><td>&nbsp;&nbsp;&nbsp;</td>"
			else
				Response.Write "<td"
				if x>0 then
					Response.Write " align=right"
				end if
				Response.Write ">"
				if x = intCols+1 then
					Response.Write "<b>"
				end if
				Response.Write "<a href=DashboardDrillDown.asp?"
				Response.Write "y=" & y
				Response.Write "&d=" & ConvertToSendSearch(Grid(0,x))
				Response.Write ">" & Grid(y, x) & "</a>"
				Response.Write "</td>"
				if x<intCols+1 then
					Response.Write "<td align=right>"
					intTotal = cint(Grid(1,x))
					if intTotal > 0 then
						Response.Write formatnumber((cint(Grid(y,x))/intTotal) * 100,1) & "%"
					end if
					Response.Write "</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				else
					Response.Write "<td align=right>"
					intTotal = cint(Grid(1,intCols+1))
					if intTotal > 0 then
						Response.Write formatnumber((cint(Grid(y,x))/intTotal) * 100,1) & "%"
					end if
					Response.Write "</td>"
					'Response.Write "</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
				end if
			end if
		next 
		Response.Write "</tr>"
	next
	Response.Write "</table>"
	
	Response.End
	
	
	strsql = "SELECT * FROM Studies WHERE OAStatus = 'L' AND StudyDirector = '" & session("FullName") & "'"
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	DisplayFields = "Study Number;Financial Client Name;Client Name;Test Substance Name;Study Type;Study Director;Department;OA Status;Total Study Cost;Protocol Issued Date;Estimated Experimental Start Date;Actual Experimental Start Date;Quoted Draft Report Date;Actual Draft Report Date"


	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		Response.Write "<h1>My Live Studies</h1>"	
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
	
	end if
	
	rs.close



%>
<!-- #include file = "includes\footer.asp"-->
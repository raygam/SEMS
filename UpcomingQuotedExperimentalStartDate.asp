
<!-- #include file = "includes\CortexConfig.asp"-->


<%
	Response.Write "<h1>Upcoming Quoted Experimental Start Date</h1>"	
	Response.Write "<p>All studies where the <b>QuotedExperimentalStartDate</b> is in the future and the <b>ActualExperimentalStartDate</b> field is blank.</p>"
	
	StudyDirector = Request.Form("StudyDirector")
	if StudyDirector = "" then
		StudyDirector = "All Study Directors"
	end if
	
	Department = Request.Form("Department")
	If Department = "" then
		Department = "All Departments"
	end if	
	
	
	Response.Write "<form method=post action=UpcomingQuotedExperimentalStartDate.asp>"
	
	Response.Write "<b>Filter Study Directors</b>  <select name=StudyDirector>"
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
	Response.Write "</select>"
	Response.Write " &nbsp; &nbsp;<b>Filter Department</b>  <select name=Department>"
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
	
	Response.Write "</select>  <input type=submit name=b1 value=Refresh></form>"
	
	d = dateadd("d",45,now())
	
	'strsql = "SELECT * FROM Studies WHERE QuotedExperimentalStartDate <= '" & formatdate(d) & "' AND ActualExperimentalStartDate IS NULL"	
	strsql = "SELECT * FROM Studies WHERE QuotedExperimentalStartDate >= '" & formatdate(Date) & " 00:00:00' AND ActualExperimentalStartDate IS NULL"	
	
	if StudyDirector <> "All Study Directors" then
		strsql = strsql & " AND StudyDirector = '" & StudyDirector & "' "
	end if
	
	if Department <> "All Departments" then
		strsql = strsql & " AND Department = '" & Department & "' "
	end if
	
	'strsql = strsql & " ORDER BY StudyNumber ASC"
	strsql = strsql & " ORDER BY QuotedExperimentalStartDate ASC"
	
	'Response.Write strsql & "<br>"
	
	DisplayFields = "Study Number;Financial ClientName;Test Substance Name;Study Description;Study Director;Department;OA Status;Total Study Cost;Protocol Issued Date;Quoted Experimental Start Date;Days Until QESD"

'Response.Write strsql
c=0
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		
		
		f = split(DisplayFields,";")
		Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		for i = 0 to ubound(f)
			Response.Write "<td><b>" 
			
			Response.Write replace(f(i),"OA","Open Acc")
			 
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
				if f(i)<>"DaysUntilQESD" then
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
				else
					Response.Write "<td align=right valign=top>"
					if isDate(rs("QuotedExperimentalStartDate"))=true then
						
						d = cint(datediff("d",now(),rs("QuotedExperimentalStartDate")))
						if d>0 and d<=10 then
							Response.Write "<font color=orange>"
						end if	
						if d<0 then
							Response.Write "<font color=red>"
						end if	
						if d>10 then
							Response.Write "<font color=green>"
						end if
						Response.Write d
						
							Response.Write "</font>"
						
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


%>
<!-- #include file = "includes\footer.asp"-->
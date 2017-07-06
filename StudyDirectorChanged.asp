
<!-- #include file = "includes\CortexConfig.asp"-->


<%
	Response.Write "<h1>Study Director Changed Report</h1>"	
	Response.Write "<p>All studies where the <b>StudyDirector</b> field has been changed in open Accounts in the last 7 days.</p>"
	
	StudyDirector = Request.Form("StudyDirector")
	if StudyDirector = "" then
		StudyDirector = "All Study Directors"
	end if
	
	Department = Request.Form("Department")
	If Department = "" then
		Department = "All Departments"
	end if	
	
	
	Response.Write "<form method=post action=StudyDirectorChanged.asp>"
	
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
	
	d = dateadd("d",-7,now())
	strsql = "SELECT * FROM Studies WHERE QuotedDraftReportDate <= '" & formatdate(d) & "' AND ActualDraftReportDate IS NULL"	
	
	if StudyDirector <> "All Study Directors" then
		strsql = strsql & " AND StudyDirector = '" & StudyDirector & "' "
	end if
	
	if Department <> "All Departments" then
		strsql = strsql & " AND Department = '" & Department & "' "
	end if
	
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	DisplayFields = "Study Number;Financial Client Name;Client Name;Test Substance Name;Study Description;Study Director;Department;OA Status;Change Date;Changed From;Changed To"

'Response.Write strsql
	set rsAudit = server.CreateObject("ADODB.Recordset")
	rsAudit.open "SELECT * FROM ChangeLog WHERE ChangeField = 'StudyDirector' AND NewValue <> OldValue AND DateStamp >= '" & formatdate(d) & "' ORDER BY RecordID Asc",db
	
	if rsAudit.eof=false or rsAudit.bof=false then
	
	f = split(DisplayFields,";")
		Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		for i = 0 to ubound(f)
			Response.Write "<td><b>" 
			
			Response.Write replace(f(i),"OA","Open Acc")
			 
			Response.Write "</b></td>"
			f(i) = replace(f(i)," ","")
		next 
		
		
		Response.Write "</tr>"
	do until rsAudit.eof=true

		strsql = "SELECT * FROM Studies WHERE StudyID = " & rsAudit("RecordID")
		if StudyDirector <> "All Study Directors" then
			strsql = strsql & " AND StudyDirector = '" & StudyDirector & "' "
		end if
	
		if Department <> "All Departments" then
			strsql = strsql & " AND Department = '" & Department & "' "
		end if
		rs.open strsql,db
	
		if rs.eof=false or rs.bof=false then
		
		
		
			
		
		
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			for i = 0 to 7
				'Response.Write "<font color=red>" & f(i) & "</font>"
				if f(i)<>"DaysUntilQDRD" then
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
					if isDate(rs("QuotedDraftReportDate"))=true then
						
						d = cint(datediff("d",now(),rs("QuotedDraftReportDate")))
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
			
			Response.Write "<td align=right valign=top>" & Formatdate(rsAudit("DateStamp")) & "</td>"
			Response.Write "<td align=right valign=top>" & rsAudit("OldValue") & "</td>"
			Response.Write "<td align=right valign=top>" & rsAudit("NewValue") & "</td>"
			'Response.Write "<td align=right valign=top>Changed in OA</td>"
			
			
			'STRATEGIC PARTNER
			
			'Response.Write "<td valign=top><a href=""ViewStudy.asp?Key=" & rs("WebKey") & """>View</a></td>"
			Response.Write "<td valign=top><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
			Response.Write "</tr>"
			c=c+1
			end if
			rs.close
			rsAudit.movenext
			
		loop
		
		Response.Write "</table>"
	
	end if
	rsAudit.close
	
	Response.Write "<br>Records: " & c


%>
<!-- #include file = "includes\footer.asp"-->
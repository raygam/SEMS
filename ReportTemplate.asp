
<!-- #include file = "includes\CortexConfig.asp"-->


<%

	strsql = "SELECT * FROM Studies WHERE OAStatus = 'L' AND StudyDirector = '" & session("FullName") & "'"
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	DisplayFields = "Study Number;Financial ClientName;Test Substance Name;Study Description;Study Director;Department;OA Status;Total Study Cost;Protocol Issued Date;Estimated Experimental Start Date;Estimated Experimental End Date"


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
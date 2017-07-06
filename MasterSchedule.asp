
<!-- #include file = "includes\CortexConfig.asp"-->
<script type='text/JavaScript' src='scw.js'></script>

<%
	StartDate = Request.Form("StartDate")
	EndDate = Request.Form("EndDate")
	RegulatoryStatus = Request.Form("RegulatoryStatus")
	if RegulatoryStatus = "" then
		RegulatoryStatus = "Any"
	end if 

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
	Response.Write "<h1>Master Schedule</h1>"	
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
		'StudyDate = "ProtocolIssuedDate"
	end if
	Response.Write "<form method=post action=MasterSchedule.asp>"
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
	
	Response.Write "</select> </td>"
	
	Response.Write " <td>&nbsp;&nbsp;&nbsp;<b>Filter Regulatory Status</b></td><td><select name=RegulatoryStatus>"
	Response.Write "<option>Any</option>"
	strTemp = "Not Set,GLP,Non-GLP"
	RegStats = split(strTemp,",")
	for i = 0 to ubound(RegStats)
	
		Response.Write "<option"
		if RegStats(i)=trim(RegulatoryStatus) then
			Response.Write " selected"
		end if
		Response.Write ">" & RegStats(i) & "</option>"	
		
	next
	
	Response.Write "</select> </td></tr>"
	
	
	'strDates="Protocol Issued Date,Quoted Experimental Start Date,In Life Completion Date,Estimated Experimental Start Date,Actual Experimental Start Date,Estimated Experimental End Date,Actual Experimental End Date,Estimated Draft Report Date,Actual Draft Report Date,Quoted Draft Report Date,Report Final Sign Date,Report Finalisation Deadline,Archive Act Date,Open Acc Start Date,Open Acc Actual End Date,Last Update,Date Stamp"
	strDates="Protocol Issued Date,Actual Experimental Start Date,Actual Experimental End Date,Estimated Draft Report Date,Quoted Draft Report Date,Actual Draft Report Date"
	StudyDates = split(strDates,",")
	Response.Write "<tr><td><b>Filter Date</b></td><td><select name=StudyDates><option>No Filter</option>"
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
	if StudyDate >"" and StudyDate <> "No Filter" then
		strsql = strsql & " AND " & replace(StudyDate," ","") & " >= '" & StartDate & " 00:00:00' AND " & replace(StudyDate," ","") & " <= '" & EndDate & " 23:59:59'"
	end if
	
	Select Case RegulatoryStatus
	Case "Any"
	
	Case "Not Set"
		strsql = strsql & " AND (RegulatoryStatus = '' OR RegulatoryStatus IS NULL)"
	Case else
		strsql = strsql & " AND RegulatoryStatus = '" & RegulatoryStatus & "'"
	
	End Select
	
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	'DisplayFields = "Study Number;Financial ClientName;Test Substance Name;Study Description;Study Director;Department;OA Status;Total Study Cost"
	
	DisplayFields = "Study Number;Client Name;Financial Client Name;Test Substance Name;Study Description;Study Director;Department;Regulatory Status;Protocol Issued Date;Actual Experimental Start Date;Actual Experimental End Date;Quoted Draft Report Date;Actual Draft Report Date;Report Final Sign Date;Archive Act Date;Test Site Details"
	
	
	
	if StudyDate>"" then
		'DisplayFields = DisplayFields & ";" & StudyDate
	end if
'Response.Write strsql
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		
		set dbExcel = server.CreateObject("ADODB.Connection")
		set rsExcel = server.CreateObject("ADODB.Recordset")
		set cmdExcel = server.CreateObject("ADODB.Command")
		
		
		dt = formatdate(now()) & mid(now(),instr(now()," ")+1)
		dt = replace(dt," ","")
		dt = replace(dt,":","")
		dt = replace(dt,"/","")
		strPath = "c:\Websites\SEMS\wwwroot\Excel\MasterSchedule" & dt & ".xlsx"
		strDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"""

        dbExcel.Open strDB
        cmdExcel.ActiveConnection = dbExcel
        
        cmdExcel.CommandText = "CREATE TABLE Studies (StudyNumber VARCHAR, ClientName VARCHAR, FinancialClientName VARCHAR, TestSubstanceName VARCHAR, StudyDescription VARCHAR, StudyDirector VARCHAR, Department VARCHAR, RegulatoryStatus VARCHAR, ProtocolIssuedDate DATE, ActualExperimentalStartDate DATE, ActualExperimentalEndDate DATE, QuotedDraftReportDate DATE, ActualDraftReportDate DATE, ReportFinalSignDate DATE, ArchiveActDate DATE, TestSiteDetails VARCHAR)"
        cmdExcel.Execute
        set cmdExcel = nothing
		
		f = split(DisplayFields,";")
		Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		for i = 0 to ubound(f)
			Response.Write "<td><b>" 
			t = f(i)
			t = replace(t,"Actual Draft Report Date","Actual Audited Draft to Client")
			t = replace(t,"Estimated Draft Report Date","Expected Audited Draft to Client")
			t = replace(t,"Archive Act Date","Actual Archive Date")
			t = replace(t,"Report Final Sign Date","Actual Final Report Issue")
			
			Response.Write replace(t,"OA","Open Acc")
			 
			Response.Write "</b></td>"
			f(i) = replace(f(i)," ","")
		next 
		
		
		Response.Write "</tr>"
		
			
		c=0
		do until rs.eof=true 
			DataList = ""
			FieldList = ""
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			for i = 0 to ubound(f)
				'Response.Write "<font color=red>" & f(i) & "</font>"
				if f(i)<>"DaysSinceRFSD" then
					FieldList = FieldList & f(i) & ","
					if f(i)<>"TestSiteDetails" then
						if isdate(rs(f(i)))= true then
							Response.Write "<td align=right valign=top>" & formatdate(rs(f(i))) 
							DataList = DataList & "'" & formatdate(rs(f(i))) & "',"
						else
					
							if f(i)="TotalStudyCost" then
								Response.Write "<td align=right valign=top>"
								if isnull(rs(f(i)))=false then
									
									Response.Write "&pound;" & formatnumber(rs(f(i)),2)
									DataList = DataList & formatnumber(rs(f(i)),2) & ","
								else
									DataList = DataList & "Null,"
								end if
							else
								Response.Write "<td valign=top>"
								if f(i)="OAStatus" then
									Response.Write DisplayOAStatus(rs(f(i)) & "")
									DataList = DataList & "'" & DisplayOAStatus(rs(f(i))) & "',"
								else
									Response.Write rs(f(i)) & ""
									if instr(f(i),"Date")=0 then
										DataList = DataList & "'" & replace(rs(f(i)) & "","'","''") & "',"
									else	
										DataList = DataList & "NULL,"
									end if
								end if
							end if
					
						end if
					else
						'TEST SITE DETAILS
						Response.Write "<td align=right valign=top>"
						DataList = DataList & "'',"
						Response.Write "</td>"
					end if
				else
					Response.Write "<td align=right valign=top>"
					if isDate(rs("ReportFinalSignDate"))=true then
						
						d = cint(datediff("d",rs("ReportFinalSignDate"),now()))
						if d>=30 then
							Response.Write "<font color=red>"
						end if	
						Response.Write d
						DataList = DataList & "'" & d & "',"
						if d>=30 then
							Response.Write "</font>"
						end if
					else
						Response.Write "NA"
						DataList = DataList & "'',"
					end if
				end if
				Response.Write "</td>"
			next
			DataList = left(DataList,len(DataList)-1)
			FieldList = left(FieldList,len(FieldList)-1)
			strsql = "INSERT INTO Studies (" & FieldList & ") VALUES (" & DataList & ")"
			'Response.Write "<hr>" & strsql & "<hr>"
			rsExcel.Open strSql, dbExcel, 1,3
			
			
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
	strPath = replace(strPath,"c:\Websites\SEMS\wwwroot\","")
	strpath = replace(strPath,"\","/")
	Response.Write " (<a target=_blank href=" & strPath & ">Download As Excel</a>)"


%>
<!-- #include file = "includes\footer.asp"-->
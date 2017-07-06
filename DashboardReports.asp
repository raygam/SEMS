<%@ Language=VBScript %>
<!-- #include file = "includes\CortexConfig.asp"-->
<%
set rs2 = Server.CreateObject("ADODB.Recordset")
Response.Write "<script type='text/JavaScript' src='scw.js'></script>"
'StartDate = Request.Form("StartDate")
'EndDate = Request.Form("EndDate")

CompanyLevel = trim(Request.Form("CompanyLevel"))

if CompanyLevel = "" then
	CompanyLevel = "All"
end if

MembershipCategory = GetVariable("MembershipCategory")
If MembershipCategory = "" then
	MembershipCategory = "All"
end if

BFAContact = GetVariable("BFAContact")
if BFAContact = "" then
	BFAContact = "All"
end if

MembershipType = GetVariable("MembershipType")
if MembershipType = "" then
	MembershipType = "BFA"
end if

DirectDebit = GetVariable("DirectDebit")
if DirectDebit = "" then
	DirectDebit = "All"
end if
SortBy = GetVariable("SortBy")
if SortBy = "" then
	SortBy = "CompanyName"
end if

SortOrder = GetVariable("SortOrder")
If SortOrder = "" then
	select case SortBy
	case "CompanyName"
		SortOrder = "ASC"
	case "AmountPaid", "DirectDebit"
		SortOrder = "DESC"	
	case else
		SortOrder = "ASC"
	end select
else
	if SortOrder = "DESC" then
		SortOrder = "ASC"
	else
		SortOrder = "DESC"
	end if		
end if

ReportFilter = replace(GetVariable("ReportFilter"),"-"," ")
if ReportFilter = "" then
	ReportFilter = "Active Memberships"
end if

'if StartDate="" then
'	StartDate = "01/" & month(dateadd("m",-1,now())) & "/" & year(dateadd("m",-1,now()))
'	EndDate = "01/" & month(now()) & "/" & year(now())
'	EndDate = DateAdd("d",-1,EndDate)
'end if

'if IsDate(StartDate)=false then
'	msg = msg & "The Start Date is not valid. "
'	EndDate = ""
'end if	
'if IsDate(EndDate)=false then
'	msg = msg & "The End Date is not valid. "
'	EndDate=""
'end if


Response.Write "<blockquote>"
Response.Write "<h1>Reports</h1>"

Response.Write "<form method=post action=DashboardReports.asp>"
Response.Write "<table border=0>"
'Response.Write "<tr><td valign=top>Start Date</td>"
'Response.Write "<td valign=top><input id=""StartDate"" name=""StartDate"" type=""text"" tabindex=""130"" value=""" & StartDate & """ />"
'Response.Write "<img src=""images/scw.gif"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('StartDate'),event);"" /></td></tr>"

'Response.Write "<tr><td valign=top>End Date</td>"
'Response.Write "<td valign=top><input  id=""EndDate"" name=""EndDate"" type=""text"" tabindex=""131"" value=""" & EndDate & """ />"
'Response.Write "<img src=""images/scw.gif"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('EndDate'),event);"" /></td></tr>"

Response.Write "<tr><td valign=top>Report Type</td>"
Response.Write "<td valign=top>"
Response.Write SimpleComboHTML("ReportFilter", ReportFilter, "All;Active Memberships;Upcoming Renewals;Recently Expired;All Expired;Renewals This Month;Recent Renewals", "", "")
Response.Write "</td></tr>"

Response.Write "<tr><td valign=top>BFA Contact </td>"
Response.Write "<td valign=top><Select name=BFAContact><option value=0>All</option>"

ReportOn = "All BFA Contacts"

rs.open "SELECT * FROM Users WHERE UserID <> 1",db

do until rs.eof=true
	Response.Write "<option value=" & rs("userid")
	if trim(BFAContact) = trim(rs("UserID")) then
		Response.Write " selected"
		ReportOn = rs("UserFirstname") & " " & rs("UserSurname")
	end if
	
	Response.Write ">" & rs("UserFirstName") & " " & rs("UserSurname") & "</option>"
	rs.movenext
loop
rs.close
Response.Write "</td>"



Response.Write "<td valign=top>Membership Type</td>"
Response.Write "<td valign=top>"
Response.Write SimpleComboHTML("MembershipType",MembershipType,"All;BFA;IOPF","","") 
Response.Write "</td>"

strsql = "SELECT MembershipCategoryCode, MembershipCategory FROM MembershipCategories "
if MembershipType <> "All" then
	strsql = strsql & " WHERE MembershipType = '" & MembershipType & "'"
end if
rs.open strsql, db
strMembershipTypes = "All"
do until rs.eof=true
	strMembershipTypes = strMembershipTypes & ";" & rs("MembershipCategoryCode") 
	rs.movenext
loop
rs.close

Response.Write "<td valign=top>&nbsp;&nbsp;&nbsp;&nbsp;Membership Category</td>"
Response.Write "<td valign=top>"
Response.Write SimpleComboHTML("MembershipCategory", MembershipCategory, strMembershipTypes, "", "")
Response.Write "</td>"

Response.Write "<td valign=top>&nbsp;&nbsp;&nbsp;&nbsp;Direct Debit</td>"
Response.Write "<td valign=top>"
Response.Write SimpleComboHTML("DirectDebit", DirectDebit, "All;No;Yes", "", "")
Response.Write "</td>"

Response.Write "</td><td>&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit name=b1 value=""Run Report""></td></tr>"

Response.Write "</table>"
Response.Write "</form>"


strDate = "AND DateStamp >= '" & day(StartDate) & "/" & monthname(month(StartDate)) & "/" & year(StartDate) & "' AND DateStamp <= '" & day(EndDate) & "/" & monthname(month(EndDate)) & "/" & year(EndDate) & "'"



'Response.Write "<br><table border=0 cellspacing=0 cellpadding=6><tr><td><b>Report Date:</b> " & day(Date) & left(monthname(month(Date)),3) & right(year(Date),2) & "</td>"	
'Response.Write "<td><b>Period Covered by Report:</b> " & day(StartDate) & left(monthname(month(startdate)),3) & right(year(startdate),2) & "-" & day(EndDate) & left(monthname(month(EndDate)),3) & right(year(EndDate),2)  & "</td></tr>"
'Response.Write "<tr><td><b>BFA Contacts Covered:</b> " & ReportOn & "</td>"
'Response.Write "<td><b>Report By:</b> " 
'rs.open "SELECT * FROM Users WHERE UserID = " & Request.Cookies("CRMLoginID"),db
'Response.Write rs("UserFirstname") & " " & rs("UserSurname") & "</td></tr></table>"
'rs.close
 
t = "<a href=FullmembershipList.asp?ReportFilter=" & replace(Reportfilter," ","-") & "&MembershipCategory=" & MembershipCategory & "&MembershipType=" & MembershipType & "&DirectDebit=" & DirectDebit & "&BFAContact=" & BFAContact & "&SortOrder=" & SortOrder & "&"
	
Response.Write "<h1>BFA Membership"
if ReportFilter > "" then
	Response.Write " - " & ReportFilter & "&nbsp;&nbsp;<font size=2>" & replace(t,"href=FullmembershipList.asp","target=_blank href=BFAExcel.asp") & "SortBy=" & SortBy & ">Export to Excel</a></font>"
end if
Response.Write "</h1>"
Response.Write "<table width=100% border=0 cellspacing=0 cellpadding=6>"
	Response.write "<tr>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=CompanyNumber>Number</a></b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Expiry>Expiry</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=LastRenewalDate>Last Renewal Date</b></td>"
	
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Category>Category</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=ContactDetails>Contact</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=CompanyName>Company Name</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Address1>Address1</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Address2>Address2</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Town>Town</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=County>County</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Postcode>Postcode</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=AmountPaid>Paid</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=DirectDebit>DD</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Telephone>Telephone</b></td>"
	Response.Write "<td valign=top><font face=arial size=2><b>" & t & "SortBy=Email>Email</b></td>"
	if BFAContact<>"0" then
		Response.Write "<td valign=top><font faoe=arial size=2><b>" & t & "SortBy=UserID>BFA Contact</b></td>"
	end if
	Response.Write "</tr>"

	
	strsql = "SELECT * FROM CompaniesView WHERE CompanyID > 0 " 
	
	If ReportFilter <> "All" then
		'Response.Write "<font color=red>" & ReportFilter & "</font>"
		select case ReportFilter
		case "Active Memberships"
			strsql = strsql & " AND (Expiry > '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "')"
		case "Upcoming Renewals"
			d = dateadd("d",60,Date)
			strsql = strsql & " AND (Expiry > '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "' AND Expiry <= '" & day(d) & "/" & monthname(month(d)) & "/" & year(d) & "')"
		case "Recently Expired"
			d = dateadd("d",-60,Date)
			strsql = strsql & " AND (Expiry < '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "' AND Expiry >= '" & day(d) & "/" & monthname(month(d)) & "/" & year(d) & "')"
		case "All Expired"
			strsql = strsql & " AND (Expiry IS NULL OR Expiry < '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "')"
		case "Renewals This Month"
			d = "01/" & monthname(month(now())) & "/" & year(now())
			strsql = strsql & " AND (LastRenewalDate <= '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "' AND LastRenewalDate >= '" & day(d) & "/" & monthname(month(d)) & "/" & year(d) & "')"
	
		case "Recent Renewals"
			d = dateadd("d",-60,Date)
			strsql = strsql & " AND (LastRenewalDate <= '" & day(now()) & "/" & monthname(month(now())) & "/" & year(now()) & "' AND LastRenewalDate >= '" & day(d) & "/" & monthname(month(d)) & "/" & year(d) & "')"
	
		end select
	end if
	if BFAContact <> "0" AND BFAContact <> "All" then
		strsql = strsql & " AND UserID = " & BFAContact
	end if
	
	if MembershipType <> "All" then
		strsql = strsql & " AND MembershipType = '" & MembershipType & "'"
	end if
	
	if MembershipCategory <> "All" then
		strsql = strsql & " AND Category = '" & MembershipCategory & "'"
	end if
	
	if DirectDebit <> "All" then
		strsql = strsql & " AND DirectDebit = '" & DirectDebit & "'"
	end if
	strsql = strsql & " ORDER BY " & SortBy & " " & SortOrder
	'Response.Write strsql & "<br>"
	
	c=0
	rs.open strSql,db
	if rs.eof=false or rs.bof=false then
		do until rs.eof=true
			Response.Write "<tr"
			if c/2 = int(c/2) then
				Response.Write " bgcolor=#E1E1FF"
			end if
			Response.Write ">"
			Response.Write "<td valign=top>" & rs("CompanyNumber") & "</td>"
			
			Response.Write "<td valign=top>" 
			if isdate(rs("Expiry") & "") = true then
				Response.write right("0" & day(rs("Expiry")),2) & "/" & left(monthname(month(rs("Expiry"))),3) & "/" & year(rs("Expiry"))
			end if
			Response.Write "</td>"
			
			Response.Write "<td valign=top>" 
			if isdate(rs("LastRenewalDate") & "") = true then
				Response.write right("0" & day(rs("LastRenewalDate")),2) & "/" & left(monthname(month(rs("LastRenewalDate"))),3) & "/" & year(rs("LastRenewalDate"))
			end if
			Response.Write "</td>"
			
			Response.Write "<td valign=top>" & rs("Category") & "</td>"
			Response.Write "<td valign=top>" & rs("ContactDetails") & "</td>"
			Response.Write "<td valign=top>" & rs("CompanyName") & "</td>"
			Response.Write "<td valign=top>" & rs("Address1") & "</td>"
			Response.Write "<td valign=top>" & rs("Address2") & "</td>"
			Response.Write "<td valign=top>" & rs("Town") & "</td>"
			Response.Write "<td valign=top>" & rs("County") & "</td>"
			Response.Write "<td valign=top>" & rs("Postcode") & "</td>"
			Response.Write "<td valign=top>" 
			if isnumeric(rs("AmountPaid"))=true then
				Response.write "&pound;" & formatnumber(rs("AmountPaid"),2)
			end if
			Response.Write "</td>"
			Response.Write "<td valign=top>" & rs("DirectDebit") & "</td>"
			Response.Write "<td valign=top>" & rs("Telephone") & "</td>"
			Response.Write "<td valign=top>" & rs("Email") & "</td>"
			
			
			
			
			if BFAContact<>"0" then
				Response.Write "<td valign=top>"
				Response.Write rs("UserFirstName") & " " & rs("UserSurname") 
				Response.Write "</td>"
			end if
			
			Response.Write "<td><a href=""ViewCompany.asp?Key=" & rs("WebKey") & """>View</a></td>"
			Response.Write "<td><a href=""EditCompanies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
			
			
			
			
			
			Response.Write "</tr>"
			c=c+1
			rs.movenext
		loop
	end if
	rs.close

	

Response.Write "</table>"
Response.Write "<br/><hr/>Records Found: " & C
Response.Write "</blockquote>"

function GetVariable(s)
	dim t
	s = trim(s)
	t = Request.Form(s)
	if t ="" then
		t = Request.QueryString(s)
	end if
	GetVariable = t
end function
%>
<!-- #include file = "includes\bottom.asp"-->

<!-- #include file = "includes\CortexConfig.asp"-->
<%
	AllRecords = trim(Request.QueryString("AllRecords"))
	if AllRecords="" then
		AllRecords = trim(Request.Form("AllRecords"))
	end if
	ShowInactive = trim(Request.QueryString("ShowInactive"))
	if ShowInactive="" then
		ShowInactive = trim(Request.Form("ShowInactive"))
	end if
	if Request.Form("ClientEdit")>"" then
		MyStudies = trim(Request.Form("MyStudies"))
		session("MyStudies")=MyStudies
	else
		MyStudies = session("MyStudies")
	end if
	
	OriginalSearch = Request.Form("Search")
	Search = replace(Request.Form("Search"),"'","''")
	if search = "" then
		search = replace(Request.QueryString("Search"),"~"," ")
		OriginalSearch = Search 
		search = replace(search,"'","''")
	end if
	if Request.Form("Search")="" then
		if Request.form("b1")="Search" or Request.QueryString("b1") = "Search" then
			search = ""
			session("SearchString")=""
		else
			search = session("SearchString")
			OriginalSearch = Search 
		end if 
	else
		session("SearchString") = search	
	end if
	ListLength = Request.Form("ListLength")
	if ListLength = "" then
		
		if Request.Cookies("ListLength")>"" then
			ListLength = Request.Cookies("ListLength")
			
		else
			ListLength = "20"
		end if
	end if
	Response.Cookies("ListLength") = ListLength
	
	d = formatdate(dateadd("y",1,now())) 
	d = month(d) & "/" & day(d) & "/" & year(d)
	
	Response.Cookies("ListLength").Expires = d
	
	if Request.Form("ClientEdit")="" and Request.querystring("Clientedit")="" then
		strsql = "SELECT * FROM Studies WHERE OAStatus = 'L' AND ArchiveActDate IS NULL AND StudyDirector = '" & replace(trim(session("LoggedInName")),"'","''") & "'"
		rs.open strsql,db
		if rs.eof=false or rs.bof=false then 
			MyStudies = "ON"
		else
			MyStudies = ""
		end if
		rs.close
	end if
%>

<blockquote>
<form method="post" action="studies.asp" id=form1 name=form1><table cellspacing="0" cellpadding="5"><tr><td valign="top"><h1>Studies</h1>&nbsp; &nbsp;</td><td valign=top></td>
<td align="right" valign="top"><a href=Studies.asp?b1=Search>Clear Search</a>&nbsp; &nbsp;


&nbsp;&nbsp;<input type="Text" name="Search" value="<%=OriginalSearch%>">
<%

'response.write "<input type=""checkbox"" name=""AllRecords"" value=""ON"""
'if AllRecords > "" then
'	Response.Write " checked"
'end if 
'Response.Write ">&nbsp;All Companies &nbsp;"

response.write "<input type=""checkbox"" name=""ShowInactive"" value=""ON"""
if ShowInactive > "" then
	Response.Write " checked"
end if 
Response.Write ">&nbsp;Show Non-Live Records &nbsp;&nbsp;"

response.write "<input type=""checkbox"" name=""MyStudies"" value=""ON"""
if MyStudies > "" then
	Response.Write " checked"
end if 
Response.Write ">&nbsp;My Studies Only &nbsp;&nbsp;"

Response.Write "&nbsp;Records per page&nbsp;&nbsp;"
s = "20;50;100"
Response.Write "<select name=ListLength>"
LL = split(s,";")
for i = 0 to ubound(LL)
	Response.Write "<option"
	if ListLength = LL(i) then
		Response.Write " selected"
	end if
	Response.Write ">" & LL(i) & "</option>"	
next
Response.Write "</select>"

%>

<input type="Submit" name="b1" value="Search"><input type=hidden name=ClientEdit value=YES></td></tr></table></form>
<%
if Request.Form("ClientEdit")="" then
	
else

end if	
	if MyStudies > "" then
		strsql = "SELECT * FROM Studies WHERE StudyDirector = '" & replace(trim(session("LoggedInName")),"'","''") & "' AND "
	else
		strsql = "SELECT * FROM Studies WHERE StudyID > 0 AND "
	end if

	
	if Search > "" then
		s = split(Search, " ")
		'strsql = strsql & " WHERE "
		for i = 0 to ubound(s)
			strsql = strsql & " (StudyNumber LIKE '%" & s(i) & "%' OR StudyDescription LIKE '%" & s(i) & "%' OR StudyDirector LIKE '%" & s(i) & "%' OR ClientName LIKE '%" & s(i) & "%' OR FinancialClientName LIKE '%" & s(i) & "%' OR TestSubstanceName LIKE '%" & s(i) & "%'  OR ProjectManagementCode LIKE '%" & s(i) & "%' OR Department LIKE '%" & s(i) & "%') AND"
		next
		
		
	end if
	strsql = left(strsql, len(strsql)-4)
	'if AllRecords = "" then
		
	'	strsql = strsql & " AND UserID = " & Request.Cookies("CRMLoginID")
	'end if
	If ShowInactive = "" then
		strsql = strsql & " AND OAStatus = 'L'"
		strsql = strsql & " AND ArchiveActDate IS NULL"
		
	'else
	'	strsql = strsql & " AND InActive IS NOT NULL"
	end if
	strsql = strsql & " ORDER BY DateStamp DESC"
	'Response.Write strsql & "<br>"
	DisplayFields = "StudyNumberNew;ClientName;TestSubstanceName;StudyDescription;StudyType;StudyDirector;Department;OAStatus;TotalStudyCost;ProtocolIssuedDate;ActualExperimentalStartDate;ActualExperimentalEndDate;ActualDraftReportDate;LastUpdate"

	PageSize = cint(ListLength)
	rs.cursorlocation = 3
	rs.cursortype = 3
	rs.pagesize = PageSize
	rs.cachesize = PageSize
	Page = trim(Request.QueryString("Page"))
	'Response.Write " - " & Page & "<hr>"
	If Page = "" then
		Page = 1	
	end if	 

	
		
		
	
	search = replace(Search," ","~")
	'Response.Write strsql & "<br>"
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		response.write "<center>"
		if cint(Page) > rs.PageCount or cint(Page) < 1 then
			Page = 1
		end if	
		
		if cint(Page)> 1 then
			Response.Write "<a href=studies.asp?Page=" & cint(Page)-1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">< Previous Page</a>&nbsp;"
		else
			Response.Write "<font color=silver><u>< Previous Page</u></font>&nbsp;"
		end if
		startpage = page - 10
		if startpage < 1 then
			startpage = 1
		end if
		endpage = page + 10
		if endpage > rs.pagecount then
			endpage = rs.pagecount
		end if
		for i = startpage to endpage
			if cint(Page) = cint(i) then
				Response.Write "<b>" & i & "</b>&nbsp;"
			else
				Response.Write "<a href=studies.asp?Page=" & i & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">" & i & "</a>&nbsp;"
			end if
				
		next
		if cint(Page)< rs.PageCount then
			Response.Write "&nbsp;<a href=studies.asp?Page=" & cint(Page)+1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & "> Next Page ></a>"
		else
			Response.Write "&nbsp;<font color=silver><u>Next Page ></u></font>"	
		end if
		Response.Write "</center><br>"
		rs.absolutepage = cint(Page)
	
		f = split(DisplayFields,";")
		Response.Write "<table cellspacing=0 cellpadding=3 border=0><tr>"
		Response.Write "<td><b>Study Number</b></td>"
		Response.Write "<td><b>Client Name</b></td>"
		Response.Write "<td><b>Test Substance</b></td>"
		Response.Write "<td><b>Study Description</b></td>"
		Response.Write "<td><b>Study Type</b></td>"
		Response.Write "<td><b>Study Director</b></td>"
		Response.Write "<td><b>Department</b></td>"
		Response.Write "<td><b>Open Acc Status</b></td>"
	
		Response.Write "<td><b>Study Cost</b></td>"
		Response.Write "<td><b>Protocol Issued Date</b></td>"
		Response.Write "<td><b>Act Exp Start Date</b></td>"
		Response.Write "<td><b>Act Exp End Date</b></td>"
		Response.Write "<td><b>Act Audited Draft Report</b></td>"
		Response.Write "<td><b>Last Updated</b></td>"
		
		Response.Write "</tr>"
		
			
		c=0
		do until rs.eof=true or c=pagesize
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
							Response.Write DisplayCurrency(rs("StudyCurrency") & "")
							Response.Write formatnumber(rs(f(i)),2)
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
				'if i=0 and rs("InActive") & "" >"" then
				'	Response.Write "&nbsp;<font color=red>(Inactive)" 
				'end if
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
	else
		Response.Write "Sorry, no matching data found."
	end if
	
	rs.close



%>
<!-- #include file = "includes\footer.asp"-->
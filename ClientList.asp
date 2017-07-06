
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
	'Response.Write "(" & Request.QueryString("Action") & " - " & Request.QueryString("WebKey") & ")<br>"
	if trim(Request.QueryString("Action"))="D" then
		if Request.QueryString("WebKey") > "" then
			t = Request.QueryString("WebKey")
			t = replace(t, "'","")
			t = replace(t, chr(34),"")
			
			strsql = "DELETE FROM ClientLookup WHERE WebKey = '" & t & "'"
			'Response.Write strsql
			rs.open strsql, db,1,3
	
		end if
	end if
	
%>
<blockquote><table width="100%" cellspacing="0" cellpadding="5"><tr><td valign="top"><h1>Client Name List</h1>&nbsp;&nbsp;<a href=EditClientList.asp>Add Client Name</a></td><td valign=top></td>
<td align="right" valign="top"><form method="post" action="ClientList.asp"><a href=Studies.asp?b1=Search>Clear Search</a>&nbsp;&nbsp;<input type="Text" name="Search" value="<%=OriginalSearch%>">
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
Response.Write ">&nbsp;Show Non-Live Records &nbsp;"

%>
<input type="Submit" name="b1" value="Search"></form></td></tr></table>
<%
if Request.Form("ClientEdit")="" then
	strsql = "SELECT * FROM ClientLookup WHERE ClientLookupID > 0 AND"
	if Search > "" then
		s = split(Search, " ")
		'strsql = strsql & " WHERE "
		for i = 0 to ubound(s)
			'strsql = strsql & " (StudyNumber LIKE '%" & s(i) & "%' OR StudyTitle LIKE '%" & s(i) & "%' OR StudyDirector LIKE '%" & s(i) & "%' OR ClientName LIKE '%" & s(i) & "%' OR FinancialClientName LIKE '%" & s(i) & "%' OR TestSubstanceName LIKE '%" & s(i) & "%') AND"
			strsql = strsql & " ClientName LIKE '" & s(i) & "%' AND"
		next
		
		
	end if
	strsql = left(strsql, len(strsql)-4)
	'if AllRecords = "" then
		
	'	strsql = strsql & " AND UserID = " & Request.Cookies("CRMLoginID")
	'end if
	If ShowInactive = "" then
		'strsql = strsql & " AND OAStatus = 'L'"
	'else
	'	strsql = strsql & " AND InActive IS NOT NULL"
	end if
	strsql = strsql & " ORDER BY ClientName ASC"
	'Response.Write strsql & "<br>"
	DisplayFields = "ClientName"
else

end if
	PageSize = 20
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
			Response.Write "<a href=ClientList.asp?Page=" & cint(Page)-1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">< Previous Page</a>&nbsp;"
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
				Response.Write "<a href=ClientList.asp?Page=" & i & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">" & i & "</a>&nbsp;"
			end if
				
		next
		if cint(Page)< rs.PageCount then
			Response.Write "&nbsp;<a href=ClientList.asp?Page=" & cint(Page)+1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & "> Next Page ></a>"
		else
			Response.Write "&nbsp;<font color=silver><u>Next Page ></u></font>"	
		end if
		Response.Write "</center><br>"
		rs.absolutepage = cint(Page)
	
		f = split(DisplayFields,";")
		Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
		Response.Write "<td><b>Client Name</b></td>"
		'Response.Write "<td><b>Financial Client Name</b></td>"
		'Response.Write "<td><b>Test Substance</b></td>"
		'Response.Write "<td><b>Study Type</b></td>"
		'Response.Write "<td><b>Study Director</b></td>"
		'Response.Write "<td><b>Department</b></td>"
		'Response.Write "<td><b>OA Status</b></td>"
		'Response.Write "<td><b>Study Cost</b></td>"
		'Response.Write "<td><b>Protocol Init Date</b></td>"
		'Response.Write "<td><b>Est Exp Start Date</b></td>"
		'Response.Write "<td><b>Est Exp End Date</b></td>"
		
		
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
				'if i=0 and rs("InActive") & "" >"" then
				'	Response.Write "&nbsp;<font color=red>(Inactive)" 
				'end if
				Response.Write "</td>"
			next
			
			'STRATEGIC PARTNER
			
			'Response.Write "<td valign=top><a href=""ViewStudy.asp?Key=" & rs("WebKey") & """>View</a></td>"
			Response.Write "<td valign=top><a href=""ClientList.asp?Action=D&WebKey=" & rs("WebKey") & """>Delete</a></td>"
			
			Response.Write "<td valign=top><a href=""EditClientList.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
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
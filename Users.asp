
<!-- #include file = "includes\CortexConfig.asp"-->
<%
	if cstr(session("SystemAdmin")) <> "-1" then
		Response.Redirect "Studies.asp"
	end if
	OriginalSearch = Request.Form("Search")
	Search = replace(Request.Form("Search"),"'","''")
	if search = "" then
		
		search = replace(Request.QueryString("Search"),"~"," ")
		if Request.Form("b1")<>"Search" and Request.QueryString("Action")<>"Clear" then
			search = session("Search")
		end if
		OriginalSearch = Search 
		search = replace(search,"'","''")
	else
		session("Search")=Search	
	end if
	AllRecords = trim(Request.QueryString("AllRecords"))
	if AllRecords="" then
		AllRecords = trim(Request.Form("AllRecords"))
	end if
	ShowInactive = trim(Request.QueryString("ShowInactive"))
	if ShowInactive="" then
		ShowInactive = trim(Request.Form("ShowInactive"))
	end if
%>
<blockquote><table width="95%" cellspacing="0" cellpadding="5"><tr><td valign="top"><h1>Users</h1></td>
<td><a href=EditUsers.asp>Add User</a></td>
<td align="right" valign="top"><form method="post" action="Users.asp"><a href=Users.asp?Action=Clear>Clear Search</a>&nbsp;&nbsp;<input type="Text" name="Search" value="<%=OriginalSearch%>">
<%

response.write "&nbsp;<input type=""checkbox"" name=""ShowInactive"" value=""ON"""
if ShowInactive > "" then
	Response.Write " checked"
end if 
Response.Write ">&nbsp;Hide Inactive Users &nbsp;"

%>
<input type="Submit" name="b1" value="Search"></form></td></tr></table>
<%
if Request.Form("ClientEdit")="" then
	strsql = "SELECT * FROM Users WHERE UserID > 0 AND "
	if Search > "" then
		s = split(Search, " ")
		'strsql = strsql & " AND Surname > 'A' AND Surname IS NOT NULL AND "
		for i = 0 to ubound(s)
			strsql = strsql & " (Surname LIKE '%" & s(i) & "%' OR Firstname LIKE '%" & s(i) & "%' OR Department LIKE '%" & s(i) & "%' OR Email LIKE '%" & s(i) & "%') AND"
		next
		
	end if
	strsql = left(strsql, len(strsql)-4)
	'if AllRecords = "" then
	'	strsql = strsql & " AND UserID = " & Request.Cookies("CRMLoginID")
	'end if
	If ShowInactive = "" then
		
	else
		strsql = strsql & " AND IsActive = -1 "
	end if
	
	strsql = strsql & " ORDER BY Surname ASC"
	'Response.Write strsql & "<br>"
	DisplayFields = "FirstName;Surname;OADepartment;Email;SystemAdmin;Archivist;BD;QA;DepartmentManager;SponsorProjectManager;TFM;IsActive" 
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
	
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		response.write "<center>"
		if cint(Page) > rs.PageCount or cint(Page) < 1 then
			Page = 1
		end if	
		
		if cint(Page)> 1 then
			Response.Write "<a href=User.asp?Page=" & cint(Page)-1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">< Previous Page</a>&nbsp;"
		else
			Response.Write "<font color=silver><u>< Previous Page</u></font>&nbsp;"
		end if
		for i = 1 to rs.PageCount
			if cint(Page) = cint(i) then
				Response.Write "<b>" & i & "</b>&nbsp;"
			else
				Response.Write "<a href=Users.asp?Page=" & i & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & ">" & i & "</a>&nbsp;"
			end if
				
		next
		if cint(Page)< rs.PageCount then
			Response.Write "&nbsp;<a href=Users.asp?Page=" & cint(Page)+1 & "&Search=" & Search & "&AllRecords=" & AllRecords & "&ShowInactive=" & ShowInactive & "> Next Page ></a>"
		else
			Response.Write "&nbsp;<font color=silver><u>Next Page ></u></font>"	
		end if
		Response.Write "</center><br>"
		rs.absolutepage = cint(Page)
	
		f = split(DisplayFields,";")
		Response.Write "<table width=""95%"" cellspacing=0 cellpadding=3 border=0><tr>"
		
		Response.Write "<td><b>Firstname</b></td>"
		Response.Write "<td><b>Surname</b></td>"
		Response.Write "<td><b>Department</b></td>"
		Response.Write "<td><b>Email</b></td>"
		Response.Write "<td align=center><b>System Admin</b></td>"
		Response.Write "<td align=center><b>Archivist</b></td>"
		Response.Write "<td align=center><b>BD</b></td>"
		Response.Write "<td align=center><b>QA</b></td>"
		Response.Write "<td align=center><b>Department Manager</b></td>"
		Response.Write "<td align=center><b>Sponsor Project Manager</b></td>"
		Response.Write "<td align=center><b>TFM</b></td>"
		
		Response.Write "<td align=center><b>Active</b></td></tr>"
		
		
			
		c=0
		do until rs.eof=true or c=pagesize
			Response.Write "<tr"
			if c/2 = int(c/2) then
				'Response.Write " bgcolor=#E1E1FF"
				Response.Write " bgcolor=#FCF3D7"
			
			end if
			Response.Write ">"
			for i = 0 to ubound(f)
				'Response.Write "<font face=red>" & f(i) & "</font>"
				if i>=4 and i<=11 then
					Response.Write "<td valign=top align=center>" 
				else
					Response.Write "<td valign=top>" 
				end if
				if c/2 <> int(c/2) then
					Response.Write "<font color=""navy"">"
				end if
				if i = 3 and f(i) & "" >"" then
					Response.Write "<a href=""mailto:" & rs(f(i)) & """>"  
				end if
				
				if i>=4 and i<=11 then
					if trim(rs(f(i)) & "") = "-1" then
						Response.Write "<img border=0 src=""images/tick.png"">"
					end if
				else
					Response.write rs(f(i))
				end if
				
				if i = 3 and f(i) & "" >"" then
					Response.Write "</a>"
				end if
				if i=2 then
					if rs("IsActive") & "" = "0" then
						Response.Write "&nbsp;<font color=red>(Inactive)</font>"
					end if
				end if
				Response.Write "</font></td>"
			next
			'Response.Write "<td valign=top><a href=""ViewCompany.asp?Key=" & rs("WebKey") & """>View</a></td>"
			'Response.Write "<td valign=top><a href=""EditCompanies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			Response.Write "<td valign=top><a href=""EditUsers.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
			
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
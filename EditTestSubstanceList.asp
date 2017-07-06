<%@ Language=VBScript %>
<%

%>
<!-- #include file = "includes\CortexConfig.asp"-->

<%




if trim(Request.QueryString("Key"))="" then

	path = left( monthname(month(now())) ,3) & year(now())
	
	WebKey = Request.Form("WebKey")
	TestSubstanceName = Request.Form("TestSubstanceName")
	
	
	
	
	
	IsActive = Request.Form("IsActive")
	if IsActive = "" then
		IsActive = "0"
	end if
	ClientEdit = Request.Form("ClientEdit")
	Email = Request.Form("Email")
	b1 = Request.Form("b1")
	

	if b1 ="Cancel" then
		Response.Redirect "TestSubstanceList.asp"
	end if
	
		
	msg = ""
	if ClientEdit="" then
		UserID = Session("LoggedIn")
	end if
	
	if TestSubstanceName = "" then
		msg = "Please enter a Test Substance Name. "
	end if
	
	if msg = "" then
		'OK to Save
		strsql = "SELECT * FROM TestSubstances WHERE WebKey = '" & WebKey & "'"
		rs.open strsql,db,1,3
		if WebKey = "" then
			rs.addnew
			WebKey = CreateWebKey()
			rs("WebKey")= WebKey
			rs("TestSubstanceName") = WebKey
			rs.update
			rs.close
			strsql = "SELECT * FROM TestSubstances WHERE WebKey = '" & WebKey & "'"
			rs.open strsql,db,1,3
			
		end if
		
		
		
		WriteToChangeLog "ChangeLogLarge", rs("TestSubstanceID"), Session("LoggedIn"), "TestSubstances", "TestSubstanceName", rs("TestSubstanceName") & "", TestSubstanceName
		rs("TestSubstanceName") = TestSubstanceName
	
		rs("LastUpdated")=now()
		rs("LastUpdatedBy") = Session("LoggedIn")
		
		rs.update
		rs.close
		
		
		Response.Redirect "TestSubstanceList.asp"
	else
		'Response.Write "Message: " & msg
		'Error - back to form
	end if
else
	WebKey = replace(Request.querystring("Key"),"'","")
	strsql = "SELECT * FROM TestSubstances WHERE WebKey = '" & WebKey & "'"
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		TestSubstanceName = rs("TestSubstanceName") & ""
		
		
		
	else
		Response.Redirect "TestSubstanceList.asp"
	end if
	rs.close
end if





ClientEdit = trim(ClientEdit)
Response.Write "<blockquote><h1>Edit "

If WebKey > "" then
	Response.Write TestSubstanceName 
else
	Response.Write "New Test Substance Name"
end if
Response.Write "</h1>"


if msg>"" and ClientEdit > "" then
	Response.Write "<font color=red><b>" & msg & "</b></font>"
end if

Response.Write "<form method=""Post"" action=""EditTestSubstanceList.asp"">"

Response.Write "<table align=""center"" cellpadding=10 border=0><tr><td valign=top>"

Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""3"">"


Response.Write "<tr><td valign=top>Test Substance Name<font color=red>*</font></td>"
Response.Write "<td valign=top>" & TextBoxHTML("TestSubstanceName",TestSubstanceName,50,100,"") & "</td></tr>"





Response.Write "</td></tr>"




Response.Write "<tr><td colspan=""2""><hr></td></tr>"

Response.Write "<tr><td colspan=""2"" align=""center""><input type=""Submit"" name=""b1"" value=""Save"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""Submit"" name=""b1"" value=""Cancel""><input type=""hidden"" name=""ClientEdit"" value=""Yes""><input type=""hidden"" name=""WebKey"" value=""" & WebKey & """></td></tr>"

Response.Write "</table>"
Response.Write "</form>"


%>
<!-- #include file = "includes\footer.asp"-->
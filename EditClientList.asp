<%@ Language=VBScript %>
<%

%>
<!-- #include file = "includes\CortexConfig.asp"-->

<%




if trim(Request.QueryString("Key"))="" then

	path = left( monthname(month(now())) ,3) & year(now())
	
	WebKey = Request.Form("WebKey")
	ClientName = Request.Form("ClientName")
	
	
	
	
	BD = Request.Form("BD")
	if BD = "" then
		BD = "0"
	end if
	
	IsActive = Request.Form("IsActive")
	if IsActive = "" then
		IsActive = "0"
	end if
	ClientEdit = Request.Form("ClientEdit")
	Email = Request.Form("Email")
	b1 = Request.Form("b1")
	

	if b1 ="Cancel" then
		Response.Redirect "ClientList.asp"
	end if
	
		
	msg = ""
	if ClientEdit="" then
		UserID = Session("LoggedIn")
	end if
	
	if ClientName = "" then
		msg = "Please enter a Firstname and Surname. "
	end if
	
	if msg = "" then
		'OK to Save
		strsql = "SELECT * FROM ClientLookup WHERE WebKey = '" & WebKey & "'"
		rs.open strsql,db,1,3
		if WebKey = "" then
			rs.addnew
			WebKey = CreateWebKey()
			rs("WebKey")= WebKey
			rs.update
			rs.close
			strsql = "SELECT * FROM ClientLookup WHERE WebKey = '" & WebKey & "'"
			rs.open strsql,db,1,3
			
		end if
		
		
		
		WriteToChangeLog "ChangeLogLarge", rs("ClientLookupID"), Session("LoggedIn"), "ClientLookup", "ClientName", rs("ClientName") & "", ClientName
		rs("ClientName") = ClientName
	
		rs("LastUpdated")=now()
		rs("LastUpdatedBy") = Session("LoggedIn")
		
		rs.update
		rs.close
		
		
		Response.Redirect "ClientList.asp"
	else
		'Response.Write "Message: " & msg
		'Error - back to form
	end if
else
	WebKey = replace(Request.querystring("Key"),"'","")
	strsql = "SELECT * FROM ClientLookup WHERE WebKey = '" & WebKey & "'"
	rs.open strsql,db
	if rs.eof=false or rs.bof=false then
		ClientName = rs("ClientName") & ""
		
		
		
	else
		Response.Redirect "Users.asp"
	end if
	rs.close
end if





ClientEdit = trim(ClientEdit)
Response.Write "<blockquote><h1>Edit "

If WebKey > "" then
	Response.Write ClientName 
else
	Response.Write "New Client Name"
end if
Response.Write "</h1>"


if msg>"" and ClientEdit > "" then
	Response.Write "<font color=red><b>" & msg & "</b></font>"
end if

Response.Write "<form method=""Post"" action=""EditClientList.asp"">"

Response.Write "<table align=""center"" cellpadding=10 border=0><tr><td valign=top>"

Response.Write "<table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""3"">"


Response.Write "<tr><td valign=top>Client Name<font color=red>*</font></td>"
Response.Write "<td valign=top>" & TextBoxHTML("ClientName",ClientName,25,40,"") & "</td></tr>"





Response.Write "</td></tr>"




Response.Write "<tr><td colspan=""2""><hr></td></tr>"

Response.Write "<tr><td colspan=""2"" align=""center""><input type=""Submit"" name=""b1"" value=""Save"">&nbsp;&nbsp;&nbsp;&nbsp;<input type=""Submit"" name=""b1"" value=""Cancel""><input type=""hidden"" name=""ClientEdit"" value=""Yes""><input type=""hidden"" name=""WebKey"" value=""" & WebKey & """></td></tr>"

Response.Write "</table>"
Response.Write "</form>"


%>
<!-- #include file = "includes\footer.asp"-->
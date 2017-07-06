<%@ Language=VBScript %>

<!--#Include File="includes\CortexConfig.asp"-->

<%

response.Expires=0
%>
<script type="text/javascript">
function toggleDisplay(divId) {
  var div = document.getElementById(divId);
  div.style.display = (div.style.display=="block" ? "none" : "block");
}
</script>




<%
set rs2 = Server.CreateObject("ADODB.Recordset")
Response.Write "<script type='text/JavaScript' src='scw.js'></script>"
StartDate = Request.Form("StartDate")
EndDate = Request.Form("EndDate")
UserID = trim(Request.Form("UserID"))
TableName = trim(Request.Form("TableName"))
TableNames = "Studies;Users"
if UserID = "" then
	UserID = "-1"
end if	


if StartDate="" then
	StartDate = "01/" & month(dateadd("m",-1,now())) & "/" & year(dateadd("m",-1,now()))
	EndDate = "01/" & month(now()) & "/" & year(now())
	EndDate = DateAdd("d",-1,EndDate)
end if

if IsDate(StartDate)=false then
	msg = msg & "The Start Date is not valid. "
	EndDate = ""
end if	
if IsDate(EndDate)=false then
	msg = msg & "The End Date is not valid. "
	EndDate=""
end if


Response.Write "<blockquote>"
Response.Write "<font color=navy size=4><h1>Auditing</h1></font><br/><br/>"

Response.Write "<form method=post action=Audit.asp id=form1 name=form1>"
Response.Write "<table border=0>"
Response.Write "<tr><td valign=top>Start Date</td>"
Response.Write "<td valign=top><input id=""StartDate"" name=""StartDate"" type=""text"" tabindex=""130"" value=""" & StartDate & """ />"
Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('StartDate'),event);"" /></td></tr>"

Response.Write "<tr><td valign=top>End Date</td>"
Response.Write "<td valign=top><input  id=""EndDate"" name=""EndDate"" type=""text"" tabindex=""131"" value=""" & EndDate & """ />"
Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('EndDate'),event);"" /></td></tr>"

Response.Write "<tr><td valign=top>User</td>"
Response.Write "<td valign=top><Select name=UserID><option value=-1>All Users</option>"

ReportOn = "All Users"

rs.open "SELECT * FROM Users WHERE IsActive = -1 AND UserID <> 1 ORDER BY Firstname ASC",db

do until rs.eof=true
	Response.Write "<option value=" & rs("userid")
	if trim(UserID) = trim(rs("UserID")) then
		Response.Write " selected"
		ReportOn = rs("Firstname") & " " & rs("Surname")
	end if
	
	Response.Write ">" & rs("FirstName") & " " & rs("Surname") & "</option>"
	rs.movenext
loop
rs.close
Response.Write "</td><td><input type=submit name=b1 value=""Run Report""></tr>"

Response.Write "</table>"
Response.Write "</form>"

strSQL = "SELECT ( "
strDate = "AND ActivityDate >= '" & day(StartDate) & "/" & monthname(month(StartDate)) & "/" & year(StartDate) & "' AND ActivityDate <= '" & day(EndDate) & "/" & monthname(month(EndDate)) & "/" & year(EndDate) & "'"



%>


<blockquote>
	
<div>
	

	
		
	
</div>


<br>
	
</div>
<!--#Include File="includes\Footer.asp"-->


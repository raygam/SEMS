<%@ Language=VBScript %>
<!-- #include file = "includes\CortexConfig.asp"-->
<%
Response.Write "<blockquote>"
Response.Write "<h1>SEMS Reports</h1>"
'Response.Write "<a href=Dashboard.asp>Dashboard Report</a> - A range of information to drill into.  <font size=1 color=red>New</font><br><br>"
Response.Write "<a href=DateReport.asp>Date Report</a> - Search for Studies where a key date falls within a date range.  <br><br>"
Response.Write "<a href=ArchivingReport.asp>Archiving Report</a> - Studies with Actual Final Report Issue Date but not archived.  <br><br>"
Response.Write "<a href=ArchivingReport2.asp>Archiving Report Finalised Studies</a> - 'Final' Studies with no Actual Final Report Issue Date and not archived.  <br><br>"

Response.Write "<a href=UpcomingQuotedDraftReportDate.asp>Upcoming Quoted Draft Report Date</a>   <br><br>"
Response.Write "<a href=UpcomingQuotedExperimentalStartDate.asp>Upcoming Quoted Experimental Start Date</a>  <br><br>"

Response.Write "<a href=NewStudies.asp>New Studies</a> - In the last 7 days.  <font size=1 color=red>New</font><br><br>"
Response.Write "<a href=StudyDirectorChanged.asp>Change of Study Director</a> - In the last 7 days.  <font size=1 color=red>New</font><br><br>"
Response.Write "<a href=EstimatedDraftReportDateChanged.asp>Expected Audited Draft to Client Date Changed</a> - In the last 7 days.  <font size=1 color=red>New</font><br><br>"
Response.Write "<a href=QuotedDraftReportDateChanged.asp>Quoted Draft Report Date Changed</a> - In the last 7 days.  <font size=1 color=red>New</font><br><br>"
Response.Write "<a href=ActualDraftReportDateChanged.asp>Actual Audited Draft to Client Date changed</a> - In the last 7 days.  <font size=1 color=red>New</font><br><br>"

Response.Write "<br><br><a href=AuditReport.asp>Audit Report</a> - Report on data changes by date, study, field and user. <font size=1 color=red>New</font><br><br>"

strUser = UCase(trim(Request.ServerVariables("AUTH_USER")))
if strUser = "SSI\HDUNN" or strUser = "SSI\FDAVIES" or strUser = "SSI\DATACRAFTADMIN" or strUser="SSI\KCOCKS" or strUser="SSI\LEARNSHAW" then
	Response.Write "<br><br><a target=_blank href=ExcelExport.aspx>Full Excel Study Export (with Triggers) Live Records Only </a> - Full Study Export in Excel (.xlsx) <font size=1 color=red>New</font><br><br>"
	Response.Write "<br><br><a target=_blank href=ExcelExportFull.aspx>Full Excel Study Export (with Triggers) All Records </a> - Full Study Export in Excel (.xlsx) <font size=1 color=red>New</font><br><br>"
	
end if

if strUser = "SSI\RGAMBARDELLA" or strUser = "SSI\FDAVIES" or strUser = "SSI\SDEAN" or strUser="SSI\MDAWSON" _
	or strUser = "SSI\PREIBACH" or strUser = "SSI\AFOURNIER" or strUser = "SSI\DMITCHELL" or strUser = "SSI\SSWALES" _
	or strUser = "SSI\RBRINHAM" or strUser = "SSI\LEARNSHAW" then
	
	response.write "<br>"
	response.Write "<a href=UpcomingSchedule.asp>Upcoming Schedule (Draft Date)</a>  <font size=1 color=red>New</font> <br><br>"
	response.Write "<a href=UpcomingSchedule_Starts.asp>Upcoming Schedule (Start Date)</a>  <font size=1 color=red>New</font> <br><br>"
	response.Write "<a href=OTD_Dept.asp>On-Time Delivery (By Dept)</a>  <font size=1 color=red>New</font> <br><br>"
	response.Write "<a href=OTS_Dept.asp>On-Time Starts (By Dept)</a>  <font size=1 color=red>New</font> <br><br>"
	response.Write "<a href=ExpectedDraftToQA.asp>Expected Draft To QA</a>  <font size=1 color=red>New</font> <br><br>"
			
end if

Response.Write "<br><br><a href=MasterSchedule.asp>Master Schedule Report</a> - <font size=1 color=red>New</font><br><br>"

Response.Write "</blockquote>"

%>
<!-- #include file = "includes\footer.asp"-->
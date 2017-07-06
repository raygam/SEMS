
<!-- #include file = "includes\CortexConfig.asp"-->


<%

	'strsql = "SELECT * FROM Studies WHERE OAStatus = 'L' AND StudyDirector = '" & session("FullName") & "'"
	strsql = "SELECT * FROM Studies WHERE OAStatus = 'L'"
	strsql = strsql & " ORDER BY StudyNumber ASC"
	'Response.Write strsql & "<br>"
	
	DisplayFields = "Study Number;FinancialClientName;ClientName;Test Substance Name;Study Description;Study Director;StudyType;StudyStatus;RegulatoryStatus;ProjectManagementCode;ProjectManager;Department;Total Study Cost;Protocol Issued Date;ActualExperimentalStartDate;ActualexperimentalEndDate;ActualDraftReportDate;QuotedDraftReportDate;ReportFinalSignDate;ReportFinalisationDeadline;ArchiveActDate"

	'Dim db As New ADODB.Connection
    'Dim rs As New ADODB.Recordset
    'db.Open "Provider=SQLNCLI10;SERVER=SHWTetra;DATABASE=SMVEUCortex;UID=CortexApp;PWD=(#GetiN#)"
    'rs.Open "SELECT * FROM Studies", db
    set dbExcel = server.CreateObject("ADODB.Connection")
    set rsExcel = server.CreateObject("ADODB.Recordset")
    f = "c:\Websites\Cortex\wwwroot\Excel\CortexKeyDataExport.xls"
    
    'f = "C:\Data\Cortex\Backup\Excel\CortexBackup" & Day(Now()) & "-" & MonthName(Month(Now())) & "-" & Year(Now()) & ".xls"
    
    set FSys = server.CreateObject("Scripting.FileSystemObject")
    FSys.CopyFile "C:\Websites\Cortex\wwwroot\Excel\CortexKeyDataTemplate.xls", f, True
    rs.open strsql, db
    dbExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & f & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=0"";"
    rsExcel.Open "[Sheet1$]", dbExcel, 1, 3
    c = 0
    f = split(DisplayFields,";")
    Do Until rs.EOF = True 
        rsExcel.AddNew
        For i = 0 To ubound(f)
			f(i) = replace(f(i)," ","")
			ff = trim(f(i))
			Response.Write i & "- " & ff & "<br>"
            rsExcel(i) = rs(ff) & ""
        Next 
        rsExcel.Update
        rs.MoveNext
        c = c + 1
        
        
    Loop
    rsExcel.Close
    dbExcel.Close
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
	
	Response.Redirect "http://smrweb:172/excel/CortexKeyDataExport.xls"

%>
<!-- #include file = "includes\footer.asp"-->
<!-- #include file = "includes\CortexConfig.asp"-->

<%

FUNCTION IsBlank(Value)
	'returns True if Empty or NULL or Zero
	If IsEmpty(Value) or IsNull(Value) Then
		IsBlank = True
	ElseIf VarType(Value) = vbString Then
		If Value = "" Then
			IsBlank = True
		End If
	ElseIf IsObject(Value) Then
		If Value Is Nothing Then
			IsBlank = True
		End If
	ElseIf IsNumeric(Value) Then
		If Value = 0 Then
			wscript.echo " Zero value found"
			IsBlank = True
		End If
	Else
		IsBlank = False
	End If
END FUNCTION

FUNCTION FlipDateView(dateToFlip)
	DIM strDateView
	DIM day, month, year
	IF IsBlank(dateToFlip) THEN
		strDateView = ""
	ELSE
		day = DatePart("d", dateToFlip)
		month = MonthName(DatePart("m", dateToFlip), True) 
		year = DatePart("yyyy",dateToFlip)

		strDateView = day & " " & Ucase(month) &" " & year
	END IF
	FlipDateView = strDateView
END FUNCTION

FUNCTION MMDDYYYY(dateToFlip)
	DIM day, month, year
      day = DatePart("d", dateToFlip)
	month = DatePart("m", dateToFlip) 
	year = DatePart("yyyy",dateToFlip)

	MMDDYYYY = month & "/" & day & "/" & year
END FUNCTION
	    
' *********************************
' ***	START HERE
' *********************************

DIM strSQL
DIM dept
DIM month
DIM year
DIM qtr
DIM qtr_months
DIM dept_string
DIM rpttype
DIM strHeader
DIM result

rpttype = Request.QueryString("type")

SELECT CASE rpttype
	CASE 1
		strHeader = "On-Time Delivery (By Department) "
	CASE ELSE
		strHeader = "Study List"
END SELECT

IF LEN(Request.QueryString("dept")) = 0 THEN
	dept = ""
ELSE
	dept = TRIM(Request.QueryString("dept"))
	SELECT CASE dept
		CASE "1"
			dept_string = "'Analytical & Physical Chemistry'"
		CASE "2"
			dept_string = "'Ecotoxicology', 'Product Development Testing'"
		CASE "3"
			dept_string =  "'E Fate'"
		CASE "4"
			dept_string = "'Regulatory'"
		CASE ELSE
			dept_string = ""
	END  SELECT
END IF

IF LEN(Request.QueryString("month")) = 0 THEN
	month = ""
ELSE
	month = TRIM(Request.QueryString("month"))
END IF

IF LEN(Request.QueryString("year")) = 0 THEN
	year = ""
ELSE
	year = TRIM(Request.QueryString("year"))
END IF

IF LEN(Request.QueryString("qtr")) = 0 THEN
	qtr = ""
ELSE
	qtr = TRIM(Request.QueryString("qtr"))
	SELECT CASE qtr
		CASE 1
			qtr_months = "1, 2, 3"
		CASE 2
			qtr_months = "4, 5, 6"
		CASE 3 
			qtr_months = "7, 8, 9"
		CASE 4
			qtr_months = "10, 11, 12"
		CASE ELSE
			qtr_months = ""
	END SELECT
END IF

response.write "<h1>&nbsp;&nbsp;" & strHeader & "</h1>"	
response.write "<h4>"
IF LEN(dept_string) > 0 THEN 
	response.write "&nbsp;&nbsp;&nbsp;Dept  :  " & dept_string & "<br>"
END IF
IF LEN(month) > 0 THEN 
	response.write "&nbsp;&nbsp;&nbsp;Month  :  " & MonthName(month, False) & "<br>"
END IF
IF LEN(year) > 0 THEN 
	response.write "&nbsp;&nbsp;&nbsp;Year  :  " & year & "<br>"
END IF
IF LEN(qtr) > 0 THEN 
	response.write "&nbsp;&nbsp;&nbsp;Quarter  :  " & qtr & "<br>"
END IF
response.write "</h4>"



' ***  BUILD SQL STATEMENT
strSQL = "SELECT * FROM Studies WHERE 1=1 AND LEN(StudyNumber) = 7 "

IF LEN(month) > 0 THEN
	strSQL = strSQL & " AND MONTH(QuotedDraftReportDate) = " & month & " " 
END IF

IF LEN(year) > 0 THEN
	strSQL = strSQL & " AND YEAR(QuotedDraftReportDate) = " & year & " " 
END IF
	
IF LEN(qtr_months) > 0 THEN
	strSQL = strSQL & " AND MONTH(QuotedDraftReportDate) IN  (" & qtr_months & ") " 
END IF

IF LEN(dept_string) > 0 THEN
	strSQL = strSQL & " AND Department IN  (" & dept_string & ") " 
END IF

strSQL = strSQL & " ORDER BY QuotedDraftReportDate, ActualDraftReportDate "

'response.Write strSQL

rs.open strSQL,db

response.write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0>"

' ***  WRITE HEADINGS
response.Write "<tr bgcolor=#E1E1FF>"
response.Write "<td width=100 align=center><B>Study Number</B></td>"
response.Write "<td width=150 align=center><B>Study Director</B></td>"
response.Write "<td width=200 align=center><B>Client Name</B></td>"
response.Write "<td width=100 align=center><B>Department</B></td>"
response.Write "<td width=100 align=center><B>Quoted Draft Report Date</B></td>"
response.Write "<td width=100 align=center><B>Actual Draft Report Date</B></td>"
response.Write "<td width=100 align=center><B>Actual Unaudited Draft To Client</B></td>"
response.Write "<td width=100 align=center><B>OTD Result</B></td>"
response.Write "</tr>"
response.Write "<tr><td colspan=8>&nbsp;</td></tr>"

DO UNTIL rs.eof=TRUE

	Response.Write "<TR>"
	Response.Write "<TD align=center><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>" & Trim(rs("StudyNumber")) & "</a></TD>"
      Response.Write "<TD>" & Trim(rs("StudyDirector")) & "</TD>"
      Response.Write "<TD>" & Trim(rs("ClientName")) & "</TD>"
      Response.Write "<TD>" & Trim(rs("Department")) & "</TD>"
      Response.Write "<TD align=center>" & FlipDateView(Trim(rs("QuotedDraftReportDate"))) & "</TD>"
      Response.Write "<TD align=center>" & FlipDateView(Trim(rs("ActualDraftReportDate"))) & "</TD>"
      Response.Write "<TD align=center>" & FlipDateView(Trim(rs("ActualUnauditedDraftToClient"))) & "</TD>"

	' ***  CALCULATE RESULT
	result = "NO"
	IF rs("QuotedDraftReportDate") > Date THEN
		result = "N/A"
	END IF
	IF TRIM(rs("Department")) = "E Fate" THEN
		IF IsBlank(rs("ActualUnauditedDraftToClient")) = False THEN
			IF cdate(MMDDYYYY(rs("ActualUnauditedDraftToClient"))) <= _
					cdate(MMDDYYYY(rs("QuotedDraftReportDate"))) THEN
				result = "YES"
			END IF
			'IF rs("ActualUnauditedDraftToClient") <= rs("QuotedDraftReportDate") THEN
			'	result = "YES"
			'END IF
		END IF
	ELSE
		IF IsBlank(rs("ActualDraftReportDate")) = False THEN
			IF cdate(MMDDYYYY(rs("ActualDraftReportDate"))) <= cdate(MMDDYYYY(rs("QuotedDraftReportDate"))) THEN
				result = "YES"
			END IF
			'IF rs("ActualDraftReportDate") <= rs("QuotedDraftReportDate") THEN
			'	result = "YES"
			'END IF
		END IF

	END IF
	Response.Write "<TD align=center "  
	IF result = "YES" THEN
		response.Write "bgcolor=#1af01c"
	ELSE
		IF result = "NO" THEN
			response.Write "bgcolor=#f72929"
		END IF
	END IF
	Response.Write ">" & result & "</TD>"
	Response.Write "</TR>"

	rs.MoveNext
LOOP
rs.Close

Response.Write "</table>"

response.Write "<br><br>"

%>

<!-- #include file = "includes\footer.asp"-->


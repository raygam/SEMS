
<!-- #include file = "includes\CortexConfig.asp"-->

<%

FUNCTION CalcCurrentStart()
	DIM daysToAdd
	DIM startDate
      DIM weekday
      weekday = DatePart("w", Now())
 	if (weekday < 7) then
		daysToAdd = (weekday) * -1
	else 
            'if (weekday = 7) then
           daysToAdd = 0
            'end if
	end if
	startDate = DateAdd("d", daysToAdd, Date())
	CalcCurrentStart = startDate
END FUNCTION
    
FUNCTION FlipDateView(dateToFlip)
	DIM strDateView
	DIM day, month, year
	day = DatePart("d", dateToFlip)
	month = MonthName(DatePart("m", dateToFlip), True) 
	year = DatePart("yyyy",dateToFlip)

	strDateView = day & " " & Ucase(month) &" " & year
	FlipDateView = strDateView
END FUNCTION

FUNCTION MMDDYYYY(dateToFlip)
	DIM day, month, year
      day = DatePart("d", dateToFlip)
	month = DatePart("m", dateToFlip) 
	year = DatePart("yyyy",dateToFlip)

	MMDDYYYY = month & "/" & day & "/" & year
END FUNCTION

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
      
	
' ***  START HERE 

dateStart = CalcCurrentStart()

response.write "<h1>&nbsp;&nbsp;Expected Draft To QA Schedule</h1>"	
response.write "<br>"

response.write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0>"

' ***  WRITE HEADING - NOW WRITING HEADINGS IN LOOP BELOW
'response.write "<TR>"
'response.write "<TD width=225></TD>"
'response.write "<TD width=100 align=center><B>Study Number</B></TD>"
'response.write "<TD width=150><B>Study Director</B></TD>"
'response.write "<TD width=200><B>Client Name</B></TD>"
'response.write "<TD width=100><B>Department</B></TD>"
'response.write "<TD width=100 align=center><B>Quoted Draft Report Date</B></TD>"
'response.write "<TD width=100 align=center><B>Expected Audited Draft to Client</B></TD>"
'response.write "<TD width=100 align=center><B>Expected Unudited Draft to Client</B></TD>"
'response.write "</TR>"
'response.write "<tr><td colspan=8>&nbsp;</td></tr>"


for i = 0 to 11

	weekStart = DateAdd("d", i * 7 , dateStart)
      weekEnd = DateAdd("d", 6, weekStart)

	strSQL = " SELECT * FROM Studies WHERE ExpectedDraftReportToQA "
	strSQL = strSQL & " BETWEEN '" & MMDDYYYY(weekStart) & "' AND '" & MMDDYYYY(weekEnd) & "'"
	strSQL = strSQL & " AND ExpectedDraftReportToQA IS NOT NULL "
	strSQL = strSQL & " ORDER BY ExpectedDraftReportToQA, StudyId"
	'response.write strSQL

      rs.open strSQL,db

	'response.write "<tr><td bgcolor=#E1E1FF ALIGN=CENTER><B>" & FlipDateView(weekStart) & " - " & FlipDateView(weekEnd) & "</B></td><td colspan=7 bgcolor=#E1E1FF>&nbsp;</td></tr>" 

	' ***  REPEAT HEADING
	response.write "<TR bgcolor=#E1E1FF>"
	response.write "<td bgcolor=#E1E1FF ALIGN=CENTER><B>" & FlipDateView(weekStart) & " - " & FlipDateView(weekEnd) & "</B></td>"
	response.write "<TD width=100 align=center><B>Study Number</B></TD>"
	response.write "<TD width=150><B>Study Director</B></TD>"
	response.write "<TD width=150><B>Client Name</B></TD>"
	response.write "<TD width=150><B>Department</B></TD>"
	response.write "<TD width=150><B>Study Type</B></TD>"
	response.write "<TD width=150 align=center><B>Expected Draft Report To QA</B></TD>"
	response.write "</TR>"
		        
      if rs.eof=false or rs.bof=false then
		do until rs.eof=true
    			Response.Write "<TR>"
                  Response.Write "<TD>&nbsp;</TD>"
                  Response.Write "<TD align=center><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>" & Trim(rs("StudyNumber")) & "</a></TD>"
                  Response.Write "<TD>" & Trim(rs("StudyDirector")) & "</TD>"
                  Response.Write "<TD>" & Trim(rs("ClientName")) & "</TD>"
                  Response.Write "<TD>" & Trim(rs("Department")) & "</TD>"
                  Response.Write "<TD>" & Trim(rs("StudyType")) & "</TD>"
			If IsBlank(rs("ExpectedDraftReportToQA")) = True Then
				Response.write "<td>&nbsp;</td>"
			Else
				response.write "<TD align=center>" & FlipDateView(Trim(rs("ExpectedDraftReportToQA"))) & "</TD>"
			End If
                  response.write "</TR>"

                  rs.movenext
            loop
	else                    
		response.write "<tr><td></td><td align=center>N/A</td><td colspan=5></td></tr>"
	end if

      response.write "<tr><td colspan=7>&nbsp;</td></tr>"
      response.write "<tr><td colspan=7>&nbsp;</td></tr>"
      rs.close
      
next

response.write "<tr><td colspan=7>&nbsp;</td></tr>"
response.write "</table>"


%>

<!-- #include file = "includes\footer.asp"-->


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

FUNCTION DisplayPct(intComp, intTotal)
	DIM strCell
	IF intTotal = 0 THEN
		strCell = "0"
	ELSE
		strCell = (intComp/intTotal) * 100
	END IF
	DisplayPct = ROUND(strCell,0) & "%"
END FUNCTION

' ***************************************
' ***  START HERE 
' ***************************************
DIM intMonth
DIM strSQL
DIM intYear
DIM strDept 

' *** TOTALS
DIM draft_deadlines, draft_completed 
DIM chem_total, eco_total, fate_total, pd_total, reg_total
DIM chem_comp, eco_comp, fate_comp, pd_comp, reg_comp
DIM chem_qtr_total, eco_qtr_total, fate_qtr_total, pd_qtr_total, reg_qtr_total
DIM chem_qtr_comp, eco_qtr_comp, fate_qtr_comp, pd_qtr_comp, reg_qtr_comp
DIM chem_ytd_total, eco_ytd_total, fate_ytd_total, pd_ytd_total, reg_ytd_total
DIM chem_ytd_comp, eco_ytd_comp, fate_ytd_comp, pd_ytd_comp, reg_ytd_comp
DIM yr_deadlines, yr_completed, qtr_deadlines, qtr_completed
DIM dept

response.write "<h1>&nbsp;&nbsp;On-Time Delivery (By Department) - 2017</h1>"	
response.write "<br>"
response.write "<br>"

response.Write "<form method=post action=OTD_Dept.asp>"

'response.Write "<select name=ddlYear>"
'response.Write "<option selected value=2017>2017</option>"
'response.Write "</select>"
'response.Write "<input type=submit value=submit>"
'response.write "<br>"
'response.write "<br>"

response.write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0>"


' ***  CAPTURE THE YEAR
intYear = 2017 'Request.Form("ddlYear")


' ***  WRITE HEADINGS
response.Write "<tr bgcolor=#E1E1FF>"
response.Write "<td width=150></td>"
response.Write "<td width=150 align=center><B>Analytical & Physical Chemistry</B></td>"
response.Write "<td width=150 align=center><B>Ecotoxicology</B></td>"
response.Write "<td width=150 align=center><B>E Fate</B></td>"
'response.Write "<td width=150 align=center><B>Product Development Testing</B></td>"
response.Write "<td width=150 align=center><B>Regulatory</B></td>"
response.Write "<td width=150 align=center><B>Overall</B></td>"
response.Write "</tr>"
response.Write "<tr><td colspan=6>&nbsp;</td></tr>"

' ***  LOOP THROUGH MONTHS
FOR intMonth = 1 TO 12

	' ***  RESET TOTALS FOR THE MONTH
	chem_total = 0
	eco_total = 0
	fate_total = 0
	pd_total = 0
	reg_total = 0
	chem_comp = 0
	eco_comp = 0
	fate_comp = 0
	pd_comp = 0
	reg_comp = 0

	strSQL = "SELECT * FROM Studies " &_
				" WHERE " &_
				" MONTH(QuotedDraftReportDate) = " & intMonth & " " &_
				" AND " &_
				" YEAR(QuotedDraftReportDate) = " & intYear & " " &_
				" AND LEN(StudyNumber) = 7 "

      rs.open strSQL,db

	strDept = "" 
	DO UNTIL rs.eof=TRUE
		
		strDept = 	TRIM(rs("Department"))

		' ***  LOOK AT EACH DEADLINE - ASSIGN TO DEPT
		SELECT CASE strDept
			CASE "Analytical & Physical Chemistry"
				chem_total =  chem_total + 1
			CASE "Ecotoxicology"
				eco_total =  eco_total + 1
			CASE "E Fate"
				fate_total = fate_total + 1
			CASE "Product Development Testing"
				eco_total =  eco_total + 1
				'pd_total = pd_total + 1
			CASE "Regulatory"
				reg_total = reg_total + 1				
		END SELECT

		' ***  SEE IF DEADLINE WAS MET.  IF EFATE, USE A DIFFERENT FIELD FOR COMPARISON
		IF TRIM(rs("Department")) = "E Fate" THEN
			IF IsBlank(rs("ActualUnauditedDraftToClient")) = False THEN
				IF cdate(MMDDYYYY(rs("ActualUnauditedDraftToClient"))) <= _
						cdate(MMDDYYYY(rs("QuotedDraftReportDate"))) THEN
					fate_comp = fate_comp + 1
				END IF
			END IF
		ELSE
			IF IsBlank(rs("ActualDraftReportDate")) = False THEN
				IF cdate(MMDDYYYY(rs("ActualDraftReportDate"))) <= cdate(MMDDYYYY(rs("QuotedDraftReportDate"))) THEN
		
					SELECT CASE strDept
						CASE "Analytical & Physical Chemistry"
							chem_comp =  chem_comp + 1
						CASE "Ecotoxicology"
							eco_comp =  eco_comp + 1
						CASE "Product Development Testing"
							'pd_comp = pd_comp + 1
							eco_comp =  eco_comp + 1
						CASE "Regulatory"
							reg_comp = reg_comp + 1				
					END SELECT

				END IF

			END IF

		END IF


		strDept = ""
		rs.movenext
	LOOP

	rs.close

	' ***  MONTH ROW
	response.Write "<tr>"
	response.Write "<td><B>" & MonthName(intMonth, False) & "</B></td>"
	response.Write "<td align=center>" & chem_comp & "/" & chem_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=1&month="& intMonth & "&year=" & intYear & "&qtr=>"  & DisplayPct(chem_comp, chem_total) & "</a></td>"
	response.Write "<td align=center>" & eco_comp & "/" & eco_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=2&month="& intMonth & "&year=" & intYear & "&qtr=>" & DisplayPct(eco_comp, eco_total) & "</a></td>"
	response.Write "<td align=center>" & fate_comp & "/" & fate_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=3&month="& intMonth & "&year=" & intYear & "&qtr=>" & DisplayPct(fate_comp, fate_total) & "</a></td>"
	'response.Write "<td align=center>" & pd_comp & "/" & pd_total & "&nbsp;&nbsp;&nbsp;" & DisplayPct(pd_comp, pd_total) & "</td>"
	response.Write "<td align=center>" & reg_comp & "/" & reg_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=4&month="& intMonth & "&year=" & intYear & "&qtr=>" & DisplayPct(reg_comp, reg_total) & "</a></td>"
	response.Write "<td align=center><B>" & (chem_comp + eco_comp + fate_comp + reg_comp) & "/" & (chem_total + eco_total + fate_total + reg_total) & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=&month="& intMonth & "&year=" & intYear & "&qtr=>" & DisplayPct(chem_comp + eco_comp + fate_comp + reg_comp, chem_total + eco_total + fate_total + reg_total) & "</a></B></td>"
	response.Write "</tr>"

	' ***  ADD TO QTR TOTALS
	chem_qtr_total = chem_qtr_total + chem_total
	chem_qtr_comp = chem_qtr_comp + chem_comp
	eco_qtr_total = eco_qtr_total + eco_total
	eco_qtr_comp = eco_qtr_comp + eco_comp
	fate_qtr_total = fate_qtr_total + fate_total
	fate_qtr_comp = fate_qtr_comp + fate_comp
	pd_qtr_total = pd_qtr_total + pd_total
	pd_qtr_comp = pd_qtr_comp + pd_comp
	reg_qtr_total = reg_qtr_total + reg_total
	reg_qtr_comp = reg_qtr_comp + reg_comp

	' ***  ADD TO YTD TOTALS
	chem_ytd_total = chem_ytd_total + chem_total
	chem_ytd_comp = chem_ytd_comp + chem_comp
	eco_ytd_total = eco_ytd_total + eco_total
	eco_ytd_comp = eco_ytd_comp + eco_comp
	fate_ytd_total = fate_ytd_total + fate_total
	fate_ytd_comp = fate_ytd_comp + fate_comp
	pd_ytd_total = pd_ytd_total + pd_total
	pd_ytd_comp = pd_ytd_comp + pd_comp
	reg_ytd_total = reg_ytd_total + reg_total
	reg_ytd_comp = reg_ytd_comp + reg_comp


	' ***  CHECK FOR END OF QTR - ADD A ROW IF IT IS
	IF intMonth = 3 OR intMonth = 6 OR intMonth = 9 OR intMonth = 12 THEN
		response.Write "<tr bgcolor=#d3d3d3>"
		response.Write "<td align=center><B>QTR " & CINT(intMonth)/3 & "</B></td>"
		response.Write "<td align=center>" & chem_qtr_comp & "/" & chem_qtr_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=1&month=&year=" & intYear & "&qtr=" & CINT(intMonth)/3 &">" & DisplayPct(chem_qtr_comp, chem_qtr_total) & "</a></td>"
		response.Write "<td align=center>" & eco_qtr_comp & "/" & eco_qtr_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=2&month=&year=" & intYear & "&qtr=" & CINT(intMonth)/3 &">" & DisplayPct(eco_qtr_comp, eco_qtr_total) & "</a></td>"
		response.Write "<td align=center>" & fate_qtr_comp & "/" & fate_qtr_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=3&month=&year=" & intYear & "&qtr=" & CINT(intMonth)/3 &">" & DisplayPct(fate_qtr_comp, fate_qtr_total) & "</a></td>"
		'response.Write "<td align=center>" & pd_qtr_comp & "/" & pd_qtr_total & "&nbsp;&nbsp;&nbsp;" & DisplayPct(pd_qtr_comp, pd_qtr_total) & "</td>"
		response.Write "<td align=center>" & reg_qtr_comp & "/" & reg_qtr_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=4&month=&year=" & intYear & "&qtr=" & CINT(intMonth)/3 &">" & DisplayPct(reg_qtr_comp, reg_qtr_total) & "</a></td>"
		response.Write "<td align=center><B>" & (chem_qtr_comp + eco_qtr_comp + fate_qtr_comp + reg_qtr_comp) & "/" & (chem_qtr_total + eco_qtr_total + fate_qtr_total + reg_qtr_total) & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=&month=&year=" & intYear & "&qtr=" & CINT(intMonth)/3 &">" & DisplayPct(chem_qtr_comp + eco_qtr_comp + fate_qtr_comp + reg_qtr_comp, chem_qtr_total + eco_qtr_total + fate_qtr_total + reg_qtr_total) & "</a></B></td>"
		response.Write "</tr>"
		response.Write "<tr><td colspan=6>&nbsp;</td></tr>"

		' ***  RESET QTR TOTALS
		chem_qtr_total = 0
		chem_qtr_comp = 0
		eco_qtr_total = 0
		eco_qtr_comp = 0
		fate_qtr_total = 0
		fate_qtr_comp = 0
		pd_qtr_total = 0
		pd_qtr_comp = 0
		reg_qtr_total = 0
		reg_qtr_comp = 0

	END IF

NEXT
    
' ***  ADD A TOTAL ROW
response.Write "<tr bgcolor=#E1E1FF>"
response.Write "<td align=center><B>TOTALS :</B></td>"
response.Write "<td align=center><B>" & chem_ytd_comp & "/" & chem_ytd_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=1&month=&year=" & intYear & "&qtr=>" & DisplayPct(chem_ytd_comp, chem_ytd_total) & "</a></B></td>"
response.Write "<td align=center><B>" & eco_ytd_comp & "/" & eco_ytd_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=2&month=&year=" & intYear & "&qtr=>" & DisplayPct(eco_ytd_comp, eco_ytd_total) & "</a></B></td>"
response.Write "<td align=center><B>" & fate_ytd_comp & "/" & fate_ytd_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=3&month=&year=" & intYear & "&qtr=>" & DisplayPct(fate_ytd_comp, fate_ytd_total) & "</a></B></td>"
'response.Write "<td align=center><B>" & pd_ytd_comp & "/" & pd_ytd_total & "&nbsp;&nbsp;&nbsp;" & DisplayPct(pd_ytd_comp, pd_ytd_total) & "</B></td>"
response.Write "<td align=center><B>" & reg_ytd_comp & "/" & reg_ytd_total & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=4&month=&year=" & intYear & "&qtr=>" & DisplayPct(reg_ytd_comp, reg_ytd_total) & "</a></B></td>"
response.Write "<td align=center><B>" & (chem_ytd_comp + eco_ytd_comp + fate_ytd_comp + reg_ytd_comp) & "/" & (chem_ytd_total + eco_ytd_total + fate_ytd_total + reg_ytd_total) & "&nbsp;&nbsp;&nbsp;<a href=OTD_StudyList.asp?type=1&dept=&month=&year=" & intYear & "&qtr=>" & DisplayPct(chem_ytd_comp + eco_ytd_comp + fate_ytd_comp + reg_ytd_comp, chem_ytd_total + eco_ytd_total + fate_ytd_total + reg_ytd_total) & "</a></B></td>"
response.Write "</tr>"

response.write "</table>"

response.Write "</form>"
response.Write "<br><br>"

%>

<!-- #include file = "includes\footer.asp"-->

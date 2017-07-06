<div>
<%
Response.Write "<br><font face=arial size=1>Page Executed in " & datediff("s",dtStart, now()) & " seconds."
%>
</div>
	</body>
</html>
<%Function ConvertToSendSearch(strSearch)
   	strSearch = replace(strSearch," ","[sp]")
	strSearch = replace(strSearch,"&","^")
	strSearch = replace(strSearch,"/","[sl]")
	strSearch = replace(strSearch,chr(34),"~")
	ConvertToSendSearch = strSearch
End Function
Function ConvertToGetSearch(strSearch)
	strSearch = replace(strSearch,"^","&")
	strSearch = replace(strSearch,"~",chr(34))
	strSearch = replace(strSearch,"[sp]"," ")
	strSearch = replace(strSearch,"[sl]","/")
	ConvertToGetSearch = strSearch
End Function
Function GetVariable(strVariable)
	s = trim(Request.Form(strVariable))
	if s="" then
		s = trim(Request.QueryString(strVariable))
	end if
	GetVariable = s
End Function
Function FormatDate(strDate)
	s = trim(strDate)
	if s & "" = "" then 
		FormatDate = ""
	else
		if isdate(s)=false then
			FormatDate = "##/###/####"
		else	
			FormatDate = right("0" & day(s),2) & "/" & left(monthname(month(s)),3) & "/" & year(s)
		end if
	end if
	
End Function
Function CreateWebKey()
	dim s, i
	for i = 1 to 16
		randomize
		s = s & chr(65+int(rnd*26))
	next
	s = s & now()
	s = replace(s,"/","")
	s = replace(s,":","")
	s = replace(s," ","")
	CreateWebKey = s
End Function
Function TextBoxHTML(FormFieldName, FormFieldValue, Width, MaxLength, Locked)
	TextBoxHTML = "<input type=""text"" name=""" & FormFieldName & """ size=""" & Width & """ MaxLength=""" & MaxLength & """ value=""" & FormFieldValue & """"
	if Locked>"" then
		TextBoxHTML = TextBoxHTML & " disabled"
	end if
	TextBoxHTML = TextBoxHTML & ">"
End Function

Function TextBoxHTMLGreen(FormFieldName, FormFieldValue, Width, MaxLength, Locked)
	TextBoxHTMLGreen = "<input style=""color: green"" type=""text"" name=""" & FormFieldName & """ size=""" & Width & """ MaxLength=""" & MaxLength & """ value=""" & FormFieldValue & """"
	if Locked>"" then
		TextBoxHTMLGreen = TextBoxHTMLGreen & " disabled"
	end if
	TextBoxHTMLGreen = TextBoxHTMLGreen & ">"
End Function

Function OptionHTML(FormFieldName, FormFieldValue, OptionList, OptionsPerRow, Locked)
	OptionHTML = "<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr>"
	OL = split(OptionList,";")
	c = 0
	for i = 0 to ubound(OL)
		c = c + 1
		if c > OptionsPerRow then
			OptionHTML = OptionHTML & "</tr>"
			c=0
		end if	
		OptionHTML = OptionHTML & "<td valign=top><input type=""Radio"" name=""" & FormFieldName & """ Value=""" & OL(i) & """"
		if trim(OL(i)) = trim(FormFieldValue) then
			OptionHTML = OptionHTML & " checked"
		end if
		if Locked>"" then
			OptionHTML = OptionHTML & " disabled"
		end if
		OptionHTML = OptionHTML & ">&nbsp;&nbsp;" & OL(i) & "&nbsp;&nbsp;</td>" 
	next
	OptionHTML = OptionHTML & "</tr></table>"	
	
End Function

Function CheckBoxHTML(FormFieldName, FormFieldValue, Locked)
	CheckBoxHTML = "<input type=""checkbox"" name=""" & FormFieldName & """ value=""Y"""
	if trim(Locked)>"" then
		CheckBoxHTML = CheckBoxHTML & " disabled"
	end if
	if trim(FormFieldValue) = "Y" or trim(FormFieldValue) = "-1"  then
		CheckBoxHTML = CheckBoxHTML & " checked"
	end if
	CheckboxHTML = CheckBoxHTML & ">"
End Function



Function SimpleComboHTML(FormFieldName, FormFieldValue, OptionList, Locked, BlankOption)
	SimpleComboHTML = "<Select name=""" & FormFieldName & """"
	if Locked>"" then
		SimpleComboHTML = SimpleComboHTML & " disabled"
	end if
	SimpleComboHTML = SimpleComboHTML & ">"
	OL = split(OptionList,";")
	if BlankOption>"" then
		SimpleComboHTML = SimpleComboHTML & "<option></option>"
	end if
	for i = 0 to ubound(OL)
		SimpleComboHTML = SimpleComboHTML & "<option"
		if trim(OL(i)) = trim(FormFieldValue) then
			SimpleComboHTML = SimpleComboHTML & " selected"
		end if
		SimpleComboHTML = SimpleComboHTML & ">" & OL(i) & "</option>"
	next
	SimpleComboHTML = SimpleComboHTML & "</select>"	

End Function

Function DataComboHTML(FormFieldName, FormFieldValue, RecordSource, Locked, BlankOption)
	DataComboHTML = "<Select name=""" & FormFieldName & """"
	if Locked>"" then
		DataComboHTML = DataComboHTML & " disabled"
	end if
	DataComboHTML = DataComboHTML & ">"
	rs2.open RecordSource, db
	if BlankOption>"" then
		DataComboHTML = DataComboHTML & "<option></option>"
	end if
	do until rs2.eof=true
		DataComboHTML = DataComboHTML & "<option value=""" & rs2(0) & """"
		if trim(rs2(0)) = trim(FormFieldValue) then
			DataComboHTML = DataComboHTML & " selected"
		end if
		DataComboHTML = DataComboHTML & ">" & rs2(1) & "</option>"
		rs2.movenext
	loop
	rs2.close
	DataComboHTML = DataComboHTML & "</select>"	


End Function

Function TextDataComboHTML(FormFieldName, FormFieldValue, RSource, Locked, BlankOption)
	TextDataComboHTML = "<Select style=""width:200px"" name=""" & FormFieldName & """"
	if Locked>"" then
		TextDataComboHTML = TextDataComboHTML & " disabled"
	end if
	TextDataComboHTML = TextDataComboHTML & ">"
	dim rsCombo,dbCombo
	set dbCombo = server.CreateObject("ADODB.Connection")
	dbCombo.open "Provider=SQLNCLI10;SERVER=SHWTetra;DATABASE=SEMS;UID=CortexApp;PWD=(#GetiN#)"
	set rsCombo = server.CreateObject("ADODB.Recordset")

	
	Rsource = trim(RSource)
	rsCombo.open RSource, dbCombo
	'rsCombo.open "SELECT LookupValue FROM LookupValues WHERE LookupID = 6", dbCombo
	if BlankOption>"" then
		TextDataComboHTML = TextDataComboHTML & "<option></option>"
	end if
	do until rsCombo.eof=true
		TextDataComboHTML = TextDataComboHTML & "<option value=""" & rsCombo(0) & """"
		if trim(rsCombo(0)) = trim(FormFieldValue) then
			TextDataComboHTML = TextDataComboHTML & " selected"
		end if
		TextDataComboHTML = TextDataComboHTML & ">" & rsCombo(0) & "</option>"
		rsCombo.movenext
	loop
	rsCombo.close
	dbCombo.close
	TextDataComboHTML = TextDataComboHTML & "</select>"	
	set rsCombo=nothing

End Function


Function PreviousDateHTML(FormFieldName, FormFieldValue,Locked)
	if FormFieldValue>"" and isdate(FormFieldValue)=true then
		d = day(FormFieldValue)
		m = month(FormFieldValue)
		y = year(FormFieldValue)
	else
		d = 0
		m = 0
		y = 0
	end if
	PreviousDateHTML = "<Select name=""" & FormFieldName & "Day" & """"
	if Locked>"" then
		PreviousDateHTML = PreviousDateHTML & " disabled"
	end if
	PreviousDateHTML = PreviousDateHTML & ">"
	PreviousDateHTML = PreviousDateHTML & "<option></option>"	
	for i = 1 to 31
		PreviousDateHTML = PreviousDateHTML & "<option"
		if trim(cstr(i))=trim(d) then
			PreviousDateHTML = PreviousDateHTML & " selected"
		end if
		PreviousDateHTML = PreviousDateHTML & ">" & i & "</option>"
	next
	PreviousDateHTML = PreviousDateHTML & "</Select> / "
	PreviousDateHTML = PreviousDateHTML & "<Select name=""" & FormFieldName & "Month" & """"
	if Locked>"" then
		PreviousDateHTML = PreviousDateHTML & " disabled"
	end if
	PreviousDateHTML = PreviousDateHTML & ">"
	PreviousDateHTML = PreviousDateHTML & "<option value=""0""></option>"
	for i = 1 to 12
		PreviousDateHTML = PreviousDateHTML & "<option value=""" & i & """"
		if trim(cstr(i))=trim(m) then
			PreviousDateHTML = PreviousDateHTML & " selected"
		end if
		PreviousDateHTML = PreviousDateHTML & ">" & monthname(i) & "</option>"
	next
	PreviousDateHTML = PreviousDateHTML & "</Select> / "
	PreviousDateHTML = PreviousDateHTML & "<Select name=""" & FormFieldName & "Year" & """"
	if Locked>"" then
		PreviousDateHTML = PreviousDateHTML & " disabled"
	end if
	PreviousDateHTML = PreviousDateHTML & ">"
	PreviousDateHTML = PreviousDateHTML & "<option></option>"		
	for i = year(now()) to 1900 step -1 
		PreviousDateHTML = PreviousDateHTML & "<option"
		if trim(cstr(i))=trim(y) then
			PreviousDateHTML = PreviousDateHTML & " selected"
		end if
		PreviousDateHTML = PreviousDateHTML & ">" & i & "</option>"
	next
	PreviousDateHTML = PreviousDateHTML & "</Select>"
	
	
	
End Function

Function DBLookup(Table, IDField, IDValue, NameField)
	set rsDBL = server.CreateObject("ADODB.Recordset")
	strsql = "SELECT * FROM [" & Table & "] WHERE " & IDField & " = " & IDValue
	rsDBL.open strsql, db
	if rsDBL.eof=false or rsDBL.bof=false then
		DBLookup = rsDBL(NameField)
	else
		DBLookup = ""
	end if
	rsDBL.close
	
End Function
Function DBLookupString(Table, IDField, IDValue, NameField)
	strsql = "SELECT * FROM [" & Table & "] WHERE " & IDField & " = '" & IDValue & "'"
	rs.open strsql, db
	if rs.eof=false or rs.bof=false then
		DBLookupString = rs(NameField)
	else
		DBLookupString = ""
	end if
	rs.close
	
End Function

Function WriteToChangeLog(LogTable, RecordID, UserID, ChangeTable, ChangeField, OldValue, NewValue)
	if trim(OldValue) <> trim(NewValue) then
		
		strsql = "INSERT INTO [" & LogTable & "] (ChangedBy, RecordID, ChangeTable, ChangeField, OldValue, NewValue) "
		'Response.Write strsql & "- " & ChangeTable & "<hr>"
		strsql = strsql & " VALUES (" & UserID & "," & RecordID & ",'" & ChangeTable & "','" & ChangeField & "','" & replace(left(OldValue,50),"'","''") & "','" & replace(left(NewValue,50),"'","''") & "')"
		'Response.Write strsql & "<hr>"
		rs2.open strsql, db, 1, 3
		WriteToChangeLog = true
		session("FieldChange")="Y"
	else
		WriteToChangeLog = false	
	end if
	
End Function

Function CompareDates(FormDate, SQLDate)
		if isdate(FormDate)=true and isdate(SQLDate)=true then
			t = year(FormDate) & "-" & right("0" & month(FormDate),2) & "-" & right("0" & day(FormDate),2)
			if formatdate(FormDate) = formatdate(SQLDate) then
			'if trim(t)=trim(SQLDate) then
				CompareDates = true
			else
				CompareDates = false
			end if
		else
			if FormDate = SQLDate then
				CompareDates = true
			else
				CompareDate = false	
			end if
		end if
End Function

Function DisplayOAStatus(Status)

	Select Case Status
		Case "L"
			DisplayOAStatus = "Live"
		Case "C"
			DisplayOAStatus = "Completed"
		Case "D"
			DisplayOAStatus = "Dead"
		Case "R"
			DisplayOAStatus = "On Hold"
		Case "S"
			DisplayOAStatus = "Suspended"
		Case "T"
			DisplayOAStatus = "Template"
		Case "V"
			DisplayOAStatus = "Closed >3M"
		Case "X"
			DisplayOAStatus = "Closed"
		Case "Y"
			DisplayOAStatus = "Closed >1M"
		Case "Z"
			DisplayOAStatus = "Complete Now"
		Case Else
			DisplayOAStatus = "Not Known"
	End Select

End Function

Function StripString(w)
	w = replace(w,chr(34),"")
	w = replace(w,"'","")
	w = replace(w,"<","")
	w = replace(w,">","")
	w = replace(w,"/","")
	w = replace(w,"\","")
	w = replace(w,",","")
	w = replace(w,"&","")
	w = replace(w,"?","")
	w = replace(w," ","")
	w = replace(w,".","")
	w = replace(w,"#","")
	w = replace(w,"%","")
	w = replace(w,"(","")
	w = replace(w,")","")
	s = ""
	for i = 1 to len(w)
		if isnumeric(mid(w,i,1))=true then
			s = s & mid(w,i,1)
		end if
	next 
	StripString = s
End Function	

Function DisplayCurrency(w)
	Select Case w
		Case "USD"
			DisplayCurrency = "$"
		Case "EUR"
			DisplayCurrency = "&euro;"
		Case "GBP"
			DisplayCurrency = "&pound;"	
		Case Else
			DisplayCurrency = ""
		End Select

End Function
%>
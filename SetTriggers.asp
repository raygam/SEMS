
<!-- #include file = "includes\CortexConfig.asp"-->
<script>
			function selectboxes(setting,trigger3) {
			  var len;
			  var el;
			  var i;
			  var form;
			  var tl;
			 
			  form = document.multi ;
			  if (form == null){return;}
			  len = form.elements.length ;
			  for (i=0 ; i<len ; i++) {
			   el = form.elements[i] ;
			   if (el.name.slice(0,trigger3.length) != trigger3) continue ;
			   el.checked = setting;
			  }
			}
			
			
			function checkThemAll(chk,trigger2){

				if(chk.checked==true){

					selectboxes(true,trigger2);

				}else{

					selectboxes(false,trigger2);

				}


			}

	</script>

<%
	Response.Expires = 0
	StartNum = Request.Form("StartNum")
	EndNum = Request.Form("EndNum")
	StartNum = StripString(StartNum)
	ClientName = Request.Form("ClientName")
	ClientName = replace(ClientName,"'","")
	ClientName = trim(replace(ClientName,chr(34),""))
	'Response.Write startnum & "-" & endnum & "-" & ClientName & "<br>"
	ClientList = Request.Form("ClientList")
	HighlightField=""
	if Request.Form("b1")="Save Data & Triggers" then
		dim s
		st = split(ClientList,";")
		
		for i = 0 to UBound(st)	
			'Response.Write s(i) & "<br>"
			StudyID = st(i)
			f = "ChangeClientName" & StudyID
			rs3.open "SELECT * FROM Studies WHERE StudyID = " & StudyID,db
			if rs3.eof=false or rs3.bof=false then
				StudyNumber = rs3("StudyNumberNew") & ""
			end if	
			rs3.close
			
			ChangeClientName = trim(Request.Form(f))
			'Response.Write "<font size=5 color=red>" & f & "</font>"
			if ChangeClientName > "" then
				'CLIENT NAME ADDED
				rs3.open "SELECT * FROM Studies WHERE StudyID = " & StudyID ,db,1,3
				if rs3.eof=false or rs3.bof=false then
					if ChangeClientName <> rs3("ClientName") then
						WriteToChangeLog "ChangeLog", rs3("StudyID"), Session("LoggedIn"), "Studies", "ClientName", rs3("ClientName") & "", ChangeClientName
					end if
					rs3("ClientName")=ChangeClientName
					
					rs3.update
				end if	
				rs3.close
			end if
			
			f = "ChangeFinancialClientName2" & StudyID
			ChangeFinancialClientName2 = trim(Request.Form(f))
			'Response.Write "<font size=5 color=red>FC2: " & ChangeFinancialClientName2 & "</font><br>"
			if ChangeFinancialClientName2 > "" then
				'CLIENT NAME ADDED
				rs3.open "SELECT * FROM Studies WHERE StudyID = " & StudyID ,db,1,3
				if rs3.eof=false or rs3.bof=false then
					if ChangeFinancialClientName2 <> rs3("FinancialClientName2") then
						WriteToChangeLog "ChangeLog", rs3("StudyID"), Session("LoggedIn"), "Studies", "FinancialClientName2", rs3("FinancialClientName2") & "", ChangeFinancialClientName2
					end if
					rs3("FinancialClientName2")=ChangeFinancialClientName2
					rs3.update
				end if	
				rs3.close
			end if
			
			
			f = "StudyNumber" & StudyID 
			StudyNumber = trim(Request.Form(f))
			
			f = "QuotedDraftReportDate" & StudyID 
			QuotedDraftReportDate = trim(Request.Form(f))
			if isdate(QuotedDraftReportDate) = true or QuotedDraftReportDate="" then
				'Save Data
				rs3.open "SELECT * FROM Studies WHERE StudyID = " & StudyID,db,1,3
				if rs3.eof=false or rs3.bof=false then
					if CompareDates(QuotedDraftReportDate, rs3("QuotedDraftReportDate")) = false then
						WriteToChangeLog "ChangeLog", rs3("StudyID"), Session("LoggedIn"), "Studies", "QuotedDraftReportDate", rs3("QuotedDraftReportDate") & "", QuotedDraftReportDate
					end if
					if QuotedDraftReportDate="" then
						rs3("QuotedDraftReportDate")= null
					else	
						rs3("QuotedDraftReportDate")=QuotedDraftReportDate
					end if
					rs3.update
				end if	
				rs3.close
			else
				HighlightField = HighlightField & "QuotedDraftReportDate" & StudyID & ";" 
				msg = msg & "<font color=red>Invalid Date QuotedDraftReportDate for " & StudyNumber & ": " & QuotedDraftReportDate & "</font> <b>This date change was NOT made.</b><br>"
			end if
			
			f = "QuotedExperimentalStartDate" & StudyID
			QuotedExperimentalStartDate = trim(Request.Form(f))
			if isdate(QuotedExperimentalStartDate) = true or QuotedExperimentalStartDate="" then
				
				'Save Data
				rs3.open "SELECT * FROM Studies WHERE StudyID = " & StudyID,db,1,3
				if rs3.eof=false or rs3.bof=false then
					if CompareDates(QuotedExperimentalStartDate, rs3("QuotedExperimentalStartDate")) = false then
						WriteToChangeLog "ChangeLog", rs3("StudyID"), Session("LoggedIn"), "Studies", "QuotedExperimentalStartDate", rs3("QuotedExperimentalStartDate") & "", QuotedExperimentalStartDate
					end if
					if QuotedExperimentalStartDate="" then
						rs3("QuotedExperimentalStartDate")=null
					else
						rs3("QuotedExperimentalStartDate")=QuotedExperimentalStartDate
					end if
					rs3.update
				end if	
				rs3.close
			else
				HighlightField = HighlightField & "QuotedExperimentalStartDate" & StudyID & ";" 
				msg = msg & "<font color=red>Invalid Date QuotedExperimentalStartDate for " & StudyNumber & ": " & QuotedExperimentalStartDate & "</font> <b>This date change was NOT made.</b><br>"
			end if
			
			'SAVE TRIGGERDATA
			f = "ProtocolIssuedDate" & StudyID
			ProtocolIssuedDate = trim(Request.Form(f))	
			a = UpdateTrigger("ProtocolIssuedDate",StudyID,ProtocolIssuedDate,StudyNumber)
			 
			f = "ActualExperimentalStartDate" & StudyID
			ActualExperimentalStartDate = trim(Request.Form(f))	
			a = UpdateTrigger("ActualExperimentalStartDate",StudyID,ActualExperimentalStartDate,StudyNumber)
			
			f = "ActualExperimentalEndDate" & StudyID
			ActualExperimentalEndDate = trim(Request.Form(f))	
			a = UpdateTrigger("ActualExperimentalEndDate",StudyID,ActualExperimentalEndDate,StudyNumber)
			
			f = "ActualDraftReportDate" & StudyID
			ActualDraftReportDate = trim(Request.Form(f))	
			a = UpdateTrigger("ActualDraftReportDate",StudyID,ActualDraftReportDate,StudyNumber)
			
			f = "ReportFinalSignDate" &StudyID
			ReportFinalSignDate = trim(Request.Form(f))	
			a = UpdateTrigger("ReportFinalSignDate",StudyID,ReportFinalSignDate,StudyNumber)
			
			
	
		next
		if msg = "" then
			msg = "<font color=green>All changes saved.</font><br>"
		end if	
		'Response.Write msg & "<br>"
	end if
	
	
	Response.write "<form method=post action=SetTriggers.asp>"
	Response.Write "Enter Start Study Number: <input type=text name=StartNum value=""" & Startnum & """>&nbsp;&nbsp;&nbsp;"
	Response.Write "Enter End Study Number: <input type=text name=EndNum value=""" & EndNum & """>&nbsp;&nbsp;&nbsp;"
	Response.Write "Enter Financial Client Name: <input type=text name=ClientName value=""" & ClientName & """>&nbsp;&nbsp;&nbsp;"
	Response.Write "<script type='text/JavaScript' src='scw.js'></script>"
	Response.Write "<input type=submit name=b1 value=Search>"
	
	Response.Write "</form>"
	
	if (StartNum>"" and EndNum >"") or ClientName>"" then
		strsql = "SELECT * FROM Studies WHERE StudyID > 0 "
	
		if StartNum>"" and EndNum >"" then
			if isnumeric(StartNum)=true and isnumeric(endnum)=true then
				strsql = strsql & " AND StudyNumberInt >= " & StartNum & " AND StudyNumberInt <= " & EndNum
			else
				msg = msg & "Please only use numeric values when searching on study numbers.<br>"
			end if
		end if
		If ClientName >"" then
			strsql = strsql & " AND FinancialClientName LIKE '" & ClientName & "%'"
		end if
	
		strsql = strsql & " ORDER BY StudyNumberNew ASC"
		'Response.Write strsql & "<br>"
	
		DisplayFields = "Study Number;Financial Client Name;Change Client Name?;Change Financial Client Name2?;Quoted Experimental<br>Start Date;Quoted Draft Report Date;Protocol Issued<br>Date;Actual Experimental<br>Start Date;Actual Experimental<br>End Date;Actual Draft<br>Report Date;Report Final<br>Sign Date"


		rs.open strsql,db
		if rs.eof=false or rs.bof=false then
			Response.Write "<h1>Set Triggers</h1>"	
			f = split(DisplayFields,";")
			if msg>""  then
				Response.Write "<font color=red>"& msg & "</font>"
			end if
			Response.Write "<form name=""multi"" id=""multi"" method=""post"" action=""SetTriggers.asp"">"
			
			Response.Write "<input type=Hidden name=StartNum value=""" & StartNum & """>"
			Response.Write "<input type=Hidden name=EndNum value=""" & EndNum & """>"
			Response.Write "<input type=Hidden name=ClientName value=""" & ClientName & """>"
			
			Response.Write "<table width=""100%"" cellspacing=0 cellpadding=3 border=0><tr>"
			Response.Write "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td align=center colspan=5 bgcolor=#E1E1FF><font color=navy><b>Please Schedule Email Triggers To Be Sent When These Dates Are Set.</b></font></td></tr>"
			for i = 0 to ubound(f)
				Response.Write "<td valign=top"
				if i > 4 then
					Response.Write " align=center"
				end if
				Response.Write "><b>" 
				
				Response.Write f(i)
				 
				Response.Write "</b></td>"
				f(i) = replace(f(i)," ","")
				f(i) = replace(f(i),"<br>","")
			next 
			
			
			Response.Write "</tr>"
			Response.Write "<tr><td colspan=5><td>"
			Response.Write "<td align=center><input name=""checkAll"" type=""checkbox"" id=""checkAll"" value=""1"" style=""width: 20px;"" onclick=""javascript:checkThemAll(this,'ProtocolIssuedDate');"" /></td>"
			Response.Write "<td align=center><input name=""checkAll"" type=""checkbox"" id=""checkAll"" value=""1"" style=""width: 20px;"" onclick=""javascript:checkThemAll(this,'ActualExperimentalStartDate');"" /></td>"
			Response.Write "<td align=center><input name=""checkAll"" type=""checkbox"" id=""checkAll"" value=""1"" style=""width: 20px;"" onclick=""javascript:checkThemAll(this,'ActualExperimentalEndDate');"" /></td>"
			Response.Write "<td align=center><input name=""checkAll"" type=""checkbox"" id=""checkAll"" value=""1"" style=""width: 20px;"" onclick=""javascript:checkThemAll(this,'ActualDraftReportDate');"" /></td>"
			Response.Write "<td align=center><input name=""checkAll"" type=""checkbox"" id=""checkAll"" value=""1"" style=""width: 20px;"" onclick=""javascript:checkThemAll(this,'ReportFinalSignDate');"" /></td>"
			
			
			
			
			Response.Write "</tr>"
				
			c=0
			ClientList = ""
			do until rs.eof=true 
				Response.Write "<tr"
				if c/2 = int(c/2) then
					Response.Write " bgcolor=#E1E1FF"
				end if
				Response.Write ">"
				for i = 0 to 1
					'Response.Write "<font color=red>" & f(i) & "</font>"
					if isdate(rs(f(i)))= true then
						Response.Write "<td align=right valign=top>" & formatdate(rs(f(i))) 
					else
						
						if f(i)="TotalStudyCost" then
							Response.Write "<td align=right valign=top>"
							if isnull(rs(f(i)))=false then
								
								Response.Write "&pound;" & formatnumber(rs(f(i)),2)
							end if
						else
							Response.Write "<td valign=top>"
							if f(i)="OAStatus" then
								Response.Write DisplayOAStatus(rs(f(i)) & "")
							else
								Response.Write rs(f(i)) 
							end if
						end if
					end if
					
					Response.Write "</td>"
				next
				ClientList = ClientList & rs("StudyID") & ";"
				
				Response.Write "<td valign=top>"
				Response.Write "<input type=text name=ClientName" & rs("StudyID") & " value=""" & rs("ClientName") & """ disabled><br>"
				'Response.Write "</td>"
				
				'CHANGE CLIENT NAME
				'Response.Write "<td valign=top>"
				Response.Write "<select name=ChangeClientName" & rs("StudyID") & " style=""width:140px"">"
				Response.Write "<option></option>"
				rs2.open "SELECT * FROM ClientLookup ORDER BY ClientName ASC",db
				do until rs2.eof=true
					Response.Write "<option>" & rs2("ClientName") & "</option>"
					rs2.movenext
				loop
				rs2.close
				Response.Write "</select>"
				Response.Write "<input type=Hidden name=StudyNumber" & rs("StudyID") & " value=""" & rs("StudyNumberNew") & """>"
				
				
				Response.Write "</td>"
				
				Response.Write "<td valign=top>"
				Response.Write "<input type=text name=ClientName" & rs("StudyID") & " value=""" & rs("FinancialClientName2") & """ disabled><br>"
				'CHANGE FINANCIAL CLIENT NAME 2
				'Response.Write "<td valign=top>"
				Response.Write "<select name=ChangeFinancialClientName2" & rs("StudyID") & " style=""width:140px"">"
				Response.Write "<option></option>"
				rs2.open "SELECT * FROM ClientLookup ORDER BY ClientName ASC",db
				do until rs2.eof=true
					Response.Write "<option>" & rs2("ClientName") & "</option>"
					rs2.movenext
				loop
				rs2.close
				Response.Write "</select>"
				Response.Write "</td>"
				
				'QUOTED EXPERIMENTAL START DATE
				Response.Write "<td valign=top>"
				Response.Write "<input id=""QuotedExperimentalStartDate" & rs("StudyID") & """ name=""QuotedExperimentalStartDate" & rs("StudyID") & """ type=""text"" tabindex=""" & c & """ value=""" & formatdate(rs("QuotedExperimentalStartDate")) & """"
				if instr(HighlightField, "QuotedExperimentalStartDate" & rs("StudyID"))>0 then
					Response.Write " style=""width:100px; background-color: #FF0000; color: #FFFFFF"" "
				else
					Response.Write " style=""width:100px"" "
				end if
				Response.Write " />"
				Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('" & "QuotedExperimentalStartDate" & rs("StudyID") & "'),event);"" />"
				Response.Write "</td>"
				
				'QUOTED DRAFT REPORT DATE
				Response.Write "<td valign=top>"
				
				Response.Write "<input id=""QuotedDraftReportDate" & rs("StudyID") & """ name=""QuotedDraftReportDate" & rs("StudyID") & """ type=""text"" tabindex=""" & c & """ value=""" & formatdate(rs("QuotedDraftReportDate")) & """"
				if instr(HighlightField, "QuotedDraftReportDate" & rs("StudyID"))>0 then
					Response.Write " style=""width:100px; background-color: #FF0000; color: #FFFFFF"" "
				else
					Response.Write " style=""width:100px"" "
				end if
				Response.Write " />"
				Response.Write "<img src=""images/inputCalendar.jpg"" title=""Click Here"" alt=""Click Here"" onclick=""scwShow(scwID('" & "QuotedDraftReportDate" & rs("StudyID") & "'),event);"" />"
				Response.Write "</td>"
				
				
				
				'Protocol Issued Date
				Response.Write "<td valign=top align=center><input type=checkbox id=ProtocolIssuedDate" & rs("StudyID") & " name=ProtocolIssuedDate" & rs("StudyID") & " value=ON"
				rs2.open "SELECT * FROM Triggers WHERE StudyID = " & rs("StudyID") & " AND TriggerField = 'ProtocolIssuedDate' AND IsActive = 1",db
				if rs2.eof=false or rs2.bof=false then
					Response.Write " checked"
				end if
				rs2.close
				Response.Write "></td>"
				
				'Actual Experimental Start Date
				Response.Write "<td valign=top align=center><input type=checkbox name=ActualExperimentalStartDate" & rs("StudyID") & " value=ON"
				rs2.open "SELECT * FROM Triggers WHERE StudyID = " & rs("StudyID") & " AND TriggerField = 'ActualExperimentalStartDate' AND IsActive = 1",db
				if rs2.eof=false or rs2.bof=false then
					Response.Write " checked"
				end if
				rs2.close
				Response.Write "></td>"
				 
				'Actual Experimental End Date 
				Response.Write "<td valign=top align=center><input type=checkbox name=ActualExperimentalEndDate" & rs("StudyID") & " value=ON"
				rs2.open "SELECT * FROM Triggers WHERE StudyID = " & rs("StudyID") & " AND TriggerField = 'ActualExperimentalEndDate' AND IsActive = 1",db
				if rs2.eof=false or rs2.bof=false then
					Response.Write " checked"
				end if
				rs2.close
				Response.Write "></td>"
				
				'Actual Draft Report Date 
				Response.Write "<td valign=top align=center><input type=checkbox name=ActualDraftReportDate" & rs("StudyID") & " value=ON"
				rs2.open "SELECT * FROM Triggers WHERE StudyID = " & rs("StudyID") & " AND TriggerField = 'ActualDraftReportDate' AND IsActive = 1",db
				if rs2.eof=false or rs2.bof=false then
					Response.Write " checked"
				end if
				rs2.close
				Response.Write "></td>"
				
				'Report Final Sign Date 
				Response.Write "<td valign=top align=center><input type=checkbox name=ReportFinalSignDate" & rs("StudyID") & " value=ON"
				rs2.open "SELECT * FROM Triggers WHERE StudyID = " & rs("StudyID") & " AND TriggerField = 'ReportFinalSignDate' AND IsActive = 1",db
				if rs2.eof=false or rs2.bof=false then
					Response.Write " checked"
				end if
				rs2.close
				Response.Write "></td>"
				
				'STRATEGIC PARTNER
				
				'Response.Write "<td valign=top><a href=""ViewStudy.asp?Key=" & rs("WebKey") & """>View</a></td>"
				'Response.Write "<td valign=top><a href=""EditStudies.asp?Key=" & rs("WebKey") & """>Edit</a></td>"
				
				Response.Write "</tr>"
				rs.movenext
				c=c+1
			loop
			
			Response.Write "</table>"
			if ClientList>"" then
			Response.Write "<br><center><input type=submit name=b1 value=""Save Data & Triggers""></center>"
			ClientList = left(ClientList, len(ClientList)-1)
			Response.Write "<input type=hidden name=ClientList value=""" & ClientList & """>"
			end if
			
		end if
	
		rs.close
	else
	
	
	end if
Response.Write "</form>"
	Function UpdateTrigger(TriggerField,StudyID,IsActive,StudyNumber)
		strsql = "SELECT * FROM Triggers WHERE TriggerField = '" & TriggerField & "' AND StudyID = " & StudyID
		'Response.Write strsql & " - " & IsActive &  "<br>"
		rs3.open strsql , db,1,3
		if rs3.eof=false or rs3.bof=false then
			if IsActive = "ON" then
				
				if rs3("IsActive")<> 1 then
					WriteToChangeLog "ChangeLog", StudyID, Session("LoggedIn"), "Triggers", TriggerField, rs3("IsActive") & "", "1" 
				end if
				rs3("IsActive")=1
			else
					
				if rs3("IsActive")<> 0 then
					WriteToChangeLog "ChangeLog", StudyID, Session("LoggedIn"), "Triggers", TriggerField, rs3("IsActive") & "", "0" 
				end if
				rs3("IsActive")=0
			end if
			rs3.update
		else
			
			If IsActive = "ON" then
				'ADD NEW TRIGGER
				rs3.addnew
				rs3("IsActive")=1
				rs3("StudyID")=StudyID
				rs3("TriggerField") = TriggerField
				rs3("CreatedBy") = session("Fullname")
				rs3("StudyNumber") = StudyNumber
				rs3.update
				WriteToChangeLog "ChangeLog", StudyID, Session("LoggedIn"), "Triggers", TriggerField, "New Record", "1" 
			end if
		end if
		rs3.close  
		UpdateTrigger = True
	End Function
%>

<!-- #include file = "includes\footer.asp"-->
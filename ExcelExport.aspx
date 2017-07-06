<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ExcelExport.aspx.vb" Inherits="ExcelExport" %>

<!--#Include File="includes\CortexConfig.aspx"-->
        <h1>Creating Excel Export....</h1>
    <%
        'On Error GoTo errhand
        Dim db As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset

        Dim rsStudies As New ADODB.Recordset
        Dim rsPrices As New ADODB.Recordset
        Dim dbSQL As New ADODB.Connection
        Dim rsSQL As New ADODB.Recordset
        Dim Grid(1000, 200)
        Dim strSql As String
        Dim strDB As String
        Dim CompanyCode As String
        Dim strPath As String
        Dim strFile As String
        Dim i As Integer
        Dim w As Long
        Dim t As String
        Dim intField As Integer
        Dim strDate As String
        Dim strURL As String

        Dim User As String
        User = UCase(Trim(Request.ServerVariables("AUTH_USER")))
        If User <> "SSI\KCOCKS" And User <> "SSI\DATACRAFTADMIN" And User <> "SSI\HDUNN" And User <> "SSI\FDAVIES" Then
            Response.Write("<br><blockquote>The user has not been granted access to this feature.</blockquote>")
            Response.End()
        End If


        'Response.Write(Server.MapPath("ProductsTemplate.aspx") & "<br>")
        'strPath = "D:\inetpub\vhosts\florismart.data-craft.co.uk\httpdocs\Excel\ProductsTemplate.xlsx"
        'strPath = "\\smrweb\c$\Websites\Cortex\wwwroot\Excel\SEMSKeyDataTemplateWithTriggers.xlsx"
        strPath = "c:\Websites\SEMS\wwwroot\Excel\SEMSKeyDataTemplateWithTriggers.xlsx"

        'strPath = "D:\inetpub\vhosts\client.florismart.com\httpdocs\Excel\ProductsTemplate.xlsx"

        strDate = Day(Date.Today) & MonthName(Month(Date.Today)).Substring(0, 3) & Year(Date.Today) & ""
        strFile = Replace(strPath, "Template", strDate)
        strURL = "http://smrweb:172/Excel/SEMSKeyData" & strDate & "WithTriggers.xlsx"
        strURL = "Excel/SEMSKeyData" & strDate & "WithTriggers.xlsx"
        
'Response.Write(strFile & "<br>")
        My.Computer.FileSystem.CopyFile(strPath, strFile, True)
        strDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile & ";Extended Properties=Excel 8.0"

        db.Open(strDB)

        strSql = "SELECT * FROM [Studies$]"

        rs.Open(strSql, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        'dbSQL.Open("Driver={SQL Server};Server=LocalHost;Database=FSCRM;Uid=FSCRMUser;Pwd=FSCRM5432#")
        dbSQL.Open("Provider=SQLNCLI10;SERVER=SHWTetra;DATABASE=SEMS;UID=CortexApp;PWD=(#GetiN#)")
        Dim f() As String
        strSql = "StudyID,StudyNumber,FinancialClientName,ClientName,TestSubstanceName,StudyDescription,StudyDirector,StudyType,StudyStatus,RegulatoryStatus,ProjectManagementCode,ProjectManager,Department,TotalStudyCost,ProtocolIssuedDate,ActualExperimentalStartDate,ActualExperimentalEndDate,ActualDraftreportDate,QuotedDraftReportDate,ReportFinalSignDate,ReportFinalisationDeadline,ArchiveActDate"


        strSql = "SELECT " & strSql & " FROM Studies WHERE OAStatus = 'L' ORDER BY StudyNumber ASC"
        rsStudies.Open(strSql, dbSQL)
        'strSql = "SELECT Products.VBNNumber, Products.ProductNameEn, Products.MinimumQuantity, ProductGroups.ProductGroupNameEN FROM ProductGroups INNER JOIN Products ON ProductGroups.VBNCode = Products.VBNCode WHERE Products.CoreProduct Is NOT NULL ORDER BY ProductGroups.ProductGroupNameEN ASC"
        'strSql = "StudyNumber,FinancialClientName,ClientName,TestSubstanceName,StudyDescription,StudyDirector,StudyType,StudyStatus,RegulatoryStatus,ProjectManagementCode,ProjectManager,Department,TotalStudyCost,ProtocolIssuedDate,ProtocolIssuedDateTrigger,ActualExperimentalStartDate,ActualExperimentalStartDateTrigger,ActualExperimentalEndDate,ActualExperimentalEndDateTrigger,ActualDraftReportDate,ActualDraftReportDateTrigger,QuotedDraftReportDate,QuotedDraftReportDateTrigger,ReportFinalSignDate,ReportFinalSignDateTrigger,ReportFinalisationDeadline,ArchiveActDate"
        strSql = "StudyNumber,FinancialClientName,ClientName,TestSubstanceName,StudyDescription,StudyDirector,StudyType,StudyStatus,RegulatoryStatus,ProjectManagementCode,ProjectManager,Department,TotalStudyCost,ProtocolIssuedDate,ProtocolIssuedDateTrigger,ActualExperimentalStartDate,ActualExperimentalStartDateTrigger,ActualExperimentalEndDate,ActualExperimentalEndDateTrigger,ActualDraftReportDate,ActualDraftReportDateTrigger,QuotedDraftReportDate,ReportFinalSignDate,ReportFinalSignDateTrigger,ReportFinalisationDeadline,ArchiveActDate"

        f = Split(strSql, ",")
        Dim ID As String
        Dim c As Long
        c = 1
        rsStudies.MoveNext()

        Do Until rsStudies.EOF = True

            rs.AddNew()
            For i = 0 To UBound(f)
                t = ""
                If InStr(f(i), "Trigger") = 0 Then
                    t = rsStudies.Fields(f(i)).Value & ""
                Else
                    'PROCESS TRIGGERS
                    strSql = "SELECT * FROM Triggers WHERE TriggerField = '" & Replace(f(i), "Trigger", "") & "' AND StudyID = " & rsStudies.Fields("StudyID").Value & " AND IsActive = 1"
                    rs2.Open(strSql, dbSQL)
                    If rs2.EOF = False Or rs2.BOF = False Then
                        t = "Y"
                    End If
                    rs2.Close()

                End If
                If t > "" Then
                    rs.Fields(i).Value = t
                End If


            Next i
            rs.Update()




            'Next



            c = c + 1

            rsStudies.MoveNext()
        Loop
        rsStudies.Close()
        rs.Close()
        strSql = "StudyID,StudyNumber,FinancialClientName,ClientName,TestSubstanceName,StudyDescription,StudyDirector,StudyType,StudyStatus,RegulatoryStatus,ProjectManagementCode,ProjectManager,Department,TotalStudyCost,ProtocolIssuedDate,ActualExperimentalStartDate,ActualExperimentalEndDate,ActualDraftreportDate,QuotedDraftReportDate,ReportFinalSignDate,ReportFinalisationDeadline,ArchiveActDate"
        'strSql = "StudyNumber,FinancialClientName,ClientName,TestSubstanceName,StudyDescription,StudyDirector,StudyType,StudyStatus,RegulatoryStatus,ProjectManagementCode,ProjectManager,Department,TotalStudyCost,ProtocolIssuedDate,ProtocolIssuedDateTrigger,ActualExperimentalStartDate,ActualExperimentalStartDateTrigger,ActualExperimentalEndDate,ActualExperimentalEndDateTrigger,ActualDraftReportDate,ActualDraftReportDateTrigger,QuotedDraftReportDate,QuotedDraftReportDateTrigger,ReportFinalSignDate,ReportFinalSignDateTrigger,ReportFinalisationDeadline,ReportFinalisationDeadlineTrigger,ArchiveActDate,ArchiveActDateTrigger"

        strSql = "SELECT " & strSql & " FROM Studies WHERE OAStatus = 'L' ORDER BY StudyNumber ASC"
        rsStudies.Open(strSql, dbSQL)




        strSql = "SELECT * FROM [Studies$]"

        rs.Open(strSql, db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        For i = 0 To UBound(f)
            rs.Fields(i).Value = DBNull.Value

            t = ""
            If InStr(f(i), "Trigger") = 0 Then
                t = rsStudies.Fields(f(i)).Value & ""
            Else
                'PROCESS TRIGGERS
                strSql = "SELECT * FROM Triggers WHERE TriggerField = '" & Replace(f(i), "Trigger", "") & "' AND StudyID = " & rsStudies.Fields("StudyID").Value & " AND IsActive = 1"
                rs2.Open(strSql, dbSQL)
                If rs2.EOF = False Or rs2.BOF = False Then
                    t = "Y"
                End If
                rs2.Close()

            End If
            If t > "" Then
                rs.Fields(i).Value = t
            End If


            't = rsStudies.Fields(f(i)).Value & ""
            'If t > "" Then
            ' rs.Fields(i).Value = t
            'End If
        Next i
        rs.Update()
        rs.Close()
        rsStudies.Close()
        db.Close()
        dbSQL.Close()

        Response.Write("Please download <a href=" & strURL & ">" & strURL & "</a>")
errhand:
        'Response.Write(w & " - " & (7 + (2 * i)))

        rs = Nothing
        db = Nothing
        dbSQL = Nothing

    %>

<!--#Include File="includes\footer.aspx"-->

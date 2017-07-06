<%@ Language=VBScript %>
<%

		set dbExcel = server.CreateObject("ADODB.Connection")
		set rsExcel = server.CreateObject("ADODB.Recordset")
		set cmdExcel = server.CreateObject("ADODB.Command")
		
		strPath = "c:\Websites\Cortex\wwwroot\Excel\TestExcel.xlsx"

        'strPath = "D:\inetpub\vhosts\client.florismart.com\httpdocs\Excel\ProductsTemplate.xlsx"

        'strDate = formatdate(Date)
        'strFile = Replace(strPath, "Template", strDate)
        'strURL = "http://smrweb:172/Excel/TestExcel.xlsx"
        'Response.Write(strFile & "<br>")
        'My.Computer.FileSystem.CopyFile(strPath, strFile, True)
        strDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"""

        dbExcel.Open strDB
        
        cmdExcel.ActiveConnection = dbExcel
        
        cmdExcel.CommandText = "CREATE TABLE Studies (IntField INT, Description VARCHAR, StartDate DATE, Amount DECIMAL)"
        cmdExcel.Execute
        set cmdExcel = nothing
        
        

        strSql = "SELECT * FROM [Studies$]"

		strsql = "INSERT INTO Studies (IntField, Description, StartDate,Amount) VALUES (17,'Test Record','17/02/1969',1.34)"

        rsExcel.Open strSql, dbExcel, 1,3
        
        rsExcel.close
		dbExcel.close
		Response.End
		rsExcel.addnew
		rsExcel(0) = 17
		rsExcel(1) = "Test Record"
		rsExcel(2) = "17/Feb/1969"
		rsExcel(3) = 1.23
		rsExcel.update
		
		
		set rsExcel = nothing
		set dbExcel = nothing
		Response.Write "Finished"
%>
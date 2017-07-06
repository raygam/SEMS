<html>
	<head>
		<title>SEMS</title>
		<link href="..\css\Cortex.css" rel="stylesheet" type="text/css" />
		<link href="..\css\CortexDashboard.css" rel="stylesheet" type="text/css" />
		<script language="javascript"> 
			
			function toggle() {
				var ele = document.getElementById("toggleText");
				var text = document.getElementById("displayText");
				if(ele.style.display == "block") {
    				ele.style.display = "none";
					text.innerHTML = "Show Report Criteria";
  				}
				else {
					ele.style.display = "block";
					text.innerHTML = "Hide Report Criteria";
				}
			} 
			
			function toggle_it(itemID){ 
     
				if ((document.getElementById(itemID).style.display == 'none')) { 
					document.getElementById(itemID).style.display = 'block' 
					
				 } else { 
					document.getElementById(itemID).style.display = 'none'; 
					
			}    
  } 
			
		</script>

	</head>
	
	<body>


		<a href="../mainmenu.asp"><img border="0" src="..\Images\SemsLogo.png"></a> &nbsp;&nbsp;<a href="javascript:window.print()"><font size=2>Print</font></a><br>
<%

dim strUserCode as string = ""
'strUserCode  = trim(Server.HtmlEncode(request.cookies("CRMLoginID").value))

if strUserCode  <>"GGGGG" then
	
	
else
	'REDIRECT TO LOGIN FAIL PAGE
	'Response.Write("<br><blockquote>The user has not been granted access to Cortex.</blockquote>")

	
	'response.write("SessionUserCode-" & strUserCode )
	'Response.End
end if
%>


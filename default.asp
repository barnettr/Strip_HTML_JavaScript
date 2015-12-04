<%@ LANGUAGE = VBScript %>
<% Option Explicit %>

<%  

srtHTML = MID(srtHTML, INSTR(srtHTML, "<html>"), INSTR(srtHTML, "</html>" +5)) 

%>


<html>
    <head>
        <title>Strip HTML Functions</title>
    </head>

    <body bgcolor="White" style="margin-left:0px; margin-right:0px; background-color:#ffffff;">
        
        <%= srtHTML %>
    
	</body>
</html>
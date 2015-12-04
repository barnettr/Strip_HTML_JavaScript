<%@ LANGUAGE = VBScript %>
<% Option Explicit %>

<!*************************
These two functions will strip out HTML and leave text. 
*************************>

<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
	
	Function stripHTML(strHTML)
	'Strips the HTML tags from strHTML

		Dim objRegExp, strOutput
		Set objRegExp = New Regexp

		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "&lt;(.|\n)+?&gt;"

		'Replace all HTML tag matches with the empty string
		strOutput = objRegExp.Replace(strHTML, "")
  
		'Replace all &lt; and &gt; with &amp;lt; and &amp;gt;
		strOutput = Replace(strOutput, "&lt;", "&amp;lt;")
		strOutput = Replace(strOutput, "&gt;", "&amp;gt;")
  
		stripHTML = strOutput    'Return the value of strOutput

		Set objRegExp = Nothing
	End Function
	
	Function stripHTML(strHTML)
	'Strips the HTML tags from strHTML using split and join

		'Ensure that strHTML contains <I>something</I>
		If len(strHTML) = 0 then
			stripHTML = strHTML
			Exit Function
		End If

		dim arysplit, i, j, strOutput

		arysplit = split(strHTML, "&lt;")
 
		'Assuming strHTML is not empty, we want to start iterating
		'from the 2nd array postition
		'if len(arysplit(0)) &gt; 0 then j = 1 else j = 0

		'Loop through each instance of the array
		for i=j to ubound(arysplit)
		'Do we find a matching &gt; sign?
			if instr(arysplit(i), "&gt;") then
				'If so, snip out all the text between the start of the string
				'and the &gt; sign
				arysplit(i) = mid(arysplit(i), instr(arysplit(i), "&gt;") + 1)
			else
				'Ah, the &lt; was was nonmatching
				'arysplit(i) = "&lt;" &amp; arysplit(i)
			end if
		next

		'Rejoin the array into a single string
		strOutput = join(arysplit, "")
  
		'Snip out the first &lt;
		strOutput = mid(strOutput, 2-j)
  
		'Convert &lt; and &gt; to &amp;lt; and &amp;gt;
		strOutput = replace(strOutput,"&gt;","&amp;gt;")
		strOutput = replace(strOutput,"&lt;","&amp;lt;")

		stripHTML = strOutput
	End Function
	
	Dim sStr
	sStr = "<a name="top" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes" xmlns:my="http://schemas.microsoft.com/office/infopath/2003/myXSD/2006-08-04T21:17:17"></a><table width="802" cellpadding="1" cellspacing="3" border="0" align="center" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:fn="http://www.w3.org/2005/02/xpath-functions" xmlns:xdt="http://www.w3.org/2005/02/xpath-datatypes" xmlns:my="http://schemas.microsoft.com/office/infopath/2003/myXSD/2006-08-04T21:17:17">
<tr>
<td width="802" bgcolor="#A3C2E0">
<table width="800" cellpadding="3" cellspacing="0" border="0">
<tr>
<td width="800" bgcolor="#336699"><font face="arial, helvetica, sans serif" size="2" color="#ffffff"><a name="Template Version"></a><b>Version:</b>  Template Version<br></font></td>
<td width="48" bgcolor="#336699" align="right"><font face="arial, helvetica, sans serif" size="2" color="#ffffff"><a href="#top" style="color: #ffffff;">top</a> ^  </font></td>
</tr>
<tr>
<td colspan="2" width="800" bgcolor="#cccccc"><font face="arial, helvetica, sans serif" size="2" color="#000000"><b>Subject Line:</b>  Don't miss these two amazing cruise offers from Disney <br></font></td>
</tr>
<tr>
<td colspan="2" width="800" bgcolor="#FFFFFF"><!--***************** START Template Version TEMPLATE*****************-->"

'Response.Write(Len(sStr))
	
</SCRIPT>


<html>
    <head>
        <title>Strip HTML Functions</title>
    </head>

    <body bgcolor="White" style="margin-left:0px; margin-right:0px; background-color:#ffffff;">
        
        <!-- Display header. -->
        <font size="4" face="arial, helvetica">
        <strong>Server Side Functions</strong></font><BR>
        <%= Response.Write(Len(sStr)) %>
    </body>
</html>
<%@ LANGUAGE="VBScript" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>TeMP:  Template eMail Previewer</title>
<%
Dim objXML, linkURL, folderStr, folderName, folderName2, folderStrL, folderStrR

linkURL = request.querystring("filePath")
folderStr = request.querystring("folder")

'Load the XML File
Set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.async = False 
Dim XSLTstyle, URLpath

'**Setup: to added new XSL Template to the TeMP

'1) Declare XSLT name, use in the Array of templates below 
Dim cm_template, fs_template, fs_template_01, fs_template_02a, fs_template_02b, tr_template, crm_template, cruises_template, ta_template, mr_template, msn_template

'2) Declare Folder name, use this in the Case Statement below
Dim cm_name, cruises_name, fs_name, fs_name_01, fs_name_02, tr_name, msn_name, crm_name, ta_name, mr_name 

'3) Associate declared vars with strings
cm_template = "cm_template.xslt"
cruises_template = "cruises_template.xslt"
fs_template = "fs_template.xslt"
fs_template_01 = "fs_template_01.xslt"
fs_template_02a = "fs_template_02a.xslt"
fs_template_02b = "fs_template_02b.xslt"
tr_template = "tr_template.xslt"
msn_template = "msn_template.xslt"
crm_template = "crm_template.xslt"
ta_template = "ta_template.xslt"
mr_template = "hybrid_template.xslt"
cm_name = "custom_email"
cruises_name = "cruises_email"
fs_name = "fare_sale"
fs_name_01 = "fare_sale_01"
fs_name_02 = "fare_sale_02"
tr_name = "travel_right"
msn_name = "msn_mail"
crm_name = "crm_mail"
ta_name= "ta_email"
mr_name= "flex_email"


'4) array of XSL Templates, get name from var declaration
XSLTstyle = Array(mr_template, cm_template, cruises_template, tr_template, crm_template, fs_template, fs_template_01, fs_template_02a, fs_template_02b, ta_template, msn_template)
'**END XSLT Setup - add condition to the Case Statement below (Step 5) and your done.

folderStrL = Left(folderStr, 6)
folderStrR = Right(folderStr, 5)
folderName2 = Right(folderStr, Len(folderStr) - Len(folderStrL))
folderName = Left(folderName2, Len(folderName2) - Len(folderStrR))
%>
<%
	'Set Upload = Server.CreateObject("Persits.Upload.1")
	dim rootPath, path
	'path = Upload.Files(1).Path
	rootPath = "C:\Inetpub\wwwroot\email\temp_images\"
	dim fs, folder
	set fs = CreateObject("Scripting.FileSystemObject")
	set folder = fs.GetFolder(rootPath)
%>
	<script language="JavaScript" type="text/JavaScript">
		if(document.images){
			expandC = new Image; 
			expandC.src="../images/minus.jpg";
			expandO = new Image;
			expandO.src="../images/plus.jpg";
		}
		function expand(expandSection){
			var exDiv = expandSection;
			var current = document.getElementById(exDiv).style.display;

			if (current == 'block')
			{
				document.getElementById(exDiv).style.display='none';
				if (document.images) 
					document.getElementById(exDiv + '_Img').src = expandO.src;
			}
			else
			{
				document.getElementById(exDiv).style.display='block';
				if (document.images) 
					document.getElementById(exDiv + '_Img').src = expandC.src;
			}
		}
		function refresh(){
			window.location.reload(true);
		}
		function loadIframe(theURL){
			document.getElementById("imgFRM").src = theURL;
		}
		</script>
		<script type="text/javascript" src="../includes/tabber.js"></script>
		<link rel=stylesheet type="text/css" href="../includes/temp_style.css"/>
	</head>
	<body bgcolor="#FFFFFF">
	<div id="faqpanel_1_answer">
	<div id="faqpanel_2_answer">
		<table cellpadding="0" cellspacing="5" width="100%">
	<tr><td colspan="2" style="font-family: verdana; font-size: 10; color: #000000;"><%= BreadCrumb(linkURL) %></td></tr></table>
	<%
'objXML.load(Server.MapPath(linkURL))
objXML.load(linkURL)

'5) Finish Setup process: Add the declared folder and XSLT names here.
Select Case folderName
   Case mr_name
       XSLTLocator(Server.MapPath(XSLTstyle(0)))
   Case cm_name
       XSLTLocator(Server.MapPath(XSLTstyle(1)))
   Case cruises_name
       XSLTLocator(Server.MapPath(XSLTstyle(2)))
   Case tr_name
       XSLTLocator(Server.MapPath(XSLTstyle(3)))
   Case crm_name
       XSLTLocator(Server.MapPath(XSLTstyle(4)))
   Case fs_name
       XSLTLocator(Server.MapPath(XSLTstyle(5)))
   Case fs_name_01
       XSLTLocator(Server.MapPath(XSLTstyle(6)))
   Case fs_name_02%>
     <div class="tabber">
		<div class="tabbertab" title="Version A"><br> 
		<% XSLTLocatorVers(Server.MapPath(XSLTstyle(7)))%>
		</div>
		<div class="tabbertab" title="Version B"><br>
		<% XSLTLocatorVers(Server.MapPath(XSLTstyle(8)))%>
		</div>
	</div><%
	Case ta_name
       XSLTLocator(Server.MapPath(XSLTstyle(9)))
    Case msn_name
       XSLTLocator(Server.MapPath(XSLTstyle(10)))
	Case Else
      Response.Write("what just happened...  don't know? Please email Craig at: <a href='mailto:ctoohey@expedia.com'>ctoohey@expedia.com</a>")
End Select
		
'Fxn to load XSLT and display page contents on demand
sub XSLTLocator(template)
	dim objXSL

	Set objXSL = Server.CreateObject("Microsoft.XMLDOM")

	objXSL.async = False
	
	objXSL.load(template)

	' Transform the XML file using the XSL stylesheet
	strHTML = objXML.transformNode(objXSL)
%>		
<table border="0" cellpadding="0" cellspacing="1">
<tr>
	<td valign="top" align="left" width="600" style="font-family: verdana; font-size: 11; color: #000000;">
		<p style="padding-bottom:3;">
		<a href="javascript:expand('Table1')" style="text-decoration: none; color: #448;"><img src='../images/plus.jpg' width="11" height="11" border="0" align="absmiddle" align="left" id="Table1_Img">&nbsp;<u>View HTML code</u></a>&nbsp;|&nbsp;
		<a href="javascript:expand('Table2')" style="text-decoration: none; color: #448;"><img src='../images/plus.jpg' width="11" height="11" border="0" align="absmiddle" align="left" id="Table2_Img">&nbsp;<u>View Text only</u></a>&nbsp;|&nbsp;
		<a href="javascript:refresh()" style="text-decoration: none; color: #448;"><img src='../images/refresh.gif' width="15" height="14" border="0" align="absmiddle">&nbsp;<u>Refresh</u></a>&nbsp;|&nbsp;
		<a href="<%= MapURL(linkURL) %>" target="_blank" style="text-decoration: none; color: #448;"><img src='../images/edit.gif' width="15" height="13" border="0" align="absmiddle">&nbsp;<u>Edit</u></a>&nbsp;|&nbsp;
		<a href="mailto:?subject=Please%20review%20email:%20%20<%=Left(folderStr,Len(folderStr)-Len("_docs"))%>&body=http://ctoohey02/email/previewer/render2.asp?filePath=<%=linkURL%>%26folder=<%=folderStr%>" target="_blank" style="text-decoration: none; color: #448;"><img src="../images/email_icon.gif" alt="" border="0" height="17" width="26" align="absmiddle"><u>eMail this page</u></a>&nbsp;|&nbsp;		
		<a href="help.html" target="_blank" style="text-decoration: none; color: #448;"><img src='../images/help.gif' width="15" height="15" border="0" align="absmiddle">&nbsp;<u>Help</u></a>
		</p>
		<div class="collapse" id="Table1" style="">
			<table border="0" bgcolor="#FFFFFF">
			<tr>
				<td align="left" style="font-family: verdana; font-size: 11; color: #000000;"><b>HTML Code</b><br>
					<form method="GET" name="" action="">
					<TEXTAREA cols="95" rows="35"><%= response.Write(strHTML) %></TEXTAREA>	
					</form>
				</td>
			</tr>
			</table>
		</div>
		<div class="collapse" id="Table2" style="">
			<table border="0" bgcolor="#FFFFFF">
			<tr>
				<td align="left" style="font-family: verdana; font-size: 11; color: #000000;"><b>Text Only</b><br>
					<form method="GET" name="" action="">
					<TEXTAREA cols="95" rows="35"><%= StripHTMLTags(strHTML) %></TEXTAREA>	
					</form>
				</td>
			</tr>
			</table>
		</div>
	</td>
</tr>
<tr>
	<td align="center"><%= strHTML %></td>
</tr>
</table>
<%end sub%>
<%'Fxn to load XSLT and display page contents on demand
sub XSLTLocatorVers(template)
	dim objXSL
	Set objXSL = Server.CreateObject("Microsoft.XMLDOM")
	objXSL.async = False
	objXSL.load(template)

	' Transform the XML file using the XSL stylesheet
	strHTML = objXML.transformNode(objXSL)
%>
<table border="0" cellpadding="0" cellspacing="1">
<tr>
	<td valign="top" align="left" width="600" style="font-family: verdana; font-size: 11; color: #000000;">
		
		<p style="padding-bottom:3;">
		<a href="javascript:expand('Table1')" style="text-decoration: none; color: #448;"><img src='../images/plus.jpg' width="11" height="11" border="0" align="absmiddle" align="left" id="Table1_Img">&nbsp;<u>View HTML code</u></a>&nbsp;|&nbsp;
		<a href="javascript:expand('Table2')" style="text-decoration: none; color: #448;"><img src='../images/plus.jpg' width="11" height="11" border="0" align="absmiddle" align="left" id="Table2_Img">&nbsp;<u>View Text only</u></a>&nbsp;|&nbsp;
		<a href="javascript:refresh()" style="text-decoration: none; color: #448;"><img src='../images/refresh.gif' width="15" height="14" border="0" align="absmiddle">&nbsp;<u>Refresh</u></a>&nbsp;|&nbsp;
		<a href="<%= MapURL(linkURL) %>" target="_blank" style="text-decoration: none; color: #448;"><img src='../images/edit.gif' width="15" height="13" border="0" align="absmiddle">&nbsp;<u>Edit</u></a>&nbsp;|&nbsp;
		<a href="mailto:?subject=Please%20review%20email:%20%20<%=Left(folderStr,Len(folderStr)-Len("_docs"))%>&body=http://ctoohey02/email/previewer/render2.asp?filePath=<%=linkURL%>%26folder=<%=folderStr%>" target="_blank" style="text-decoration: none; color: #448;"><img src="../images/email_icon.gif" alt="" border="0" height="17" width="26" align="absmiddle"><u>eMail this page</u></a>&nbsp;|&nbsp;		
		<a href="help.html" target="_blank" style="text-decoration: none; color: #448;"><img src='../images/help.gif' width="15" height="15" border="0" align="absmiddle">&nbsp;<u>Help</u></a>
		</p>
		<div class="collapse" id="Table1" style="">
			<table border="0" bgcolor="#FFFFFF">
			<tr>
				<td align="left" style="font-family: verdana; font-size: 11; color: #000000;"><b>HTML Code</b><br>
					<form method="GET" name="" action="" ID="Form1">
					<TEXTAREA cols="95" rows="35"><%= response.Write(strHTML) %></TEXTAREA>	
					</form>
				</td>
			</tr>
			</table>
		</div>
		<div class="collapse" id="Table2" style="">
			<table border="0" bgcolor="#FFFFFF" ID="Table6">
			<tr>
				<td align="left" style="font-family: verdana; font-size: 11; color: #000000;"><b>Text Only</b><br>
					<form method="GET" name="" action="" ID="Form2">
					<TEXTAREA cols="95" rows="35"><%= StripHTMLTags(strHTML) %></TEXTAREA>	
					</form>
				</td>
			</tr>
			</table>
		</div>
	</td>
</tr>
<tr>
	<td align="center"><%= strHTML %></td>
</tr>
</table>
<%end sub%>
</body>	
</html>
<%
function MapURL(path)
	dim rootPath, url
	rootPath = Server.MapPath("/")
	url = Right(path, Len(path) - Len(rootPath))
	MapURL = Replace(url, "\", "/")
	'response.Write(path)
end function

function BreadCrumb(path)
	dim rootPath, url, url2
	rootPath = Server.MapPath("/")
	url = Right(path, Len(path) - Len(rootPath))
	url2 = Left(url, Len(url) - Len(".xml"))
	BreadCrumb = Replace(url2, "\", " -> ")
	'response.Write(path)
end function

'This function uses Regular Expressions to strip HTML tags from a string  
Public Function StripHTMLTags(HTMLString)

Set RegularExpressionObject = New RegExp

With RegularExpressionObject
 .Pattern = "<[^>]+>"
 .IgnoreCase = True
 .Global = True
End With

StripHTMLTags = RegularExpressionObject.Replace(HTMLString, "")

Set RegularExpressionObject = nothing

End Function

'This function uses Regular Expressions to replace ASCII symbol code in a string 
Public Function ReplaceSymbols(HTMLString)

Set RegularExpressionObject = New RegExp

With RegularExpressionObject
 .Pattern = "<[ ]*[&amp;]*(#133;|#151;|#174;|#169;|nbsp;|lt;|gt;)*>"
 .IgnoreCase = True
 .Global = True
End With

ReplaceSymbols = RegularExpressionObject.Replace(HTMLString, "&")

Set RegularExpressionObject = nothing

End Function

'Report.EnableSafeEncoding = true


Set objXML = Nothing
Set objXSL = Nothing

%>

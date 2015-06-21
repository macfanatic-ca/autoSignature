--This script will read attributes from Active Directory (AD), place them into an HTML template, import HTML into Microsoft Outlook (2013 & 2016 tested).  A connection to AD is manditory for the information to populate.
--Created on February 24th, 2015
--Written by Jon Yergatian
--V1.1 - March 12, 2015
--Added workflow which exports to an html file on the users desktop and opens it in safari for a more manual process. Great for webmail users.
--V1.1.2 - April 09, 2015
--Renamed and corrected Identifier
--V1.1.3 - May 06, 2015
--Fixed a bug which would ignore any JobTitle not containing a space, such as CEO
--V1.1.5 - May 07, 2015
--Fixed a syntax error causing an invalid mailto link being generated.

--Set loction of applet.icns
set appicon to path to resource "applet.icns"

--Set location of image
set pathtoimg to "http://www.company.com/images/signatureimage.png"

--Set Company Name
set companyname to "comany name here"

--Set Company Website
set companysite to "http://www.company.com"

--Grabs current username
set user to do shell script "whoami"

--Reads specifed AD attributes and stores relevant info
set fullname to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " RealName | awk -F 'RealName:' '{ print $1 }'"

set email to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " EMailAddress | awk '{ print $2 }'"

set phonenumber to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " PhoneNumber | awk -F 'PhoneNumber:' '{ print $1 }'"

set ext to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " ipPhone | awk -F 'dsAttrTypeNative:ipPhone:' '{ print $2 }'"

set jobtitle to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " JobTitle | awk -F 'JobTitle:' '{ print $1 }'"
if jobtitle = "" then
	set jobtitle to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read Users/" & user & " JobTitle | awk '{ FS = \":[ 	]*|[ 	]+\" }''{ print $2 }'"
end if

set street to do shell script "dscl '/Active Directory/DOMAIN/search.path' -read /Users/" & user & " Street | awk -F 'Street:' '{ print $1 }'"

--Additional Information
display dialog "You're about to create a new signature.
Would you like to manually add your Skype and Mobile contact information?" with icon appicon buttons {"Yes", "No", "Cancel"} default button 2
if result = {button returned:"Yes"} then
	display dialog "Please enter your Skype information
	Format: all lowercase" default answer "username" with icon appicon
	set skypeName to text returned of result
	display dialog "Please enter your Mobile information
	Format: 1 (123) 456-7890" default answer "1 (123) 456-7890" with icon appicon
	set mobileNumber to text returned of result
	set sigHTML to "<html>
	<head>
		<meta http-equiv=\"Content-Type\" content=\"text/html; charset=iso-8859-1\">
	</head>
    <body>
    </body>
	<body>
<table table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\" style=\"font-family:Lucida Sans, Lucida Sans Unicode, Verdana, sans-serif; text-align:left; font-size:8px; padding:0px;\" class=\"webfont\">
<tr>
</tr>
<tr>
<hr size=\"2\" color=\"#00afd0\" width=\"100%\" align=\"LEFT\">
</tr>
<tr height=\"2\">
<td></td>
</tr>
<tr>
<td width=\"679\"><img src=\"" & pathtoimg & "\" width=\"180\" height=\"33\" alt=\"" & companyname & "\"></td>
</tr>
<tr height=\"10\">
<td></td>
</tr>
<tr>
<td><span style=\"font-size:14px; color:#555555\">" & fullname & " | <span style=\"color:#00aed0\">" & jobtitle & "</span></span></td>
</tr>
<tr height=\"10\">
<td></td>
</tr>
<tr>
<td><p style=\"font-size:11px;margin: 0 0 5px 0;\"><span style=\"color:#00afd0;\">" & phonenumber & " ext." & ext & "  | <a style=\"color:#00afd0; text-decoration:none;\" href=mailto:" & email & ">" & email & "</a> | <a style=\"color:#00afd0; text-decoration:none;\" href=\"" & companysite & "" title=\" & companyname & \">" & companysite & "</a></span></p>
<p style=\"font-size:11px;margin: 0 0 5px 0;\"><span style=\"color:#00afd0;\">" & mobileNumber & " (mobile) | <a style=\"color:#00afd0; text-decoration:none;\" >" & skypeName & " (skype)</a></p>
<p style=\"font-size:11px;margin:0 0 10px 0;\"><span style=\"font-size:11px; color:#555555\">" & companyname & " | " & street & "</span></p></td>
</tr>
<tr height=\"10\">
<td></td>
</tr>
</table>
</body>
</html>"
else if result = {button returned:"No"} then
	set sigHTML to "<html>
	<head>
		<meta http-equiv=\"Content-Type\" content=\"text/html; charset=iso-8859-1\">
	</head>
    <body>
    </body>
	<body>
<table table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\" style=\"font-family:Lucida Sans, Lucida Sans Unicode, Verdana, sans-serif; text-align:left; font-size:8px; padding:0px;\" class=\"webfont\">
<tr>
</tr>
<tr>
<hr size=\"2\" color=\"#00afd0\" width=\"100%\" align=\"LEFT\">
</tr>
<tr height=\"2\">
<td></td>
</tr>
<tr>
<td width=\"679\"><img src=\"" & pathtoimg & "\" width=\"180\" height=\"33\" alt=\"" & comanyname & "\"></td>
</tr>
<tr height=\"10\">
<td></td>
</tr>
<tr>
<td><span style=\"font-size:14px; color:#555555\">" & fullname & " | <span style=\"color:#00aed0\">" & jobtitle & "</span></span></td>
</tr>
<tr height=\"10\">
<td></td>
</tr>
<tr>
<td><p style=\"font-size:11px;margin: 0 0 5px 0;\"><span style=\"color:#00afd0;\">" & phonenumber & " ext." & ext & "  | <a style=\"color:#00afd0; text-decoration:none;\" href=mailto:" & email & ">" & email & "</a> | <a style=\"color:#00afd0; text-decoration:none;\" href=\"" & companysite & "" title=\"" & companyname & "\">" & companysite & "</a></span></p>
<p style=\"font-size:11px;margin:0 0 10px 0;\"><span style=\"font-size:11px; color:#555555\"> " & companyname & " | " & street & "</span></p></td>
</tr>
<tr height=\"10\"
<td></td>
</tr>
</table>
</body>
</html>"
end if

--Allow user to choose between web or app
display dialog "Do you prefer using the Outlook App or the Outlook Webmail?" with icon appicon buttons {"App", "Webmail"} default button 1
--Proceed with App
if result = {button returned:"App"} then
	--Offer user option to name signature
	display dialog "You may want to provide a unique name for your new signature, especially if you have multiple already. " default answer "My New Signature" buttons {"Cancel", "Lets Go!"} with icon appicon default button 2
	set sigtitle to text returned of result
	tell application "AppleScript Utility"
		--Adds AD info to HTML template > forwards to Microsoft Outlook (2013 & 2016 tested)
		make new «class cSig» with properties {name:"" & sigtitle & "", «class ctnt»:"" & sigHTML & ""}
	end tell
	--Directs user to setting default signature
	display dialog "Almost Done!

Please navigate to Outlook > Preferences > Signatures

Set your new signature as Default for:

	New Messages

	Replies/forwards" with icon appicon buttons {"Got it"} default button 1
else if result = {button returned:"Webmail"} then
	--Give user instructions on what to do next.
	display dialog "Your new signature will open in Safari
Copy and Paste it into the Email Signature field within Outlook Webmail" with icon appicon buttons {"Got it"} default button 1
	--Set location for html file
	set fileLocation to (path to desktop as text) & "New Signature.html"
	--Write HTML file using sigHTML above
	set htmlFile to open for access fileLocation with write permission
	write sigHTML to htmlFile
	close access htmlFile
	--Open HTML file in Safari for preview and Copy/Paste
	tell application "Safari"
		run
		activate
		set frontmost to true
		open file fileLocation
	end tell
end if

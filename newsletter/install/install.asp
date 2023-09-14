<% 
on error resume next
    Sub trace(strText)
	    Response.Write "Debug: "&strText&"<br />"&vbcrlf
    End Sub

    Sub catch(sText,sText2)

	    If Err.Number <> 0 then
		    Call trace(sText&" - "&err.description)
	    Else
		    Call trace(sText&" - no error")
	    End If
	    If sText2 <> "" Then
		    Call trace(sText&" - "&sText2)
	    End If

	    on error goto 0
    End Sub

    Const ForReading = 1
    Const TristateUseDefault = -2

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = DBvalue

		If fieldvalue <> "" AND Not IsNull(fieldvalue) Then
		
			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((delete)*(select)*(update)*(into)*(drop)*(insert)*(declare)*(xp_)*(union)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue=replace(fieldvalue,"'","''")

		End If

		DBEncode = fieldvalue

	End Function    
%>
<!DOCTYPE HTML>
<html>
<head>
    <title>Install</title>
    <link type="text/css" rel="stylesheet" href="../assets/css/main.css" />
</head>
<body>
    <div id="main" class="container" align="center" style="margin-top: -75px;">
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <header>
                    <h2>EZNewsletter Installation</h2>
                </header>
            </div>
        </div>
    </div>
    <%
  If Trim(Request.QueryString("step")) = "one" Then
    %>
    <div id="main" class="container" align="center" style="margin-top: -100px;">
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <form action="install.asp?setsql=y" method="post">

                    <header>
                        <h2>MSSQL Database</h2>
                    </header>
                    <div class="row">
                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="svrname" style="text-align: left;">Server Host Name or IP Address</label>
                            <input type="text" id="svrname" name="svrname" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="dbname" style="text-align: left;">Database Name</label>
                            <input type="text" id="dbname" name="dbname" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="dbid" style="text-align: left;">Database Login</label>
                            <input type="text" id="dbid" name="dbid" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="dbpwd" style="text-align: left;">Database Password</label>
                            <input type="password" id="dbpwd" name="dbpwd" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="dbprefix" style="text-align: left;">Table Prefix</label>
                            <input type="text" id="dbprefix" name="dbprefix" value="eznws_" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="12u 12u$(medium)">
                            <input class="button" type="submit" name="submit" value="Continue">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <%
  ElseIf Request.QueryString("setsql") = "y" Then

    %>
    <div id="main" class="container" align="center">
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <%

    msdbserver = Trim(Request.Form("svrname"))
    msdb = Trim(Request.Form("dbname"))
	msdbid = Trim(Request.Form("dbid"))
    msdbpwd = Trim(Request.Form("dbpwd"))
    msdbprefix = Trim(Request.Form("dbprefix"))

    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Set rsCommon = Server.CreateObject("ADODB.Recordset")

    Response.Write "Creating Database Tables<br /><br />"
	Response.Write "Creating admin table...<br />"
	Response.Flush

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Conn.Execute "CREATE TABLE "&msdbprefix&"admin (" & _
    "[adminID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_admin] PRIMARY KEY," & _
    "[name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[pwd] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[salt] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[send]  [bit] NULL ," & _
    "[addresses]  [bit] NULL ," & _
    "[images]  [bit] NULL ," & _
    "[templates]  [bit] NULL ," & _
    "[options]  [bit] NULL ," & _
    "[db_rights]  [bit] NULL ," & _
    "[admins_rights]  [bit] NULL ," & _
    "[arights]  [bit] NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating admin table...<br />"
	Response.Flush

    Conn.Execute "INSERT INTO "&msdbprefix&"admin ([name],[pwd],[salt],[send],[addresses],[images],[templates],[options],[db_rights],[admins_rights],[arights]) VALUES ('admin','EB36FB0C1F1A92A838AA1ECAAD4AB6E3B5257103','833D1','True','True','True','True','True','True','True','True')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating settings table...<br />"
    Response.Flush

    Conn.Execute "CREATE TABLE "&msdbprefix&"settings (" & _
    "[settingID] [numeric](10, 0) IDENTITY (1, 1) CONSTRAINT [PK_settings] PRIMARY KEY," & _
    "[site_title] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[domain_name] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[smtp_server] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[email_address] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[smtp_password] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[smtp_port] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[rewrite] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[confirm_email] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
    "[aspupload]  [bit] NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating Admin Messages table...<br />"
    Response.Flush

    Conn.Execute "CREATE TABLE "&msdbprefix&"messages " & _
    "([messageID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_messages] PRIMARY KEY," & _
    "[msg] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[message] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating Admin Messages table...<br />"
	Response.Flush

	Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('eas','The email was successfully added  to the database!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('aid','That Email Address is already in the database!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ds','The delete action was successful!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nea','You forgot to enter an email address!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('uls','File(s) uploaded successful!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ids','File(s) deleted successful!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nadmin','You can not change Admins info.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('del','Template deleted!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('das','You successfully deleted the Admin.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('adad','You have successfully added an Admin.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nt','That name has been taken.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('tc','Template Created!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('car','You have successfully modified Admin Rights.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nwst','Newsletter Sent!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ulf','Upload failed!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ant','Admin name taken!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('ftna','Sorry, only JPG, PNG & GIF files are allowed.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('fex','File already exists!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nimg','File is not an image!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('tus','Template updated successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('mus','Messages updated successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('error','An unknown error has occurred<br>Please contact support!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('siu','Site info updated successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('cpwds','Password changed successfully!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"messages([msg],[message]) VALUES('nar','You do not have sufficient rights to view this page!')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating User Messages table...<br />"
    Response.Flush

    Conn.Execute "CREATE TABLE "&msdbprefix&"endMsg " & _
    "([endMsgID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_endMsg] PRIMARY KEY," & _
    "[endMsgName] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[endMsg] [nvarchar] (MAX) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating User Messages table...<br />"
	Response.Flush

	Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('thanks','<h2>Thank you!</h2><br />#EMAIL# was added to our list<br />A confirmation email was sent please follow the instructions in it.<br />Be sure to check your Junk/Spam folder!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('confirmed','<h2>Success!</h2><br />#EMAIL# has been confirmed.<br /><br />Thank you for subscribing to the #SITETITLE# newsletter.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('confirmerr','<h2>Sorry!</h2><br />There was a problem and we could not confirm #EMAIL#<br />Please try again or contact support.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('alreadysubbed','<h2>OOPS!</h2><br />It seems that #EMAIL# is already subscribed!<br />While we appreciate your enthusiasm you can only subscribe once!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('thankserr','<h2>Sorry?</h2><br />We were able to add #EMAIL# to our list. But could not send the confirmation email.<br />Please contact support!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('adderr','<h2>Sorry!</h2><br />There was a problem and we could not add #EMAIL# to our list.<br />Please try again or contact support.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('notfound','<h2>Sorry!</h2><br />We could not find #EMAIL# in our database.<br />Please contact support.')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('removed','<h2>Sorry!</h2>There was a problem and we could not remove #EMAIL# from our list.<br />Please contact support!')"
    Conn.Execute "INSERT INTO "&msdbprefix&"endMsg([endMsgName],[endMsg]) VALUES('canceled','<h2>Success!</h2><br />#EMAIL# was removed from our list.<br />We are sorry to see you go!<br />Thanks<br />The #SITETITLE# Newsletter team')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

	Response.Write "Creating Newsletter Addresses table...<br />"
	Response.Flush

    Conn.Execute "CREATE TABLE "&msdbprefix&"addresses " & _
    "([NewsID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_addresses] PRIMARY KEY," & _
    "[email] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[datDate] [smalldatetime] NULL ," & _
    "[confirm] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[token] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating Newsletter table...<br />"
    Response.Flush

    Conn.Execute "CREATE TABLE "&msdbprefix&"newsletter " & _
    "([newsletterID] [numeric] IDENTITY (1, 1) CONSTRAINT [PK_newsLetter] PRIMARY KEY," & _
    "[news_title] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[news_save] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[news_description] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
    "[news_body] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
    ")"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Populating newsletter table...<br />"
    Response.Flush

    Conn.Execute "INSERT INTO "&msdbprefix&"newsletter([news_title],[news_save],[news_description],[news_body]) VALUES('Template One','template','A simple starting point!','<header><h2 style=""text-align:center;"">Template One</h2></header><content><div style=""text-align:center""><article><span style=""font-size:16px;""> Hello World!</span></article></div></content><footer><div style=""text-align:center"">&copy; 2017 All Rights Reserved</div></footer>')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Conn.Execute "INSERT INTO "&msdbprefix&"newsletter([news_title],[news_save],[news_description],[news_body]) VALUES('Draft one','draft','Another Simple Starting Point.','<h2>Draft one</h2><table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""width:500px""><tbody><tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td></tr></tbody></table>')"

	If Err.Number <> 0 then
		response.Write "Error: "&err.description&"<br />"
	End If
	on error goto 0

    Response.Write "Creating database tables...Complete!<br />"
    Response.Flush

    Response.Write "<br /><br />"
                %>
            </div>
        </div>
    </div>
    <div id="main" class="container" align="center">
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <form action="install.asp?step=two" method="post">
                    <input type="hidden" name="msdbserver" value="<%= msdbserver %>">
                    <input type="hidden" name="msdb" value="<%= msdb %>">
                    <input type="hidden" name="msdbid" value="<%= msdbid %>">
                    <input type="hidden" name="msdbpwd" value="<%= msdbpwd %>">
                    <input type="hidden" name="msdbprefix" value="<%= msdbprefix %>">
                    <header>
                        <h3><span class="first">You have successfully installed the MSSQL Database<br />
                            Please click the button below to continue</span></h3>
                    </header>
                    <div class="row">
                        <div class="12u 12u$(medium)">
                            <input class="button" type="submit" name="submit" value="Continue">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <%
		Conn.Close: Set Conn = Nothing

  ElseIf Request.QueryString("step") = "two" Then
    %>
    <div id="main" class="container" align="center">
        <div class="row 50%">
            <div class="12u 12u$(medium)">
                <form action="install.asp?step=three" method="post">
                    <input type="hidden" name="msdbserver" value="<%= Trim(Request.Form("msdbserver")) %>">
                    <input type="hidden" name="msdb" value="<%= Trim(Request.Form("msdb")) %>">
                    <input type="hidden" name="msdbid" value="<%= Trim(Request.Form("msdbid")) %>">
                    <input type="hidden" name="msdbpwd" value="<%= Trim(Request.Form("msdbpwd")) %>">
                    <input type="hidden" name="msdbprefix" value="<%= Trim(Request.Form("msdbprefix")) %>">
                    <input type="hidden" name="PhyPath" value="<%= strPhysPath %>" />
                    <header>
                        <h2>Path Settings</h2>
                    </header>
                    <div class="row">
                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="bdir" style="text-align: left;">Base Directory</label>
                            <input type="text" id="bdir" name="bdir" value="<%= Request.ServerVariables("APPL_PHYSICAL_PATH") %>" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="dir" style="text-align: left;">EZNewsletter Directory</label>
                            <input type="text" id="dir" name="dir" value="/newsletter/" size="40" />
                        </div>
                        <div class="4u 1u$"><span></span></div>
                        <div class="12u 12u$(medium)">
                            <input class="button" type="submit" name="submit" value="Continue">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <%
  ElseIf Request.QueryString("step") = "three" Then

    strPageFileName = Server.MapPath("../includes/config.asp")

    Set objPageFileFSO = CreateObject("Scripting.FileSystemObject")

    If objPageFileFSO.FileExists(strPageFileName) Then
        Set objPageFileTs = objPageFileFSO.OpenTextFile(strPageFileName, 2)
    Else
        Set objPageFileTs = objPageFileFSO.CreateTextFile(strPageFileName)
    End If

    strPageEntry = Chr(60) & Chr(37) & vbcrlf & _
    "baseDir=""" & Trim(Request.Form("bdir")) & """" & vbcrlf & _
    "strDir=""" & Trim(Request.Form("dir")) & """" & vbcrlf & _
    "msdbprefix=""" & Trim(Request.Form("msdbprefix")) & """" & vbcrlf & _
    "msdbserver=""" & Trim(Request.Form("msdbserver")) & """" & vbcrlf & _
    "msdb=""" & Trim(Request.Form("msdb")) & """" & vbcrlf & _
    "msdbid=""" & Trim(Request.Form("msdbid") )& """" & vbcrlf & _
    "msdbpwd=""" & Trim(Request.Form("msdbpwd")) & """" & vbcrlf & _
    Chr(37) & Chr(62)

    objPageFileTs.WriteLine strPageEntry

    objPageFileTs.Close

    Response.Redirect "install.asp?step=four"

  ElseIf Request.QueryString("step") = "four" Then
    %>
    <div id="main" class="container" style="margin-top: -100px;">
        <div class="row">
            <div class="12u 12u$(medium)" style="text-align: center;">
                <form action="install.asp?step=five" method="post">
                    <header>
                        <h2>Other stuff</h2>
                    </header>
                    <div class="row">

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="sitetitle" style="text-align: left;">Site title</label>
                            <input type="text" id="sitetitle" name="sitetitle" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="domainname" style="text-align: left;">Domain name</label>
                            <input type="text" id="domainname" name="domainname" value="<%= Request.ServerVariables("SERVER_NAME") %>" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="smtpserver" style="text-align: left;">SMTP Server <span style="font-size: 12px;">(If your not sure leave it blank)</span></label>
                            <input type="text" id="smtpserver" name="smtpserver" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="smtpserver" style="text-align: left;">SMTP Port <span style="font-size: 12px;">(If your not sure leave it blank)</span></label>
                            <input type="text" id="smtpport" name="smtpport" value="587" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="emailaddress" style="text-align: left;"> Your Email Address <span style="font-size: 12px;">(If your not sure leave it blank)</span></label>
                            <input type="text" id="emailaddress" name="emailaddress" />
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <label for="smtppwd" style="text-align: left;">SMTP Password <span style="font-size: 12px;">(If your not sure leave it blank)</span></label>
                            <input type="password" id="smtppwd" name="smtppwd" required>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="-4u 4u 12u$(medium)" style="padding-bottom: 20px;">
                            <div class="row" style="padding-bottom: 20px;">
                                <header><strong>Do you have ASPUpload installed <span style="font-size: 12px;">(If your not sure check no)</span></strong></header>
                                <div class="6u 12u$(medium)">
                                    <input type="radio" id="upload-yes" name="aspupload" value="on">
                                    <label for="upload-yes">Yes</label>
                                </div>
                                <div class="6u 12u$(medium)" style="padding-bottom: 20px;">
                                    <input type="radio" id="upload-no" name="aspupload" value="off" checked>
                                    <label for="upload-no">No</label>
                                </div>
                            </div>
                        </div>
                        <div class="4u 1u$"><span></span></div>

                        <div class="12u 12u$(medium)">
                            <input class="button" type="submit" name="submit" value="Continue">
                        </div>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <%
  ElseIf Request("step") = "five" Then
    %><!-- #include file="../includes/config.asp"--><%
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    blnAspupload = False
    If Trim(Request.Form("aspupload")) = "on" Then blnAspupload = True
    strConfirmEmail = "<p>Thank you #EMAIL# for subscribing to the #SITETITLE# Newsletter<br /><br />Please confirm your subscription by clicking on the link below.<br /><br />#CONFIRMREWRITE#<br /><br />You received this email because someone&nbsp;submitted this email address to our mailing list.<br />If you did not subscribe or wish to be removed from our list - click on the link below<br /><br />#CANCELREWRITE#<br /><br />Our Thanks<br />#SITETITLE#</p><p style=""text-align:center""><span style=""font-size:12px"">#CR# Copyright #YEAR# #SITETITLE# |&nbsp;<a href="""" target=""_blank"">Privacy Policy</a></span></p>"
    Conn.Execute "INSERT INTO "&msdbprefix&"settings ([site_title],[domain_name],[smtp_server],[email_address],[smtp_password],[smtp_port],[rewrite],[confirm_email],[aspupload]) VALUES ('"&DBEncode(Request.Form("sitetitle"))&"','"&DBEncode(Request.Form("domainname"))&"','"&DBEncode(Request.Form("smtpserver"))&"','"&DBEncode(Request.Form("emailaddress"))&"','"&DBEncode(Request.Form("smtppwd"))&"','"&Trim(Request.Form("smtpport"))&"','1','"&strConfirmEmail&"','"&blnAspupload&"')"

    Conn.Close: Set Conn = Nothing

    Response.Redirect "install.asp?step=done"

  ElseIf Request("step") = "done" Then
    %>
    <div id="main" class="container">
        <div class="row">
            <div class="12u 12u$(medium)" style="text-align: center;">
                <span class="first">Success!
                    <br />
                    You have successfully configured EZNewsletter!
                    <br />
                    The next step is to change your password.
                    <br />
                    Click on the link below and login to admin.
                    <br />
                    Click on "Password" in the left options menu and change your password.
                    <br /><br />
                    <a class="first" href="../admin/login.asp">Login</a>
                </span>
            </div>
        </div>
    </div>
    <% Else %>
    <div id="main" class="container" style="margin-top: -75px;">
        <div class="row">
            <div class="12u 12u$(medium)" style="text-align: center;">
                <span class="first">You are about to install EZNewsletter.
	               <br>
                    Please follow the instructions carefully!
	                <br>
                    <br>
                    <input class="button" type="button" onclick="parent.location='install.asp?step=one'" value="Continue">
                    <br>
                    <br>
                </span>
            </div>
        </div>
    </div>
    <% End If %>
    <br />
</body>
</html>
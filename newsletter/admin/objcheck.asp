<%@Language="VBScript"%>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

on error resume next
response.buffer = true
server.scripttimeout = 1000

installedCOMs = 0
onNum = 0
lastUpdate = "4/3/2003"
newVersion = False

' The Components
' format: comObject|comURL|comName|comCategory|comCategory2
com = "CDONTS.NewMail|http://www.microsoft.com|CDONTS (free)|1|"
com = com & vbnewline & "MSWC.NextLink|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp7pmc.asp|Microsoft Content Linking Component|0|"
com = com & vbnewline & "MSWC.BrowserType|http://msdn.microsoft.com/library/default.asp?url=/library/en-us/iisref/html/psdk/asp/comp3xx0.asp|Microsoft Browser Capability|2|"
com = com & vbnewline & "MSWC.ContentRotator|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp09dg.asp|Microsoft Content Rotator|0|"
com = com & vbnewline & "MSWC.AdRotator|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp59f8.asp|Microsoft Ad Rotator|0|"
com = com & vbnewline & "MSWC.PermissionChecker|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp3hf8.asp|Microsoft Permission Checker Component|0|"
com = com & vbnewline & "MSWC.Status|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp1qt0.asp|Microsoft Status Component|0|"
com = com & vbnewline & "MSWC.Tools|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp7g8k.asp|Microsoft Tools Component|0|"
com = com & vbnewline & "MSWC.PageCounter|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp00vo.asp|Microsoft Page Counter Component|0|"
com = com & vbnewline & "MSWC.IISLog|http://msdn.microsoft.com/library/en-us/iisref/html/psdk/asp/comp6i5w.asp|Microsoft Logging Utility Component|0|"
com = com & vbnewline & "MSXML2.ServerXMLHTTP|http://msdn.microsoft.com/library/en-us/xmlsdk30/htm/xmobjxmldomserverxmlhttp_using_directly.asp|Microsoft ServerXMLHTTP|13|"
com = com & vbnewline & "Microsoft.XMLDOM|http://www.microsoft.com|Microsoft XMLDOM Component|13|"
com = com & vbnewline & "Microsoft.XMLHTTP|http://www.microsoft.com|Microsoft XMLHTTP Component|13|"
com = com & vbnewline & "Scripting.FileSystemObject|http://www.microsoft.com|MicrosoftFileSystem Object|6|"
com = com & vbnewline & "ADOX.Catalog|http://www.microsoft.com|MicroSoft ADOX Catalog|0|"
com = com & vbnewline & "WScript.Shell|http://www.microsoft.com|Windows Script Shell|0|"
com = com & vbnewline & "WScript.Network|http://www.microsoft.com|Windows Script Network|0|"
com = com & vbnewline & "ADODB.Connection|http://www.microsoft.com|ADODB.Connection|0|"
com = com & vbnewline & "ADODB.Command|http://www.microsoft.com|ADODB.Command|0|"
com = com & vbnewline & "ADODB.Recordset|http://www.microsoft.com|ADODB.Recordset|0|"
com = com & vbnewline & "Scripting.Dictionary|http://www.microsoft.com|Scripting.Dictionary|0|"
com = com & vbnewline & "ASPFileUpload.File|http://support.microsoft.com/default.aspx?scid=kb;EN-US;q299692|MicroSoft File Upload|3|0"
com = com & vbnewline & "Scripting.Encoder|http://www.microsoft.com|Script Encoder|0|"
com = com & vbnewline & "Msxml2.DOMDocument.3.0|http://www.microsoft.com|Microsoft XMLDOM 3.0 Component|13|"
com = com & vbnewline & "Msxml2.DOMDocument.4.0|http://msdn.microsoft.com/downloads/default.asp?url=/downloads/topic.asp?url=/msdn-files/028/000/072/topic.xml|Microsoft XMLDOM 4.0 Component|13|"
com = com & vbnewline & "SMTPsvg.Mailer|http://www.serverobjects.com|Server Objects - ASPMail|1|"
com = com & vbnewline & "SMTPsvg.Mailer|http://www.serverobjects.com|Server Objects - ASPQMail|1|"
com = com & vbnewline & "AspImage.Image|http://www.serverobjects.com|Server Objects - ASPImage|4|"
com = com & vbnewline & "POP3svg.Mailer|http://www.serverobjects.com|Server Objects - ASPPop3|1|"
com = com & vbnewline & "AspNNTP.Conn|http://www.serverobjects.com|Server Objects - AspNNTP|0|"
com = com & vbnewline & "AspFile.FileObj|http://www.serverobjects.com|ServerObjects - AspFile|6|"
com = com & vbnewline & "AspConv.Expert|http://www.serverobjects.com|ServerObjects - AspConv|0|"
com = com & vbnewline & "AspHTTP.Conn|http://www.serverobjects.com|ServerObjects - AspHTTP|0|"
com = com & vbnewline & "AspDNS.Lookup|http://www.serverobjects.com|ServerObjects - AspDNS|0|"
com = com & vbnewline & "AspMX.Lookup|http://www.serverobjects.com|ServerObjects - AspMX|1|"
com = com & vbnewline & "WaitFor.Comp|http://www.serverobjects.com|ServerObjects - Waitfor (free)|0|"
com = com & vbnewline & "LastMod.FileObj|http://www.serverobjects.com|ServerObjects - Last Modified (free)|6|"
com = com & vbnewline & "ImgSize.Check|http://www.serverobjects.com|ServerObjects - Image Size (free)|4|"
com = com & vbnewline & "GuidMakr.GUID|http://www.serverobjects.com|ServerObjects - GUID Maker (free)|0|"
com = com & vbnewline & "ASPsvg.Process|http://www.serverobjects.com|ServerObjects - AspProc (free)|0|"
com = com & vbnewline & "AspPing.Conn|http://www.serverobjects.com|ServerObjects - AspPing (free)|0|"
com = com & vbnewline & "AspInet.FTP|http://www.serverobjects.com|ServerObjects - AspInet (free)|0|"
com = com & vbnewline & "ASPExec.Execute|http://www.serverobjects.com|ServerObjects - AspExec (free)|0|"
com = com & vbnewline & "AspCrypt.Crypt|http://www.serverobjects.com|ServerObjects - AspCryp (free)|9|"
com = com & vbnewline & "Bible.Lookup|http://www.serverobjects.com|ServerObjects - AspBible (free)|0|"
com = com & vbnewline & "SoftArtisans.SAFile|http://www.softartisans.com|SoftArtisians Fileup|3|"
com = com & vbnewline & "SoftArtisans.FileManager|http://www.softartisans.com|SoftArtisians FileManager|6|"
com = com & vbnewline & "SoftArtisans.XFRequest|http://www.softartisans.com|SoftArtisians X-File|6|"
com = com & vbnewline & "SoftArtisans.FileManagerTX|http://www.softartisans.com|SoftArtisians FileManagerTX|6|"
com = com & vbnewline & "SoftArtisans.SASessionPro.1|http://www.softartisans.com|SoftArtisans SA-Session Pro|0|"
com = com & vbnewline & "SMUM.XCheck.1|http://www.softartisans.com|SoftArtisians Check (form validator)|11|"
com = com & vbnewline & "Softartisans.Archive|http://www.softartisans.com|SoftArtisans Archive|6|"
com = com & vbnewline & "SoftArtisans.SMTPMail|http://www.softartisans.com|SoftArtisans SMTPmail|1|"
com = com & vbnewline & "Softartisans.ExcelWriter|http://www.softartisans.com|SoftArtisans Excel Writer|5|"
com = com & vbnewline & "SoftArtisans.Groups|http://www.softartisans.com|SoftArtisans.Groups (SA-Admin)|9|"
com = com & vbnewline & "SoftArtisans.Performance|http://www.softartisans.com|SoftArtisians.Performance (SA-Admin)|9|"
com = com & vbnewline & "SoftArtisans.RAS|http://www.softartisans.com|SoftArtisans.RAS (SA-Admin)|9|"
com = com & vbnewline & "SoftArtisans.Shares|http://www.softartisans.com|SoftArtisans.Shares (SA-Admin)|9|"
com = com & vbnewline & "SoftArtisans.User|http://www.softartisans.com|SoftArtisans.User (SA-Admin)|9|"
com = com & vbnewline & "Jmail.smtpmail|http://www.dimac.net|w3 JMail|1|"
com = com & vbnewline & "w3sitetree.tree|http://www.dimac.net|w3 Site Tree : www.dimac.net|0|"
com = com & vbnewline & "w3.upload|http://www.dimac.net|w3 Upload|3|"
com = com & vbnewline & "w3.netutils|http://www.dimac.net|w3 Utils|0|"
com = com & vbnewline & "Socket.TCP|http://www.dimac.net|w3 Sockets|0|"
com = com & vbnewline & "w3.netutils|http://www.dimac.net|w3 NetDebug|0|"
com = com & vbnewline & "Persits.MailSender|http://www.persits.com|Persits - ASPEmail|1|"
com = com & vbnewline & "Persits.Upload.1|http://www.persits.com|Persits - ASPUpload|3|"
com = com & vbnewline & "Persits.Jpeg|http://www.persits.com|Persits - AspJpeg|4|"
com = com & vbnewline & "Persits.Grid|http://www.persits.com|Persits - AspGrid|0|"
com = com & vbnewline & "Persits.AspUser|http://www.persits.com|Persits - AspUser|9|"
com = com & vbnewline & "Persits.CryptoManager|http://www.persits.com|Persits - AspEncrypt|9|"
com = com & vbnewline & "ADISCON.SimpleMail.1|http://www.simplemail.adiscon.com/en|SimpleMail|1|"
com = com & vbnewline & "CalendarCom.CalendarStuff|http://www.devguru.com|DevGuru - dgcalendar|0|"
com = com & vbnewline & "dgEncrypt.Key|http://www.devguru.com|DevGuru - dgEncrypt|9|"
com = com & vbnewline & "dgFileUpload.dgUpload|http://www.devguru.com|DevGuru - dgFileup|3|"
com = com & vbnewline & "dgReport.Report|http://www.devguru.com|DevGuru - dgReport|0|"
com = com & vbnewline & "dgSort.QuickSort|http://www.devguru.com|DevGuru - dgSort|0|"
com = com & vbnewline & "dgTree.Tree|http://www.devguru.com|DevGuru - dgTree|0|"
com = com & vbnewline & "Dundas.Mailer|http://www.dundas.com|Dundas - ASPMailer|1|"
com = com & vbnewline & "Dundas.PieChartServer.2|http://www.dundas.com|Dundas - Pie Chart Server Control|7|"
com = com & vbnewline & "Dundas.Upload|http://www.dundas.com|Dundas - Upload|3|"
com = com & vbnewline & "EasyMail.SMTP.5|http://www.quiksoft.com|Quicksoft - EasyMail (free)|1|"
com = com & vbnewline & "AspPing.Conn|http://www.15seconds.com/component/pg000229.htm|ASP Ping|0|"
com = com & vbnewline & "Dynu.CreditCard|http://www.dynu.com|Dynu CreditCard|10|11"
com = com & vbnewline & "Dynu.DateTime|http://www.dynu.com|Dynu DateTime|0|"
com = com & vbnewline & "Dynu.DNS|http://www.dynu.com|Dynu DNS|0|"
com = com & vbnewline & "Dynu.Exec|http://www.dynu.com|Dynu Exec|0|"
com = com & vbnewline & "Dynu.Email|http://www.dynu.com|Dynu Email|1|"
com = com & vbnewline & "Dynu.Encrypt|http://www.dynu.com|Dynu Encrypt|9|"
com = com & vbnewline & "Dynu.FileUtil|http://www.dynu.com|Dynu File|6|"
com = com & vbnewline & "Dynu.FTP|http://www.dynu.com|Dynu FTP|0|6"
com = com & vbnewline & "Dynu.HTTP|http://www.dynu.com|Dynu HTTP|0|"
com = com & vbnewline & "Dynu.POP3|http://www.dynu.com|Dynu POP3|1|"
com = com & vbnewline & "Dynu.Ping|http://www.dynu.com|Dynu Ping|0|"
com = com & vbnewline & "Dynu.TCPSocket|http://www.dynu.com|Dynu TCPSocket|0|"
com = com & vbnewline & "Dynu.StringUtil|http://www.dynu.com|Dynu String|0|"
com = com & vbnewline & "Dynu.Upload|http://www.dynu.com|Dynu Upload|3|"
com = com & vbnewline & "Dynu.Wait|http://www.dynu.com|Dynu Wait|0|"
com = com & vbnewline & "Dynu.Whois|http://www.dynu.com|Dynu Whois|0|"
com = com & vbnewline & "MP_Mikys_ASP.Password|http://www.mikys-asp.nykoping.net/Password|ASP Password|9|"
com = com & vbnewline & "S3Weather.Current|http://www.softshell.net|S3 Weather Component (free)|0|"
com = com & vbnewline & "AuthNetSSLConnect.SSLPost|http://www.authorize.net|Authorize.Net Transaction COM (free)|10|11"
com = com & vbnewline & "HexValidEmail.Connection|http://www.hexillion.com|Hexillion - HexValidEmail|1|11"
com = com & vbnewline & "Hexillion.HexIcmp|http://www.hexillion.com|Hexillion - HexIcmp|0|"
com = com & vbnewline & "Hexillion.HexLookup|http://www.hexillion.com|Hexillion - HexLookup|0|"
com = com & vbnewline & "Hexillion.HexTcpQuery|http://www.hexillion.com|Hexillion - HexTcpQuery|0|"
com = com & vbnewline & "HexDns.Connection|http://www.hexillion.com|Hexillion - HexDSN|0|"
com = com & vbnewline & "ocxQmail.ocxQmailCtrl.1|http://www.flicks.com|Flicks - ocxQmail|1|"
com = com & vbnewline & "OCXHTTP.OCXHttpCtrl.1|http://www.flicks.com|Flicks - OCXHttp|0|"
com = com & vbnewline & "ocxQmail.ocxQmailCtrl.1|http://www.flicks.com|Flicks - OCXQMail|1|"
com = com & vbnewline & "VASPTV.ASPTreeView|http://www.visualasp.com|VisualASP - TreeView|0|"
com = com & vbnewline & "VASPLV.ASPListView|http://www.visualasp.com|VisualASP - ListView|0|"
com = com & vbnewline & "VASPMV.ASPMonthView|http://www.visualasp.com|VisualASP - MonthView|0|"
com = com & vbnewline & "VASPTB.ASPTabView|http://www.visualasp.com|VisualASP - TabView|0|"
com = com & vbnewline & "ASPWordToy.WordToy|http://www.asptoys.com|ASP Toys - WordToy (Word Converter)|6|"
com = com & vbnewline & "ASPTabToy.TabToy|http://www.asptoys.com|ASP Toys - TabToy|0|"
com = com & vbnewline & "aspZipCodeToy.ZipCodeToy|http://www.asptoys.com|ASP Toys - ASP ZipCodeToy|0|11"
com = com & vbnewline & "ASPCryptToy.CryptToy|http://www.asptoys.com|ASP Toys - CryptToy|9|"
com = com & vbnewline & "Convert.t2h|http://members.home.net/pjsteele/asp|CONVERT - string/html/text manipulation (free)|0|"
com = com & vbnewline & "APDocConv.Object|http://www.activepdf.com|activePDF - DocConverter|5|"
com = com & vbnewline & "APWebGrabber.Object|http://www.activepdf.com|activePDF - WebGrabber|5|"
com = com & vbnewline & "APServer.Object|http://www.activepdf.com|activePDF - activePDF Server|5|"
com = com & vbnewline & "APSpool.Object|http://www.activepdf.com|activePDF - Spooler|5|"
com = com & vbnewline & "APToolkit.Object|http://www.activepdf.com|activePDF - Toolkit|5|"
com = com & vbnewline & "shotgraph.image|http://www.shotgraph.com|Shot Graph|7|"
com = com & vbnewline & "IntrChart.Chart|http://www.compsysaus.com.au|IntrChart|7|"
com = com & vbnewline & "IntrSQL.Query|http://www.compsysaus.com.au|IntrSQL|0|"
com = com & vbnewline & "IntrPWD.Validate|http://www.compsysaus.com.au|IntrPWD|9|"
com = com & vbnewline & "IntrCard.Credit|http://www.compsysaus.com.au|IntrCard|0|11"
com = com & vbnewline & "AspSmartImage.SmartImage|http://www.aspsmart.com|ASP Smart - aspSmartImage|4|"
com = com & vbnewline & "AspSmartChat.SmartChat|http://www.aspsmart.com|ASP Smart - aspSmartChat|0|"
com = com & vbnewline & "AspSmartFile.SmartFile|http://www.aspsmart.com|ASP Smart - aspSmartFile|6|"
com = com & vbnewline & "aspSmartMenu.SmartMenuPopUp|http://www.aspsmart.com|ASP Smart - aspSmartMenu|0|"
com = com & vbnewline & "AspSmartDate.SmartDate|http://www.aspsmart.com|ASP Smart - aspSmartDate|0|"
com = com & vbnewline & "AspSmartUpload.SmartUpload|http://www.aspsmart.com|ASP Smart - aspSmartUpload|3|"
com = com & vbnewline & "aspSmartMail.SmartMail|http://www.aspsmart.com|ASP Smart - aspSmartMail|1|"
com = com & vbnewline & "aspSmartCache.SmartCache|http://www.aspsmart.com|ASP Smart - aspSmartCache|0|"
com = com & vbnewline & "xAuthorize.Charge|http://www.xauthorize.com|xAuthorize CC|10|11"
com = com & vbnewline & "acDesktop.Desktop|http://www.activecomponents.nu|acDesktop|0|"
com = com & vbnewline & "acNetwork.DNS|http://www.activecomponents.nu|acNetwork|0|"
com = com & vbnewline & "acSMTP.Smtp|http://www.activecomponents.nu|acSMTP SSL|9|"
com = com & vbnewline & "Temperature.Conversion|http://asp.myscripting.com/activextemp.asp|Temperature Conversion|0|"
com = com & vbnewline & "cyScape.browserObj|http://www.cyscape.com|BrowserHawk|2|11"
com = com & vbnewline & "dkQmail.Qmail||dkQMail|1|"
com = com & vbnewline & "Geocel.Mailer|http://www.geocel.com|GeoCel|1|"
com = com & vbnewline & "iismail.iismail.1||IISMail|1|"
com = com & vbnewline & "SmtpMail.SmtpMail.1||SMTP|1|"
com = com & vbnewline & "OpenX2.Connection|http://www.openx.ca|OpenX|1|"
com = com & vbnewline & "ABMailer.Mailman|http://www.absoftwarex.com/abmailer|ABMailer|1|"
com = com & vbnewline & "c2geread.Message|http://www.componentstogo.com|C2GEread|1|"
com = com & vbnewline & "C2G.SCM|http://www.componentstogo.com|C2GSCM|0|8"
com = com & vbnewline & "C2GSCM.Service|http://www.componentstogo.com|C2GSCM|8|0"
com = com & vbnewline & "C2G.SCAN|http://www.componentstogo.com|C2GSCAN|0|"
com = com & vbnewline & "C2G.whois|http://www.componentstogo.com|C2GWHOIS |0|"
com = com & vbnewline & "c2g.http|http://www.componentstogo.com|C2GHttp |0|"
com = com & vbnewline & "C2G.Ping|http://www.componentstogo.com|C2GPing|0|"
com = com & vbnewline & "C2G.Tracert|http://www.componentstogo.com|C2GTracert|0|"
com = com & vbnewline & "ANUPLOAD.OBJ|http://www.adminsystem.net/webapp/popcom|ANPOP|1|"
com = com & vbnewline & "ASPXP.Mail|http://aspxp.com/free_stuff/aspxpmail|ASPXPMail (free)|1|"
com = com & vbnewline & "ActiveMessenger.Message|http://www.infomentum.com|ActiveMessenger|1|"
com = com & vbnewline & "ActiveFile.Post|http://www.infomentum.com|ActiveFile|3|"
com = com & vbnewline & "ActiveNavigator.Toolbar|http://www.infomentum.com|ActiveNavigator|0|"
com = com & vbnewline & "ActiveProfile.Profile|http://www.infomentum.com|ActiveProfile|2|9"
com = com & vbnewline & "DartZip.Zip.1|http://www.dart.com|Dart Zip Compression Tool|6|"
com = com & vbnewline & "Dart.Ftp.1|http://www.dart.com|Dart FTP Tool|6|0"
com = com & vbnewline & "Dart.Pop.1|http://www.dart.com|Dart POP Mail|1|"
com = com & vbnewline & "Dart.Ping.1|http://www.dart.com|Dart Ping|0|"
com = com & vbnewline & "Dart.Dns.1|http://www.dart.com|Dart DNS|0|"
com = com & vbnewline & "Dart.Smtp.1|http://www.dart.com|Dart SMTP|1|"
com = com & vbnewline & "Dart.Telnet.1|http://www.dart.com|Dart PowerTCP Telnet Tool|0|"
com = com & vbnewline & "Dart.Http.1|http://www.dart.com|Dart HTTP|0|"
com = com & vbnewline & "Dart.Tcp.1|http://www.dart.com|Dart TCP|0|"
com = com & vbnewline & "Dart.WebPage.1|http://www.dart.com|Dart WebPage|0|"
com = com & vbnewline & "Dart.WebASP.1|http://www.dart.com|Dart ASP|0|"
com = com & vbnewline & "Dart.Message.1|http://www.dart.com|Dart Message|0|"
com = com & vbnewline & "Dart.Manager.1|http://www.dart.com|Dart Manager|0|"
com = com & vbnewline & "quicktab.quicktabs|http://www.webintel.net|Quicktab|0|"
com = com & vbnewline & "waspzip.waspzip|http://www.webintel.net|Wasp Zip|6|5"
com = com & vbnewline & "easyBarCode.aspBarCode|http://www.mitdata.com|aspEasyBarCode|7|0"
com = com & vbnewline & "aspZip.EasyZIP|http://www.mitdata.com|aspEasyZIP|6|5"
com = com & vbnewline & "aspPDF.EasyPDF|http://www.mitdata.com|aspEasyPDF|5|6"
com = com & vbnewline & "aspCrypt.EasyCRYPT|http://www.mitdata.com|aspEasyCRYPT|9|"
com = com & vbnewline & "objBarGraph.DrawChart|http://www.livesoup.com/bargraph.asp|BarGraph (free)|7|"
com = com & vbnewline & "LyfUpload.UploadFile|http://www.21jsp.com|LyfUpload (free)|3|"
com = com & vbnewline & "lyfimage.image|http://www.21jsp.com|LyfImage (free)|4|7"
com = com & vbnewline & "ASPControlHost.Host|http://release-systems.8m.com/asphost.html|ASPControlHost|7|4"
com = com & vbnewline & "GSServer.GSServerProp|http://www.graphicsserver.com|Graphics Server|4|7"
com = com & vbnewline & "ASPPicture.Picture|http://www.unchanged.net|ASPPicture|4|"
com = com & vbnewline & "COMobjectsNET.IconGrabber|http://www.comobjects.net|COMobjects.NET Icon Grabber|4|"
com = com & vbnewline & "COMobjects.NET.PictureProcessor|http://www.comobjects.net|COMobjects.NET Picture Processor|4|"
com = com & vbnewline & "COMobjectsNET.PictureGalleryPro|http://www.comobjects.net|COMobjects.NET Picture Gallery Pro|4|"
com = com & vbnewline & "COMobjectsNET.Colorizer|http://www.comobjects.net|COMobjects.NET Colorizer|4|"
com = com & vbnewline & "COMobjectsNET.PieChart|http://www.comobjects.net|COMobjects.NET 3D Pie Chart|7|4"
com = com & vbnewline & "ChartDirector.API|http://www.advsofteng.com|ChartDirector|7|"
com = com & vbnewline & "Stonebroom.ASPointer|http://www.stonebroom.com|Stonebroom.ASPointer|13|5"
com = com & vbnewline & "Stonebroom.ASP2XML|http://www.stonebroom.com|Stonebroom.ASP2XML|13|5"
com = com & vbnewline & "Stonebroom.RegEx|http://www.stonebroom.com|Stonebroom.RegEx|0|"
com = com & vbnewline & "Stonebroom.RemoteZip|http://www.stonebroom.com|Stonebroom.RemoteZip|5|6"
com = com & vbnewline & "Stonebroom.SaveForm|http://www.stonebroom.com|Stonebroom.SaveForm|12|"
com = com & vbnewline & "Stonebroom.ServerZip|http://www.stonebroom.com|Stonebroom.ServerZip|5|6"
com = com & vbnewline & "Stonebroom.XSLTransform|http://www.stonebroom.com|Stonebroom.XSLTransform|13|5"
com = com & vbnewline & "OpenX.DBMail|http://www.openx.ca|OpenX DBMail|1|12"
com = com & vbnewline & "com.comsoltech.CGI|http://www.comsoltech.com|com.comsoltech.CGI (free)|12|"
com = com & vbnewline & "Datafun.FormBoy|http://www.datafun.net|FormBoy|12|10"
com = com & vbnewline & "AddressTools.ZIPCheck|http://www.addresstools.com|AddressTools - ZIPCheck|11|12"
com = com & vbnewline & "AddressTools.EmailCheck|http://www.addresstools.com|AddressTools - EmailCheck|11|12"
com = com & vbnewline & "VisualSoft.Mail.1|http://www.visualmart.com|VisualSoft Mail|1|"
com = com & vbnewline & "VisualSoft.BLOWFISHCrypt.1|http://www.visualmart.com|VisualSoft Crypt|9|"
com = com & vbnewline & "VisualSoft.FTP.1|http://www.visualmart.com|VisualSoft FTP|6|0"
com = com & vbnewline & "VisualSoft.HTTP.1|http://www.visualmart.com|VisualSoft HTTP|2|0"
com = com & vbnewline & "VisualSoft.Chart.1|http://www.visualmart.com|VisualSoft Chart|7|"
com = com & vbnewline & "VisualSoft.DMXML.1|http://www.visualmart.com|VisualSoft XMLPro|13|"
com = com & vbnewline & "VisualSoft.DataAdmin.1|http://www.visualmart.com|VisualSoft DataAdmin|0|"
com = com & vbnewline & "QwerkSoft.FormSlam|http://www.qwerksoft.com|Form Slam|12|11"
com = com & vbnewline & "SiteAdmin.AdminTools|http://components.sitetown.com|SiteSecurity|9|"
com = com & vbnewline & "SiteSecurity.Login|http://components.sitetown.com|SiteSecurity|9|"
com = com & vbnewline & "FileDownload.Manager|http://components.sitetown.com|File Download|6|0"
com = com & vbnewline & "EasyDb.Database|http://components.sitetown.com|Easy DB|0|"
com = com & vbnewline & "AbsoluteHttp.Conn|http://www.speeq.com|AbsoluteHTTP|0|"
com = com & vbnewline & "ASPCharge.CC|http://www.bluesquirrel.com|A$PCharge|10|11"
com = com & vbnewline & "ProjectDisplay.Charts|http://www.aspkey.com|ASPkey ProjectDisplay|0|"
com = com & vbnewline & "IPWorksASP.SOAP|www.nsoftware.com|IP Works Soap|13|"
com = com & vbnewline & "IPWorksASP.FileMailer|www.nsoftware.com|IP Works FileMailer|1|6"
com = com & vbnewline & "IPWorksASP.FTP|www.nsoftware.com|IP Works FTP|0|"
com = com & vbnewline & "IPWorksASP.HTMLMailer|www.nsoftware.com|IP Works HTMLMailer|1|"
com = com & vbnewline & "IPWorksASP.HTTP|www.nsoftware.com|IP Works HTTP|13|0"
com = com & vbnewline & "IPWorksASP.ICMPPort|www.nsoftware.com|IP Works ICMPPort|0|"
com = com & vbnewline & "IPWorksASP.IMAP|www.nsoftware.com|IP Works IMAP|0|"
com = com & vbnewline & "IPWorksASP.IPInfo|www.nsoftware.com|IP Works IPInfo|0|"
com = com & vbnewline & "IPWorksASP.IPPort|www.nsoftware.com|IP Works IPPort|0|"
com = com & vbnewline & "IPWorksASP.LDAP|www.nsoftware.com|IP Works LDAP|0|"
com = com & vbnewline & "IPWorksASP.MCast|www.nsoftware.com|IP Works MCast|0|"
com = com & vbnewline & "IPWorksASP.MIME|www.nsoftware.com|IP Works MIME|1|"
com = com & vbnewline & "IPWorksASP.MX|www.nsoftware.com|IP Works MX|1|"
com = com & vbnewline & "IPWorksASP.NetClock|www.nsoftware.com|IP Works NetClock|0|"
com = com & vbnewline & "IPWorksASP.NetCode|www.nsoftware.com|IP Works NetCode|0|"
com = com & vbnewline & "IPWorksASP.NetDial|www.nsoftware.com|IP Works NetDial|0|"
com = com & vbnewline & "IPWorksASP.NNTP|www.nsoftware.com|IP Works NNTP|0|"
com = com & vbnewline & "IPWorksASP.Ping|www.nsoftware.com|IP Works Ping|0|"
com = com & vbnewline & "IPWorksASP.POP|www.nsoftware.com|IP Works POP|1|"
com = com & vbnewline & "IPWorksASP.RCP|www.nsoftware.com|IP Works RCP|6|0"
com = com & vbnewline & "IPWorksASP.Rexec|www.nsoftware.com|IP Works Rexec|0|"
com = com & vbnewline & "IPWorksASP.Rshell|www.nsoftware.com|IP Works Rshell|0|"
com = com & vbnewline & "IPWorksASP.SMTP|www.nsoftware.com|IP Works SMTP|1|"
com = com & vbnewline & "IPWorksASP.SNMP|www.nsoftware.com|IP Works SNMP|1|0"
com = com & vbnewline & "IPWorksASP.SNPP|www.nsoftware.com|IP Works SNPP|13|0"
com = com & vbnewline & "IPWorksASP.Telnet|www.nsoftware.com|IP Works Telnet|0|"
com = com & vbnewline & "IPWorksASP.TFTP|www.nsoftware.com|IP Works TFTP|0|"
com = com & vbnewline & "IPWorksASP.TraceRoute|www.nsoftware.com|IP Works TraceRoute|0|"
com = com & vbnewline & "IPWorksASP.UDPPort|www.nsoftware.com|IP Works UDPPort|0|"
com = com & vbnewline & "IPWorksASP.WebForm|www.nsoftware.com|IP Works WebForm|12|"
com = com & vbnewline & "IPWorksASP.WebUpload|www.nsoftware.com|IP Works WebUpload|3|"
com = com & vbnewline & "IPWorksASP.Whois|www.nsoftware.com|IP Works Whois|0|"
com = com & vbnewline & "IPWorksASP.XMLp|www.nsoftware.com|IP Works XMLp|13|"
com = com & vbnewline & "iisCC.cc|http://www.iiscart.com|IIS Cart - iisCARTcc|0|11"
com = com & vbnewline & "Coalesys.CSPanelBar.2|http://www.coalesys.com|CSPanelBar|0|"
com = com & vbnewline & "Coalesys.CSWebMenu.1|http://www.coalesys.com|CSWebMenu|0|3"
com = com & vbnewline & "TCPIP.DNS|http://www.pstruh.cz/help/tcpip/library.htm|Simple DNS+Traceroute|0|"
com = com & vbnewline & "DrWFM.fm|http://www.dataroad.sk/dr/drwfm/default.asp|DrWebFileManager|6|"
com = com & vbnewline & "id3.id3get|http://www.infinitemonkeys.ws/infinitemonkeys|Atrax ID3.ID3Get|0|"
com = com & vbnewline & "Atrax.ComboBox|http://www.infinitemonkeys.ws/infinitemonkeys|Atrax ComboBox|0|"
com = com & vbnewline & "Atrax.URLGrabber|http://www.infinitemonkeys.ws/infinitemonkeys|Atrax URLGrabber|0|13"
com = com & vbnewline & "Atrax.Whois|http://www.infinitemonkeys.ws|Atrax Whois|0|"
com = com & vbnewline & "SOFTWING.ASPEventlog|http://www.alphasierrapapa.com|Asp Event log (FREE)|8|0"
com = com & vbnewline & "Softwing.EventLogReader|http://www.alphasierrapapa.com|Event Log Reader (FREE)|0|"
com = com & vbnewline & "Softwing.AspQPerfCounters|http://www.alphasierrapapa.com|AspQPerfCounters|8|0"
com = com & vbnewline & "SOFTWING.AspTear|http://www.alphasierrapapa.com|AspTear|8|0"
com = com & vbnewline & "AspTouch.TouchIt|http://www.alphasierrapapa.com|AspTouch TouchIt (FREE)|8|0"
com = com & vbnewline & "Softwing.FileCache.1|http://www.alphasierrapapa.com|Softwing FileCache (FREE)|8|0"
com = com & vbnewline & "Softwing.LocaleFormatter|http://www.alphasierrapapa.com|LocaleFormatter (FREE)|0|"
com = com & vbnewline & "Softwing.MacBinary|http://www.alphasierrapapa.com|MacBinary Xtraction (FREE)|6|"
com = com & vbnewline & "Softwing.OdbcRegTool|http://www.alphasierrapapa.com|OdbcRegTool (FREE)|8|0"
com = com & vbnewline & "Softwing.Profiler|http://www.alphasierrapapa.com|Softwing ASP Script Speed Profiler (FREE)|0|"
com = com & vbnewline & "AlphaSierraPapa.AspRegSvr|http://www.alphasierrapapa.com|RegServer [component registration via ASP] (FREE!!)|8|0"
com = com & vbnewline & "Softwing.VersionInfo|http://www.alphasierrapapa.com|VersionInfo|8|0"
com = com & vbnewline & "w3info.w3info.1|http://www.alphasierrapapa.com|W3 Info|0|"
com = com & vbnewline & "SoftwingXSB.ShoppingBag|http://www.alphasierrapapa.com|Softwing ShoppingBag|10|"
com = com & vbnewline & "crossoft.quickcal|http://www.quickgallery.com|Quick Calendar|0|"
com = com & vbnewline & "crossoft.wapsplash|http://www.quickgallery.com|QuickDeck|0|"
com = com & vbnewline & "crossoft.waplist|http://www.quickgallery.com|QuickDeck|0|"
com = com & vbnewline & "crossoft.remotescript|http://www.quickgallery.com|QuickList|0|"
com = com & vbnewline & "crossoft.quicklist|http://www.quickgallery.com|QuickList|0|"
com = com & vbnewline & "crossoft.quicktable|http://www.quickgallery.com|QuickTable|0|"
com = com & vbnewline & "OneTouchASP.StrFunctions|http://www.1touchasp.com|1Touch|0|"
com = com & vbnewline & "ZmeYsoft.Hashes.MD5|http://www.newobjects.com|ZmeYsoft MD5 Hash|9|0"
com = com & vbnewline & "binarysendfile.BinFileSend|http://www.newobjects.com|Binarysendfile component|0|"
com = com & vbnewline & "werkslib.mp3exp|http://www.marban.at/download/aspmp3.zip|werk3AT - MP3|0|"
com = com & vbnewline & "TreeGen.Tree|http://www.treegen.com|Tree Gen|0|"
com = com & vbnewline & "Text2Tree150d.tree|http://www.asp-components.de|Text2Tree|0|"
com = com & vbnewline & "ASPBarChart100d.chart|http://www.asp-components.de|Bar Chart|8|0"
com = com & vbnewline & "AspWebCal120d.webcal|http://www.asp-components.de|ASP WebCalendar|0|"
com = com & vbnewline & "ScriptUtils.ASPForm|http://pstruh.cz/help/ScptUtl/library.htm|Simple Upload|3|0"
com = com & vbnewline & "ScriptUtils.ByteArray|http://pstruh.cz/help/ScptUtl/library.htm|Simple Download|0|"
com = com & vbnewline & "ScriptUtils.Kernel|http://pstruh.cz/help/ScptUtl/library.htm|ASP Timing|0|"
com = com & vbnewline & "Scribe.ScribeDOM|http://www.innuvo.com|ScribeDOM|13|"
com = com & vbnewline & "ANPOP.POPMSG|http://www.adminsystem.net|ANPOP |1|"
com = com & vbnewline & "ANSMTP.OBJ|http://www.adminsystem.net|ANSMTP|1|"
com = com & vbnewline & "ANUPLOAD.OBJ|http://www.adminsystem.net|ANUPLOAD (free)|3|"
com = com & vbnewline & "VoiceShot.VoiceShot|http://www.voiceshot.com/api/readme.htm|ASP Call|0|"
com = com & vbnewline & "SimplePageASP.SNPP|http://www.rushweb.com|SimplePageASP SNPP|0|"
com = com & vbnewline & "khttp.inet|http://www.rainfall.com|KHTTP|13|0"
com = com & vbnewline & "OCXHTTP.OCXHttpCtrl.1|http://www.flicks.com|Flicks OCXHttp|13|0"
com = com & vbnewline & "URLFetch.URLFetch|http://www.screen-scraper.com|URLFetch|13|0"
com = com & vbnewline & "Dundas.Mailer|http://www.dundas.com|Dundas Mailer|1|"
com = com & vbnewline & "Dundas.Mailer.1|http://www.dundas.com|Dundas Mailer|1|"
com = com & vbnewline & "Dundas.PieChartServer.1|http://www.dundas.com|Dundas PieChartServer|7|"
com = com & vbnewline & "Dundas.Upload|http://www.dundas.com|Dundas Upload|3|"
com = com & vbnewline & "Dundas.Upload.2|http://www.dundas.com|Dundas Upload|3|"
com = com & vbnewline & "Dundas.ChartServer|http://www.dundas.com|Dundas ChartServer|7|"
com = com & vbnewline & "Dundas.ChartServer2D.1|http://www.dundas.com|Dundas ChartServer 2D|7|"
com = com & vbnewline & "ABCUpload4.XForm|http://www.websupergoo.com|ABC Upload|3|"
com = com & vbnewline & "ABCpdf3.Doc|http://www.websupergoo.com|ABC PDF|0|"
com = com & vbnewline & "ImageGlue5.Canvas|http://www.websupergoo.com|Image Glue|4|"
com = com & vbnewline & "ImageEffects.FX|http://www.websupergoo.com|Image Effects|4|"
com = com & vbnewline & "ABCDrawHTML.Page|http://www.websupergoo.com|ABC Draw HTML|4|0"
com = com & vbnewline & "ABCCrypto2.Crypto|http://www.websupergoo.com|ABC Crypto|9|"
com = com & vbnewline & "MetaFiler2.File|http://www.websupergoo.com|MetaFiler|4|"
com = com & vbnewline & "XceedSoftware.XceedZip|http://www.xceedsoft.com|XceedZip|5|"
com = com & vbnewline & "Xceed.BinaryEncoding|http://www.xceedsoft.com|Xceed Binary Encoding|0|"
com = com & vbnewline & "Xceed.Base64Encoding|http://www.xceedsoft.com|Xceed Base 64 Encoding|0|"
com = com & vbnewline & "Xceed.Encryption|http://www.xceedsoft.com|Xceed Encryption|9|0"
com = com & vbnewline & "Xceed.TwofishEncryptionMethod|http://www.xceedsoft.com|Xceed Two fish Encryption Method|9|0"
com = com & vbnewline & "Xceed.HavalHashingMethod|http://www.xceedsoft.com|Xceed Haval Hashing Method|9|0"
com = com & vbnewline & "XceedSoftware.XceedFtp|http://www.xceedsoft.com|Xceed Ftp|8|0"
com = com & vbnewline & "Xceed.StreamingCompression|http://www.xceedsoft.com|Xceed Streaming Compression|0|"
com = com & vbnewline & "Xceed.DeflateCompression|http://www.xceedsoft.com|Xceed Deflate Compression|0|"
com = com & vbnewline & "CDO.Message|http://www.microsoft.com|CDO Message|1|"
'com = com & vbnewline & "||||"
'com = com & vbnewline & "||||"
'com = com & vbnewline & "||||"
'com = com & vbnewline & "||||"
'com = com & vbnewline & "||||"

com = Split(com, vbnewline)


cat = "Miscellaneous"				' 0
cat = cat & "|Email"				' 1
cat = cat & "|Browser"				' 2
cat = cat & "|Upload"				' 3
cat = cat & "|Image"				' 4
cat = cat & "|Documents"			' 5
cat = cat & "|File Management"		' 6
cat = cat & "|Graphs & Charts"		' 7
cat = cat & "|Server Management"	' 8
cat = cat & "|Users & Security"		' 9
cat = cat & "|E-Commerce"			' 10
cat = cat & "|Validation"			' 11
cat = cat & "|Forms"				' 12
cat = cat & "|XML"					' 13


cat = Split(cat, "|")

	if (isnumeric(request("show"))) then show = CInt(request("show")) else show = 1
		if (show > 3) then show = 1
	if (isnumeric(request("showCat")) AND request("showCat") <> "") then showCat = CInt(request("showCat")) else showCat = "all"
		if isNumeric(showCat) then
			if (showCat > UBound(cat)) then showCat = "all"
		end if
checkVersion = getHTML("http://www.pensaworks.com/tutorials/com_version.asp")
if (checkVersion <> lastUpdate) then newVersion = True
%>
<HTML>
<HEAD>
<TITLE>ASP Component Test</TITLE>
<SCRIPT language=JavaScript>
<!--
function BringUpWindow(webpage) {
     var url = webpage;
     var hWnd = window.open(url,"Mailer_Popup","width=425,height=325,resizable=yes,scrollbars=yes,status=yes");
	 if (window.focus) {hWnd.focus()}
     if (hWnd != null) {
	      if (hWnd.opener == null) {
		   hWnd.opener = self; window.name = "home";
		   hWnd.location.href=url;
		   }
	 } else {
    }
   }
// -->
</SCRIPT>
</HEAD>
<body bgcolor="#ffffff" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginwidth="0" marginheight="0">
<%
if request("comID") <> "" then
comID = request("comID")
comDetails = Split(com(comID), "|")
comCreate = comDetails(0)
comURL = comDetails(1)
comName = comDetails(2)
comCat = comDetails(3)
comCat2 = comDetails(4)
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <tr>
    <td bgcolor="#000080"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Component Details</font></b></font></td>
  </tr>
</table>
<%
Set b = New ProgIDInfo
Set a = b.LoadProgID(comCreate)
if a.Description <> "" then
%>
			<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
			<b>Component:</b> <%=comName%><br>
			<b>Website:</b> <% if comURL <> "" then %><a href="<%=comURL%>" target="_blank"><%=comURL%></a><% end if %><br>
			<b>Category(s):</b> <%=cat(comCat)%><% if comCat2 <> "" then %> | <%=cat(comCat2)%><% end if %><br>
			<b>Description:</b> <%= a.Description %><br>
		    <b>DLLName:</b> <%=a.DLLName%><br>
			<b>ProgID:</b> <%=a.ProgID%><br>
			<b>ClsID:</b> <%=a.ClsID%><br>
			<b>Path:</b> <%=a.Path%><br>
			<b>TypeLib:</b> <%=a.TypeLib%><br>
			<b>Version:</b> <%=a.Version%><br>
			</font>
      <% else %>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>No Information could be found for:</b> <%=comName%></font>
      <% end if %>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p align="center"><a href=# onClick="self.close();"><b>Close Window</b></a></p>
</font>
<% else %>
<table border="0" width="100%" cellspacing="0" cellpadding="3">
  <tr>
    <td width="100%" bgcolor="#676767">
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="4" color="mintcream">ASP Component Test</font></b></div>
    </td>
  </tr>
</table>

<% response.flush() %>
<table border="0" align="center" cellspacing="2" cellpadding="4">
  <tr bgcolor="#000080">
    <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#676767">#</font></b></td>
    <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#676767">Category</font></b></td>
    <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#676767">Status</font></b></td>
    <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#676767">Details</font></b></td>
    <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#676767">Com</font></b></td>
  </tr>
<%
	for i = 0 to UBound(com)
		comDetails = split(com(i), "|")
		display = false
		display2 = false
		comCreate = comDetails(0)
		comURL = comDetails(1)
		comName = comDetails(2)
		comCat = CInt(comDetails(3))
		comCat2 = CInt(comDetails(4))
		installed = IsObjInstalled(comCreate)

		if (NOT Installed) then display = false else display = true

		if isnumeric(showCat) then
			if (comCat = showCat or comCat2 = showCat) then display2 = true else display2 = false
		else
			display2 = true
		end if

	    if (display AND display2) then
	        onNum = onNum + 1
                if (onNum Mod 2) Then 
%>
  <tr>
<% else %>
  <tr bgcolor="#CCCCCC">
<% end If %>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=(onNum)%></b></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><% if IsNumeric(showCat) then %><%=cat(showCat)%><% else %><%=cat(comCat)%><% end if %></font></td>
    <td>
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>
	    <% if NOT installed then %>
		<font color="#FF0000">Not Installed</font>
	    <%
		else
		installedCOMs = installedComs + 1
		%>
		<font color="#009933">Installed</font>
	    <% end if %>
	</b></font></div>
    </td>
    <td>
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
        <% if NOT installed then %>
      None Available
        <% else %>
      <a href="Javascript:BringUpWindow('<%=Mid(request.servervariables("SCRIPT_NAME"), InstrRev(request.servervariables("SCRIPT_NAME"), "/") + 1)%>?comID=<%=i%>')">COM Details</a>
        <% end if %>
      </font></div>
    </td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><% if comURL <> "" then %><a href="<%=comURL%>" target="_blank"><%=comName%></a><% else %><%=comName%><% end if %></font></td>
  </tr>
  <%
end if
installed = "" : comCreate = "" : comURL = "" : comName = "" : comCat = "" : comCat2 = ""
  next
  response.flush()
  %>
<% if onNum = 0 then %>
<tr>
    <td colspan="5">
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>You do not have any components installed for your selections.</b></font></div>
    </td>
  </tr>
<% end if %>
</table>
	<div align="center">
  <p><font size="3" face="Verdana, Arial, Helvetica, sans-serif">You have a total of <b><%=installedCOMs%></b> COMs installed out of <b>327</b> checked.</font></p>
</div>
<% end if %>
</BODY>
</HTML>
<%
function IsObjInstalled(strClassString)
 IsObjInstalled = false : Err = 0
	 Set testObj = Server.CreateObject(strClassString)
		 if (0 = Err) then IsObjInstalled = true else IsObjInstalled = false
	 Set testObj = nothing
end function

Class Program
	Public Description, ClsID, ProgID, Path, TypeLib, Version, DLLName
End Class

Class ProgIDInfo
	Private WshShell, sCVProgID, oFSO

	Private Sub Class_Initialize()
		On Error Resume Next
		set oFSO = CreateObject("Scripting.FileSystemObject")
		Set WshShell = CreateObject("WScript.Shell")
	End Sub

	Private Sub Class_Terminate()
		If IsObject(WshShell) Then Set WshShell = Nothing
		If IsObject(oFSO) Then set oFSO = Nothing
	End Sub

	Private Function IIf(byval conditions, byval trueval, byval falseval)
		if cbool(conditions) then IIf = trueval else IIf = falseval
	End Function

	Public Function LoadProgID(ByVal sProgramID)
		Dim sTmpProg, oTmp, sRegBase, sDesc, sClsID
		Dim sPath, sTypeLib, sProgID, sVers, sPathSpec
		If IsObject(WshShell) Then
			On Error Resume Next
			sCVProgID = WshShell.RegRead("HKCR\" & _
				sProgramID & "\CurVer\")
			sTmpProg = IIf(Err.Number = 0, sCVProgID, sProgramID)

			sRegBase = "HKCR\" & sTmpProg
			sDesc = WshShell.RegRead(sRegBase & "\")
			sClsID = WshShell.RegRead(sRegBase & "\clsid\")
			sRegBase = "HKCR\CLSID\" & sClsID
			sPath = WshShell.RegRead(sRegBase & "\InprocServer32\")
			sPath = WshShell.ExpandEnvironmentStrings(sPath)
			sTypeLib = WshShell.RegRead(sRegBase & "\TypeLib\")
			sProgID = WshShell.RegRead(sRegBase & "\ProgID\")
			sVers = oFSO.getFileVersion(sPath)
			sPathSpec = right(sPath, len(sPath) - _
				instrrev(sPath, "\"))

			Set oTmp = New Program
			oTmp.Description = sDesc
			oTmp.ClsID = IIf(sClsID <> "", sClsID, "undetermined")
			oTmp.Path = IIf(sPath <> "", sPath, "undetermined")
			oTmp.TypeLib = IIf(sTypeLib <> "", _
				sTypeLib, "undetermined")
			oTmp.ProgID = IIf(sProgID <> "", _
				sProgID, "undetermined")
			oTmp.DLLName = IIf(sPathSpec <> "", _
				sPathSpec, "undetermined")
			oTmp.Version = IIf(sVers <> "", sVers, "undetermined")
			Set LoadProgID = oTmp
		Else
			Set LoadProgID = Nothing
		End If
	End Function
End Class

function getHTML(strURL)
  dim objXMLHTTP, strReturn
  Set objXMLHTTP = SErver.CreateObject("Microsoft.XMLHTTP")
  objXMLHTTP.Open "GET", strURL, False
  objXMLHTTP.Send
  getHTML = objXMLHTTP.responseText
  Set objXMLHTTP = Nothing
end function
%>
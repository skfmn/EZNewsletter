<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <div class="row 50%">
        <div class="-1u 10u$ 12u$(medium)">
            <header>
                <h2>EZNewsletter FAQS</h2>
            </header>
        </div>
    </div>
    <div class="row 50%">
        <div class="-1u 10u 12u$(medium)">
            <div id="accordion">
                <h3>Whats the difference between a Template and a Draft?</h3>
                <div style="text-align: left;">A Template is a finished Draft that can be used for sending Newsletters.<br />
                    A Draft is an unfinished Template not suitable for sending Newsletters.<br />
                    In fact, you can't send a Draft. You have to save it as a Template first.<br />
                    <br />
                    If you're still confused sacrifice an iPhone to the APP gods!</div>
                <h3>What does modRewrite do?</h3>
                <div style="text-align: left;">modRewrite cleans URLS.<br />
                    <br />
                    In this case it changes all the forward facing URLs (Sign Up, Confirm, Cancel)<br />
                    from for example "../includes/process.asp" to "/thankyou/" for the Sign Up Thank You page<br />
                    <br />
                    If this makes no sense, sacrifice two iPhones to the APP gods!</div>
                <h3>Why do some people get the emails and others don't?</h3>
                <div style="text-align: left;">
                    There could be a few reasons this is happening.

                    <br />
                    <br />
                    <ol>
                        <li style="padding-bottom: 10px;">Your SMTP Server may be blocking some addresses from free accounts like hotmail.com, gmail.com and others. Check with your hosting provider to see if they can white list them.</li>
                        <li style="padding-bottom: 10px;">Your Settings may be set wrong. The SMTP Email address domain should match the SMTP Server domain name. Some shared hosting providers use aliases for Server Names. Contact your provider to see if they can help you in that regard.</li>
                        <li style="padding-bottom: 10px;">You may not have an SPF (Sender Policy Framework) Record set up for your domain name. You should add a TXT record to your DNS Settings so they look something like this (You should be able to get the correct one from you Domain Provider):<br />
                            <pre>
                    <code>
v=spf1 include:yourmailserver.com -all
                    </code>
                </pre>
                        </li>
                        <li style="padding-bottom: 10px;">You may be trying to send too many emails at one time. Try sending fewer emails at a time.</li>
                        <li style="padding-bottom: 10px;">You may have hit your sending limit set by your provider. Check with them to find out if it can be changed.</li>
                        <li>The APP gods are displeased! You must sacrifice three iPhones to appease them!</li>
                    </ol>
                </div>
                <h3>What SMTP Port should I use?</h3>
                <div style="text-align: left;">
                    There are five choices:

                    <ol>
                        <li>25 = Default port. Not secure, Most providers block port 25 for incoming mail. Should only be used as a last option.</li>
                        <li>80 = Some providers like GoDaddy require you to use this secure port.</li>
                        <li>465 = SSL (Secure Socket Layer) Use this if your mail server uses a SSL.</li>
                        <li>587 = TLS (Transport layer security) Recommended! Secure, most providers require this port for SMTP.</li>
                        <li>2525 = Mirror of 587. Use this if port 587 is blocked.</li>
                    </ol>
                    If none of these work sacrifice five iPhones to the APP gods!

                </div>
                <h3>Credits</h3>
                <div style="text-align: left;">
                    <ul>
                        <li>Editor ------> <a target="_blank" href="https://ckeditor.com/">CKEditor</a></li>
                        <li>JQuery ----> <a target="_blank" href="https://api.jqueryui.com/">JQuery UI</a></li>
                        <li>Pop-Ups --> <a target="_blank" href="http://fancyapps.com/fancybox/">Fancybox Version 2</a></li>
                        <li>Icons ------> <a target="_blank" href="https://fontawesome.com/">Font-awesome</a></li>
                    </ul>
                    This APP was built with <a href="https://github.com/ajlkn/baseline"><strong>Baseline</strong></a> a boilerplate for creating new projects.

                    <br />
                    <br />
                    The APP gods are pleased with this!

                </div>
            </div>
             <div class="12u$">
                <hr class="major" style="margin: 1em 0;" />
            </div>
            <div class="12u$">
                <style>
	                div.thenews {
		                float: left;
		                position: relative;
		                display: block;
		                width: 100%;
	                }

	                div.show-more {
		                float: left;
		                position: absolute;
		                left: 0;
		                bottom: 0;
		                display: block;
		                height: 30px;
		                padding-top: 2px;
		                background-color: #5a5a5a;
		                color: #ffffff;
		                cursor: pointer;
		                width: 100%;
		                border-radius: 4px;
	                }
                </style>
                <script type="text/javascript">
                    function toggle_visibility(e) {
                        if (document.getElementById(e), $(".thenews").toggleClass("expanded"), "80px" == document.getElementById("show-news").style.height) {
                            var t = document.getElementById("show-news").scrollHeight;
                            document.getElementById("show-news").style.height = (t + 30).valueOf() + "px", document.getElementById("more-news").innerHTML = "<i class='fa fa-arrow-up'></i> Show Less <i class='fa fa-arrow-up'></i></i>"
                        } else document.getElementById("show-news").style.height = "80px", document.getElementById("more-news").innerHTML = "<i class='fa fa-arrow-down'></i> Show More <i class='fa fa-arrow-down'></i>"
                    }
                </script>
                <div class="thenews" style="height: 80px; overflow: hidden; transition: height 2s linear; -webkit-transition: height 2s linear;" id="show-news">
                    <h3>Server Variables</h3>
                    <div class="row">
                        <div class="3u 12u$(medium)" style="border-bottom:solid 1px #000000;"><strong>Variable Name</strong></div>
                        <div class="9u$ 12u$(medium)" style="border-bottom:solid 1px #000000;text-align:left;"><strong>Value</strong></div>
                        <%
                        For each Key in Request.ServerVariables
                            Response.Write "<div class=""3u 12u$(medium)"" style=""border-bottom:solid 1px #000000;"">" & Key & "</div>"&vbcrlf
	                        Response.Write "<div class=""9u 12u$(medium)"" style=""border-bottom:solid 1px #000000;""><span style=""word-break: break-all;"">&nbsp;" & Request.ServerVariables(Key)& "</div>"&vbcrlf
                        Next
                        %>
                    </div>
                    <div id="more-news" class="show-more" style="text-align: center;" onclick="toggle_visibility('show-news')">
                        <i class="fa fa-arrow-down"></i> Show More <i class="fa fa-arrow-down"></i>
                    </div>
                </div>
            </div>
            <div class="12u$">
                <hr class="major" style="margin: 1em 0;" />
            </div>
            <div class="12u$" id="objcheck">
                <h3>ASP Components</h3>
                <div class="row">
                    <div class="5u 12u$(medium)">
                        <form action="faqs.asp#objcheck" method="post" onclick="document.getElementById('aspinner').style.display = 'inline';">
                            <input type="hidden" name="objcheck" value="show" />
                            <input class="button" type="submit" name="submit" value="Check ASP Components" />
                        </form>
                    </div>
                    <div class="7u$ 12u$(medium)" style="padding-top:7px;">
                        <p id="aspinner" title="Loading..." style="float:left;display:none;">
                            <span class="fa-stack fa-1x" title="Loading...">
                                <i class="fa fa-spinner fa-lg fa-stack-1x fa-inverse fa-spin" style="color:#71b2f0;font-size:30px;"></i>
                            </span>&nbsp;&nbsp;<span>Testing 327 components. This may take a while to load</span>
                        </p>
                    </div>
                </div>
                <% If Request.Form("objcheck") = "show" Then %>
                <script>document.getElementById("aspinner").style.display = "none";</script>
                <%= getResponse("http://dev.aspjunction.com/newsletter/admin/objcheck.asp") %>
                <% End If %>
            </div>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->
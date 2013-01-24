About
=====
When I migrated a classic ASP Website I noticed that ASPMail
(or SMTPSvg.Mailer) from ServerObjects no longer works under Windows Server
2008, let alone 2008 R2 and 64bit. Others have failed as I did, noticing the
same error (Server.CreateObject failed, error 800703e6 - Invalid access
to memory location) - with no solution or to resort to other mailing methods,
for example CDOSYS.

Not wanting to change any code, I resorted to writing a simple wrapper around
the .NET MailMessage class to create a component interface-compatible to the
SMTPSvg.Mailer class.

Far away from beeing identical in behavior and only tested for the specific
website, you can download the source here.

Binary Download
===============
bin/1.0/SMTPSvg.dll

Installation
============
The component must be installed in the GAC to use it from classic ASP. Contact
the MSDN Library on how to do that.
<http://msdn.microsoft.com/en-us/library/dkkx7f79.aspx>

Lincense
========
GPL2 or newer, contact me if you feel you need something else.
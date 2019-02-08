# cmwt
## ConfigMgr Web Tools

## Update 02/07/2019

Adapted the installation guide from Microsoft Word to markdown.  See the new *_InstallGuide.md* file

## Update 05/23/2017

Merged updates into master branch.  Download repository will be decommissioned in lieu of using the cmwt repo native "download" (zip file) features provided by GitHub.

## Update 12/07/2016 - Latest Raw Code

Wiki will be updated when finalized.  This is going to be daily build updates on the path to releasing another ZIP milestone release.

## Update 12/05/2016 - Latest Build Uploaded

Refer to the CMWT GitHub Wiki pages for more information

## Update 12/04/2016 - General Info

I have to say, I didn't expect the response to this that I've seen so far.  I'm really blown away (thank you!).  Some feedback has already pointed out some issues that might trip some people up, so I wanted to post them here:

* The global.asa file has the path to _config.txt hard-coded as "F:\CMWT\\_config.txt" - You may need to edit that to correct the drive letter and/or path.  The first indication that this is incorrect, is an error message "-2147467259: [Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified"  The permanent "fix" for this will be posted in the next build, which is to replace that line within global.asa with this: ConfigFile = Server.MapPath("_config.txt")

* Some questions have come up about how the communications channel is managed between CMWT, SCCM and SQL.  Basically, all read operations are performed via SQL ADO requests, and all write operations are performed via the SMS (WMI) provider channel.  The exception to this are the AD Tools, which use ADSI with Secure LDAP connections.  For the SQL read ops, I chose this because my benchmarking seemed to show a significant performance benefit over read ops via WMI.  In no case are SQL write-ops performed against the SCCM database.  The CMWT database is separate and uses SQL ADO for read and write operations.

## Update 12/03/2016 - Massive Overhaul

Based on user feedback, the application has undergone a massive rewrite.  It is still based on Classic ASP (get mad, I don't care, it still works just fine), but the UI has been replaced with an semi-quasi-Azure-ish kind of menu motif with the "hamburger" control thing at the top left.  The custom reports feature has been rewritten.  The database setup script is new.  The configuration settings script is new.  The installation guide is new (and actually tested by someone other than myself).  Enjoy!

Tested with Configuration Manager 1610 on Windows Server 2012 R2 using SQL Server 2014 SP1.  Client tested on Chrome and IE 11 on Windows 10 (up to 14971).

## Overview

CMWT is a simple, ASP web application, which provides a browser interface to view and manage a Microsoft System Center Configuration Manager primary site.  The intent is to expose a subset of the features provided in the traditional client "admin console", rather than attempt to replace it entirely.  The current version is focused on being lightweight, responsive, and intuitive.  Not all features are fully developed.  Future development will depend on public interest and feedback.  Almost anything is possible to add or improve upon in this project, so if there's something you want or need, let me know.

There are a few files which invoke the "Canvas" JavaScript library for generating graphical charts.  This is not vital to the project, but included for convenience.  If you intend to use this project in a business environment, note that Canvas requires a separate license and compliance with separate terms.  The Installation and Configuration guide is included in the .ZIP archive within the "docs" subfolder.

# cmwt
ConfigMgr Web Tools

## Update 12/03/2016 - Massive Overhaul

Bsaed on user feedback, the application has undergone a massive rewrite.  It is still based on Classic ASP (get mad, I don't care, it still works just fine), but the UI has been replaced with an semi-quasi-Azure-ish kind of menu motif with the "hamburger" control thing at the top left.  The custom reports feature has been rewritten.  The database setup script is new.  The configuration settings script is new.  The installation guide is new (and actually tested by someone other than myself).  Enjoy!

## Overview

CMWT is a simple, ASP web application, which provides a browser interface to view and manage a Microsoft System Center Configuration Manager primary site.  The intent is to expose a subset of the features provided in the traditional client "admin console", rather than attempt to replace it entirely.  The current version is focused on being lightweight, responsive, and intuitive.  Not all features are fully developed.  Future development will depend on public interest and feedback.  Almost anything is possible to add or improve upon in this project, so if there's something you want or need, let me know.

There are a few files which invoke the "Canvas" JavaScript library for generating graphical charts.  This is not vital to the project, but included for convenience.  If you intend to use this project in a business environment, note that Canvas requires a separate license and compliance with separate terms.  The Installation and Configuration guide is included in the .ZIP archive within the "docs" subfolder.

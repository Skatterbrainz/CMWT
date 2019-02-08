# CMWT Installation Guide

## Overview

This document explains how to install and configure CMWT on a System Center Configuration Manager site system.  
Note that the site system must have the SMS Provider role.  CMWT works with CAS and standalone primary site hierarchies.  

CMWT has been tested on Windows Server 2012 R2, SQL Server 2014 and Configuration Manager 1610 (5.00.8458.1000), 
using Microsoft Internet Explorer, Microsoft Edge and Google Chrome web browsers. Note that features may not 
behave identically in different browsers.  It should work equally as well on Windows Server 2016.

## Installation Process

File System Preparation

  1.	Create a Folder on the Site Server named CMWT (e.g. F:\CMWT)
  2.	Extract the ZIP contents (files and folders) into the CMWT target folder

## CMWT Configuration Settings

There are two (2) modes for configuring global settings for CMWT: Express and Manual.  Express configuration 
uses a script to walk through the settings individually.  Manual mode involves locating and editing the 
"_config.txt" settings file.  For details about settings, refer to Appendix B, and C.

### Express Mode

 1.	Double-click the script file "config.vbs" located in the CMWT installation folder.
 2.	Review and modify the values for each setting to suit your environment
 3.	When finished, the settings are written to the _config.txt file, and the original is backed up as _config.bak.

### Manual Mode

 1.	Edit the file _config.txt” , located in the CMWT installation folder.
 2.	Review and modify the values for each setting to suit your environment

### Permissions

 1.	Configure NTFS permissions on the CMWT folder
 2.	Refer to the following example for NTFS security settings.  Essentially, make sure that whatever account is 
 used by the IIS application pool to read the CMWT physical folder contents has Read permissions on the physical folder.

### Database Preparation

Note: The CMWT database can reside on the same SQL Server instance as the ConfigMgr database, or under a separate 
instance, or on a separate SQL Server host altogether.  If you choose to place the CMWT database on the same 
SQL Server instance as ConfigMgr, be sure to account for performance tuning to give ConfigMgr higher priority to resources.

_NOTE:_ Installing third-party databases on the same instance as a Configuration Manager SQL instance is not supported
by Microsoft and my violate licensing terms.  We recommend using a separate/different SQL instance.

 *	Open SQL Server Management Studio
 *	Connect to the CM database instance
 *	Create a new Database named “CMWT”
 *	Click File / Open
 *	Browse to locate the file *"cmwt_db_setup.sql"*
 *	When it opens in SSMS, click Run (or press F5)


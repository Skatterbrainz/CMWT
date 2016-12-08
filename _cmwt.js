/*
****************************************************************
Filename..: cmwt.js
Author....: David M. Stein
Date......: 12/07/2016
Purpose...: javascript library
****************************************************************
*/

function explorer (computer) {
	var x = "file://"+computer+"/c$";
	var y = window.open(x, "mywin");
}

function ccmlogs (computer) {
	var x = "file://"+computer+"/c$/windows/ccm/logs";
	var y = window.open(x, "mywin");
}

function manage (computer) {
	var cmd = "compmgmt.msc -a /computer="+computer;
  	try   {
  		var objShell = new ActiveXObject("wscript.shell");
  		objShell.Run(cmd);
  	} catch(e) {
  		alert(e);
  	}
}

function winrs (computer) {
	var cmd = "file://cm01/cmwt$/tools/scripts/winrs.bat";
  	try   {
  		var objShell = new ActiveXObject("wscript.shell");
  		objShell.Run(cmd+" "+computer);
  	} catch(e) {
  		alert(e);
  	}
}

function rdp (computer) {
	var cmd = "mstsc -v "+computer;
  	try   {
  		var objShell = new ActiveXObject("wscript.shell");
  		objShell.Run(cmd);
  	} catch(e) {
  		alert(e);
  	}
}

function cmwthelp () {
	alert("These buttons are only enabled for IE browser sessions.\nYou must also have the URL in your Local Intranet zone\nand have enabled \"unsafe\" ActiveX Scripts in that zone.");
}

//Modify the path to your applications!
var rdp='C:\\Windows\\System32\\mstsc.exe';
var ssh='C:\\Program Files\\PuTTY\\putty.exe';
var vnc='C:\\Program Files\\RealVNC\\VNC Viewer\\vncviewer.exe';
var telnet='C:\\Program Files\\PuTTY\\putty.exe';
var ping='C:\\Windows\\System32\\ping.exe';
var sftp='C:\\Program Files\\FileZilla FTP Client\\filezilla.exe';

var logfile = "c:\\ggconnect\\ggconnect-log.csv";

//set command line arguments as destination	
var destination=(WScript.Arguments(0));

//find out the URI you are opening
var uri=destination.substring(0, destination.indexOf(":"));
var uriFormat=uri + "://";

var currentdate = new Date(); 
	if (currentdate.getMonth() < 10)
	{
		month = '0' + currentdate.getMonth();
	}else{
		month = currentdate.getMonth();
	}
	if (currentdate.getDate() < 10)
	{
		date = '0' + currentdate.getDate();
	}else{
		date = currentdate.getDate();
	}
	if (currentdate.getHours() < 10)
	{
		hours = '0' + currentdate.getHours();
	}else{
		hours = currentdate.getHours();
	}
	if (currentdate.getMinutes() < 10)
	{
		minutes = '0' + currentdate.getMinutes();
	}else{
		minutes = currentdate.getMinutes();
	}
	if (currentdate.getSeconds() < 10)
	{
		seconds = '0' + currentdate.getSeconds();
	}else{
		seconds = currentdate.getSeconds();
	}
var datetime = currentdate.getFullYear() + "/" 
				+ month + "/"
                + date + ","  
                + hours + ":"  
                + minutes + ":" 
                + seconds;
	
//open the appropriate application
switch(uri)
{
	case "rdp":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Exec(rdp + " /v:" + destination)
		break;
	case "ping":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Run(ping + " " + destination + " -t")
		break;
	case "vnc":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Exec(vnc + " " + destination)
		break;
	case "ssh":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Exec(ssh + " " + destination)
		break;
	case "sftp":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Exec(sftp + " sftp://" + destination + " --logontype=ask")
		break;
	case "telnet":
		cleanDestination();
		var ws = new ActiveXObject("WScript.Shell")
		ws.Exec(telnet + " telnet://" + destination)
		break;
	
}

function writelogfile(filename, output)
{
	var fso  = new ActiveXObject("Scripting.FileSystemObject");
	var fh = fso.OpenTextFile(filename, 8, true);
	fh.WriteLine(output); 
	fh.Close(); 
}

function cleanDestination()
{
	if(destination.indexOf('|') >= 0 || destination.indexOf('&') >= 0 || destination.indexOf('<') >= 0 || destination.indexOf('>') >= 0 || destination.indexOf('%3E') >= 0 || destination.indexOf('%3C') >= 0 || destination.indexOf('%7C') >= 0){
				WScript.Echo("|, <, > or & found in your link, bad robot.\nyou can no longer trust links from this source\nThe command you clicked was " + destination + "\nFor the moment you are about to see an error but atleast it didn't run the command and has been logged!");
				writelogfile(logfile, "blocked" + "," + destination + "," + datetime + ",aborted");
				throw '';
	}
	destination=destination.replace(uriFormat, '');
	destination=destination.replace('/', '');;
	destination=destination.replace('\\', '');
	writelogfile(logfile, uri + "," + destination +"," + datetime + ",run");
}

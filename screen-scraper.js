/* 
	This script screen scrapes data from the <UNDEFINED> website.
	- Dump files are placed into the "dump" folder and a subfolder with the category name.
	- In the script, cat is short for category. Manuf is short for manufacturer.
*/

/* -------------- USER INPUTS -------------------------------------------- */
var wait = 1000; 	// standard pause in between http calls
var rndWait = 2000; // random pause range (0-X000) to randomize intervals
var manufLimit = null; // null or some value. Used for testing, limit the number read attempts.
var multiCats = []; // empty, one or many values
/* ------------------------------------------------------------------------*/

var urlPrefix = "http://www.undefined.com/folder/";
var urlCats = urlPrefix + "software-categories.asp";
var CR = "\n";
var log = "log.txt";
var totCats = 0;
var totManufs = 0;
var clearLog = false;

var m = {
	WRITE   : 2, 
	CREATE  : 1,
	APPEND	: 8,
	ASCII	: 0
}

var status = {
	OK 		: 200
}


/* -----------------------------------------------------------------------------------
	Get a list of href anchors from a url
 ---------------------------------------------------------------------------------- */	
var getUrls = function(url) {

	var urls = [];
	var xhr = httpGet(url);
	
	if (xhr.status === status.OK) {		
		
		var html = xhr.responseText;
		
		// Parse through each line of the html and grab urls and labels
		var parsedHtml = html.split(CR);
		
		for (var i = 0; i <= parsedHtml.length; i++) {
			
			var line = parsedHtml[i];
			if (line) {
				if (line.indexOf("<!-- #esid-results -->") > -1) break;
				if (line.indexOf("<a href=") > -1) {
					
					var href = null;				
					line.replace(/[^<]*(<a href="([^"]+)">([^<]+)<\/a>)/g, function () { // pull out href anchors
						href = Array.prototype.slice.call(arguments, 1, 4);
					});				
				
					if (href) {
						var url = href[1];
						var label = href[2];
					
						if (!existsInArray(urls, label, 2)) // avoid duplicates
							urls.push([url, label]);	
					}
				}
			}
		}
		
	} else {
		WScript.stdout.write(String.format("Failed to parse url: {0}, Error: {1}", url, xhr.status));
	}

	return urls;
};


/* -----------------------------------------------------------------------------------
	Read contents of manufacturer http response contents and save to file
 ---------------------------------------------------------------------------------- */	
var dumpPageContents = function(manufUrl, catLabel, manufLabel, folderSuffix) {

	if (manufUrl.indexOf("software-company-applications.asp") < 0) return;
	
	var xhr = httpGet(manufUrl);
	
	if (xhr.status === status.OK) {
		var html = xhr.responseText;
		saveFile(catLabel, manufLabel, html, folderSuffix);
	} else {
		WScript.stdout.write(String.format("Failed to dump page: {0}, Error: {1}", manufUrl, xhr.status));		
	}
};


/* -----------------------------------------------------------------------------------
	Save http response contents to file. 
	Ex: dump\[category] - [code1] - [code2]\[category][Manufacturer].html
	    dump\Access - 43232900 - 43232901\AccessAccess Control Technology.html
 ---------------------------------------------------------------------------------- */	
var saveFile = function(catLabel, manufLabel, html, folderSuffix) {

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var file = catLabel + "-" + manufLabel;
	var file = file.replace(/[^A-Z\d\s]/gi, '').replace(/\s+/g, ' '); // clean up
	var dir = "dump\\" + catLabel + " - " + folderSuffix + "\\";
	
	try {
		
		if (!fso.FolderExists(dir))
			fso.CreateFolder(dir);
		
		var output = fso.OpenTextFile(dir + file + ".html", m.WRITE, m.CREATE, m.ASCII);		
		output.Write(html);
		output.Close();	
		WScript.stdout.write(" [Dumped!]");
		
	} catch (e) {
		writeToLog(String.format("{0}FAILED: could not dump {1}.", this.CR, file));
	}	
};


/* -----------------------------------------------------------------------------------
	Find item in one or two-dimensional array
 ---------------------------------------------------------------------------------- */	
var existsInArray = function(array, itemToFind, dimensions) {	
	
	for (var i = 0, len = array.length; i < len; i++) {
		
		if (dimensions == 1) {
			if(array[i] === itemToFind)
				return true;
		} else {
			if(array[i][1] === itemToFind)
				return true;
		}
	}

	return false;
};


/* -----------------------------------------------------------------------------------
	Write to log file in current directory
 ---------------------------------------------------------------------------------- */	
var writeToLog = function(line) {

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var file = "log.txt";
	var output = fso.OpenTextFile(file, m.APPEND, m.CREATE, m.ASCII);		
	output.Write(line);
	output.Close();
	WScript.stdout.write(line);
};


/* -----------------------------------------------------------------------------------
	Return a random number from 1-X (seconds) + default wait time
 ---------------------------------------------------------------------------------- */	
var randomizer = function() {
	
	return Math.floor(Math.random() * 
		this.rndWait) + 
			this.wait;
};


/* -----------------------------------------------------------------------------------
	Make http call and return http object
 ---------------------------------------------------------------------------------- */	
var httpGet = function(url) {
	
	var xhr = new ActiveXObject("MSXML2.serverXMLHTTP.6.0");
	xhr.open("GET", url);
	WScript.Sleep(1000); // Give it a moment to open website
	xhr.send();
	return xhr;
};


var deleteLogFile = function() {
	
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	if (fso.fileExists(this.log)) fso.DeleteFile(this.log);
};


if (!String.format) {
  
  String.format = function(format) {
    var args = [].slice.call(arguments, 1);
		return format.replace(/{(\d+)}/g, function(match, number) { 
			return typeof args[number] != 'undefined' ? args[number] : match;
		});
  };
};


writeToLog("----------------------------Starting Script-----------------------------");

if (this.clearLog) deleteLogFile
var catUrls = getUrls(urlCats);
var manufUrls = null;
var folderSuffix = null;

// Loop through categories
 for (var i = 0; i <= catUrls.length; i++) {
  
	var catUrlArr = catUrls[i];  	
	if (!catUrlArr) continue;
	var catUrl = catUrlArr[0];
	var catLabel = catUrlArr[1];
	
	// Get the url's query paramters and pass in as suffix for folder name
	if (catUrl.indexOf("software-category-applications.asp") > -1) {
		var idx = catUrl.indexOf("?mc=");
		var folderSuffix = catUrl.substr(idx+4);
		folderSuffix = folderSuffix.replace("&id=", " - ");
	}		
	
	// This is the end of their page, bail
	if (catLabel == "Express Labs" ) break;
	
	// Are we only processing categories in multiCats[]?
	if (this.multiCats.length != 0) {
		if (!existsInArray(this.multiCats, catLabel, 1)) {
			WScript.stdout.write(String.format("{0}SKIPPING category {1} due to not in multiCats[].", this.CR, catLabel));
			continue;			
		}
	}

	var logMsg = String.format("{0}#{1}: CATEGORY QUERY - {2} - {3} -> {4}", 	    
		this.CR,
		this.totCats+1,
		new Date(), 
		catLabel, 
		catUrl);
		
	writeToLog(logMsg);
	
	// Grab list of manufacturer urls from current category page
	manufUrls = getUrls(this.urlPrefix + catUrl);
	var ctr = 0;

	// Loop through manufacturer urls and dump page to file
	for (var y = 0; y <= manufUrls.length; y++) {
		
		if (manufLimit) {
			if (y == (manufLimit+2)) {			
				writeToLog(String.format("{0}     Manufacturer limit of {1} reached...", this.CR, this.manufLimit));
				break;
			}
		}

		var rndWait = randomizer();
		WScript.Sleep(rndWait);
		
		var manufUrlArr = manufUrls[y];		
		if (!manufUrlArr) continue;		
		var manufUrl = this.urlPrefix + manufUrlArr[0];
		manufUrl = manufUrl.replace("\/.\/", "\/");
		var manufLabel = manufUrlArr[1];		
		var curDate = new Date();
		
		// Skip if we run into this
		if (manufUrl.indexOf("pc-application.asp") > -1) continue;
		
		logMsg = String.format("{0}     #{1} MANUF QUERY - {2} - {3} -> {4} -> {5}   |   paused {6} seconds", 
			this.CR,
			++ctr,
			new Date(), 
			catLabel, 
			manufLabel,
			manufUrl,
			rndWait / 1000);
		
		writeToLog(logMsg);					
		dumpPageContents(manufUrl, catLabel, manufLabel, folderSuffix);
		this.totManufs++;
    }
	
	this.totCats++;
};

	
writeToLog(String.format("{0}{0}Total Categories processed: {1}, Total Manufacturer links processed: {2}",
	this.CR,
	this.totCats,
	this.totManufs));
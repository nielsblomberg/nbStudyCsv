/*
	© Niels Blomberg, 15-16 April 2020
	Sum data from several XML files in grouped CSV files
*/
// SETTINGS
var setTypes=["Memory"]	// Select 1 or more types for output
var setLogPath=".\\nbStudyCsv.log";
var setLogOverwrite=true;  	// true=create new file, false=append to old fileCreatedDate
var setVerbose=false; 		// true=show many details; false=show 1 line at the beginning, changes made during CSV, errors found
var setCombinedMarker=true;	// Use marker combinations divided by slash?
var setDecimal=","	;		// Can be , or .	Currently unused
var setSeparator=";";		// Can be ; or ,	CSV field separator, defined in Windows Regional settings
var setXmlPath=".\\XML_Input";	// Path of XML input files. Needs to be available
var setCsvPath=".\\CSV_Output";	// Path of CSV output files. Will be created if necesary
var setRyPrefix=["Lymphocytes/Single#1/Single#2/B cells"]	// The beginning of each line

function objLogFile(pPath,pOverwrite) {
function Xschryf(tekst) {
	this.file.WriteLine(new Date()+": "+tekst);
}
function Xsluit() {
	this.file.Close();
	this.file=null;
	this.fso=null;
}
this.schryf=Xschryf;
this.sluit=Xsluit;
this.fso = new ActiveXObject( "Scripting.FilesystemObject" );
// https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
this.file = this.fso.OpenTextFile(pPath,pOverwrite?2:8,true,-2);
}

var logFyl=new objLogFile(setLogPath,setLogOverwrite);
logFyl.schryf("Start");
var hetObject={};
try {
	var mapXml;
	try {
		mapXml = logFyl.fso.GetFolder(setXmlPath);
		if (setVerbose) logFyl.schryf("XML folder: "+mapXml.Path);
	} catch(e) {
		throw {name:"custom error",message:"XML folder "+setXmlPath+" is missing"} 
	}
	
	var mapCsv;
	var foutjeBedankt=false;
	try {
		mapCsv = logFyl.fso.GetFolder(setCsvPath);
		if (setVerbose) logFyl.schryf("CSV folder: "+mapCsv.Path);
	} catch(e) {foutjeBedankt=true;}
    if (foutjeBedankt) try {
		foutjeBedankt=false;
		mapCsv = logFyl.fso.CreateFolder(setCsvPath);
		if (setVerbose) logFyl.schryf("CSV folder created: "+mapCsv.Path);
	} catch(e) {
		throw {name:"custom error",message:"CSV folder "+setCsvPath+" is missing"} 
	}
	
	var fc = new Enumerator(mapXml.Files);
	for (fc.moveFirst();!fc.atEnd(); fc.moveNext()) { 
		var fXml=fc.item();
		try {
			// Open het bestand
			var	objXml = new ActiveXObject("MSXML2.DOMDocument");
			objXml.async = false; 
			objXml.setProperty("SelectionLanguage", "XPath");
			objXml.load(fXml.Path);
			var xmlThTop=objXml.documentElement;
			if (xmlThTop.nodeName!="Table") throw {name:"incorrect outer tab",message:"expecting 'Table'"}
			// All XPath: https://devhints.io/xpath

			// Bepaal de naam van de studie
			/* OUWE: 
			var study=fXml.Name+""; // Not sure this is necessary, but the file should never be renamed
			if (study.length==study.toLowerCase().indexOf("_%.xml")+6) {
			study=study.substr(0,study.length-6);
			} else if (study.length==study.toLowerCase().indexOf(".xml")+4) {
				study=study.substr(0,study.length-4);
			}
			*/
			var study=xmlThTop.selectSingleNode("@workspaceTitle").text;
			var stCategory=
				(study.toUpperCase().indexOf('CSA')!=-1||fXml.Name.toUpperCase().indexOf('CSA')!=-1)?'CSA':
				(study.toUpperCase().indexOf('DFR')!=-1||fXml.Name.toUpperCase().indexOf('DFR')!=-1)?'DFR':
				'RA';
			logFyl.schryf("Opening "+fXml.Path+" ; Study "+study+" ; Category "+stCategory);
			if (hetObject[stCategory]==null) hetObject[stCategory]={};
			var objCat=hetObject[stCategory];

			// Haal de kolom de kop?
			var kopMarker=xmlThTop.selectSingleNode("@name").text;
			var heeftAlleMarkers=kopMarker=="%";
			if (setVerbose) logFyl.schryf("Alle markers? "+heeftAlleMarkers);
			
			// Struin de kolommmen af
			var xmlKolommen=xmlThTop.selectSingleNode("Columns");
			var xmlPrefix=xmlKolommen.selectSingleNode("Column[@id='0']/@analysisPath").text;
			var isPrefixOnbekend=true;
			for (var telPref=0;telPref<setRyPrefix.length ;telPref++) 
			if (xmlPrefix.indexOf(setRyPrefix[telPref])==0) {
				isPrefixOnbekend=false;
				if (xmlPrefix== setRyPrefix[telPref]) {
					if (setVerbose) logFyl.schryf("Known prefix found: \""+xmlPrefix+"\"");
				} else {
					var foutRegel="First line starting with \""+xmlPrefix+"\"";
					xmlPrefix=setRyPrefix[telPref];
					foutRegel+="; searching for prefix \""+xmlPrefix+"\""
					if (setVerbose) logFyl.schryf(foutRegel);
				}
			}
			if (isPrefixOnbekend) {
				logFyl.schryf("WARNING: Found unknown prefix \""+xmlPrefix+"\"");
			}
			if (setVerbose) logFyl.schryf(xmlPrefix);
			for (var teller in setTypes) {
				var typeStudy=setTypes[teller];
				if (setVerbose) logFyl.schryf("Type Study: "+typeStudy);
				if (objCat[typeStudy]==null) objCat[typeStudy]={};
				var objTypeStudy=objCat[typeStudy];
				
				var deZoeker="Column[starts-with(@analysisPath,\""+xmlPrefix+"/\") and "+
					(heeftAlleMarkers
					?"contains(@analysisPath,\"/"+typeStudy+"/\")]"
					:"contains(@analysisPath,\"/"+typeStudy+"\") and substring(@analysisPath, string-length(@analysisPath) - string-length(\"/"+typeStudy+"\") +1) = \"/"+typeStudy+"\"]"); //Helaas werkt ends-with niet! Oplossing gevonden https://stackoverflow.com/questions/22436789/xpath-ends-with-does-not-work
				if (setVerbose) logFyl.schryf(deZoeker);
				var ryNoodKolommen=xmlKolommen.selectNodes(deZoeker);
				for (var nodeNo=0;nodeNo<ryNoodKolommen.length;nodeNo++) {
					var nuNood=ryNoodKolommen[nodeNo];
					var idNode=nuNood.selectSingleNode("@id").text;
					var anapadNood=nuNood.selectSingleNode("@analysisPath").text;
					if (false && setVerbose) logFyl.schryf(anapadNood);
					
					// The marker, defining a file
					var markerNood=heeftAlleMarkers
						?anapadNood.substr(anapadNood.indexOf("/"+typeStudy+"/")+typeStudy.length+2,anapadNood.length)
						:kopMarker;
					if (setCombinedMarker || markerNood.indexOf("/")==-1) { // Slash in the marker? This is a combination, which is optionallly OK
						// The slash is OK, but not in the output file name; get rid of it!
						while (markerNood.indexOf("/")!=-1) markerNood=markerNood.replace("/","_");
					
						// The group, X-axis value
						var GroupNood=anapadNood.substring(xmlPrefix.length+1,anapadNood.indexOf("/"+typeStudy));
					
						if (objTypeStudy[markerNood]==null) objTypeStudy[markerNood]={studies:{},groups:{}}
						objFileInfo=objTypeStudy[markerNood];
						if (objFileInfo.studies[study]==null) objFileInfo.studies[study]={};
						objStudy=objFileInfo.studies[study];
						if (objFileInfo.groups[GroupNood]==null) objFileInfo.groups[GroupNood]="";
					
						var ryNoodWaarden=xmlThTop.selectNodes("Row/Entry[@colId="+idNode+" and text() != \"\"]")
						var waardeNood=ryNoodWaarden.length==0?null:ryNoodWaarden[0].text;
						var waardeFout=ryNoodWaarden.length<2?null:ryNoodWaarden[0].text;
						if (ryNoodWaarden.length>1) try {
							waardeFout+="("+ryNoodWaarden[0].selectSingleNode("../@sample").text+")";
						} catch(e) {}
					
						for (var telNW=1; telNW<ryNoodWaarden.length; telNW++) 
						{
							if (waardeNood=="0")waardeNood=ryNoodWaarden[telNW].text;
							waardeFout+=" ; "+ryNoodWaarden[telNW].text;
							try {
								waardeFout+="("+ryNoodWaarden[telNW].selectSingleNode("../@sample").text+")";
							} catch(e) {}
						}
						if (waardeNood==null) {
							logFyl.schryf("WARNING No Value found: study "+study+", category "+stCategory+", marker "+markerNood+", group "+GroupNood);
						} else {
							objStudy[GroupNood]=waardeNood;
							if (waardeFout!=null) logFyl.schryf("WARNING multiple values found: "+waardeFout+"; Selected value "+waardeNood+" for study "+study+", category "+stCategory+", marker "+markerNood+", group "+GroupNood);						
							else if (setVerbose) logFyl.schryf("Value "+waardeNood+" found study "+study+", category "+stCategory+", marker "+markerNood+", group "+GroupNood);
						}
					} // Combined Marker
				} // for nodeNo in kolommmen ryNoodKolommen
				/* ----- */ 
			}
			objXml=null;
		} catch(e) {
			logFyl.schryf("ERROR: Moving to next XML after "+e.name+": "+e.message);
		} 		

	}
	
	if (setVerbose) logFyl.schryf("Creating CSV files");
	for (var telCat in hetObject) 
	for (var telTypeStudy in hetObject[telCat])
	for (var telMarker in hetObject[telCat][telTypeStudy])
	{
		var objFileInfo=hetObject[telCat][telTypeStudy][telMarker];
		var fylNaam=telCat+"_"+telTypeStudy+"_"+telMarker+".csv";
		var fylCSV;
		try {
			try {
				fylCSV=mapCsv.CreateTextFile(fylNaam,true,false);
			} catch (e) {
				throw {name:e.name, message:"Is the file open in Excel? If so, pleacse close. Original message: "+e.message}
			}
			if (setVerbose) logFyl.schryf("Category "+telCat+", Type "+telTypeStudy+", Marker "+telMarker+": "+objFileInfo);
		
			// Kopregel
			var ryGroepen=[];
			fylCSV.Write("\"Study\"");
			for (telGroup in objFileInfo.groups) {
				fylCSV.Write(setSeparator+"\""+telGroup+"\"");
				ryGroepen[ryGroepen.length]=telGroup;
			}
			fylCSV.WriteLine("");
	
			for (telStudy in objFileInfo.studies) {
				fylCSV.Write("\""+telStudy+"\"");
				rijWaarden=objFileInfo.studies[telStudy];
				for (var teller=0;teller<ryGroepen.length;teller++) {
					var waardeNu=rijWaarden[ryGroepen[teller]];
					if (waardeNu==null) fylCSV.Write(setSeparator+"\"\"");
					else fylCSV.Write(setSeparator+waardeNu);
				}
				fylCSV.WriteLine("");			
			}
			logFyl.schryf("Created "+mapCsv.path+"\\"+fylNaam)
		} catch (e) {
			logFyl.schryf("ERROR when creating "+mapCsv.path+"\\"+fylNaam+" : "+e.name+": "+e.message);
		}
		
		try {
			fylCSV.Close();
		} catch (e) {
			// Dit kan falen als het bestand niet geopend is; dat hoef de gebruiker niet te weten
		}
		fylCSV=null;
	}

} catch (e) {
	logFyl.schryf("ERROR: Unable to proceed after "+e.name+": "+e.message);
}

logFyl.schryf("Stop");
logFyl.sluit();
logFyl = null;
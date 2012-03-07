/* Acrobat 7.0 or later required
Prior to running this script for the first time, be sure the global variable global.startBatch is undefined. If necessary, execute delete global.startBatch from the console.
*/

/*
Author: Luke Pederson
Naval Postgraduate School - Dudley Knox Library
Email: lbpeders@nps.edu
*/

try{
//Trim
function trim(str) {
		str = str.replace(/\s{2,}/g," ");
		str = str.replace(/^\s+|\s+$/g,"");
		str = str.replace(/\(\s/g,"(");
		str = str.replace(/\s\)/g,")");
		return str;
}

function toTitleCase(str)
{
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function fixAuthor(str){
	//Seperate authors
	str = str.replace(/,/g,";");
	str = str.replace(/(\s+);/g,";");
	str = str.replace(/;(\s+)/g,";");
	
	//Now we swap first and last names -> last, first m.
	//Multiple Authors
	if(str.search(";") != -1)
	{
		var Authors = str.split(";");
		str = "";
		len = Authors.length;
		for(var i = 0 ; i < len ; i++)
		{
			if(Authors[i].search(" ") != -1 && Authors[i].length > 0)
			{
				//Grab last name & first name
				var lastName = Authors[i].substring(Authors[i].lastIndexOf(" "),Authors[i].length);
				var firstName = Authors[i].substring(0,Authors[i].lastIndexOf(" "));
				Authors[i] = lastName + ", " + firstName;
				Authors[i] = trim(Authors[i]);
				str += Authors[i];
				if((i+1) != len)
					str += ";";
			}
			else
			{
				str += Authors[i];
				if((i+1) != len)
					str += ";";
			}
		}
	}
	//1 Author
	else
	{
		if(str.search(" ") != -1 && str.length > 0)
		{
			var lastName = str.substring(str.lastIndexOf(" "),str.length);
			var firstName = str.substring(0,str.lastIndexOf(" "));
			str = lastName + ", " + firstName;
			str = trim(str);
		}
		else
		{
			str = trim(str);
		}
	}
	return str;
}


if ( typeof global.startBatch == "undefined" ) {
	global.startBatch = true;
	
	// When we begin, we create a blank doc in the viewer to hold the
	// attachment.
	global.myContainer = app.newDoc();
	global.myContainer.addWatermarkFromText("The generated spreadsheet is an attachment of this document.\nTo locate: Open the \"View\" dropdown in the Menu Bar\nThen locate \"Navigation Panels\", and open \"Attachments\".",0,font.Helv,20,color.black);
	
	// Create an attachment and some fields separated by tabs
	global.myContainer.createDataObject({
	cName: "mySummary.xls",
	cValue: "FileName\tContributorAuthor\tTitle\tDate\tDateISO\tPublisher\tAbstract\tSubjectTerms\tAdvisor\tSecondReader\tDegreeName\tDegreeLevel\tDegreeDiscipline\tAffiliation\r\n"});
}
/*
else
{
	delete global.startBatch;
}
*/

/************************** Extract first 7 Pages *********************/
    if(this.numPages > 6)
    {
		var OLine = "";
        var dataLine = "";

        /* Spell Check a Document */
        var Word, numWords, i, j;
        for ( i = 0; i < 6; i++ )
        {
            numWords = this.getPageNumWords(i);
            for ( j = 0; j < numWords; j++)
                {
                    Word = this.getPageNthWord(i,j,false);
                        if ( Word != null )
                        {
                               dataLine += Word;
                       }
                }
        }
		
		//Remove extraneous white space
		dataLine = dataLine.replace(/\s{2,}/g," ");
		
		// Partial Dataline Pg 5 - Thesis Advisor and Second Reader
		var pDataLine = "";
		for ( i = 3; i < 6; i++ )
        {
            numWords = this.getPageNumWords(i);
            for ( j = 0; j < numWords; j++)
                {
                    Word = this.getPageNthWord(i,j,false);
                        if ( Word != null )
                        {
                               pDataLine += Word;
                       }
                }
        }
		
		// Another partial Dataline pg5 - Affiliation (Navy, Army, etc.)
		// Check below for usage...
		var apDataLine = pDataLine;
		
		//This gets set in Author
		var affAuthor = "";
		
/****************************** End Extract ************************************/
	
	/**************************** Special Chars & Variable Ini ***************************/
	//Replace special characters - Everything in brackets is the white-list [whitelist]
	dataLine = dataLine.replace(/[^\w\s\n\r\/.,\[\]{}\\~`;;"'\|!@#$%^&*()+=_-]/gi, "")
	
	var captured = [];
	var len;
	var re = "";
	
	//For checking whether to get [advisor and second reader] from [1st or 5th] page...
	var Flag = true;
	
	//Insert file name is first tab
	var fileName = this.documentFileName;
	OLine += fileName + "\t";
	console.println("");
	console.println("**********Begin: Errors for " + fileName + "*************");
	/*********************************** End Special Char *********************************/
	
	
/*******************************************Begin: Finding Advisor*********************************************/
	//Narrow down 5th page data
	re = /(?:Approved)(?:\s{0,2})(?:by)(?:[:]{1})((.|\n)*)(?:iii)(?:\s{0,2})(?:\r\n)(?:This)(?:\s+)(?:Page)/i;
	try
	{	
		//Variant 1 "Approved by: ---- iii\nThis Page..."
		if(re.test(pDataLine))
		{
			captured = re.exec(pDataLine);
			if(trim(captured[1]))
			{
				Flag = true;
				pDataLine = captured[1];
			}
			else
			{
				if(trim(captured[0]))
				{
					Flag = true;
					pDataLine = captured[0];
				}
				else
				{
					Flag = false;
				}
			}
		}
		else
		{
			re = /(?:Approved)(?:\s{0,2})(?:by)(?:[:]{1})((.|\n)*)(?:iv)(?:\s{0,2})(?:\n)(?:This)(?:\s+)(?:Page)/i;
			//Variant 2 "Approved by: ---- iv\nThis Page..."
			if(re.test(pDataLine))
			{
				captured = re.exec(pDataLine);
				if(trim(captured[1]))
				{
					Flag = true;
					pDataLine = captured[1];
				}
				else
				{
					if(trim(captured[0]))
					{
						Flag = true;
						pDataLine = captured[0];
					}
					else
					{
						Flag = false;
					}
				}
			}
			else
			{
				Flag = false;
			}	
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}	
	var Advisor = "";
	var SReader = "";
	var AFlag = false; //Advisor flag (found?)
	var SFlag = false; //Second Reader flag (found?)
	try
	{
		//Data from 5th page was extracted (Thesis advisor & second reader)
		if(Flag == true)
		{	
			pDataLine = pDataLine.replace(/(?:Approved)(?:\s{0,2})(?:by)(?:[:]{1})/i,"");
			pDataLine = pDataLine.substring(0,pDataLine.length-3);
			captured = pDataLine.split("\n");
			len = captured.length - 1;
			
			for(var i = 0 ; i < len ; i++)
			{	
				//Variant 1 - "Thesis Advisor"
				re = /(Thesis)(\s+)(Advisor)/i;
				if(captured[i].search(re) != -1)
				{
					if((i-1) >= 0)
					{
						if(captured[i-1].search(",") != -1)
						{
							captured[i-1] = captured[i-1].replace(/(,)(.*)/i,"");
						}
						if(Advisor.length > 0)
						{
							Advisor += ";" + trim(captured[i-1]);
							AFlag = true;
						}
						else
						{
							//This line breaks loop - No idea why...
							//captured[i-1] = fixAuthor(captured[i-1]);
							captured[i-1] = trim(captured[i-1]);
							Advisor = captured[i-1];
							AFlag = true;
						}	
					}
				}
				else
				{
					//Variant 2 - "Thesis Co-Advisor"
					re = /(Thesis)(\s+)(Co-Advisor)/i;
					if(captured[i].search(re) != -1)
					{
						if((i-1) >= 0)
						{
							if(captured[i-1].search(",") != -1)
							{
								captured[i-1] = captured[i-1].replace(/(,)(.*)/,"");
							}
							if(Advisor.length > 0)
							{
								Advisor += ";" + trim(captured[i-1]);
								AFlag = true;
							}
							else
							{
								Advisor = trim(captured[i-1]);
								AFlag = true;
							}
						}
					}
				}
				//Second Reader - Variant 1 - "Second Reader"
				re = /(Second)(\s+)(Reader)/i;
				if(captured[i].search(re) != -1)
				{
					if((i-1) >= 0)
					{	
						if(captured[i-1].search(",") != -1)
						{
							captured[i-1] = captured[i-1].replace(/(,)(.*)/,"");
						}
						if(SReader.length > 0)
						{
							SReader += trim(captured[i-1]);
							SFlag = true;
						}
						else
						{
							SReader = trim(captured[i-1]);
							SFlag = true;
						}
					}
				}
			}
		} //End 5th page
		else
		//Unable to get 5th page, use 1st
		{
			//Variant 1 - "Thesis Advisor:"
			re = /(?:Thesis)(?:\s{0,4})(?:Advisor)(?:[:]{0,3})(?:\s{0,4})(.*)(?:\n)/i;
			if(re.test(dataLine))
			{
				captured = re.exec(dataLine);
				len = captured.length;
				if(trim(captured[1]))
				{
					Advisor = trim(captured[1]);
					Advisor = fixAuthor(Advisor);
					AFlag = true;
				}
				else
				{
					if(trim(captured[0]))
					{
						console.println("Thesis Advisor: There was a problem with the match seperation, so the ENTIRE match was used...");
						Advisor += trim(captured[0]);
						AFlag = true;
					}
					else
					{
						//Advisor = "ERROR";
						console.println("Unable to capture advisor...");
					}
				}
				captured = [];
				re = "";
				len = 0;
			}
			else
			{
				//Variant 2 - "Thesis Co-Advisors:"
				re = /(?:Thesis)(?:\s{0,4})(?:Co-)(?:Advisor)(?:s?)(?:[:]{0,3})(?:\s{0,4})(.*)(?:\n)/i;
				if(re.test(dataLine))
				{
					captured = re.exec(dataLine);
					if(trim(captured[1]))
					{
						Advisor = trim(captured[1]);
						Advisor = fixAuthor(Advisor);
						AFlag = true;
					}
					else
					{
					if(trim(captured[0]))
						{
							console.println("Thesis Advisor: There was a problem with the match seperation, so the ENTIRE match was used...");
							Advisor += trim(captured[0]);
							AFlag = true;
						}
						else
						{
							//Advisor = "ERROR";
							console.println("Unable to capture advisor...");
						}
					}
					captured = [];
					re = "";
					len = 0;
				}
				else
				{
					//Advisor = "ERROR";
					console.println("Fail in matching Advisor...");
				}
			}
		} //End first Page
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	
/*******************************************End: Finding Advisor*********************************************/
	
/*******************************************Begin: Finding Second Reader *********************************************/
	//SReader - Grab all words until new line... OR if new line NEXT line until new line <---- Goal
	
	//If Second Reader was not caught from Pg5
	if(Flag == false)
	{
		re = /(?:Second)(?:\s{0,4})(?:Reader)(?:[:]{0,3})(.*)(?:\n)/i;
		try{
			if(re.test(dataLine))
			{
				captured = re.exec(dataLine);
				len = captured.length;
				if(trim(captured[1]))
				{
					SReader = trim(captured[1]);
					SReader = fixAuthor(SReader);
					SFlag = true;
				}
				else
				{
					if(trim(captured[0]))
					{
						console.println("Second Reader: There was a problem with the match seperation, so the ENTIRE match was used...");
						SReader += trim(captured[0]);
						SFlag = true;
					}
					else
					{
						//SReader = "ERROR";
						console.println("Unable to capture SReader...");
					}
				}
				captured = [];
				re = "";
				len = 0;
			}
			else
			{
				//SReader = "ERROR";
				console.println("Fail in matching SReader...");
			}
		}
		catch(e)
		{
			console.println("Error on line " + e.lineNumber + ": " + e);
		}
	}
/*******************************************End: Finding Second Reader*********************************************/
	
/***************** Strip New Lines ****************/	
//Get rid of all new line characters...

	//Replace newline
	dataLine = dataLine.replace(/\n/g,"");
	dataLine = dataLine.replace(/\r/g,"");
/**************************************************/	
	
/*******************************************Begin: Finding authors*********************************************/

	re = /(?:6\.?\s{0,4})(?:AUTHOR)(?:\s{0,4})(?:\(?)(?:s?)(?:\)?)(?:\s{0,4})(\D*)(?:\d\.?)/i;
	try{
		if(re.test(dataLine))
		{
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				captured[1] = trim(captured[1]);
				
				//Replacing " and " with ", "
				//fixAuthor handles commas
				captured[1] = captured[1].replace(/(,?)(\s{1,4})(and)(\s{1,4})/gi,", ");
				OLine += fixAuthor(captured[1]);
				
				//Used for searching for "(author)\n(affiliation)\n"
				affAuthor = captured[1];			
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Author: There was a problem with the match seperation, so the ENTIRE match was used...");
					OLine += trim(captured[0]);
				}
				else
				{
					OLine += "ERROR";
					console.println("Unable to capture author...");
				}
			}
			captured = [];
			re = "";
			len = 0;
		}
		else
		{
			OLine += "ERROR";
			console.println("Fail in matching author...");
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";
	
/*******************************************End: Finding authors*********************************************/

/*******************************************Begin: Finding Title*********************************************/

	re = /(?:4\.?\s{0,4})(?:TITLE)(?:\s{0,4})(?:AND?)(?:\s{0,4})(?:SUBTITLE?)(?:[:]{0,3})(?:\s{0,4})(.*)(?:6\.?\s{0,4}Author)/i;
	try
	{
		if(re.test(dataLine))
			{
			captured = [];
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				captured[1] = captured[1].replace(/(5\.?\s{0,4})(Fund)(.*)/i,"");
				OLine += trim(captured[1]);
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Title: There was a problem with the match seperation, so the ENTIRE match was used...");
					OLine += trim(captured[0]);
				}
				else
				{
					OLine += "ERROR";
					console.println("Unable to capture title...");
				}
			}
		}
		else
		{
			re = /(?:4\.?\s{0,4})(?:TITLE)(?:[:]{0,3})(?:\s{0,4})(.*)(?:6\.?\s{0,4}Author)/i;
			if(re.test(dataLine))
			{
				captured = [];
				captured = re.exec(dataLine);
				len = captured.length;
				if(trim(captured[1]))
				{
					captured[1] = captured[1].replace(/(5\.?\s{0,4})(Fund)(.*)/i,"");
					OLine += trim(captured[1]);
				}
				else
				{
					if(trim(captured[0]))
					{
						console.println("Title: There was a problem with the match seperation, so the ENTIRE match was used...");
						OLine += trim(captured[0]);
					}
					else
					{
						OLine += "ERROR";
						console.println("Unable to capture title...");
					}
				}
			}
			else
			{
				OLine += "ERROR";
				console.println("Unable to match Title...");
			}
		}
		captured = [];
		re = "";
		len = 0;
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";

/*******************************************End: Finding Title*********************************************/

/*******************************************Begin: Finding Date & Prep ISO*********************************************/
	var dateISO = "";
	re = /(?:2\.?\s{0,4})(?:REPORT)(?:\s{0,4})(?:DATE?)(?:[:]{0,3})(?:\s{0,4})(.*)(?:3\.?\s{0,4})(?:REPORT)/i;
	try
	{
		if(re.test(dataLine))
			{
			captured = [];
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				OLine += trim(captured[1]);
				dateISO = trim(captured[1]);
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Date: There was a problem with the match seperation, so the ENTIRE match was used...");
					OLine += trim(captured[0]);
				}
				else
				{
					OLine += "ERROR";
					console.println("Unable to capture Date...");
				}
			}
			captured = [];
			re = "";
			len = 0;
		}
		else
		{
			OLine += "ERROR";
			console.println("Unable to match Date...");
		}
	}
	catch(e)
	{
		console.println(fileName + ": Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";

/*******************************************End: Finding Date & Prep ISO*********************************************/

/*******************************************End: Finding ISO*********************************************/
	try
	{
		if(dateISO.length > 0)
		{
			//ISO
			re = /\d{4}/;
			var year = re.exec(dateISO);
			re = /[a-zA-Z]*/;
			var month = re.exec(dateISO);
			if(trim(month[0]))
			{
				if(month[0].search(/Jan/i) != -1)
					OLine += year + "-" + "01";
				if(month[0].search(/Feb/i) != -1)
					OLine += year + "-" + "02";
				if(month[0].search(/Mar/i) != -1)
					OLine += year + "-" + "03";
				if(month[0].search(/Apr/i) != -1)
					OLine += year + "-" + "04";
				if(month[0].search(/May/i) != -1)
					OLine += year + "-" + "05";
				if(month[0].search(/Jun/i) != -1)
					OLine += year + "-" + "06";
				if(month[0].search(/Jul/i) != -1)
					OLine += year + "-" + "07";
				if(month[0].search(/Aug/i) != -1)
					OLine += year + "-" + "08";
				if(month[0].search(/Sep/i) != -1)
					OLine += year + "-" + "09";	
				if(month[0].search(/Oct/i) != -1)
					OLine += year + "-" + "10";
				if(month[0].search(/Nov/i) != -1)
					OLine += year + "-" + "11";
				if(month[0].search(/Dec/i) != -1)
					OLine += year + "-" + "12";
				//END ISO
			}
			else
			{
				OLine += "ERROR";
				console.println("Unable to convert to ISO...");
			}
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";
	dateISO = "";
	re = "";
/*******************************************End: Finding ISO*********************************************/

/******************************************* Begin: Publisher ********************************************/
//This is going to be a blanket statement...
	OLine += "Monterey, California: Naval Postgraduate School";
	OLine += "\t";
/******************************************* End: Publisher ********************************************/

/*******************************************Begin: Finding Abstract*********************************************/
	//(?:(?:(?:\(?)(?:\s{0,2}?)(?:maximum?)(?:\s+?)(?:200?)(?:\s+?)(?:words?)(?:\s{0,2}?)(?:\)?)){0,1})
	re = /(?:13\.?\s{0,4})(?:ABSTRACT)(?:[:]{0,3})(?:\s{0,4})(.*)(?:15\.?\s{0,4})(?:NUMBER)/i;
	try
	{
		if(re.test(dataLine))
		{
			captured = [];
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				captured[1] = captured[1].replace(/(\()(\s{0,4})(maximum)(\s{0,4})(\d{3})(\s{0,4})(words)(\s{0,4})(\))(\s{0,4})/i,"");
				captured[1] = captured[1].replace(/(14)(\.?)(\s{0,4})(Subject)(\s{0,4})(Term)(.*)/i,"");
				OLine += trim(captured[1]);
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Abstract: There was a problem with the match seperation, so the ENTIRE match was used...");
					OLine += trim(captured[0]);
				}
				else
				{
					OLine += "ERROR";
					console.println("Unable to capture Abstract...");
				}
			}
			captured = [];
			re = "";
			len = 0;
		}
		else
		{
			OLine += "ERROR";
			console.println("Unable to match Abstract...");
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";

/*******************************************End: Finding Abstract*********************************************/

/*******************************************Begin: Finding Subject Terms*********************************************/

	re = /(?:14\.?\s{0,4})(?:SUBJECT)(?:\s{0,4})(?:TERMS)(?:[:]{0,3})(?:\s{0,4})(.*)(?:16.?\s{0,4})(?:PRICE)/i;
	try
	{
		if(re.test(dataLine))
		{
			captured = [];
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				captured[1] = captured[1].replace(/,/g,";");
				captured[1] = captured[1].replace(/(\s+);/g,";");
				captured[1] = captured[1].replace(/;(\s+)/g,";");
				captured[1] = captured[1].replace(/(15\.?)(\s+?)(.*)/i,"");
				OLine += trim(captured[1]);
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Subject Terms: There was a problem with the match seperation, so the ENTIRE match was used...");
					OLine += trim(captured[0]);
				}
				else
				{
					OLine += "ERROR";
					console.println("Unable to capture Subject Terms...");
				}
			}
			captured = [];
			re = "";
			len = 0;
		}
		else
		{
			OLine += "ERROR";
			console.println("Unable to match Subject Terms...");
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
	OLine += "\t";


/*******************************************End: Finding Subject Terms*********************************************/

/*******************************************Begin: Finding Degree *********************************************/
	//Degree - Grab all words until new line... OR if new line NEXT line until new line <---- Goal
	var DegreeName = "";
	var DegreeLevel = "";
	var DegreeDiscipline = "";
	var DFlag = false;
	re = /(?:for)(?:\s{0,4})(?:the)(?:\s{0,4})(?:degree)(?:\s{0,4})(?:of)(?:\s{0,4})(.*)(?:from)(?:\s{0,4})(?:the)/i;
	try{
		if(re.test(dataLine))
		{
			captured = re.exec(dataLine);
			len = captured.length;
			if(trim(captured[1]))
			{
				DFlag = true;
				DegreeName = trim(captured[1]);
				DegreeName = DegreeName.toLowerCase();
				DegreeName = toTitleCase(DegreeName);
				DegreeLevel = "other";
				re = /master/i;
				if(re.test(DegreeName))
				{
					DegreeLevel = "master's";
				}
				re = /doctor/i;
				if(re.test(DegreeName))
				{
					DegreeLevel = "doctoral";
				}
				re = /(\s+)(in)(\s+)/i;
				if(re.test(DegreeName))
				{
					DegreeDiscipline = DegreeName.substring(DegreeName.indexOf(" In ")+4,DegreeName.length);
				}
			}
			else
			{
				if(trim(captured[0]))
				{
					console.println("Degree: There was a problem with the match seperation, so the ENTIRE match was used...");
					DegreeName += trim(captured[0]);
					DFlag = true;
				}
				else
				{
					//OLine += "ERROR";
					console.println("Unable to capture Degree...");
				}
			}
			captured = [];
			re = "";
			len = 0;
		}
		else
		{
			//OLine += "ERROR";
			console.println("Fail in matching Degree...");
		}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
/*******************************************End: Finding Degree*********************************************/

/****************************************** Begin: Affiliation **********************************************/
	// Uses apDataLine
	try
	{
		if(affAuthor.length > 0)
		{
			//Seperate authors
			affAuthor = affAuthor.replace(/,/g,";");
			affAuthor = affAuthor.replace(/(\s+);/g,";");
			affAuthor = affAuthor.replace(/;(\s+)/g,";");

			if(affAuthor.search(";") != -1)
			{
				//Now we have our array of authors...
				var affAuthors = affAuthor.split(";");
				affAuthor = "";
				len = affAuthors.length;
				
				//temporary...
				var reAff = new RegExp("regex","g");
				var affFlag = false;
				var tempAuthor = [];
				var tempLen;
				
				//For each author, try to find affiliation
				for(var i = 0 ; i < len ; i++)
				{
					if(affAuthors[i].search(" ") != -1)
					{
						tempAuthor = affAuthors[i].split(" ");
						tempLen = tempAuthor.length;
						if(tempAuthor[0] && tempAuthor[tempLen-1])
						{
							reAff = new RegExp("(?:" + tempAuthor[0] + ")(?:.*)(?:" + tempAuthor[tempLen-1] + ")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
						}
						else
						{
							reAff = new RegExp("(?:" + affAuthors[i] + ")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
						}
					}
					else
					{
						reAff = new RegExp("(?:" + affAuthors[i]+")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
					}
					if(reAff.test(apDataLine))
					{
						captured = reAff.exec(apDataLine);
						if(trim(captured[1]))
						{
							//Found affiliation
							affFlag = true;
							if(affAuthor.length > 0)
							{
								affAuthor += ";";
								affAuthor += trim(captured[1]);
							}
							else
							{
								affAuthor += trim(captured[1]);
							}
						}
						else
						{
							if(trim(captured[0]))
							{
								console.println("Affilation: There was a problem with the match seperation, so the ENTIRE match was used...");
								if(affAuthor.length > 0)
								{
									affAuthor += ";";
									affAuthor += trim(captured[1]);
								}
								else
								{
									affAuthor += trim(captured[1]);
								}
								affFlag = true;
							}
							else
							{
								console.println("Unable to capture Affiliation...");
							}
						}
					}
					else
					{
						console.println("Unable to match Affiliation...");
					}
				}
			}
			else // 1 author
			{
				//temporary...
				var reAff = new RegExp("regex","g");
				var affFlag = false;
				var tempAuthor = [];
				var tempLen;
				
				if(affAuthor.search(" ") != -1)
				{
					tempAuthor = affAuthor.split(" ");
					tempLen = tempAuthor.length;
					if(tempAuthor[0] && tempAuthor[tempLen-1])
					{
						reAff = new RegExp("(?:" + tempAuthor[0] + ")(?:.*)(?:" + tempAuthor[tempLen-1] + ")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
					}
					else
					{
						reAff = new RegExp("(?:" + affAuthor + ")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
					}
				}
				else
				{
					reAff = new RegExp("(?:" + affAuthor + ")(?:\\s+?)(?:\\n)(.*)(?:\\n)","i");
				}
				affAuthor = "";
				
				if(reAff.test(apDataLine))
				{
					captured = reAff.exec(apDataLine);
					if(trim(captured[1]))
					{
						//Found affiliation
						affFlag = true;
						if(affAuthor.length > 0)
						{
							affAuthor += ";";
							affAuthor += trim(captured[1]);
						}
						else
						{
							affAuthor += trim(captured[1]);
						}
					}
					else
					{
						if(trim(captured[0]))
						{
							console.println("Affilation: There was a problem with the match seperation, so the ENTIRE match was used...");
							if(affAuthor.length > 0)
							{
								affAuthor += ";";
								affAuthor += trim(captured[1]);
							}
							else
							{
								affAuthor += trim(captured[1]);
							}
							affFlag = true;
						}
						else
						{
							console.println("Unable to capture Affiliation...");
						}
					}
				}
				else
				{
					console.println("Unable to match Affiliation...");
				}
			}
		}
		//else
		//{
			//OLine += "ERROR";
		//}
	}
	catch(e)
	{
		console.println("Error on line " + e.lineNumber + ": " + e);
	}
/****************************************** End: Affiliation **********************************************/

/******** Advisor ******/
//The search was up above, because it had to be prior to removal of new line chars
	if(AFlag == true)
	{
		OLine += fixAuthor(Advisor);
	}
	else
	{
		OLine += "ERROR";
	}
	OLine += "\t";
/***********************/

/******** SReader ******/
//The search was up above, because it had to be prior to removal of new line chars
	if(SFlag == true)
	{
		OLine += fixAuthor(SReader);
	}
	else
	{
		OLine += "ERROR";
	}
	OLine += "\t";
/***********************/

/********* Degree Stuff *********/
	if(DFlag == true)
	{
		OLine += trim(DegreeName) + "\t";
		OLine += trim(DegreeLevel) + "\t";
		OLine += trim(DegreeDiscipline) + "\t";
	}
	else
	{
		OLine += "ERROR\tERROR\tERROR\t";
	}
/********************************/

/******** Affiliation *************/
	if(affFlag == true)
	{
		OLine += affAuthor;
	}
	else
	{
		OLine += "ERROR";
	}
/***********************************/

	console.println("********* End of Errors for " + fileName + "**********");
	OLine += "\r\n";

        // Get the data object contents as a file stream 
       var oFile = global.myContainer.getDataObjectContents("mySummary.xls"); 
	   
	   // Convert the stream to a string
       var cFile = util.stringFromStream(oFile, "utf-8");
        
       //Concatenate the new lines.
        cFile += OLine;
        OLine = "";
     
       //Convert back to a file stream
        oFile = util.streamFromString( cFile, "utf-8" ); 
      
       //and update the file attachment
        global.myContainer.setDataObjectContents({cName: "mySummary.xls", oStream: oFile });  
}  
    else
    {console.println("Error: " + this.documentFileName + " has less than 7 pages");}
} 
catch(e) 
{
    console.println("Error on line " + e.lineNumber + ": " + e); 
	delete typeof global.startBatch;
    event.rc = false; // abort batch
}
/* PHII_UkrMTS.js*/
/* function ParseInvoice(Excel,filename)
*	Excel: reference to an Excel application object
*	filename: path of the invoice file
*
*	Loops through the invoice via FindInvoiceSection() and extracts all of the call data.
*
*	Returns a phone call data structure.
*/

function ParseInvoice(EXL,filename){
	var	i = 1,
		pnum,
		stime,
		svname,
		InvoiceSheet,
		isheet,
		//IVDATA = {"name":"","headers":[],"numbers":[],"data":[]},
		dhs = [],
		dns = [],
		data = [],
		hcell,
		hpos,
		curpos=0;

	//Initialization
	stime = +new Date();
	try{	InvoiceSheet = EXL.Workbooks.Open(filename);
		isheet = InvoiceSheet.WorkSheets(1);
	}catch(err){
		alert("Error! " + err.message);
		return 0;
	}
		//Start Actual Parsing
	function FindInvoiceSection(){ //can we improve efficiency here?
		var	tnum = "",
			ncell;
		do{
			if(hpos<i){return "";} //end of data
			i = hpos+50; //headers are at least that long; just save some time.

			if(curpos<hpos){ //we need to find new call data
				curpos = parseInt(			
					isheet.Columns(7).Find(":",			//look for start of call data; we don't just loop over every row and check for
					isheet.Cells(i,7)).Address.split("$")[2],10);	//new headers at the same time because that would miss the case where the last
				if(curpos<i){return "";} //end of data			//number in the invoice has no data. So, just in case, we check that first.
			} //otherwise, the cursor is either already at the data for this header,
			//or this header has no data, and we need to get the next one

			ncell = isheet.Columns(1).Find("Контракт №",isheet.Cells(i,1)); //now get the next header
			hpos = parseInt(ncell.Address.split("$")[2],10);

			//if the next header occurs after the next call block, we get the number
			//for the last header and set i at the beginning of that call block
			if(hpos>curpos){
				tnum = (/: (\d+)/).exec(hcell.Text)[1]; // get the number for that Section
				i=curpos;
			}//otherwise, tnum remains "" and the loop repeats
			hcell = ncell; //cash the next header location for the next time through
		} while(tnum===""); // if a new header showed up, we search again
		return tnum;
	}

	function hasNumber(arr,sval){ //we can't just use "in" or "indexOf" because of prefixes
		var	entry,
			loop,
			elen,
			slen = sval.length;
		for(loop=arr.length;loop--;){
			entry = arr[loop];
			elen = entry.length;
			if(	//take advantage of short-circuiting
				(entry == sval) || //and test the hopefully most common cases first
				((slen > elen && entry == sval.substr(slen-elen)) ||
				(elen > slen && entry.substr(elen-slen) == sval))
			){return true;} //take advantage of short-circuiting
		}
		return false;
	}

	function parBlock(line,hnum){
		for(var cnum="X";cnum!=="";line++){
			var recline = isheet.Range("C"+line+":H"+line).Cells;
			cnum = recline(1,3).Text;
			if(hasNumber(dhs,cnum)){continue;} //skip duplicates
			//switch(recline(1,1).Text.substring(0,2)){
			switch(recline(1,1).Text){
				/*date/time,duration,outgoing,incoming*/
				//case "Ви":
				case "Вихідні дзвінки ":
				case "Вихідні повідом.":
					data.push([+new Date(recline(1,4).Text+" "+recline(1,5).Text),
						+new Date("70/01/01 "+recline(1,6).Text),
						hnum,cnum]);
					break;
				//case "Вх":
				case "Вхідні дзвінки  ":
				case "Вхідні повідом. ":
					data.push([+new Date(recline(1,4).Text+" "+recline(1,5).Text),
						+new Date("70/01/01 "+recline(1,6).Text),
						cnum,hnum]);
					break;
				default: continue;
			}
			if(!(isNaN(cnum) || hasNumber(dns,cnum))){dns.push(cnum);}
		}
	}

	window.moveTo(screen.width+10, screen.height+10);	//The window has a tendency to freeze, so we want it out of the way
	try{
		var blockLocs = [];
		WShell.Popup("Parsing Headers.",1,"Invoice Importer",64);

		hcell = isheet.Columns(1).Find("Контракт №",isheet.Cells(i,1)); //prime the pump
		hpos = parseInt(hcell.Address.split("$")[2],10);
		pnum = FindInvoiceSection();

		while(pnum!==""){ //queue up all of the headers in the invoice
			//WShell.Popup(i+":"+pnum,1,"Invoice Importer",64);
			dhs.push(pnum);
			dns.push(pnum);
			blockLocs.push(i);
			pnum = FindInvoiceSection();
		}
		WShell.Popup("Found "+blockLocs.length+" data blocks.",1,"Invoice Importer",64);
		for(var w=0;w<blockLocs.length;){
			parBlock(blockLocs[w],dhs[w]);
			w++;
			if(!(w%3)){WShell.Popup("Completed %"+Math.floor(100*w/blockLocs.length),1,"Invoice Importer",64);}
		}
	}catch(err){ //clean-up and tell about the problem if there's an error in one of the records
		InvoiceSheet.Close(false);
		window.moveTo(200,200);
		alert("Error encountered in processing on line " + i + ": " + err.message);
		return 0;
	}
	var d = (/\d+\.\d+\.\d+/).exec(isheet.Cells(2,1).Text)[0].split(".");
	var m = parseInt(d[1].substr(0,1)=="0"?d[1].substr(1,1):d[1],10)-1;
	svname = d[2]+["01","02","03","04","05","06","07","08","09","10","11","12"][m]+".cdat";
	InvoiceSheet.Close(false);

	dhs.sort();
	dns.sort();
	data.sort(/*function(a,b){return a[0]-b[0];}*/);

	window.moveTo(200,200);
	statbox.value+="\nExecution took "+((+new Date() - stime) / 1000)+" seconds\n";

	return {"name":svname,"headers":dhs,"numbers":dns,"data":data};
}
Array.prototype.removeAt = function(val){
	for(var i=val;i<this.length-1;i++) this[i] = this[i+1];
	if(val<this.length) this.pop();
}

String.prototype.trim = function(){return this.replace(/^\s\s*/, '').replace(/\s\s*$/, '');};

function ParseInvoice(){
	var Dir1, Dir2, etime, i, tdate;	//Initialization

	var svname = savename.value;
	var pnums = number.value.split(",");
	var pagecurs = [];
	var pages = [];
	var stime = new Date();

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var log = fso.OpenTextFile("logfile.txt", 2, true, -1);

	var path = fso.GetAbsolutePathName(".");
	var spath = path.substring(0,path.lastIndexOf("\\")) + "\\Phone Records";
	if(!fso.FolderExists(spath)) fso.CreateFolder(spath);
	spath += "\\";
	path += "\\";
	Excel = new ActiveXObject("Excel.Application");
	Excel.Visible = 0;
	tdate = parseInt(tdbox.value);
	try{
		ListSheet = Excel.Workbooks.Open(plist1.value);
		Dir1 = MakeDirectory(ListSheet.WorkSheets("CList")); //This is the complete directory for the first part of the month
		ListSheet.Close(false);
		if(tdate != 32){
			ListSheet = Excel.Workbooks.Open(plist2.value);
			Dir2 = MakeDirectory(ListSheet.WorkSheets("Clist")); //This is the directory of changes post-transfer
			ListSheet.Close(false);
		}else{Dir2 = null;}
		InvoiceSheet = Excel.Workbooks.Open(invf.value);
		GraphSheet = Excel.Workbooks.Open(path + "sourcefiles\\pgm.template");
	}catch(err){
		Excel.Quit();
		alert("Error In Initialization! " + err.message);
		return 0;
	}			//automatic name assignment if no output file name was typed.
	if(!svname){
		var d = (/\d+\.\d+\.\d+/).exec(InvoiceSheet.Worksheets(1).Cells(2,1).Text)[0].split(".");
		var m = parseInt(d[1].substr(0,1)=="0"?d[1].substr(1,1):d[1])-1;
		svname = "Tracking" + ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sept","Oct","Nov","Dec"][m]+d[2];
	}
	GraphSheet.SaveAs(spath + svname);

	//Start Actual Parsing
        i=1;
	try{
		var ivdata = InvoiceSheet.WorkSheets(1);
		for(var j=0;j<pnums.length;j++){ //clean up the search list
			pnums[j] = pnums[j].trim();
			var testcell = ivdata.Columns(5).Find(pnums[j],ivdata.Cells(1,5));
			if(!testcell){
				log.WriteLine(pnums[j] + " does not occur in the invoice.");
				pnums.removeAt(j);
				j--;
			}else{
				//log.WriteLine(pnums[j] + " first occurs at line " + testcell.Address.split("$")[2]);
				pagecurs.push(3);
				pages.push(CopyPage(GraphSheet,pnums[j]));
			}
		}

		if(pnums.length==0){
			log.WriteLine("No numbers found.");
			alert("No matching numbers found.");
			GraphSheet.Close(false);
			Excel.Quit();
			return;
		}
		
		/*
		I. Get the current header / missionary phone number
		II. Loop
			1. Get the position of the next header
			2. Loop
				a. Look at first phone number
				b. find instances of that number up until the end of the section and copy them into that number's page
			3. Set next header as the current header.
		*/
		
		var hcell = ivdata.Columns(1).Find("Контракт №",ivdata.Cells(i,1)); //initial stuff
		var nextpos = parseInt(hcell.Address.split("$")[2]);
		var nextnum = (/: (\d+)/).exec(hcell.Text)[1]; // number for that Section

		while(nextnum){
			var curpos = nextpos;
			var curnum = nextnum;
			i=curpos+1;
			hcell = ivdata.Columns(1).Find("Контракт №",ivdata.Cells(i,1)); //initial stuff
			nextpos = parseInt(hcell.Address.split("$")[2]);
			nextnum = (nextpos<i)?"":(/: (\d+)/).exec(hcell.Text)[1]; // number for that Section
			
			if(Dir1[curnum] || (Dir2 && Dir2[curnum])){
				log.WriteLine("Current number: " + curnum);
				for(var j=0;j<pnums.length;j++){
					log.WriteLine("Checking for "+pnums[j]);
					var recline = parseInt(ivdata.Columns(5).Find(pnums[j],ivdata.Cells(i,5)).Address.split("$")[2]);
					var latch = recline;
					if(recline<i || recline>=nextpos) continue;
					do{
						var curdir = ((new Date(ivdata.Cells(recline,6).Text)).getDate()>tdate)?Dir2:Dir1;
						log.WriteLine("recline: " + recline + " date: " + ivdata.Cells(recline,6).Text+ " name: " + curdir[curnum]);

						if(CopyRecord(recline,ivdata,pagecurs[j],curnum,pages[j],curdir)) pagecurs[j]++;
						recline = parseInt(ivdata.Columns(5).Find(pnums[j],ivdata.Cells(recline,5)).Address.split("$")[2]);
					}while(recline < nextpos && recline>=latch);
				}
			}
		}
	}catch(err){ //clean-up and tell about the problem if there's an error in one of the records
		log.close();
		InvoiceSheet.Close(false);
		GraphSheet.Save(); //Backup!
		GraphSheet.Close(false);
		Excel.Quit();
		window.moveTo(200, 200);
		alert("Error! " + err.message);
		return;
	}
	log.close();
	etime = ((new Date()).getTime() - stime.getTime()) / 1000;
	InvoiceSheet.Close(false);
	GraphSheet.Worksheets("1").Cells.Value = "";
	GraphSheet.Worksheets("1").Delete();
	GraphSheet.Save();
	GraphSheet.Close(false);
	Excel.Quit();
	window.moveTo(200,200);
	alert("Invoice Processing Completed\nGraphs Saved as "+spath+svname);
	alert("Execution took " + etime + " seconds");
}

function CopyRecord(icursor,ivdata,ccursor,cnum,cpage,dir){
	if(!dir[cnum]) return false;
	cpage.Cells(ccursor,2).Value = dir[cnum];
	switch(ivdata.Cells(icursor, 3).Text.substring(0,2)){
			case "Ви":	cpage.Cells(ccursor,1).Value = "Outgoing";
					break;
			case "Вх":	cpage.Cells(ccursor,1).Value = "Incoming";
					break;
			default:	cpage.Cells(ccursor,1).Value = "Service";
	}
	cpage.Range("C" + ccursor + ":E" + ccursor).Value = ivdata.Range("F" + icursor + ":H" + icursor).Value;
	return true;
}

function MakeDirectory(dsheet){
	var inc, area, dir, phnum, dname;
	dir = [];
	inc = 4;
	while(dsheet.Cells(inc,1).Text != ""){
		area = dsheet.Cells(inc, 2).Text;
		if(area != "Extra" ){
			m1 = dsheet.Cells(inc,3).Text;
			m2 = dsheet.Cells(inc,4).Text.substr(0,12);
			if(m1 == "Closed") m1 = "";
			if(m2 == ""){dname = (m1=="")?dsheet.Cells(inc, 2).Text.substr(0,31):m1.substr(0,30);}
			else{
				var namelen;
				m1 = m1.substr(0,12);
				namelen = m1.length + m2.length + 2;
				dname = area.substr(0, 31 - namelen) + " " + m1 + " " + m2;
			}
			phnum = dsheet.Cells(inc, 1).Text;
			phnum = phnum.substring(phnum.length-9);
			dir[phnum] = dname;
		}
		inc++;
	}
	return dir;
}		

function CopyPage(csheet,number){
	csheet.Sheets("1").Copy(null,csheet.Worksheets(csheet.Worksheets.Count));
	csheet.ActiveSheet.name = number;
	return csheet.ActiveSheet;
}
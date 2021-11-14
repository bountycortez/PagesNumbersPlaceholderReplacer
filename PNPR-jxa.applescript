
(() => {
//function run(argv){
	'use strict';
	
	//version
	var version="V1.0 20211114 by BF"
 
	//startup
	var app = Application.currentApplication();
	app.includeStandardAdditions = true;

	console.log('Starting PagesNumbersPlaceholderReplacer PNPR '+version+'...' + app.currentDate());
	//debugger;

	//================================================================================
	//variables
	var voice = 'Alex';
	var voicespeed = 200;
	var voicepitch = 30;
	var voicemodulation = 100;
	var alerttext;
	var alertmessage;
	var appname = app.name;
	var sysevents =  Application('System Events');
	var pages = Application('Pages');
	var numbers = Application('Numbers');

	console.log(appname);


	//================================================================================
	//bring app to front
	app.activate();
	app.say("Please check!", {
 	    using: voice,
    	speakingRate: voicespeed,
    	pitch: voicepitch,
    	modulation: voicemodulation
	});
	alerttext = 'WELCOME TO\n PagesNumbersPlaceholderReplacer\n(PNPR '+version+')\n\nARE YOU READY?\nPLEASE CHECK THE FOLLOWING:';
	var txt1='Numbers document (e.g. PNPR.numbers) with first row as header containing values according to Pages placeholder tags\nmust be open now\nand datarows (optional after filtering)\nmust be seleted now.';
	var txt2='Pages document made from template with Placeholders (e.g. PNPR.template)\nmust be open now.';
	var txt3='If more than one row is selected, additional Pages documents based on template of open document will be created!';
	alertmessage = '\n1. '+txt1+'\n\n2. '+txt2+'\n\n3. '+txt3+'\n';
	app.displayAlert(alerttext, {
					message: alertmessage,
					as: 'critical',
					buttons: ['Dont Continue', 'Continue'],
    				defaultButton: 'Continue',
    				cancelButton: 'Dont Continue'
	});
	app.say("Let's get ready to rumble!", {
 	    using: voice,
    	speakingRate: voicespeed,
    	pitch: voicepitch,
    	modulation: voicemodulation
	});

	
	//================================================================================
	//make numbers and pages visible
	pages.activate();
	delay(1);
	pages.window.visible = true;

	numbers.activate();
	delay(1);
	numbers.window.visible = true;

	//bring to frontmost
	try {
		sysevents.processes['Pages'].frontmost = true;
	} catch(err) {
		const e = new Error('\nPages not running, run aborted!', { cause: err }); e.errorNumber = -600; 
		throw e;
	}

	try {
		sysevents.processes['Numbers'].frontmost = true;
	} catch(err) {
		const e = new Error('\nPages not running, run aborted!', { cause: err }); e.errorNumber = -600; 
		throw e;
	}

	//================================================================================
	//aktuelle dokumente	
	try {
		var numbersdoc=numbers.documents[0];
		var numbersfile= numbersdoc.file();
		console.log('Numbers Doc File: ' + numbersfile);		
	} catch(err){
		const e = new Error('\nCannot determine Numbers document filename, run aborted!', { cause: err }); e.errorNumber = -38; 
		throw e;
	}

	try {
		var pagesdoc=pages.documents[0];
		var pagesfile=pagesdoc.file();
		console.log('Pages Doc File: ' + pagesfile);
		var pagesdocname=pagesdoc.name();
		console.log('Pages Doc Name: ' + pagesdocname);
	} catch(err){
		const e = new Error('\nCannot determine Pages document filename, run aborted!', { cause: err }); e.errorNumber = -38; 
		throw e;
	}

	try {
		var pagesdoctemplate=pagesdoc.documentTemplate();
		var pagestemplatename=pagesdoc.documentTemplate().name();
		console.log('Pages Document Template Name: ' + pagestemplatename);

	} catch(err){
		const e = new Error('\nCannot determine Pages document template, run aborted!', { cause: err }); e.errorNumber = -38; 
		throw e;
	}

	try {
		var template=pages.templates[pagestemplatename];
		console.log('Pages Template Name: ' + template.name());
	} catch(err){
		const e = new Error('\nCannot determine Pages template <'+ pagestemplatename+'>, run aborted!', { cause: err }); e.errorNumber = -38; 
		throw e;
	}

	try {
		var placeholders=pagesdoc.placeholderTexts;
		var placeholdercount=placeholders.length;
		console.log('Pages Placeholder Count: '+placeholdercount);

		//make placeholder texts unique
		//var placeholdertags = placeholders().filter((v,i,a)=>a.indexOf(v)==i);
		var placeholdertags=[...new Set(placeholders())]; //most elegant solution
		var placeholdertaglength= placeholdertags.length;
		console.log('Pages Placeholder Tags ('+placeholdertaglength+'): '+placeholdertags);	
	} catch(err) {
		const e = new Error('\nCannot determine Pages document placeholder count, run aborted!', { cause: err }); e.errorNumber = -38; 
		throw e;
	}

	//================================================================================
	//numbers document with headers and selection?
	const sheet=numbers.documents[0].activeSheet;
	const table=sheet.tables[0];
	const sheetname=sheet.name();
	const tablename=table.name();
	const headerrowcount=table.headerRowCount();
	console.log('Numbers Sheet Name: <' + sheetname + '>');
	console.log('Numbers Table Name: <' + tablename + '>');
	console.log('Numbers Header Row Count: <' + headerrowcount + '>');
	
	if(headerrowcount!=1){
		const e = new Error('\nNumbers Sheet <' + sheetname +'> Table <' + tablename + '> header row count must be 1!\n\nRun aborted!', { cause: null }); e.errorNumber = -192; 
		throw e;
	}
	
	//table can be filtered
	//if(!table.filtered()){
	//	const e = new Error('\nNumbers Sheet <' + sheetname +'> Table <' + tablename + '> not filtered!\n\nRun aborted!'); e.errorNumber = -192; 
	//	throw e;
	//}
	if(table.filtered()){
		console.log('TABLE IS FILTERED!');
	}else{
		console.log('TABLE IS NOT FILTERED!');
	}
	
	//get selection
	try {
		var selectionrange = table.selectionRange;
		var selectionrows = selectionrange.rows;
		var selectioncount = selectionrows.length;
		
	} catch(err) {
	    //console.log('NO SELECTION!');
		const e = new Error('\nGetting Numbers Selection failed!\n\nRun aborted!', { cause: err }); e.errorNumber = -192; 
		throw e;
	}

	//get headers
	try {
		var cellrange=table.cellRange;
		var rows=cellrange.rows;
		var rowcount=rows.length;
		var headercells=rows[0].cells();
		var headerheight=rows[0].height();
		var headerlength=headercells.length
		var zuersetzen;
		var value;
		var anzahl;
		var headertags=[];
		var numberstags=[];
			
		console.log('TABLE ROW COUNT: ' + rowcount);
	    console.log('HEADER ROW HEIGHT: ' + headerheight);
	    console.log('HEADER COLUMN COUNT: ' + headerlength);
		console.log('SELECTION ROW COUNT: ' + selectioncount);

		
		//show headers
		for(let i=0;i<headerlength;i++) {
			value=headercells[i]().formattedValue();
			console.log('HEADER VALUE[' + i +']: ' + value);
			headertags.push(value);
			
			zuersetzen=placeholders.whose({tag: value });
			anzahl=zuersetzen.length;
			if(anzahl==0){
				console.log('+++ WARNING: Keinen Placeholder <'+value+'> gefunden!');
			}else{
				console.log('Necessary Replacements for Placeholder <'+value+'>: '+anzahl);
				numberstags.push(value);
			}
		}		
	} catch(err) {
		const e = new Error('\nCannot determine Numbers document header, run aborted!', { cause: err }); e.errorNumber = -192; 
		throw e;
	}
		
	if(selectioncount==0){
		const e = new Error('\nNo rows selected, run aborted!', { cause: null }); e.errorNumber = -192; 
		throw e;
	}


	console.log('=============================================================');
	console.log('HEADER  TAGS: '+headertags);
	console.log('NUMBERS TAGS: '+numberstags);
	
	var selectioncells= selectionrows[0].cells();
	var selectionlength=selectioncells.length;
	//selection length must be equal to headertags length!
	//check here!!!!!!!
	if(headertags.length!=selectionlength){
		const e = new Error('\nNot enough cells selected (must be '+selectionlength+' per selected row), run aborted!', { cause: null }); e.errorNumber = -192; 
		throw e;
	}
	console.log('SELECTION LENGTH OK (taken from selection row 0): ' + selectionlength);
	console.log('=============================================================');

	
	//================================================================================
	//get numbers data rows and replace placeholders
	var newdoc=pagesdoc;
	var toreplace;
	var tag;
	var restanzahl;
	var index=-1;
	
	//loop over every data row
	for(let i=0;i<selectioncount;i++) {

		try {
									
			selectioncells = selectionrows[i].cells();
			
			//selection length for this row
			selectionlength = selectioncells.length;
			//selection count must be equal to headertags length!
			//check here!!!!!!!
			if(headertags.length!=selectionlength){
				const e = new Error('\nNot enough cells selected (must be '+selectionlength+' in selected row '+i+'), run aborted!', { cause: null }); e.errorNumber = -192; 
				throw e;
			}
	    	console.log('DATAROW[' + i + '] length: '+selectionlength+', height:' + selectionrows[i].height());	

			//loop over numbers tags
			for(let j=0;j<numberstags.length;j++) {
				tag=numberstags[j];
				index=headertags.indexOf(tag);
				if(index<0){
					console.log('No Numbers header tag <'+tag+'> found!');
					continue;
				}else{
					console.log('Index for Numbers header tag <'+tag+'>: '+index);
				}
				toreplace=newdoc.placeholderTexts.whose({tag:tag});
				if(index<selectionlength){
					value=selectioncells[index]().formattedValue();
					console.log('DATAROW[' + i + ']: Value for tag <'+tag+'> is <'+value+'>');
				}else{
					console.log('DATAROW[' + i + ']: No value for tag <'+tag+'>, using empty value...');
					value='';
				}
				let anzahl=toreplace.length;
				console.log('Necessary Replacements for tag <'+tag+'>: '+anzahl);
				for(let k=0;k<anzahl;k++) {
					toreplace[0].set(value); //0 weil es reduziert sich ja jedesmal
				}
			}
			
			//remaining placeholders in this document
			var remainingplaceholders=newdoc.placeholderTexts;
			var remainingcount=remainingplaceholders.length;
			console.log('Remaining Pages Placeholder Count: '+remainingcount);
			for(let k=0;k<remainingcount;k++) {
				console.log('Remaining Tag '+k+': '+remainingplaceholders()[k]);
			}
			
			//================================================================================
			//rename doc (not allowed!)
			//newdoc.name='PNR-NR'+i;

			//================================================================================
			if(i<selectioncount-1){
				//make new document from template	
				newdoc = pages.Document({
					documentTemplate: template
				});
				newdoc.make();
				pages.activate();
				delay(1);
				pages.window.visible = true;
			}

		} catch(err) {
			const e = new Error('\nCannot determine Numbers document datarow #' + i +', run aborted!', { cause: err }); e.errorNumber = -192; 
			throw e;
		}
	}


	//================================================================================
	//================================================================================
	//================================================================================
	//end of run
    console.log('+++ END OF CODE +++');
	app.say("Finished! Check if result is as expected!", {
 	    using: voice,
    	speakingRate: voicespeed,
    	pitch: voicepitch,
    	modulation: voicemodulation
	});
	alerttext = '\n\nFINISHED!\nCreated '+ selectioncount+' Pages documents\n\nCheck if everything is as expected!\n\n(This message becomes closed after 5 seconds.)';
	alertmessage = '----------';
	app.displayDialog(alerttext, {
					message: alertmessage,
					as: 'informational',
					buttons: ['OK'],
    				defaultButton: 'OK',
					givingUpAfter: 5,
					withIcon: 'stop'
	});
    console.log('+++ BYE! +++');
	return 0;
		
})()


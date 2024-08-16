Office.onReady(()=>{ // Office is ready.
	startup();
});

var context:Excel.RequestContext;
var metaDataSheet:worksheetWrapper|undefined=undefined;
var sheetData:{client:string,template:string,version:string,sheetId:string}|undefined;

var myConsole:{console:HTMLElement|null,log:(text:string)=>void,clear:()=>void} = {
	"console":null,
	"log":(text:string)=>{
		console.log(text);
		if(console==null)return;
		myConsole.console!.innerText+=text+"\n";
	},
	"clear":()=>{
		if(myConsole.console==null)return;
		myConsole.console!.innerText="";
	}
}
var inputElement:HTMLElement;
function addHtmlButton(buttonText:string,func:(()=>Promise<void>)) {
	const button = document.createElement("button"); button.innerText = buttonText;
	button.style.marginTop = "1em"; button.style.marginLeft = "1em"; button.id = buttonText;
	button.onclick = func; inputElement.appendChild(button);
	inputElement.appendChild(document.createElement("br"));
}
async function addMetaDataHTMLTextInput(name:string,startValue:string,onChanged:((value:string)=>Promise<void>)):Promise<void> {
	var data:{[key:string]:any} = await readMetadataObject();
	var value = startValue;
	if (data[name]!=undefined) { value=data[name]; } else {
		await clearMetadataObject();
		data[name]=value;
		await setMetadataObject(data);
	}
	const inputLabel:HTMLLabelElement = document.createElement("label");inputElement.appendChild(inputLabel);
	inputLabel.style.marginTop = "1em"; inputLabel.style.marginLeft = "1em";
	inputLabel.id="metaData"+name+"Label"; inputLabel.htmlFor="metaData"+name; inputLabel.innerText=name;

	const input = document.createElement("input"); input.type="text"; input.value = value;
	input.style.marginTop = "1em"; input.style.marginLeft = "1em"; input.id = name;
	inputElement.appendChild(input);
	input.onchange = async () => {
		var data:{[key:string]:any} = await readMetadataObject();
		data[name]=input.value;
		await clearMetadataObject();
		await setMetadataObject(data);
		await onChanged(input.value);
	};
	inputElement.appendChild(document.createElement("br"));
}

async function getMetaDataSheetIsNull():Promise<boolean> {
	if (metaDataSheet==undefined) metaDataSheet = new worksheetWrapper(context);
	metaDataSheet.getWorksheet("MetaData");
	return metaDataSheet.isNullObject.asyncGet();
}
async function getOrMakeMetadataSheet():Promise<void> {
	// if the sheet is null, make it, otherwise the function just returns
	if (await getMetaDataSheetIsNull()) {
		await metaDataSheet!.addWorksheet("MetaData");
		metaDataSheet!.getRange("$A$1").values.set([["num Settings"]]).wrapText.set(true).columnWidth.set((13*11+7)/2);
		metaDataSheet!.getRange("$B$1").values.set([[0]]).wrapText.set(true).columnWidth.set(300);
		metaDataSheet!.hide();
		await context.sync();
	}
}
async function readMetadataObject():Promise<{[key:string]:any}> {
	// if the metadataSheet is not found obviously return an empty object
	if (await getMetaDataSheetIsNull()) return {};
	var data:{[key:string]:any} = {};
	var numRows=(await metaDataSheet!.getRange("$B$1").values.asyncGet())[0][0];
	if (numRows>0) data=Object.fromEntries(await metaDataSheet!.getRange("$A$2:$B$"+(numRows+1)).values.asyncGet())
	return data;
}
async function clearMetadataObject():Promise<void> {
	if (await getMetaDataSheetIsNull()) return;
	var numRows=(await metaDataSheet!.getRange("$B$1").values.asyncGet())[0][0];
	if (numRows>0) metaDataSheet!.getRange("$A$2:$B$"+numRows).clear();
	metaDataSheet!.getRange("$A$1").values.set([["num Settings"]]).wrapText.set(true).columnWidth.set((13*11+7)/2);
	metaDataSheet!.getRange("$B$1").values.set([[0]]).wrapText.set(true).columnWidth.set(300);
	await context.sync();
}
async function setMetadataObject(data:{[key:string]:any}):Promise<void> {
	await getOrMakeMetadataSheet();
	// set data
	var entries:[string,any][] = Object.entries(data);
	var longestStrLen:number=12;// length of "num Settings" + 1
	for (let i = 0; i < entries.length; i++) {
		if (entries[i][0].length>longestStrLen)longestStrLen=entries[i][0].length;
	}
	metaDataSheet!.getRange("$A$2:$B$"+(entries.length+1)).values.set(entries);
	metaDataSheet!.getRange("$A$1").values.set([["num Settings"]]).wrapText.set(true).columnWidth.set(((longestStrLen+1)*11+7)/2);
	metaDataSheet!.getRange("$B$1").values.set([[entries.length]]).wrapText.set(true).columnWidth.set(300);
	metaDataSheet!.hide();
	await context.sync();
}

async function setMetadataValue(key:string,value:any) {
	if (key=="num Settings" || key=="Client" || key=="Template" || key=="Version" || key=="Sheet Id") return;
	await getOrMakeMetadataSheet();
	var data:{[key:string]:any} = await readMetadataObject();
	await clearMetadataObject();
	data[key]=value;
	await setMetadataObject(data);
}
async function getMetadataValue(key:string):Promise<any|undefined> {
	if (await getMetaDataSheetIsNull()) return undefined;
	var data:{[key:string]:any} = await readMetadataObject();
	return data[key];
}

const DefaultValues:{client:string,template:string,version:string} = {client:"DEV",template:"Testing",version:"1.0"};
async function startup() {
	//console.clear();
	inputElement=document.getElementById("input")!;// get input div
	myConsole.console = document.getElementById("console");// get console div for output
	context = await RequestContextAsync();

	var data:{[key:string]:any} = await readMetadataObject();
	const hasMetadataSheet:boolean = (!await getMetaDataSheetIsNull());
	if (hasMetadataSheet) {
		if ((await metaDataSheet!.getRange("$A$1").values.asyncGet())[0][0]!="num Settings") {
			myConsole.log("Wrong version of Template handler, register new template and copy in data."); return;
		}
	}
	if (data["Client"]!=undefined && data["Template"]!=undefined && data["Version"]!=undefined && data["Sheet Id"]!=undefined) {
		myConsole.log("Metadata found");
		sheetData={
			client:(data["Client"] as string),
			template:data["Template"] as string,
			version:(data["Version"] as string).split("\"")[1],
			sheetId:data["Sheet Id"] as string
		};
		validateAndLoad();
	} else {
		myConsole.log("No MetaData found.");
		// prevent registering a workbook that has anything other than Sheet1 and a metadatasheet
		const worksheets:worksheetWrapper[] = await worksheetWrapper.getWorksheets(context);
		if (hasMetadataSheet) {
			// the "num settings" metadata setting must be 0 ("num settings" is always the first settings and therefor the value is in B1)
			var numSettings=(await metaDataSheet!.getRange("$B$1").values.asyncGet())[0][0];
			if (numSettings!=0) { myConsole.log("You can only register a template on an empty workbook."); return; }
			// there can anly be 2 sheets if on of them is the metadata sheet
			if (worksheets.length!=2) { myConsole.log("You can only register a template on an empty workbook."); return; }
			// the names of the 2 sheets must be either "Sheet1" and then "MetaData" or "MetaData" and then "Sheet1"
			const name1:string = await worksheets[0].name.asyncGet();
			const name2:string = await worksheets[1].name.asyncGet();
			if (!((name1=="Sheet1") && (name2=="MetaData")) && !((name1=="MetaData") && (name2=="Sheet1"))) { myConsole.log("You can only register a template on an empty workbook."); return; }
		} else {
			// there can only be one sheet, and it must be named "Sheet1"
			if (worksheets.length!=1) { myConsole.log("You can only register a template on an empty workbook."); return; }
			if ((await worksheets[0].name.asyncGet())!="Sheet1") { myConsole.log("You can only register a template on an empty workbook."); return; }
		}
		
		// client input
		const clientLabel:HTMLLabelElement = document.createElement("label");inputElement.appendChild(clientLabel);
		clientLabel.id="metaDataLabel"; clientLabel.htmlFor="metaDataClient"; clientLabel.innerText="Client";
		const client:HTMLInputElement = document.createElement("input");inputElement.appendChild(client);
		client.id="metaDataClient"; client.name="client"; client.value=DefaultValues.client;
		inputElement.appendChild(document.createElement("br"));
		// template name input
		const templates = await fetchJson("/templates",{});
		console.log(templates);
		const templateLabel:HTMLLabelElement = document.createElement("label");inputElement.appendChild(templateLabel);
		templateLabel.id="metaDataLabel"; templateLabel.htmlFor="metaDataTemplate"; templateLabel.innerText="Template";
		const template:HTMLSelectElement = document.createElement("select");inputElement.appendChild(template);
		template.id="metaDataTemplate"; template.name="template";
		for (let i = 0; i < templates.length; i++) {
			const option = document.createElement("option");template.appendChild(option);
			option.value=option.innerText=templates[i];
		}
		template.value=DefaultValues.template;
		inputElement.appendChild(document.createElement("br"));
		// template version input
		const versionLabel:HTMLLabelElement = document.createElement("label");inputElement.appendChild(versionLabel);
		versionLabel.id="metaDataLabel"; versionLabel.htmlFor="metaDataVersion"; versionLabel.innerText="Version";
		const version:HTMLInputElement = document.createElement("input");inputElement.appendChild(version);
		version.id="metaDataVersion"; version.name="version"; version.value=DefaultValues.version;
		inputElement.appendChild(document.createElement("br"));
		// register button
		const button:HTMLButtonElement = document.createElement("button");inputElement.appendChild(button);
		button.id="addMetaData"; button.innerText="Register Sheet"; button.onclick=register;
	}
}
async function register() {
	//register sheet with server
	myConsole.log("Registering worksheet.");
	var client:string = (document.getElementById("metaDataClient")! as HTMLInputElement).value;
	var template:string = (document.getElementById("metaDataTemplate")! as HTMLInputElement).value;
	var version:string = (document.getElementById("metaDataVersion")! as HTMLInputElement).value;
	var out:any = await fetchJsonPost("/register",{template,client,version},{});
	// if server returned empty object or null, say it failed and return
	if (out==null||Object.keys(out).length==0) { myConsole.log("Registration failed.");return; }
	const sheetId:string=out.sheetId; myConsole.log("Registration succeeded.");
	// set values in metadata sheet
	sheetData={client, template, version, sheetId};
	await setMetadataObject({
		"Client":client,
		"Template":template,
		"Version":"\""+version+"\"",
		"Sheet Id":sheetId
	});
	// clear out html
	inputElement.innerHTML="";
	// continue
	validateAndLoad();
}
async function validateAndLoad() {
	if (sheetData==undefined) return;
	// send validation post request
	var validation:any = await fetchJsonPost("/validate",sheetData,{});
	if (!validation.status) {myConsole.log("Validation failed.");return;}
	myConsole.log("Validation succeeded.");

	// if successfull, set taskpane to auto open
	Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument",true);
	await Office.context.document.settings.saveAsync();

	// load script for template with credentials
	let script = document.createElement('script');
	script.src = "/template/"+sheetData!.template+"/"+sheetData!.version+".js?client="+sheetData!.client+"&sheetId="+sheetData?.sheetId;
	script.type="text/javascript";
	document.head.appendChild(script);
	await new Promise<void>((resolve:()=>void)=>{script.onload=resolve;});// wait for the script to load
	myConsole.log("Script loaded.");
	// after script has been loaded, create the instance of the template class
	var template:templateInterface = generateTemplate(context,myConsole);
}
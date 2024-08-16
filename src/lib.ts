function generateUUID():string {
	var a = (new Date()).getTime();//Timestamp
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
		var b = Math.random() * 16;//random number between 0 and 16
		b = (a + b)%16 | 0;
		a = Math.floor(a/16);
		return (c === 'x' ? b : (b & 0x3 | 0x8)).toString(16);
	});
}
async function fetchJson(url:string,headers:{[keys:string]:any}):Promise<any> {
    var tmpHeaders = headers||{};
    tmpHeaders["Content-Type"] = "application/json";
    return new Promise<string>((resolve:(value:string)=>void) => {
        fetch(url,{
            method:"GET",
            mode:"no-cors",
            headers:tmpHeaders
        }).then((response:Response)=>response.text())
        .then((text:string)=>resolve(JSON.parse(text)));
    });
};
async function fetchJsonPost(url:string,data:any,headers:{[keys:string]:any}):Promise<any> {
    var tmpHeaders = headers||{};
    tmpHeaders["Content-Type"] = "application/json";
    return new Promise<string>((resolve:(value:string)=>void) => {
        fetch(url,{
            method:"POST",
            mode:"no-cors",
            headers:tmpHeaders,
            body: JSON.stringify(data)
        }).then((response:Response)=>response.text())
        .then((text:string)=>resolve(JSON.parse(text)));
    });
};
function dateStringToJsDate(dateString:string):Date {
	const dateStringSplit:string[] = dateString.split("/");
	const date:Date = new Date();
	date.setUTCMonth(Number(dateStringSplit[0])-1);
	date.setUTCDate(Number(dateStringSplit[1]));
	date.setUTCFullYear(Number(dateStringSplit[2]));
	return date;
}
function jsDateToUtcDateString(date:Date):string {
	return date.getUTCMonth()+1+"/"+date.getUTCDate()+"/"+date.getUTCFullYear();
}


function jsDateToExcelDate(jsDate:Date):number {
	const excelEpoch:Date = dateStringToJsDate("12/30/1899");// Excel's epoch date
	const millisecondsPerDay:number = 24 * 60 * 60 * 1000; // Number of milliseconds in a day
	return  Math.floor((jsDate.getTime() - excelEpoch.getTime()) / millisecondsPerDay);
}
function dateStringToExcelDate(dateString:string):number {
	return jsDateToExcelDate(dateStringToJsDate(dateString));
}
function ExcelDateToJsDate(excelDate:number):Date {
	const excelEpoch:Date = dateStringToJsDate("12/30/1899");// Excel's epoch date
	const millisecondsPerDay:number = 24 * 60 * 60 * 1000; // Number of milliseconds in a day
	return new Date(excelDate*millisecondsPerDay+excelEpoch.getTime());
}
function excelDateToUtcDateString(excelDate:number):string {
	return jsDateToUtcDateString(ExcelDateToJsDate(excelDate));
}
function EOMONTH(excelDate:number):number {
	const dateString:string = excelDateToUtcDateString(excelDate);
	const [Month,Day,Year]:number[] = dateString.split("/").map((val:string):number=>{return Number(val);});
	return dateStringToExcelDate(Month%12+1+"/1/"+(Year+Math.floor(Month/12)))-1;
}
function WORKDAY(excelDate:number, days:number, holidays:number[]):number {
	var finalDate:number = excelDate;
	if (days>=0) {
		for (let i = 0; i < days; i++) {
			finalDate++;
			while (true) {
				const adjustedDateWeekdate:number = ExcelDateToJsDate(finalDate).getUTCDay();
				if ((adjustedDateWeekdate==0) || (adjustedDateWeekdate==6) || holidays.includes(finalDate)) {
					finalDate++;
				}
				else break;
			}
		}
		return finalDate;
	} else {
		for (let i = 0; i < days; i++) {
			finalDate--;
			while (true) {
				const adjustedDateWeekdate:number = ExcelDateToJsDate(finalDate).getUTCDay();
				if ((adjustedDateWeekdate==0) || (adjustedDateWeekdate==6) || holidays.includes(finalDate))
					finalDate--;
				else break;
			}
		}
		return finalDate;
	}
}
class rangeWrapper {
	context:Excel.RequestContext;
	range:Excel.Range|undefined;
	worksheet:worksheetWrapper;
	constructor(context:Excel.RequestContext) { this.context=context; this.worksheet=new worksheetWrapper(this.context);}
	
	getSelectedRange():rangeWrapper {
		this.range = this.context.workbook.getSelectedRange();
		this.worksheet = worksheetWrapper.Wrap(this.range.worksheet);
		return this;
	}
	getWorksheetRange(worksheet:Excel.Worksheet,address:string):rangeWrapper {
		this.range = worksheet.getRange(address);
		this.worksheet = worksheetWrapper.Wrap(this.range.worksheet);
		return this;
	}
	
	clear():rangeWrapper {
		if (this.range==null) return this;
		this.range.clear();
		return this;
	}
	select():rangeWrapper {
		if (this.range != null) this.range.select();
		return this;
	}
	merge():rangeWrapper {
		if (this.range != null) this.range.merge();
		return this;
	}
	lock():rangeWrapper {
		if (this.range != null) this.range.format.protection.locked=true;
		return this;
	}
	unlock():rangeWrapper {
		if (this.range != null) this.range.format.protection.locked=false;
		return this;
	}
	
	hideFormulas():rangeWrapper {
		if (this.range != null) this.range.format.protection.formulaHidden=true;
		return this;
	}
	unhideFormulas():rangeWrapper {
		if (this.range != null) this.range.format.protection.formulaHidden=false;
		return this;
	}
	setBorder(index:"EdgeBottom"|"EdgeTop"|"EdgeLeft"|"EdgeRight"|"InsideVertical"|"InsideHorizontal"|"DiagonalDown"|"DiagonalUp",color:string,weight:Excel.BorderWeight|"Hairline"|"Thin"|"Medium"|"Thick"="Thin", style:Excel.BorderLineStyle|"None"|"Continuous"|"Dash"|"DashDot"|"DashDotDot"|"Dot"|"Double"|"SlantDashDot"="Continuous"):rangeWrapper {
		if (this.range==null) return this;
		const border:Excel.RangeBorder = this.range.format.borders.getItem(index);
		border.color=color;
		border.weight=weight;
		border.style=style;
		return this;
	}
	setBorderEdges(color:string,weight:Excel.BorderWeight|"Hairline"|"Thin"|"Medium"|"Thick"):rangeWrapper {
		const list:("EdgeBottom"|"EdgeTop"|"EdgeLeft"|"EdgeRight"|"InsideVertical"|"InsideHorizontal")[]=["EdgeBottom","EdgeTop","EdgeLeft","EdgeRight","InsideVertical","InsideHorizontal"];
		for (let i = 0; i < list.length; i++) { this.setBorder(list[i],color,weight) }
		return this;
	}
	setBorderBox(color:string,weight:Excel.BorderWeight|"Hairline"|"Thin"|"Medium"|"Thick"):rangeWrapper {
		const list:("EdgeBottom"|"EdgeTop"|"EdgeLeft"|"EdgeRight"|"InsideVertical"|"InsideHorizontal")[]=["EdgeBottom","EdgeTop","EdgeLeft","EdgeRight"];
		for (let i = 0; i < list.length; i++) { this.setBorder(list[i],color,weight) }
		return this;
	}
	setFontSize(size:number):rangeWrapper {
		if (this.range==null) return this;
		this.range!.format.font.size=size;
		return this;
	}
	setFontColor(color:string):rangeWrapper {
		if (this.range==null) return this;
		this.range!.format.font.color=color;
		return this;
	}
	sort(fields:Excel.SortField[],matchCase?:boolean|undefined,hasHeaders?:boolean|undefined,orientation?:Excel.SortOrientation|undefined,method?:Excel.SortMethod|undefined):rangeWrapper {
		if (this.range==null) return this;
		this.range.sort.apply(fields,matchCase,hasHeaders,orientation,method);
		return this;
	}
	
	values:{set:(values:any[][])=>rangeWrapper,load:()=>rangeWrapper,get:()=>any[][],asyncGet:()=>Promise<any[][]>} = {
		"set":(values:any[][]):rangeWrapper=>{
			if (this.range!=null) this.range.values=values;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("values");
			return this;
		},
		"get":()=>{
			if (this.range==null) return [];
			else return this.range.values;
		},
		"asyncGet":async()=>{
			this.values.load();
			await this.context.sync();
			return this.values.get();
		}
	}
	visibleValues:{asyncGet:()=>Promise<any[][]>} = {
		"asyncGet":async ():Promise<any[][]>=>{
			if (this.range==null) return [];
			const rngVw:Excel.RangeView = this.range!.getVisibleView();
			rngVw.load("values");
			await this.context.sync();
			return rngVw.values;
		}
	}
	text:{load:()=>rangeWrapper,get:()=>string[][],asyncGet:()=>Promise<string[][]>} = {
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("text");
			return this;
		},
		"get":()=>{
			if (this.range==null) return [];
			else return this.range.text;
		},
		"asyncGet":async()=>{
			this.text.load();
			await this.context.sync();
			return this.text.get();
		}
	}
	visibleText:{asyncGet:()=>Promise<any[][]>} = {
		"asyncGet":async ():Promise<any[][]>=>{
			if (this.range==null) return [];
			const rngVw:Excel.RangeView = this.range!.getVisibleView();
			rngVw.load("text");
			await this.context.sync();
			return rngVw.text;
		}
	}
	formulas:{load:()=>rangeWrapper,get:()=>string[][],asyncGet:()=>Promise<string[][]>} = {
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("formulas");
			return this;
		},
		"get":()=>{
			if (this.range==null) return [];
			else return this.range.formulas;
		},
		"asyncGet":async()=>{
			this.formulas.load();
			await this.context.sync();
			return this.formulas.get();
		}
	}
	numberFormat:{set:(numberFormat:string[][])=>rangeWrapper,load:()=>rangeWrapper,get:()=>string[][],asyncGet:()=>Promise<string[][]>} = {
		"set":(numberFormat:string[][]):rangeWrapper=>{
			if (this.range!=null) this.range.numberFormat=numberFormat;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("numberFormat");
			return this;
		},
		"get":()=>{
			if (this.range==null) return [];
			else return this.range.numberFormat;
		},
		"asyncGet":async()=>{
			this.numberFormat.load();
			await this.context.sync();
			return this.numberFormat.get();
		}
	}
	columnWidth:{set:(width:number)=>rangeWrapper,load:()=>rangeWrapper,get:()=>number,asyncGet:()=>Promise<number>} = {
		"set":(width:number):rangeWrapper=>{
			if (this.range!=null) this.range.format.columnWidth=width;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/columnWidth");
			return this;
		},
		"get":()=>{
			if (this.range==null) return 0;
			else return this.range.format.columnWidth;
		},
		"asyncGet":async()=>{
			this.columnWidth.load();
			await this.context.sync();
			return this.columnWidth.get();
		}
	}
	rowHeight:{set:(height:number)=>rangeWrapper,load:()=>rangeWrapper,get:()=>number,asyncGet:()=>Promise<number>} = {
		"set":(height:number):rangeWrapper=>{
			if (this.range!=null) this.range.format.rowHeight=height;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/rowHeight");
			return this;
		},
		"get":()=>{
			if (this.range==null) return 0;
			else return this.range.format.rowHeight;
		},
		"asyncGet":async()=>{
			this.rowHeight.load();
			await this.context.sync();
			return this.rowHeight.get();
		}
	}
	/** HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange") */
	fill:{set:(color:string)=>rangeWrapper,load:()=>rangeWrapper,get:()=>string,clear:()=>rangeWrapper,asyncGet:()=>Promise<string>} = {
		"set":(color:string):rangeWrapper=>{
			if (this.range!=null) this.range.format.fill.color=color;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/fill/color");
			return this;
		},
		"get":()=>{
			if (this.range==null) return "000000";
			else return this.range.format.fill.color;
		},
		"clear":()=>{
			if (this.range!=null) this.range.format.fill.clear();
			return this;
		},
		"asyncGet":async()=>{
			this.fill.load();
			await this.context.sync();
			return this.fill.get();
		}
	}
	bold:{set:(bold:boolean)=>rangeWrapper,load:()=>rangeWrapper,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"set":(bold:boolean):rangeWrapper=>{
			if (this.range!=null) this.range.format.font.bold=bold;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/font/bold");
			return this;
		},
		"get":()=>{
			if (this.range==null) return false;
			else return this.range.format.font.bold;
		},
		"asyncGet":async()=>{
			this.bold.load();
			await this.context.sync();
			return this.bold.get();
		}
	}
	italic:{set:(italic:boolean)=>rangeWrapper,load:()=>rangeWrapper,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"set":(italic:boolean):rangeWrapper=>{
			if (this.range!=null) this.range.format.font.italic=italic;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/font/italic");
			return this;
		},
		"get":()=>{
			if (this.range==null) return false;
			else return this.range.format.font.italic;
		},
		"asyncGet":async()=>{
			this.italic.load();
			await this.context.sync();
			return this.italic.get();
		}
	}
	wrapText:{set:(wrapText:boolean)=>rangeWrapper,load:()=>rangeWrapper,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"set":(wrapText:boolean):rangeWrapper=>{
			if (this.range!=null) this.range.format.wrapText=wrapText;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/wrapText");
			return this;
		},
		"get":()=>{
			if (this.range==null) return false;
			else return this.range.format.wrapText;
		},
		"asyncGet":async()=>{
			this.wrapText.load();
			await this.context.sync();
			return this.wrapText.get();
		}
	}
	verticalAlignment:{set:(verticalAlignment:"Center"|"Justify"|"Distributed"|"Top"|"Bottom")=>rangeWrapper,load:()=>rangeWrapper,get:()=>("Center"|"Justify"|"Distributed"|"Top"|"Bottom"),asyncGet:()=>Promise<"Center"|"Justify"|"Distributed"|"Top"|"Bottom">} = {
		"set":(verticalAlignment:"Center"|"Justify"|"Distributed"|"Top"|"Bottom"):rangeWrapper=>{
			if (this.range!=null) this.range.format.verticalAlignment=verticalAlignment;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/verticalAlignment");
			return this;
		},
		"get":()=>{
			if (this.range==null) return "Justify";
			else return this.range.format.verticalAlignment;
		},
		"asyncGet":async()=>{
			this.verticalAlignment.load();
			await this.context.sync();
			return this.verticalAlignment.get();
		}
	}
	horizontalAlignment:{set:(horizontalAlignment:"Center"|"Justify"|"Distributed"|"General"|"Left"|"Right"|"Fill"|"CenterAcrossSelection")=>rangeWrapper,load:()=>rangeWrapper,get:()=>("Center"|"Justify"|"Distributed"|"General"|"Left"|"Right"|"Fill"|"CenterAcrossSelection"),asyncGet:()=>Promise<"Center"|"Justify"|"Distributed"|"General"|"Left"|"Right"|"Fill"|"CenterAcrossSelection">} = {
		"set":(horizontalAlignment:"Center"|"Justify"|"Distributed"|"General"|"Left"|"Right"|"Fill"|"CenterAcrossSelection"):rangeWrapper=>{
			if (this.range!=null) this.range.format.horizontalAlignment=horizontalAlignment;
			return this;
		},
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("format/horizontalAlignment");
			return this;
		},
		"get":()=>{
			if (this.range==null) return "General";
			else return this.range.format.horizontalAlignment;
		},
		"asyncGet":async()=>{
			this.horizontalAlignment.load();
			await this.context.sync();
			return this.horizontalAlignment.get();
		}
	}
	address:{load:()=>rangeWrapper,get:()=>string,asyncGet:()=>Promise<string>} = {
		"load":():rangeWrapper=>{
			if (this.range!=null) this.range.load("address");
			return this;
		},
		"get":()=>{
			if (this.address==null) return "";
			else return this.range!.address;
		},
		"asyncGet":async()=>{
			this.address.load();
			await this.context.sync();
			return this.address.get();
		}
	}

	track():rangeWrapper {
		if (this.range!=null) this.context.trackedObjects.add(this.range!);
		return this;
	}
	untrack():rangeWrapper {
		if (this.range!=null) this.context.trackedObjects.remove(this.range!);
		return this;
	}

	static Wrap(range:Excel.Range):rangeWrapper {
		const wrapped:rangeWrapper = new rangeWrapper(range.context);
		wrapped.range = range; return wrapped;
	}
}
class worksheetWrapper {
	context:Excel.RequestContext;
	worksheet:Excel.Worksheet|undefined;
	constructor(context:Excel.RequestContext) { this.context=context; }
	
	async addWorksheet(name:string):Promise<worksheetWrapper> {
		var worksheets:Excel.WorksheetCollection = this.context.workbook.worksheets;
		this.worksheet=worksheets.add(name);
		await this.context.sync();
		return this;
	}
	getWorksheet(name:string):worksheetWrapper {
		var worksheets:Excel.WorksheetCollection = this.context.workbook.worksheets;
		this.worksheet=worksheets.getItemOrNullObject(name);
		return this;
	}
	getActiveWorksheet():worksheetWrapper {
		this.worksheet=this.context.workbook.worksheets.getActiveWorksheet();
		return this;
	}
	get usedRange():rangeWrapper {
		if (this.worksheet==null) return (new rangeWrapper(this.context));
		return rangeWrapper.Wrap(this.worksheet.getUsedRange());
	}
	getUsedRange():rangeWrapper {
		if (this.worksheet==null) return (new rangeWrapper(this.context));
		var range:rangeWrapper = new rangeWrapper(this.context);
		range.range=this.worksheet.getUsedRange();
		return range;
	}
	getRange(address:string):rangeWrapper {
		if (this.worksheet==null) return (new rangeWrapper(this.context));
		return (new rangeWrapper(this.context)).getWorksheetRange(this.worksheet,address)
	}
	activate():worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.activate();
		return this;
	}
	protect(options?: Excel.WorksheetProtectionOptions | undefined, password?: string | undefined):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.protection.protect(options,password);
		return this;
	}
	unprotect(password?: string | undefined):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.protection.unprotect(password);
		return this;
	}
	hide():worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.visibility="Hidden";
		return this;
	}
	unhide():worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.visibility="Visible";
		return this;
	}
	setCustomProperty(name:string,value:string):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.customProperties.add(name,value);
		return this;
	}
	addNamedRange(name:string,reference:rangeWrapper):worksheetWrapper {
		if (this.worksheet==null) return this;
		if (reference.range==null) return this;
		this.worksheet.names.add(name,reference.range);
		return this;
	}
	deleteNamedRange(name:string):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.names.getItem(name).delete();
		return this;
	}
	freezeRange(range:rangeWrapper):worksheetWrapper {
		if (this.worksheet==null) return this;
		if (range.range==null) return this;
		this.worksheet.freezePanes.freezeAt(range.range);
		return this;
	}
	freezeRows(num:number):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.freezePanes.freezeRows(num);
		return this;
	}
	freezeColumns(num:number):worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.freezePanes.freezeColumns(num);
		return this;
	}
	unFreeze():worksheetWrapper {
		if (this.worksheet==null) return this;
		this.worksheet.freezePanes.unfreeze();
		return this;
	}
	
	name:{set:(name:string)=>worksheetWrapper,load:()=>worksheetWrapper,get:()=>string,asyncGet:()=>Promise<string>} = {
		"set":(name:string):worksheetWrapper=>{
			if (this.worksheet!=null) this.worksheet!.name=name;
			return this;
		},
		"load":()=>{
			if (this.worksheet!=null) this.worksheet.load("name");
			return this;
		},
		"get":()=>{
			if (this.worksheet==null) return "";
			else return this.worksheet.name;
		},
		"asyncGet":async()=>{
			this.name.load();
			await this.context.sync();
			return this.name.get();
		}
	}
	position:{load:()=>worksheetWrapper,get:()=>number,asyncGet:()=>Promise<number>} = {
		"load":():worksheetWrapper=>{
			if (this.worksheet!=null) this.worksheet.load("position");
			return this;
		},
		"get":()=>{
			if (this.worksheet==null) return -1;
			else return this.worksheet.position;
		},
		"asyncGet":async()=>{
			this.position.load();
			await this.context.sync();
			return this.position.get();
		}
	}
	isNullObject:{load:()=>void,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"load":()=>{
			this.worksheet!.load("isNullObject")
		},
		"get":()=>{
			if (this.worksheet==null) return true;
			else return this.worksheet.isNullObject;
		},
		"asyncGet":async()=>{
			this.isNullObject.load();
			await this.context.sync();
			return this.isNullObject.get();
		}
	}
	tabColor:{set:(tabColor:string)=>worksheetWrapper,load:()=>worksheetWrapper,get:()=>string,asyncGet:()=>Promise<string>} = {
		"set":(tabColor:string):worksheetWrapper=>{
			if (this.worksheet!=null) this.worksheet.tabColor=tabColor;
			return this;
		},
		"load":():worksheetWrapper=>{
			if (this.worksheet!=null) this.worksheet.load("tabColor");
			return this;
		},
		"get":()=>{
			if (this.worksheet==null) return "";
			else return this.worksheet.tabColor;
		},
		"asyncGet":async()=>{
			this.tabColor.load();
			await this.context.sync();
			return this.tabColor.get();
		}
	}

	track():worksheetWrapper {
		if (this.worksheet!=null) this.context.trackedObjects.add(this.worksheet!);
		return this;
	}
	untrack():worksheetWrapper {
		if (this.worksheet!=null) this.context.trackedObjects.remove(this.worksheet!);
		return this;
	}

	static Wrap(worksheet:Excel.Worksheet):worksheetWrapper {
		const wrapped:worksheetWrapper = new worksheetWrapper(worksheet.context);
		wrapped.worksheet = worksheet; return wrapped;
	}
	static async getWorksheets(context:Excel.RequestContext):Promise<worksheetWrapper[]> {
		var worksheets:Excel.WorksheetCollection = context.workbook.worksheets;
		worksheets.load("items");
		await context.sync();
		return worksheets.items.map((el:Excel.Worksheet)=>worksheetWrapper.Wrap(el));
	}
}
class tableWrapper {
	context:Excel.RequestContext;
	worksheet:worksheetWrapper;
	table:Excel.Table|undefined;
	constructor(worksheet:worksheetWrapper) {
		this.worksheet=worksheet;
		this.context=this.worksheet.context;
	}
	async addTable(name:string, range:string, hasHeaders:boolean):Promise<tableWrapper> {
		var tables:Excel.TableCollection = this.worksheet.worksheet!.tables;
		this.table=tables.add(range,hasHeaders);
		this.table.name=name;
		await this.context.sync();
		return this;
	}
	getTable(name:string):tableWrapper {
		var tables:Excel.TableCollection = this.worksheet.worksheet!.tables;
		this.table = tables.getItemOrNullObject(name);
		return this;
	}
	getRange():rangeWrapper {
		var range:rangeWrapper = new rangeWrapper(this.context);
		if (this.table!=null) range.range=this.table!.getRange();
		return range;
	}
	getDataBodyRange():rangeWrapper {
		if (this.table==null) return new rangeWrapper(this.context);
		else return rangeWrapper.Wrap(this.table!.getDataBodyRange());
	}
	convertToRange():rangeWrapper {
		var range:rangeWrapper = new rangeWrapper(this.context);
		if (this.table!=null) range.range=this.table!.convertToRange();
		return range;
	}
	setColumnWidths(widths:number[]):tableWrapper {
		if (this.table==null) return this;
		for (let i = 0; i < widths.length; i++) {
			this.columns.getColumn(i).columnWidth.set(widths[i]);
		}
		return this;
	}
	async resize(test:rangeWrapper|string):Promise<tableWrapper> {
		this.table!.resize((typeof test == "string")?test:(await test.address.asyncGet()));
		await this.context.sync();
		return this;
	}
	style:{set:(style:string)=>tableWrapper,load:()=>tableWrapper,get:()=>string,asyncGet:()=>Promise<string>} = {
		"set":(style:string):tableWrapper=>{
			if (this.table!=null) this.table!.style=style;
			return this;
		},
		"load":():tableWrapper=>{
			if (this.table!=null) this.table!.load("style");
			return this;
		},
		"get":():string=>{
			if (this.table==null) return "";
			return this.table!.style;
		},
		"asyncGet":async():Promise<string>=>{
			if (this.table==null) return "";
			this.style.load();
			await this.context.sync();
			return this.style.get();
		}
	}
	showTotals:{set:(name:boolean)=>tableWrapper,load:()=>tableWrapper,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"set":(value:boolean):tableWrapper=>{
			if (this.table!=null) this.table!.showTotals=value;
			return this;
		},
		"load":():tableWrapper=>{
			if (this.table!=null) this.table!.load("showTotals");
			return this;
		},
		"get":():boolean=>{
			if (this.table==null) return false;
			return this.table!.showTotals;
		},
		"asyncGet":async():Promise<boolean>=>{
			if (this.table==null) return false;
			this.showTotals.load();
			await this.context.sync();
			return this.showTotals.get();
		}
	}
	name:{set:(name:string)=>tableWrapper,load:()=>tableWrapper,get:()=>string,asyncGet:()=>Promise<string>} = {
		"set":(name:string):tableWrapper=>{
			if (this.table!=null) this.table!.name=name;
			return this;
		},
		"load":()=>{
			if (this.table!=null) this.table.load("name");
			return this;
		},
		"get":()=>{
			if (this.table==null) return "";
			return this.table.name;
		},
		"asyncGet":async()=>{
			if (this.table==null) return "";
			this.name.load();
			await this.context.sync();
			return this.name.get();
		}
	}
	headers:{ set:(values:any[][])=>tableWrapper,getRange:()=>rangeWrapper,asyncGet:()=>Promise<any[][]> } = {
		"set":(values:any[][]):tableWrapper=>{
			if (this.table!=null) this.table.getHeaderRowRange().values=values;
			return this;
		},
		"getRange":():rangeWrapper=>{
			var range:rangeWrapper = new rangeWrapper(this.context);
			if (this.table!=null) range.range=this.table!.getHeaderRowRange();
			return range;
		},
		"asyncGet":async():Promise<any[][]>=>{
			if (this.table==null) return [];
			var range:Excel.Range = this.table.getHeaderRowRange().load("values");
			await this.context.sync();
			return range.values;
		}
	}
	body:{ set:(values:any[][])=>tableWrapper,getRange:()=>rangeWrapper,asyncGet:()=>Promise<any[][]> } = {
		"set":(values:any[][]):tableWrapper=>{
			if (this.table!=null) this.table.getDataBodyRange().values=values;
			return this;
		},
		"getRange":():rangeWrapper=>{
			var range:rangeWrapper = new rangeWrapper(this.context);
			if (this.table!=null) range.range=this.table!.getDataBodyRange();
			return range;
		},
		"asyncGet":async():Promise<any[][]>=>{
			if (this.table==null) return [];
			var range:Excel.Range = this.table.getDataBodyRange().load("values");
			await this.context.sync();
			return range.values;
		}
	}
	totals:{ set:(values:any[][])=>tableWrapper,getRange:()=>rangeWrapper,asyncGet:()=>Promise<any[][]> } = {
		"set":(values:any[][]):tableWrapper=>{
			if (this.table!=null) this.table.getTotalRowRange().values=values;
			return this;
		},
		"getRange":():rangeWrapper=>{
			var range:rangeWrapper = new rangeWrapper(this.context);
			if (this.table!=null) range.range=this.table!.getTotalRowRange();
			return range;
		},
		"asyncGet":async():Promise<any[][]>=>{
			if (this.table==null) return [];
			var range:Excel.Range = this.table.getTotalRowRange().load("values");
			await this.context.sync();
			return range.values;
		}
	}
	rows:{ setRow:(index:number,values:any[])=>tableWrapper,asyncGet:()=>Promise<Excel.TableRow[]>,getRow:(index:number)=>rangeWrapper} = {
		"setRow":(index:number,values:any[]):tableWrapper=>{
			if (this.table!=null) this.table!.getDataBodyRange().getRow(index).values=values;
			return this;
		},
		"asyncGet":async():Promise<Excel.TableRow[]>=>{
			if (this.table==null) return [];
			var tableRows:Excel.TableRowCollection = this.table.rows.load("items");
			await this.context.sync();
			return tableRows.items;
		},
		"getRow":(index:number):rangeWrapper=>{
			var range:rangeWrapper = new rangeWrapper(this.context);
			if (this.table!=null) range.range=this.table!.getDataBodyRange().getRow(index);
			return range;
		}
	}
	columns:{ setColumn:(index:number,values:any[])=>tableWrapper,asyncGet:()=>Promise<Excel.TableRow[]>,getColumn:(index:number)=>rangeWrapper} = {
		"setColumn":(index:number,values:any[]):tableWrapper=>{
			if (this.table!=null) this.table!.getDataBodyRange().getColumn(index).values=values;
			return this;
		},
		"asyncGet":async():Promise<Excel.TableRow[]>=>{
			if (this.table==null) return [];
			var tableColumns:Excel.TableColumnCollection = this.table.columns.load("items");
			await this.context.sync();
			return tableColumns.items;
		},
		"getColumn":(index:number):rangeWrapper=>{
			var range:rangeWrapper = new rangeWrapper(this.context);
			if (this.table!=null) range.range=this.table!.getDataBodyRange().getColumn(index);
			return range;
		}
	}
	isNullObject:{load:()=>void,get:()=>boolean,asyncGet:()=>Promise<boolean>} = {
		"load":()=>{
			this.table!.load("isNullObject")
		},
		"get":()=>{
			if (this.table==null) return true;
			else return this.table.isNullObject;
		},
		"asyncGet":async()=>{
			this.isNullObject.load();
			await this.context.sync();
			return this.isNullObject.get();
		}
	}
	values:{asyncGet:()=>Promise<any[][]>} = {
		"asyncGet":async ():Promise<any[][]>=>{
			if (this.table==null) return [];
			const rng:Excel.Range = this.table!.getDataBodyRange();
			rng.load("values");
			await this.context.sync();
			return rng.values;
		}
	}
	visibleValues:{asyncGet:()=>Promise<any[][]>} = {
		"asyncGet":async ():Promise<any[][]>=>{
			if (this.table==null) return [];
			const rngVw:Excel.RangeView = this.table!.getDataBodyRange().getVisibleView();
			rngVw.load("values");
			await this.context.sync();
			return rngVw.values;
		}
	}
	text:{asyncGet:()=>Promise<any[][]>} = {
		"asyncGet":async ():Promise<any[][]>=>{
			if (this.table==null) return [];
			const rng:Excel.Range = this.table!.getDataBodyRange();
			rng.load("text");
			await this.context.sync();
			return rng.text;
		}
	}
	visibleText:{asyncGet:()=>Promise<string[][]>} = {
		"asyncGet":async ():Promise<string[][]>=>{
			if (this.table==null) return [];
			const rngVw:Excel.RangeView = this.table!.getRange().getVisibleView()
			rngVw.load("text");
			await this.context.sync();
			var tmp:string[][] = rngVw.text;
            tmp.splice(0,1);//remove first row
			return tmp;
		}
	}
	track():tableWrapper {
		if (this.table!=null) this.context.trackedObjects.add(this.table!);
		return this;
	}
	untrack():tableWrapper {
		if (this.table!=null) this.context.trackedObjects.remove(this.table!);
		return this;
	}
}

async function RequestContextAsync():Promise<Excel.RequestContext> {
	return new Promise<Excel.RequestContext>((resolve:(value:Excel.RequestContext)=>void)=>{
		Excel.run(async (context:Excel.RequestContext) => {
			resolve(context);
		});
	});
}
const DefaultNumberFormat:string = "0";
const PercentageNumberFormat:string = "0%";
const AccountNumberFormat:string = "_(#,##0.00_);_((#,##0.00);_(\"-\"??_);_(@_)";
const AccountMoneyNumberFormat:string = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
const DateNumberFormat:string = "m/d/yyyy";
const MonthDayNumberFormat:string = "m/d";
const WeekdayNumberFormat:string = "dddd";

type myConsoleType = {console:HTMLElement|null,log:(text:string)=>void,clear:()=>void};

var columnsAlphebet: string[] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", 
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", 
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", 
    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ"];
//#region settings
interface TableSheetSettings {
    headerOverrideA1:boolean;
    doLockA1:boolean;
    headers:string[][];
    headersFontSize:number;
    numBufferLines:number;
}
interface DataSheetSettings {
    headerOverrideA1:boolean;
    headers:string[][];
    headersFontSize:number;
}
interface GuidanceSheetSettings {
    values:string[][];
    bold:boolean[][];
    fontSize:number[][];
}
interface TableSheetColumnSettings {
    isInputColumn:boolean;
    name:string;
    numberFormat:string;
    columnWidth:number;
    alignment:"Center" | "Justify" | "Distributed" | "General" | "Left" | "Right" | "Fill" | "CenterAcrossSelection";
    bgColor?:string;
    wrapText?:boolean;
    hasTotal?:boolean;
    totalType?:"CntA"|"Cnt"|"Sum"|"Avg"|"Custom";
    totalCustomValue?:string;
}
interface DataSheetColumnSettings {
    name:string;
    columnWidth:number;
    alignment:"Center" | "Justify" | "Distributed" | "General" | "Left" | "Right" | "Fill" | "CenterAcrossSelection";
}
//#endregion settings
interface TableSheetColumnData {
    isInputColumn:boolean;
    name:string;
    numberFormat:string;
    columnWidth:number;
    alignment:"Center" | "Justify" | "Distributed" | "General" | "Left" | "Right" | "Fill" | "CenterAcrossSelection";
    bgColor?:string;
    wrapText?:boolean;
    hasTotal?:boolean;
    totalType?:"CntA"|"Cnt"|"Sum"|"Avg"|"Custom";
    totalCustomValue?:string;

    letter:string;
    sum:number;
    count:number;
    countN:number;
}
interface DataSheetColumnData {
    name:string;
    columnWidth:number;
    alignment:"Center" | "Justify" | "Distributed" | "General" | "Left" | "Right" | "Fill" | "CenterAcrossSelection";
    letter:string;
}

class TableSheetColumnGroup {
    public sheetHandler:TableSheetHandler;
    public columns:number[];
    public isDirty:boolean;
    public hasInit:boolean;
    
    private process:()=>Promise<any[][]>;
    constructor(_sheetHandler:TableSheetHandler,_columns:number[],_process:()=>Promise<any[][]>) {
        this.sheetHandler=_sheetHandler;
        this.columns=_columns;
        this.process=_process;
        this.isDirty=true;
        this.hasInit=false;
    }
    static loopDetector:number[]=[];
    setDirty():void {
        if (this.isDirty) return;
        this.isDirty=true;
        this.sheetHandler.worksheet.tabColor.set("#CCCC00");
        this.sheetHandler.unprotect()
        for (let i = 0; i < this.columns.length; i++)
            this.sheetHandler.worksheet.getRange("$"+this.sheetHandler.columns[this.columns[i]].letter+"$"+(1+this.sheetHandler.settings.headers.length)).fill.set("#CCCC00");
        this.sheetHandler.protect()
        // mark any dependent on a column within this group as dirty as well
        for (const column of this.columns) {
            const dependents:TableSheetColumnGroup[] = this.sheetHandler.columnDependents[column];
            for (const dependent of dependents) {
                dependent.setDirty();
            }
        }
    }
    async init():Promise<void> {
        // make sure TableSheetHandler is aware that we are getting the values, so it can keep track of dependecies
        TableSheetHandler.currColumnGroup=this;
        const values:any[][] = await this.process();
        TableSheetHandler.currColumnGroup=undefined;
        this.sheetHandler.setColumns(this.columns,values);
        this.isDirty=false;
        this.hasInit=true;
        this.sheetHandler.worksheet.tabColor.set("");
        this.sheetHandler.unprotect()
        for (let i = 0; i < this.columns.length; i++)
            this.sheetHandler.worksheet.getRange("$"+this.sheetHandler.columns[this.columns[i]].letter+"$"+(1+this.sheetHandler.settings.headers.length)).fill.clear();
        this.sheetHandler.protect()
    }
    async clean():Promise<void> {
        this.sheetHandler.setColumns(this.columns,await this.process());
        this.isDirty=false;
        this.sheetHandler.worksheet.tabColor.set("");
        this.sheetHandler.unprotect()
        for (let i = 0; i < this.columns.length; i++)
            this.sheetHandler.worksheet.getRange("$"+this.sheetHandler.columns[this.columns[i]].letter+"$"+(1+this.sheetHandler.settings.headers.length)).fill.clear();
        this.sheetHandler.protect()
    }
}

abstract class sheetHandlerAbstract {
	context:Excel.RequestContext;
    htmlConsole:myConsoleType;
    templateHandler:TemplateHandler;
    name:string;
    worksheet:worksheetWrapper;
    constructor(_context: Excel.RequestContext, _htmlConsole: myConsoleType,_templateHandler:TemplateHandler,_name:string) {
        this.context=_context;
        this.htmlConsole=_htmlConsole;
        this.templateHandler=_templateHandler;
        this.name=_name;
        this.worksheet=new worksheetWrapper(context);
    }
    abstract init():Promise<void>;
}

interface DataSetQueueItem {
    address:string;
    data:[any][];
}
class TableSheetHandler extends sheetHandlerAbstract {
    table:tableWrapper;
    settings:TableSheetSettings;
    protectionOptions:Excel.WorksheetProtectionOptions={
        allowFormatColumns: false,
        allowInsertColumns: false,
        allowDeleteColumns: false,
        allowFormatRows: false,
        allowInsertRows: false,
        allowDeleteRows: false,

        allowFormatCells: false,
        allowEditScenarios: false,
        allowPivotTables: false,
        allowInsertHyperlinks: true,

        allowAutoFilter: false,
        allowSort: false,
        allowEditObjects: true
    }
    isProtected:boolean=false;

    public columnByName:{[name:string]:number}={};
    public columns:TableSheetColumnData[]=[];
    public static currColumnGroup:TableSheetColumnGroup|undefined=undefined;
    public columnDependents:TableSheetColumnGroup[][] = []; // list of column group dependents by column
    public columnGroupsByColumn:{[key:string]:number} = {};
    private columnValidators:(((newValue:any)=>any|undefined)|undefined)[] = [];

    private data:any[][]=[];

    private suppressOnProtectionChanged:number=0;
    private suppressOnSelectionChanged:number=0;
    public isCursorLocked:boolean = false;
    public constructor(_context: Excel.RequestContext, _htmlConsole: myConsoleType,_templateHandler:TemplateHandler,_name:string,_settings:TableSheetSettings) {
        super(_context,_htmlConsole,_templateHandler,_name);
        this.settings=_settings;
        this.table=new tableWrapper(this.worksheet);
    }
    public async init():Promise<void> {
        // check if worksheet is null, if it is, create it
        this.worksheet.getWorksheet(this.name);
        if (await this.worksheet.isNullObject.asyncGet()) {
            this.context.application.suspendScreenUpdatingUntilNextSync();
            await this.worksheet.addWorksheet(this.name);
        } else this.worksheet.unprotect(this.name+"PASSWORD");
        // check if table is null, if it is, create it
        const tableName:string = this.name+"TABLE";
        this.table.getTable(tableName);
        if (await this.table.isNullObject.asyncGet()) await this.table.addTable(tableName,this.getTableAddress(),true);
        // read the data and resize to enforce 
        await this.readData();
        this.setHeaders();
        this.worksheet.freezeRows(this.settings.headers.length+1)
        this.unprotect();
        await this.resize();
        this.setTotals();
        this.setFormat();
        this.protect();
        // add event callbacks
        this.worksheet.worksheet!.onNameChanged.add((async (args:Excel.WorksheetNameChangedEventArgs)=>{this.worksheet.worksheet!.name=this.name;await this.context.sync();}).bind(this))
        this.worksheet.worksheet!.onProtectionChanged.add((async (args:Excel.WorksheetProtectionChangedEventArgs)=>{
            if (this.suppressOnProtectionChanged==0 && args.isProtected==false){
                console.log(args);
                this.isProtected=false;
                this.protect(); await this.context.sync();
            } else if (this.suppressOnProtectionChanged>0 && args.isProtected==false) this.suppressOnProtectionChanged--;
        }).bind(this)); // prevents user from unprotecting the sheet
        this.worksheet.worksheet!.onVisibilityChanged.add((async (args:Excel.WorksheetVisibilityChangedEventArgs)=>{if (args.visibilityAfter!="Visible")this.worksheet.unhide();await this.context.sync();}).bind(this));
        this.worksheet.worksheet!.onSelectionChanged.add((async (args:Excel.WorksheetSelectionChangedEventArgs)=>{
            if (this.suppressOnSelectionChanged>0) { this.suppressOnSelectionChanged--; return; }
            else if (this.templateHandler.isCursorLocked) { console.log("setting selected range to \""+this.name+"!$A$1\""); this.suppressOnSelectionChanged++; this.worksheet.getRange("$A$1").select(); await this.context.sync(); }
        }).bind(this));
        this.worksheet.worksheet!.onChanged.add(this.onChanged.bind(this));
        // sync to apply all changes
        await this.context.sync();
    }
    anyRowHasTotals:boolean = false;
    public addColumn(settings:TableSheetColumnSettings):void {
        this.columnByName[settings.name]=this.columns.length;
        // store settings along with the column "letter" A,B,C,D, etc..
        const columnData:TableSheetColumnData = {
            isInputColumn: settings.isInputColumn,
            name: settings.name,
            numberFormat: settings.numberFormat,
            columnWidth: settings.columnWidth,
            alignment: settings.alignment,
            bgColor: settings.bgColor,
            wrapText: settings.wrapText,
            hasTotal: settings.hasTotal&&(settings.totalType!=undefined)&&(settings.totalType!="Custom"||settings.totalCustomValue!=null),
            totalType: settings.totalType,
            totalCustomValue: settings.totalCustomValue,

            letter: columnsAlphebet[this.columns.length],
            sum:0,
            count:0,
            countN:0
        };
        this.columns.push(columnData);
        this.columnValidators.push(undefined);
        if (columnData.hasTotal) this.anyRowHasTotals=true;
        this.columnDependents.push([]);
    }
    // sets function to validate all input, it takes in the input and returns either the new value or undefined, which sets it to the old value
    public setColumnValidation(name:string,fn:((newValue:any)=>any|undefined)) {
        // get the index of the column by its name
        let index:number|undefined = this.columnByName[name];
        if (index==undefined) { console.error("Could not find column \""+name+"\""); return []; }
        if (!this.columns[index].isInputColumn) { console.error("Validation functions can only be added to input columns."); return []; }
        this.columnValidators[this.columnByName[name]]=fn;
    }

    public async cleanColumn(name:string):Promise<void> {
        // clean any column group that the column you are trying to read belongs to
        // if it is not dirt the function will return immediately
        const columnGroupIndex:number|undefined = this.columnGroupsByColumn[name];
        if (columnGroupIndex!=undefined) {
            const group:TableSheetColumnGroup = this.templateHandler.columnGroups[columnGroupIndex];
            if (TableSheetColumnGroup.loopDetector.includes(columnGroupIndex)) { this.htmlConsole.log("LOOP DETECTED"); TableSheetColumnGroup.loopDetector=[]; return; }
            TableSheetColumnGroup.loopDetector.push(columnGroupIndex);
            if (group.isDirty) {
                if (group.hasInit) return group.clean();
                else return group.init();
            }
            if (TableSheetColumnGroup.loopDetector.length!=0) TableSheetColumnGroup.loopDetector.pop();// should pop the columnGroupIndex off the list
        }
    }
    public async getColumn(name:string):Promise<any[]> {
        // get the index of the column by its name
        let index:number|undefined = this.columnByName[name];
        if (index==undefined) { console.error("Could not find column \""+name+"\""); return []; }
        // clean the column before reading it if its needed
        await this.cleanColumn(name);
        // add column group to columns dependencies if needed
        if (TableSheetHandler.currColumnGroup!=undefined) this.columnDependents[index].push(TableSheetHandler.currColumnGroup);
        // return data for the column requested
        return this.data.map((el:any[])=>el[index!]);
    }
    public async getColumns(names:string[]):Promise<any[][]> {
        // parse the names of the columns into their indices
        let indices:number[] = [];
        for (const name of names) {
            const index:number|undefined = this.columnByName[name];
            if (index==undefined) { console.error("Could not find column \""+name+"\""); return []; }
            await this.cleanColumn(name);
            indices.push(index);
        }
        // add column group to each columns dependencies if needed
        if (TableSheetHandler.currColumnGroup!=undefined) for (const index of indices) this.columnDependents[index].push(TableSheetHandler.currColumnGroup);
        // return data for each column in the order it was requested
        return this.data.map((el:any[])=>{ return indices.map((index:number)=>el[index]); });
    }

    public preProcess():void {
        // pop completely empty lines and lines containing only calculated values
        for (let i = this.data.length-1; i >= 0; i--) {
            var lineEmpty:boolean=true;
            for (let j = 0; j < this.data[i].length; j++) {
                if (this.data[i][j]!=""&&this.columns[j].isInputColumn) {lineEmpty=false;break;}
            }
            if (lineEmpty) this.data.pop();
            else break;
        }
    }
    private queue:DataSetQueueItem[] = [];
    public setColumns(columns:(string[]|number[]),values:any[][]):void {
        if (values.length==0) return;
        if (values[0].length!=columns.length) { console.error("Invalid data width settings columns \""+columns.join(",")+"\" on sheet \""+this.name+"\"."); console.log("data:");console.log(values); return; }
        let indices:number[] = [];
        // parse "columns" into indices whether it is the names of the columns or the column indices already
        if (typeof columns[0] == "string") {
            for (const name of columns) {
                const index:number|undefined = this.columnByName[name];
                if (index==undefined) { console.error("Could not find column \""+name+"\""); return; }
                indices.push(index);
            }
        } else indices.push(...(columns as number[]));
        // push empty lines if there is no room to fit new data
        if (values.length>this.data.length) {
            const emptyLine:string = JSON.stringify(this.columns.map(()=>""));// json string of a row with the correct number of columns filled with empty strings
            for (let i = this.data.length; i <= values.length; i++) {
                this.data.push(JSON.parse(emptyLine) as any[]);
            }
        }
        // set column data in "this.data" and restructure data into individual columns to set with the excel api in postProcess
        for (let i = 0; i < indices.length; i++) {
            let columnData:[any][] = [];
            for (let y = 0; y < values.length; y ++) {
                columnData.push([values[y][i]]);
                this.data[y][indices[i]]=values[y][i];
            }
            // push empty values if new data is shorter than old data
            for (let y = values.length; y < this.data.length; y++) {
                this.data[y][indices[i]]="";
            }
            const letter:string = this.columns[indices[i]].letter;
            this.queue.push({
                address:"$"+letter+"$"+(2+this.settings.headers.length)+":$"+letter+"$"+(values.length+(1+this.settings.headers.length)),
                data:columnData
            });
        }
        // pop completely empty lines
        for (let i = this.data.length-1; i >= 0; i--) {
            var lineEmpty:boolean=true;
            for (let j = 0; j < this.data[i].length; j++) {
                if (this.data[i][j]!="") {lineEmpty=false;break;}
            }
            if (lineEmpty) this.data.pop();
            else break;
        }
        // updated sums of changed columns
        for (const i of indices) {
            this.columns[i].sum=0;
            this.columns[i].count=0;
            this.columns[i].countN=0;
            for (let j = 0; j < this.data.length; j++) {
                const cell:any = this.data[j][i];
                if (cell === "") continue;
                this.columns[i].count++;
                if (typeof cell == "number") {
                    this.columns[i].countN++;
                    this.columns[i].sum+=cell;
                }
            }
        }
    }
    public async postProcess():Promise<void> {
        // resize data, set data from the queue, clear the queue, reset formatting, and sync.
        this.unprotect();
        await this.resize();
        for (let i = 0; i < this.queue.length; i++)
            this.worksheet.getRange(this.queue[i].address).values.set(this.queue[i].data);
        this.queue=[];
        this.setTotals();
        this.setFormat();
        this.protect();
        await this.context.sync();
    }

    //#region get address functions
    // address which is from A1 to column CA and 1000 lines below where the table ends
    private getSheetAddress():string {
        return "$A$1:$CA$"+(1+this.settings.headers.length+Math.max(this.data.length+this.settings.numBufferLines,1)+1000);
    }
    // address containing rows of table headers body and totals but all the way to column CA
    private getUnlockedAreaAddress():string {
        return "$A$"+(1+this.settings.headers.length)+":$CA$"+(1+this.settings.headers.length+Math.max(this.data.length+this.settings.numBufferLines,1)+(this.anyRowHasTotals?1:0));
    }
    // address containing table headers body and totals
    private getTableAddress():string {
        return "$A$"+(1+this.settings.headers.length)+":$"+columnsAlphebet[this.columns.length-1]+"$"+(1+this.settings.headers.length+Math.max(this.data.length+this.settings.numBufferLines,1));
    }
    // address containing just table headers
    private getHeaderAddress():string {
        const row:number = 1+this.settings.headers.length;
        return "$A$"+row+":$"+columnsAlphebet[this.columns.length-1]+"$"+row;
    }
    // address containing just table body
    private getBodyAddress():string {
        return "$A$"+(2+this.settings.headers.length)+":$"+columnsAlphebet[this.columns.length-1]+"$"+(1+this.settings.headers.length+Math.max(this.data.length+this.settings.numBufferLines,1));
    }
    // address containing just table totals
    private getTotalsAddress():string {
        if (!this.anyRowHasTotals) return "";
        const row:number = 1+this.settings.headers.length+this.data.length+this.settings.numBufferLines+1;
        return "$A$"+row+":$"+columnsAlphebet[this.columns.length-1]+"$"+row;
    }
    //#endregion get address functions

    private lastTableSize:number = -1;
    // does a full read of all data in the table.
    public async readData():Promise<void> {
        // read data from table
        this.data=await this.table.values.asyncGet();
        this.lastTableSize = this.data.length;
        for (let x = 0; x < this.columns.length; x++) {
            const columnValidator:((input:any)=>(any|undefined))|undefined = this.columnValidators[x];
            if (columnValidator!=undefined) {
                let columnChanged:boolean = false;
                let lowestRowChanged:number = -1;
                let highestRowChanged:number = 0;
                for (let y = 0; y < this.data.length; y++) {
                    const oldValue:any = this.data[y][x];
                    let newValue:any = columnValidator!(oldValue);
                    if (newValue==undefined) newValue=this.data[y][x];
                    if (oldValue!=newValue) {
                        columnChanged=true;
                        if (lowestRowChanged==-1) lowestRowChanged = y;
                        highestRowChanged = y;
                        this.data[y][x]=newValue;
                    }
                }
                if (columnChanged) {
                    let data:any[][] = [];
                    for (let y = lowestRowChanged; y <= highestRowChanged; y++)
                        data.push([this.data[y][x]]);
                    this.worksheet.getRange("$"+this.columns[x].letter+"$"+(lowestRowChanged+(2+this.settings.headers.length))+":$"+this.columns[x].letter+"$"+(highestRowChanged+(2+this.settings.headers.length))).values.set(data);
                    await this.context.sync();
                }
            }
        }
        // pop completely empty lines
        for (let i = this.data.length-1; i >= 0; i--) {
            var lineEmpty:boolean=true;
            for (let j = 0; j < this.data[i].length; j++) {
                if (this.data[i][j]!="") {lineEmpty=false;break;}
            }
            if (lineEmpty) this.data.pop();
            else break;
        }
        // calculate totals 
        for (let i = 0; i < this.columns.length; i++) {
            if (this.columns[i].isInputColumn) {
                for (let j = 0; j < this.data.length; j++) {
                    const cell:any = this.data[j][i];
                    if (cell === "") continue;
                    this.columns[i].count++;
                    if (typeof cell == "number") {
                        this.columns[i].countN++;
                        this.columns[i].sum+=cell;
                    }
                }
            }
        }
    }
    private async onChanged(args:Excel.WorksheetChangedEventArgs):Promise<void> {
        if (args.triggerSource=="ThisLocalAddin") return;// dont check for changes from the add-in itself
        const isSingleCell = !args.address.includes(":");
        // parse the beginning and endings rows and columns of the range that was modified
        // if it was just one cell rowStart==rowEnd and columnStart==columnEnd
        let rowStart:number;
        let columnStart:number;
        let rowEnd:number;
        let columnEnd:number;
        if (isSingleCell) {
            const address:string = args.address;
            rowStart = rowEnd = parseInt(address.replace(/\D/g,""));
            columnStart = columnEnd = columnsAlphebet.indexOf(address.replace(rowEnd.toString(),""))
        } else {
            const address:[string,string]=args.address.split(":") as [string,string];
            rowStart = parseInt(address[0].replace(/\D/g,""));
            columnStart = columnsAlphebet.indexOf(address[0].replace(rowStart.toString(),""));
            rowEnd = parseInt(address[1].replace(/\D/g,""));
            columnEnd = columnsAlphebet.indexOf(address[1].replace(rowEnd.toString(),""));
        }
        if (rowStart>rowEnd || columnStart>columnEnd) { this.htmlConsole.log("ERROR"); return;}// ERROR
        // adjust values for indexing
        rowStart-=(2+this.settings.headers.length);
        rowEnd-=(2+this.settings.headers.length);
        // check if any data that was changed was outside of the allowed range, and reset it if it was.
        const highestRowAllowed=this.data.length+this.settings.numBufferLines-1;
        const highestColumnAllowed=this.columns.length-1;
        if (rowStart==-1) { this.setHeaders(); rowStart=0; if (isSingleCell) return; }// if the data overrode the headers
        if (rowEnd==-1) return;// if range also ended on the headers row, just return
        if (rowStart>highestRowAllowed+1) { this.worksheet.getRange(args.address).clear().select(); await this.context.sync(); return;}// if changed area row is completely out of range of the table
        if (columnStart>highestColumnAllowed) { this.worksheet.getRange(args.address).clear().select(); await this.context.sync(); return;}// if changed area column is completely out of range of the table
        if (rowEnd>highestRowAllowed) {
            if (columnEnd>highestColumnAllowed) {
                const bottomRight:string = ":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+(2+this.settings.headers.length));
                const topRight1:string = "$"+columnsAlphebet[columnStart]+"$"+(this.data.length+(2+this.settings.headers.length));
                const topRight2:string = "$"+columnsAlphebet[this.columns.length]+"$"+(rowStart+(2+this.settings.headers.length));
                this.worksheet.getRange(topRight1+bottomRight).clear();
                this.worksheet.getRange(topRight2+bottomRight).clear();
                columnEnd=highestColumnAllowed;
            } else {
                this.worksheet.getRange("$"+columnsAlphebet[columnStart]+"$"+(this.data.length+(2+this.settings.headers.length))+":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+(2+this.settings.headers.length))).clear();
            }
            if (rowStart<=highestRowAllowed+1) this.setTotals();
            rowEnd=highestRowAllowed; await this.context.sync();
        } else if (columnEnd>highestColumnAllowed) {
            this.worksheet.getRange("$"+columnsAlphebet[this.columns.length]+"$"+(rowStart+(2+this.settings.headers.length))+":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+(2+this.settings.headers.length))).clear();
            columnEnd=highestColumnAllowed; await this.context.sync();
        }
        // if the new data was outside of the range of "this.data" add empty lines, (somewhere in the buffer lines)
        if (rowEnd>=this.data.length) {
            const emptyLine:string = JSON.stringify(this.columns.map(()=>""));// json string of a row with the correct number of columns filled with empty strings
            for (let i = this.data.length; i <= rowEnd; i++) {
                this.data.push(JSON.parse(emptyLine) as any[]);
            }
        }

        // validate that data was actually changed and if it wasn't an input column, reset it
        if (isSingleCell) {
            if (JSON.stringify(args.details.valueBefore)===JSON.stringify(args.details.valueAfter)) return;// return if there was no change
            if (this.columns[columnStart].isInputColumn) {
                // if the column validator function is undefined, just use the input, if the validator returns undefined, reset it to the old value
                const oldValue:any = this.data[rowStart][columnStart];
                let newValue:any = args.details.valueAfter;
                const columnValidator:((input:any)=>(any|undefined))|undefined = this.columnValidators[columnStart];
                if (columnValidator!=undefined) {
                    newValue = columnValidator!(newValue);
                    if (newValue==undefined) newValue=oldValue;
                    if (oldValue!=newValue) {
                        this.data[rowStart][columnStart]=newValue;
                        this.worksheet.getRange("$"+this.columns[columnStart].letter+"$"+(rowStart+(2+this.settings.headers.length))).values.set([[newValue]]);
                        await this.context.sync();
                    }
                } else
                    this.data[rowStart][columnStart]=args.details.valueAfter;
            } else { this.worksheet.getRange("$"+this.columns[columnStart].letter+"$"+(rowStart+(2+this.settings.headers.length))).values.set([[this.data[rowStart][columnStart]]]); await this.context.sync(); }
        } else {
            const newDataAddress:string = "$"+columnsAlphebet[columnStart]+"$"+(rowStart+(2+this.settings.headers.length))+":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+(2+this.settings.headers.length));
            let newData:any[][] = await this.worksheet.getRange(newDataAddress).values.asyncGet();
            let changed:boolean = false;
            let inputChanged:boolean = false;
            for (let x = columnStart; x <= columnEnd; x++) {
                if (this.columns[x].isInputColumn) {
                    for (let y = rowStart; y <= rowEnd; y++) {
                        const oldValue:any = this.data[y][x];
                        const rawInput:any = newData[y-rowStart][x-columnStart];
                        if (oldValue==rawInput) continue;
                        let newValue:any = rawInput;
                        // if the column validator function is undefined, just use the input, if the validator returns undefined, reset it to the old value
                        const columnValidator:((input:any)=>(any|undefined))|undefined = this.columnValidators[x];
                        if (columnValidator!=undefined) {
                            newValue = columnValidator!(newValue);
                            if (newValue==undefined) newValue=oldValue;
                            else if (newValue!=rawInput) {
                                inputChanged=true;
                                newData[y-rowStart][x-columnStart]=newValue;
                            }
                        }
                        if (oldValue==newValue) continue;
                        changed=true;
                        this.data[y][x]=newValue;
                        // maintain correct total values
                        if (oldValue !== "") {
                            this.columns[x].count--;
                            if (typeof oldValue == "number") {
                                this.columns[x].countN--;
                                this.columns[x].sum-=oldValue;
                            }
                        }
                        if (newValue !== "") {
                            this.columns[x].count++;
                            if (typeof newValue == "number") {
                                this.columns[x].countN++;
                                this.columns[x].sum+=newValue;
                            }
                        }
                    }
                } else {
                    for (let y = rowStart; y <= rowEnd; y++) {
                        if (!inputChanged && this.data[y][x]!=newData[y-rowStart][x-columnStart]) inputChanged=true;
                        newData[y-rowStart][x-columnStart]=this.data[y][x];
                    }
                }
            }
            if (inputChanged) {
                this.worksheet.getRange(newDataAddress).values.set(newData);
                this.context.sync();
            }
            if (!changed) return;// return if there was no change
        }
        // pop completely empty lines
        for (let i = this.data.length-1; i >= 0; i--) {
            var lineEmpty:boolean=true;
            for (let j = 0; j < this.data[i].length; j++) {
                if (this.data[i][j]!="") {lineEmpty=false;break;}
            }
            if (lineEmpty) this.data.pop();
            else break;
        }
        // resize table and update totals
        this.unprotect();
        await this.resize();
        this.setTotals();
        this.protect();
        // find groups that have now become "dirty"
        for (let x = columnStart; x <= columnEnd; x++) {
            if (this.columns[x].isInputColumn) {
                const columnGroups:TableSheetColumnGroup[] = this.columnDependents[x];
                for (let i = 0; i < columnGroups.length; i++) {
                    columnGroups[i].setDirty();
                }
            }
        }
        await this.context.sync();
    }

    //#region formatting functions
    public unprotect() {
        if (!this.isProtected) return;
        this.isProtected=false;
        this.suppressOnProtectionChanged++;
        this.worksheet.unprotect(this.name+"PASSWORD");
    }
    public protect() {
        // dont try to protect it if it is already
        if (this.isProtected) return;
        this.isProtected=true;
        this.worksheet.getRange(this.getSheetAddress()).lock().hideFormulas();
        this.worksheet.getRange(this.getUnlockedAreaAddress()).unlock().unhideFormulas();
        if (!this.settings.doLockA1) this.worksheet.getRange("$A$1").unlock().unhideFormulas();
        this.worksheet.protect(this.protectionOptions,this.name+"PASSWORD");
    }
    private async resize():Promise<void> {
        this.worksheet.getRange("$A$"+(this.data.length+2+this.settings.headers.length)+":$"+this.columns[this.columns.length-1].letter+"$"+(1+this.settings.headers.length+Math.max(this.data.length+this.settings.numBufferLines+1,this.lastTableSize))).clear();
        if ((this.data.length+this.settings.numBufferLines+1)==this.lastTableSize&&this.lastTableSize!=-1) return;
        this.table.showTotals.set(false);
        await this.table.resize(this.getTableAddress());
        this.table.showTotals.set(this.anyRowHasTotals);
        this.lastTableSize=this.data.length+this.settings.numBufferLines+1;
        return;
    }
    private setHeaders():void {
        // set table headers
        this.worksheet.getRange(this.getHeaderAddress())
            .values.set([this.columns.map((el:TableSheetColumnSettings)=>el.name)])
            .setFontColor("#000000")
            .bold.set(true)
            .wrapText.set(true)
            .verticalAlignment.set("Top")
            .horizontalAlignment.set("Center");
        // set sheet headers
        if (this.settings.headers.length==0 || this.settings.headers[0].length==0) return;
        if (this.settings.headerOverrideA1) {
            this.worksheet.getRange("$A$1:$"+columnsAlphebet[this.settings.headers[0].length-1]+"$"+this.settings.headers.length)
                .values.set(this.settings.headers)
                .setFontSize(this.settings.headersFontSize);
        } else {
            // set font size for whole area
            this.worksheet.getRange("$A$1:$"+columnsAlphebet[this.settings.headers[0].length-1]+"$"+this.settings.headers.length)
                .setFontSize(this.settings.headersFontSize);
            // set other data just not cell A1
            if (this.settings.headers[0].length>1)
                this.worksheet.getRange("$B$1:$"+columnsAlphebet[this.settings.headers[0].length-1]+"$1").values.set([this.settings.headers[0].filter((cell:any,index:number)=>index!=0)]);
            if (this.settings.headers.length>1)
                this.worksheet.getRange("$A$2:$"+columnsAlphebet[this.settings.headers[0].length-1]+"$"+this.settings.headers.length).values.set(this.settings.headers.filter((row:any[],index:number)=>index!=0));
        }
    }
    private setTotals():void {
        this.worksheet.getRange(this.getTotalsAddress())
            .values.set([this.columns.map((column:TableSheetColumnData)=>{
                if (!column.hasTotal) return "";
                switch (column.totalType) {
                    case "Cnt":
                        return column.countN;
                    case "CntA":
                        return column.count;
                    case "Sum":
                        return column.sum;
                    case "Avg":
                        return column.sum/(column.countN||1);
                    case "Custom":
                        return column.totalCustomValue!;
                    default:
                        return "";
                }
            })])
            .numberFormat.set(
                [this.columns.map((data:TableSheetColumnData)=>data.numberFormat)]
            );
        //console.log(this.getTotalsAddress())
        //console.log([this.columns.map((data:TableSheetColumnData)=>data.numberFormat)])
    }
    private setFormat():void {
        if (this.data.length==0) return;
        const upperLimit:number = 2+this.settings.headers.length;
        const lowerLimit:number = 1+this.settings.headers.length+this.data.length+this.settings.numBufferLines+(this.anyRowHasTotals?1:0);
        for (const column of this.columns) {
            const widthRange:rangeWrapper = this.worksheet.getRange("$"+column.letter+"$"+upperLimit)
            if (column.columnWidth==undefined) widthRange.columnWidth.set((column.columnWidth*11+7)/2);// convert character to pixels
            else if (column.columnWidth==0) widthRange.columnWidth.set(0);
            else widthRange.columnWidth.set(86);// (15*11+7)/2 or 15 characters

            const range:rangeWrapper = this.worksheet.getRange("$"+column.letter+"$"+upperLimit+":$"+column.letter+"$"+lowerLimit);
            range.numberFormat.set([[column.numberFormat||"@"]]);
            range.horizontalAlignment.set(column.alignment);
            if (column.bgColor!=undefined) range.fill.set(column.bgColor!);
            if (column.wrapText!=undefined) range.wrapText.set(column.wrapText!);
        }
    }
    //#endregion
}
class DataSheetHandler extends sheetHandlerAbstract {
    //public settings:DataSheetSettings;

    public columnByName:{[name:string]:number}={};
    public columns:DataSheetColumnData[]=[];
    public static currColumnGroup:TableSheetColumnGroup|undefined=undefined;
    public columnDependents:TableSheetColumnGroup[][] = []; // list of column group dependents by column
    
    private data:any[][]=[];

    constructor(_context: Excel.RequestContext, _htmlConsole: myConsoleType,_templateHandler:TemplateHandler,_name:string/*,_settings:DataSheetSettings*/) {
        super(_context,_htmlConsole,_templateHandler,_name);
        //this.settings=_settings;
    }
    public async init():Promise<void> {
        this.worksheet.getWorksheet(this.name);
        if (await this.worksheet.isNullObject.asyncGet())
            await this.worksheet.addWorksheet(this.name);

        this.worksheet.getRange("$A$1:$"+this.columns[this.columns.length-1].letter+"$1")
            .values.set([this.columns.map((el:DataSheetColumnSettings)=>el.name)])
            .setFontColor("#000000")
            .bold.set(true)
            .wrapText.set(true)
            .verticalAlignment.set("Top")
            .horizontalAlignment.set("Center");
        await this.readData();
        this.setFormat();
        
        this.worksheet.worksheet!.onNameChanged.add((async (args:Excel.WorksheetNameChangedEventArgs)=>{this.worksheet.worksheet!.name=this.name;await this.context.sync();}).bind(this));
        this.worksheet.worksheet!.onVisibilityChanged.add((async (args:Excel.WorksheetVisibilityChangedEventArgs)=>{if (args.visibilityAfter!="Visible")this.worksheet.unhide();await this.context.sync();}).bind(this));
        this.worksheet.worksheet!.onChanged.add(this.onChanged.bind(this));
        await this.context.sync();
    }
    public addColumn(settings:DataSheetColumnSettings):void {
        this.columnByName[settings.name]=this.columns.length;
        this.columns.push({
            name: settings.name,
            columnWidth: settings.columnWidth,
            alignment: settings.alignment,
            letter:columnsAlphebet[this.columns.length]
        });
        this.columnDependents.push([]);
    }

    public getColumn(name:string):any[] {
        let index:number|undefined = this.columnByName[name];
        if (index==undefined) return [];
        if (TableSheetHandler.currColumnGroup!=undefined) this.columnDependents[index].push(TableSheetHandler.currColumnGroup);
        return this.data.map((el:any[])=>el[index!]);
    }
    public getColumns(names:string[]):any[][] {
        let indices:number[] = [];
        for (const name of names) {
            const index:number|undefined = this.columnByName[name];
            if (index==undefined) { console.error("Could not find column \""+name+"\""); return []; }
            indices.push(index);
        }
        if (TableSheetHandler.currColumnGroup!=undefined) for (const index of indices) this.columnDependents[index].push(TableSheetHandler.currColumnGroup);
        return this.data.map((el:any[])=>{ return indices.map((index:number)=>el[index]); });
    }

    public async readData() : Promise<void> {
        var currentLine:number = 2;
        const linesAtATime:number = 200;
        this.data=[];
        while (true) {
            const values:any[][] = await this.worksheet.getRange("$A$"+currentLine+":$"+this.columns[this.columns.length-1].letter+"$"+(currentLine+linesAtATime)).values.asyncGet();
            var chunkIsEmpty:boolean=true;
            for (let i = 0; i < linesAtATime; i++) {
                var columnIsEmpty:boolean=true;
                for (let j = 0; j < this.columns.length; j++) {
                    if (values[i][j]!="") { columnIsEmpty=false;break; }
                }
                if (!columnIsEmpty) { chunkIsEmpty=false; this.data.push(values[i]); }
            }
            if (chunkIsEmpty) break;
            currentLine+=linesAtATime;
        }
    }
    private async onChanged(args:Excel.WorksheetChangedEventArgs):Promise<void> {
        if (args.triggerSource=="ThisLocalAddin") return;// dont check for changes from the add-in itself
        const isSingleCell = !args.address.includes(":");
        let rowStart:number;
        let columnStart:number;
        let rowEnd:number;
        let columnEnd:number;
        if (isSingleCell) {
            const address:string = args.address;
            rowStart = rowEnd = parseInt(address.replace(/\D/g,""));
            columnStart = columnEnd = columnsAlphebet.indexOf(address.replace(rowEnd.toString(),""))
        } else {
            const address:[string,string]=args.address.split(":") as [string,string];
            rowStart = parseInt(address[0].replace(/\D/g,""));
            columnStart = columnsAlphebet.indexOf(address[0].replace(rowStart.toString(),""));
            rowEnd = parseInt(address[1].replace(/\D/g,""));
            columnEnd = columnsAlphebet.indexOf(address[1].replace(rowEnd.toString(),""));
        }
        if (rowStart>rowEnd || columnStart>columnEnd) { this.htmlConsole.log("ERROR"); return;}// ERROR
        rowStart-=2; rowEnd-=2;

        const highestColumnAllowed=this.columns.length-1;
        if (rowStart==-1) { /*this.setHeaders();*/ rowStart=0; if (isSingleCell) return; }// if the data overrode the headers
        if (rowEnd==-1) return;// if range also ended on the headers row, just return
        if (columnStart>highestColumnAllowed) { this.worksheet.getRange(args.address).clear().select(); await this.context.sync(); return;}// if changed area column is completely out of range of the table
        if (columnEnd>highestColumnAllowed) {
            this.worksheet.getRange("$"+columnsAlphebet[this.columns.length]+"$"+(rowStart+2)+":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+2)).clear();
            columnEnd=highestColumnAllowed; await this.context.sync();
        }

        // if the new data was outside of the range of "this.data" add empty lines, (somewhere in the buffer lines)
        const extendedData:boolean = rowEnd>=this.data.length;
        if (extendedData) {
            const emptyLine:string = JSON.stringify(this.columns.map(()=>""));// json string of a row with the correct number of columns filled with empty strings
            for (let i = this.data.length; i <= rowEnd; i++) {
                this.data.push(JSON.parse(emptyLine) as any[]);
            }
        }
        if (isSingleCell) {
            if (JSON.stringify(args.details.valueBefore)===JSON.stringify(args.details.valueAfter)) return;// return if there was no change
            this.data[rowStart][columnStart]=args.details.valueAfter;
        } else {
            let newData:any[][] = await this.worksheet.getRange("$"+columnsAlphebet[columnStart]+"$"+(rowStart+2)+":$"+columnsAlphebet[columnEnd]+"$"+(rowEnd+2)).values.asyncGet();
            let changed:boolean = false;
            for (let y = rowStart; y <= rowEnd; y++) {
                var line:any[] = [];
                for (let x = columnStart; x <= columnEnd; x++) {
                    line.push(this.data[y][x]);
                    if (this.data[y][x]!=newData[y-rowStart][x-columnStart]) changed=true;
                    this.data[y][x]=newData[y-rowStart][x-columnStart];
                }
            }
            if (!changed) return;// return if there was no change
        }
        // check if there are now empty lines at the end of the data that werent there before
        let poppedData:boolean = false;
        for (let i = this.data.length-1; i >= 0; i--) {
            var lineEmpty:boolean=true;
            for (let j = 0; j < this.data[i].length; j++) {
                if (this.data[i][j]!="") {lineEmpty=false;break;}
            }
            if (lineEmpty) { poppedData=true; this.data.pop(); }
            else break;
        }
        // if the data changed size, reset the format
        if (extendedData||poppedData) this.setFormat();
        // find groups that have now become "dirty"
        for (let x = columnStart; x <= columnEnd; x++) {
            const columnGroups:TableSheetColumnGroup[] = this.columnDependents[x];
            for (let i = 0; i < columnGroups.length; i++) {
                columnGroups[i].setDirty();
            }
        }
        await this.context.sync();
    }

    private setFormat():void {
        const upperLimit:number = 2;
        const lowerLimit:number = 1+Math.max(this.data.length,1);
        for (const column of this.columns) {
            if (column.columnWidth==undefined) this.worksheet.getRange("$"+column.letter+"$1").columnWidth.set((column.columnWidth*11+7)/2);// convert character to pixels
            else if (column.columnWidth==0) this.worksheet.getRange("$"+column.letter+"$1").columnWidth.set(0);
            else this.worksheet.getRange("$"+column.letter+"$1").columnWidth.set(86);// (15*11+7)/2 or 15 characters

            const range:rangeWrapper = this.worksheet.getRange("$"+column.letter+"$"+upperLimit+":$"+column.letter+"$"+lowerLimit);
            range.horizontalAlignment.set(column.alignment);
        }
    }
}
class GuidanceSheetHandler extends sheetHandlerAbstract {
    settings:GuidanceSheetSettings;
    constructor(_context: Excel.RequestContext, _htmlConsole: myConsoleType,_templateHandler:TemplateHandler,_settings:GuidanceSheetSettings) {
        super(_context,_htmlConsole,_templateHandler,"Guidance");
        this.settings=_settings;
    }
    async init():Promise<void> {
        this.worksheet.getWorksheet(this.name);
        if (await this.worksheet.isNullObject.asyncGet()) {
            await this.worksheet.addWorksheet(this.name);
            // set guidance data
        }
    }
}
class TemplateHandler {
	context:Excel.RequestContext;
    htmlConsole:myConsoleType;

    tableSheetsByName:{[key:string]:number} = {};
    dataSheetsByName:{[key:string]:number} = {};
    tableSheets:TableSheetHandler[] = [];
    dataSheets:DataSheetHandler[] = [];
    guidanceSheet:GuidanceSheetHandler|undefined=undefined;
    SheetInitOrder:sheetHandlerAbstract[]=[];

    columnGroups:TableSheetColumnGroup[] = [];

    constructor(context: Excel.RequestContext, _htmlConsole: myConsoleType) {
        this.context=context;this.htmlConsole=_htmlConsole;
        addHtmlButton("Process",this.process.bind(this));
    }
    public activeWorksheet:worksheetWrapper|undefined;
    public activeRange:rangeWrapper|undefined;
    public isCursorLocked:boolean = false;
    private async lockCursor() {
        (new worksheetWrapper(this.context)).getActiveWorksheet().getRange("$A$1").select();
        this.context.sync();
        this.activeRange = (new rangeWrapper(this.context)).getSelectedRange().track();
        this.isCursorLocked=true;
    }
    private async unlockCursor() {
        if (!this.isCursorLocked) return;
        this.isCursorLocked=false;
        this.activeRange!.select().untrack();
        this.activeWorksheet=undefined;
        this.activeRange=undefined;
        await this.context.sync();
    }
    async init():Promise<void> {
        //this.context.application.calculationMode=Excel.CalculationMode.manual;
        //this.context.application.suspendScreenUpdatingUntilNextSync();
        // setup sheets
        for (let i = 0; i < this.SheetInitOrder.length; i++) {
            await this.SheetInitOrder[i].init();
            this.SheetInitOrder[i].worksheet.worksheet!.position=i;
        }
        if (this.guidanceSheet!=undefined) { await this.guidanceSheet.init(); this.guidanceSheet.worksheet.worksheet!.position=this.SheetInitOrder.length; }

        // delete "Sheet1"
        var Sheet1:worksheetWrapper = new worksheetWrapper(this.context).getWorksheet("Sheet1");
        Sheet1.worksheet!.delete();
        await this.context.sync();

        // setup column groups
        for (let i = 0; i < this.tableSheets.length; i++) this.tableSheets[i].preProcess();
        await this.lockCursor();
        for (let i = 0; i < this.columnGroups.length; i++) await this.columnGroups[i].init();
        for (let i = 0; i < this.tableSheets.length; i++) await this.tableSheets[i].postProcess();
        await this.unlockCursor();
        //this.context.application.calculationMode=Excel.CalculationMode.automatic;
    }
    addTableSheet(name:string,settings:TableSheetSettings):TableSheetHandler {
        this.tableSheetsByName[name]=this.tableSheets.length;
        const tmp:TableSheetHandler = new TableSheetHandler(this.context,this.htmlConsole,this,name,settings);
        this.SheetInitOrder.push(tmp);
        this.tableSheets.push(tmp);
        return tmp;
    }
    addDataSheet(name:string/*,settings:DataSheetSettings*/):DataSheetHandler {
        this.dataSheetsByName[name]=this.dataSheets.length;
        const tmp:DataSheetHandler = new DataSheetHandler(this.context,this.htmlConsole,this,name/*,settings*/);
        this.SheetInitOrder.push(tmp);
        this.dataSheets.push(tmp);
        return tmp;
    }
    async setGuidanceSheet(sheet:GuidanceSheetHandler) {
        this.guidanceSheet=sheet;
    }
    addColumnGroup(sheetHandler: TableSheetHandler, columns: string[], process: () => Promise<any[][]>) {
        let indices:number[] = [];
        for (let i = 0; i < columns.length; i++) {
            const index:number|undefined = sheetHandler.columnByName[columns[i]];
            if (index==undefined) { console.error("Could not find column \""+columns[i]+"\""); return; }
            if (sheetHandler.columns[index].isInputColumn) { console.error("Cannot create a column group setting the values of an input column"); return; }
            indices.push(index);
        }
        const columnGroup:TableSheetColumnGroup = new TableSheetColumnGroup(sheetHandler,indices,process);
        for (let i = 0; i < columns.length; i++) {
            if (sheetHandler.columnGroupsByColumn[columns[i]]!=undefined) { console.error("More than one column group may not contain the same column."); return; }
            else sheetHandler.columnGroupsByColumn[columns[i]]=this.columnGroups.length;
        }
        this.columnGroups.push(columnGroup);
    }
    //#region specialized column groups
    mapColumn(toSheet:TableSheetHandler, mappedColumn:string, mapFunction:(cell:any)=>any, fromSheet:TableSheetHandler, fromColumn:string) {
        this.addColumnGroup(toSheet,[mappedColumn],(async ():Promise<any[]> => {
            return (await fromSheet.getColumns([fromColumn])).map(([cell]:any[])=>[((cell!="")?mapFunction(cell):"")]);
        }).bind(this));
    }
    aliasColumns(toSheet:TableSheetHandler, toColumns:string[], fromSheet:TableSheetHandler|DataSheetHandler, fromColumns:string[]) {
        if (fromColumns.length!=toColumns.length) { console.error("Number or source and destination columns must match to alias columns."); return; }
        if (fromColumns.length==0) { console.error("Cannot alias 0 columns to 0 columns."); return; }
        for (let i = 0; i < toColumns.length; i++) {
            this.addColumnGroup(toSheet,[toColumns[i]],(async ():Promise<any[]> => {
                return await fromSheet.getColumns([fromColumns[i]]);
            }).bind(this));
        }
    }
    sumColumns(sumSheet:TableSheetHandler, sumColumn:string, fromSheet:TableSheetHandler, fromColumns:string[]) {
        if (fromColumns.length==0) { console.error("Cannot sum 0 columns."); return; }
        this.addColumnGroup(sumSheet,[sumColumn],(async ():Promise<any[]> => {
            return (await fromSheet.getColumns(fromColumns)).map((row:number[])=>{
                const sum:number = row.reduce((accumulator:number,cell:number)=>accumulator+cell);
                return [sum];
            });
        }).bind(this));
    }
    averageColumns(avgSheet:TableSheetHandler, avgColumn:string, fromSheet:TableSheetHandler, fromColumns:string[]) {
        if (fromColumns.length==0) { console.error("Cannot average 0 columns."); return; }
        this.addColumnGroup(avgSheet,[avgColumn],(async ():Promise<any[]> => {
            return (await fromSheet.getColumns(fromColumns)).map((row:number[])=>{
                const sum:number = row.reduce((accumulator:number,cell:number)=>accumulator+cell);
                return [((sum!=0)?sum/fromColumns.length:0)];
            });
        }).bind(this));
    }
    //#endregion specialized column groups

    async process():Promise<void> {
        for (let i = 0; i < this.tableSheets.length; i++) this.tableSheets[i].preProcess();
        await this.lockCursor();
        for (let i = 0; i < this.columnGroups.length; i++) await this.columnGroups[i].clean();
        for (let i = 0; i < this.tableSheets.length; i++) await this.tableSheets[i].postProcess();
        await this.unlockCursor();
    }
}
abstract class templateInterface {
	context:Excel.RequestContext;
	htmlConsole:myConsoleType;
    constructor(context: Excel.RequestContext, _htmlConsole: myConsoleType) {
        this.context=context;this.htmlConsole=_htmlConsole;
    }
}
var generateTemplate:((context:Excel.RequestContext,_htmlConsole:myConsoleType)=>templateInterface);
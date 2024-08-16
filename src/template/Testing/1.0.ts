class TestingTemplate extends templateInterface {
    templateHandler:TemplateHandler;
    TestingDataSheet:DataSheetHandler;
    TestingSheet:TableSheetHandler;
    byKey2Sheet:TableSheetHandler;
    byKey3Sheet:TableSheetHandler;
    constructor(_context: Excel.RequestContext, _htmlConsole: myConsoleType) {
        super(_context, _htmlConsole);
        this.templateHandler = new TemplateHandler(context,this.htmlConsole);

        this.TestingDataSheet = this.templateHandler.addDataSheet("TestingData");
        this.TestingDataSheet.addColumn({ name:"Column 1", columnWidth:25, alignment:"Left" });
        this.TestingDataSheet.addColumn({ name:"Column 2", columnWidth:25, alignment:"Left" });
        this.TestingDataSheet.addColumn({ name:"Column 3", columnWidth:25, alignment:"Left" });
        this.TestingDataSheet.addColumn({ name:"Column 4", columnWidth:25, alignment:"Left" });
        this.TestingDataSheet.addColumn({ name:"Column 5", columnWidth:25, alignment:"Left" });
        this.TestingDataSheet.addColumn({ name:"Column 6", columnWidth:25, alignment:"Left" });

        this.TestingSheet = this.templateHandler.addTableSheet("Testing",{
            headerOverrideA1:false,
            doLockA1:false,
            headers:[[""],["By Date"]],
            headersFontSize:16,
            numBufferLines:15
        });
        this.TestingSheet.addColumn({ isInputColumn:true , name:"Date",       numberFormat:DateNumberFormat   , columnWidth:12, alignment:"Center", hasTotal:true, totalType:"Cnt" });
        this.TestingSheet.addColumn({ isInputColumn:true , name:"Key 2",      numberFormat:DefaultNumberFormat, columnWidth:15, alignment:"Center", hasTotal:true, totalType:"Cnt" });
        this.TestingSheet.addColumn({ isInputColumn:true , name:"Key 3",      numberFormat:DefaultNumberFormat, columnWidth:15, alignment:"Center", hasTotal:true, totalType:"Cnt" });
        this.TestingSheet.addColumn({ isInputColumn:true , name:"Amount 1",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left", hasTotal:true, totalType:"Sum" });
        this.TestingSheet.addColumn({ isInputColumn:true , name:"Amount 2",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left", hasTotal:true, totalType:"Sum" });
        this.TestingSheet.addColumn({ isInputColumn:false, name:"Sum",        numberFormat:AccountNumberFormat, columnWidth:25, alignment:"Left", hasTotal:true, totalType:"Avg" });
        this.TestingSheet.addColumn({ isInputColumn:false, name:"DoubleSum",  numberFormat:AccountNumberFormat, columnWidth:25, alignment:"Left", hasTotal:true, totalType:"Sum" });
        this.TestingSheet.addColumn({ isInputColumn:false, name:"DataCopy 1", numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left", hasTotal:true, totalType:"Sum" });
        this.TestingSheet.addColumn({ isInputColumn:false, name:"DataCopy 2", numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left", hasTotal:true, totalType:"Sum" });
        // key columns
        this.TestingSheet.setColumnValidation("Key 2",(input:any)=>{
            if (typeof input == "string")
                return input.toUpperCase();
            else return undefined;
        });
        this.TestingSheet.setColumnValidation("Key 3",(input:any)=>{
            if (typeof input == "string")
                return input.toUpperCase();
            else return undefined;
        });
        // Sum column
        this.templateHandler.sumColumns(this.TestingSheet,"Sum",this.TestingSheet,["Amount 1","Amount 2"]);
        // DoubleSum column
        this.templateHandler.mapColumn(this.TestingSheet,"DoubleSum",(sum:any)=>sum*2,this.TestingSheet,"Sum");
        // Data Copy Columns
        this.templateHandler.aliasColumns(this.TestingSheet,["DataCopy 1","DataCopy 2"],this.TestingDataSheet,["Column 1","Column 2"]);
        
        this.byKey2Sheet = this.templateHandler.addTableSheet("ByKey2",{
            headerOverrideA1:true,
            doLockA1:true,
            headers:[["=Testing!$A$1"],["By Key 2"]],
            headersFontSize:16,
            numBufferLines:0
        });
        this.byKey2Sheet.addColumn({ isInputColumn:false, name:"Key 2",      numberFormat:DefaultNumberFormat, columnWidth:15, alignment:"Center", hasTotal:true, totalType:"Custom", totalCustomValue:"Totals:" });
        this.byKey2Sheet.addColumn({ isInputColumn:false, name:"Amount 1",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left"  , hasTotal:true, totalType:"Sum" });
        this.byKey2Sheet.addColumn({ isInputColumn:false, name:"Amount 2",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left"  , hasTotal:true, totalType:"Sum" });
        this.byKey2Sheet.addColumn({ isInputColumn:false, name:"DataCopy 2", numberFormat:AccountNumberFormat, columnWidth:25, alignment:"Left"  , hasTotal:true, totalType:"Sum" });

        this.templateHandler.addColumnGroup(this.byKey2Sheet,["Key 2","Amount 1","Amount 2","DataCopy 2"],(async ():Promise<any[]> => {
            const Columns:any[] = await this.TestingSheet.getColumns(["Key 2","Amount 1","Amount 2","DataCopy 2"]);
            var map:{[key:string]:[number,number,number]} = {};
            for (let i = 0; i < Columns.length; i++) {
                const [key2,amount1,amount2,datacopy2]:[string,number,number,number] = Columns[i];
                if ((typeof key2 == undefined)||(typeof amount1 != "number")||(typeof amount2 != "number")||(typeof datacopy2 != "number")) continue;
                else {
                    if (map[key2]!=undefined) { map[key2][0]+=amount1;map[key2][1]+=amount2;map[key2][2]+=datacopy2; }
                    else map[key2]=[amount1,amount2,datacopy2];
                }
            }
            const entries:[string,[number,number,number]][] = Object.entries(map);
            if (entries.length==0) return [["",0,0,0]];
            return entries.map(([key2,[amount1,amount2,datacopy2]]:[string,[number,number,number]])=>[key2,amount1,amount2,datacopy2]);
        }).bind(this));
        
        this.byKey3Sheet = this.templateHandler.addTableSheet("ByKey3",{
            headerOverrideA1:true,
            doLockA1:true,
            headers:[["=Testing!$A$1"],["By Key 3"]],
            headersFontSize:16,
            numBufferLines:0
        });
        this.byKey3Sheet.addColumn({ isInputColumn:false, name:"Key 3",      numberFormat:DefaultNumberFormat, columnWidth:15, alignment:"Center", hasTotal:true, totalType:"Custom", totalCustomValue:"Totals:" });
        this.byKey3Sheet.addColumn({ isInputColumn:false, name:"Amount 1",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left"  , hasTotal:true, totalType:"Sum" });
        this.byKey3Sheet.addColumn({ isInputColumn:false, name:"Amount 2",   numberFormat:AccountNumberFormat, columnWidth:20, alignment:"Left"  , hasTotal:true, totalType:"Sum" });
        this.byKey3Sheet.addColumn({ isInputColumn:false, name:"Average",    numberFormat:AccountNumberFormat, columnWidth:25, alignment:"Left"  , hasTotal:true, totalType:"Avg" });
        this.byKey3Sheet.addColumn({ isInputColumn:false, name:"DataCopy 1", numberFormat:AccountNumberFormat, columnWidth:25, alignment:"Left"  , hasTotal:true, totalType:"Sum" });
        this.templateHandler.addColumnGroup(this.byKey3Sheet,["Key 3","Amount 1","Amount 2","DataCopy 1"],(async ():Promise<any[]> => {
            const Columns:any[] = await this.TestingSheet.getColumns(["Key 3","Amount 1","Amount 2","DataCopy 1"]);
            var map:{[key:string]:[number,number,number]} = {};
            for (let i = 0; i < Columns.length; i++) {
                const [key3,amount1,amount2,datacopy1]:[string,number,number,number] = Columns[i];
                if ((typeof key3 == undefined)||(typeof amount1 != "number")||(typeof amount2 != "number")||(typeof datacopy1 != "number")) continue;
                else {
                    if (map[key3]!=undefined) { map[key3][0]+=amount1;map[key3][1]+=amount2;map[key3][2]+=datacopy1; }
                    else map[key3]=[amount1,amount2,datacopy1];
                }
            }
            const entries:[string,[number,number,number]][] = Object.entries(map);
            if (entries.length==0) return [["",0,0,0,0]];
            return entries.map(([key3,[amount1,amount2,datacopy1]]:[string,[number,number,number]])=>[key3,amount1,amount2,datacopy1]);
        }).bind(this));
        this.templateHandler.averageColumns(this.byKey3Sheet,"Average",this.byKey3Sheet,["Amount 1","Amount 2"]);

        (async () => {
            await addMetaDataHTMLTextInput("metadata","test",async()=>{});
            await this.templateHandler.init();
        }).bind(this)();
    }
}
generateTemplate = (context: Excel.RequestContext, myConsole: myConsoleType): templateInterface => (new TestingTemplate(context, myConsole) as templateInterface);
namespace Guidance {
    var columns: string[] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", 
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", 
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM"];
    export interface Content {
        text:any[][];
        columnWidths?:number[];
        bold?:boolean[][];
        fontSizes?:number[][];
    }
    export class Handler {
        context:Excel.RequestContext;
        sheet:worksheetWrapper;
        password:string;
        text:any[][];
        columnWidths:number[];
        bold:boolean[][];
        fontSizes:number[][];
        
        constructor(context:Excel.RequestContext,password:string,content:Content) {
            this.context=context;
            this.sheet=new worksheetWrapper(this.context);
            this.password=password
            this.text=content.text;
            this.columnWidths=content.columnWidths||[];
            this.bold=content.bold||[];
            this.fontSizes=content.fontSizes||[];
        }
        async create() {
            this.sheet.getWorksheet("Guidance").track();
            await this.context.sync();
            if (!(await this.sheet.isNullObject.asyncGet())) return;// do nothing if it already exists

            await this.sheet.addWorksheet("Guidance");
            this.sheet.getRange("$A$1:$"+columns[this.text[0].length-1]+"$"+this.text.length).values.set(this.text).wrapText.set(true).verticalAlignment.set("Top");
            for (let i = 0; i < this.columnWidths.length; i++) {
                this.sheet.getRange("$"+columns[i]+"$1").columnWidth.set((this.columnWidths[i]*11+7)/2);
            }
            for (let i = 0; i < this.fontSizes.length; i++) {
                for (let j = 0; j < this.fontSizes[i].length; j++) {
                    if (this.fontSizes[i][j]!=null) this.sheet.getRange("$"+columns[j]+"$"+(i+1)).setFontSize(this.fontSizes[i][j]);
                }
            }
            for (let i = 0; i < this.bold.length; i++) {
                for (let j = 0; j < this.bold[i].length; j++) {
                    if (this.bold[i][j]) this.sheet.getRange("$"+columns[j]+"$"+(i+1)).bold.set(this.bold[i][j])
                }
            }
            this.sheet.protect(undefined,this.password);
            await this.context.sync();
        }
    }
}
console.clear();
const express = require("express");
var bodyParser = require('body-parser');
function generateUUID():string {
	var a = (new Date()).getTime();//Timestamp
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
		var b = Math.random() * 16;//random number between 0 and 16
		b = (a + b)%16 | 0;
		a = Math.floor(a/16);
		return (c === 'x' ? b : (b & 0x3 | 0x8)).toString(16);
	});
}
/**
 * @param {string} str
 * @param {string} char
 * @returns {number} how many times {str} contains {char}
 */
const StrNumIncludes:((str:string,char:string)=>number)=(str:string,char:string)=>str.split(char).length-1;
import * as fs from "fs";
type databaseType = {
    clients:{[client:string]:{[template:string]:string}},
    sheets:{[client:string]:{[template:string]:{[version:string]:string[]}}},
    stats:{[sheetId:string]:{created?:number,lastAccessed?:number}}
}
class Database {
    static DataBasePath:string=__dirname+"/clientData.json";
    static data:databaseType = Database.get();
    static get():databaseType {
        var database:any = "";
        if (fs.existsSync(Database.DataBasePath)) database = fs.readFileSync(Database.DataBasePath);
        if (database!="") database=JSON.parse(database);
        else database={};
        return database as databaseType;
    }
    static set(_data:databaseType) {
        this.data=_data;
    }
    static validateClient(client:string,template:string,version:string,sheet:string):boolean {
        var database:any = Database.get();
        if ( database.sheets                                     ==null) return false;
        if ( database.sheets[client]                             ==null) return false;
        if ( database.sheets[client][template]                   ==null) return false;
        if ( database.sheets[client][template][version]          ==null) return false;
        if (!database.sheets[client][template][version].includes(sheet)) return false;
        //set lastAccessed stat
        if ( database.stats        ==null) database.stats={};
        if ( database.stats[sheet] ==null) database.stats[sheet]={};
        database.stats[sheet].lastAccessed = (new Date()).getTime();
        //save and return
        Database.set(database);Database.save();
        return true;
    }
    static registerSheet(client:string,template:string,version:string):boolean {
        var database:any = Database.get();
        if ( database.clients                              ==null) return false;
        if ( database.clients[client]                      ==null) return false;
        if ( database.clients[client][template]            ==null) return false;
        if (!database.clients[client][template].includes(version)) return false;
        return true;
    }
    static addSheet(client:string,template:string,version:string,sheet:string):void {
        var database:any = Database.get();
        if (database.sheets                           ==null) database.sheets={};
        if (database.sheets[client]                   ==null) database.sheets[client]={};
        if (database.sheets[client][template]         ==null) database.sheets[client][template]={};
        if (database.sheets[client][template][version]==null) database.sheets[client][template][version]=[];
        database.sheets[client][template][version].push(sheet);
        //set created stat
        if ( database.stats        ==null) database.stats={};
        if ( database.stats[sheet] ==null) database.stats[sheet]={};
        database.stats[sheet].created = (new Date()).getTime();
        database.stats[sheet].client = client;
        database.stats[sheet].type = template;
        database.stats[sheet].version = version;
        Database.set(database);Database.save();
    }
    static save() {
        fs.writeFileSync(Database.DataBasePath,JSON.stringify(Database.data,null,3));
    }
}

const app:any = express();
app.use(express.urlencoded({ extended: true }));
app.use(bodyParser.text());
app.use((req:any, res:any, next:any)=>{ if ((typeof req.body)=="string" && req.body.length!=0) req.body = JSON.parse(req.body); next(); });
const staticDist:any = express.static(__dirname+"/dist/");
app.use("/"      ,(req:any, res:any, next:any)=>{
    if (req.method!="GET") {next();return;}
    const pathname:string = req._parsedUrl.pathname;
    const pathnameSplt:string[] = pathname.split("/");
    if (StrNumIncludes(pathname,"/")==1) {
        staticDist(req,res,next);
    } else if (pathname.startsWith("/template")) {
        const query:{[key:string]:string} = req._parsedUrl.query!=null?( Object.fromEntries( req._parsedUrl.query.split("&").map((el:string)=>el.split("=")) ) ):{};
        const template:string = pathnameSplt[2];
        const fileSplt:string[] = pathnameSplt[pathnameSplt.length-1].split(".");fileSplt.pop();//split on "."s and remove file extention
        const file:string = fileSplt.join(".");
        if (Database.validateClient(query.client,template,file,query.sheetId)) {
            staticDist(req,res,next);
        } else {//send error template
            next();
        }
    } else next();
});
app.use("/assets",express.static(__dirname+"/assets/"));
app.post("/register"  ,(req:any,res:any)=>{
    res.setHeader("Content-Type", "application/json");
    if(Database.registerSheet(req.body.client,req.body.template,req.body.version)) {
        const sheetId:string = generateUUID();
        res.end(JSON.stringify({sheetId}));
        console.log("Registered sheet: "+sheetId+" .");
        Database.addSheet(req.body.client,req.body.template,req.body.version,sheetId);
    } else res.end(JSON.stringify({}));
})
app.post("/validate"  ,(req:any,res:any)=>{
    res.setHeader("Content-Type", "application/json");
    if (Database.validateClient(req.body.client,req.body.template,req.body.version,req.body.sheetId)) {
        res.end(JSON.stringify({status:true}));
    } else res.end(JSON.stringify({status:false}));
})
app.get("/templates",(req:any,res:any)=>{
    res.setHeader("Content-Type", "application/json");
    const templates:string[] = Object.keys(Database.get()["clients"]["DEV"]);
    res.end(JSON.stringify(templates));
});
const PORT:number = 62085;
app.listen(PORT,()=>console.log("Excel Add-in server is running at: http://localhost:"+PORT));
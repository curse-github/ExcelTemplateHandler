"use strict";
exports.__esModule = true;
console.clear();
var express = require("express");
var bodyParser = require('body-parser');
function generateUUID() {
    var a = (new Date()).getTime(); //Timestamp
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var b = Math.random() * 16; //random number between 0 and 16
        b = (a + b) % 16 | 0;
        a = Math.floor(a / 16);
        return (c === 'x' ? b : (b & 0x3 | 0x8)).toString(16);
    });
}
/**
 * @param {string} str
 * @param {string} char
 * @returns {number} how many times {str} contains {char}
 */
var StrNumIncludes = function (str, char) { return str.split(char).length - 1; };
var fs = require("fs");
var Database = /** @class */ (function () {
    function Database() {
    }
    Database.get = function () {
        var database = "";
        if (fs.existsSync(Database.DataBasePath))
            database = fs.readFileSync(Database.DataBasePath);
        if (database != "")
            database = JSON.parse(database);
        else
            database = {};
        return database;
    };
    Database.set = function (_data) {
        this.data = _data;
    };
    Database.validateClient = function (client, template, version, sheet) {
        var database = Database.get();
        if (database.sheets == null)
            return false;
        if (database.sheets[client] == null)
            return false;
        if (database.sheets[client][template] == null)
            return false;
        if (database.sheets[client][template][version] == null)
            return false;
        if (!database.sheets[client][template][version].includes(sheet))
            return false;
        //set lastAccessed stat
        if (database.stats == null)
            database.stats = {};
        if (database.stats[sheet] == null)
            database.stats[sheet] = {};
        database.stats[sheet].lastAccessed = (new Date()).getTime();
        //save and return
        Database.set(database);
        Database.save();
        return true;
    };
    Database.registerSheet = function (client, template, version) {
        var database = Database.get();
        if (database.clients == null)
            return false;
        if (database.clients[client] == null)
            return false;
        if (database.clients[client][template] == null)
            return false;
        if (!database.clients[client][template].includes(version))
            return false;
        return true;
    };
    Database.addSheet = function (client, template, version, sheet) {
        var database = Database.get();
        if (database.sheets == null)
            database.sheets = {};
        if (database.sheets[client] == null)
            database.sheets[client] = {};
        if (database.sheets[client][template] == null)
            database.sheets[client][template] = {};
        if (database.sheets[client][template][version] == null)
            database.sheets[client][template][version] = [];
        database.sheets[client][template][version].push(sheet);
        //set created stat
        if (database.stats == null)
            database.stats = {};
        if (database.stats[sheet] == null)
            database.stats[sheet] = {};
        database.stats[sheet].created = (new Date()).getTime();
        database.stats[sheet].client = client;
        database.stats[sheet].type = template;
        database.stats[sheet].version = version;
        Database.set(database);
        Database.save();
    };
    Database.save = function () {
        fs.writeFileSync(Database.DataBasePath, JSON.stringify(Database.data, null, 3));
    };
    Database.DataBasePath = __dirname + "/clientData.json";
    Database.data = Database.get();
    return Database;
}());
var app = express();
app.use(express.urlencoded({ extended: true }));
app.use(bodyParser.text());
app.use(function (req, res, next) { if ((typeof req.body) == "string" && req.body.length != 0)
    req.body = JSON.parse(req.body); next(); });
var staticDist = express.static(__dirname + "/dist/");
app.use("/", function (req, res, next) {
    if (req.method != "GET") {
        next();
        return;
    }
    var pathname = req._parsedUrl.pathname;
    var pathnameSplt = pathname.split("/");
    if (StrNumIncludes(pathname, "/") == 1) {
        staticDist(req, res, next);
    }
    else if (pathname.startsWith("/template")) {
        var query = req._parsedUrl.query != null ? (Object.fromEntries(req._parsedUrl.query.split("&").map(function (el) { return el.split("="); }))) : {};
        var template = pathnameSplt[2];
        var fileSplt = pathnameSplt[pathnameSplt.length - 1].split(".");
        fileSplt.pop(); //split on "."s and remove file extention
        var file = fileSplt.join(".");
        if (Database.validateClient(query.client, template, file, query.sheetId)) {
            staticDist(req, res, next);
        }
        else { //send error template
            next();
        }
    }
    else
        next();
});
app.use("/assets", express.static(__dirname + "/assets/"));
app.post("/register", function (req, res) {
    res.setHeader("Content-Type", "application/json");
    if (Database.registerSheet(req.body.client, req.body.template, req.body.version)) {
        var sheetId = generateUUID();
        res.end(JSON.stringify({ sheetId: sheetId }));
        console.log("Registered sheet: " + sheetId + " .");
        Database.addSheet(req.body.client, req.body.template, req.body.version, sheetId);
    }
    else
        res.end(JSON.stringify({}));
});
app.post("/validate", function (req, res) {
    res.setHeader("Content-Type", "application/json");
    if (Database.validateClient(req.body.client, req.body.template, req.body.version, req.body.sheetId)) {
        res.end(JSON.stringify({ status: true }));
    }
    else
        res.end(JSON.stringify({ status: false }));
});
app.get("/templates", function (req, res) {
    res.setHeader("Content-Type", "application/json");
    var templates = Object.keys(Database.get()["clients"]["DEV"]);
    res.end(JSON.stringify(templates));
});
var PORT = 62085;
app.listen(PORT, function () { return console.log("Excel Add-in server is running at: http://localhost:" + PORT); });

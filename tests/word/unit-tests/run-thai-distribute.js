"use strict";

var fs = require("fs");
var path = require("path");
var vm = require("vm");

var root = path.resolve(__dirname, "../../..");
var sandbox = { console: console, window: {} };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(root, "word/Editor/Thai/thai-dict.js"), "utf8"), sandbox, { filename: "thai-dict.js" });
vm.runInContext(fs.readFileSync(path.join(root, "word/Editor/Thai/thai-break.js"), "utf8"), sandbox, { filename: "thai-break.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "thai-distribute.js"), "utf8"), sandbox, { filename: "thai-distribute.js" });

var failures = sandbox.window.AscTestThaiDistribute.run();
if (failures.length)
{
	console.error(failures.join("\n"));
	process.exit(1);
}
console.log("thai-distribute: ok");

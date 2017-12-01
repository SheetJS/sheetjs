var readFirstSheet = require("../").readFirstSheet;
console.log(readFirstSheet("a,b,c\n1,2,3\n4,5,6", {type:"binary"}));
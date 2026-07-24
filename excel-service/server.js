const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json({limit:"50mb"}));


const PORT = process.env.PORT || 3000;


// Download Folder
const downloadDir = path.join(__dirname,"downloads");

if(!fs.existsSync(downloadDir)){
    fs.mkdirSync(downloadDir);
}


// Generate Excel API
app.post("/api/excel/generate",(req,res)=>{

try{

const body = req.body;

console.log("Request Received");


// ================================
// Jira Bug Report
// ================================

if(
    body.testCases &&
    body.testCases.length &&
    body.testCases[0].key
){

console.log("Generating Jira Bug Report");


const bugs = body.testCases;

const workbook = XLSX.utils.book_new();


// -----------------------------
// Sheet 1 : Bug Details
// -----------------------------

const bugHeaders=[
    "Key",
    "Summary",
    "Status",
    "Reporter",
    "Priority",
    "Assignee",
    "Project"
];


const bugRows=[
    bugHeaders,

    ...bugs.map(bug=>[
        bug.key || "",
        bug.summary || "",
        bug.status || "",
        bug.reporter || "",
        bug.priority || "",
        bug.assignee || "",
        bug.project || ""
    ])
];


const bugSheet =
XLSX.utils.aoa_to_sheet(bugRows);



bugSheet["!cols"]=[
    {wch:15},
    {wch:70},
    {wch:25},
    {wch:25},
    {wch:15},
    {wch:25},
    {wch:15}
];


XLSX.utils.book_append_sheet(
    workbook,
    bugSheet,
    "Bug Details"
);



// -----------------------------
// Sheet 2 : Summary
// -----------------------------


const summaryRows=[

["Jira Bug Report Summary",""],

["Project",body.projectKey || ""],

["Total Bugs",body.totalBugs || bugs.length],

["",""],

["Status","Count"]

];


Object.entries(body.statusCount || {})
.forEach(([key,value])=>{

summaryRows.push([
    key,
    value
]);

});


summaryRows.push(["",""]);

summaryRows.push([
    "Priority",
    "Count"
]);


Object.entries(body.priorityCount || {})
.forEach(([key,value])=>{

summaryRows.push([
    key,
    value
]);

});



const summarySheet =
XLSX.utils.aoa_to_sheet(summaryRows);


summarySheet["!cols"]=[
    {wch:40},
    {wch:15}
];


XLSX.utils.book_append_sheet(
    workbook,
    summarySheet,
    "Summary"
);



// -----------------------------
// Sheet 3 : Date Trend
// -----------------------------


const trendRows=[

[
"Date",
"Bug Count"
]

];


Object.entries(body.dateTrend || {})
.forEach(([date,count])=>{

trendRows.push([
    date,
    count
]);

});



const trendSheet =
XLSX.utils.aoa_to_sheet(trendRows);


trendSheet["!cols"]=[
    {wch:20},
    {wch:15}
];



XLSX.utils.book_append_sheet(
    workbook,
    trendSheet,
    "Date Trend"
);



// File Name

const fileName =
`${body.issue_key || "Jira_Bug_Report"}.xlsx`;


const filePath =
path.join(downloadDir,fileName);



XLSX.writeFile(
    workbook,
    filePath
);



return res.json({

success:true,

issue_key:
body.issue_key,

totalRows:
bugs.length,

fileName:fileName,

downloadUrl:
`${req.protocol}://${req.get("host")}/downloads/${fileName}`

});


}



// ================================
// AI Test Case Excel
// ================================

else{


console.log("Generating Test Case Excel");


const data = body.testCases || [];


const headers=[

"Test Suite",
"Name",
"Importance",
"Summary",
"Precondition",
"Action",
"Expected Result",
"Actual Result",
"Test Step Status",
"Test Case Status"

];



const rows=[

headers,

...data.map(tc=>[

tc["Test Suite"] || "",

tc["Name"] || "",

tc["Importance"] || "",

tc["Summary"] || "",

tc["Precondition"] || "",

tc["Action"] || "",

tc["Expected Result"] || "",

tc["Actual Result"] || "",

tc["Test Step Status"] || "",

tc["Test Case Status"] || ""

])

];



const worksheet =
XLSX.utils.aoa_to_sheet(rows);



worksheet["!cols"]=[

{wch:25},
{wch:30},
{wch:15},
{wch:50},
{wch:50},
{wch:60},
{wch:60},
{wch:20},
{wch:20},
{wch:20}

];



const workbook =
XLSX.utils.book_new();



XLSX.utils.book_append_sheet(
 workbook,
 worksheet,
 "Test Cases"
);



const fileName =
`${body.issue_key || "TestCases"}.xlsx`;



const filePath =
path.join(downloadDir,fileName);



XLSX.writeFile(
 workbook,
 filePath
);



return res.json({

success:true,

fileName:fileName,

totalRows:data.length,

downloadUrl:
`${req.protocol}://${req.get("host")}/downloads/${fileName}`

});


}


}
catch(error){

console.log(error);

res.status(500).json({

error:error.message

});

}


});



// Static Download
app.use(
"/downloads",
express.static(downloadDir)
);



app.listen(PORT,()=>{

console.log(
`Excel Service running on port ${PORT}`
);

});
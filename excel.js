//functions
var emails = [];
var mobiles = [];

function checkEmail(email){
// console.log("email: "+email)
var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
var result=re.test(String(email).toLowerCase());
if(result==true){
    emails.push({"email ": email});
}}

function checkMobile(mobile){
// console.log("mobile: "+mobile)
var re1=/^\d{10}$/;
var re2=/^\d{11}$/;
var re3=/^\d{14}$/;

var result1=re1.test(mobile) && mobile.toString()[0]=='1';
var result2=re2.test(mobile) && mobile.toString()[0]=='0';
var result3=re3.test(mobile) && mobile.toString()[0]=='+';

if(result1==true){
   var newM1 = "\t+0020"+ mobile.toString();
   mobiles.push({'mobile':newM1});
} else if (result2==true) {
   newM2="\t+002"+mobile.toString();
   mobiles.push({'mobile':newM2});
} else if (result3==true) {
   newM3= mobile.toString();
   mobiles.push({'mobile':newM3});
  }
}

function procSheet(excelSheet){
  // console.log(excelSheet);
    for(var i =0;i<excelSheet.length;i++){
      if(excelSheet[i].email!==undefined) {checkEmail(excelSheet[i].email);}
      if(excelSheet[i].mobile!==undefined) {checkMobile(excelSheet[i].mobile);}
    }
}

//reading

var XLSX = require('xlsx')
var workbook = XLSX.readFile('All.xlsx');
var sheet_name_list = workbook.SheetNames;
console.log(sheet_name_list);
console.log(workbook.SheetNames.length);

for(var i=0;i<workbook.SheetNames.length;i++) {
    var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[i]]);
    procSheet(xlData);
}

console.log(emails);
console.log(mobiles);

// Writting in csv

const ObjectsToCsv = require('objects-to-csv')
const csvE = new ObjectsToCsv(emails);
csvE.toDisk('./emaillist.csv')

const csvM = new ObjectsToCsv(mobiles);
csvM.toDisk('./mobilelist.csv')

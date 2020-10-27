const Excel=require('exceljs');
var workbook = new Excel.Workbook();
const json={};
const labels=[];
(async function(){
 workbook = await workbook.xlsx.readFile('Financial Sample.xlsx');
 workbook.eachSheet(function(worksheet, sheetId) {
  if(!json[sheetId]) json[sheetId]={};
  worksheet.eachRow(function(row, rowNumber) {
   if(rowNumber==1)
    row.values.forEach(v=>labels.push(v));
   else {
    json[sheetId]['row:'+(rowNumber)]={};
    row.values.forEach((v,i)=>json[sheetId]['row:'+(rowNumber)][labels[i-1]]=v);
   }
  });
 });
 console.log(json);
})();

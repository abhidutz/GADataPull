/*

Enter the profile IDs of all the profiles you wanna fetch data.
IDs goes in an array. 

*/

 var allProfiles = [
       "ga:303000", 
       "ga:303001", 
       "ga:303002", 
       "ga:839003", 
       "ga:4134004",
       "ga:303005", 
       "ga:9597006"   
];

function getdates(){
   var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboards");
    sheet.getRange(2,3 , 1, 1).setValue(getLastNdays(396));
  sheet.getRange(2,4, 1, 1).setValue(getLastNdays(366));
  sheet.getRange(3,3 , 1, 1).setValue(getLastNdays(31));
  sheet.getRange(3,4, 1, 1).setValue(getLastNdays(1));
}

function populateYearlyData(){
populateData(allProfiles,'year');
}

function populatePast30(){
 
     var profiles = [
       "ga:303000", 
       "ga:303001", 
       "ga:303002", 
       "ga:839003", 
       "ga:4134004",
       "ga:303005", 
       "ga:9597006"   
  
];
   
   populateData(profiles,'past30');
   populateData(profiles,'past365');
  try{
  var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboards");
  sheet.getRange(2,3 , 1, 1).setValue(getLastNdays(396));
  sheet.getRange(2,4, 1, 1).setValue(getLastNdays(366));
  sheet.getRange(3,3 , 1, 1).setValue(getLastNdays(31));
  sheet.getRange(3,4, 1, 1).setValue(getLastNdays(1));
  }
  catch(error){}
  
}

function compareCustomData(){

populateData(allProfiles,'customCurrent');
  populateData(allProfiles,'customPast');

}


function populateData(profiles, type){
   for (i = 0; i < profiles.length; ++i) 
  {
  
    var results = getReportDataForProfile(profiles[i],type);       
    outputToSpreadsheet(results,i,type); 
    if(type !='daily' ){ 
    for(k = 0; k < 3; ++k) {    
    var deviceTypeResults = getDeviceData(profiles[i],k,type);
      outDeviceDataToSheet(deviceTypeResults,k,i,type);
       }
    }
  }

}


function populateDailyData() {
   
     var profiles = [
       "ga:303000", 
       "ga:303001", 
       "ga:303002", 
       "ga:839003", 
       "ga:4134004",
       "ga:303005", 
       "ga:9597006"     
];
  
   populateData(profiles,'daily');
 
}



function getReportDataForProfile(Profile,type) {
  var startDate,endDate,optArgs;
   

  switch(type){
   case 'past30':
       
       startDate = getLastNdays(31); 
       endDate = getLastNdays(1);    
       optArgs = {'start-index': '1','max-results': '10000'};
   break;

 case 'past365':
    
       startDate = getLastNdays(396);
       endDate = getLastNdays(366);  
       optArgs = {'start-index': '1','max-results': '10000'};
   break;
      
    case 'daily':
       startDate = getLastNdays(2);   // yesterday
       endDate = getLastNdays(2);  
       optArgs = {'dimensions': 'ga:year,ga:month,ga:week,ga:date','start-index': '1','max-results': '10000'};
   break;
      
case 'year':
       startDate="2008-01-01";   // yesterday
       endDate = getLastNdays(1);      // yesterday
       optArgs = {'dimensions': 'ga:year','start-index': '1','max-results': '10000'};
   break;
      
      
case 'customCurrent':
      
       var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Custom Comparison");     
      startDate=sheet.getRange(2,2 , 1, 1).getValue();
      startDate=Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');
     endDate=sheet.getRange(2,3, 1, 1).getValue();
      endDate=Utilities.formatDate(endDate, 'GMT', 'yyyy-MM-dd');
       optArgs = {'start-index': '1','max-results': '10000'};
   break;
      
      
      
case 'customPast':
        var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Custom Comparison");
    startDate=sheet.getRange(3,2 , 1, 1).getValue();
      startDate=Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');
      endDate=sheet.getRange(3,3, 1, 1).getValue();
      endDate=Utilities.formatDate(endDate, 'GMT', 'yyyy-MM-dd');
       optArgs = {'start-index': '1','max-results': '10000'};
   break;
      
  }
  
  
  // Make a request to the API.
  var results = Analytics.Data.Ga.get(
      Profile,                  // Table id (format ga:xxxxxx).
      startDate,                // Start-date (format yyyy-MM-dd).
      endDate,                  // End-date (format yyyy-MM-dd).
      'ga:visitors,ga:pageviews,ga:visitBounceRate,ga:adsenseRevenue', // Comma seperated list of metrics.
      optArgs);

  
  if (results.getRows()) {
    return results;
  } else {
    throw new Error('No profiles found');
  }
}

function getDeviceData (Profile,a,type) {
 var startDate,endDate,optArgs;
 
   var filterData="";
  switch(a)
  {
    case 0:
      filterData = "ga:deviceCategory=@desktop";
      break;
  case 1:
      filterData = "ga:deviceCategory=@mobile";
      break;
      case 2:
      filterData = "ga:deviceCategory=@tablet";
      break;
  }

  
  
  switch(type){
   case 'past30':
       startDate = getLastNdays(31);   // yesterday
       endDate = getLastNdays(1);      // yesterday
       optArgs = {'start-index': '1','max-results': '10000','filters' : filterData};
   break;

 case 'past365':
       startDate = getLastNdays(396);   // month ago last year
       endDate = getLastNdays(366);      // same date last year
       optArgs = {'start-index': '1','max-results': '10000','filters' : filterData};
   break;
      
   case 'customCurrent':
       var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Custom Comparison");
    startDate=sheet.getRange(2,2 , 1, 1).getValue();
      startDate=Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');
      endDate=sheet.getRange(2,3, 1, 1).getValue();
      endDate=Utilities.formatDate(endDate, 'GMT', 'yyyy-MM-dd');
       optArgs = {'start-index': '1','max-results': '10000','filters' : filterData};
   break;

 case 'customPast':
       var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Custom Comparison");
    startDate=sheet.getRange(3,2 , 1, 1).getValue();
      startDate=Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');
      endDate=sheet.getRange(3,3, 1, 1).getValue();
      endDate=Utilities.formatDate(endDate, 'GMT', 'yyyy-MM-dd');
       optArgs = {'start-index': '1','max-results': '10000','filters' : filterData};
   break;
  
      
    case 'daily':
       startDate = getLastNdays(2); 
       endDate = getLastNdays(2);      // yesterday
       optArgs = {'dimensions': 'ga:date','start-index': '1','max-results': '10000'};
   break;
       case 'year':
       startDate="2008-01-01";   // yesterday
       endDate = getLastNdays(1);      // yesterday
      
       optArgs = {'dimensions': 'ga:year','start-index': '1','max-results': '10000','filters' : filterData};
   break;
  }
  
  
  // Make a request to the API.
  var results = Analytics.Data.Ga.get(
      Profile,                  // Table id (format ga:xxxxxx).
      startDate,                // Start-date (format yyyy-MM-dd).
      endDate,                  // End-date (format yyyy-MM-dd).
      'ga:visitors', // Comma seperated list of metrics.
      optArgs);

  
  if (results.getRows()) {
    return results;
  } else {
    throw new Error('No profiles found');
  }
}

function getLastNdays(nDaysAgo) {
  var today = new Date(); 
  var before = new Date();
  before.setDate(today.getDate() - nDaysAgo);
  return Utilities.formatDate(before, 'GMT', 'yyyy-MM-dd');
}




function outDeviceDataToSheet(results,typeIndex,sheetIndex,type) 
{
  var sName="";
  switch(type){
    case 'daily':
  switch(sheetIndex){
    case 0: 
      sName="Private";
      break;      
      case 1: 
      sName="Public";
      break;      
      case 2: 
      sName="Boarding";
      break;
      case 3: 
      sName="Community";
      break;
      case 4: 
      sName="Poker";
      break;
      case 5: 
      sName="Fish";
      break;
      case 6: 
      sName="Seo";
      break;
  }
      break;
    case 'past30':
      sName="Past30";
      break;
      case 'past365':
      sName="Past30";
      break;
      
      
       case 'customCurrent':
      sName="Custom";
      break;
      
      case 'customPast':
      sName="Custom";
      break;
   
        case 'year':
      sName="Year";
      break;
  }
    var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sName);

  
  var writeIndex = sheet.getLastRow();
  switch(type){
   case 'past30':
     sheet.getRange(sheetIndex*3+2, typeIndex+6, results.getRows().length, 1)
      .setValues(results.getRows());

   break;

 case 'past365':
      sheet.getRange(sheetIndex*3+3, typeIndex+6, results.getRows().length, 1)
      .setValues(results.getRows());

   break;
        case 'customCurrent':
     sheet.getRange(sheetIndex*3+2, typeIndex+6, results.getRows().length, 1)
      .setValues(results.getRows());

   break;

 case 'customPast':
      sheet.getRange(sheetIndex*3+3, typeIndex+6, results.getRows().length, 1)
      .setValues(results.getRows());

   break;
      
     case 'daily':
    
  sheet.getRange(writeIndex,typeIndex*2+9 , results.getRows().length, 2)
      .setValues(results.getRows());
      break;
     case 'year':
    
  sheet.getRange(sheetIndex*11+2,typeIndex*2+6 , results.getRows().length, 2)
      .setValues(results.getRows());
      break;
      
  }
    
}




function outputToSpreadsheet(results,i,type) {
  
  var sName="";
  switch(type){
    case 'daily':
  switch(i){
    case 0: 
      sName="Private";
      break;      
      case 1: 
      sName="Public";
      break;      
      case 2: 
      sName="Boarding";
      break;
      case 3: 
      sName="Community";
      break;
      case 4: 
      sName="Poker";
      break;
      case 5: 
      sName="Fish";
      break;
      case 6: 
      sName="Seo";
      break;
  }
      break;
    case 'past30':
      sName="Past30";
      break;
      case 'past365':
      sName="Past30";
      break;
      
       case 'customCurrent':
      sName="Custom";
      break;
      case 'customPast':
      sName="Custom";
      break;
      case 'year':
      sName="Year";
      break;
  }
    var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sName);

   
   switch(type){
   case 'past30':
     sheet.getRange(i*3+2, 2, results.getRows().length, 4)
      .setValues(results.getRows());

   break;

 case 'past365':
      sheet.getRange(i*3+3, 2, results.getRows().length, 4)
      .setValues(results.getRows());

   break;
      
        case 'customCurrent':
     sheet.getRange(i*3+2, 2, results.getRows().length, 4)
      .setValues(results.getRows());

   break;

 case 'customPast':
      sheet.getRange(i*3+3, 2, results.getRows().length, 4)
      .setValues(results.getRows());

   break;
       
       
     case 'daily':
       for(var d=0;d<results.getRows().length;d++){
  var val = results.rows[d][3];
    results.rows[d][3]=val.substring(0,4)+"-"+val.substring(4,6)+"-"+val.substring(6,8);
  }
  var writeIndex = sheet.getLastRow()+1;
 
  sheet.getRange(writeIndex, 1, results.getRows().length, 8)
      .setValues(results.getRows());
   break;
 
  
  case 'year':
  var yearWriteIndex;
   switch(i){
    case 0: 
      yearWriteIndex=2;
      break;      
      case 1: 
      yearWriteIndex=13;
      break;      
      case 2: 
      yearWriteIndex=24;
      break;
      case 3: 
      yearWriteIndex=35;
      break;
      case 4: 
      yearWriteIndex=46;
      break;
      case 5: 
      yearWriteIndex=57;
      break;
      case 6: 
      yearWriteIndex=68;
      break;
  }
   
  sheet.getRange(yearWriteIndex, 1, results.getRows().length, 5)
      .setValues(results.getRows());
   break;
  }
  }
      

//Function to export dashboard on excel to spreadsheet



function exportDashboard() {
  try{
  var originalSpreadsheet = SpreadsheetApp.getActive();
  
  var message = "Please see attached, Monthly Dashboard."; 
  var subject = "Monthly Website Dashboard";
  var contacts = originalSpreadsheet.getSheetByName("ExportEmailList");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(2, 1, numRows, 1).getValues();
 var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet = originalSpreadsheet.getSheetByName("Dashboards");
  var exportSheet = originalSpreadsheet.getSheetByName("ExportSheet");
   sheet.getRange("A1:P37").copyTo(exportSheet.getRange("A1:P37"),{contentsOnly:true});
  exportSheet.copyTo(newSpreadsheet);
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();

  var pdf = DocsList.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:'Monthly Dashboard.pdf',content:pdf, mimeType:'application/pdf'};
  MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
  DocsList.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}
  catch(error){}
}

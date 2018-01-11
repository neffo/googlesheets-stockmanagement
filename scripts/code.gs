function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Transaction Menu')
  .addItem('Reset/Clear WorkOrder', 'clearWorkOrderRange')
  .addItem('Reset/Clear PickList', 'clearOrderRange')
  .addSeparator()
  .addItem('Submit & Print WorkOrder', 'addWorkOrderToTransactions')
  .addItem('Submit & Print PickList', 'addOrderPicklistToTransactions')
  //.addSeparator()
  //.addItem('Export PDF/Print this Sheet', 'exportSheet')
  //.addItem('Export PickList', 'exportPicklist')
  .addToUi();
}


function exportWorkOrder () {
  exportSheet('Prod Worksheet','name@company.com','Exported Workorder');
}

function exportPicklist () {
  exportSheet('Order Picklist','name@company.com','Exported Picklist');
}

function exportSheet() {
/*  var message = "Please see attached"; // Could make it a pop-up perhaps, but out of wine today
  var key = 
  
  //var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:'Exported-'+sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};
 
  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true); */
  
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    var gid = sheet.getSheetId();

    var pdfOpts = '&size=A4&fzr=false&portrait=false&fitw=true&gridlines=false&comments=false&printcomments=false&printnotes=false&notes=false&printtitle=false&sheetnames=false&pagenum=1,2&attachment=false&gid='+gid;

    var row2 = sheet.getMaxRows()-1;
    var printRange = '&c1=0' + '&r1=0' + '&c2=8' + '&r2='+row2; // B2:APn
    var url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf' + pdfOpts + printRange;

    var app = UiApp.createApplication().setWidth(200).setHeight(50);
    app.setTitle('Print this sheet');

    var link = app.createAnchor('Download PDF to print', url).setTarget('_new');

    app.add(link);

    ss.show(app);

}

function display_link_to_file (file, type) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = file.getUrl();
    var app = UiApp.createApplication().setWidth(200).setHeight(50);
    app.setTitle('Print this '+type);

    var link = app.createAnchor('Link to download or print PDF', url).setTarget('_new');

    app.add(link);

    ss.show(app);
}

function downloadurl(format, col) {
  
  if (format == '')
    format = 'pdf';
  //SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var gid = sheet.getSheetId();

  var pdfOpts = '&size=A4&fzr=false&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&attachment=false&gid='+gid;

  if (col == 0)
    col = 8;
  var row2 = sheet.getMaxRows()-1;
  var printRange = '&c1=0' + '&r1=0' + '&c2='+col + '&r2='+row2; // B2:APn
  var url = ss.getUrl().replace(/edit$/, '') + 'export?format='+format + pdfOpts + printRange;
  
  return url;
}

/*function exportSheet(sheetName, emailTo, subject) {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive();
  
  // Set the message to attach to the email.
  var message = "Please see attached"; // Could make it a pop-up perhaps, but out of wine today
  
  // Get Project Name from Cell A1
  //var projectname = originalSpreadsheet.getRange("A1:A1").getValues(); 
  // Get Reporting Period from Cell B3
  //var period = originalSpreadsheet.getRange("B3:B3").getValues(); 
  // Construct the Subject Line
  //var subject = projectname + " - Weekly Status Sheet - " + period;
 
      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  //var contacts = originalSpreadsheet.getSheetByName("Contacts");
  //var numRows = contacts.getLastRow();
  //var emailTo = contacts.getRange(2, 2, numRows, 1).getValues();
 
  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack
 
  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); //previous was getActiveSheet();
  var projectname = SpreadsheetApp.getActiveSpreadsheet();
  sheet = originalSpreadsheet.getActiveSheet();
  var tempsheet = projectname.insertSheet("temp sheet");
  
  sheet.getRange("B1:H37").copyTo(tempsheet.getRange("B1:H37"), {contentsOnly:true});
  sheet.getRange("B1:H37").copyTo(tempsheet.getRange("B1:H37"), {formatsOnly:true});
  
 
  
  tempsheet.copyTo(newSpreadsheet);
  
    
  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  //var expRange = newSpreadsheet.getActiveSheet().getRange("A1:G31");
  //sheet.getRange(1, 1, 37, 8).copyTo(expRange, {contentsOnly:true});
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:'Exported-'+sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};
 
  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject, message, {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true); 
}*/

function clearOrderRange() { 
  var sheet = SpreadsheetApp.getActive().getSheetByName('Order Picklist');
  sheet.getRange('B10:D29').clearContent();
  sheet.getRange('D6').setValue("=now()"); // date, set to current date
  sheet.getRange('D5').clearContent(); // order number
  sheet.getRange('D7').clearContent(); // description
  sheet.getRange('F5').clearContent(); // username
  
  for (var row=0; row<20; row++) {
    sheet.getRange(10+row,5).setValue("="+"L"+(10+row));
  }
}

function clearWorkOrderRange() { 
  var sheet = SpreadsheetApp.getActive().getSheetByName('Prod Workorder');

  sheet.getRange('E4').clearContent(); // product code
  sheet.getRange('E5').setValue(1); // order quantity
  sheet.getRange('E6').setValue(""); // min expiry
  sheet.getRange('E9').clearContent(); // username
  sheet.getRange('E10').setValue("WORKORDER"); // description
  sheet.getRange('E11').setValue("=now()"); // date, set to current date
  //sheet.getRange('D12').clearContent(); // description
  sheet.getRange('E13').setValue("=\"Manufacture/Dispense/Packing of \"&E3&\" (\"&E4&\") \"&E12&\" x \"&E5");
  sheet.getRange('E12').setValue("=J11");
  
  for (var row=0; row<17; row++) {
    sheet.getRange(16+row,5).setValue("="+"K"+(16+row));
  }
  
}

function addWorkOrderToTransactions() {
  // showAlertOk('ERROR','This is just a placeholder.');
  //return 0;
  
  var shWorkOrder = SpreadsheetApp.getActive().getSheetByName('Prod Workorder');
  var shTrans = SpreadsheetApp.getActive().getSheetByName('Transactions');
  var startRow; // = shTrans.getRange('O2:O1000').getNumRows();
  var validData = validateWorkOrder();
  
  startRow = getFirstEmptyRow()+1;
  
  // just debugging stuff
  //shWorkOrder.getRange('M2').setValue(startRow);
  //shWorkOrder.getRange('M3').setValue(validData);
  
  //return 0;
  
  //if (validData === false && showAlertQuestion('Confirm', 'Print this work order first?')==1) {
  //  exportSheet();
  //  return 0;
  //}
  
  if (validData === false && showAlertQuestion('Confirm', 'Add workorder details to transaction log?')==1 ) {
    var orderCode;
    var orderDate;
    var orderOp;
    var orderDesc;
    var orderProd;
    var orderLot;
    var orderCount;
    
    var itemCode;
    var itemCount;
    var itemLot;
    var itemPrice;

    orderCode = shWorkOrder.getRange('E10').getValue();
    orderDate = shWorkOrder.getRange('E11').getValue();
    orderOp = shWorkOrder.getRange('E9').getValue();
    orderDesc = shWorkOrder.getRange('E13').getValue();
    orderProd = shWorkOrder.getRange('E4').getValue();
    orderLot = shWorkOrder.getRange('E12').getValue();
    orderCount = shWorkOrder.getRange('E5').getValue();
    
    var date = Utilities.formatDate(new Date(), "AEST", "yyyyMMdd");
    var filename = date+"-Workorder "+orderProd+" "+orderLot+" x "+orderCount+" ("+orderOp+")";
    var file = saveSpreadsheetAsPDF("Prod Workorder", filename);
    
    var range = shWorkOrder.getRange('B16:E37');
    var values = range.getValues();
    
    for (var row=0; row<values.length; row++) {
      if (values[row][0]!="" && values[row][3]!="") {
        itemCode = values[row][0];
        itemCount = values[row][2];
        if (itemCode != 'MISC') 
          itemLot = values[row][3];
        else
          itemLot = 'MISC';
        
        Logger.log('Trans: '+startRow+'->'+orderDate+orderOp+itemCode+itemLot+orderCode+itemCount+itemPrice+orderDesc);
        
        shTrans.getRange(startRow,1).setValue(orderDate);
        shTrans.getRange(startRow,2).setValue(orderOp);
        shTrans.getRange(startRow,3).setValue(itemCode);
        shTrans.getRange(startRow,4).setValue(itemLot);
        shTrans.getRange(startRow,5).setValue(orderCode);
        shTrans.getRange(startRow,6).setValue(-itemCount); // negative because we are issuing (removing from stock)
        //shTrans.getRange(startRow,7).setValue(itemPrice);
        shTrans.getRange(startRow,11).setValue(orderDesc);
        
        startRow++;
      }
    }
    
    // add transaction for the finished product
    shTrans.getRange(startRow,1).setValue(orderDate);
    shTrans.getRange(startRow,2).setValue(orderOp);
    shTrans.getRange(startRow,3).setValue(orderProd);
    shTrans.getRange(startRow,4).setValue(orderLot);
    shTrans.getRange(startRow,5).setValue(orderCode);
    shTrans.getRange(startRow,6).setValue(orderCount); // positive because we are creating (removing from stock)
    //shTrans.getRange(startRow,7).setValue(itemPrice);
    shTrans.getRange(startRow,11).setValue(orderDesc);
    
    // remove retention amount
    var retention = shWorkOrder.getRange('G5').getValue();
    if (retention != 0 ) {
      startRow++;
      shTrans.getRange(startRow,1).setValue(orderDate);
      shTrans.getRange(startRow,2).setValue(orderOp);
      shTrans.getRange(startRow,3).setValue(orderProd);
      shTrans.getRange(startRow,4).setValue(orderLot);
      shTrans.getRange(startRow,5).setValue("RETENTION");
      shTrans.getRange(startRow,6).setValue(-retention); // negative because we are issuing (removing from stock)
      shTrans.getRange(startRow,11).setValue("Held for retention");
    }
    
    
    if (showAlertQuestion('Confirm', 'Clear order sheet?')==1)
      clearWorkOrderRange();
    
    display_link_to_file (file, "Work Order");
  }
}

function addOrderPicklistToTransactions() {
  var shPicklist = SpreadsheetApp.getActive().getSheetByName('Order Picklist');
  var shTrans = SpreadsheetApp.getActive().getSheetByName('Transactions');
  var startRow; // = shTrans.getRange('O2:O1000').getNumRows();
  var validData = validatePicklist();
  
  startRow = getFirstEmptyRow()+1;
  
  // just debugging stuff
  //shPicklist.getRange('J12').setValue(startRow);
  //shPicklist.getRange('J13').setValue(validData);
  
  //if (validData === false && showAlertQuestion('Confirm', 'Print this pick list?')==1) {
  //  exportSheet();
  //  return;
  //}
  
  if (validData === false && showAlertQuestion('Confirm', 'Add order details to transaction log?')==1 ) {
    var orderCode;
    var orderDate;
    var orderOp;
    var orderDesc;
    
    var itemCode;
    var itemCount;
    var itemLot;
    var itemPrice;

    orderCode = shPicklist.getRange('D5').getValue();
    orderDate = shPicklist.getRange('D6').getValue();
    orderOp = shPicklist.getRange('F5').getValue();
    orderDesc = shPicklist.getRange('D7').getValue();
    
    var date = Utilities.formatDate(new Date(), "AEST", "yyyyMMdd");
    var filename = date+"-Picklist "+orderCode+" ("+orderOp+")";
    var file = saveSpreadsheetAsPDF("Order Picklist", filename);
    
    var range = shPicklist.getRange('B10:E29');
    var values = range.getValues();
    
    for (var row=0; row<values.length; row++) {
      if (values[row][0]!="") {
        itemCode = values[row][0];
        itemCount = values[row][1];
        itemPrice = values[row][2];
        if (itemCode != 'MISC') 
          itemLot = values[row][3];
        else
          itemLot = 'MISC';
        
        Logger.log('Trans: '+startRow+'->'+orderDate+orderOp+itemCode+itemLot+orderCode+itemCount+itemPrice+orderDesc);
        
        shTrans.getRange(startRow,1).setValue(orderDate);
        shTrans.getRange(startRow,2).setValue(orderOp);
        shTrans.getRange(startRow,3).setValue(itemCode);
        shTrans.getRange(startRow,4).setValue(itemLot);
        shTrans.getRange(startRow,5).setValue(orderCode);
        shTrans.getRange(startRow,6).setValue(-itemCount); // negative because we are issuing (removing from stock)
        shTrans.getRange(startRow,7).setValue(itemPrice);
        shTrans.getRange(startRow,11).setValue(orderDesc);
        
        startRow++;
      }
    }
    
    if (showAlertQuestion('Confirm', 'Clear order sheet?')==1)
      clearOrderRange();
    
    display_link_to_file (file, "Pick list");
  }
}

function getFirstEmptyRow(skiprows) {
  var shTrans = SpreadsheetApp.getActive().getSheetByName('Transactions');
  var range = shTrans.getRange('O2:O5000');
  var values = range.getValues();
  var row = skiprows != null ? skiprows : 0;
  for (; row<values.length; row++) {
    if (!values[row].join("")) break;
  }
  return (row+1);
}

// returns false if everything is ok, true on error
function validatePicklist() {
  var shPicklist = SpreadsheetApp.getActive().getSheetByName('Order Picklist');
  var ret = true;
  // all the top fields completed yeah?
  
  if ( shPicklist.getRange('D5')=="" ||
    shPicklist.getRange('D6')=="" ||
      shPicklist.getRange('D7')=="" ||
        shPicklist.getRange('F5')=="" ) {
    showAlertOk('ERROR','Please complete order data fields.');
    return true;
  }
  
  var range = shPicklist.getRange('B10:E29');
  var values = range.getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if (values[row][0]!="") {
      Logger.log('Row ' + row + ' 1st cell = ' + values[row][0] + ' RET = '+ret);
      ret = ret && (values[row][1]!="" && values[row][1] > -10000 && values[row][1] < 10000); // is empty OR isn't valid number
      // MLC160426: modified the check above, we don't care about negative numbers anymore
      Logger.log('Row ' + row + ' 2nd cell = ' + values[row][1] + ' RET = '+ret);
      ret = ret && (values[row][2]=="" || values[row][2] > 0); // isn't empty and is negative 
      Logger.log('Row ' + row + ' 3rd cell = ' + values[row][2] + ' RET = '+ret);
      ret = ret && (values[row][0] == 'MISC' || values[row][3] != '#N/A'); // a lot is available
      Logger.log('Row ' + row + ' 4th cell = '+ values[row][3]+ ' RET = '+ret);
    }
    if (!ret)
      break;
  }
  
  if (!ret) {
    showAlertOk('ERROR','Invalid line item data at row '+(row+1)+'. Please confirm pass status and expiry.');
  }
  
  return (!ret);
}

// returns false if everything is ok, true on error
function validateWorkOrder() {
  var shWorkOrder = SpreadsheetApp.getActive().getSheetByName('Prod Workorder');
  var ret = true;
  
  if (shWorkOrder == null) {
    showAlertOk('ERROR','Invalid production worksheet name? It should be "Prod Workorder".');
    return true; 
  }
    
  // all the top fields completed yeah? (FIXME: this is broken)
  if ( shWorkOrder.getRange('D4')=="" ||
    shWorkOrder.getRange('D5')=="" ||
      shWorkOrder.getRange('D9').length == 0 ||
        shWorkOrder.getRange('D10')=="" ||
          shWorkOrder.getRange('D11')=="" ||
            shWorkOrder.getRange('D12')=="" ) {
   
    showAlertOk('ERROR','Please all complete order data fields.');
    return true;
  }
  
  var range = shWorkOrder.getRange('B15:N36');
  var values = range.getValues();
  var row = 0;
  var errors = new Array();
  for (var row=0; row<values.length; row++) {
    //Logger.log('Row ' + row + ' code = '+ values[row][0].substr(0,2)+' is a product.');
    if ( /*(typeof values[row][0] !== 'number') && values[row][0].substr(0,2)=="A-" */ values[row][12] == "Current" ) { // is this a listed product (substr breaks if value is a number so check this first)
      Logger.log('Row ' + row + ' code = '+ values[row][0]+' is a product.');
      
      ret = ret && (values[row][1]!="" && typeof values[row][1] == 'number' /*values[row][1] > 0 && values[row][1] < 10000*/); // is empty OR isn't valid number
      if (!ret)
        errors[errors.length] = "Row "+row+ " Cell 2 is empty or isn't a number. ("+values[row][1]+","+(typeof values[row][1])+") ";
      ret = ret && values[row][3] != '#N/A' && values[row][3] != '' ; // the lot has an expiry assigned
      //if (!ret)
      //  errors[errors.length] = "Row "+row+ " Cell 4 the lot has no expiry assigned. ";
      ret = ret && values[row][3] != '#N/A'; // the lot has an expiry assigned
      if (!ret)
        errors[errors.length] = "Row "+row+ " Cell 4 the lot has no expiry assigned. "
      ret = ret && values[row][4] == 'PASS'; // a lot is passed
      if (!ret)
        errors[errors.length] = "Row "+row+ " Cell 4 the lot is not passed. "
      
        


    } else { // else not a product, we don't care about this
      Logger.log('Row ' + row + ' code = '+ values[row][0]+' is NOT a product.');
      ret = true; // because it's not a product it won't be added to anything so we can safely ignore it
    }
    if (!ret)
      break;
    Logger.log('Row ' + row + ' is ok.');
  }
  
  if (!ret) {
    showAlertOk('ERROR','Invalid line item data at row '+(row+1)+'. Please confirm pass status and expiry. [First error: '+errors[0]+']');
  }
  
  return (!ret);
}

function showAlertQuestion(title, message) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     title,
     message,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert('Confirmation received.');
    return 1;
  } else {
    // User clicked "No" or X in the title bar.
    //ui.alert('Permission denied.');
    return 0;
  }
}

function showAlertOk(title, message) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     title,
     message,
      ui.ButtonSet.OK);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert('Confirmation received.');
    return 1;
  } else {
    // User clicked "No" or X in the title bar.
    //ui.alert('Permission denied.');
    return 0;
  }
  
}

function testFindEmptyRow() {
  var d = new Date();
  var timeStart = d.getTime();
  var emptyRow = getFirstEmptyRow();
  var timeEnd = d.getTime();
  
  //showAlertOk("Processing Time", "Elapsed: "+(timeEnd - timeStart)+"");
  Logger.log("Processing Time Elapsed: "+(timeEnd - timeStart));
}

/* Send Spreadsheet in an email as PDF, automatically */
// stolen from: https://ctrlq.org/code/19869-email-google-spreadsheets-pdf

function saveSpreadsheetAsPDF(sheetname, filename) {
  
  if (sheetname == null)
    sheetname = "Prod Workorder";
  
  if (filename == null)
    filename = "PSS Export";
  
  // Send the PDF of the spreadsheet to this email address
  var email = "name@company.com"; 
  
  // Get the currently active spreadsheet URL (link)
  // Or use SpreadsheetApp.openByUrl("<<SPREADSHEET URL>>");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetname);

  // Subject of email message
  var subject = "PDF generated from spreadsheet " + ss.getName(); 

  // Email Body can  be HTML too with your logo image - see ctrlq.org/html-mail
  var body = "Exported PDF link below.";
  
  // Base URL
  var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
  
  /* Specify PDF export parameters
  From: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579
  */
  
  var col = 9;
  var row2 = sheet.getMaxRows()-1;

  var printRange = '&c1=0' + '&r1=0' + '&c2='+col + '&r2='+row2; // B2:APn

  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&size=a4'                       // paper size legal / letter / A4
  + '&portrait=false'                    // orientation, false for landscape
  + '&fitw=true'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + printRange                           // our specific columns and rows
  + '&gid=';                             // the sheet's Id
  
  
  var token = ScriptApp.getOAuthToken();
  var sheets = ss.getSheets(); 
  
  //make an empty array to hold your fetched blobs  
  var blobs = [];
  
  //for (var i=0; i<sheets.length; i++) {
  if (1) {
    
    // Convert individual worksheets to PDF
    var response = UrlFetchApp.fetch(url + url_ext + sheet.getSheetId(), {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });
    
    //convert the response to a blob and store in our array
    blob = response.getBlob().setName(filename + '.pdf');
    
  }
  
  //create new blob that is a zip file containing our blob array
  //var zipBlob = Utilities.zip(blobs).setName(ss.getName() + '.zip'); 
  
  //optional: save the file to the root folder of Google Drive
  //DriveApp.createFile(blob);
  
  var folders = DriveApp.getFoldersByName("ExportedWorkorders");
  
  if (!folders.hasNext()) {
    Logger.log ("Folder not found!");
    DriveApp.createFolder("ExportedWorkorders");
    folders = DriveApp.getFoldersByName("ExportedWorkorders");
    //return;
  }
  
  var toFolder = folders.next();
  
  var file = toFolder.createFile(blob);
  
  // Define the scope
  Logger.log("Storage Space used: " + DriveApp.getStorageUsed());
  
  // If allowed to send emails, send the email with the PDF attachment
  /*if (MailApp.getRemainingDailyQuota() > 0) 
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[blob]     
    });  */
  
  return file;
}

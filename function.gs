function processNewAccountForm(formObject) {  
  
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  const timestamp = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(userEmail);
  const brandName = formObject.newBrand;
  //  Check for exisiting brand name. Change to search for all new brand then compare name? filter? Make faster.
  const brandColumn = mySheet.getRange('D:D');
  const brandValues = brandColumn.getValues();
  let i = 1;
  for(i=1; i < brandValues.length; i++) {
    if(brandValues[i].toString().toLowerCase().trim() == brandName.toString().toLowerCase().trim() && mySheet.getRange(i+1,5).getValue() == 'New Brand'){       
      //   add error message if name already exisit
      throw new Error("Brand already exists.");        
    }       
  };  
  //    Create folder and PDF with build instructions
  const parentFolder = DriveApp.getFolderById("folderid");  
  const newFolder = parentFolder.createFolder(brandName);
  const docFile = newFolder.createFile('New Build.pdf', 
                                       'ššæš®š»š± š”š®šŗš²:  ' + formObject.newBrand + 
                                       '\nšš¼šš»ššæš:  ' + formObject.country +
                                       '\nš¦šš®šš²:  ' + formObject.state + 
                                       '\nšš¶šš:  ' + formObject.city + 
                                       '\nš”ššŗšÆš²šæ š¼š³ šš¼š°š®šš¶š¼š»š:  ' + formObject.locations + 
                                       '\nšššš¶š»š²šš š§šš½š²:  ' + formObject.businessType +
                                       '\nšš»šš²š“šæš®šš¶š¼š»:  ' + formObject.integration + 
                                       '\nšš¼šŗš½š¹š²šš¶šš:  ' + formObject.complexity + 
                                       '\nš£šæš¶š¼šæš¶šš š„š®šš¶š»š“:  ' + formObject.priority +
                                       '\nšš¶š¹š² š¦š¼ššæš°š²:  ' + formObject.menuSource +
                                       '\nšš¼ šš¼š š»š²š²š± š®š» š¶šš²šŗ šµš¶š±š±š²š»/š±š¶šš®šÆš¹š²š±?:  ' + formObject.disableItemOption +
                                       '\nšŖšµš¶š°šµ š¶šš²šŗ(š) ššµš¼šš¹š± šÆš² š±š¶šš®šÆš¹š²š±?:  ' + formObject.itemDisable +
                                       '\nšš¼ šš¼š š»š²š²š± š® šŗš¼š±š¶š³š¶š²šæ šµš¶š±š±š²š»/š±š¶šš®šÆš¹š²š±?:  ' + formObject.disableModOption +  
                                       '\nšŖšµš¶š°šµ šŗš¼š±š¶š³š¶š²šæš ššµš¼šš¹š± šÆš² š±š¶šš®šÆš¹š²š±?:  ' + formObject.modDisable +  
                                       '\nšš¼ šš¼š š»š²š²š± šš¼ šš½š±š®šš² š® š½šæš¶š°š²?:  ' + formObject.updatePrice +  
                                       '\nšŖšµš¶š°šµ š¶šš²šŗ/šŗš¼š±š¶š³š¶š²šæ š»š²š²š±š š® š½šæš¶š°š² šš½š±š®šš²?:  ' + formObject.priceUpdate +                                        
                                       '\nš¦šµš¼šš¹š± ššµš¶š šÆšæš®š»š± šµš®šš² šš®š¹š²š šš®š š®š±š±š²š±? ' + formObject.tax +
                                       '\nšš¼ššæš š¼š³ š¼š½š²šæš®šš¶š¼š»: ' + formObject.hours +                                       
                                       '\nšš»š š¼ššµš²šæ š¶š»š³š¼šæšŗš®šš¶š¼š» š»š²š²š±š²š± š¼š» ššµš² šŗš²š»š?: ' + formObject.otherUpdates + 
                                       '\nššæš² ššµš²šæš² šš½š²š°š¶š®š¹ š¶š»šššæšš°šš¶š¼š»š/š»š¼šš²š š³š¼šæ ššµš¶š šÆšæš®š»š±?:  ' + formObject.specialInstructions, 
                                       MimeType.PDF);       
  const fileURL = docFile.getUrl(); 
  
  //    Create Spreadsheet in Brand folder. Activity log.
  const name = brandName + " Activity Log";
  const id = newFolder.getId();
  const resource = {
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{id: id}]      
  }; 
  
  const optionalArgs={supportsTeamDrives: true};  
  const fileJson = Drive.Files.insert(resource, null, optionalArgs);
  const fileId = fileJson.id;
  //  Add Lock
  const lock = LockService.getDocumentLock();
  lock.waitLock(120000);
  //    add if lock statment?
  try {   
    const lastRow = mySheet.getLastRow();
    //    Add data to array
    const newEntry = [
      lastRow, 
      timestamp, 
      timestamp, 
      formObject.newBrand, 
      'New Brand', 
      formObject.businessType, 
      formObject.integration, 
      '=HYPERLINK("'+formObject.menuSource+'")', 
      userEmail,
      fileURL,
      ,
      ,
      ,
      formObject.complexity,
      ,
      formObject.city,
      formObject.state,
      formObject.country,
      fileId,
      "New Brand",
      formObject.priority,
      formObject.locations
      ,
    ];   
    
    //  Add header row to spreadsheet
    const newSheet = SpreadsheetApp.openById(fileId);    
    const sheetRange = newSheet.getSheetByName("Sheet1").getRange(1,1,1,23);  
    const headers = [
      ['šš»š±š²š',
       'ššæš®š»š± š¼šæš¶š“š¶š»š®š¹ š°šæš²š®šš² š±š®šš²',
       'š§š¶šŗš² šš»',
       'ššæš®š»š± š”š®šŗš²',
       'š„š²š¾šš²šš šæš²š®šš¼š»',
       'šššš¶š»š²šš ššš½š²',
       'šš»šš²š“šæš®šš¶š¼š»',
       'š š²š»š š¦š¼ššæš°š²',
       'ššæš²š®šš²š± šÆš',
       'š š²š»š šš»šššæšš°šš¶š¼š»š',
       'šššš¶š“š»š²š± šš¼?',
       'š§š¶šŗš² š¢šš',
       'šš¼šŗš½š¹š²šš²š± šÆš:',
       'šš¼šŗš½š¹š²šš¶šš š„š®šš¶š»š“',
       'šš¼ š¹š¶šš² š±š®šš² š®š»š± šš¶šŗš²',
       'šš¶šš',
       'š¦šš®šš²/š£šæš¼šš¶š»š°š²',
       'šš¼šš»ššæš',
       'šš°šš¶šš¶šš šš¼š“',
       'š§šš½š²',
       'š£šæš¶š¼šæš¶šš š„š®šš¶š»š“', 	
       'šš¼š°š®šš¶š¼š»š', 
       'š¤š š¦š°š¼šæš²']
    ];    
    sheetRange.setValues(headers);
    
    //   Add data to last row in main tracker  
    mySheet.appendRow(newEntry);
    
    //   Copy data to spreadsheet brand 
    const activitySheet = newSheet.getSheetByName("Sheet1") 
    activitySheet.appendRow(newEntry);
    
    //   Flush changes before releasing lock
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
    throw new Error("System is busy.");
  } finally {       
    lock.releaseLock();          
  }
  return brandName
};



function processUpdateBrandForm(formObject) {
  
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  
  const activeCell = mySheet.getActiveCell();
  const cellValue = activeCell.getValue();
  const timestamp = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(userEmail);
  const index = Number(formObject.index) + 1;
  let qaScore = "";
  if(formObject.requestReason == "appovedQA") {
    qaScore = formObject.qaScore;
    }

  
  //  Update time out time stamp & completed by 
  let timeOut = mySheet.getRange(index,12);
  let completedBy = mySheet.getRange(index,13);
  
  
  //  Get data from current row
  const rowRange = mySheet.getRange(index,1,1,23);    
  let brandInfo = rowRange.getValues();
  let brandName = brandInfo[0][3];
  
  
  
  let assignee = brandInfo[0][10];
  Logger.log(assignee);
  if(formObject.requestReason == "wflAssign"){
    Logger.log(formObject.assignTo)
    assignee = formObject.assignTo;
    Logger.log(assignee);
  } 
  //  Get folder ID by using parent of Sheet    
  const spreadsheetFile = DriveApp.getFileById(brandInfo[0][18]);
  const folderId = spreadsheetFile.getParents().next().getId();    
  const brandFolder = DriveApp.getFolderById(folderId);
  //  Create a new PDF with instructions in brand folder.
  const docFile = brandFolder.createFile(formObject.requestReason +'.pdf', 
                                         'š„š²š¾šš²šš š„š²š®šš¼š»:  ' + formObject.requestReason + 
                                         '\nšš»šššæšš°šš¶š¼š»š:  ' + formObject.instructions, 
                                         MimeType.PDF);   
  const docFileURL = docFile.getUrl();
  
  
    if (formObject.requestReason == brandInfo[0][4] && formObject.requestReason !== "wflAssign"){
    throw new Error("You have selected the same request reason that this brand is on.");   
    } else if(timeOut.getValue() !== "") {
     throw new Error("Timestamp Already Added.");
    } else {
      timeOut.setValue(timestamp);
    completedBy.setValue(userEmail);  
  }; 
  //  Add Lock
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {  
 
    const lastRow = mySheet.getLastRow();
    
    //  Update data with new reason, user, in timestamp
    const updatedEntry = [
      lastRow, 
      brandInfo[0][1], 
      timestamp,
      brandInfo[0][3],
      formObject.requestReason, 
      brandInfo[0][5], 
      brandInfo[0][6], 
      '=HYPERLINK("'+brandInfo[0][7]+'")', 
        userEmail,
          //       update below? New pdf for each action required e.g. additionl info? basically any communications add userName to Doc so we can see.
          '=HYPERLINK("'+docFileURL+'")', 
            assignee,
              ,
                ,
                  brandInfo[0][13],
                    ,
                      brandInfo[0][15],
                        brandInfo[0][16],
                          brandInfo[0][17],
                            brandInfo[0][18],
                              brandInfo[0][19],
                                brandInfo[0][20],
                                  brandInfo[0][21],
                                    qaScore
                                    
                                
                            ];
    //  Add new row to main tracker
    mySheet.appendRow(updatedEntry);
    
    //  Add new row to activity logger sheet
    const activityLog = SpreadsheetApp.openById(brandInfo[0][18]);
    const activitySheet = activityLog.getSheetByName("Sheet1");
    const activitySheetLastRow = activitySheet.getLastRow();
    activitySheet.appendRow(updatedEntry);
    
    //  Add out Timestampin Activitiy tracker
    const trackerRange = activitySheet.getRange(activitySheetLastRow,12);
    const completeByRange = activitySheet.getRange(activitySheetLastRow,13);
    trackerRange.setValue(timestamp);
    completeByRange.setValue(userEmail);
    
    SpreadsheetApp.flush();
  } catch(e) {    
    Logger.log(e);
    throw new Error("System is Busy."); 
  } finally {  
    lock.releaseLock();          
  };  
  return brandName
};


function processGoLiveForm(formObject) {
  
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  
  const activeCell = mySheet.getActiveCell();
  const cellValue = activeCell.getValue();
  const timestamp = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(userEmail);
  const index = Number(formObject.index) + 1;
  
  //Get data from current row
  const rowRange = mySheet.getRange(index,1,1,23);    
  const brandInfo = rowRange.getValues();
  const brandName = (brandInfo[0][3]);

    //  Update time out time stamp
    let timeOut = mySheet.getRange(index,12);
    let completedBy = mySheet.getRange(index,13);
    
    if(timeOut.getValue() == "") {
      timeOut.setValue(timestamp);
      completedBy.setValue(userEmail);
    } else {
      throw new Error("TimestampAlready Added.");      
    };
    
    //  Get folder ID by using parent of Sheet    
    const spreadsheetFile = DriveApp.getFileById(brandInfo[0][18]);
    const folderId = spreadsheetFile.getParents().next().getId();
    const brandFolder = DriveApp.getFolderById(folderId)
    //  Create a new PDF with instructions in brand folder.
    const docFile = brandFolder.createFile('GoLive.pdf', 
                                           'š§š¶šŗš² šš°š°š¼šš»š šŖš²š»š šš¶šš²:  ' + timestamp + 
                                           '\nšØšš²šæ š§šµš®š š§š¼š¼šø šš°š°š¼šš»š šš¶šš²:  ' + userEmail, 
                                           MimeType.PDF);   
    //  const docFileURL = docFile.getUrl(); <- Not needed unless link to go live doc is preferable? Posisble user and reason? or final commnets.  P3?
    
    //  Get Lock 
    const lock = LockService.getDocumentLock();
    lock.waitLock(60000);
    try {
      const lastRow = mySheet.getLastRow();
      
      //  Update data with new reason, user, in timestamp
      const updatedEntry = [
        lastRow, 
        brandInfo[0][1], 
        timestamp, 
        brandInfo[0][3], 
        "Go Live", 
        brandInfo[0][5], 
        brandInfo[0][6], 
         '=HYPERLINK("'+brandInfo[0][7]+'")',  
          userEmail,
            "Account Live",
              brandInfo[0][10],
                timestamp,
                  userEmail,
                    brandInfo[0][13],
                      timestamp,      
                        brandInfo[0][15],
                          brandInfo[0][16],
                            brandInfo[0][17],
                              brandInfo[0][18],
                                brandInfo[0][19],
                                   brandInfo[0][20],
                                      brandInfo[0][21],
                                         brandInfo[0][22]
                              ];
      //  Add new row to main sheet
      mySheet.appendRow(updatedEntry);        
      
      //  Add new row to activity logger sheet
      const activityLog = SpreadsheetApp.openById(brandInfo[0][18])
      const activitySheet = activityLog.getSheetByName("Sheet1")
      const activitySheetLastRow = activitySheet.getLastRow();
      activitySheet.appendRow(updatedEntry);
      
      //  Add out Timestampin Activitiy tracker
    const trackerRange = activitySheet.getRange(activitySheetLastRow,12);
    const completeByRange = activitySheet.getRange(activitySheetLastRow,13);
    trackerRange.setValue(timestamp);
    completeByRange.setValue(userEmail);
   
      //  Flush changes to spreadsheet and release lock  
      SpreadsheetApp.flush();
    } catch(e) {
       Logger.log(e);
       throw new Error("System is Busy.");      
    } finally {       
      lock.releaseLock();          
    };   
  return brandName
};

function processAdditionalEditForm(formObject) {
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  
  const timestamp = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(userEmail);
  const index = Number(formObject.index) + 1;
  
  //  Update time out time stamp & completed by 
  let isLive = mySheet.getRange(index,5);
  let completedBy = mySheet.getRange(index,13);
  
//  add check that brand is live already or throw error
  if(isLive.getValue() == "Go Live") {
    
 
  //  Get data from current row
  const rowRange = mySheet.getRange(index,1,1,23);    
  let brandInfo = rowRange.getValues();
  let brandName = brandInfo[0][3];
  //  Get folder ID by using parent of Sheet    
  Logger.log(brandInfo[0][18])
  const spreadsheetFile = DriveApp.getFileById(brandInfo[0][18]);
  const folderId = spreadsheetFile.getParents().next().getId();    
  const brandFolder = DriveApp.getFolderById(folderId);
  //  Create a new PDF with instructions in brand folder.
  const docFile = brandFolder.createFile('AdditionalEdit.pdf',                                        
                                       '\nšš¶š¹š² š¦š¼ššæš°š²:  ' + formObject.menuSource +
                                       '\nšš¼ šš¼š š»š²š²š± š®š» š¶šš²šŗ šµš¶š±š±š²š»/š±š¶šš®šÆš¹š²š±?:  ' + formObject.disableItemOption +
                                       '\nšŖšµš¶š°šµ š¶šš²šŗ(š) ššµš¼šš¹š± šÆš² š±š¶šš®šÆš¹š²š±?:  ' + formObject.itemDisable +
                                       '\nšš¼ šš¼š š»š²š²š± š® šŗš¼š±š¶š³š¶š²šæ šµš¶š±š±š²š»/š±š¶šš®šÆš¹š²š±?:  ' + formObject.disableModOption +  
                                       '\nšŖšµš¶š°šµ šŗš¼š±š¶š³š¶š²šæš ššµš¼šš¹š± šÆš² š±š¶šš®šÆš¹š²š±?:  ' + formObject.modDisable +  
                                       '\nšš¼ šš¼š š»š²š²š± šš¼ šš½š±š®šš² š® š½šæš¶š°š²?:  ' + formObject.updatePrice +  
                                       '\nšŖšµš¶š°šµ š¶šš²šŗ/šŗš¼š±š¶š³š¶š²šæ š»š²š²š±š š® š½šæš¶š°š² šš½š±š®šš²?:  ' + formObject.priceUpdate + 
                                       '\nšš»š š¼ššµš²šæ š¶š»š³š¼šæšŗš®šš¶š¼š» š»š²š²š±š²š± š¼š» ššµš² šŗš²š»š?: ' + formObject.otherUpdates + 
                                       '\nššæš² ššµš²šæš² šš½š²š°š¶š®š¹ š¶š»šššæšš°šš¶š¼š»š/š»š¼šš²š š³š¼šæ ššµš¶š šÆšæš®š»š±?:  ' + formObject.specialInstructions,                                         
                                         MimeType.PDF);   
  const docFileURL = docFile.getUrl();
  
  //  Add Lock
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {         
    const lastRow = mySheet.getLastRow();
    
    //  Update data with new reason, user, in timestamp
    const updatedEntry = [
      lastRow, 
      brandInfo[0][1], 
      timestamp,
      brandInfo[0][3],
      'New Edit', 
      brandInfo[0][5], 
      brandInfo[0][6], 
      '=HYPERLINK("'+brandInfo[0][7]+'")', 
        userEmail,
          //       update below? New pdf for each action required e.g. additionl info? basically any communications add userName to Doc so we can see.
          '=HYPERLINK("'+docFileURL+'")', 
            ,
              ,
                ,
                  brandInfo[0][13],
                    ,
                      brandInfo[0][15],
                        brandInfo[0][16],
                          brandInfo[0][17],
                            brandInfo[0][18],
                              "Additional Edits",
                                brandInfo[0][20],
                                      brandInfo[0][21],
                                         
                            ];
    //  Add new row to main tracker
    mySheet.appendRow(updatedEntry);
    
    //  Add new row to activity logger sheet
    const activityLog = SpreadsheetApp.openById(brandInfo[0][18]);
    const activitySheet = activityLog.getSheetByName("Sheet1");
    const activitySheetLastRow = activitySheet.getLastRow();
    activitySheet.appendRow(updatedEntry);
    
    //  Add out Timestampin Activitiy tracker
    const trackerRange = activitySheet.getRange(activitySheetLastRow,12);
    const completeByRange = activitySheet.getRange(activitySheetLastRow,13);
    trackerRange.setValue(timestamp);
    completeByRange.setValue(userEmail);
    
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(e);
    throw new Error("System is Busy.");
  } finally {  
    lock.releaseLock();          
  };  
  return brandName
  } else {
  throw new Error("Account is not live.");
  };  
}


function processBreakForm(formObject) {
const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  
  const activeCell = mySheet.getActiveCell();
  const cellValue = activeCell.getValue();
  const timestamp = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy HH:mm:ss");
  const userEmail = Session.getActiveUser().getEmail();
  Logger.log(userEmail);
  const index = Number(formObject.index) + 1;
  
  //  Update time out time stamp & completed by 
  let timeOut = mySheet.getRange(index,12);
  let completedBy = mySheet.getRange(index,13);

  //  Get data from current row
  const rowRange = mySheet.getRange(index,1,1,23);    
  let brandInfo = rowRange.getValues();
  let brandName = brandInfo[0][3];
  //  Get folder ID by using parent of Sheet    
  const spreadsheetFile = DriveApp.getFileById(brandInfo[0][18]);
  const folderId = spreadsheetFile.getParents().next().getId();    
  const brandFolder = DriveApp.getFolderById(folderId);
  
    if(timeOut.getValue() == "") {
    timeOut.setValue(timestamp);
    completedBy.setValue(userEmail);
  } else {
   throw new Error("Timestamp Already Added.");
  }; 
  //  Add Lock
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {  
 
    const lastRow = mySheet.getLastRow();
    
    //  Update data with new reason, user, in timestamp
    const updatedEntry = [
      lastRow, 
      brandInfo[0][1], 
      timestamp,
      brandInfo[0][3],
      formObject.timeReason, 
      brandInfo[0][5], 
      brandInfo[0][6], 
      '=HYPERLINK("'+brandInfo[0][7]+'")', 
        userEmail,
          //       update below? New pdf for each action required e.g. additionl info? basically any communications add userName to Doc so we can see.
          brandInfo[0][9], 
            brandInfo[0][10],
              ,
                ,
                  brandInfo[0][13],
                    ,
                      brandInfo[0][15],
                        brandInfo[0][16],
                          brandInfo[0][17],
                            brandInfo[0][18],
                              brandInfo[0][19],
                                brandInfo[0][20],
                                  brandInfo[0][21],
                                    
                                    
                                
                            ];
    //  Add new row to main tracker
    mySheet.appendRow(updatedEntry);
    
    //  Add new row to activity logger sheet
    const activityLog = SpreadsheetApp.openById(brandInfo[0][18]);
    const activitySheet = activityLog.getSheetByName("Sheet1");
    const activitySheetLastRow = activitySheet.getLastRow();
    activitySheet.appendRow(updatedEntry);
    
    //  Add out Timestampin Activitiy tracker
    const trackerRange = activitySheet.getRange(activitySheetLastRow,12);
    const completeByRange = activitySheet.getRange(activitySheetLastRow,13);
    trackerRange.setValue(timestamp);
    completeByRange.setValue(userEmail);
    
    SpreadsheetApp.flush();
  } catch(e) {    
    Logger.log(e);
    throw new Error("System is Busy."); 
  } finally {  
    lock.releaseLock();          
  };  
  return brandName
};





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
                                       '𝗕𝗿𝗮𝗻𝗱 𝗡𝗮𝗺𝗲:  ' + formObject.newBrand + 
                                       '\n𝗖𝗼𝘂𝗻𝘁𝗿𝘆:  ' + formObject.country +
                                       '\n𝗦𝘁𝗮𝘁𝗲:  ' + formObject.state + 
                                       '\n𝗖𝗶𝘁𝘆:  ' + formObject.city + 
                                       '\n𝗡𝘂𝗺𝗯𝗲𝗿 𝗼𝗳 𝗟𝗼𝗰𝗮𝘁𝗶𝗼𝗻𝘀:  ' + formObject.locations + 
                                       '\n𝗕𝘂𝘀𝗶𝗻𝗲𝘀𝘀 𝗧𝘆𝗽𝗲:  ' + formObject.businessType +
                                       '\n𝗜𝗻𝘁𝗲𝗴𝗿𝗮𝘁𝗶𝗼𝗻:  ' + formObject.integration + 
                                       '\n𝗖𝗼𝗺𝗽𝗹𝗲𝘅𝗶𝘁𝘆:  ' + formObject.complexity + 
                                       '\n𝗣𝗿𝗶𝗼𝗿𝗶𝘁𝘆 𝗥𝗮𝘁𝗶𝗻𝗴:  ' + formObject.priority +
                                       '\n𝗙𝗶𝗹𝗲 𝗦𝗼𝘂𝗿𝗰𝗲:  ' + formObject.menuSource +
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝗮𝗻 𝗶𝘁𝗲𝗺 𝗵𝗶𝗱𝗱𝗲𝗻/𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.disableItemOption +
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗶𝘁𝗲𝗺(𝘀) 𝘀𝗵𝗼𝘂𝗹𝗱 𝗯𝗲 𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.itemDisable +
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝗮 𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿 𝗵𝗶𝗱𝗱𝗲𝗻/𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.disableModOption +  
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿𝘀 𝘀𝗵𝗼𝘂𝗹𝗱 𝗯𝗲 𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.modDisable +  
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝘁𝗼 𝘂𝗽𝗱𝗮𝘁𝗲 𝗮 𝗽𝗿𝗶𝗰𝗲?:  ' + formObject.updatePrice +  
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗶𝘁𝗲𝗺/𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿 𝗻𝗲𝗲𝗱𝘀 𝗮 𝗽𝗿𝗶𝗰𝗲 𝘂𝗽𝗱𝗮𝘁𝗲?:  ' + formObject.priceUpdate +                                        
                                       '\n𝗦𝗵𝗼𝘂𝗹𝗱 𝘁𝗵𝗶𝘀 𝗯𝗿𝗮𝗻𝗱 𝗵𝗮𝘃𝗲 𝘀𝗮𝗹𝗲𝘀 𝘁𝗮𝘅 𝗮𝗱𝗱𝗲𝗱? ' + formObject.tax +
                                       '\n𝗛𝗼𝘂𝗿𝘀 𝗼𝗳 𝗼𝗽𝗲𝗿𝗮𝘁𝗶𝗼𝗻: ' + formObject.hours +                                       
                                       '\n𝗔𝗻𝘆 𝗼𝘁𝗵𝗲𝗿 𝗶𝗻𝗳𝗼𝗿𝗺𝗮𝘁𝗶𝗼𝗻 𝗻𝗲𝗲𝗱𝗲𝗱 𝗼𝗻 𝘁𝗵𝗲 𝗺𝗲𝗻𝘂?: ' + formObject.otherUpdates + 
                                       '\n𝗔𝗿𝗲 𝘁𝗵𝗲𝗿𝗲 𝘀𝗽𝗲𝗰𝗶𝗮𝗹 𝗶𝗻𝘀𝘁𝗿𝘂𝗰𝘁𝗶𝗼𝗻𝘀/𝗻𝗼𝘁𝗲𝘀 𝗳𝗼𝗿 𝘁𝗵𝗶𝘀 𝗯𝗿𝗮𝗻𝗱?:  ' + formObject.specialInstructions, 
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
      ['𝗜𝗻𝗱𝗲𝘅',
       '𝗕𝗿𝗮𝗻𝗱 𝗼𝗿𝗶𝗴𝗶𝗻𝗮𝗹 𝗰𝗿𝗲𝗮𝘁𝗲 𝗱𝗮𝘁𝗲',
       '𝗧𝗶𝗺𝗲 𝗜𝗻',
       '𝗕𝗿𝗮𝗻𝗱 𝗡𝗮𝗺𝗲',
       '𝗥𝗲𝗾𝘂𝗲𝘀𝘁 𝗿𝗲𝗮𝘀𝗼𝗻',
       '𝗕𝘂𝘀𝗶𝗻𝗲𝘀𝘀 𝘁𝘆𝗽𝗲',
       '𝗜𝗻𝘁𝗲𝗴𝗿𝗮𝘁𝗶𝗼𝗻',
       '𝗠𝗲𝗻𝘂 𝗦𝗼𝘂𝗿𝗰𝗲',
       '𝗖𝗿𝗲𝗮𝘁𝗲𝗱 𝗯𝘆',
       '𝗠𝗲𝗻𝘂 𝗜𝗻𝘀𝘁𝗿𝘂𝗰𝘁𝗶𝗼𝗻𝘀',
       '𝗔𝘀𝘀𝗶𝗴𝗻𝗲𝗱 𝘁𝗼?',
       '𝗧𝗶𝗺𝗲 𝗢𝘂𝘁',
       '𝗖𝗼𝗺𝗽𝗹𝗲𝘁𝗲𝗱 𝗯𝘆:',
       '𝗖𝗼𝗺𝗽𝗹𝗲𝘅𝗶𝘁𝘆 𝗥𝗮𝘁𝗶𝗻𝗴',
       '𝗚𝗼 𝗹𝗶𝘃𝗲 𝗱𝗮𝘁𝗲 𝗮𝗻𝗱 𝘁𝗶𝗺𝗲',
       '𝗖𝗶𝘁𝘆',
       '𝗦𝘁𝗮𝘁𝗲/𝗣𝗿𝗼𝘃𝗶𝗻𝗰𝗲',
       '𝗖𝗼𝘂𝗻𝘁𝗿𝘆',
       '𝗔𝗰𝘁𝗶𝘃𝗶𝘁𝘆 𝗟𝗼𝗴',
       '𝗧𝘆𝗽𝗲',
       '𝗣𝗿𝗶𝗼𝗿𝗶𝘁𝘆 𝗥𝗮𝘁𝗶𝗻𝗴', 	
       '𝗟𝗼𝗰𝗮𝘁𝗶𝗼𝗻𝘀', 
       '𝗤𝗔 𝗦𝗰𝗼𝗿𝗲']
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
                                         '𝗥𝗲𝗾𝘂𝗲𝘀𝘁 𝗥𝗲𝗮𝘀𝗼𝗻:  ' + formObject.requestReason + 
                                         '\n𝗜𝗻𝘀𝘁𝗿𝘂𝗰𝘁𝗶𝗼𝗻𝘀:  ' + formObject.instructions, 
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
                                           '𝗧𝗶𝗺𝗲 𝗔𝗰𝗰𝗼𝘂𝗻𝘁 𝗪𝗲𝗻𝘁 𝗟𝗶𝘃𝗲:  ' + timestamp + 
                                           '\n𝗨𝘀𝗲𝗿 𝗧𝗵𝗮𝘁 𝗧𝗼𝗼𝗸 𝗔𝗰𝗰𝗼𝘂𝗻𝘁 𝗟𝗶𝘃𝗲:  ' + userEmail, 
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
                                       '\n𝗙𝗶𝗹𝗲 𝗦𝗼𝘂𝗿𝗰𝗲:  ' + formObject.menuSource +
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝗮𝗻 𝗶𝘁𝗲𝗺 𝗵𝗶𝗱𝗱𝗲𝗻/𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.disableItemOption +
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗶𝘁𝗲𝗺(𝘀) 𝘀𝗵𝗼𝘂𝗹𝗱 𝗯𝗲 𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.itemDisable +
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝗮 𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿 𝗵𝗶𝗱𝗱𝗲𝗻/𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.disableModOption +  
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿𝘀 𝘀𝗵𝗼𝘂𝗹𝗱 𝗯𝗲 𝗱𝗶𝘀𝗮𝗯𝗹𝗲𝗱?:  ' + formObject.modDisable +  
                                       '\n𝗗𝗼 𝘆𝗼𝘂 𝗻𝗲𝗲𝗱 𝘁𝗼 𝘂𝗽𝗱𝗮𝘁𝗲 𝗮 𝗽𝗿𝗶𝗰𝗲?:  ' + formObject.updatePrice +  
                                       '\n𝗪𝗵𝗶𝗰𝗵 𝗶𝘁𝗲𝗺/𝗺𝗼𝗱𝗶𝗳𝗶𝗲𝗿 𝗻𝗲𝗲𝗱𝘀 𝗮 𝗽𝗿𝗶𝗰𝗲 𝘂𝗽𝗱𝗮𝘁𝗲?:  ' + formObject.priceUpdate + 
                                       '\n𝗔𝗻𝘆 𝗼𝘁𝗵𝗲𝗿 𝗶𝗻𝗳𝗼𝗿𝗺𝗮𝘁𝗶𝗼𝗻 𝗻𝗲𝗲𝗱𝗲𝗱 𝗼𝗻 𝘁𝗵𝗲 𝗺𝗲𝗻𝘂?: ' + formObject.otherUpdates + 
                                       '\n𝗔𝗿𝗲 𝘁𝗵𝗲𝗿𝗲 𝘀𝗽𝗲𝗰𝗶𝗮𝗹 𝗶𝗻𝘀𝘁𝗿𝘂𝗰𝘁𝗶𝗼𝗻𝘀/𝗻𝗼𝘁𝗲𝘀 𝗳𝗼𝗿 𝘁𝗵𝗶𝘀 𝗯𝗿𝗮𝗻𝗱?:  ' + formObject.specialInstructions,                                         
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





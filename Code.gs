//global variables
const sheetId = "foo";
const Route = {};
Route.path = function(route, callback){
  Route[route] = callback;
}

function onOpen() {
  Logger.log(Session.getActiveUser().getEmail());
  SpreadsheetApp.getUi() 
  .createMenu('Menu')
  .addItem('Open Web App', 'openUrl')   
  .addToUi();
}

function doGet(e){
  Logger.log(Session.getActiveUser().getEmail());
  Route.path("newAccountForm",loadNewForm);
  Route.path("updateBrandForm", loadUpdateForm);
  Route.path("goLive", loadGoLiveForm);
  Route.path("additionalEdits", loadAdditionalEditsForm);
  Route.path("breaks", loadBreaksForm);
  Route.path("reports", loadReportsPage);
  
  if(Route[e.parameters.v]) {
    return Route[e.parameters.v](); 
  } else {
    return render("home") 
  }
};

//function doPost(e) {
// 
//  return render("error", {e:e})
//  
//}  

function loadNewForm() {  
  return render("addNewAccount");  
}

function loadUpdateForm(){ 
//  add agentList  
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Options");   
  const agentList = mySheet.getRange(2,3,mySheet.getRange("C2").getDataRegion().getLastRow(),1).getValues();
  
  const repList = agentList.map(function(r){
  return '<option value="' + r[0] + '">'+ r[0] + '</option>';
  }).join(''); 
  return render("updateBrand", {repList:repList});  
}

function loadGoLiveForm () {
return render("goLive");
}

function loadAdditionalEditsForm () {
return render("additionalEdits");
}

function loadBreaksForm(){
  return render("breaks");
}

function loadReportsPage(){
  return render("reports");
}

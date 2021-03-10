function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openUrl(){
  var html = HtmlService.createHtmlOutput('<html><script>'
                                          +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
                                          +'var a = document.createElement("a"); a.href="'+'https://script.google.com/a/google.com/macros/s/AKfycbxW67ahkQeyiaddysKECw3_H5iZNc8epSIgI1vRvXEkN1ekVyUL/exec'+'"; a.target="_blank";'
                                          +'if(document.createEvent){'
                                          +'  var event=document.createEvent("MouseEvents");'
                                          +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
                                          +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
                                          +'}else{ a.click() }'
                                          +'close();'
                                          +'</script>'
                                          // Offer URL as clickable link in case above code fails.
                                          +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+'https://script.google.com/a/google.com/macros/s/AKfycbxW67ahkQeyiaddysKECw3_H5iZNc8epSIgI1vRvXEkN1ekVyUL/exec'+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
                                          +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
                                          +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function render(file,argsObject) {
  const tmp = HtmlService.createTemplateFromFile(file);
  if(argsObject) {
    const keys = Object.keys(argsObject);
    keys.forEach(function(key){
      tmp[key] = argsObject[key];
    })   
  }  // END IF  
  return tmp.evaluate();  
}

function matchBrand(indexData) {
  
  const sheetActive = SpreadsheetApp.openById(sheetId);
  const mySheet = sheetActive.getSheetByName("Sheet1");
  const index = parseInt(indexData)
  
  //  range of index values
  
  let brandMatch = mySheet.getRange(index+1,4).getValue();
  
  if (brandMatch.length < 1 || index == 0) {
    return "Brand Not Found Please Check Index"
  } else {
    return brandMatch
  }
}

function onFormSubmit(e) {
 
  
  Logger.log(e.namedValues);

  

}

function myFunction(){
  // Was separated line by line for debugging purposes
  // You have to go to Edition > Triggers > Form being sent
  // See https://stackoverflow.com/questions/43429161/how-to-get-form-values-in-the-submit-event-handler to install in a form Directly
  
  var sheet = SpreadsheetApp.getActive();
  var a = ScriptApp.newTrigger("onFormSubmit");
  var b = a.forSpreadsheet(sheet);
  var c = b.onFormSubmit();
  var d = c.create();
}
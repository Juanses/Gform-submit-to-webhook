function onFormSubmit(e) {
  var comm = new CommunicationClass();
  var colnames = ["Nom du médecin","Numéro de la question","Motif du suivi"];
  var dataname = ["medecin","question","motif"];
  var data = {};
  for (i in colnames){
    data[dataname[i]]=e.namedValues[colnames[i]][0]
  }
  comm.sendtowebhook("https://hook.integromat.com/h1nm6yyw3qzpeo7akkt8nl6txnqubfcs",data);
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
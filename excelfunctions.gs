var ExcelClass = function(){
  /*
  GIST - Start - 
  var excel = new ExcelClass();
  excel.setSheetbyName("RH");
  var cols = ["nom","prenom","fonction","email","telephone","ddn","ldn","adresse","ddc","feuille"];
  excel.setCols(cols);
  var values = excel.getallvalues(2,1,cols.length);
  var feuillename = excel.get_colnumber_fromname("feuille");
  for (var i = 0; i < values.length; i++) {
  Logger.log(values[0][feuillename]);
  }
  
  function doGet(e) {
  var excel = new ExcelClass();
  excel.setsheetbyId("1Bz5ZcdWvpBOudh0nd2EwapFIpspWZs0W4zjmof4IZUI",0);
  var cols = ["email","prenom","nom","spe","telephone"];
  excel.setCols(cols);
  var values = excel.getallvalues(2,1,cols.length);
  var result = {};
  for (var i = 0; i < values.length; i++) {
  result[i] = {};
  for (var j = 0; j < cols.length; j++) {
  //Logger.log(values[i][excel.get_colnumber_fromname(cols[j])]);
  result[i][cols[j]]= values[i][excel.get_colnumber_fromname(cols[j])];
  }
  }
  return ContentService.createTextOutput(JSON.stringify(result))
  .setMimeType(ContentService.MimeType.JSON);
  }
  
  */
  
  //*************** SPREADSHEET START ***************
  
  //SET SHEET FROM SHEET NUMBER [FROM 0 TO X]
  this.setSpreadhsheetbyDefault = function(){
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  //SET SHEET FROM SHEET NUMBER [FROM 0 TO X]
  this.setSpreadhsheetbyId = function(id){
    this.ss = SpreadsheetApp.openById(id);
  }
  
  //*************** SPREADSHEET END ***************
  
  
  //*************** SHEET START ***************
  
  //SET SHEET FROM SHEET NUMBER [FROM 0 TO X]
  this.setSheetbyNumber = function (number){
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getSheets()[number]; 
  }
  
  //SET SHEET FROM SHEET NAME
  this.setSheetbyName = function (name){
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getSheetByName(name); 
  }
  
  //SET SHEET FROM SPREEDSHEET ID
  this.setsheetbyId = function(id,sheetnumber){
    this.ss = SpreadsheetApp.openById(id);
    this.sheet = this.ss.getSheets()[sheetnumber];
  }
  
  //*************** SHEET END ***************
  
  //*************** COLUMNS START ***************
  
  //GET THE LETTER OF THE COLUMN CORRESPONDING TO THE COLS TITLE
  this.getColletterfromCol = function (colname){
    var index = this.cols.indexOf(colname);
    return String.fromCharCode(index+65);
  }
  //GET COL NUMBER BASED ON COLNAME
  this.get_colnumber_fromname = function (name){
    return this.cols.indexOf(name);
  }
  
  //SET THE COLS IN EXCEL OBJECT
  this.setCols = function (cols){
    this.cols = cols;
  }
  
  //*************** COLUMNS END ***************
  
  //*************** VALUES START ***************
  
  //SET VALUE IN CELL BASED ON ROW AND COLNAME
  this.setValueinCol = function (value,row,colname){
    var index = this.cols.indexOf(colname);
    if (index != -1) {var range = this.sheet.getRange(row,index+1);range.setValue(value);}
  }
  
  //SET VALUE IN CELL BASED ON ROW AND COLNAME
  this.setFormulainCol = function (value,row,colname){
    var index = this.cols.indexOf(colname);
    var range = this.sheet.getRange(row,index+1);
    range.setFormula(value);
  }
  
  //GET LAST ROW OF THE SHEET
  this.getLastRow = function(){
    this.lastrow = this.sheet.getLastRow()+1;
    return this.lastrow;
  }
  
  this.appendvalue = function(value,colname){
    var row = this.sheet.getLastRow()+1;
    this.setValueinCol(value,row,colname);
  }
  
  this.appendvalues = function(values){
    //values object = {colname : value, colname : value ....}
    var row = this.sheet.getLastRow()+1;
    for (var key in values) {
      var index = this.cols.indexOf(key);
       if (index != -1) {var range = this.sheet.getRange(row,index+1);range.setValue(values[key]);}

    }  
  }
  
  //GET ALL THE VALUES IN A RANGE
  this.getallvalues = function(startline,startcolumn,endcolumn){
    var lastrow = this.sheet.getLastRow();
    return this.sheet.getRange(startline,startcolumn,lastrow-(startline-1),endcolumn).getValues(); 
  }
  
  //GET ALL THE VALUES AND CREATE AN OBJECT FROM A ROW BASED ON THE COLUMNS 
  this.getvaluesfromrow = function(values,startrow,rownumber){
    var row;
    var rowdata = {};
    for (var i = 0; i < values.length; i++) {
      if (i == (rownumber-startrow)){
        for(key in this.cols){
          rowdata[this.cols[key]] = values[i][key];
        }
        break;
      }
    }   
    return rowdata;
  } 
  
  this.getSpecificDatafromRow = function(values,selectedarray){
    var rowdata = {}; 
    for (var j= 0; j < selectedarray.length ; j++) {
      rowdata[selectedarray[j]] = values[this.get_colnumber_fromname(selectedarray[j])];
    }
    return rowdata; 
  }
  
  //*************** VALUES END ***************
  
  //*************** OTHER START ***************
  
  this.exportToPdf = function(){
    //return file
    return DriveApp.createFile(this.ss.getBlob());
  }
  
  
  //SET FONT COLOR IN CELL BASED ON ROW AND COLNAME
  this.setcolorincol = function (color,row,colname){
    var index = this.cols.indexOf(colname);
    var range = this.sheet.getRange(row,index+1);
    range.setFontColor(color);
  }
  
  //SET DATA VALIDATION
  this.setDataValidationbyRange = function (sheet1,optionsrange,sheet2,validation){
    var range1 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheet1].getRange(optionsrange);
    var range2 = SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheet2].getRange(validation);
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range1).build();
    range2.setDataValidation(rule);
  }
  
  //DATA SEARCH
  this.searchspreadsheet = function (spreadsheetid,value,col){
    var unsuscribedbsheet = SpreadsheetApp.openById(spreadsheetid).getSheets()[0];
    var sheetlastRow = unsuscribedbsheet.getLastRow();
    // A CHANGER var values = unsuscribedbsheet.getRange(2,1,sheetlastRow-1,15).getValues();
    // A CHANGER var enum=2;
    var found=-1;
    values.forEach(function(row){
      if (row[col-1] == value){
        found = enum;
        return found;
      }
      enum++;
    });
    return found;
  }
  
  //REMOVE DUPLICATES
  this.removeduplicates = function (spreadsheetid,sheet,startrow,col){
    //I remove duplicates and return the number of unique responses
    var responses = SpreadsheetApp.openById(spreadsheetid);
    var responsessheet = responses.getSheets()[sheet];
    var responsestRow = responsessheet.getLastRow();
    if (responsestRow != 1){
      //The form has responses
      var values = responsessheet.getRange(startrow,col,responsestRow-(startrow-1),1).getValues();
      var buffer = []; 
      var continuer = true;
      var i = 0;
      while (continuer) {
        if (buffer.indexOf(values[i][0]) === -1) {
          buffer.push(values[i][0]);
          i++;
        }
        else{
          responsessheet.deleteRow(i+2);
          values.splice(i, 1);
          i = buffer.length;
        }
        continuer = (buffer.length == values.length)?false:true;
      }
      return buffer.length;
    }
    else{
      return 0;
    }
  }
  //*************** OTHER END ***************
}

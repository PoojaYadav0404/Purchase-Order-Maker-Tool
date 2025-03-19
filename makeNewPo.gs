function onOpen() {
  hideSheets();
  NewMenu(); 
}

function NewMenu(){
  SpreadsheetApp.getUi()
  .createMenu("NEW MENU")
  .addItem("Make New PO", "makeNewPo")
  .addSeparator()
  .addItem("Add Row", "AddNewRow")
  .addSeparator()
  .addItem("Request Approval", "requestApproval")
  .addSeparator()
  .addItem("Approve", "approved")
  .addSeparator()
  .addToUi();

}

function hideSheets(){
  const HiddenSheets = ["PO Mst Fmt", "Supplier Mst", "PO Database", "Mail Temp", "SKU Mst"];
  const SS = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  SS.forEach(function(sheet){
    if(HiddenSheets.includes(sheet.getName())){
      sheet.hideSheet();
    }
  }); 

}

function makeNewPo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var PIMstSheet = ss.getSheetByName("PO Mst Fmt");
  var SplMstSheet = ss.getSheetByName("Supplier Mst");
  var NewPINo = Number(SplMstSheet.getRange("A2").getValue()) + 1;
  var ThisYear = Utilities.formatDate(new Date(), "GMT+05:30", "yy");
  //Logger.log(Test);
  var YearCycle = [];
  if (new Date().getMonth()> 2){
    YearCycle.push(Number(Utilities.formatDate(new Date(), "GMY+05:30", "yyyy")) + "-" + (Number(ThisYear) + 1));
  
  }else if (new Date().getMonth() === 2){
    YearCycle.push((Number(Utilities.formatDate(new Date(), "GMT+05:30", "yyyy")) - 1) + "-" + ThisYear)

  }else if (new Date().getMonth() < 2){
    YearCycle.push((Number(Utilities.formatDate(new Date(), "GMT+05:30", "yyyy")) - 1) + "-" + ThisYear)
  }

  PIMstSheet.activate();
  var ActiveSheet = ss.duplicateActiveSheet().setName(NewPINo + "/" + Utilities.formatDate(new Date(), "GMT+05:30", "dd-MMM-yyyy") + ' Customer Name').activate();
   
  Logger.log(YearCycle)

  ActiveSheet.getRange("I6").setValue([["PO",YearCycle,NewPINo].join("/")]);
  SplMstSheet.getRange("A2").setValue(NewPINo);
 
  PIMstSheet.hideSheet();
  
  hideSheets();
   
}

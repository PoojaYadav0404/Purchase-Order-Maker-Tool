function AddNewRow() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var WorkingSheets = SS.getActiveSheet();
  var LastRow = WorkingSheets.getLastRow();
  var MasterSheets = ["Instruction" ,"PO Mst Fmt", "Supplier Mst", "SKU Mst", "PO List", "PO Database", "Mail Temp", "PO Data Filter", "SKU Filter"];
  if (MasterSheets.includes(WorkingSheets.getName())) {
    SS.toast("Adding row in a wrong sheet.", "Warning!!!");

  } else {
    var AddRowPosition = WorkingSheets.getLastRow() - 10;
    var Ui = SpreadsheetApp.getUi();
    var Alert = Ui.prompt("Add New Row", "How many rows do you want to add?", Ui.ButtonSet.OK_CANCEL);
    var NoOfRowsToAdd = Alert.getResponseText();
    WorkingSheets.insertRowsAfter(AddRowPosition, NoOfRowsToAdd);

    WorkingSheets.getRange("J17").setValue("=ArrayFormula($I17:$I" + (Number(AddRowPosition) + Number(NoOfRowsToAdd)) + "*$H17:$H" + (Number(AddRowPosition) + Number(NoOfRowsToAdd)) + ")")
    WorkingSheets.getRange(WorkingSheets.getLastRow()-9, 10).setValue(['=SUM($J17:$J' + (Number(AddRowPosition) + Number(NoOfRowsToAdd) )  + ')'])
    WorkingSheets.getRange(17, 1, WorkingSheets.getLastRow() - 10-16,10).setBorder(false, false, false, false, false, false)
    WorkingSheets.getRange(17, 1, WorkingSheets.getLastRow() - 10-16,10).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
    WorkingSheets.getRange(17, 1, WorkingSheets.getLastRow() - 10-16,10).setBorder(null, null, null, null, true, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    
    WorkingSheets.getRange(17, 2, WorkingSheets.getLastRow() - 10-16,3).setBorder(false, false, false, false, false, false)
    WorkingSheets.getRange(17, 2, WorkingSheets.getLastRow() - 10-16,3).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    var MrgCells = WorkingSheets.getRange(17, 1, WorkingSheets.getLastRow() - 10-16, 4).getValues();
    MrgCells.forEach(function(e, index){
      WorkingSheets.getRange(17+ index, 2, 1, 3).mergeAcross();
    })

    SS.toast("New Row added successfully.", "Successfull!!!!");

  }

}

function onEdit() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var SplMstSheet = SS.getSheetByName("Supplier Mst");
  var PoListSheet = SS.getSheetByName("PO List");
  var SkuMst = SS.getSheetByName("SKU Mst");
  var MasterSheets = ["Instruction" ,"PO Mst Fmt", "Supplier Mst", "SKU Mst", "PO List", "Mail Temp", "PO Data Filter", "SKU Filter"];
  var ActiveSheet = SS.getActiveSheet();
  var SheetName = ActiveSheet.getName();
  var Vendor = ActiveSheet.getRange("B6").getValue();
  var Row = ActiveSheet.getActiveCell().getRow();
  var Col = ActiveSheet.getActiveCell().getColumn();
  var PoNo = ActiveSheet.getRange("I6").getValue();
  var PoDate = ActiveSheet.getRange("I7").getValue();
  var FmtPoDate = Utilities.formatDate(new Date(PoDate),"GMT+05:30", "dd-MMM-yy")

  if (!MasterSheets.includes(SheetName) && Row === 6 && Col === 2 && (PoNo === "" || PoDate === "")) {
    var Ui = SpreadsheetApp.getUi();
    Ui.alert("Alert!!,", "Either Proforma no. or Date or both are blank", Ui.ButtonSet.OK);
    ActiveSheet.getRange("B6").clearContent();

  } else if (!MasterSheets.includes(SheetName) && PoNo != "" && (Row === 6 && Col === 2) || (Row ===7 && Col === 9) ) {
    var PoNoSplit = PoNo.split("/")[2];
    ActiveSheet.setName(PoNoSplit + "/" + FmtPoDate + " " + Vendor);

    var SplyData = SplMstSheet.getRange(2, 3, SplMstSheet.getLastRow() - 1, SplMstSheet.getLastColumn()-2).getValues();
    var SplMstSheetData = SplyData.filter(function (row) {
      return row.some(function (cell) {
        return cell === Vendor
      })
    });
    Logger.log(SplMstSheetData)
    ActiveSheet.getRange("B7").setValue(SplMstSheetData[0][1]); //address
    ActiveSheet.getRange("B11").setValue(SplMstSheetData[0][5]); //GST
    ActiveSheet.getRange("I9").setValue(SplMstSheetData[0][6]); //TYPE OF MATERIAL
    //ActiveSheet.getRange("C13").setValue(SplMstSheetData[0][8]); //Payment Terms
    //ActiveSheet.getRange(ActiveSheet.getLastRow() - 10, 3).setValue(SplMstSheetData[0][7]); //Incoterms
    //ActiveSheet.getRange(ActiveSheet.getLastRow() - 8, 3).setValue(SplMstSheetData[0][9]); //packing

  } else if (!MasterSheets.includes(SheetName) && Row > 16 && Row <= ActiveSheet.getLastRow()-10 && Col === 1 && ActiveSheet.getRange(Row, 1).getValue()!="" ) {

    var SkuMstData = SkuMst.getRange(2, 1, SkuMst.getLastRow() - 1, SkuMst.getLastColumn()).getValues();
    var ActiveVals = ActiveSheet.getRange(Row, Col).getValue();
    var Arr = [];
    SkuMstData.forEach(function(row){
      if (ActiveVals===row[9]) {

        //CATEOGUE, ITEM NAME, COLOUR
        //[row[2], row[1], row[7]].join(" ")
        Arr.push([row[2], row[1], row[6]].join(", "), "", "", `=image("${row[12]}")`);
        Logger.log(row[12])

      }
    })
    Logger.log(Arr.length)
    //Logger.log(ActiveSheet.getRange(Row, Col).getA1Notation())
    if (Arr.length > 0) {

      ActiveSheet.getRange(Row, 2, 1, Arr.length).setValues([Arr]);

    } else {
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "Either duplicate entry or Match do not found.", Ui.ButtonSet.OK)
      ActiveSheet.getRange(Row, Col).clearContent();

    }

  } 

}

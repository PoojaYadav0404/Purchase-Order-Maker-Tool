function requestApproval() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var MasterSheets = ["Instruction" ,"PO Mst Fmt", "Supplier Mst", "SKU Mst", "PO List", "PO Database", "Mail Temp", "PO Data Filter", "SKU Filter"];
  var MailTemp = SS.getSheetByName("Mail Temp");
  var PoListSheet = SS.getSheetByName("PO List")
  var ActiveSheet = SS.getActiveSheet();
  var Vendor = ActiveSheet.getRange("B6").getValue();
  var PoNo = ActiveSheet.getRange("I6").getValue();
  var PoDate = ActiveSheet.getRange("I7").getValue();
  var DelTime = ActiveSheet.getRange("I10").getValue();
  var PreparedBy = ActiveSheet.getRange(ActiveSheet.getLastRow()-1, 4).getValue();

  if (!MasterSheets.includes(ActiveSheet.getName()) && Vendor != ""  && PoDate != "" && DelTime != "" && PoNo != "" && PreparedBy!="") {

    var InspData = PoListSheet.getRange(2, 1, PoListSheet.getLastRow()-1, PoListSheet.getLastColumn()).getValues();
    var RecordFound = InspData.filter(r => r[0] == PoNo);

    if (RecordFound.length === 0) {
      var PoReportLink = "https://docs.google.com/spreadsheets/d/17r4L/edit?gid=" + ActiveSheet.getSheetId() + "#gid=" + ActiveSheet.getSheetId();
      var url_base = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", SS.getId());
      var url_ext = 'export?exportFormat=pdf&format=pdf' //export as pdf/ csv. xls
        // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
        //+
        //('&id=' + ActiveSheet.getSheetId())
        +
        '&size=A4' // paper size
        +
        '&portrait=true' // orientation, false for landscape
        +
        '&scale=2' //1=Normal 100% / 2= Fit to Width / 3=Fit to height / 4=Fit to Page
        //+
        //'&fitw=true' // fit to width, false for actual size
        +
        '&top_margin=0.50' //All four margins must be set!
        +
        '&bottom_margin=0.50' //All four margins must be set!
        +
        '&left_margin=0.50' //All four margins must be set!
        +
        '&right_margin=0.50' //All four margins must be set!
        +
        '&sheetnames=false&printtitle=false&pagenumbers=false' //hide optional headers and footers
        +
        '&gridlines=false' // hide gridlines
        +
        '&fzr=false' // do not repeat row headers (frozen rows) on each page
        + 
        '&gid=';

      var options = {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        }
      }

      var Response = UrlFetchApp.fetch(url_base + url_ext + ActiveSheet.getSheetId(), options);
      var InspSheetBlob = Response.getBlob().setName(PoNo + " " + Vendor + ".pdf");
      var UserId = Session.getActiveUser().getEmail();
      var MailName = MailTemp.getRange("B2").getValue();
      var To = MailTemp.getRange("B3").getValue();
      var CC = MailTemp.getRange("B4").getValue();
      var Sub = PoNo + MailTemp.getRange("B5").getValue() + Vendor;
      var ReplyTo = MailTemp.getRange("B7").getValue() ;
      var Body = MailTemp.getRange("B6").getValue().replace("{po no.}",PoNo).replace("{vendor}",Vendor).replace("{link}", PoReportLink ) ;

      GmailApp.sendEmail(To, Sub, "", {
        name:MailName, 
        htmlBody: Body, 
        replyTo: UserId, 
        cc: CC, 
        attachments: InspSheetBlob })

      var Arr = [PoNo,Vendor, PoDate, DelTime + " days", PoReportLink, "Not approved"];
      PoListSheet.getRange(PoListSheet.getLastRow() + 1, 1, 1, Arr.length).setValues([Arr]);
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Approval Mail Status!!", "Purchase Order send succesfully for approval.", Ui.ButtonSet.OK);

    } else {
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "This Purchase order no. already exist.", Ui.ButtonSet.OK);
    }
  } else {
    var Ui = SpreadsheetApp.getUi();
    Ui.alert("Alert!!", "Either Wrong Sheet or Important details(i.e. Vendor, Po No., PO Date, Delivery Time, Prepared by) are missing.", Ui.ButtonSet.OK)

  }
  
}

function approved() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var MasterSheets = ["Instruction" ,"PO Mst Fmt", "Supplier Mst", "SKU Mst", "PO List", "PO Database", "Mail Temp", "PO Data Filter", "SKU Filter"];
  var MailTemp = SS.getSheetByName("Mail Temp");
  var PoListSheet = SS.getSheetByName("PO List");
  var DatabaseSheet = SS.getSheetByName("PO Database");
  var ActiveSheet = SS.getActiveSheet();
  var Vendor = ActiveSheet.getRange("B6").getValue();
  var VGst = ActiveSheet.getRange("B11").getValue();
  var PoNo = ActiveSheet.getRange("I6").getValue();
  var PoDate = ActiveSheet.getRange("I7").getValue();
  var DelTime = ActiveSheet.getRange("I10").getValue();
  var LotType = ActiveSheet.getRange("I11").getValue();
  var PreparedBy = ActiveSheet.getRange(ActiveSheet.getLastRow()-1, 4).getValue();
  var UserId = Session.getActiveUser().getEmail();
  var OldSub = PoNo + MailTemp.getRange("E5").getValue() + Vendor;
  var RepBody = MailTemp.getRange("E6").getValue();
  var Thread = GmailApp.getInboxThreads(0, 200);

  if (!MasterSheets.includes(ActiveSheet.getName()) && (UserId == "rajatbajaj@bajato.com" || UserId=="ppc@bajato.com" || UserId=="ops@bajato.com" || UserId=="purchase@bajato.com" || UserId=="bajatoanoop@gmail.com") && Vendor != ""  && PoDate != "" && DelTime != "" && PoNo != "" && PreparedBy!="") {

    var OrderArr = (!PoListSheet.getLastRow() < 1) ? PoListSheet.getRange(2, 1, PoListSheet.getLastRow() - 1, 1).getValues().map(insp => { return insp.toString() }) : [];
    var Row = OrderArr.indexOf(PoNo.toString()) + 2;
    //Logger.log(OrderArr);
    //Logger.log(Row);
    var ThreadId = []
    for (i = 0; i < Thread.length; i++) {
      Logger.log(Thread[i].getFirstMessageSubject())
      if (Thread[i].getFirstMessageSubject() == OldSub) {
        //Logger.log(Thread[i].getFirstMessageSubject())
        ThreadId.push(Thread[i].getId());
        break
      }
    }
    Logger.log(ThreadId.length)

    if(Row > 1 && ThreadId.length === 1){

      var FolderId = "1rSZR-tf_CYwTlg-IpN71SeYb0aYDgY11"
      var MasterFormatId = "1xxa7L2-r1mj5Dzul9hwtLEY_IDK5TIjUuU9Sedniwcw";
      var MainDrive = DriveApp.getFolderById(FolderId);
      try {
        var CustomerFdr = MainDrive.getFoldersByName(Vendor).next();
      } catch (e) {
        var CustomerFdr = MainDrive.createFolder(Vendor);
      }

      var NewSheetId = DriveApp.getFileById(MasterFormatId).makeCopy(CustomerFdr).setName(PoNo + "//" + Vendor).getId();
      var NewSheet = SpreadsheetApp.openById(NewSheetId);
      ActiveSheet.copyTo(NewSheet);

      var sheet = NewSheet.getSheetByName('Sheet1');
      NewSheet.deleteSheet(sheet);

      var FinalNewSheet = NewSheet.getSheets()[0].setName(PoNo + "//" + Vendor).activate();
      FinalNewSheet.getRange(1, 1, FinalNewSheet.getLastRow(), FinalNewSheet.getLastColumn()).activate();
      FinalNewSheet.getActiveRangeList().setFontFamily('Calibri');

      //Delete data validations
      FinalNewSheet.getRange(1, 1, FinalNewSheet.getLastRow(), FinalNewSheet.getLastColumn()).clearDataValidations();
      FinalNewSheet.getRange(16, 1).setValue("Sr. No.");
      /*
      var IndexRange = FinalNewSheet.getRange(17, 1, FinalNewSheet.getLastRow()-10-16, 1).getValues();
      IndexRange.forEach(function(e, index){
        if(e[1]!=""){
          FinalNewSheet.getRange(17 + index, 1).setValue(index + 1);  
        }
        
      }) */

      //adding data in database sheet
      DatabaseSheet.getRange("I2:I").clearContent();
      var PoDataRange = ActiveSheet.getRange(17, 1, ActiveSheet.getLastRow()-10-16, 10).getValues();
      PoDataRange.forEach(function(c){
        if(c[1]!=""){
          var ValArr = [Vendor, VGst, PoNo, PoDate, DelTime + " days", LotType, c[0], c[1], "", c[5], c[6], c[7], c[8], c[9]];
          Logger.log(ValArr[0].length);
          DatabaseSheet.getRange(DatabaseSheet.getLastRow()+1, 1, 1, ValArr.length).setValues([ValArr]);
          
        }
      });
      DatabaseSheet.getRange("I2").setFormula(`=ArrayFormula(iferror(image(VLOOKUP($G2:$G, 'SKU Mst'!$J:$M, 4,0))))`)

      // ceating PDF for mail
      var url_base = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", SS.getId());
      var url_ext = 'export?exportFormat=pdf&format=pdf' //export as pdf/ csv. xls
        // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
        // following parameters are optional...
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
      var PoSheetBlob = Response.getBlob().setName(PoNo + " " + Vendor + ".pdf");
      
      //GmailApp.sendEmail(To, Sub, "", { cc: CC, htmlBody: RepBody, attachments: PoSheetBlob, replyTo: ReplyTo });

      GmailApp.getThreadById(ThreadId.toString()).replyAll("", { htmlBody: RepBody, attachments: PoSheetBlob });


      var OrderArr = (!PoListSheet.getLastRow() < 1) ? PoListSheet.getRange(2, 1, PoListSheet.getLastRow() - 1, 1).getValues().map(insp => { return insp.toString() }) : [];
      var Row = OrderArr.indexOf(PoNo.toString()) + 2;
      Logger.log(Row)
      PoListSheet.getRange(Row, 5).setValue([NewSheet.getUrl()]);
      PoListSheet.getRange(Row, 6).setValue(["Approved"]);
      PoListSheet.getRange(Row, 7).setValue([new Date()]);
      PoListSheet.getRange(Row, 8).setValue(UserId);
      
      SS.deleteActiveSheet();

      SS.toast("Purchase Order created succesfully.", "Successful!!");

    } else{ var Ui = SpreadsheetApp.getUi();
      Ui.alert("###ERROR###", "Till now, this Purchase Order is not requested for approval.", Ui.ButtonSet.OK )

    }

  } else { var Ui = SpreadsheetApp.getUi();
    Ui.alert("###ERROR###", "Either Wrong Sheet or Important details(i.e. Vendor, Po No., PO Date, Delivery Time, Prepared by) are missing.", Ui.ButtonSet.OK ) }

};

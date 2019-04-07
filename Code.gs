function getOutputFolder() {
    const rootFolderName = "oreshkin-ind";
    var folder = DriveApp.getFoldersByName(rootFolderName);
    if (folder.hasNext()) {
        Logger.log('/oreshkin-ind folder exists')
        return folder.next()
    } else {
        var folderNew = DriveApp.createFolder(rootFolderName);
        Logger.log('/oreshkin-ind folder created');
        return folderNew;
    }
}

function test() {
    var dict = {};

    dict['key'] = "testing";

    console.log(dict);
}

function createNextSheet(index) {
    var outputFolder = getOutputFolder();
    var newSheet = SpreadsheetApp.create("Individuals-" + index)
    var temp = DriveApp.getFileById(newSheet.getId());
    outputFolder.addFile(temp)
    DriveApp.getRootFolder().removeFile(temp);
    return newSheet;
}

function createIndSheetGroup(address) {
    var outputFolder = getOutputFolder();
    var newSheet = SpreadsheetApp.create(address)
    var temp = DriveApp.getFileById(newSheet.getId());
    outputFolder.addFile(temp)
    DriveApp.getRootFolder().removeFile(temp);
    return newSheet;
}

function createIndividualReports() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Заказы");;
    var range = sheet.getRange("C2:C" + sheet.getLastRow())
    var data = range.getValues();

    var ranges = range.getMergedRanges();

//    var orders = [];
    var reOrdered = {};//map
    for (var i = 0; i < ranges.length; i++) {
        var individualInfo = {};
        var currrange = ranges[i];

        var startRow = currrange.getRow();
        var lastRow = currrange.getLastRow();
        var column = currrange.getColumn();

        var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow)+1, 6).getValues();
        individualInfo.orderNum = neededRange[0][1];
        individualInfo.name = neededRange[0][2];
        individualInfo.address = neededRange[0][3];
        individualInfo.price = neededRange[0][5];

        individualInfo.positions = [];
        for (var k = 0; k <= (lastRow - startRow); k++) {
            // var name = names[k];
            var position = neededRange[k][0];
            individualInfo.positions[k] = position;
        }


        if (reOrdered[individualInfo.address] == null) {
            reOrdered[individualInfo.address] = [];
        }
        var aggrByAdress = reOrdered[individualInfo.address];
        aggrByAdress.push(individualInfo);
    }
  //TODO:
  
  
  
  for (var key in reOrdered) {
    var ordersForAdress = reOrdered[key];
    
    var outputSheet =  createIndSheetGroup(key); //todo: from template
    
    for (var indOrdIndex in ordersForAdress){
      //var ss = SpreadsheetApp.getActiveSpreadsheet();
      //var templateSheet = ss.getSheetByName('Sales');
      //ss.insertSheet({template: templateSheet});
      var orders = ordersForAdress[indOrdIndex];
      
      var  indSheet = outputSheet.insertSheet();
      indSheet.setName(outputSheet.name);
      
      
      indSheet.getRange("B3").setValue(orders[i].address);
      indSheet.getRange("B4").setValue(orders[i].name);
//       indSheet.getRange("B5").setValue(orders[i].);//todo: phone
      indSheet.getRange("B5").setValue(orders[i].orderNum);
      indSheet.getRange("B6").setValue(orders[i].price);
      
        for (var r = 0; r < orders[i].positions.length; r++) {
            indSheet.getRange(7 + r, 1).setValue(orders[i].positions[r]);
        }

    }
    
 
    
  }
  
  
  
    for (var i = 0; i < ranges.length; i++) {
        //   var startCorpDate = new Date();
        var newSheet;
        if (i % 50 == 0) {
            newSheet = createNextSheet(i);
        }
        var currrange = ranges[i];

        var startRow = currrange.getRow();
        var lastRow = currrange.getLastRow();
        var column = currrange.getColumn();

        var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow)+1, 6).getValues();

        var orderNum = neededRange[0][1];
        var name = neededRange[0][2];
        var address = neededRange[0][3];
        var price = neededRange[0][5];
        var positions = [];
        for (var k = 0; k <= (lastRow - startRow); k++) { //todo: check ranges!
            // var name = names[k];
            var position = neededRange[k][0];
            positions[k] = position;
        }

        orders[i] = {};
        orders[i].orderNum = orderNum;
        orders[i].name = name;
        orders[i].address = address;
        orders[i].price = price;
        orders[i].positions = positions;

        ///generate indovidial reports


        var namedSheet = newSheet.insertSheet();
        namedSheet.setName(orders[i].name + " " + orders[i].orderNum);

        namedSheet.getRange("A2").setValue(orders[i].address);
        namedSheet.getRange("A3").setValue(orders[i].name);
        namedSheet.getRange("A4").setValue(orders[i].orderNum);

        namedSheet.getRange("A5").setValue(orders[i].price);

        namedSheet.setColumnWidth(1, 200);

        namedSheet.getRange("D4").setValue("https://oreshkinspb.ru");

        for (var r = 0; r < orders[i].positions.length; r++) {
            namedSheet.getRange(7 + r, 1).setValue(orders[i].positions[r]);
        }

        var trashSheet = newSheet.getSheetByName("Sheet1");
        if (trashSheet != null) {
            newSheet.deleteSheet(trashSheet);
        }
    }
}
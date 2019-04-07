function createIndividualReports() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Заказы");;
    var range = sheet.getRange("C2:C" + sheet.getLastRow())
    var data = range.getValues();

    var ranges = range.getMergedRanges();

    var reOrdered = {}; //map
    for (var i = 0; i < ranges.length; i++) {
        var individualInfo = {};
        var currrange = ranges[i];

        var startRow = currrange.getRow();
        var lastRow = currrange.getLastRow();
        var column = currrange.getColumn();

        var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow) + 1, 10).getValues();
        individualInfo.orderNum = neededRange[0][1];
        individualInfo.name = neededRange[0][2];
        individualInfo.phone = neededRange[0][3];
        individualInfo.address = neededRange[0][4];
        individualInfo.price = neededRange[0][6];

        individualInfo.positions = [];
        for (var k = 0; k <= (lastRow - startRow); k++) {
            // var name = names[k];
            var posWithAmt = {};
            posWithAmt.position = neededRange[k][7];
            posWithAmt.amt = neededRange[k][9];
            individualInfo.positions[k] = posWithAmt;
        }


        if (reOrdered[individualInfo.address] == null) {
            reOrdered[individualInfo.address] = [];
        }
        var aggrByAdress = reOrdered[individualInfo.address];
        aggrByAdress.push(individualInfo);
    }

    var ss = SpreadsheetApp.openById("1RcNI5v_BOalBi8qckNPuLqS_A0cEgml276SPRAx5lj0");
    var templateSheet = ss.getSheetByName('template');

    for (var key in reOrdered) {
        Logger.log("Started: " + key)
        var ordersForAdress = reOrdered[key];

        var outputSheet = createIndSheetGroup(key);
        if (outputSheet.getSheetByName("Sheet1") == null) {
            Logger.log("Skippping: " + key);
            continue;
        }
        for (var indOrdIndex in ordersForAdress) {

            var orders = ordersForAdress[indOrdIndex];
            var indSheet = templateSheet.copyTo(outputSheet)

            var sName = orders.name + " " + orders.orderNum;
            if (outputSheet.getSheetByName(sName) != null) {
                Logger.log("Skippping sheet: " + sName);
                continue;
            }
            indSheet.setName(sName);

            indSheet.getRange("A2").setValue(orders.address);
            indSheet.getRange("A3").setValue(orders.name + " (" + orders.phone + ")");
            //       indSheet.getRange("B5").setValue(orders[i].);//todo: phone
            indSheet.getRange("A4").setValue(orders.orderNum);
            indSheet.getRange("A5").setValue(orders.price);

            for (var r = 0; r < orders.positions.length; r++) {
                indSheet.getRange(8 + r, 1).setValue(orders.positions[r].position);
                indSheet.getRange(8 + r, 2).setValue(orders.positions[r].amt);
                indSheet.insertRowBefore(8 + r + 1);
            }

        }
        outputSheet.deleteSheet(outputSheet.getSheetByName("Sheet1"));
        Logger.log("Finished: " + key)
    }
}

function getOutputFolder() {
    const rootFolderName = "oreshkin-ind";
    var folder = DriveApp.getFoldersByName(rootFolderName);
    if (folder.hasNext()) {
        Logger.log(rootFolderName + ' folder exists')
        return folder.next()
    } else {
        var folderNew = DriveApp.createFolder(rootFolderName);
        Logger.log(rootFolderName + ' folder created');
        return folderNew;
    }
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
    var filesWithSameName = DriveApp.getFilesByName(address);
    if (filesWithSameName.hasNext()) {
        return SpreadsheetApp.open(filesWithSameName.next());
    }
    var newSheet = SpreadsheetApp.create(address)
    var temp = DriveApp.getFileById(newSheet.getId());
    outputFolder.addFile(temp)
    DriveApp.getRootFolder().removeFile(temp);
    return newSheet;
}
function createIndividualReports() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Заказы");
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange("C2:C" + lastRow)

    var ranges = range.getMergedRanges();

    var reOrdered = {}; //map
    console.log("Iterating over merged started")
    reorderOverMerged(reOrdered, ranges, sheet);
    console.log("Iterating over merged finished")


    console.log("Iterating over not-merged started")
    var notMergedRanges = getNotMergedRanges(range, sheet);
    reorderOverMerged(reOrdered, notMergedRanges, sheet);
    console.log("Iterating over not-merged finished")
    createIndReportsFromInMemory(reOrdered);

    console.log("Finished ALL")
}

function getNotMergedRanges(range, sheet) {
    console.log("Iterating over not-merged started")
    var notMergedRanges = [];
    for (var i = 1; i <= range.getLastRow() - 1; i++) { //C2
        var ccc = range.getCell(i, 1);
        if (ccc.isPartOfMerge()) {
            continue;
        }
        var cellRange = sheet.getRange(ccc.getRow(), ccc.getColumn());
        notMergedRanges.push(cellRange);
    }
    return notMergedRanges;
}

function createIndReportsFromInMemory(reOrdered) {
    var ss = SpreadsheetApp.openById("1RcNI5v_BOalBi8qckNPuLqS_A0cEgml276SPRAx5lj0");
    var templateSheet = ss.getSheetByName('template');

    for (var key in reOrdered) {
        console.log("Started: " + key)
        var ordersForAdress = reOrdered[key];
        var outputSheet = createIndSheetGroup( "ind-"+key);

        for (var indOrdIndex in ordersForAdress) {
            var orders = ordersForAdress[indOrdIndex];
            var sName = orders.name + " " + orders.orderNum;
            if (outputSheet.getSheetByName(sName) != null) {
                console.log("Skippping: " + sName + " in " + key);
                continue;
            }

            var indSheet = templateSheet.copyTo(outputSheet)

            indSheet.setName(sName);

            indSheet.getRange("A5").setValue(orders.address);
            indSheet.getRange("A2").setValue(orders.name + " (" + orders.phone + ")");
            //       indSheet.getRange("B5").setValue(orders[i].);//todo: phone
            indSheet.getRange("A3").setValue(orders.orderNum);
            indSheet.getRange("A4").setValue(orders.price);

            for (var r = 0; r < orders.positions.length; r++) {
                var posCell = indSheet.getRange(8 + r, 1);
                posCell.setValue(orders.positions[r].position);
                posCell.setBorder(true, true, true, true, false, false);
                var amtCell = indSheet.getRange(8 + r, 2);
                amtCell.setBorder(true, true, true, true, false, false);
                amtCell.setValue(orders.positions[r].amt);
                indSheet.insertRowBefore(8 + r + 1);
            }

        }
        var toDelete = outputSheet.getSheetByName("Sheet1");
        if (toDelete != null) {
            outputSheet.deleteSheet(toDelete);
        }
        console.log("Finished: " + key)
    }
}

function reorderOverMerged(reOrdered, ranges, sheet) {
    for (var i = 0; i < ranges.length; i++) {
        var individualInfo = {};
        var currrange = ranges[i];

        var startRow = currrange.getRow();
        var lastRow = currrange.getLastRow();
        var column = currrange.getColumn();

        var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow) + 1, 11).getValues();
        individualInfo.orderNum = neededRange[0][1];
        individualInfo.name = neededRange[0][2];
        individualInfo.phone = neededRange[0][3];
        individualInfo.address = neededRange[0][4];
        individualInfo.price = neededRange[0][6];

        individualInfo.positions = [];
        for (var k = 0; k <= (lastRow - startRow); k++) {
            var posWithAmt = {};
            posWithAmt.position = neededRange[k][7];
            posWithAmt.amt = neededRange[k][10];
            individualInfo.positions[k] = posWithAmt;
        }

        if (reOrdered[individualInfo.address] == null) {
            reOrdered[individualInfo.address] = [];
        }
        var aggrByAdress = reOrdered[individualInfo.address];
        aggrByAdress.push(individualInfo);
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

function createIndSheetGroup(address) {
    var filesWithSameName = DriveApp.getFilesByName(address);
    if (filesWithSameName.hasNext()) {
        return SpreadsheetApp.open(filesWithSameName.next());
    }
    var outputFolder = getOutputFolder();
    var newSheet = SpreadsheetApp.create(address)
    var temp = DriveApp.getFileById(newSheet.getId());
    outputFolder.addFile(temp)
    DriveApp.getRootFolder().removeFile(temp);
    return newSheet;
}
function dumpJson() {
  	console.log("Scrtipt has started") 
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sheet = ss.getSheetByName("Заказы");
	var lastRow = sheet.getLastRow();
	var range = sheet.getRange("B2:B" + lastRow);


    var documentProperties = PropertiesService.getDocumentProperties();
    { //unmerged rows
		var notMergedRanges = [];
		var lastProcessedRow = documentProperties.getProperty("lastProcessedRowS");
		if (lastProcessedRow == null) {
			lastProcessedRow = 1;
			console.log("Starting unmerged rows from beginning")
		} else {
			console.log("recovering from unmerged row " + lastProcessedRow)
		}
		for (var i = lastProcessedRow; i <= range.getLastRow() - 1; i++) { //C2
			var ccc = range.getCell(i, 1);
			if (ccc.isPartOfMerge()) {
              documentProperties.setProperty("lastProcessedRowS", i);		
              continue;
			}
			var cellRange = sheet.getRange(ccc.getRow(), ccc.getColumn());
			precessOneNotMergedRow(sheet, cellRange);
			documentProperties.setProperty("lastProcessedRowS", i)
		}
	}

	//merged
	{
		console.log("Starting merged rows")
		var mergedRanges = range.getMergedRanges();
            
      //documentProperties.deleteProperty("lastProcessedRowM")//todo: deltee
      var lastProcessedRow = documentProperties.getProperty("lastProcessedRowM");

      
		if (lastProcessedRow == null) {
			lastProcessedRow = 0;
			console.log("Starting merged rows from beginning")
		} else {
			console.log("recovering from merged row " + lastProcessedRow)
		}
   

		for (var i = lastProcessedRow; i < mergedRanges.length; i++) {
			var individualInfo = {};
			var currrange = mergedRanges[Number(i)];
    

			var startRow = currrange.getRow();
			var lastRow = currrange.getLastRow();
			var column = currrange.getColumn();

			var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow) + 1, 11).getValues();
			individualInfo.orderNum = neededRange[0][1].substring(1);
			individualInfo.name = neededRange[0][2];
			individualInfo.phone = neededRange[0][3];
			individualInfo.address = neededRange[0][4];
			individualInfo.price = neededRange[0][6];

			individualInfo.positions = [];
			for (var k = 0; k <= (lastRow - startRow); k++) {
				var posWithAmt = {};
				posWithAmt.position = neededRange[k][7];
				posWithAmt.amt = neededRange[k][10];
				individualInfo.positions[k] = posWithAmt;
			}

			var jsoned = JSON.stringify(individualInfo, null, 4);
			var fileName = individualInfo.orderNum + ".json";
			console.log("File M1 " + fileName + " created")

			createOrAppendFile(individualInfo.orderNum + ".json", jsoned)
			documentProperties.setProperty("lastProcessedRowM", (i+0))
		}
      console.log("Processing merged finished")
	}

}


function precessOneNotMergedRow(sheet, notMergedRange) {
	var documentProperties = PropertiesService.getDocumentProperties();
	var individualInfo = {};
	var currrange = notMergedRange;

	var startRow = currrange.getRow();
	var lastRow = currrange.getLastRow();
	var column = currrange.getColumn();

	var neededRange = sheet.getRange(startRow, 1, (lastRow - startRow) + 1, 11).getValues();
	individualInfo.orderNum = neededRange[0][1].substring(1);

	if (documentProperties.getProperty(individualInfo.orderNum) != null) {
		console.log("skipping 1 " + individualInfo.orderNum);
		return;
	} else {
		documentProperties.setProperty(individualInfo.orderNum, "");
	}

	individualInfo.name = neededRange[0][2];
	individualInfo.phone = "" + neededRange[0][3];
	individualInfo.address = neededRange[0][4];
	individualInfo.price = neededRange[0][6];

	individualInfo.positions = [];
	for (var k = 0; k <= (lastRow - startRow); k++) {
		var posWithAmt = {};
		posWithAmt.position = neededRange[k][7];
		posWithAmt.amt = neededRange[k][10];
		individualInfo.positions[k] = posWithAmt;
	}
	var jsoned = JSON.stringify(individualInfo, null, 4);
	var fileName = individualInfo.orderNum + ".json";
	console.log("File S1 " + fileName + " created")

	createOrAppendFile(individualInfo.orderNum + ".json", jsoned)
}

function createOrAppendFile(fileName, content) {
	var folderName = "jsons";

	// get list of folders with matching name
	var folderList = DriveApp.getFoldersByName(folderName);
	if (folderList.hasNext()) {
		// found matching folder
		var folder = folderList.next();

		// search for files with matching name
		var fileList = folder.getFilesByName(fileName);

		if (fileList.hasNext()) {
			// found matching file - append text
			var file = fileList.next();
			//file.setContent(content);
		} else {
			// file not found - create new
			folder.createFile(fileName, content);
		}
	}
}

function clearProps() {
	var documentProperties = PropertiesService.getDocumentProperties();
	documentProperties.deleteAllProperties();
}

function getProps() {
	var documentProperties = PropertiesService.getDocumentProperties();
	return documentProperties.getKeys();
}
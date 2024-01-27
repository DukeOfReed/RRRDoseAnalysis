//Called every ~3 minuites, data has been saved and the last itteration of the script terminated, calls MakeSpreadsheet with its optional parameter to start the sheet running again! 
function SheetKickStart()
{
  //      LOCAL VARIABLES
  //Get Database "SERVER CHECKPOINT"
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var serverCheckPoint = sheetActive.getSheetByName("SERVER CHECKPOINT");

  //Checkpoint Functions
  if(serverCheckPoint.getRange(2, 3).getCell(1, 1).getValue() == "Waiting For Kick")
  {
    //If C2 of SERVER CHECKPOINT is set to "Waiting for Kick" set it to nil, then load all the data stored in the spreadsheet, clearing as you go
    serverCheckPoint.getRange(2, 3).getCell(1, 1).setValue("")
    var loadingData = [];
    //"loading data" LD[0] is the name of the spreadsheet being saved
    loadingData.push(serverCheckPoint.getRange(2, 1).getCell(1, 1).getValue());
    serverCheckPoint.getRange(2, 1).getCell(1, 1).setValue("");
    //"loading data" LD[1] is the year the spreadsheet starts
    loadingData.push(serverCheckPoint.getRange(3, 1).getCell(1, 1).getValue());
    serverCheckPoint.getRange(3, 1).getCell(1, 1).setValue("");
    //"loading data" LD[2] is the number of quarters being analyzed
    loadingData.push(serverCheckPoint.getRange(4, 1).getCell(1, 1).getValue());
    serverCheckPoint.getRange(4, 1).getCell(1, 1).setValue("");
    //"loading data" LD[3] is an array whos 0th element is the step that proscessing stopped on. how many rows of data in have ~been~ calculated. 
    loadingData.push([serverCheckPoint.getRange(2,2).getCell(1, 1).getValue()]);
    serverCheckPoint.getRange(2,2).getCell(1, 1).setValue("");
    var gatheringData = true;
    var i = 5;
    //Starting at A5 we copy all the remaining data into LD[3][0].
    while(gatheringData)
    {
      if(serverCheckPoint.getRange(i, 1).getCell(1, 1).getValue() != "")
      {
        loadingData[3].push(serverCheckPoint.getRange(i, 1).getCell(1, 1).getValue());
        serverCheckPoint.getRange(i, 1).getCell(1, 1).setValue("");
        console.log(loadingData[3][i]);
        i++;
      } else {
        gatheringData = false;
      }
    }

    MakeSpreadsheet(loadingData);
  }

}

//Checks if 10 minuites of runtime have elapsed. If not, updates last checkpoint for debugging and continues on. If so, stores run data for SheetKickStart to put in a pickUp
function CheckPoint(startTime, currentFunction, storedData)
{
  //      LOCAL VARIABLES
  //Get Database "SERVER CHECKPOINT"
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var serverCheckPoint = sheetActive.getSheetByName("SERVER CHECKPOINT");
  var currentRuntime = new Date();
  var timeout = 15;

  //Update Checkpoint Log
  serverCheckPoint.getRange(2,4).getCell(1, 1).setValue("Last Checkpoint on function " + currentFunction);
  serverCheckPoint.getRange(3,4).getCell(1, 1).setValue("Last Checkpoint at time " + currentRuntime);
  serverCheckPoint.getRange(4,4).getCell(1, 1).setValue("Script start-time was " + startTime);
  serverCheckPoint.getRange(5,4).getCell(1, 1).setValue("Time until next save: " + (timeout - getMinDiff(startTime, currentRuntime)));

  //Check if 10 mins of run-time have elapsed. We will choose checkpoints such that any step of this sheet can run alone in 20 mins. 
  if(Math.abs(getMinDiff(startTime, currentRuntime)) > timeout)
  {
    serverCheckPoint.getRange(2,3).getCell(1, 1).setValue("Waiting For Kick");
    serverCheckPoint.getRange(2,2).getCell(1, 1).setValue(currentFunction);
    for(i = 0; i < storedData.length; i++)
    {
      console.log("Stored data length: " + storedData.length + " at item: " + i);
      serverCheckPoint.getRange(i+2,1).getCell(1, 1).setValue(storedData[i]);
    }
    TerminateScript();
  }
}

//Throws an Error to terminate script after checkpoint data is saved.
function TerminateScript()
{
  throw new Error("Timeout Pending. Saved Project waiting for Kick Start")
}

//Checks if some cell has the same value as a report/sheet on file.
function IsUniqueName(cell_reference) 
{
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = sheetActive.getSheets();
  
  if(cell_reference == "") 
  {
    return false;
  }

  for (var i=0 ; i<allSheets.length ; i++) 
  {
    if(allSheets[i].getName() == cell_reference)
    {
      return false;
    }
  }
  
  return true;
}

//Gets the Modulus of a by b: divMod(a, b) = a % b.
function divMod(a, b) 
{
  return [Math.trunc(a / b), a % b];
}

//No idea what this does. It is used in indexToCol though!
function divModSheets(n)
{
		var ab = divMod(n, 26)
    if(ab[1] == 0)
    {
        return [ab[0] - 1, ab[1] + 26]
    }
    return ab
}

//Takes a column number and gives the string column ref for making cell formula.
function indexToCol(num)
{
	var alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
	var chars = [];
  var numD = [num, 0];
  while(numD[0] > 0)
  {
  	numD = divModSheets(numD[0])
    chars = chars.concat(alphabet[numD[1] - 1])
  }
  return chars.reverse().join("");
}

//Gets the diference between two data objects in minuites.
function getMinDiff(t1, t2)
{
  var hd=t1.valueOf();
  var td=t2.valueOf();
  var sec=1000;
  var min=60*sec;
  var hour=60*min;
  var day=24*hour;
  var diff=td-hd;
  var days=Math.floor(diff/day);
  var hours=Math.floor(diff%day/hour);
  var minutes=Math.floor(diff%day%hour/min);
  return minutes
}

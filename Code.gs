function appendTrackingRowToSheet(sheetName, volunteer, location)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.appendRow([volunteer, location]);
}


function appendDurationRowToSheet(sheetName, accepted, volunteers)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.appendRow([accepted, volunteers]);
}


function appendResponseTimeRowToSheet(sheetName, dayNum, hourNum)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.appendRow([dayNum, hourNum]);
}



function appendReviewRowToSheet(sheetName, description, volunteer)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.appendRow([description, "", false, volunteer]);
}

function appendVolunteerToScoreList(sheetName, volunteer)
{
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheetByName(sheetName);
 sheet.appendRow([volunteer, 0]);
}

function deleteLastRow(sheetName)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  sheet.deleteRow(rows);

}

function appendVolunteerToBlackList(sheetName, volunteer)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")
  sheet.appendRow([volunteer, date]);
}

function checkIfExists(sheetName, volunteer)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  for (var row=2; row <= rows; row++) {
    var mailCell=selection.getCell(row,1);
    if (mailCell.getValue()===volunteer)
      return true;
  }
  return false;
}


function updateScore(sheetName, volunteer)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  for (var row=2; row <= rows; row++) {
    var mailCell=selection.getCell(row,1);
    if (!mailCell.isBlank() && mailCell.getValue()===volunteer) {
      var scoreCell=selection.getCell(row,2);
      scoreCell.setValue(scoreCell.getValue()+1);
    }
  }
}

function blacklistScore(sheetName, volunteer)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  for (var row=2; row <= rows; row++) {
    var mailCell=selection.getCell(row,1);
    if (!mailCell.isBlank() && mailCell.getValue()===volunteer) {
      var scoreCell=selection.getCell(row,2);
      scoreCell.setValue(-1);
    }
  }
}




function getScore(sheetName, volunteer)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  for (var row=2; row <= rows; row++) {
    var mailCell=selection.getCell(row,1);
    if (!mailCell.isBlank() && mailCell.getValue()===volunteer) {
      var scoreCell=selection.getCell(row,2);
      var score=scoreCell.getValue();
      return score;
    }
  }
  return 0;
}
  


function getNumberOfVolunteers(sheetName)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  return rows-1;
}


function getNumberOfAcceptedRequests(sheetName)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  var count=0;
  var acceptStatus=false;
  for (var row=2; row <= rows; row++) {
    acceptStatus=selection.getCell(row,5).getValue();
    if(acceptStatus==true){
      count=count+1;
    }
  }
  return count;
}


function payReward(sheetName, volunteer, price)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  var selection=sheet.getDataRange();
  var columns=selection.getNumColumns();
  var rows=selection.getNumRows();
  for (var row=2; row <= rows; row++) {
    var mailCell=selection.getCell(row,1);
    if (!mailCell.isBlank() && mailCell.getValue()===volunteer) {
      var scoreCell=selection.getCell(row,2);
      var score=scoreCell.getValue()-price;
      scoreCell.setValue(score);
      return 1;
    }
  }
  return 0;
}



function onReview(e)
{
  var spreadS = e.source;
  var sheet = spreadS.getActiveSheet();
  var sheetName = sheet.getName();
  var activeRange = sheet.getActiveRange();
  var activeRow = activeRange.getRow();
  var activeColumn = activeRange.getColumn();
  if(sheetName === "Reviews"){
    var reportStatus=sheet.getRange(activeRow,3).getValue();
    var volunteerMail = sheet.getRange(activeRow, 4).getValue();
    if(reportStatus==true)
    {
      appendVolunteerToBlackList("Blacklist",volunteerMail);
      blacklistScore("Score",volunteerMail);
    }
  }
}


function onRewardClaim(e)
{
  var spreadS = e.source;
  var sheet = spreadS.getActiveSheet();
  var sheetName = sheet.getName();
  var activeRange = sheet.getActiveRange();
  var activeRow = activeRange.getRow();
  var activeColumn = activeRange.getColumn();
  if(sheetName === "Rewards"){
    var claimAddress=sheet.getRange(activeRow,4).getValue();
    if(claimAddress!="")
    {
      var volunteerScore=getScore("Score", claimAddress);
      var rewardItem=sheet.getRange(activeRow,1).getValue();
      var neededScore=sheet.getRange(activeRow,3).getValue();
      if(volunteerScore>=neededScore)
      {
        var subject=rewardItem;
        var message=generateDiscountCode();
        payReward("Score",claimAddress,neededScore);
        MailApp.sendEmail(claimAddress, subject, message);
        sheet.getRange(activeRow,4).setValue("");
      }
    }
  }
}

function generateDiscountCode()
{
  var first=(Math.random()*100);
  var second=(Math.random()*100);
  var third=(Math.random()*100);
 // var currentDate = Utilities.formatDate(new Date(), "GMT", "### EEEE - dd/MM/yyyy")
  //return first;
  //return first+second+third;
 // return getCurrentDayNumber()+" "+getCurrentHourNumber();
 return "COUPON"+getCurrentDayNumber()+getCurrentHourNumber();
}

function getCurrentDayNumber()
{
  var currentDate = Utilities.formatDate(new Date(), "GMT+1", "### EEEE - dd/MM/yyyy");
  var day = currentDate.split(" ");
  if(day[1]==="Monday") return 1;
  if(day[1]==="Tuesday") return 2;
  if(day[1]==="Wednesday") return 3;
  if(day[1]==="Thursday") return 4;
  if(day[1]==="Friday") return 5;
  if(day[1]==="Saturday") return 6;
  if(day[1]==="Sunday") return 7;
}


function getCurrentHourNumber() {
  var currentDate = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd'*'HH");
  var h=currentDate.split("*");
  return h[1];
}




function onRequestStatusModification(e) {
  var spreadS = e.source;
  var sheet = spreadS.getActiveSheet();
  var sheetName = sheet.getName();
  var activeRange = sheet.getActiveRange();
  var activeRow = activeRange.getRow();
  var activeColumn = activeRange.getColumn();
  


  if(sheetName === "Requests"){
    var volunteerMail = sheet.getRange(activeRow, 7).getValue();
    var completeStatus=sheet.getRange(activeRow,8).getValue();
    if(completeStatus==true)
    {
      var description=sheet.getRange(activeRow,3).getValue();
      appendReviewRowToSheet("Reviews",description,volunteerMail);
      updateScore("Score", volunteerMail);
      deleteLastRow("PredictAcceptDuration");
      var totalVolunteersAvailable=getNumberOfVolunteers("Score")-getNumberOfVolunteers("Blacklist");
      appendDurationRowToSheet("PredictAcceptDuration",getNumberOfAcceptedRequests("Requests"),totalVolunteersAvailable);
    }
    else
    {
      var acceptStatus=sheet.getRange(activeRow,5).getValue();
      if(acceptStatus==true)
      {
        var exists=checkIfExists("Score",volunteerMail);
        if(!exists)
          appendVolunteerToScoreList("Score",volunteerMail);

        var destinationEmail = sheet.getRange(activeRow, 6).getValue();
        var location=sheet.getRange(activeRow,1).getValue();

        var subject="Help accepted";
        var message="A volunteer will arrive soon!";
        deleteLastRow("PredictAcceptDuration");
       // appendDurationRowToSheet("PredictAcceptDuration",getNumberOfAcceptedRequests("Requests"), getNumberOfVolunteers("Score"));

        var totalVolunteersAvailable=getNumberOfVolunteers("Score")-getNumberOfVolunteers("Blacklist");
        appendDurationRowToSheet("PredictAcceptDuration",getNumberOfAcceptedRequests("Requests"),totalVolunteersAvailable);
      // var message="Estimated time:"+estimatedTime+"from"+volunteerMail;
        appendTrackingRowToSheet("Tracking", volunteerMail, location);
        //updateScore("Score",volunteerMail);
        MailApp.sendEmail(destinationEmail, subject, message);
      

      }
    }
  }
}







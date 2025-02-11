function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('طلب إجازة');
  menu.addItem('بالتاريخ', 'vecationByDate');
  menu.addItem('بالأيام ابداءً من اليوم', 'vecationByDays');
  menu.addToUi();

}

function vecationByDate(){
  
  const activeSheet = SpreadsheetApp.getActiveSheet();
  var cell = activeSheet.getActiveCell();
  var row = cell.getRow();
  
  var empLeaveDays = activeSheet.getRange('F' + row).getValue();
  // console.log('Current row: '+ row);
  // console.log('Current cell: '+ cell.getA1Notation());
  console.log('Emp leave days: '+ empLeaveDays);

  
  var ui = SpreadsheetApp.getUi(); 

  var result = ui.prompt(
      'Enter Vacation Date',
      'Enter Start and End Dates as (dd/mm/yyyy - dd/mm/yyyy)',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var vacationDate = result.getResponseText();
  if (button == ui.Button.OK) {
    
    // verifyDates(vacationDate, empLeaveDays)
    var empName = activeSheet.getRange('A' + row).getValue();
    var empID = activeSheet.getRange('C' + row).getValue();
    
    let empData = verifyDates(vacationDate, empLeaveDays);
    console.log(empData);
    confirmVacation(empData[0], empData[1], empData[2], empName, empID)
    SpreadsheetApp.getActiveSheet().getRange('F' + row).setValue(empLeaveDays - empData[2]);

  }

}

function verifyDates(dates, empLeaveDays) {  
  const regex = /\b(\d{1,2}\/\d{1,2}\/\d{4})\s*-\s*(\d{1,2}\/\d{1,2}\/\d{4})\b/; 
  let ui = SpreadsheetApp.getUi();
  // Extract dates
  const match = regex.exec(dates);
  if (!match || !regex.test(dates)) {
    ui.alert("Date range is not valid please Try Again");
    return;
  }

  const startDateStr = match[1]; // Start date as string
  const endDateStr = match[2];   // End date as string

  console.log('Vacation Date From', parseDate(startDateStr), 'TO', parseDate(endDateStr))

  let startDate = parseDate(startDateStr);
  let endDate = parseDate(endDateStr);

  // check dates string exists
  if (!startDate || !endDate) {
    ui.alert("⚠️" + startDate + " - " + endDate + "Invalid date range format. Please Try Again");
    return;
  }
  // check if date is in the past
  if (isPastDate(startDate)) {
    ui.alert("⚠️ Start date is in the past.Please Enter Valid dates");
    return;
  }
  // check if start date exceed end date
  if (startDate > endDate) {
    ui.alert("⚠️❗Start date cannot be after end date.Please Enter Valid dates");
    return;
  }
  
  const numOfDays = (endDate - startDate) /  (1000 * 60 * 60 * 24);
  
  // check number of Days
  if ( numOfDays <= 0){
    ui.alert("⚠️❗Error, please enter valid dates");
    return;
  }

  // check if employee have enough leave days
  if ( numOfDays >= empLeaveDays){
    ui.alert("⚠️❗This Employee doesn't have enough leave days❗");
    return;

  }

  return [startDateStr, endDateStr, numOfDays];
  
}
function confirmVacation (startDate, endDate, numOfDays, empName, empID){
  let ui = SpreadsheetApp.getUi();
  
  ui.alert(`✅ ${numOfDays} Leave Days confirmed from ${startDate} To ${endDate} for ${empName} with ID: ${empID}`)


}
function parseDate(dateStr) {
  const [day, month, year] = dateStr.split('/').map(Number);
  const date = new Date(year, month - 1, day); // Month is 0-indexed in JavaScript
  // Check if the date is valid
  if (isNaN(date.getTime())) {
    return null; // Invalid date
  }
  return date;
}

function isPastDate(date) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Remove time part for comparison
  return date < today;
}

function vecationByDays(){
  let ui = SpreadsheetApp.getUi();
  ui.alert("🔜 Coming Soon 🔜")
}

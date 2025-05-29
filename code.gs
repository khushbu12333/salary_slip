/**
 * Sample Google Apps Script file.
 * You can use this script to automate tasks in Google Sheets, Docs, etc.
 */

//code.gs
function createAndSendSalSlips() {
  var empID = "";
  var empName = "";
  var empEmail = "";
  var empDesg = "";
  var empDept = "";

  var noOfDaysWorked = 0;
  var salMonth = "";

  var basicSal = 0;
  var hra = 0;
  var ConveyanceAllowence = 0;
  var MealAllowence = 0;
  var MedicalAllowence = 0;
  var PersonalPay = 0;

  var proffTax = 0;
  var Advance = 0;
  var totalIncome = 0;
  var totalDeduction = 0;
  var netSal = 0;
  var InWord = 0;
  var BankName = 0;
  var IFSCCode = 0;
  var AccNO = 0;
  var CurrentSalary = 0;

  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var salSheet = spSheet.getSheetByName("Salary");

  var salaryDetailsFolder = DriveApp.getFolderById("1jdHyhb2zfzhxfTgzoKsg0NfxJEcF84m5");
  var salaryTemplate = DriveApp.getFileById("1jkbAQJ6k7u0S6x5HgBfTDgd5n5b9TIxp210bBZVSTWs");

  var totalRows = salSheet.getLastRow();

  for (var rowNo = 2; rowNo <= totalRows; rowNo++) {
    empID = salSheet.getRange("A" + rowNo).getDisplayValue();
    empName = salSheet.getRange("B" + rowNo).getDisplayValue();
    empDesg = salSheet.getRange("C" + rowNo).getDisplayValue();
    empDept = salSheet.getRange("D" + rowNo).getDisplayValue();
    empEmail = salSheet.getRange("E" + rowNo).getDisplayValue();
    salMonth = salSheet.getRange("F" + rowNo).getDisplayValue();
    noOfDaysWorked = salSheet.getRange("H" + rowNo).getDisplayValue();

    basicSal = salSheet.getRange("I" + rowNo).getDisplayValue();
    hra = salSheet.getRange("J" + rowNo).getDisplayValue();
    ConveyanceAllowence = salSheet.getRange("K" + rowNo).getDisplayValue();
    MealAllowence = salSheet.getRange("L" + rowNo).getDisplayValue();
    MedicalAllowence = salSheet.getRange("M" + rowNo).getDisplayValue();
    PersonalPay = salSheet.getRange("N" + rowNo).getDisplayValue();
    totalIncome = salSheet.getRange("O" + rowNo).getDisplayValue();
    proffTax = salSheet.getRange("P" + rowNo).getDisplayValue();
    Advance = salSheet.getRange("Q" + rowNo).getDisplayValue();
    totalDeduction = salSheet.getRange("R" + rowNo).getDisplayValue();
    netSal = salSheet.getRange("S" + rowNo).getDisplayValue();
    InWord = salSheet.getRange("T" + rowNo).getDisplayValue();
    BankName = salSheet.getRange("U" + rowNo).getDisplayValue();
    IFSCCode = salSheet.getRange("V" + rowNo).getDisplayValue();
    AccNO = salSheet.getRange("W" + rowNo).getDisplayValue();
    CurrentSalary = salSheet.getRange("X" + rowNo).getDisplayValue();

    // Skip if email is missing or blank
    if (!empEmail || empEmail.trim() === "") {
      Logger.log("Row " + rowNo + ": No email found for employee " + empName + " (ID: " + empID + "). Skipping.");
      continue;
    }

    var rawSalFile = salaryTemplate.makeCopy(salaryDetailsFolder);
    var rawFile = DocumentApp.openById(rawSalFile.getId());
    var rawFileContent = rawFile.getBody();

    rawFileContent.replaceText("EMP_ID_XXXX", empID);
    rawFileContent.replaceText("EMP_NAME_XXXX", empName);
    rawFileContent.replaceText("DESG_XXXX", empDesg);

    rawFileContent.replaceText("MONTH_XXXX", salMonth);
    rawFileContent.replaceText("DAYS_XXXX", noOfDaysWorked);
    rawFileContent.replaceText("DEPT_XXXX", empDept);

    rawFileContent.replaceText("BASIC_SAL_XXXX", basicSal);
    rawFileContent.replaceText("HRA_XXXX", hra);
    rawFileContent.replaceText("CA_XXXX", ConveyanceAllowence);
    rawFileContent.replaceText("MA_XXXX", MealAllowence);
    rawFileContent.replaceText("MLA_XXXX", MedicalAllowence);
    rawFileContent.replaceText("PA_XXXX", PersonalPay);

    rawFileContent.replaceText("PT_XXXX", proffTax);
    rawFileContent.replaceText("ADV_XXXX", Advance);

    rawFileContent.replaceText("TI_XXXX", totalIncome);
    rawFileContent.replaceText("TD_XXXX", totalDeduction);
    rawFileContent.replaceText("NSP_XXXX", netSal);
    rawFileContent.replaceText("WO_XXXX", InWord);
    rawFileContent.replaceText("BN_XXXX", BankName);
    rawFileContent.replaceText("IFSC_XXXX", IFSCCode);
    rawFileContent.replaceText("AC_XXXX", AccNO);
    rawFileContent.replaceText("BN_XXXX", CurrentSalary);

    rawFile.saveAndClose();
    var salSlip = rawFile.getAs(MimeType.PDF);
    var salPDF = salaryDetailsFolder.createFile(salSlip).setName("Salary_" + empID);

    salaryDetailsFolder.removeFile(rawSalFile);

    var mailSubject = "Salary Slip";
    var mailBody = "Please find the salary slip for the month of " + salMonth + " attached.";

    GmailApp.sendEmail(empEmail, mailSubject, mailBody, {
      attachments: [salPDF.getAs(MimeType.PDF)]
    });
  }
}
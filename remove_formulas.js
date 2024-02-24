function main(workbook: ExcelScript.Workbook) {
  // copy and paste as values entire excel sheets
  let sheetNames = ["Sheet1", "Sheet2", "Sheet3"];
  
  for (let i = 0; i < sheetNames.length; i++) {
    let sheet = workbook.getWorksheet(sheetNames[i]);
    let rangeToCopy = sheet.getUsedRange();
    rangeToCopy.copyFrom(rangeToCopy, ExcelScript.RangeCopyType.values, false);
  }
  
  workbook.save();
}

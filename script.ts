/**
* This script copies the values from a current worksheet to another selected worksheet.
*/
function main(workbook: ExcelScript.Workbook) {

  let main_worksheet = 'TOUT';

  let usedRange = workbook.getActiveWorksheet().getUsedRange();
  // Select the mainworksheet
  let newSheet = workbook.getWorksheet(main_worksheet)
  // Copy the values from the used range to the new worksheet.
  let copyType = ExcelScript.RangeCopyType.values;
  let targetRange = newSheet.getRangeByIndexes(
    usedRange.getRowIndex(),
    usedRange.getColumnIndex(),
    usedRange.getRowCount(),
    usedRange.getColumnCount()
  );
  targetRange.copyFrom(usedRange, copyType);

  // Switch the view to the main worksheet.
  newSheet.activate();
}

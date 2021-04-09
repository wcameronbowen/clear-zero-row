function main(workbook: ExcelScript.Workbook)
{
  let range = workbook.getActiveWorksheet().getRange("A1:D34");
  let rangeValues = range.getValues();

  rangeValues.forEach((rowItem, rowIndex) => {

    rangeValues[rowIndex].forEach((columnItem, columnIndex) => {
      let columnValue = columnItem as number;
      if (columnValue.valueOf() == 0) {
        if (columnIndex == 3) {
          range.getRow(rowIndex).delete(ExcelScript.DeleteShiftDirection.up);
        }
      }
    });
  });
}

import { z } from "zod";
const RowSchema = z.object({
  ID: z.number(),
  Category: z.string(),
  Subcategory: z.string(),
  Price: z.number(),
});

export const DataValidation = async (args) => {
  try {
    Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();
      const changedCellIndex = args.address.substring(1);
      //console.log("The Change is the address is", changedCellIndex);
      const entireRow = sheet.getRange(`A${changedCellIndex}:D${changedCellIndex}`);
      entireRow.load("values");
      await context.sync();
      //console.log("The Entire Row is", entireRow.values);
      const anyEmptyStringInRow = entireRow.values[0].some((value) => typeof value === "string" && value === "");
      //console.log("Any Empty String value is", anyEmptyStringInRow);
      if (!anyEmptyStringInRow) {
        const tables = sheet.tables.load("name, items");
        await context.sync();
        const requiredTable = tables.items.find((table) => table.name === "Table2");
        const tableRange = requiredTable.getRange().getUsedRange();
        tableRange.load("values");
        await context.sync();
        const tableData = tableRange.values;
        const tableHeaders = tableData[0];
        const dataArray = tableData.slice(1).map((row) => {
          const obj = {};
          tableHeaders.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
        const range = sheet.getUsedRange();
        range.load("values");
        await context.sync();
        dataArray.forEach(async (data, index) => {
          try {
            RowSchema.parse(data);
            const cellRange = sheet.getRange(`A${index + 2}:D${index + 2}`);
            // cellRange.format.fill.load();
            // await context.sync();
            // cellRange.format.fill.load("color")
            // await context.sync()
            // if(cellRange.format.fill.color){
              cellRange.format.fill.clear();
              cellRange.format.font.color = "black";
              await context.sync();
              console.log("Changing the Format Properties");
          } catch (error) {
            console.log("The Error in Some Cell is", error);
            const cellRange = sheet.getRange(`A${index + 2}:D${index + 2}`);
            cellRange.format.fill.color = "Red";
            cellRange.format.font.color = "white";
            await context.sync();
            console.log("Something is Wrong in the Row", index + 2);
          }
        });
      }
    });
  } catch (error) {
    console.log("Error:", error);
  }
};

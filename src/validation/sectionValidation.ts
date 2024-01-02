import { z } from "zod";
const FirstSectionSchema = z.object({
  "SKU Description": z.string(),
  "Customer": z.string(),
  "Customer SKU ID": z.union([z.number(),z.string()]),
  "Status": z.string(),
  "EAN": z.union([z.number(), z.string()]),
  "GCAS": z.union([z.string(), z.number()]),
  "Category": z.string(),
  "Brand": z.string(),
  "Segment": z.string(),
  "Sub-Segment": z.string(),
});

export const SectionValidation =  () => {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Data 4");
    if (sheet) {
      const selectedrange = sheet.getRange("A:A").getUsedRange();
      selectedrange.load("values");
      await context.sync();
      let cellIndex = 0;
      for (let i = 0; i < selectedrange.values.length; i++) {
        if (selectedrange.values[i].toString() == "SKU Description") {
          cellIndex = i;
          break;
        }
      }
      console.log("The Cell Index value is", cellIndex + 1);
      const headerRange = sheet.getRange(`${cellIndex+1}:${cellIndex+1}`).getUsedRange();
      headerRange.load("values");
      await context.sync();
      const headerValues = headerRange.values[0].slice(0, 11);
      const kCellRange = sheet.getRange("K:K").getUsedRange();
      await context.sync();
      kCellRange.load("address");
      await context.sync();
      const lastCellindex = kCellRange.address.split(":")[1].slice(1);
      const valueRange = sheet.getRange(`A11:K${lastCellindex}`);
      valueRange.load("values");
      await context.sync();
      const dataValues = valueRange.values;
      console.log("The Total Values here is ", dataValues);
      const dataArray = dataValues.map((row) => {
        const obj = {};
        headerValues.forEach((header, index) => {
          obj[header] = row[index];
        });
        return obj;
      });

      console.log("The Object is", dataArray);
      dataArray.forEach(async (data, index) => {
        try {
          FirstSectionSchema.parse(data);
          const cellRange = sheet.getRange(`A${index + 12}:K${index + 12}`);
          cellRange.format.fill.clear();
          cellRange.format.font.color = "black";
          await context.sync();
          console.log("Changing the Format Properties");
        } catch (error) {
          console.log("The Error in Some Cell is", error);
          const cellRange = sheet.getRange(`A${index + 12}:K${index + 12}`);
          cellRange.format.fill.color = "Red";
          cellRange.format.font.color = "white";
          await context.sync();
          console.log("Something is Wrong in the Row", index + 12);
        }
      });
    }
  });
};

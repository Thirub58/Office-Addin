import { z } from "zod";
const OdsAndSdSectionSchema = z.object({
  "Distribution": z.string(),
  "SUF": z.string(),
  "Shelf Price incl GST": z.union([z.number(), z.string()]),
  "RRP incl GST": z.string(),
  "LPTT": z.union([z.number(), z.string()]),
  "SLOG": z.union([z.string(), z.number()]),
  "PPD": z.union([z.number(), z.string()]),
  "Ullage": z.union([z.number(), z.string()]),
  "Cont. Terms": z.union([z.number(), z.string()]),
  "Deferred Deal": z.union([z.number(), z.string()]),
  "On-Invoice": z.string(),
  "%Margin @RRP": z.number(),
});


export const OdsAndSdSection =  ()=>{
Excel.run(async(context)=>{
const sheet=context.workbook.worksheets.getItem("Data 4")
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
    const headerRange=sheet.getRange("N10:Y10")
    headerRange.load("Values")
    await context.sync()
    const lastCellRange = sheet.getRange("Y:Y").getUsedRange();
    await context.sync();
    lastCellRange.load("address");
    await context.sync();
    const lastCellindex = lastCellRange.address.split(":")[1].slice(1);
    const valueRange = sheet.getRange(`N11:Y${lastCellindex}`);
    valueRange.load("values");
    await context.sync();
    const dataValues = valueRange.values;
    const headerValues = headerRange.values[0].slice(0, 11);
    console.log("The Header Values is",headerValues)
    const odsData=dataValues.map(row=>{
        const obj={}
        headerValues.forEach((header,index)=>{
            obj[header] = row[index];
        })
        return obj
    })
    odsData.forEach(async (data)=>{
        try {
            OdsAndSdSectionSchema.parse(data);
            const cellRange = sheet.getRange(`N${cellIndex + 12}:Y${cellIndex + 12}`);
            cellRange.format.fill.clear();
            cellRange.format.font.color = "black";
            await context.sync();
            console.log("Changing the Format Properties");
          } catch (error) {
            console.log("The Error in Some Cell is", error);
            const cellRange = sheet.getRange(`N${cellIndex + 12}:Y${cellIndex + 12}`);
            cellRange.format.fill.color = "Red";
            cellRange.format.font.color = "white";
            await context.sync();
            console.log("Something is Wrong in the Row", cellIndex + 12);
          }
    })
    

}
})
}
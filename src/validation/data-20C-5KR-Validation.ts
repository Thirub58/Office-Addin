import { sharing } from "webpack";
import { data_20Columns_Schema } from "../validationSchema/data-20C-5KR";
export const data_20C_5KR_Validation = () => {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const headerRange = sheet.getRange("A1:T1");
        headerRange.load("values");
        await context.sync();
        const headerValues = headerRange.values[0];
        console.log("The Header Values are", headerValues);
        const dataRange = sheet.getRange(`A2:T5000`);
        dataRange.load("values");
        await context.sync();

        //Generate the Array of Data Object
        const dataArray = dataRange.values.map((dataRow) => {
            const dataObject = {};
            headerValues.forEach((header, index) => {
                dataObject[header] = dataRow[index];
            });
            return dataObject;
        });
        console.log("The Data to be validated is", dataArray);

        //Validation of Each Row
        dataArray.forEach(async (dataObj, index) => {
            try {
                const isValidated = data_20Columns_Schema.parse(dataObj);
                const validRange = sheet.getRange(`A${index + 2}:T${index + 2}`)
                // validRange.load("format")
                // await context.sync()
                validRange.format.fill.clear()
                validRange.format.font.color = "black"
                await context.sync()
            } catch (error) {
                console.log("The Length of errors in row is",error.errors[0])
                console.log("The Error in the Range", `A${index + 2}:T${index + 2}`)
                const errorCell = error.errors[0].path[0]
                //console.log("The Header Values are", headerValues)
                // const errorCellIndex=headerValues.find(errorCell)
                //console.log("The Error Cell values are", errorCell)
                //console.log("The Error Cell index is", headerValues.indexOf(errorCell)+1)
                const columnIndex=String.fromCharCode(65+ headerValues.indexOf(errorCell))
                const errorColumn=`${columnIndex}${index+2}`
                console.log("The Error is at specific column",errorColumn)
                console.log("The Error Meesage", error.errors[0].message,)
                const invalidRange = sheet.getRange(`A${index + 2}:T${index + 2}`)
                invalidRange.format.fill.color = "Red"
                invalidRange.format.font.color = "white"
                invalidRange.select() // To Activate the invalid row
                await context.sync()

            }
        });
    });
};

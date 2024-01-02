import e from "express";
export const loadData=async()=>{
Excel.run(async(context)=>{
    console.log("The data is inside the")
    const sheet=context.workbook.worksheets.getActiveWorksheet()
    const response = await fetch("http://localhost:3010/data");
    const responseData = await response.json();
    console.log("Response data is",responseData,responseData.length)
      if (responseData.length > 0) {
        const dataValues=responseData.map((data)=>Object.values(data))
        console.log("The Value of the Object is",dataValues)
        const rowCount = responseData.length;
        const endColumn = String.fromCharCode(65 + dataValues[0].length - 1);
        console.log("The EndColumn is",(endColumn))
        console.log(`The Column Size is A2:${endColumn}${rowCount+1}`)
        const range = sheet.getRange(`A2:${endColumn}${rowCount+1}`);
        range.values = dataValues;
        await context.sync()
        console.log("The Response Data is uploaded")
      }
})
}
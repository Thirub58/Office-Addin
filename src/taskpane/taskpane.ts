/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { SectionValidation } from "../validation/sectionValidation";
import { DataValidation } from "../functions/dataValidation";
import { OdsAndSdSection } from "../validation/odsandsdValidation";
import { loadData } from "../functions/loaddata";
import { data_20C_5KR_Validation } from "../validation/data-20C-5KR-Validation";
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("loaddata").addEventListener('click',loadData)
    document.getElementById("validatedata").addEventListener('click',data_20C_5KR_Validation)
    try {
      await Excel.run(async (context) => {
        const workbook=context.workbook
        workbook.load('name')
        await context.sync()
        if(workbook.name==="ResponsiveCheck.xlsx"){
          const sheet=workbook.worksheets.getActiveWorksheet()
          // loadData()
        }
      });
    } catch (error) {
      console.error("Validation error:", error);
    }
  }
});

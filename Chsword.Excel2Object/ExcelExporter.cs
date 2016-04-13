﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Chsword.Excel2Object.Internal;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Chsword.Excel2Object
{
    public class ExcelExporter
    {
        public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data, ExcelType excelType)
        {

            IWorkbook workbook = null;
            switch (excelType)
            {
                case ExcelType.Xls:
                    workbook = new HSSFWorkbook();
                    break;
                case ExcelType.Xlsx:
                    workbook = new XSSFWorkbook();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("excelType");
            }
            var sheet = workbook.CreateSheet();
            var attrDict = ExcelUtil.GetExportAttrDict<TModel>();
            var attrArray = attrDict.OrderBy(c => c.Value.Order).ToArray();
            for (int i = 0; i < attrArray.Length; i++)
            {
                sheet.SetColumnWidth(i, 50 * 256);
            }
            var headerRow = sheet.CreateRow(0);

            for (int i = 0; i < attrArray.Length; i++)
            {
                headerRow.CreateCell(i).SetCellValue(attrArray[i].Value.Title);
            }
            int rowNumber = 1;
            foreach (var item in data.Where(c => c != null))
            {
                var row = sheet.CreateRow(rowNumber++);
                for (int i = 0; i < attrArray.Length; i++)
                {
                    row.CreateCell(i).SetCellValue((attrArray[i].Key.GetValue(item, null) ?? "").ToString());
                }
            }
            using (var output = new MemoryStream())
            {
                workbook.Write(output);
                var bytes = output.ToArray();
                return bytes;
            }
        }
        public byte[] ObjectToExcelBytes<TModel>(IEnumerable<TModel> data)
        {
            return this.ObjectToExcelBytes(data, ExcelType.Xls);
        }
    }
}

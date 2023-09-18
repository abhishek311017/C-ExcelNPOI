using System;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace ConsoleApp2
{
    class Program
    {
        static void Main()
        {
            // Sample DataTable
            DataTable dt = new DataTable();
            dt.Columns.Add("State", typeof(string));
            dt.Columns.Add("Value", typeof(int));
            dt.Rows.Add("CA", 1);
            dt.Rows.Add("CA", 2);
            dt.Rows.Add("TX", 3);
            dt.Rows.Add("TX", 4);
            dt.Rows.Add("TX", 5);
            dt.Rows.Add("TX", 6);
            dt.Rows.Add("FL", 8);
            dt.Rows.Add("FL", 9);
            dt.Rows.Add("FL", 10);
            dt.Rows.Add("FL", 11);
            dt.Rows.Add("F", 1);
            dt.Rows.Add("F", 1);
            dt.Rows.Add("F", 1);
            dt.Rows.Add("F", 1);

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell headerCell = headerRow.CreateCell(i);
                headerCell.SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            int startRow = 1;
            int endRow = 1;

            for (int i = 1; i <= dt.Rows.Count; i++)
            {
                if (i == dt.Rows.Count || dt.Rows[i]["State"].ToString() != dt.Rows[i - 1]["State"].ToString())
                {
                    endRow = i;
                    if (startRow != endRow)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));
                    }
                    startRow = i + 1;
                }
            }

            using (FileStream fs = new FileStream("output.xlsx", FileMode.Create))
            {
                workbook.Write(fs);
            }
        }
    }

}

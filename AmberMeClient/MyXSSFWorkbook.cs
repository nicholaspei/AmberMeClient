using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace AmberMeClient
{
    public class MyXSSFWorkbook
    {
 
        public IList<Task> GetTaskList(Stream stream)
        {
            var list = new List<Task>();
            var workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheet("个人周报");
            var employeeName = sheet.GetRow(2).Cells[15].ToString();
            int maxrow = sheet.LastRowNum;
            for (int i = 4; i < maxrow-1; i++)
			{
                var currentRow = sheet.GetRow(i);
                if (currentRow != null)
                {
                    var cells = currentRow.Cells;
                    var task = new Task();
                    task.EmpName = employeeName;
                    foreach (var cell in cells)
                    {
                        switch (cell.ColumnIndex) { 
                            case 1:
                                task.TaskNum = this.CellValueTyper(cell).ToString();
                                break;
                            case 2:
                                task.TaskName = this.CellValueTyper(cell).ToString();
                                break;
                            case 3:
                                task.ProjectNum = this.CellValueTyper(cell).ToString();
                                break;
                            case 4:
                                task.TaskType = this.CellValueTyper(cell).ToString();
                                break;
                            case 5:
                                task.InPlan = this.CellValueTyper(cell).ToString();
                                break;
                            case 6:
                                task.Description = this.CellValueTyper(cell).ToString();
                                break;
                            case 7:
                                if (cell.CellType == CellType.Blank)
                                    task.Mon = 0;
                                else
                                task.Mon =(double)this.CellValueTyper(cell);
                                break;
                            case 8:
                                if (cell.CellType == CellType.Blank)
                                    task.Tue = 0;
                                else
                                task.Tue = (double)this.CellValueTyper(cell);
                                break;
                            case 9:
                                if (cell.CellType == CellType.Blank)
                                    task.Wen = 0;
                                else
                                task.Wen = (double)this.CellValueTyper(cell);
                                break;
                            case 10:
                                if (cell.CellType == CellType.Blank)
                                    task.Thr = 0;
                                else
                                task.Thr = (double)this.CellValueTyper(cell);
                                break;
                            case 11:
                                if (cell.CellType == CellType.Blank)
                                    task.Fir = 0;
                                else
                                task.Fir = (double)this.CellValueTyper(cell);
                                break;
                            case 12:
                                if (cell.CellType == CellType.Blank)
                                    task.San = 0;
                                else
                                task.San = (double)this.CellValueTyper(cell);
                                break;
                            case 13:
                                if (cell.CellType == CellType.Blank)
                                    task.Sun = 0;
                                else
                                task.Sun = (double)this.CellValueTyper(cell);
                                break;
                            case 14:
                                //if (cell.CellType == CellType.Blank)
                                //    task.Sum = 0;
                                //else
                                //task.Sum = (double)this.CellValueTyper(cell);
                                task.Sum = 0;
                                break;
                            case 15:
                                task.Percent = this.CellValueTyper(cell).ToString();
                                break;
                            case 16:
                                task.Result = this.CellValueTyper(cell).ToString();
                                break;
                            case 17:
                                task.WillFinDate = this.CellValueTyper(cell).ToString();
                                break;
                            case 18:
                                task.WillFinMD = this.CellValueTyper(cell).ToString();
                                break;
                            case 19:
                                task.Advince = this.CellValueTyper(cell).ToString();
                                break;
                            default:                               
                                break;
                        }
                    }
                    list.Add(task);
                }
            }
            return list;
        }

        private object CellValueTyper(ICell cell)
        {
            if (cell.CellType == CellType.Blank)
                return string.Empty;
            if (cell.CellType == CellType.Boolean)
                return cell.BooleanCellValue;
            if (cell.CellType == CellType.Numeric)
                return cell.NumericCellValue;
            if (cell.CellType == CellType.String)
                return cell.StringCellValue;
            return cell.ToString();
        }

        public void WriteToFile(IList<Task> list, Stream stream) {
            var workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheet("个人周报汇总表");
            for (int i = 4; i < list.Count+4; i++)
            {
               var row= sheet.GetRow(i);
               for (int j = 1; j < 21; j++)
               {
                   switch (j) { 
                       case 1:
                           row.GetCell(j).SetCellValue(list[i-4].EmpName);
                           break;
                   }
               }
            }

            FileStream fs = File.Create("周报汇总.xlsx");
            workbook.Write(fs);
            fs.Close();
        }
    }
}

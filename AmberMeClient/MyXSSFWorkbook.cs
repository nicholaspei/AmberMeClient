using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

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
            if(employeeName==string.Empty)
                employeeName = sheet.GetRow(2).Cells[16].ToString();
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
                                if (cell.CellType == CellType.Blank)
                                    task.Percent = 0.0;
                                else
                                task.Percent = (double)this.CellValueTyper(cell);
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
            var workbook = new HSSFWorkbook(stream);
            var sheet = workbook.GetSheet("个人周报汇总表");
            for (int i = 4; i < list.Count+4; i++)
            {
               var row= sheet.GetRow(i);
               for (int j = 0; j < 20; j++)
               {
                   switch (j) { 
                       case 0:
                           row.GetCell(j).SetCellValue(list[i-4].EmpName);
                           break;
                       case 1:
                           row.GetCell(j).SetCellValue(list[i - 4].TaskNum);
                           break;
                       case 2:
                           row.GetCell(j).SetCellValue(list[i - 4].TaskName);
                           break;
                       case 3:
                           row.GetCell(j).SetCellValue(list[i - 4].ProjectNum);
                           break;
                       case 4:
                           row.GetCell(j).SetCellValue(list[i - 4].TaskType);
                           break;
                       case 5:
                           row.GetCell(j).SetCellValue(list[i - 4].InPlan);
                           break;
                       case 6:
                           row.GetCell(j).SetCellValue(list[i - 4].Description);
                           break;
                       case 7:
                           if(list[i-4].Mon==0.0)
                           row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i-4].Mon);
                           break;                     
                       case 8:
                           if (list[i - 4].Tue == 0.0)
                               row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].Tue);
                           break;
                       case 9:
                           if (list[i - 4].Wen == 0.0)
                           row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].Wen);
                           break;
                       case 10:
                           if (list[i - 4].Thr == 0.0)
                           row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].Thr);
                           break;
                       case 11:
                           if (list[i - 4].Fir == 0.0)
                           row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].Fir);
                           break;
                       case 12:
                           if (list[i - 4].San == 0.0)
                               row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].San);
                           break;
                       case 13:
                           if (list[i - 4].Sun == 0.0)
                               row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue(list[i - 4].Sun);
                           break;
                       case 14:
                           row.GetCell(j).SetCellFormula(string.Format("SUM(H{0}:N{0})", i+1));                          
                           break;
                       case 15:                       
                           //var cellStyle = workbook.CreateCellStyle();
                           //cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.0%");
                           //row.GetCell(j).CellStyle = cellStyle;
                           if (list[i - 4].Percent == 0.0)
                               row.GetCell(j).SetCellValue("");
                           else
                           row.GetCell(j).SetCellValue((list[i - 4].Percent*100.0).ToString()+"%");
                           break;
                       case 16:
                           row.GetCell(j).SetCellValue(list[i - 4].Result);
                           break;
                       case 17:
                           row.GetCell(j).SetCellValue(list[i - 4].WillFinDate);
                           break;
                       case 18:
                           row.GetCell(j).SetCellValue(list[i - 4].WillFinMD);
                           break;
                       case 19:
                           row.GetCell(j).SetCellValue(list[i - 4].Advince);
                           break;
                       default:
                           break;
                   }
               }
            }

            FileStream fs = File.Create("周报汇总.xls");
            workbook.Write(fs);
            fs.Close();
        }

        public IList<Task> GetAllTask(Stream stream)
        {
            var list = new List<Task>();
            var workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheet("个人周报汇总表");
            int maxrow = sheet.LastRowNum;
            for (int i = 4; i < maxrow - 1; i++) {
                var currentRow = sheet.GetRow(i);
                var cells = currentRow.Cells;
                var task = new Task();
                if (currentRow != null)
                {
                    foreach (var cell in cells)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 0:
                                task.EmpName = this.CellValueTyper(cell).ToString();
                                break;
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
                                if (cell.CellType == CellType.Blank||cell.CellType==CellType.String)
                                    task.Mon = 0;
                                else
                                    task.Mon = (double)this.CellValueTyper(cell);
                                break;
                            case 8:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.Tue = 0;
                                else
                                    task.Tue = (double)this.CellValueTyper(cell);
                                break;
                            case 9:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.Wen = 0;
                                else
                                    task.Wen = (double)this.CellValueTyper(cell);
                                break;
                            case 10:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.Thr = 0;
                                else
                                    task.Thr = (double)this.CellValueTyper(cell);
                                break;
                            case 11:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.Fir = 0;
                                else
                                    task.Fir = (double)this.CellValueTyper(cell);
                                break;
                            case 12:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.San = 0;
                                else
                                    task.San = (double)this.CellValueTyper(cell);
                                break;
                            case 13:
                                if (cell.CellType == CellType.Blank || cell.CellType == CellType.String)
                                    task.Sun = 0;
                                else
                                    task.Sun = (double)this.CellValueTyper(cell);
                                break;
                            case 14:
                                if (cell.CellType == CellType.Blank)
                                    task.Sum = 0;
                                else
                                    task.Sum = cell.NumericCellValue;
                              //  task.Sum = 0;
                                break;
                            case 15:
                                if (cell.CellType == CellType.Blank)
                                    task.Percent = 0.0;
                                else
                                    task.PercentStr = cell.ToString();
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

        public void WriteAllTaskToFile(IList<Task> list,IList<NextWeekTask> nwlist, Stream stream)
        {
            var workbook = new HSSFWorkbook(stream);
            var sheet = workbook.GetSheet("项目周报");
            for (int i = 5; i < list.Count + 5; i++)
            {
                var row = sheet.GetRow(i);                
                if(row.Cells[0].ToString()=="Amber")
                {

                }
                else
                {                   
                    sheet.CopyRow(i - 1, i);
                    row = sheet.GetRow(i);
                }

                for (int j = 0; j < 20; j++)
                {
                    switch (j)
                    {
                        case 1:
                            row.GetCell(j).SetCellValue(list[i - 5].TaskNum);
                            break;
                        case 6:
                            row.GetCell(j).SetCellValue(list[i - 5].Description);
                            break;
                        case 14:
                            row.GetCell(j).SetCellValue(list[i - 5].Sum);
                            break;
                        case 16:
                            row.GetCell(j).SetCellValue(list[i - 5].EmpName);
                            break;
                        case 17:
                            row.GetCell(j).SetCellValue(list[i - 5].Result);
                            break;
                        default:
                            break;
                    }
                }
            }
          
            //开始写入下周计划
            for (int i = 0; i < nwlist.Count;i++)
            {
                var nwstartrow = sheet.LastRowNum;
                var row = sheet.GetRow(nwstartrow+1);
                if(row==null)
                {
                    sheet.CopyRow(nwstartrow,nwstartrow+1);
                    row=sheet.GetRow(nwstartrow+1);
                }
                for (int j = 0; j < 20; j++)
                {
                    switch (j)
                    {    
                        case 0:
                            row.Cells[0].SetCellValue(nwlist[i].TaskNum);
                            break;
                        case 1:
                            row.Cells[1].SetCellValue(nwlist[i].TaskName);
                            break;
                        case 2:
                            row.Cells[2].SetCellValue(nwlist[i].ProjectNum);
                            break;
                        case 3:
                            row.Cells[3].SetCellValue(nwlist[i].TaskType);
                            break;
                        case 4:
                            row.Cells[4].SetCellValue(nwlist[i].InPlan);
                            break;
                        case 5:
                            row.Cells[5].SetCellValue(nwlist[i].Description);
                            break;
                        case 6:
                            row.Cells[6].SetCellValue(nwlist[i].Percent);
                            break;
                        case 7:
                            row.Cells[7].SetCellValue(nwlist[i].StartTime);
                            break;
                        case 8:
                            row.Cells[8].SetCellValue(nwlist[i].EndTime);
                            break;
                        case 9:
                            row.Cells[9].SetCellValue(nwlist[i].ManDay);
                            break;
                        case 10:
                            row.Cells[10].SetCellValue(nwlist[i].Employee);
                            break;
                        case 11:
                            row.Cells[11].SetCellValue(nwlist[i].PlanEmployee);
                            break;
                        case 12:
                            row.Cells[12].SetCellValue(nwlist[i].Result);
                            break;
                        case 13:
                            row.Cells[13].SetCellValue(nwlist[i].Commit);
                            break;
                        default:
                            break;
                    }
                }
            }

            FileStream fs = File.Create("成品目标.xls");
            workbook.Write(fs);
            fs.Close();
        }

        public IList<NextWeekTask> getNextWeekTask(Stream nwstream)
        {
            var list = new List<NextWeekTask>();
            var workbook = new XSSFWorkbook(nwstream);
            var sheet = workbook.GetSheet("项目周报");
            int maxrow = sheet.LastRowNum-1;
            for (int i = 5; i < maxrow - 1; i++)
            {
                var currentRow = sheet.GetRow(i);
                var cells = currentRow.Cells;
                var task = new NextWeekTask();
                if (currentRow != null)
                {
                    foreach (var cell in cells)
                    {
                        switch (cell.ColumnIndex)
                        {
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
                                task.Percent = this.CellValueTyper(cell).ToString();
                                break;
                            case 8:
                                task.StartTime = this.CellValueTyper(cell).ToString();
                                break;
                            case 9:
                                task.EndTime = this.CellValueTyper(cell).ToString();
                                break;                           
                            case 10:
                                try
                                {
                                task.ManDay =int.Parse(this.CellValueTyper(cell).ToString());
                                }
                                catch{
                                 task.ManDay=0;
                                }
                                break;
                            case 11:
                                task.Employee = this.CellValueTyper(cell).ToString();
                                break;
                            case 12:
                                task.PlanEmployee = this.CellValueTyper(cell).ToString();
                                break;
                            case 13:
                                task.Result = this.CellValueTyper(cell).ToString();
                                break;
                            case 14:
                                task.Commit = this.CellValueTyper(cell).ToString();
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
    }
}

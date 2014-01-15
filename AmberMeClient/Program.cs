using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace AmberMeClient
{
    class Program
    {
        static void Main(string[] args)
        {

            var streamManager = new StreamManager();
            var bookManager = new MyXSSFWorkbook();
            var stream = streamManager.GetStreamByName("个人周报汇总-TS中行总行BJ13003FW.xlsx");
            var nwstream = streamManager.GetStreamByName("项目下周计划周报-运维数据库.xlsx");
            var list = bookManager.GetAllTask(stream);
            var nwlist = bookManager.getNextWeekTask(nwstream);
            var mergeList = new List<Task>();
            var taskNumList = list.OrderBy(s=>s.TaskNum).Select(s =>new { s.TaskNum,s.TaskName}).Distinct();

            foreach (var taskdis in taskNumList)
            {
                var sameNumTasks = list.Where(l => l.TaskNum == taskdis.TaskNum);

                var taskContent = string.Empty;
                double workTime = 0.0;
                var workers = string.Empty;
                var allresult=string.Empty;
                //合并
                foreach (var task in sameNumTasks)
                { 
                  //工作内容合并
                    taskContent += task.Description ;
                    taskContent += "\n";

                 //工时合并
                    workTime = workTime + task.Mon + task.Tue + task.Wen + task.Thr + task.Fir + task.San + task.Sun;

                //人员合并
                    workers += task.EmpName;
                    workers += "、";

                //提交结果合并
                    allresult+=task.Result;
                    allresult+="\n";
                }
                var mergetask = new Task();
                mergetask.TaskNum = taskdis.TaskNum;
                mergetask.TaskName = taskdis.TaskName;
                mergetask.Result = allresult;
                mergetask.Sum = workTime;
                mergetask.Description = taskContent;
                mergetask.EmpName = workers;
                mergetask.Advince = string.Empty;               
                mergeList.Add(mergetask);
            }

            var writeFile = streamManager.GetWriteStream("成品目标template.xls");
            bookManager.WriteAllTaskToFile(mergeList, nwlist, writeFile);
            Console.Read();
        }

        static void WeeklyReport()
        {
            Console.WriteLine("开始寻找当前文件夹下的所有.xlsx文件，请耐心等待");
            var list = new List<Task>();
            var fileManager = new FileManager();
            var streamManager = new StreamManager();
            var bookManager = new MyXSSFWorkbook();
            var files = fileManager.GetAllFileNames();
            Console.WriteLine("开始导出数据.......");

            foreach (var file in files)
            {
                try
                {
                    var stream = streamManager.GetStreamByName(file);
                    var currentlist = bookManager.GetTaskList(stream);
                    Console.WriteLine("导出{0}的周报成功 :)", file);
                    list.AddRange(currentlist);
                }
                catch
                {
                    Console.WriteLine("导出{0}的周报出错啦。请查看Excel是否格式正确", file);
                }
            }

            // 需要写入数据的文件
            Console.WriteLine("开始写入数据......");
            var writeFile = streamManager.GetWriteStream("项目周报-项目编号-汇报日期测试.xls");
            bookManager.WriteToFile(list, writeFile);
            Console.WriteLine("，导出成功，导出的文件名为：《周报汇总.xls》，请输入任何字符结束程序");
        }
    }
}

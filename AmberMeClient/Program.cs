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
                        Console.WriteLine("导出{0}的周报出错啦。请查看Excel是否格式正确",file);
                    }
                }
            
            // 需要写入数据的文件
            Console.WriteLine("开始写入数据......");
            var writeFile = streamManager.GetWriteStream("项目周报-项目编号-汇报日期测试.xls");
            bookManager.WriteToFile(list, writeFile);
            Console.WriteLine("，导出成功，导出的文件名为：《周报汇总.xls》，请输入任何字符结束程序");
            Console.Read();
        }

      
    }
}

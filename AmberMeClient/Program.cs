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
            var fileManager = new FileManager();
            var streamManager=new StreamManager();
            var bookManager = new MyXSSFWorkbook();
            var files = fileManager.GetAllFileNames();
            var stream = streamManager.GetStreamByName(files[0]);
            var list = bookManager.GetTaskList(stream);

           // 需要写入数据的文件
            var writeFile = streamManager.GetWriteStream("项目周报-项目编号-汇报日期测试.xls");
            bookManager.WriteToFile(list, writeFile);
           
            Console.Read();
        }

      
    }
}

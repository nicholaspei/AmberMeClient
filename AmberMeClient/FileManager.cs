using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace AmberMeClient
{
    /// <summary>
    /// 获取当前文件夹下所有的.xlsx文件
    /// </summary>
    public class FileManager
    {
        public IList<string> GetAllFileNames() {
            var directoryPath = Directory.GetCurrentDirectory();
            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);
            FileInfo[] files = dirInfo.GetFiles();
            var list = new List<string>();
            foreach (var file in files)
            {
                if (file.Extension == ".xlsx" && file.Name != "项目周报-项目编号-汇报日期测试.xlsx")
                {
                    list.Add(file.Name);
                }
            }

            return list;
        }
    }
}

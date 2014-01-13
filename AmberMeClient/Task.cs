using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AmberMeClient
{
    public class Task
    {
        public string EmpName { get; set; } //员工编号

        public string TaskNum { get; set; }

        public string TaskName { get; set; }

        public string ProjectNum { get; set; }

        public string TaskType { get; set; }

        public string InPlan { get; set; }

        public string  Description { get; set; }

        public double Mon { get; set; }

        public double Tue { get; set; }

        public double Wen { get; set; }

        public double Thr { get; set; }

        public double Fir { get; set; }

        public double San { get; set; }

        public double Sun { get; set; }

        public double Sum { get; set; }

        public double Percent { get; set; }

        public string PercentStr { get; set; }

        public string Result { get; set; }

        public string WillFinDate { get; set; }

        public string WillFinMD { get; set; }

        public string Advince { get; set; }
    }
}

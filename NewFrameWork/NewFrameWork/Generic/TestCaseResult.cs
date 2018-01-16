using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewFrameWork.Generic
{
    internal class TestCaseResult
    {
        public string CompletionStatus;
        public string DiffImage;
        public string Exception;
        public string GoldImage;
        public string TestImage;
        public string Errorimage;
        public string Actualresult;
        public string Executiontime;
        public string Testdata;

        public TestCaseResult()
        {
            CompletionStatus = string.Empty;
            DiffImage = string.Empty;
            Exception = string.Empty;
            GoldImage = string.Empty;
            TestImage = string.Empty;
            Errorimage = string.Empty;
            Actualresult = string.Empty;
            Executiontime = "00:00:00";
            Testdata = string.Empty;
        }
    }
}

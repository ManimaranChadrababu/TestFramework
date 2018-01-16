using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using NewFrameWork.Test;
using NewFrameWork.Selenium;
using System.Diagnostics;
using System.Reflection;

namespace NewFrameWork.Generic
{
   public class Runner
    {
       
        internal static void RunCase(ConfigValues conval)
        {
            HtmlReport htmlRep = null;

            conval.Browser = ConfigurationManager.AppSettings["browser"];
            var testcase = new TestVP1();

            Driver.Instance.Quits();
            htmlRep = new HtmlReport("TestVP1", false);
            String[] header =
               {
                    "Test case ID","Step_Seq","Test Case Title", "Steps", "Expected Result", "Completion Status","Error Image", "Exceptions(if any)", "Test Data","Execution time"
                };

            htmlRep.AddHeader(header);


        }
        public static void invokTest(string method,TestVP1 testcase ,ConfigValues configval)
        {
            MethodInfo testCaseMethod = testcase.GetType().GetMethod(method);
            List<TestCaseResult> Testresult = new List<TestCaseResult>();
            Stopwatch watch = new Stopwatch();
            watch.Start();
            Driver.NewInstance();
            Testresult = (List<TestCaseResult>)testCaseMethod.Invoke(testcase, new object[] { configval.datasetindex, testcasedata });
            watch.Stop();
            Driver.Instance.Quits();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewFrameWork.Generic
{
    public struct ConfigValues
    {  
        public string Browser { get; set; }
        public string BrowserType { get; set; }
        public string FFDriverPath { get; set; }
        public string ChromeDriverPath { get; set; }
        public string IEDriverPath { get; set; }
        public string AppURL { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class LaborproModel
    {
        public int Id { get; set; }
        public string FeatureFileName { get; set; }
        public string ScenarioName { get; set; }
        public string SmokeTest { get; set; }
        public string RegressionTest { get; set; }
    }
}

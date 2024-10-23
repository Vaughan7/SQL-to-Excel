using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    internal class EmailValueSet
    {
        [JsonProperty("Queries")]
        public Dictionary<string, EmailValue> emailValues { get; set; }
    }
}

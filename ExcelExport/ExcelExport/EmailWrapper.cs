using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    internal class EmailWrapper
    {
        [JsonProperty("JsonValues")]
        public EmailValueSet emailValueSet { get; set; }
    }
}

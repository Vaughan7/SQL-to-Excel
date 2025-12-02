using DocumentFormat.OpenXml.Office2013.Word;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    internal class EmailValue
    {
        [JsonProperty("Address")]
        public string[] address { get; set; }

        [JsonProperty("CC")]
        public string[] cc { get; set; }

        [JsonProperty("Subject")]
        public string subject { get; set; }

        [JsonProperty("Body")] 
        public string body { get; set; }

        [JsonProperty("Pivot Table")]
        public bool pivotTable { get; set; }

        [JsonProperty("Row Labels")]
        public int[] rowLabels { get; set; }

        [JsonProperty("Column Labels")]
        public int[] columnLabels { get; set;}

        [JsonProperty("Value Labels")]
        public int[] valueLabels { get; set; }

        [JsonProperty("Value Functions")]
        public string[] valueFunctions { get; set; }

        [JsonProperty("Pivot Table 2")]
        public bool pivotTable2 { get; set; }

        [JsonProperty("Row Labels 2")]
        public int[] rowLabels2 { get; set; }

        [JsonProperty("Column Labels 2")]
        public int[] columnLabels2 { get; set; }

        [JsonProperty("Value Labels 2")]
        public int[] valueLabels2 { get; set; }

        [JsonProperty("Value Functions 2")]
        public string[] valueFunctions2 { get; set; }

        [JsonProperty("Pivot Table 3")]
        public bool pivotTable3 { get; set; }

        [JsonProperty("Row Labels 3")]
        public int[] rowLabels3 { get; set; }

        [JsonProperty("Column Labels 3")]
        public int[] columnLabels3 { get; set; }

        [JsonProperty("Value Labels 3")]
        public int[] valueLabels3 { get; set; }

        [JsonProperty("Value Functions 3")]
        public string[] valueFunctions3 { get; set; }

        [JsonProperty("Duplicate Columns")]
        public int[] duplicateColumns { get; set; }

        [JsonProperty("Collapsed Field")]
        public bool collapseField {  get; set; }

        [JsonProperty("Move Σ Value")]
        public bool moveΣValue { get; set; }

        [JsonProperty("Move Σ Value 2")]
        public bool moveΣValue2 { get; set; }

        [JsonProperty("Filter")]
        public int[] filter { get; set; }


    }
}

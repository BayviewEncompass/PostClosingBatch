using System.Collections.Generic;
using Newtonsoft.Json;

namespace PostClosingBatch
{
    public class BatchUpdatePC
    {
        [JsonProperty("RepName")]
        public string RepName { get; set; }

        [JsonProperty("Col1")]
        public string Col1 { get; set; }

        [JsonProperty("Col2")]
        public string Col2 { get; set; }

        [JsonProperty("Col3")]
        public string Col3 { get; set; }

        [JsonProperty("Col4")]
        public string Col4 { get; set; }

        [JsonProperty("Col5")]
        public string Col5 { get; set; }
    }
    public class BatchUpdateDB
    {
        [JsonProperty("BatchUpdatePC")]
        public List<BatchUpdatePC> BatchUpdatePC { get; set; }

    }
}

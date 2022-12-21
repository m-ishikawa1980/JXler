using System.Collections.Generic;

namespace JXler.Models
{
    public class Settings
    {
        public bool OutputXlsLink { get; set; }
        public bool WriteBlankColumn { get; set; }
        public string IndexSheet { get; set; }
        public string BasePath { get; set; }
        public XlsCellStyles XlsCellStyles { get; set; }
        public List<JsonXls> JsonXlsHash { get; set; }
        public ExecPtn ExecPtn { get; set; }
        public string Path { get; set; }

        public Settings()
        {
        }
    }

    public class XlsCellStyles
    {
        public string Text { get; set; }
        public string NumberInt { get; set; }
        public string NumberFloat { get; set; }
        public string DateTime { get; set; }

    }

    public class JsonXls
    {
        public int No { get; set; }
        public string JsonPath { get; set; }
        public string JsonName { get; set; }
        public string Action { get; set; }
        public string XlsPath { get; set; }
        public string XlsName { get; set; }
    }

}
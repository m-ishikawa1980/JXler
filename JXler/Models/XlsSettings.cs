namespace JXler.Models
{
    public class XlsSettings
    {
        public bool OutputXlsLink;
        public bool WriteBlankColumn;
        public string IndexSheet;
        public string SheetName;
        public string ChildSheet;
        public int IndexNum;
        public XlsCellStyles XlsCellStyles;

        public XlsSettings()
        {
        }

        public XlsSettings(Settings settings) {
            IndexSheet = settings.IndexSheet;
            SheetName = settings.IndexSheet;
            WriteBlankColumn = settings.WriteBlankColumn;
            OutputXlsLink = settings.OutputXlsLink;
            XlsCellStyles = settings.XlsCellStyles;
        }
    }
}

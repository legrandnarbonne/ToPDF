
namespace ToPDF.Classes
{
    public class WorkingConfig
    {
        public string SourceFolder { get; set; }
        public string DestinationFolder { get; set; }

        public bool OverWriteIfExist { get; set; }

        public bool DeleteAfterConvert { get; set; }
    }
}

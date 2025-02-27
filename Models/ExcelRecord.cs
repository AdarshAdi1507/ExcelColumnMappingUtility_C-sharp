namespace ExcelProcessor.Models
{
    public class ExcelRecord
    {
        public string DrawingNumber { get; set; }
        public string Revision { get; set; }
        public string Title { get; set; }
        public string ProjectNo { get; set; }
        public string RevisionText { get; set; }
        public string RailPart { get; set; }
        public string Part { get; set; }
        public string RemotePath { get; set; }
        public string FileName1 { get; set; }
        public string FileName2 { get; set; }
        public string FileName3 { get; set; }
        public string FileName4 { get; set; }
        public string FileName5 { get; set; }
        public string FileName6 { get; set; }
        public string Supersedes { get; set; }
        public string MigratedFrom { get; set; }

        public ExcelRecord()
        {
            DrawingNumber = string.Empty;
            Revision = string.Empty;
            Title = string.Empty;
            ProjectNo = string.Empty;
            RevisionText = string.Empty;
            RailPart = string.Empty;
            Part = string.Empty;
            RemotePath = string.Empty;
            FileName1 = string.Empty;
            FileName2 = string.Empty;
            FileName3 = string.Empty;
            FileName4 = string.Empty;
            FileName5 = string.Empty;
            FileName6 = string.Empty;
            Supersedes = string.Empty;
            MigratedFrom = string.Empty;
        }
    }
}
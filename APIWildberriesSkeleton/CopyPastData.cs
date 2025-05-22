namespace APIWildberries
{
    public struct CopyPastData(string spreadSheetCopyId, string listNameCopy,
        string spreadSheetPastId, string listNamePast,
        int startRowIndexCopy, int endRowIndexCopy, string startColumnCopy, string endColumnCopy,
        int startRowIndexPast, int endRowIndexPast, string startColumnPast, string endColumnPast)
    {
        public string SpreadSheetCopyId { get; set; } = spreadSheetCopyId;
        public string SpreadSheetPastId { get; set; } = spreadSheetPastId;
        public string ListNameCopy { get; set; } = listNameCopy;
        public string ListNamePast { get; set; } = listNamePast;
        public int StartRowIndexCopy { get; set; } = startRowIndexCopy;
        public int EndRowIndexCopy { get; set; } = endRowIndexCopy;
        public string StartColumnCopy { get; set; } = startColumnCopy;
        public string EndColumnCopy { get; set; } = endColumnCopy;
        public int StartRowIndexPast { get; set; } = startRowIndexPast;
        public int EndRowIndexPast { get; set; } = endRowIndexPast;
        public string StartColumnPast { get; set; } = startColumnPast;
        public string EndColumnPast { get; set; } = endColumnPast;
    }
}
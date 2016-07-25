namespace MintScrape.Interfaces {
    public interface IExcelMapper {
        bool SetLastTransactionToZero { get; set; }
        void WriteToWorkSheet();
    }
}
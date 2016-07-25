namespace MintScrape.Interfaces {
    public interface IAccount {
        string AccountName { get; set; }
        double Balance { get; set; }
        string NickName { get; set; }
        string LastUpdated { get; set; }
        int ExcelSheetRow { get; set; }
    }
}
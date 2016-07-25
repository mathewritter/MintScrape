using MintScrape.Interfaces;

namespace MintScrape.Core {
    public class Account : IAccount {
        public string AccountName { get; set; }
        public double Balance { get; set; }
        public int ExcelSheetRow { get; set; }
        public string LastUpdated { get; set; }
        public string NickName { get; set; }
    }
}
using System.Collections.Generic;

namespace MintScrape.Interfaces {
    public interface IHtmlParserUtility {
        IEnumerable<IAccount> GetAccountData();
        void WriteAccountDataToConsole();
    }
}
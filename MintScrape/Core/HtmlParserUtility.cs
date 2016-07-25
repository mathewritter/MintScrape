using System;
using System.Collections.Generic;
using MintScrape.Infrastructure;
using MintScrape.Interfaces;
using Ninject;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;

namespace MintScrape.Core {
    internal class HtmlParserUtility : IHtmlParserUtility {
        private readonly RemoteWebDriver _driver;

        /// <summary>
        ///     Constructor.
        /// </summary>
        /// <param name="driver"></param>
        public HtmlParserUtility(RemoteWebDriver driver) {
            _driver = driver;
        }

        /// <summary>
        ///     Get account data from Mint.
        /// </summary>
        /// <returns>Collection of Accounts.</returns>
        public IEnumerable<IAccount> GetAccountData() {
            var iocModule = new IocModule();
            _driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5));
            var accountElements = _driver.FindElementsByClassName("accounts-data-li");
            var accounts = new List<IAccount>();
            foreach (var accountElement in accountElements) {
                var iocKernel = new StandardKernel(new IocModule());
                var account = iocKernel.Get<IAccount>();
                var balance = accountElement.FindElement(By.ClassName("balance")).Text;
                account.Balance = !string.IsNullOrEmpty(balance) && balance.Length > 1
                    ? Math.Abs(Convert.ToDouble(balance.Substring(1)))
                    : 0;
                account.AccountName = accountElement.FindElement(By.ClassName("accountName")).Text;
                account.NickName = accountElement.FindElement(By.ClassName("nickname")).Text;
                account.LastUpdated = accountElement.FindElement(By.ClassName("last-updated")).Text;
                if (!string.IsNullOrEmpty(account.NickName)) {
                    accounts.Add(account);
                }
            }
            return accounts;
        }

        /// <summary>
        ///     Outputs the account data.
        /// </summary>
        public void WriteAccountDataToConsole() {
            Console.WriteLine();
            var accounts = GetAccountData();
            foreach (var account in accounts) {
                Console.WriteLine(account.AccountName);
                Console.WriteLine(account.NickName);
                Console.WriteLine(account.Balance);
                Console.WriteLine(account.LastUpdated);
                Console.WriteLine(Environment.NewLine);
            }
        }
    }
}
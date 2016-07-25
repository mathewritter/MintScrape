using System;
using System.Configuration;
using MintScrape.Infrastructure;
using MintScrape.Interfaces;
using MintScrape.Window;
using Ninject;
using Ninject.Parameters;
using OpenQA.Selenium.Remote;

namespace MintScrape {
    public class Program {
        /// <summary>
        ///     Prompt for command.
        /// </summary>
        private static void Prompt() {
            Console.WriteLine("**********************************************************");
            Console.WriteLine("Choose Option:");
            Console.WriteLine("  [R]ead Account Data");
            Console.WriteLine("  [W]rite to Spreadsheet");
            Console.WriteLine("  [O]verwrite Transaction Column and Write to Spreadsheet");
            Console.WriteLine("  [E]xit");
            Console.WriteLine("**********************************************************");
            Position.ToTop();
        }

        /// <summary>
        ///     Print "MintScrape" in ASCII characters.
        /// </summary>
        private static void PrintAppName() {
            Console.WriteLine("888b     d888 d8b          888    .d8888b.");
            Console.WriteLine("8888b   d8888 Y8P          888   d88P  Y88b");
            Console.WriteLine("88888b.d88888              888   Y88b.");
            Console.WriteLine("888Y88888P888 888 88888b.  888888  Y888b.    .d8888b 888d888 8888b. 88888b.   .d88b.");
            Console.WriteLine("888 Y888P 888 888 888  88b 888       Y88b. d88P.    888PP       88b 888 888b d8P  Y8b");
            Console.WriteLine("888  Y8P  888 888 888  888 888         .888 888     888.   .d888888 888 8888 88888888");
            Console.WriteLine("888   v   888 888 888  888 Y88b. Y88b d88P Y88b.    888    888  888 888 d88P Y8b.");
            Console.WriteLine("888       888 888 888  888  Y888 Y8888PP    Y8888P  888    Y8888888 88888P    Y88888");
            Console.WriteLine("                                                                    888");
            Console.WriteLine("                                                                    888");
            Console.WriteLine("                                                                    888");
        }

        /// <summary>
        ///     Main method.
        /// </summary>
        /// <param name="args"></param>
        private static void Main(string[] args) {
            Size.Resize();
            PrintAppName();

            var iocKernel = new StandardKernel(new IocModule());
            using (var driver = iocKernel.Get<RemoteWebDriver>()) {
                //Login to mint.com
                var userAccount = iocKernel.Get<IUserAccount>(new ConstructorArgument("driver", driver));
                userAccount.Login();

                // Display account data in console
                var htmlParserUtility =
                    iocKernel.Get<IHtmlParserUtility>(new ConstructorArgument("driver", driver));
                htmlParserUtility.WriteAccountDataToConsole();

                // Prompt for command
                Prompt();
                var enteredValue = Console.ReadLine() ?? string.Empty;

                // While refreshing or writing
                while (enteredValue.Equals("R", StringComparison.CurrentCultureIgnoreCase) ||
                       enteredValue.Equals("W", StringComparison.CurrentCultureIgnoreCase) ||
                       enteredValue.Equals("O", StringComparison.CurrentCultureIgnoreCase)) {
                    // Refresh data in console
                    if (enteredValue.Equals("R", StringComparison.CurrentCultureIgnoreCase)) {
                        htmlParserUtility.WriteAccountDataToConsole();
                        Prompt();
                        enteredValue = Console.ReadLine() ?? string.Empty;
                    }
                    // Write data to Excel workbook
                    else if (enteredValue.Equals("W", StringComparison.CurrentCultureIgnoreCase) ||
                             enteredValue.Equals("O", StringComparison.CurrentCultureIgnoreCase)) {
                        var accounts = htmlParserUtility.GetAccountData();

                        var excelMapper = iocKernel.Get<IExcelMapper>(
                            new ConstructorArgument("accounts", accounts),
                            new ConstructorArgument("workbookPath",
                                ConfigurationManager.AppSettings["ExcelWorkBookPath"]),
                            new ConstructorArgument("sheetName", ConfigurationManager.AppSettings["ExcelSheetName"]));

                        // If the entered value is O then overwrite the transaction columns with 0
                        if (enteredValue.Equals("O", StringComparison.CurrentCultureIgnoreCase)) {
                            excelMapper.SetLastTransactionToZero = true;
                        }
                        excelMapper.WriteToWorkSheet();

                        Prompt();
                        enteredValue = Console.ReadLine() ?? string.Empty;
                    }
                    else {
                        userAccount.Logout();
                    }
                }
            }
        }
    }
}
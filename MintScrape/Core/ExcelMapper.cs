using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using MintScrape.Interfaces;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MintScrape.Core {
    /// <summary>
    ///     Maps the accounts' values to their coordinates in the Excel workbook.
    /// </summary>
    public class ExcelMapper : IExcelMapper {
        private readonly IEnumerable<IAccount> _accounts;
        private readonly string _sheetName;
        private readonly string _workbookPath;
        private readonly int _accountNameColumnNumber;
        private readonly int _balanceColumnNumber;
        private readonly int _transactionColumnNumber;

        public ExcelMapper(IEnumerable<IAccount> accounts, string workbookPath, string sheetName,
            bool setLastTransactionToZero = false) {
            _accounts = accounts;
            _workbookPath = workbookPath;
            _sheetName = sheetName;
            SetLastTransactionToZero = setLastTransactionToZero;
            SetAccountsExcelRowField();
            _balanceColumnNumber = int.Parse(ConfigurationManager.AppSettings["BalanceColumnNumber"]);
            _accountNameColumnNumber = int.Parse(ConfigurationManager.AppSettings["AccountNameColumnNumber"]);
            _transactionColumnNumber = int.Parse(ConfigurationManager.AppSettings["PaymentColumnNumber"]);
        }

        /// <summary>
        ///     Determines whether to set the transaction column to zero.  Useful for clearing out old transactions.  Only applies
        ///     to mapped accounts.  Other transactions will remain.
        /// </summary>
        public bool SetLastTransactionToZero { get; set; }

        /// <summary>
        ///     Write account information to Excel spreadhseet.
        /// </summary>
        public void WriteToWorkSheet() {
            Console.WriteLine(@"Writing to {0} on '{1}'", _workbookPath, _sheetName);
            Console.WriteLine(Environment.NewLine);

            var errorAccounts = new List<IAccount>();
            try {
                using (var stream = new FileStream(_workbookPath, FileMode.Open, FileAccess.ReadWrite)) {
                    var wb = new XSSFWorkbook(stream);
                    var sheet = wb.GetSheet(_sheetName);

                    foreach (var account in _accounts.Where(x => !string.IsNullOrEmpty(x.NickName))) {
                        var row = sheet.GetRow(account.ExcelSheetRow);
                        if (row != null && row.RowNum > 0) {
                            Console.WriteLine("Setting {0} => {1} on row {2}", account.NickName, account.Balance,
                                row.RowNum);
                            row.Cells[_balanceColumnNumber].SetCellValue(account.Balance);
                            if (SetLastTransactionToZero) {
                                row.Cells[_transactionColumnNumber].SetCellValue(0); // Sets the last payment made back to 0    
                            }
                        } else {
                            errorAccounts.Add(account);
                        }
                    }

                    // Re-evaluates all formula cells to be sure they are calculated correctly after setting the new values
                    XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb);

                    // Writes to workbook
                    wb.Write(new FileStream(_workbookPath, FileMode.Create, FileAccess.Write));
                }

                Console.WriteLine("Done writing to spreadsheet.");
                Console.WriteLine(Environment.NewLine);
                foreach (var errorAccount in errorAccounts) {
                    Console.WriteLine("Could not set {0} => {1}", errorAccount.NickName, errorAccount.Balance);
                }
                Console.WriteLine(Environment.NewLine);
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                Console.WriteLine(Environment.NewLine);
            }
        }

        /// <summary>
        ///     Gets account name to row number dictionary.
        /// </summary>
        /// <returns>Account name to row number dictionary.</returns>
        private Dictionary<string, int> GetAccountNameToRowNumDictionary() {
            // Loop through and populate the row number for each account object
            using (var fs = new FileStream(_workbookPath, FileMode.Open)) {
                // Get workbook
                var wb = new XSSFWorkbook(fs);

                // Get sheet
                var sheet = wb.GetSheet(_sheetName);

                // Build out a dictionary of ExcelColName and RowNumber
                var colToRowDic = new Dictionary<string, int>();

                // Loop through first 150 rows 
                for (var i = 0; i < 150; i++) {
                    var row = sheet.GetRow(i);
                    var cell = row?.GetCell(0);

                    // If account name isn't a string or formula then continue to next iteration
                    if (cell == null || (cell.CellType != CellType.String && cell.CellType != CellType.Formula))
                        continue;

                    // Get account name from cell value
                    var accountName = row.GetCell(_accountNameColumnNumber, MissingCellPolicy.CREATE_NULL_AS_BLANK).RichStringCellValue.String;

                    // Add row name and row num to dictionary
                    if (!string.IsNullOrEmpty(accountName)) {
                        try {
                            colToRowDic.Add(accountName, row.RowNum);
                        } catch (Exception ex) {
                            Console.WriteLine("Error at row {0} - {1} - {2}: ", i, accountName, ex.Message);
                            Console.WriteLine();
                            Console.WriteLine(Environment.NewLine);
                        }
                    }
                }
                return colToRowDic;
            }
        }

        /// <summary>
        ///     Sets the ExcelSheetRow property for each account.
        /// </summary>
        private void SetAccountsExcelRowField() {
            // Loop through and populate the row number for each account object
            var dictionary = GetAccountNameToRowNumDictionary();
            foreach (var account in _accounts) {
                var rowNum = 0;
                dictionary.TryGetValue(account.NickName, out rowNum);
                if (rowNum > 0) {
                    account.ExcelSheetRow = rowNum;
                }
            }
        }
    }
}
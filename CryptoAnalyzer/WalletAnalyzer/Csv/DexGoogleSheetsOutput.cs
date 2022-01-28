using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using WebScraper;
using WebScraper.WebScrapers.EtherscanDex;

namespace WalletAnalyzer
{
    public class DexGoogleSheetsOutput : IDexOutput
    {
        private readonly OutputOptions _config;
        private ICellStyle _dateFormatStyle;
        private readonly string _dateFormatString;
        private readonly int _indexFirstRowForDexTable;

        public DexGoogleSheetsOutput(IOptions<OutputOptions> config)
        {
            _config = config.Value;
            _dateFormatString = _config.DateFormatForXlsx;
            _indexFirstRowForDexTable = _config.IndexFirstRowDexTable;
        }

        public void DoOutput(string outputName, string tokenHash, DexTableOutputDto table, string timeElapsed, int nmRows)
        {
            var pathName = _config.Path;
            var fullPath = pathName + '/' + outputName + ".xlsx";
            Directory.CreateDirectory(pathName);

            try
            {
                var workbook = new XSSFWorkbook();
                var sheet1 = workbook.CreateSheet();

                _dateFormatStyle = workbook.CreateCellStyle();
                _dateFormatStyle.DataFormat = workbook.CreateDataFormat().GetFormat(_dateFormatString);

                WriteScrapeInfo(sheet1.CreateRow(0), timeElapsed, nmRows);
                WriteTokenInfo(sheet1.CreateRow(1), table.TokenName, tokenHash);

                AddFormulaForLastSell(_indexFirstRowForDexTable, table.Rows, $"P${_indexFirstRowForDexTable+1}:R", $"R${_indexFirstRowForDexTable+1}:R", "D");
                
                AddDexTableRows(sheet1, _indexFirstRowForDexTable-1, 0, table.Rows);
                AddFormulaFilterSellersOnlyTable(
                    sheet1.GetRow(_indexFirstRowForDexTable).CreateCell(15),
                    $"B${_indexFirstRowForDexTable+1}:D",
                    $"C${_indexFirstRowForDexTable+1}:C",
                    DexAction.Sell.ToString());

                CreateOutputFile(workbook, fullPath);
            }
            catch (IOException)
            {
                throw;
            }
            catch
            {
                System.Console.WriteLine("Error outputing xlsx");
                State.ExitAndLog(new StackTrace()); //fix
            }
        }

        private void AddFormulaForLastSell(int dexTableFirstRowIndex, List<DexRowOutputDto> rows, string sellersTable, string sellerTableHashArray, string mainTableFromHashColumnLetter)
        {
            for (var i = 0; i < rows.Count; i++)
            {
                var currentRow = dexTableFirstRowIndex+1 + i;

                var sellersTableRange = XlsxIndirectRange(sellersTable);
                var sellerTableHashArrayRange = XlsxIndirectRange(sellerTableHashArray);

                rows[i].LastSell = $"INDEX({sellersTableRange},MATCH({mainTableFromHashColumnLetter}{currentRow},{sellerTableHashArrayRange},0),1)";
            }
        }

        private void WriteScrapeInfo(IRow row, string timeElapsed, int nmRows)
        {
            row.CreateCell(0).SetCellValue("Time elapsed");
            row.CreateCell(1).SetCellValue(timeElapsed);

            row.CreateCell(3).SetCellValue("Rows scraped");
            row.CreateCell(4).SetCellValue(nmRows);
        }

        private void WriteTokenInfo(IRow row, string tokenName, string tokenHash)
        {
            row.CreateCell(0).SetCellValue("Token Name");
            row.CreateCell(1).SetCellValue(tokenName);

            row.CreateCell(3).SetCellValue("Token Hash");
            row.CreateCell(4).SetCellValue(tokenHash);
        }

        private void AddFormulaFilterSellersOnlyTable(ICell cell, string tableFor_txnDate_Action_FromHash, string actionArray, string sellActionString)
        {
            var rangeActionArray = XlsxIndirectRange(actionArray);
            var rangeTableFor_txnDate_Action_FromHash = XlsxIndirectRange(tableFor_txnDate_Action_FromHash);
            var formula = $"FILTER({rangeTableFor_txnDate_Action_FromHash}, {rangeActionArray}=\"{sellActionString}\")";
            try
            {
                cell.SetCellFormula(formula);
            }
            catch
            {
                var t = 0;

            }
        }

        private void CreateOutputFile(XSSFWorkbook workbook, string path)
        {
            var outputStream = File.Create(path);
            workbook.Write(outputStream);
            workbook.Close();
        }

        private void AddDexTableRows(ISheet sheet, int initIndexForTableHeaderRow, int initColumnIndex, List<DexRowOutputDto> rows)
        {
            var headerXlsxRow = sheet.CreateRow(initIndexForTableHeaderRow);
            headerXlsxRow.CreateCell(initColumnIndex).SetCellValue("TxnHash");
            headerXlsxRow.CreateCell(initColumnIndex+1).SetCellValue("TxnDate");
            headerXlsxRow.CreateCell(initColumnIndex+2).SetCellValue("Action");
            headerXlsxRow.CreateCell(initColumnIndex+3).SetCellValue("FromHash");
            headerXlsxRow.CreateCell(initColumnIndex+4).SetCellValue("ToHash");
            headerXlsxRow.CreateCell(initColumnIndex+5).SetCellValue("LastSell");

            for (var i = 0; i < rows.Count; i++)
            {
                var dexRow = rows[i];
                var xlsxRow = sheet.CreateRow(initIndexForTableHeaderRow+1 + i);
                xlsxRow.CreateCell(initColumnIndex).SetCellValue(dexRow.TxnHash);

                var txnDateCell = xlsxRow.CreateCell(initColumnIndex + 1);
                txnDateCell.SetCellValue(dexRow.TxnDate);
                txnDateCell.CellStyle = _dateFormatStyle;

                xlsxRow.CreateCell(initColumnIndex + 2).SetCellValue(dexRow.Action.ToString());
                xlsxRow.CreateCell(initColumnIndex + 3).SetCellValue(dexRow.FromHash);
                xlsxRow.CreateCell(initColumnIndex + 4).SetCellValue(dexRow.ToHash);

                try
                {
                    var cell = xlsxRow.CreateCell(initColumnIndex + 5);
                    cell.SetCellFormula(dexRow.LastSell);
                }
                catch
                {
                    System.Console.WriteLine("Error setting formula");
                    throw; //fix
                }
            }
        }

        private string XlsxIndirectRange(string range)  //NPOI throws exception for range like A3:A which would be valid in Google sheets, but not valid in Excel
        {
            return $"INDIRECT(\"{range}\")";  //Google sheets function INDIRECT solves this
        }
    }
}


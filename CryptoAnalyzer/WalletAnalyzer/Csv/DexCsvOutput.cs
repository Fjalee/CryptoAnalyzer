using CsvHelper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using WebScraper;
using WebScraper.Parsers;

namespace WalletAnalyzer
{
    public class DexCsvOutput : IDexOutput
    {
        private readonly OutputOptions _config;

        public DexCsvOutput(IOptions<OutputOptions> config)
        {
            _config = config.Value;
        }

        public void DoOutput(string outputName, string tokenHash, DexTableOutputDto table, string timeElapsed, int nmRows)
        {
            var pathName = _config.Path;
            var fullPath = pathName + '/' + outputName + ".csv";

            AddGoogleSpreadsheetsFormulaForLastSell(5, table.Rows, "I$3:K", "K$3:K", "D");

            Directory.CreateDirectory(pathName);

            try
            {
                using (var writer = new StreamWriter(fullPath))
                using (var csvWriter = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    //WriteCsvOptions(csvWriter);
                    WriteScrapeInfo(csvWriter, timeElapsed, nmRows);
                    WriteTokenInfo(csvWriter, table.TokenName, tokenHash);
                    AddGoogleSpreadsheetsFormulaForSellersOnlyTable(csvWriter);

                    csvWriter.NextRecord();
                    csvWriter.Context.RegisterClassMap<DexRowOutputDtoMap>();
                    csvWriter.WriteRecords(table.Rows);
                }
            }
            catch (IOException)
            {
                throw;
            }
            catch
            {
                State.ExitAndLog(new StackTrace());
            }
        }

        private void AddGoogleSpreadsheetsFormulaForSellersOnlyTable(CsvWriter csvWriter)
        {
            SkipCells(csvWriter, 7);
            csvWriter.WriteField("Sell");
            csvWriter.WriteField("=FILTER(B$5:D, C$5:C=H3)");
        }

        private void AddGoogleSpreadsheetsFormulaForLastSell(int firstRowIndexSpreadsheet, List<DexRowOutputDto> rows, string sellersTable, string sellerTableHashArray, string mainTableFromHashColumnLetter)
        {
            for (var i = 0; i < rows.Count; i++)
            {
                var currentRow = firstRowIndexSpreadsheet + i;
                rows[i].LastSell = $"=INDEX({sellersTable},MATCH({mainTableFromHashColumnLetter}{currentRow},{sellerTableHashArray},0),1)";

            }
        }

        private void SkipCells(CsvWriter csvWriter, int amount)
        {
            for (var i = 0; i < amount; i++)
            {
                csvWriter.WriteField("");
            }
        }

        private void WriteCsvOptions(CsvWriter csvWriter)
        {
            csvWriter.WriteField("sep=,");
            csvWriter.NextRecord();
        }

        private void WriteScrapeInfo(CsvWriter csvWriter, string timeElapsed, int nmRows)
        {
            csvWriter.WriteField("Time elapsed: ");
            csvWriter.WriteField(timeElapsed);
            csvWriter.WriteField("");
            csvWriter.WriteField("Rows scraped: ");
            csvWriter.WriteField(nmRows);
            csvWriter.NextRecord();
        }

        private void WriteTokenInfo(CsvWriter csvWriter, string tokenName, string tokenHash)
        {
            csvWriter.WriteField("Token Name: ");
            csvWriter.WriteField(tokenName);
            csvWriter.WriteField("");
            csvWriter.WriteField("Token Hash: ");
            csvWriter.WriteField(tokenHash);
            csvWriter.NextRecord();
        }
    }
}

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

        public void DoOutput(string outputName, string tokenHash, DexTableCsvOutputDto table, string timeElapsed, int nmRows)
        {
            var pathName = _config.Path;
            var fullPath = pathName + '/' + outputName + ".csv";

            Directory.CreateDirectory(pathName);

            try
            {
                using (var writer = new StreamWriter(fullPath))
                using (var csvWriter = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    WriteCsvOptions(csvWriter);
                    WriteScrapeInfo(csvWriter, timeElapsed, nmRows);
                    WriteTokenInfo(csvWriter, table.TokenName, tokenHash);

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

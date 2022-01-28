using Microsoft.Extensions.Options;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using WebScraper;

namespace WalletAnalyzer
{
    public class DexGoogleSpreadsheetsOutput : IDexOutput
    {
        private readonly OutputOptions _config;

        public DexGoogleSpreadsheetsOutput(IOptions<OutputOptions> config)
        {
            _config = config.Value;
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
                OutputDexTableRows(sheet1, 0 , 0, table.Rows);

                CreateOutputFile(workbook, fullPath);
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

        private void CreateOutputFile(XSSFWorkbook workbook, string path)
        {

            if (File.Exists(path))
            {
                System.Console.WriteLine($"file {path} already exists");
                State.ExitAndLog(new StackTrace());
            }
            var outputStream = File.Create(path);
            workbook.Write(outputStream);
            workbook.Close();
        }

        private void OutputDexTableRows(ISheet sheet, int initRowIndex, int initColumnIndex, List<DexRowOutputDto> rows)
        {
             var headerXlsxRow = sheet.CreateRow(initRowIndex);
            headerXlsxRow.CreateCell(initColumnIndex).SetCellValue("TxnHash");
            headerXlsxRow.CreateCell(initColumnIndex+1).SetCellValue("TxnDate");
            headerXlsxRow.CreateCell(initColumnIndex+2).SetCellValue("Action");
            headerXlsxRow.CreateCell(initColumnIndex+3).SetCellValue("FromHash");
            headerXlsxRow.CreateCell(initColumnIndex+4).SetCellValue("ToHash");
            headerXlsxRow.CreateCell(initColumnIndex+5).SetCellValue("LastSell");

            for (var i = 0; i < rows.Count; i++)
            {
                var dexRow = rows[i];
                var xlsxRow = sheet.CreateRow(initRowIndex+1 + i);
                xlsxRow.CreateCell(initColumnIndex).SetCellValue(dexRow.TxnHash);
                xlsxRow.CreateCell(initColumnIndex + 1).SetCellValue(dexRow.TxnDate);
                xlsxRow.CreateCell(initColumnIndex + 2).SetCellValue(dexRow.Action.ToString());
                xlsxRow.CreateCell(initColumnIndex + 3).SetCellValue(dexRow.FromHash);
                xlsxRow.CreateCell(initColumnIndex + 4).SetCellValue(dexRow.ToHash);
                xlsxRow.CreateCell(initColumnIndex + 5).SetCellValue(dexRow.LastSell);
            }
        }
    }
}


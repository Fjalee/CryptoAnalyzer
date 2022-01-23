using AutoMapper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using WebScraper;
using WebScraper.WebScrapers;

namespace WalletAnalyzer
{
    public class DexCollector : IDexCollector
    {
        private readonly IDexScraperFactory _dexScrapperFactory;
        private readonly IDexOutput _dexOutput;
        private readonly IMapper _mapper;
        private readonly List<DexRow> _allNewRows = new List<DexRow>();
        private readonly DexTable _tableToOutput = new DexTable();
        private readonly Stopwatch _stopwatch = new Stopwatch();

        private int _totalRowsScraped = 0;
        private long _msWorthOfDataOutputed = 0;
        private bool _isNeededSaveAsap = false;

        public DexCollector(IDexScraperFactory dexScraperFactory, IDexOutput dexOutput, IMapper mapper)
        {
            _dexScrapperFactory = dexScraperFactory;
            _dexOutput = dexOutput;
            _mapper = mapper;
        }

        public async Task Start(string url, string tokenHash, int sleepTimeMs, int appendPeriodInMs, int nmRowsToScrape)
        {
            var scrapeStartDate = DateTime.Now.ToString("yyyy_MM_dd_HHmm");
            var dexScrapper = _dexScrapperFactory.CreateScrapper(url);
            var outputName = "";
            var isLastOutputDone = false;
            //Trace.Listeners.Add(new TextWriterTraceListener(ConfigurationManager.AppSettings.Get("LOG_PATH")));
            //Trace.AutoFlush = true;
            //fix

            _stopwatch.Start();

            while(_totalRowsScraped < nmRowsToScrape)
            {
                var pageDexTable = await dexScrapper.ScrapeCurrentPageTable();

                if(pageDexTable == null)
                {
                    break;
                }

                _allNewRows.AddRange(pageDexTable.Rows);
                _totalRowsScraped += pageDexTable.Rows.Count;

                ManageTableName(pageDexTable.TokenName);

                Thread.Sleep(sleepTimeMs);

                if (_msWorthOfDataOutputed + appendPeriodInMs < _stopwatch.ElapsedMilliseconds
                    || _isNeededSaveAsap)
                {
                    outputName = $"{_tableToOutput.TokenName}_{scrapeStartDate}";
                    isLastOutputDone = TryOutput(outputName, tokenHash);
                }

                dexScrapper.GoToNextPage();
            }

            while (!isLastOutputDone)
            {
                isLastOutputDone = TryOutput(outputName, tokenHash);
            }
        }

        private void ManageTableName(string newTableName)
        {

            if (_tableToOutput.TokenName == null)
            {
                _tableToOutput.TokenName = newTableName;
            }
            else
            {
                if (_tableToOutput.TokenName != newTableName)
                {
                    State.ExitAndLog(new StackTrace()); //fix
                }
            }
        }

        private bool TryOutput(string outputName, string tokenHash)
        {
            _tableToOutput.Rows.AddRange(_allNewRows);
            _allNewRows.Clear();

            var ts = _stopwatch.Elapsed;
            var timeOutput = string.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

            try
            {
                var outputTable = _mapper.Map<DexTableOutputDto>(_tableToOutput);
                _dexOutput.DoOutput(outputName, tokenHash, outputTable, timeOutput, _totalRowsScraped);
                _msWorthOfDataOutputed = _stopwatch.ElapsedMilliseconds;

                if (_isNeededSaveAsap)
                {
                    _isNeededSaveAsap = false;
                    Console.WriteLine("Output file closed. Scraped data was added SUCCESSFULLY...");
                }

                Console.WriteLine("Appended data scraped in " + timeOutput);
                return true;
            }
            catch(IOException)
            {
                _isNeededSaveAsap = true;
                Console.WriteLine("Scraped data could not be added. Please close the output file...");
                return false;
            }
        }
    }
}

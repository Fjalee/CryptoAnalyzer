using System;
using WebScraper.Parsers;
using WebScraper.WebScrapers.EtherscanDex;

namespace WebScraper.WebScrapers
{
    public class DexRow : IRow
    {
        public string TxnHash { get; set; }
        public DateTime TxnDate { get; set; }
        public DexAction Action { get; set; }
        public string FromHash{ get; set; }
        public string ToHash{ get; set; }
    }
}